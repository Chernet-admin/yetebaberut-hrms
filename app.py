import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import sqlite3
import io
import os
import random
import base64
import subprocess
import sys
import calendar
import json
from datetime import datetime, date, timedelta
import re

import xlsxwriter
import psycopg2
import psycopg2.extensions

st.set_page_config(page_title="Yetebaberut GSP — HRMS",page_icon="🏢",layout="wide",initial_sidebar_state="expanded")

# ══════════════════════════════════════════════════════════════════
# POSTGRES COMPATIBILITY SHIM
# The rest of this app was written against sqlite3's interface
# (conn.execute(sql,params) with "?" placeholders, PRAGMA table_info,
# INSERT OR IGNORE, BLOB columns, sqlite3.IntegrityError, etc).
# Rather than rewrite every one of the ~150 call sites, this shim wraps
# a real psycopg2 connection so it behaves like sqlite3 from the rest
# of the app's point of view — translating syntax differences at the
# point of execute() instead.
# ══════════════════════════════════════════════════════════════════

_PRAGMA_TABLE_INFO_RE = re.compile(r"PRAGMA\s+table_info\((\w+)\)", re.IGNORECASE)

def _pg_prepare_sql(sql):
    # ? -> %s placeholders
    sql = sql.replace("?", "%s")
    # BLOB -> BYTEA (schema statements only; harmless elsewhere since
    # the literal word BLOB never appears in this app's data/queries)
    sql = re.sub(r"\bBLOB\b", "BYTEA", sql, flags=re.IGNORECASE)
    # SQLite autoincrement -> Postgres serial
    sql = re.sub(r"INTEGER\s+PRIMARY\s+KEY\s+AUTOINCREMENT", "SERIAL PRIMARY KEY", sql, flags=re.IGNORECASE)
    # INSERT OR IGNORE -> INSERT ... ON CONFLICT DO NOTHING
    if re.search(r"INSERT\s+OR\s+IGNORE", sql, flags=re.IGNORECASE):
        sql = re.sub(r"INSERT\s+OR\s+IGNORE", "INSERT", sql, flags=re.IGNORECASE)
        sql = sql.rstrip().rstrip(";") + " ON CONFLICT DO NOTHING"
    return sql

def _pg_prepare_params(params):
    if not params: return params
    out=[]
    for p in params:
        if isinstance(p,(bytes,memoryview)):
            out.append(psycopg2.Binary(bytes(p)))
        else:
            out.append(p)
    return type(params)(out) if not isinstance(params,list) else out

class _PGCursor:
    def __init__(self, real_cursor, conn_wrapper):
        self._c=real_cursor
        self._w=conn_wrapper
        self._pragma_rows=None

    def execute(self, sql, params=()):
        m=_PRAGMA_TABLE_INFO_RE.search(sql)
        if m:
            table=m.group(1)
            try:
                self._c.execute("""SELECT column_name,data_type,is_nullable
                    FROM information_schema.columns WHERE table_name=%s ORDER BY ordinal_position""",(table,))
                rows=self._c.fetchall()
                # Shape like sqlite's PRAGMA table_info: (cid,name,type,notnull,dflt,pk)
                self._pragma_rows=[(i,r[0],r[1],0 if r[2]=='YES' else 1,None,0) for i,r in enumerate(rows)]
            except Exception:
                self._pragma_rows=[]
            return self
        if re.match(r"^\s*PRAGMA", sql, flags=re.IGNORECASE):
            return self  # no-op: PRAGMA journal_mode/synchronous/etc has no Postgres equivalent needed
        self._pragma_rows=None
        prepped=_pg_prepare_sql(sql)
        pparams=_pg_prepare_params(params)
        try:
            self._c.execute(prepped, pparams)
        except psycopg2.errors.UniqueViolation as e:
            self._w._conn.rollback(); raise sqlite3.IntegrityError(str(e))
        except psycopg2.errors.NotNullViolation as e:
            self._w._conn.rollback(); raise sqlite3.IntegrityError(str(e))
        except psycopg2.errors.ForeignKeyViolation as e:
            self._w._conn.rollback(); raise sqlite3.IntegrityError(str(e))
        except psycopg2.IntegrityError as e:
            self._w._conn.rollback(); raise sqlite3.IntegrityError(str(e))
        except psycopg2.Error as e:
            self._w._conn.rollback(); raise sqlite3.OperationalError(str(e))
        return self

    def executemany(self, sql, params_seq):
        if re.search(r"INSERT\s+OR\s+IGNORE|INSERT\s+INTO", sql, flags=re.IGNORECASE):
            prepped=_pg_prepare_sql(sql)
            try:
                for p in params_seq:
                    self._c.execute(prepped,_pg_prepare_params(p))
            except psycopg2.Error as e:
                self._w._conn.rollback(); raise sqlite3.OperationalError(str(e))
            return self
        prepped=_pg_prepare_sql(sql)
        try:
            self._c.executemany(prepped,[_pg_prepare_params(p) for p in params_seq])
        except psycopg2.Error as e:
            self._w._conn.rollback(); raise sqlite3.OperationalError(str(e))
        return self

    def fetchone(self):
        if self._pragma_rows is not None:
            return self._pragma_rows.pop(0) if self._pragma_rows else None
        return self._c.fetchone()

    def fetchall(self):
        if self._pragma_rows is not None:
            r=self._pragma_rows; self._pragma_rows=[]; return r
        return self._c.fetchall()

    @property
    def description(self): return self._c.description
    @property
    def rowcount(self): return self._c.rowcount

class _PGConnWrapper:
    """Mimics sqlite3.Connection's interface (execute/executemany shortcuts,
    commit/close/cursor) on top of a real psycopg2 connection."""
    def __init__(self, real_conn):
        self._conn=real_conn

    def cursor(self):
        return _PGCursor(self._conn.cursor(), self)

    def execute(self, sql, params=()):
        cur=self.cursor(); cur.execute(sql, params); return cur

    def executemany(self, sql, params_seq):
        cur=self.cursor(); cur.executemany(sql, params_seq); return cur

    def commit(self): self._conn.commit()
    def rollback(self): self._conn.rollback()
    def close(self): self._conn.close()

def pg_read_sql(sql, conn, params=None):
    """Drop-in replacement for pd.read_sql_query against our Postgres shim
    (pandas' read_sql_query only special-cases sqlite3.Connection/SQLAlchemy,
    so a plain wrapped psycopg2 connection needs this instead)."""
    cur=conn.cursor()
    cur.execute(sql, params or ())
    cols=[d[0] for d in cur.description] if cur.description else []
    rows=cur.fetchall()
    return pd.DataFrame(rows, columns=cols)

def get_conn():
    # Railway (and most PaaS Postgres add-ons) inject a single DATABASE_URL
    # env var. Prefer that; fall back to individual PG* env vars; fall back
    # to st.secrets for local development with a .streamlit/secrets.toml.
    database_url = os.environ.get("DATABASE_URL")
    if database_url:
        # Railway's URL sometimes starts with postgres:// — psycopg2 accepts either.
        real_conn = psycopg2.connect(database_url, sslmode=os.environ.get("PGSSLMODE", "prefer"), connect_timeout=10)
        return _PGConnWrapper(real_conn)

    if all(k in os.environ for k in ("PGHOST","PGPORT","PGDATABASE","PGUSER","PGPASSWORD")):
        real_conn = psycopg2.connect(
            host=os.environ["PGHOST"],
            port=os.environ["PGPORT"],
            dbname=os.environ["PGDATABASE"],
            user=os.environ["PGUSER"],
            password=os.environ["PGPASSWORD"],
            sslmode=os.environ.get("PGSSLMODE", "prefer"),
            connect_timeout=10,
        )
        return _PGConnWrapper(real_conn)

    real_conn = psycopg2.connect(
        host=st.secrets["postgres"]["host"],
        port=st.secrets["postgres"]["port"],
        dbname=st.secrets["postgres"]["database"],
        user=st.secrets["postgres"]["user"],
        password=st.secrets["postgres"]["password"],
        sslmode="require",
        connect_timeout=10,
    )
    return _PGConnWrapper(real_conn)

@st.cache_data
def get_holidays(year):
    return {
        date(year,1,7):"Christmas (Genna)",date(year,1,19):"Epiphany (Timkat)",
        date(year,3,2):"Adwa Victory Day",date(year,4,9):"Easter (Fasika)",
        date(year,5,1):"Labour Day",date(year,5,5):"Patriots Victory Day",
        date(year,5,28):"Derg Downfall Day",date(year,9,11):"New Year (Enkutatash)",
        date(year,9,27):"Meskel",date(year,11,8):"Mawlid",
    }

# ════════════════════════════════════════════════════════
# DATABASE SCHEMA — Performance-indexed for 10,000+ records
# ════════════════════════════════════════════════════════
def init_db():
    conn=get_conn(); c=conn.cursor()
    c.execute("""CREATE TABLE IF NOT EXISTS employees(
        emp_id TEXT PRIMARY KEY,full_name TEXT NOT NULL,division TEXT NOT NULL,
        cost_center TEXT,
        contact TEXT,email TEXT,house_address TEXT,woreda TEXT,subcity TEXT,kebele TEXT,
        resident_id TEXT,place_of_birth TEXT,age INTEGER,sex TEXT,
        marital_status TEXT,nationality TEXT,religion TEXT,
        emergency_contact_name TEXT,emergency_contact_phone TEXT,
        blood_type TEXT,tin_number TEXT,pension_number TEXT,
        bank_name TEXT,bank_account TEXT,
        edu_background TEXT,field_of_graduate TEXT,graduation_year TEXT,institution_name TEXT,
        current_status TEXT DEFAULT 'Pending Screening',
        job_title TEXT,employment_type TEXT,start_date TEXT,contract_end_date TEXT,
        weekly_dayoff TEXT DEFAULT 'Sunday',
        basic_salary REAL DEFAULT 0,registration_date TEXT,notes TEXT,
        photo_name TEXT,photo_data BLOB,
        edu_doc_name TEXT,edu_doc_data BLOB,
        forensic_doc_name TEXT,forensic_doc_data BLOB,
        id_scan_name TEXT,id_scan_data BLOB,
        medical_doc_name TEXT,medical_doc_data BLOB,
        guarantee_letter_name TEXT,guarantee_letter_data BLOB,
        police_clearance_name TEXT,police_clearance_data BLOB,
        contract_doc_name TEXT,contract_doc_data BLOB,
        first_doc_name TEXT,first_doc_data BLOB)""")

    c.execute("PRAGMA table_info(employees)")
    ex=[col[1] for col in c.fetchall()]
    migrations={"division":"TEXT","cost_center":"TEXT","weekly_dayoff":"TEXT DEFAULT 'Sunday'",
        "marital_status":"TEXT","nationality":"TEXT","religion":"TEXT",
        "emergency_contact_name":"TEXT","emergency_contact_phone":"TEXT","blood_type":"TEXT",
        "tin_number":"TEXT","pension_number":"TEXT","bank_name":"TEXT","bank_account":"TEXT",
        "graduation_year":"TEXT","institution_name":"TEXT","job_title":"TEXT","employment_type":"TEXT",
        "start_date":"TEXT","contract_end_date":"TEXT","basic_salary":"REAL DEFAULT 0","notes":"TEXT",
        "photo_name":"TEXT","photo_data":"BLOB","edu_doc_name":"TEXT","edu_doc_data":"BLOB",
        "forensic_doc_name":"TEXT","forensic_doc_data":"BLOB","id_scan_name":"TEXT","id_scan_data":"BLOB",
        "medical_doc_name":"TEXT","medical_doc_data":"BLOB","guarantee_letter_name":"TEXT",
        "guarantee_letter_data":"BLOB","police_clearance_name":"TEXT","police_clearance_data":"BLOB",
        "contract_doc_name":"TEXT","contract_doc_data":"BLOB","first_doc_name":"TEXT","first_doc_data":"BLOB",
        "department":"TEXT","annual_leave_entitlement":"INTEGER DEFAULT 20",
        # ── Official HR Record Format fields (added to match the standardized
        # company template — Personal / Address / Emergency Contact /
        # Mortgage Condition (Guarantor) / Financial & IDs / Education /
        # Employment & Division groups) ──
        "contact2":"TEXT","city":"TEXT","national_id_number":"TEXT",
        "emergency_contact_city":"TEXT","emergency_contact_subcity":"TEXT","emergency_contact_woreda":"TEXT",
        "guarantor_name":"TEXT","guarantor_phone":"TEXT","guarantor_city":"TEXT",
        "guarantor_subcity":"TEXT","guarantor_woreda":"TEXT","guarantor_company_id":"TEXT",
        "guarantor_company_name":"TEXT","guarantor_letter_number":"TEXT","guarantor_date_written":"TEXT"}
    for col,typ in migrations.items():
        if col not in ex:
            try: c.execute(f"ALTER TABLE employees ADD COLUMN {col} {typ}"); conn.commit()
            except: pass
    # Migrate old department -> division if needed
    try:
        c.execute("SELECT COUNT(*) FROM employees WHERE division IS NULL AND department IS NOT NULL")
        if c.fetchone()[0]>0:
            c.execute("UPDATE employees SET division=department WHERE division IS NULL")
            conn.commit()
    except: pass
    # One-time backfill: employees created before "Contact 01/02" split still
    # have their phone number in the original `contact` column only — leave
    # it there (it IS Contact 01) so no data is lost.

    # ── Repair legacy 'department' column: some databases still carry it
    # as NOT NULL from an older schema version, which blocks every new
    # employee insert (division is the real column used everywhere now).
    try:
        c.execute("PRAGMA table_info(employees)")
        dept_info=[col for col in c.fetchall() if col[1]=="department"]
        if dept_info and dept_info[0][3]==1:  # notnull flag set
            try:
                c.execute("ALTER TABLE employees DROP COLUMN department")
                conn.commit()
            except Exception:
                # Older SQLite without DROP COLUMN support: keep the column
                # but make sure it's never left NULL going forward.
                c.execute("UPDATE employees SET department=division WHERE department IS NULL")
                conn.commit()
    except: pass

    # ── PERFORMANCE INDEXES (critical for 10,000+ records) ──
    for idx_sql in [
        "CREATE INDEX IF NOT EXISTS idx_emp_status ON employees(current_status)",
        "CREATE INDEX IF NOT EXISTS idx_emp_division ON employees(division)",
        "CREATE INDEX IF NOT EXISTS idx_emp_costcenter ON employees(cost_center)",
        "CREATE INDEX IF NOT EXISTS idx_emp_name ON employees(full_name)",
    ]:
        try: c.execute(idx_sql)
        except: pass
    conn.commit()

    c.execute("""CREATE TABLE IF NOT EXISTS payroll(
        id INTEGER PRIMARY KEY AUTOINCREMENT,emp_id TEXT,month TEXT,
        basic_salary REAL DEFAULT 0,transport_allowance REAL DEFAULT 0,
        housing_allowance REAL DEFAULT 0,other_allowance REAL DEFAULT 0,
        income_tax REAL DEFAULT 0,pension_employee REAL DEFAULT 0,
        pension_employer REAL DEFAULT 0,other_deductions REAL DEFAULT 0,
        fine_amount REAL DEFAULT 0,fine_days INTEGER DEFAULT 0,
        sick_leave_days INTEGER DEFAULT 0,annual_leave_days INTEGER DEFAULT 0,
        maternity_leave_days INTEGER DEFAULT 0,mourning_leave_days INTEGER DEFAULT 0,
        unpaid_leave_days INTEGER DEFAULT 0,absent_days INTEGER DEFAULT 0,
        holiday_days INTEGER DEFAULT 0,dayoff_days INTEGER DEFAULT 4,
        gross_salary REAL DEFAULT 0,net_salary REAL DEFAULT 0,
        payment_status TEXT DEFAULT 'Pending',notes TEXT,created_at TEXT)""")
    c.execute("PRAGMA table_info(payroll)")
    pcols=[col[1] for col in c.fetchall()]
    for pcol,ptyp in {"gross_salary":"REAL DEFAULT 0","absent_days":"INTEGER DEFAULT 0",
        "holiday_days":"INTEGER DEFAULT 0","dayoff_days":"INTEGER DEFAULT 4",
        "fine_days":"INTEGER DEFAULT 0","fine_amount":"REAL DEFAULT 0",
        "sick_leave_days":"INTEGER DEFAULT 0","annual_leave_days":"INTEGER DEFAULT 0",
        "maternity_leave_days":"INTEGER DEFAULT 0","mourning_leave_days":"INTEGER DEFAULT 0",
        "unpaid_leave_days":"INTEGER DEFAULT 0","pension_employer":"REAL DEFAULT 0",
        "full_name":"TEXT","division":"TEXT","cost_center":"TEXT"}.items():
        if pcol not in pcols:
            try: c.execute(f"ALTER TABLE payroll ADD COLUMN {pcol} {ptyp}"); conn.commit()
            except: pass
    c.execute("CREATE INDEX IF NOT EXISTS idx_payroll_emp ON payroll(emp_id)")
    c.execute("CREATE INDEX IF NOT EXISTS idx_payroll_month ON payroll(month)")

    # ── PAYROLL SUBMISSION & APPROVAL WORKFLOW ──
    # Supervisor compiles a cost center's attendance for the month and submits.
    # Payroll Section reviews and approves before salaries are released.
    c.execute("""CREATE TABLE IF NOT EXISTS payroll_submissions(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        cost_center TEXT NOT NULL,
        division TEXT NOT NULL,
        month TEXT NOT NULL,
        submitted_by TEXT,
        submitted_at TEXT,
        status TEXT DEFAULT 'Pending Approval',
        reviewed_by TEXT,
        reviewed_at TEXT,
        review_notes TEXT,
        employee_count INTEGER DEFAULT 0,
        total_net_amount REAL DEFAULT 0,
        data_snapshot TEXT,
        UNIQUE(cost_center,month))""")
    c.execute("CREATE INDEX IF NOT EXISTS idx_subm_cc ON payroll_submissions(cost_center)")
    c.execute("CREATE INDEX IF NOT EXISTS idx_subm_status ON payroll_submissions(status)")

    c.execute("""CREATE TABLE IF NOT EXISTS fine_letters(
        id INTEGER PRIMARY KEY AUTOINCREMENT,emp_id TEXT,month TEXT,
        issue_date TEXT,fine_reason TEXT,fine_type TEXT DEFAULT 'Disciplinary',
        fine_days INTEGER DEFAULT 0,fine_amount REAL DEFAULT 0,
        letter_name TEXT,letter_data BLOB,
        applied_to_payroll TEXT DEFAULT 'No',created_at TEXT)""")
    c.execute("PRAGMA table_info(fine_letters)")
    flc=[col[1] for col in c.fetchall()]
    for flcol,fltyp in {"month":"TEXT","fine_type":"TEXT DEFAULT 'Disciplinary'","fine_reason":"TEXT",
        "record_status":"TEXT DEFAULT 'Active'","cancelled_by":"TEXT","cancelled_at":"TEXT","cancel_reason":"TEXT",
        "compensated_days":"INTEGER DEFAULT 0","compensation_notes":"TEXT"}.items():
        if flcol not in flc:
            try: c.execute(f"ALTER TABLE fine_letters ADD COLUMN {flcol} {fltyp}"); conn.commit()
            except: pass
    c.execute("CREATE INDEX IF NOT EXISTS idx_fine_emp ON fine_letters(emp_id)")

    c.execute("""CREATE TABLE IF NOT EXISTS leave_records(
        id INTEGER PRIMARY KEY AUTOINCREMENT,emp_id TEXT,leave_type TEXT,
        start_date TEXT,end_date TEXT,days_taken INTEGER DEFAULT 0,
        is_paid INTEGER DEFAULT 1,daily_rate REAL DEFAULT 0,
        deduction_amount REAL DEFAULT 0,approved_by TEXT,
        status TEXT DEFAULT 'Approved',notes TEXT,created_at TEXT)""")
    c.execute("PRAGMA table_info(leave_records)")
    lrc=[col[1] for col in c.fetchall()]
    for lrcol,lrtyp in {"edited_by":"TEXT","edited_at":"TEXT","cancelled_by":"TEXT","cancelled_at":"TEXT","cancel_reason":"TEXT","doc_name":"TEXT","doc_data":"BLOB"}.items():
        if lrcol not in lrc:
            try: c.execute(f"ALTER TABLE leave_records ADD COLUMN {lrcol} {lrtyp}"); conn.commit()
            except: pass
    c.execute("CREATE INDEX IF NOT EXISTS idx_leave_emp ON leave_records(emp_id)")

    # weekly_dayoff lives on employees table now — keep schedule table for overrides/history
    c.execute("""CREATE TABLE IF NOT EXISTS dayoff_schedule(
        id INTEGER PRIMARY KEY AUTOINCREMENT,emp_id TEXT,
        month TEXT,dayoff_date TEXT,dayoff_notes TEXT,created_at TEXT)""")
    c.execute("CREATE INDEX IF NOT EXISTS idx_dayoff_emp ON dayoff_schedule(emp_id)")

    c.execute("""CREATE TABLE IF NOT EXISTS absent_records(
        id INTEGER PRIMARY KEY AUTOINCREMENT,emp_id TEXT,
        absent_date TEXT,reason TEXT,is_excused INTEGER DEFAULT 0,created_at TEXT)""")
    c.execute("PRAGMA table_info(absent_records)")
    arc=[col[1] for col in c.fetchall()]
    for arcol,artyp in {"record_status":"TEXT DEFAULT 'Active'","cancelled_by":"TEXT","cancelled_at":"TEXT",
        "cancel_reason":"TEXT","is_compensated":"INTEGER DEFAULT 0","compensation_date":"TEXT","compensation_notes":"TEXT"}.items():
        if arcol not in arc:
            try: c.execute(f"ALTER TABLE absent_records ADD COLUMN {arcol} {artyp}"); conn.commit()
            except: pass
    c.execute("CREATE INDEX IF NOT EXISTS idx_absent_emp ON absent_records(emp_id)")

    # ── RECYCLE BIN ──
    # Deleted employees and key records are moved here instead of
    # being permanently erased, so a Manager can restore them.
    c.execute("""CREATE TABLE IF NOT EXISTS recycle_bin(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        record_type TEXT NOT NULL,
        record_id TEXT NOT NULL,
        record_label TEXT,
        record_data TEXT,
        deleted_by TEXT,
        deleted_at TEXT,
        restored INTEGER DEFAULT 0)""")
    c.execute("CREATE INDEX IF NOT EXISTS idx_recycle_type ON recycle_bin(record_type)")

    c.execute("""CREATE TABLE IF NOT EXISTS employee_documents(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        emp_id TEXT NOT NULL,
        doc_type TEXT,
        doc_name TEXT,
        doc_data BLOB,
        notes TEXT,
        uploaded_by TEXT,
        uploaded_at TEXT)""")
    c.execute("CREATE INDEX IF NOT EXISTS idx_empdoc_emp ON employee_documents(emp_id)")

    c.execute("""CREATE TABLE IF NOT EXISTS daily_status_records(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        emp_id TEXT NOT NULL,
        status TEXT NOT NULL,
        leave_type TEXT,
        start_date TEXT NOT NULL,
        end_date TEXT NOT NULL,
        num_days INTEGER,
        reason TEXT,
        doc_name TEXT,
        doc_data BLOB,
        supervisor_id TEXT,
        submitted_at TEXT,
        workflow_stage TEXT DEFAULT 'Pending HR Review',
        gs_reviewed_by TEXT, gs_reviewed_at TEXT, gs_comments TEXT,
        hr_reviewed_by TEXT, hr_reviewed_at TEXT, hr_comments TEXT,
        linked_leave_id INTEGER, linked_absent_id INTEGER)""")
    c.execute("CREATE INDEX IF NOT EXISTS idx_dsr_emp ON daily_status_records(emp_id)")
    c.execute("CREATE INDEX IF NOT EXISTS idx_dsr_stage ON daily_status_records(workflow_stage)")

    c.execute("""CREATE TABLE IF NOT EXISTS system_users(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,password TEXT NOT NULL,
        role TEXT NOT NULL DEFAULT 'Data Officer',full_name TEXT,email TEXT,
        permissions TEXT DEFAULT 'view_only',is_active INTEGER DEFAULT 1,
        assigned_division TEXT,
        created_by TEXT,created_at TEXT,last_login TEXT)""")
    c.execute("PRAGMA table_info(system_users)")
    su_cols=[col[1] for col in c.fetchall()]
    if "assigned_division" not in su_cols:
        try: c.execute("ALTER TABLE system_users ADD COLUMN assigned_division TEXT"); conn.commit()
        except: pass
    if "nav_access" not in su_cols:
        try: c.execute("ALTER TABLE system_users ADD COLUMN nav_access TEXT"); conn.commit()
        except: pass
    c.execute("SELECT COUNT(*) FROM system_users")
    if c.fetchone()[0]==0:
        now=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        c.executemany("INSERT OR IGNORE INTO system_users(username,password,role,full_name,permissions,is_active,assigned_division,created_by,created_at)VALUES(?,?,?,?,?,1,?,'system',?)",[
            ("ygs_manager","secure2026","Manager","System Manager","full",None,now),
            ("ygs_officer","data2026","Data Officer","Data Officer","view_only",None,now),
            ("ygs_payroll","payroll2026","Payroll Section","Payroll Officer","payroll_approve",None,now),
        ]); conn.commit()

    # ── DEMOTION & STATUS MANAGEMENT ──
    # HR submits a grade/title/salary change request; Manager approves it;
    # once approved, Finance is considered notified and the employee record
    # (job_title, grade if tracked via job_title, basic_salary) is updated.
    c.execute("""CREATE TABLE IF NOT EXISTS demotion_records(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        emp_id TEXT NOT NULL,
        change_type TEXT DEFAULT 'Demotion',
        previous_title TEXT,
        previous_salary REAL,
        new_title TEXT,
        new_salary REAL,
        reason TEXT,
        submitted_by TEXT,
        submitted_at TEXT,
        status TEXT DEFAULT 'Pending Manager Approval',
        reviewed_by TEXT,
        reviewed_at TEXT,
        review_notes TEXT,
        finance_notified INTEGER DEFAULT 0,
        applied INTEGER DEFAULT 0)""")
    c.execute("CREATE INDEX IF NOT EXISTS idx_demotion_emp ON demotion_records(emp_id)")
    c.execute("CREATE INDEX IF NOT EXISTS idx_demotion_status ON demotion_records(status)")

    # ── COST CENTERS table (manually created per division) ──
    c.execute("""CREATE TABLE IF NOT EXISTS cost_centers(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        code TEXT UNIQUE NOT NULL,name TEXT NOT NULL,division TEXT NOT NULL,
        budget REAL DEFAULT 0,description TEXT,is_active INTEGER DEFAULT 1,
        created_by TEXT,created_at TEXT)""")

    # ── DIVISIONS table (manually created — same pattern as Cost Centers) ──
    # Divisions are no longer a hardcoded list baked into the app. A Manager
    # creates them here first (Cost Centers page → Divisions tab), and every
    # division dropdown across the system (Applicant Intake, Employee
    # Directory, Employee Profile, Payroll, Supervisor assignment, etc.)
    # reads from this table.
    c.execute("""CREATE TABLE IF NOT EXISTS divisions(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE NOT NULL,
        description TEXT,
        is_active INTEGER DEFAULT 1,
        created_by TEXT,created_at TEXT)""")
    c.execute("CREATE INDEX IF NOT EXISTS idx_division_active ON divisions(is_active)")

    # ── SYSTEM SETTINGS (applicant gate control + leave/overtime policy) ──
    c.execute("""CREATE TABLE IF NOT EXISTS system_settings(
        key TEXT PRIMARY KEY, value TEXT, updated_by TEXT, updated_at TEXT)""")
    c.execute("SELECT COUNT(*) FROM system_settings WHERE key='applications_open'")
    if c.fetchone()[0]==0:
        c.execute("INSERT INTO system_settings(key,value,updated_by,updated_at) VALUES('applications_open','1','system',?)",
            (datetime.now().strftime("%Y-%m-%d %H:%M:%S"),))
        conn.commit()

    # Leave / overtime / holiday policy defaults (Ethiopian Labour Proc. 1156/2019)
    # Manager can adjust these in Public Holidays — all payroll calculations read from here.
    POLICY_DEFAULTS = {
        "policy_annual_leave_days":"20",
        "policy_sick_leave_full_months":"1",
        "policy_sick_leave_half_months":"2",
        "policy_maternity_leave_days":"90",
        "policy_paternity_leave_days":"3",
        "policy_mourning_leave_days":"3",
        "policy_working_days_per_month":"26",
        "policy_overtime_weekday":"1.25",
        "policy_overtime_weekend":"1.5",
        "policy_overtime_holiday":"2.0",
        "policy_holiday_payment_status":"Paid",
        "policy_dayoff_payment_status":"Paid",
        "policy_sick_payment_status":"Paid",
        "policy_unpaid_leave_payment_status":"Unpaid",
    }
    now_p=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for pk,pv in POLICY_DEFAULTS.items():
        c.execute("SELECT COUNT(*) FROM system_settings WHERE key=?",(pk,))
        if c.fetchone()[0]==0:
            c.execute("INSERT INTO system_settings(key,value,updated_by,updated_at) VALUES(?,?,'system',?)",(pk,pv,now_p))
    conn.commit()

    conn.commit(); conn.close()

init_db()

# ════════════════════════════════════════════════════════
# SMART DAY-OFF CALCULATOR
# Given a weekday name (e.g. "Monday") and a month/year,
# automatically computes every calendar date that falls on
# that weekday — fully autonomous, recalculates every time.
# ════════════════════════════════════════════════════════
WEEKDAY_MAP={"Monday":0,"Tuesday":1,"Wednesday":2,"Thursday":3,"Friday":4,"Saturday":5,"Sunday":6}

def get_dayoff_dates(weekday_name, year, month):
    """Return list of date objects in given year/month matching the chosen weekday."""
    target=WEEKDAY_MAP.get(weekday_name,6)
    _,days_in_month=calendar.monthrange(year,month)
    dates=[]
    for d in range(1,days_in_month+1):
        dt=date(year,month,d)
        if dt.weekday()==target:
            dates.append(dt)
    return dates

def count_dayoffs_in_month(weekday_name, year, month):
    return len(get_dayoff_dates(weekday_name,year,month))

# ════════════════════════════════════════════════════════
# CACHED QUERIES — fast at 10,000+ records
# ════════════════════════════════════════════════════════
@st.cache_data(ttl=20)
def get_stats():
    conn=get_conn()
    df=pg_read_sql("SELECT current_status,COUNT(*) as c FROM employees GROUP BY current_status",conn)
    conn.close()
    return dict(zip(df['current_status'],df['c']))

@st.cache_data(ttl=20)
def get_division_list():
    """Divisions are manually created by the Manager (Cost Centers page →
    Divisions tab) — there is no hardcoded default list any more. Legacy
    records that still carry a division name not (yet) registered in the
    divisions table are appended at the end so old data never disappears
    from filters/dropdowns."""
    conn=get_conn()
    df=pg_read_sql("SELECT name FROM divisions WHERE is_active=1 ORDER BY name",conn)
    legacy=pg_read_sql("SELECT DISTINCT division FROM employees WHERE division IS NOT NULL AND division!='' ORDER BY division",conn)
    conn.close()
    created=df['name'].tolist() if len(df)>0 else []
    extra=[d for d in legacy['division'].tolist() if d and d not in created]
    return created+extra

@st.cache_data(ttl=20)
def get_cost_centers(division=None):
    conn=get_conn()
    if division and division!="All":
        df=pg_read_sql("SELECT * FROM cost_centers WHERE division=? AND is_active=1 ORDER BY code",conn,params=(division,))
    else:
        df=pg_read_sql("SELECT * FROM cost_centers WHERE is_active=1 ORDER BY division,code",conn)
    conn.close()
    return df

@st.cache_data(ttl=20)
def get_emp_list_cached():
    conn=get_conn()
    df=pg_read_sql("SELECT emp_id,full_name,division,cost_center,current_status FROM employees ORDER BY emp_id LIMIT 5000",conn)
    conn.close(); return df

@st.cache_data(ttl=15)
def count_records(status_filter,div_filter,cc_filter,search):
    conn=get_conn()
    q="SELECT COUNT(*) as c FROM employees WHERE 1=1"
    p=[]
    if status_filter and status_filter!="All": q+=" AND current_status=?"; p.append(status_filter)
    if div_filter and div_filter!="All": q+=" AND division=?"; p.append(div_filter)
    if cc_filter and cc_filter!="All": q+=" AND cost_center=?"; p.append(cc_filter)
    if search: q+=" AND (full_name LIKE ? OR emp_id LIKE ?)"; p.extend([f"%{search}%",f"%{search}%"])
    total=pg_read_sql(q,conn,params=p).iloc[0]['c']; conn.close(); return int(total)

@st.cache_data(ttl=15)
def query_records(status_filter,div_filter,cc_filter,search,page=1,page_size=50):
    conn=get_conn()
    q="""SELECT emp_id,full_name,job_title,division,cost_center,sex,basic_salary,contact,current_status,registration_date
         FROM employees WHERE 1=1"""
    p=[]
    if status_filter and status_filter!="All": q+=" AND current_status=?"; p.append(status_filter)
    if div_filter and div_filter!="All": q+=" AND division=?"; p.append(div_filter)
    if cc_filter and cc_filter!="All": q+=" AND cost_center=?"; p.append(cc_filter)
    if search: q+=" AND (full_name LIKE ? OR emp_id LIKE ?)"; p.extend([f"%{search}%",f"%{search}%"])
    offset = max(page-1,0)*page_size
    q+=" ORDER BY emp_id LIMIT ? OFFSET ?"; p.extend([page_size,offset])
    df=pg_read_sql(q,conn,params=p); conn.close()
    # Replace database NULLs with a clean placeholder instead of showing literal "None"
    df = df.fillna("—")
    for col in df.columns:
        df[col] = df[col].apply(lambda v: "—" if v=="None" or v is None else v)
    return df

@st.cache_data(ttl=10)
def get_employee(eid):
    """Loads the employee's core profile + document NAMES, but deliberately
    excludes the heavy document BLOB columns (edu_doc_data, id_scan_data,
    etc.) so a routine profile view doesn't drag megabytes of file bytes
    across the network on every load. Use get_employee_document_blob() to
    fetch one specific document's bytes only when the user asks to view it.
    photo_data is kept since the profile header always shows a thumbnail."""
    heavy_blob_cols = {
        "edu_doc_data","forensic_doc_data","id_scan_data","medical_doc_data",
        "guarantee_letter_data","police_clearance_data","contract_doc_data","first_doc_data",
    }
    conn=get_conn(); cur=conn.cursor()
    cur.execute("PRAGMA table_info(employees)")
    all_cols=[c[1] for c in cur.fetchall()]
    select_cols=[c for c in all_cols if c not in heavy_blob_cols]
    cur.execute(f"SELECT {','.join(select_cols)} FROM employees WHERE emp_id=?",(eid,))
    row=cur.fetchone(); conn.close()
    if not row:
        return None
    result = dict(zip(select_cols,row))
    for c in heavy_blob_cols:
        result[c] = None  # placeholder; fetched on demand via get_employee_document_blob()
    # psycopg2 returns BYTEA columns as memoryview, which st.cache_data
    # cannot pickle — convert to plain bytes so caching this function works
    # for employees who actually have a photo uploaded.
    if result.get("photo_data") is not None:
        result["photo_data"] = bytes(result["photo_data"])
    return result

def get_employee_document_blob(eid, data_col):
    """Fetch a single document's bytes on demand — called only when the
    user clicks to preview/download that specific document, not as part
    of the routine profile load."""
    conn=get_conn(); cur=conn.cursor()
    cur.execute(f"SELECT {data_col} FROM employees WHERE emp_id=?",(eid,))
    row=cur.fetchone(); conn.close()
    return row[0] if row else None


@st.cache_data(ttl=15)
def get_active_employee_payroll_summary():
    """All active employees (incl. on leave types) with their current month financial summary."""
    conn=get_conn()
    df=pg_read_sql("""SELECT emp_id,full_name,division,cost_center,basic_salary,current_status
        FROM employees WHERE current_status != 'Terminated' ORDER BY emp_id LIMIT 5000""",conn)
    conn.close(); return df

@st.cache_data(ttl=10)
def get_setting(key, default="1"):
    conn=get_conn(); cur=conn.cursor()
    cur.execute("SELECT value FROM system_settings WHERE key=?",(key,))
    row=cur.fetchone(); conn.close()
    return row[0] if row else default

def set_setting(key, value, user):
    conn=get_conn()
    conn.execute("INSERT INTO system_settings(key,value,updated_by,updated_at) VALUES(?,?,?,?) ON CONFLICT(key) DO UPDATE SET value=?,updated_by=?,updated_at=?",
        (key,value,user,datetime.now().strftime("%Y-%m-%d %H:%M:%S"),value,user,datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    conn.commit(); conn.close()
    get_setting.clear()

# ════════════════════════════════════════════════════════
# PAYROLL HELPERS
# ════════════════════════════════════════════════════════
def eth_tax(g):
    if g<=600:return 0
    elif g<=1650:return g*0.10-60
    elif g<=3200:return g*0.15-142.5
    elif g<=5250:return g*0.20-302.5
    elif g<=7800:return g*0.25-565
    elif g<=10900:return g*0.30-955
    else:return g*0.35-1500

def calc_pay(basic,transport,housing,other,fine,unpaid,absent,extra):
    try: working_days = float(get_setting("policy_working_days_per_month","26"))
    except: working_days = 26
    gross=basic+transport+housing+other; daily=basic/working_days
    tax=eth_tax(gross); pen=basic*0.07; pen_er=basic*0.11
    net=max(gross-tax-pen-fine-(daily*(unpaid+absent))-extra,0)
    return round(net,2),round(tax,2),round(pen,2),round(pen_er,2),round(daily,2),round(gross,2)

def b64file(data,name):
    if not data or not name: return None,None
    return base64.b64encode(bytes(data)).decode(), name.split(".")[-1].lower()

def get_annual_leave_balance(emp_id, year, entitlement):
    """Returns (used_days, remaining_days) of Annual Leave for emp_id in the given year."""
    conn=get_conn()
    used=pg_read_sql("""SELECT COALESCE(SUM(days_taken),0) as used FROM leave_records
        WHERE emp_id=? AND leave_type='Annual Leave' AND status != 'Cancelled'
        AND start_date LIKE ?""",conn,params=(emp_id,f"{year}%"))
    conn.close()
    used_days=int(used.iloc[0]['used']) if len(used)>0 else 0
    return used_days, max(entitlement-used_days,0)

def preview_html(data,name,label="Document"):
    if not data or not name:
        return f'<div style="text-align:center;padding:28px;color:#6B7FA3"><div style="font-size:36px;opacity:0.3"></div><div style="font-size:12px;margin-top:6px">No {label} uploaded</div></div>'
    b64,ext=b64file(data,name)
    if ext in ["jpg","jpeg","png","gif","webp"]:
        return f'<div style="text-align:center"><img src="data:image/{ext};base64,{b64}" style="max-width:100%;max-height:460px;border-radius:10px;border:1px solid rgba(212,168,71,0.2);object-fit:contain"/><div style="font-size:10px;color:#6B7FA3;margin-top:5px">{name}</div></div>'
    elif ext=="pdf":
        return f'<div style="width:100%;height:480px;border-radius:10px;overflow:hidden;border:1px solid rgba(212,168,71,0.2)"><iframe src="data:application/pdf;base64,{b64}" width="100%" height="100%" style="border:none"></iframe></div><div style="font-size:10px;color:#6B7FA3;margin-top:4px;text-align:center">{name}</div>'
    return f'<div style="background:#131F38;border-radius:8px;padding:12px;color:#94A8C8;font-size:12px"> {name} — download to open</div>'

def print_slip(emp,pay,company="Yetebaberut General Service Provider"):
    bs=float(pay.get("basic_salary",0)); daily=bs/26
    return f"""<!DOCTYPE html><html><head><meta charset="UTF-8"><style>
body{{font-family:Arial,sans-serif;margin:0;padding:14px;color:#000}}
.slip{{border:2px solid #D4A847;border-radius:8px;padding:20px;max-width:750px;margin:auto}}
.header{{text-align:center;border-bottom:2px solid #D4A847;padding-bottom:10px;margin-bottom:16px}}
.co{{font-size:18px;font-weight:bold;color:#0D1526}}.ti{{font-size:12px;color:#666;margin-top:2px}}
.grid{{display:grid;grid-template-columns:1fr 1fr;gap:8px;background:#F8F8F8;padding:10px;border-radius:6px;margin-bottom:14px}}
.lbl{{font-size:9px;color:#888;text-transform:uppercase}}.val{{font-size:12px;font-weight:500}}
table{{width:100%;border-collapse:collapse;margin-bottom:12px}}
th{{background:#0D1526;color:#D4A847;padding:6px 10px;font-size:10px;text-align:left;text-transform:uppercase}}
td{{padding:6px 10px;font-size:12px;border-bottom:1px solid #eee}}
.ded{{color:#c0392b}}.add{{color:#27ae60}}
.nr{{background:#0D1526}}.nr td{{color:#D4A847;padding:10px;font-weight:bold;font-size:14px}}
.footer{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:16px;text-align:center;margin-top:18px}}
.sig{{border-top:1px solid #000;margin-top:34px;font-size:10px;color:#666;padding-top:3px}}
@media print{{body{{margin:0}}}}
</style></head><body><div class="slip">
<div class="header"><div class="co">{company}</div>
<div class="ti">PAYROLL STATEMENT — {pay.get("month","")}</div>
<div style="font-size:9px;color:#888;margin-top:2px">Addis Ababa, Ethiopia | info@yetebaberut.com | {datetime.now().strftime("%Y-%m-%d")}</div></div>
<div class="grid">
<div><div class="lbl">Employee ID</div><div class="val">{emp.get("emp_id","")}</div></div>
<div><div class="lbl">Full Name</div><div class="val">{emp.get("full_name","")}</div></div>
<div><div class="lbl">Division</div><div class="val">{emp.get("division","")}</div></div>
<div><div class="lbl">Cost Center</div><div class="val">{emp.get("cost_center","—")}</div></div>
<div><div class="lbl">Job Title</div><div class="val">{emp.get("job_title","—")}</div></div>
<div><div class="lbl">Bank / Account</div><div class="val">{emp.get("bank_name","—")} / {emp.get("bank_account","—")}</div></div>
</div>
<table><tr><th>EARNINGS</th><th style="text-align:right">ETB</th></tr>
<tr><td>Basic Salary</td><td style="text-align:right">{float(pay.get("basic_salary",0)):,.2f}</td></tr>
<tr><td>Transport Allowance</td><td style="text-align:right">{float(pay.get("transport_allowance",0)):,.2f}</td></tr>
<tr><td>Housing Allowance</td><td style="text-align:right">{float(pay.get("housing_allowance",0)):,.2f}</td></tr>
<tr><td>Other Allowance</td><td style="text-align:right">{float(pay.get("other_allowance",0)):,.2f}</td></tr>
<tr style="font-weight:bold;background:#f5f5f5"><td>GROSS</td><td style="text-align:right">{float(pay.get("gross_salary",0)):,.2f}</td></tr></table>
<table><tr><th>DEDUCTIONS</th><th style="text-align:right">ETB</th></tr>
<tr><td class="ded">Income Tax</td><td class="ded" style="text-align:right">-{float(pay.get("income_tax",0)):,.2f}</td></tr>
<tr><td class="ded">Employee Pension 7%</td><td class="ded" style="text-align:right">-{float(pay.get("pension_employee",0)):,.2f}</td></tr>
<tr><td class="ded">Fines ({pay.get("fine_days",0)} days)</td><td class="ded" style="text-align:right">-{float(pay.get("fine_amount",0)):,.2f}</td></tr>
<tr><td class="ded">Unpaid Leave ({pay.get("unpaid_leave_days",0)} days)</td><td class="ded" style="text-align:right">-{float(pay.get("unpaid_leave_days",0))*daily:,.2f}</td></tr>
<tr><td class="ded">Absent ({pay.get("absent_days",0)} days)</td><td class="ded" style="text-align:right">-{float(pay.get("absent_days",0))*daily:,.2f}</td></tr>
<tr><td class="ded">Other Deductions</td><td class="ded" style="text-align:right">-{float(pay.get("other_deductions",0)):,.2f}</td></tr>
<tr><td class="add">Paid Leave (Sick/Annual/Mat/Mourning)</td><td class="add" style="text-align:right"> Paid</td></tr>
<tr><td class="add">Day-Off ({pay.get("dayoff_days",4)} days — {pay.get("dayoff_weekday","Sunday")}s)</td><td class="add" style="text-align:right"> Paid</td></tr>
<tr><td class="add">Public Holidays ({pay.get("holiday_days",0)} days)</td><td class="add" style="text-align:right"> Paid</td></tr></table>
<table><tr class="nr"><td>NET SALARY</td><td style="text-align:right;font-size:16px">ETB {float(pay.get("net_salary",0)):,.2f}</td></tr></table>
<table><tr><th>EMPLOYER</th><th style="text-align:right">ETB</th></tr>
<tr><td class="add">Employer Pension 11%</td><td class="add" style="text-align:right">{float(pay.get("pension_employer",0)):,.2f}</td></tr></table>
<div class="footer">
<div><div class="sig">Employee Signature</div></div>
<div><div class="sig">HR Officer</div></div>
<div><div class="sig">Manager / Director</div></div></div></div>
<script>window.onload=function(){{window.print()}}</script></body></html>"""

def soft_delete(record_type, record_id, record_label, record_data_dict, deleted_by):
    """Move a record's data into the recycle bin before deleting it for real."""
    conn=get_conn()
    conn.execute("INSERT INTO recycle_bin(record_type,record_id,record_label,record_data,deleted_by,deleted_at,restored)VALUES(?,?,?,?,?,?,0)",
        (record_type, str(record_id), record_label, json.dumps(record_data_dict, default=str), deleted_by, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    conn.commit(); conn.close()

def export_excel(df):
    safe=[c for c in df.columns if not c.endswith("_data")]
    buf=io.BytesIO()
    with pd.ExcelWriter(buf,engine="xlsxwriter") as wr:
        df[safe].to_excel(wr,index=False,sheet_name="Employees")
        wb=wr.book; ws=wr.sheets["Employees"]
        hf=wb.add_format({"bold":True,"bg_color":"#0D1526","font_color":"#D4A847","border":1,"font_size":11})
        cf=wb.add_format({"bg_color":"#060B18","font_color":"#E8EEF7","border":1,"font_size":10})
        af=wb.add_format({"bg_color":"#0A1020","font_color":"#C8D8F0","border":1,"font_size":10})
        for i,col in enumerate(safe):
            ws.write(0,i,col,hf); ws.set_column(i,i,max(len(col)+4,16))
        for ri in range(1,len(df)+1):
            for ci,col in enumerate(safe):
                val=df[safe].iloc[ri-1,ci]
                ws.write(ri,ci,str(val) if val is not None else "",cf if ri%2==0 else af)
    return buf.getvalue()

# ════════════════════════════════════════════════════════
# OFFICIAL HR RECORD FORMAT — "NEW HR Personal Record Data Format 2026"
# Standardized company template. Every full employee-record export (and
# the printable single-employee record) uses this exact grouping and
# column order:
#   Personal Information | Address | Emergency Contact |
#   Mortgage Condition (Guarantor) | Financial & IDs | Education |
#   Employment & Division
# ════════════════════════════════════════════════════════
HR_FORMAT_GROUPS = [
    ("Personal Information", [
        ("Employee ID","emp_id"),("Full Name","full_name"),("Contact 01","contact"),
        ("Contact 02","contact2"),("Email","email"),("Sex","sex"),("Marital","marital_status"),
        ("Nationality","nationality"),("Religion","religion"),("Age","age"),
        ("Place of Birth","place_of_birth"),("Blood","blood_type"),("Resident ID","resident_id"),
    ]),
    ("Address", [
        ("City","city"),("Subcity","subcity"),("Woreda","woreda"),("House","house_address"),
    ]),
    ("Emergency Contact", [
        ("Name","emergency_contact_name"),("Phone","emergency_contact_phone"),
        ("City","emergency_contact_city"),("Subcity","emergency_contact_subcity"),
        ("Woreda","emergency_contact_woreda"),
    ]),
    ("Mortgage Condition", [
        ("Name","guarantor_name"),("Phone","guarantor_phone"),("City","guarantor_city"),
        ("Subcity","guarantor_subcity"),("Woreda","guarantor_woreda"),
        ("Company ID","guarantor_company_id"),("Company Name","guarantor_company_name"),
        ("Letter Number","guarantor_letter_number"),("Date Written","guarantor_date_written"),
    ]),
    ("Financial & IDs", [
        ("National Id Number","national_id_number"),("TIN Id Number","tin_number"),
        ("Pension Id Number","pension_number"),("Bank Name","bank_name"),("Account","bank_account"),
    ]),
    ("Education", [
        ("Level","edu_background"),("Field","field_of_graduate"),("Grad Year","graduation_year"),
        ("Institution","institution_name"),
    ]),
    ("Employment & Division", [
        ("Job Title","job_title"),("Type","employment_type"),("Division","division"),
        ("Cost Center","cost_center"),("Basic Salary","basic_salary"),
        ("Weekly Day-Off","weekly_dayoff"),("Start Date (YYYY-MM-DD)","start_date"),
        ("Contract End (YYYY-MM-DD)","contract_end_date"),("Status","current_status"),
        ("Internal Notes","notes"),
    ]),
]
HR_FORMAT_FIELDS = [f for _,fields in HR_FORMAT_GROUPS for _,f in fields]

def export_excel_hr_official(df):
    """Export employee records to the standardized company template
    (NEW HR Personal Record Data Format 2026) — two header rows: a merged
    group-label row, then the exact field-name row, in the exact column
    order the company uses for every official HR register."""
    buf=io.BytesIO()
    with pd.ExcelWriter(buf,engine="xlsxwriter") as wr:
        wb=wr.book
        ws=wb.add_worksheet("Data")
        grp_fmt=wb.add_format({"bold":True,"bg_color":"#0D1526","font_color":"#D4A847","border":1,
            "align":"center","valign":"vcenter","font_size":11})
        hdr_fmt=wb.add_format({"bold":True,"bg_color":"#D4A847","font_color":"#0D1526","border":1,
            "align":"center","valign":"vcenter","font_size":10})
        cell_fmt=wb.add_format({"bg_color":"#060B18","font_color":"#E8EEF7","border":1,"font_size":10})
        alt_fmt=wb.add_format({"bg_color":"#0A1020","font_color":"#C8D8F0","border":1,"font_size":10})

        col=0
        for group_label,fields in HR_FORMAT_GROUPS:
            span=len(fields)
            if span>1:
                ws.merge_range(0,col,0,col+span-1,group_label,grp_fmt)
            else:
                ws.write(0,col,group_label,grp_fmt)
            for label,_ in fields:
                ws.write(1,col,label,hdr_fmt)
                ws.set_column(col,col,max(len(label)+3,14))
                col+=1

        for ri in range(len(df)):
            row=df.iloc[ri]
            for ci,field in enumerate(HR_FORMAT_FIELDS):
                val=row[field] if field in df.columns else ""
                if val is None or (isinstance(val,float) and pd.isna(val)): val=""
                ws.write(2+ri,ci,val,cell_fmt if ri%2==0 else alt_fmt)
        ws.freeze_panes(2,1)
    return buf.getvalue()

def print_employee_record_html(emp,company="Yetebaberut General Service Provider"):
    """Single-employee printable record in the same standardized grouping
    and order as the Excel export — for a full personal-record printout
    (not the payroll slip, which stays a separate document)."""
    sections=""
    for group_label,fields in HR_FORMAT_GROUPS:
        rows="".join([
            f'<tr><td class="lbl">{label}</td><td class="val">{emp.get(field,"") or "—"}</td></tr>'
            for label,field in fields])
        sections+=f'<div class="grp"><div class="grp-title">{group_label}</div><table>{rows}</table></div>'
    return f"""<!DOCTYPE html><html><head><meta charset="UTF-8"><style>
body{{font-family:Arial,sans-serif;margin:0;padding:18px;color:#000}}
.wrap{{max-width:820px;margin:auto;border:2px solid #D4A847;border-radius:8px;padding:20px}}
.header{{text-align:center;border-bottom:2px solid #D4A847;padding-bottom:10px;margin-bottom:14px}}
.co{{font-size:18px;font-weight:bold;color:#0D1526}}.ti{{font-size:12px;color:#666;margin-top:2px}}
.grp{{margin-bottom:12px}}
.grp-title{{background:#0D1526;color:#D4A847;font-size:11px;font-weight:bold;text-transform:uppercase;
  padding:5px 10px;border-radius:4px 4px 0 0}}
table{{width:100%;border-collapse:collapse}}
td{{padding:5px 10px;font-size:12px;border:1px solid #ddd}}
.lbl{{background:#F8F8F8;color:#555;font-weight:600;width:34%}}
.val{{color:#111}}
@media print{{body{{margin:0}}}}
</style></head><body><div class="wrap">
<div class="header"><div class="co">{company}</div>
<div class="ti">OFFICIAL EMPLOYEE PERSONAL RECORD</div>
<div style="font-size:9px;color:#888;margin-top:2px">Addis Ababa, Ethiopia | Printed {datetime.now().strftime("%Y-%m-%d")}</div></div>
{sections}
</div>
<script>window.onload=function(){{window.print()}}</script></body></html>"""

# ════════════════════════════════════════════════════════
# CLEANUP — remove any auto-seeded cost centers.
# This app used to auto-create a default set of cost centers on first
# run. The user manages cost centers manually, so this now does the
# opposite: it removes any leftover system-created ones instead of
# creating new ones, and never seeds anything going forward.
# ════════════════════════════════════════════════════════
def seed_if_empty():
    conn=get_conn()
    conn.execute("DELETE FROM cost_centers WHERE created_by='system'")
    conn.execute("DELETE FROM divisions WHERE created_by='system'")
    conn.commit(); conn.close()

seed_if_empty()
