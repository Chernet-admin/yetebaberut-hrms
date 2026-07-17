import streamlit as st
import pandas as pd
import sqlite3
import io
import random
import base64
import subprocess
import sys
import calendar
import json
from datetime import datetime, date, timedelta

import xlsxwriter

st.set_page_config(page_title="Yetebaberut GSP — HRMS",layout="wide",initial_sidebar_state="collapsed")
DB_FILE="yetebaberut_enterprise.db"

def get_conn():
    conn = sqlite3.connect(DB_FILE, check_same_thread=False, timeout=30)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    conn.execute("PRAGMA cache_size=-64000")
    conn.execute("PRAGMA temp_store=MEMORY")
    return conn

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
        "department":"TEXT"}
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
        "unpaid_leave_days":"INTEGER DEFAULT 0","pension_employer":"REAL DEFAULT 0"}.items():
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
    for lrcol,lrtyp in {"edited_by":"TEXT","edited_at":"TEXT","cancelled_by":"TEXT","cancelled_at":"TEXT","cancel_reason":"TEXT"}.items():
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

    # ── COST CENTERS table (manually created per division) ──
    c.execute("""CREATE TABLE IF NOT EXISTS cost_centers(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        code TEXT UNIQUE NOT NULL,name TEXT NOT NULL,division TEXT NOT NULL,
        budget REAL DEFAULT 0,description TEXT,is_active INTEGER DEFAULT 1,
        created_by TEXT,created_at TEXT)""")

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

try: init_db()
except Exception as _e:
    import os
    if os.path.exists(DB_FILE): os.remove(DB_FILE)
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
    df=pd.read_sql_query("SELECT current_status,COUNT(*) as c FROM employees GROUP BY current_status",conn)
    conn.close()
    return dict(zip(df['current_status'],df['c']))

@st.cache_data(ttl=20)
def get_division_list():
    conn=get_conn()
    df=pd.read_sql_query("SELECT DISTINCT division FROM employees WHERE division IS NOT NULL ORDER BY division",conn)
    conn.close()
    base=["Catering","MRO","Appearance","Ramp","Cargo"]
    extra=[d for d in df['division'].tolist() if d not in base]
    return base+extra

@st.cache_data(ttl=20)
def get_cost_centers(division=None):
    conn=get_conn()
    if division and division!="All":
        df=pd.read_sql_query("SELECT * FROM cost_centers WHERE division=? AND is_active=1 ORDER BY code",conn,params=(division,))
    else:
        df=pd.read_sql_query("SELECT * FROM cost_centers WHERE is_active=1 ORDER BY division,code",conn)
    conn.close()
    return df

@st.cache_data(ttl=20)
def get_emp_list_cached():
    conn=get_conn()
    df=pd.read_sql_query("SELECT emp_id,full_name,division,cost_center,current_status FROM employees ORDER BY emp_id LIMIT 5000",conn)
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
    total=pd.read_sql_query(q,conn,params=p).iloc[0]['c']; conn.close(); return int(total)

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
    df=pd.read_sql_query(q,conn,params=p); conn.close()
    # Replace database NULLs with a clean placeholder instead of showing literal "None"
    df = df.fillna("—")
    for col in df.columns:
        df[col] = df[col].apply(lambda v: "—" if v=="None" or v is None else v)
    return df

@st.cache_data(ttl=10)
def get_employee(eid):
    conn=get_conn(); cur=conn.cursor()
    cur.execute("SELECT * FROM employees WHERE emp_id=?",(eid,))
    cols=[d[0] for d in cur.description]; row=cur.fetchone(); conn.close()
    return dict(zip(cols,row)) if row else None

@st.cache_data(ttl=15)
def get_active_employee_payroll_summary():
    """All active employees (incl. on leave types) with their current month financial summary."""
    conn=get_conn()
    df=pd.read_sql_query("""SELECT emp_id,full_name,division,cost_center,basic_salary,current_status
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
# SEED DATA — Professional sample dataset (only runs once)
# ════════════════════════════════════════════════════════
def seed_if_empty():
    conn=get_conn(); c=conn.cursor()
    c.execute("SELECT COUNT(*) FROM employees")
    if c.fetchone()[0]>0:
        conn.close(); return
    F=["Abebe","Chaltu","Almaz","Yohannes","Bekele","Makeda","Tariku","Saba","Dawit","Fatuma","Selam","Elias","Hana","Tigist"]
    FA=["Kebede","Chala","Tadesse","Tekle","Bogale","Mengistu","Assefa","Haile","Gezahagn","Zewdu"]
    G=["Balcha","Tafa","Alemu","Desta","Kassa","Fikru","Tolossa","Girma","Woldemariam"]
    DIVS=["Catering","MRO","Appearance","Ramp","Cargo"]
    CC_MAP={"Catering":["CC-CAT-01","CC-CAT-02"],"MRO":["CC-MRO-01","CC-MRO-02"],
            "Appearance":["CC-APP-01"],"Ramp":["CC-RMP-01","CC-RMP-02"],"Cargo":["CC-CGO-01"]}
    E=["High School Graduate","TVET Diploma","BSc/BA Degree","MSc/MA Post-Graduate"]
    J=["Service Agent","Ground Handler","Cargo Operator","Ramp Supervisor","Catering Staff","MRO Technician"]
    B=["CBE","Awash Bank","Abyssinia Bank","Dashen Bank"]
    BL=["A+","A-","B+","B-","O+","O-","AB+","AB-"]
    SC=["Bole","Yeka","Kirkos","Nifas Silk","Akaky Kaliti","Lideta","Gulele","Kolfe"]
    RL=["Orthodox","Muslim","Protestant","Catholic"]
    MR=["Single","Married","Divorced","Widowed"]
    ET=["Permanent","Contract","Temporary"]
    WD=["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"]
    dists={"Active Deployment":720,"Pending Screening":200,"Terminated":50,"Pre-Employment Process":20,"On Leave":10}
    recs,gc=[],1
    for status,count in dists.items():
        for _ in range(count):
            eid=f"YGS-{1000+gc}"; name=f"{random.choice(F)} {random.choice(FA)} {random.choice(G)}"
            sal=round(random.uniform(4500,18000),2)
            div=random.choice(DIVS)
            cc=random.choice(CC_MAP[div])
            recs.append((eid,name,div,cc,f"+2519{random.randint(10000000,99999999)}",
                f"{name.split()[0].lower()}{gc}@gsp.local",f"H-No {random.randint(100,999)}",
                f"Woreda {random.randint(1,12)}",random.choice(SC),f"0{random.randint(1,9)}",
                f"RES-{random.randint(100000,999999)}","Addis Ababa",random.randint(20,55),
                random.choice(["Male","Female"]),random.choice(MR),"Ethiopian",random.choice(RL),
                random.choice(F),f"+2519{random.randint(10000000,99999999)}",random.choice(BL),
                f"TIN-{random.randint(1000000,9999999)}",f"PEN-{random.randint(100000,999999)}",
                random.choice(B),f"{random.randint(1000000000,9999999999)}",
                random.choice(E),"Operational Logistics",str(random.randint(2005,2022)),
                random.choice(["AAU","AASTU","Hawassa Univ","Mekele Univ"]),
                status,random.choice(J),random.choice(ET),random.choice(WD),
                f"2024-{random.randint(1,12):02d}-{random.randint(1,28):02d}",
                f"2027-{random.randint(1,12):02d}-{random.randint(1,28):02d}",sal,
                f"2026-05-{random.randint(10,28)}",""))
            gc+=1
    c.executemany("""INSERT INTO employees(emp_id,full_name,division,cost_center,contact,email,house_address,woreda,
        subcity,kebele,resident_id,place_of_birth,age,sex,marital_status,nationality,religion,
        emergency_contact_name,emergency_contact_phone,blood_type,tin_number,pension_number,
        bank_name,bank_account,edu_background,field_of_graduate,graduation_year,institution_name,
        current_status,job_title,employment_type,weekly_dayoff,start_date,contract_end_date,basic_salary,
        registration_date,notes)VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",recs)

    # Seed cost centers
    now=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cc_recs=[]
    for div,ccs in CC_MAP.items():
        for cc in ccs:
            cc_recs.append((cc,f"{div} Operations — {cc}",div,round(random.uniform(200000,800000),2),f"Primary cost center for {div} division",1,"system",now))
    c.executemany("INSERT OR IGNORE INTO cost_centers(code,name,division,budget,description,is_active,created_by,created_at)VALUES(?,?,?,?,?,?,?,?)",cc_recs)

    conn.commit(); conn.close()

seed_if_empty()

# ════════════════════════════════════════════════════════
# CSS — Aurora Night Design System
# ════════════════════════════════════════════════════════
st.markdown("""<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&family=Cinzel:wght@700&display=swap');
*{box-sizing:border-box}
.stApp{background:#060B18 !important;color:#E8EEF7 !important;font-family:'Inter',sans-serif !important}
.yh{background:linear-gradient(135deg,#0D1526,#0A1020,#0D1830);border-bottom:1px solid rgba(212,168,71,0.25);
  padding:14px 28px;margin:-1rem -1rem 0;position:relative}
.yh::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;
  background:linear-gradient(90deg,transparent,#D4A847,#F0C96B,#D4A847,transparent)}
.hb{font-family:'Cinzel',serif !important;font-size:clamp(14px,2vw,22px) !important;
  font-weight:700 !important;color:#F0C96B !important;letter-spacing:.06em !important;margin:0 !important}
.ht{font-size:10px;color:#6B7FA3;letter-spacing:.08em;text-transform:uppercase;margin-top:2px}
.cs{display:flex;background:#0D1526;border:1px solid rgba(212,168,71,0.15);border-radius:8px;overflow:hidden;margin:10px 0}
.ci{flex:1;padding:7px 12px;font-size:11px;color:#94A8C8;border-right:1px solid rgba(255,255,255,0.06)}
.ci:last-child{border-right:none}
.card{background:#0D1526;border:1px solid rgba(255,255,255,0.07);border-radius:12px;
  padding:16px;margin-bottom:12px;position:relative;overflow:hidden}
.card::before{content:'';position:absolute;top:0;left:0;right:0;height:1px;
  background:linear-gradient(90deg,transparent,rgba(212,168,71,0.3),transparent)}
.card-gold{background:linear-gradient(135deg,#0D1526,#12192E);border:1px solid rgba(212,168,71,0.3)}
.ey{font-size:9px;font-weight:600;letter-spacing:.16em;text-transform:uppercase;color:#D4A847;
  margin-bottom:4px;display:flex;align-items:center;gap:6px}
.ey::before{content:'';display:inline-block;width:14px;height:1px;background:#D4A847}
.tl{font-family:'Cinzel',serif;font-size:clamp(15px,2vw,24px);font-weight:700;color:#F0C96B;margin-bottom:5px}
.mg{display:grid;grid-template-columns:repeat(7,1fr);gap:7px;margin-bottom:16px}
.mb{background:#0D1526;border:1px solid rgba(255,255,255,0.07);border-radius:10px;
  padding:11px 8px;text-align:center;position:relative;overflow:hidden}
.mb::after{content:'';position:absolute;bottom:0;left:0;right:0;height:2px}
.mg-gold::after{background:linear-gradient(90deg,#D4A847,#F0C96B)}.mg-green::after{background:#10B981}
.mg-cyan::after{background:#38BDF8}.mg-amber::after{background:#F59E0B}
.mg-red::after{background:#EF4444}.mg-purple::after{background:#A855F7}.mg-teal::after{background:#14B8A6}
.ml{font-size:8px;font-weight:600;letter-spacing:.08em;text-transform:uppercase;margin-bottom:3px}
.ml-gold{color:#D4A847}.ml-green{color:#10B981}.ml-cyan{color:#38BDF8}
.ml-amber{color:#F59E0B}.ml-red{color:#EF4444}.ml-purple{color:#A855F7}.ml-teal{color:#14B8A6}
.mv{font-family:'Cinzel',serif;font-size:18px;font-weight:700;color:#E8EEF7;line-height:1}
div.stButton>button,div.stFormSubmitButton>button,button[kind="formSubmit"],button[kind="secondaryFormSubmit"],button[kind="primaryFormSubmit"]{
  background:linear-gradient(135deg,#1A6B3C,#22C55E) !important;color:#fff !important;
  border:none !important;border-radius:7px !important;padding:7px 14px !important;
  font-weight:600 !important;font-size:12px !important;transition:all .15s !important}
div.stButton>button:hover,div.stFormSubmitButton>button:hover,button[kind="formSubmit"]:hover{
  box-shadow:0 4px 16px rgba(34,197,94,0.4) !important;transform:translateY(-1px) !important}
div.stDownloadButton>button{background:linear-gradient(135deg,#0369A1,#0284C7) !important;
  color:#fff !important;border:none !important;border-radius:7px !important;font-weight:600 !important}
.stTextInput>div>div>input,.stNumberInput>div>div>input{background:#131F38 !important;
  border:1px solid rgba(255,255,255,0.1) !important;border-radius:7px !important;color:#E8EEF7 !important}
.stTextArea textarea{background:#131F38 !important;border:1px solid rgba(255,255,255,0.1) !important;
  border-radius:7px !important;color:#E8EEF7 !important}
.stSelectbox>div>div{background:#131F38 !important;border:1px solid rgba(255,255,255,0.1) !important;
  border-radius:7px !important;color:#E8EEF7 !important}
label{color:#94A8C8 !important;font-size:11px !important;font-weight:500 !important}
.stTabs [data-baseweb="tab-list"]{background:#0D1526 !important;border-radius:8px !important;
  padding:3px !important;border:1px solid rgba(255,255,255,0.07) !important;gap:2px !important}
.stTabs [data-baseweb="tab"]{border-radius:6px !important;color:#6B7FA3 !important;
  font-weight:500 !important;font-size:11px !important;padding:6px 12px !important}
.stTabs [aria-selected="true"]{background:linear-gradient(135deg,#1A4B6B,#0284C7) !important;color:#fff !important}
.stTabs [data-baseweb="tab-panel"]{background:#0D1526 !important;border:1px solid rgba(255,255,255,0.07) !important;
  border-radius:10px !important;padding:14px !important;margin-top:5px !important}
.stDataFrame{border-radius:9px !important;overflow:hidden !important}
hr{border:none !important;border-top:1px solid rgba(255,255,255,0.07) !important;margin:12px 0 !important}
.stFileUploader>div{background:#131F38 !important;border:1px dashed rgba(212,168,71,0.3) !important;border-radius:9px !important}
.sb{background:rgba(16,185,129,0.1);border:1px solid rgba(16,185,129,0.3);border-radius:8px;
  padding:7px 10px;text-align:center;font-size:11px;color:#6EE7B7;margin-bottom:6px}
.db{display:inline-flex;align-items:center;gap:3px;font-size:10px;padding:2px 8px;
  border-radius:10px;font-weight:500;margin:2px}
.db-up{background:rgba(16,185,129,0.15);color:#34D399;border:1px solid rgba(16,185,129,0.3)}
.db-mis{background:rgba(239,68,68,0.1);color:#FCA5A5;border:1px solid rgba(239,68,68,0.2)}
.ps{background:linear-gradient(135deg,#0D1526,#0A1A10);border:1px solid rgba(16,185,129,0.25);
  border-radius:12px;padding:14px 18px;margin-top:12px}
.pr{display:flex;justify-content:space-between;align-items:center;padding:4px 0;
  border-bottom:1px solid rgba(255,255,255,0.05);font-size:12px}
.pl{color:#94A8C8}.pv{color:#E8EEF7;font-weight:500}
.pd{color:#FCA5A5}.pn{color:#34D399;font-family:'Cinzel',serif;font-size:16px;font-weight:700}
.ic{background:linear-gradient(135deg,#0D1526,#0A1830);border:1px solid rgba(212,168,71,0.35);
  border-radius:14px;padding:20px;position:relative;overflow:hidden}
.ic::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;
  background:linear-gradient(90deg,#D4A847,#F0C96B,#D4A847)}
.ifl{font-size:9px;color:#6B7FA3;letter-spacing:.07em;text-transform:uppercase;margin-bottom:1px}
.ifv{font-size:12px;color:#E8EEF7;font-weight:500}
.ifv-g{color:#F0C96B;font-family:'Cinzel',serif;font-size:14px;font-weight:700}
.sp{display:inline-block;padding:2px 10px;border-radius:20px;font-size:10px;font-weight:600;letter-spacing:.03em}
.sa{background:rgba(16,185,129,0.15);color:#34D399;border:1px solid rgba(16,185,129,0.3)}
.spe{background:rgba(56,189,248,0.1);color:#7DD3FC;border:1px solid rgba(56,189,248,0.2)}
.spr{background:rgba(245,158,11,0.1);color:#FCD34D;border:1px solid rgba(245,158,11,0.2)}
.sl{background:rgba(168,85,247,0.1);color:#C4B5FD;border:1px solid rgba(168,85,247,0.2)}
.st{background:rgba(239,68,68,0.1);color:#FCA5A5;border:1px solid rgba(239,68,68,0.2)}
.pb{background:#0A0F1E;border:1px solid rgba(212,168,71,0.15);border-radius:10px;padding:10px;margin-top:6px}
.gt{font-family:'Cinzel',serif;color:#F0C96B !important;font-size:17px;font-weight:700;margin-bottom:10px}
.gl{color:#D4A847 !important;font-weight:600;font-size:10px;margin-top:7px;margin-bottom:2px;
  letter-spacing:.07em;text-transform:uppercase}
.fs{font-size:9px;font-weight:600;color:#D4A847;letter-spacing:.09em;text-transform:uppercase;
  margin:9px 0 5px;display:flex;align-items:center;gap:6px}
.fs::before{content:'';display:inline-block;width:10px;height:1px;background:#D4A847}
.hch{display:inline-block;padding:2px 8px;border-radius:20px;font-size:10px;font-weight:500;
  margin:2px;background:rgba(212,168,71,0.12);color:#F0C96B;border:1px solid rgba(212,168,71,0.25)}
.scan{background:#0A1020;border:2px dashed rgba(56,189,248,0.3);border-radius:10px;
  padding:16px;text-align:center;margin-bottom:10px}
.gate-banner{background:linear-gradient(135deg,#1A0A0A,#0D1526);border:1px solid rgba(239,68,68,0.3);
  border-radius:12px;padding:14px 18px;margin-bottom:14px;display:flex;align-items:center;gap:12px}
.gate-open{background:linear-gradient(135deg,#0A1A10,#0D1526);border:1px solid rgba(16,185,129,0.3)}
.cc-tag{display:inline-block;padding:2px 9px;border-radius:8px;font-size:10px;font-weight:600;
  background:rgba(168,85,247,0.12);color:#C4B5FD;border:1px solid rgba(168,85,247,0.25);margin-left:6px}
::-webkit-scrollbar{width:4px;height:4px}
::-webkit-scrollbar-track{background:#060B18}
::-webkit-scrollbar-thumb{background:rgba(212,168,71,0.3);border-radius:2px}
</style>""",unsafe_allow_html=True)

# ════════════════════════════════════════════════════════
# SESSION STATE
# ════════════════════════════════════════════════════════
if '572' not in st.session_state:
    st.session_state['572']={'ur':{'ygs_manager':{'pw':'secure2026','role':'Manager'},'ygs_officer':{'pw':'data2026','role':'Data Officer'}}}
for k,v in {"role":None,"uid":None,"view":"Home","eid":None,"full_name":"","assigned_division":None,"nav_access_json":None}.items():
    if k not in st.session_state: st.session_state[k]=v

# ════════════════════════════════════════════════════════
# HEADER
# ════════════════════════════════════════════════════════
h1,h2=st.columns([3.5,1])
with h1:
    welcome_text = f"Welcome, {st.session_state.full_name}" if st.session_state.role else "Welcome"
    st.markdown(f"""<div class="yh">
      <div style="font-size:11px;color:#10B981;letter-spacing:.08em;text-transform:uppercase;margin-bottom:4px">{welcome_text}</div>
      <div class="hb">YETEBABERUT · GENERAL SERVICE PROVIDER</div>
      <div class="ht">Human Resource Management System | Addis Ababa, Ethiopia | Est. 2015</div>
    </div>""",unsafe_allow_html=True)
with h2:
    st.write(""); st.write("")
    if not st.session_state.role:
        if st.button("LOGIN",use_container_width=True): st.session_state.show_login=True
        @st.dialog("Portal Login")
        def login():
            st.markdown("""<style>
            div[data-testid="stDialog"] input{
                background:#FFFFFF !important;color:#0D1526 !important;
                border:1px solid #D4A847 !important;border-radius:8px !important}
            div[data-testid="stDialog"] label{color:#10B981 !important;font-weight:600 !important}
            div[data-testid="stDialog"] button[kind="formSubmit"],
            div[data-testid="stDialog"] div.stFormSubmitButton button{
                background:linear-gradient(135deg,#1A6B3C,#22C55E) !important;
                color:#FFFFFF !important;border:none !important}
            </style>""",unsafe_allow_html=True)
            st.markdown('<div class="gt" style="color:#10B981 !important">Yetebaberut Login</div>',unsafe_allow_html=True)
            with st.form("lf"):
                st.markdown('<div class="gl" style="color:#10B981 !important">User ID</div>',unsafe_allow_html=True)
                u=st.text_input("u",label_visibility="collapsed")
                st.markdown('<div class="gl" style="color:#10B981 !important">Password</div>',unsafe_allow_html=True)
                p=st.text_input("p",type="password",label_visibility="collapsed")
                if st.form_submit_button("Enter",use_container_width=True):
                    conn_l=get_conn(); cur_l=conn_l.cursor()
                    cur_l.execute("SELECT username,password,role,full_name,is_active,assigned_division,nav_access FROM system_users WHERE username=?",(u,))
                    db_user=cur_l.fetchone()
                    if db_user and db_user[1]==p and db_user[4]==1:
                        st.session_state.role=db_user[2]; st.session_state.uid=db_user[0]
                        st.session_state.full_name=db_user[3] or db_user[0]
                        st.session_state.assigned_division=db_user[5]
                        st.session_state.nav_access_json=db_user[6]
                        cur_l.execute("UPDATE system_users SET last_login=? WHERE username=?",(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),u,))
                        conn_l.commit(); conn_l.close()
                        st.session_state.view="Applicant Intake"; st.toast(f"Welcome {st.session_state.full_name}!"); st.rerun()
                    else:
                        conn_l.close()
                        ur=st.session_state['572']['ur']
                        if u in ur and ur[u]['pw']==p:
                            st.session_state.role=ur[u]['role']; st.session_state.uid=u
                            st.session_state.full_name=u
                            st.session_state.view="Applicant Intake"; st.toast(f"Welcome {st.session_state.role}!"); st.rerun()
                        else: st.error("Invalid credentials or account disabled.")
        if st.session_state.get("show_login"): st.session_state.show_login=False; login()
    else:
        st.markdown(f'<div class="sb"> <b>{st.session_state.role}</b><br><span style="font-size:10px;color:#6B7FA3">{st.session_state.uid}</span></div>',unsafe_allow_html=True)
        if st.button("LOGOUT",use_container_width=True):
            st.session_state.role=None; st.session_state.uid=None
            st.session_state.eid=None; st.session_state.view="Home"
            st.session_state.nav_access_json=None; st.session_state.assigned_division=None
            st.rerun()

st.markdown("""<div class="cs">
  <div class="ci"> Addis Ababa, Ethiopia</div>
  <div class="ci"> info@yetebaberut.com</div>
  <div class="ci"> +251 911 000 000</div>
  <div class="ci"> @YGSP_Global</div>
</div>""",unsafe_allow_html=True)
st.markdown("<hr>",unsafe_allow_html=True)

# ════════════════════════════════════════════════════════
# NAVIGATION ACCESS DEFINITIONS
# Master list of every navigable view. A Manager can grant
# a custom subset (with view/edit/delete permission per
# view) to any user created in Administration. If a user
# has no custom nav_access saved, the role default applies.
# ════════════════════════════════════════════════════════
ALL_NAV_VIEWS = ["Home","Applicant Intake","Employee Directory","Employee Profile",
    "Supervisor Console","Payroll","Payroll Approvals","Leave & Discipline",
    "Public Holidays","Cost Centers","Recycle Bin","Administration"]

ROLE_DEFAULT_VIEWS = {
    "Supervisor": ["Home","Supervisor Console","Public Holidays"],
    "Payroll Section": ["Home","Payroll Approvals","Payroll","Employee Directory","Employee Profile","Public Holidays","Cost Centers"],
    "Manager": ["Home","Applicant Intake","Employee Directory","Employee Profile","Payroll",
        "Payroll Approvals","Leave & Discipline","Public Holidays","Cost Centers","Recycle Bin","Administration"],
}
ROLE_DEFAULT_FALLBACK = ["Home","Applicant Intake","Employee Directory","Employee Profile",
    "Payroll","Payroll Approvals","Leave & Discipline","Public Holidays","Cost Centers"]

def get_user_nav_views(role, nav_access_json):
    """Returns the ordered list of views this user can see."""
    if nav_access_json:
        try:
            access = json.loads(nav_access_json)
            views = [v for v in ALL_NAV_VIEWS if v in access]
            if views: return views
        except: pass
    return ROLE_DEFAULT_VIEWS.get(role, ROLE_DEFAULT_FALLBACK)

def get_user_view_permission(view_name, role, nav_access_json):
    """Returns one of 'view','edit','both','full_control' for a given view.
    'view' = read-only. 'edit' = can change data, no delete.
    'both' = view and edit. 'full_control' = view, edit and delete."""
    if nav_access_json:
        try:
            access = json.loads(nav_access_json)
            if view_name in access:
                return access[view_name]
        except: pass
    return "full_control" if role=="Manager" else "view"

def can_edit(view_name, role, nav_access_json):
    return get_user_view_permission(view_name, role, nav_access_json) in ("edit","both","full_control")

def can_delete(view_name, role, nav_access_json):
    return get_user_view_permission(view_name, role, nav_access_json) == "full_control"

# ════════════════════════════════════════════════════════
# NAVIGATION — grouped vertical sidebar, left side
# ════════════════════════════════════════════════════════
if st.session_state.role:
    ROLE = st.session_state.role
    VIEWS = get_user_nav_views(ROLE, st.session_state.get("nav_access_json"))
else:
    st.session_state.view="Home"
    VIEWS=["Home"]

V=st.session_state.view

# Group views under labeled sections for fast scanning at scale.
NAV_GROUPS = [
    ("OVERVIEW", ["Home"]),
    ("RECRUITMENT", ["Applicant Intake"]),
    ("WORKFORCE", ["Employee Directory","Employee Profile","Supervisor Console"]),
    ("FINANCE", ["Payroll","Payroll Approvals","Cost Centers"]),
    ("HR OPERATIONS", ["Leave & Discipline"]),
    ("REFERENCE", ["Public Holidays"]),
    ("SYSTEM", ["Recycle Bin","Administration"]),
]

st.markdown("""<style>
.nav-sidebar{padding-right:4px}
.nav-group-label{font-size:8px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;
    color:#6B7FA3;margin:10px 0 4px 4px}
.nav-group-label:first-child{margin-top:0}
.nav-sidebar div.stButton button{
    width:100% !important; text-align:left !important; justify-content:flex-start !important;
    margin-bottom:2px !important; padding:5px 10px !important; min-height:30px !important;
    background:#0D1526 !important; color:#94A8C8 !important; font-size:11px !important;
    border:1px solid rgba(255,255,255,0.06) !important;
    box-shadow:none !important; font-weight:500 !important; border-radius:6px !important;}
.nav-sidebar div.stButton button:hover{
    background:#131F38 !important; color:#E8EEF7 !important;
    border-color:rgba(212,168,71,0.3) !important;}
.nav-sidebar .nav-active button{
    background:linear-gradient(135deg,#1A6B3C,#22C55E) !important;
    color:#fff !important; font-weight:600 !important;
    border-color:transparent !important;}
.nav-footer{margin-top:10px;padding:6px 8px;border-top:1px solid rgba(255,255,255,0.07);
    font-size:9px;color:#6B7FA3;text-align:center}
.hamburger-bar{display:flex;align-items:center;background:#0D1526;
    border-bottom:1px solid rgba(212,168,71,0.2);padding:6px 14px;
    margin:-1rem -1rem 0 -1rem;gap:12px}
.hamburger-bar .hb-title{font-size:11px;color:#6B7FA3;letter-spacing:.06em;
    text-transform:uppercase;font-weight:600}
div[data-testid="stHorizontalBlock"] > div:first-child div.stButton button[kind="secondary"]{
    background:#0D1526 !important;border:1px solid rgba(212,168,71,0.35) !important;
    color:#D4A847 !important;font-size:16px !important;font-weight:700 !important;
    padding:2px 10px !important;min-height:34px !important;min-width:42px !important;
    border-radius:6px !important;line-height:1 !important;}
</style>""",unsafe_allow_html=True)

if "sidebar_open" not in st.session_state: st.session_state.sidebar_open = True

# ── Hamburger toggle sits in a dedicated bar at the very top, before the columns split ──
if st.session_state.role:
    hb_col, hb_label = st.columns([0.08, 1])
    with hb_col:
        toggle_icon = "☰"
        if st.button(toggle_icon, key="nav_toggle", help="Toggle navigation menu"):
            st.session_state.sidebar_open = not st.session_state.sidebar_open
            st.rerun()
    with hb_label:
        current_view_label = st.session_state.view or "Home"
        st.markdown(f'<div class="hb-title" style="padding-top:6px">Navigation &nbsp;|&nbsp; <span style="color:#F0C96B">{current_view_label}</span></div>', unsafe_allow_html=True)

if st.session_state.role and st.session_state.sidebar_open:
    nav_col, main_col = st.columns([1, 5.5])
elif st.session_state.role and not st.session_state.sidebar_open:
    nav_col, main_col = None, st.container()
else:
    nav_col, main_col = None, st.container()

if st.session_state.role and st.session_state.sidebar_open:
    with nav_col:
        st.markdown('<div class="nav-sidebar">', unsafe_allow_html=True)
        for group_label, group_views in NAV_GROUPS:
            visible_in_group = [v for v in group_views if v in VIEWS]
            if not visible_in_group: continue
            st.markdown(f'<div class="nav-group-label">{group_label}</div>', unsafe_allow_html=True)
            for v in visible_in_group:
                is_active = (st.session_state.view == v)
                if is_active: st.markdown('<div class="nav-active">', unsafe_allow_html=True)
                if st.button(v, use_container_width=True, key=f"nav_{v}"):
                    st.session_state.view=v; st.rerun()
                if is_active: st.markdown('</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="nav-footer">Yetebaberut HRMS<br>v2.0</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

main_block = main_col.container() if st.session_state.role else main_col

def sclass(s):
    return {"Active Deployment":"sa","Pending Screening":"spe","Pre-Employment Process":"spr","On Leave":"sl","Terminated":"st"}.get(s,"spe")

# ════════════════════════════════════════════════════════
# HOME — with Applicant Gate Control
# ════════════════════════════════════════════════════════
with main_block:
    if V=="Home":
        if st.session_state.role:
            st.markdown('<div class="ey">Live Status — Today</div>',unsafe_allow_html=True)
            st.markdown('<div class="tl">Workforce Attendance Snapshot</div>',unsafe_allow_html=True)
            today_str = date.today().isoformat()
            today_weekday = date.today().strftime("%A")
            today_holidays = get_holidays(date.today().year)
            is_holiday_today = date.today() in today_holidays

            seg1,seg2 = st.columns([1,1])
            with seg1:
                seg_mode = st.selectbox("Segregate by",["All Employees","By Division","By Cost Center"],key="live_seg_mode")
            with seg2:
                seg_value="All"
                if seg_mode=="By Division":
                    seg_value=st.selectbox("Division",get_division_list(),key="live_seg_div")
                elif seg_mode=="By Cost Center":
                    all_cc_live=get_cost_centers()
                    cc_opts_live=all_cc_live['code'].tolist() if len(all_cc_live)>0 else []
                    if cc_opts_live: seg_value=st.selectbox("Cost Center",cc_opts_live,key="live_seg_cc")
                    else: st.info("No cost centers exist yet.")

            conn=get_conn()
            base_filter=""; base_param=[]
            if seg_mode=="By Division" and seg_value!="All":
                base_filter=" AND division=?"; base_param=[seg_value]
            elif seg_mode=="By Cost Center" and seg_value!="All":
                base_filter=" AND cost_center=?"; base_param=[seg_value]

            total_in_scope=pd.read_sql_query(f"SELECT COUNT(*) as c FROM employees WHERE current_status='Active Deployment'{base_filter}",conn,params=base_param).iloc[0]['c']
            on_leave_today=pd.read_sql_query(f"""SELECT COUNT(DISTINCT lr.emp_id) as c FROM leave_records lr
                JOIN employees e ON lr.emp_id=e.emp_id
                WHERE lr.status='Approved' AND lr.start_date<=? AND lr.end_date>=?{base_filter}""",
                conn,params=[today_str,today_str]+base_param).iloc[0]['c']
            absent_today=pd.read_sql_query(f"""SELECT COUNT(DISTINCT ar.emp_id) as c FROM absent_records ar
                JOIN employees e ON ar.emp_id=e.emp_id
                WHERE ar.absent_date=? AND COALESCE(ar.record_status,'Active')='Active'{base_filter}""",
                conn,params=[today_str]+base_param).iloc[0]['c']
            dayoff_today=pd.read_sql_query(f"""SELECT COUNT(*) as c FROM employees
                WHERE current_status='Active Deployment' AND weekly_dayoff=?{base_filter}""",
                conn,params=[today_weekday]+base_param).iloc[0]['c']
            conn.close()

            if is_holiday_today:
                dayoff_today = total_in_scope

            on_work_today = max(total_in_scope - on_leave_today - absent_today - dayoff_today, 0)

            holiday_note = f" — {today_holidays.get(date.today(),'')}" if is_holiday_today else ""
            st.markdown(f'<div style="font-size:11px;color:#6B7FA3;margin-bottom:8px">{today_str} ({today_weekday}){holiday_note} — scope: <b style="color:#F0C96B">{seg_value if seg_value!="All" else "All Employees"}</b></div>',unsafe_allow_html=True)

            st.markdown(f"""<div class="mg" style="grid-template-columns:repeat(5,1fr)">
              <div class="mb mg-green"><div class="ml ml-green">On Work Today</div><div class="mv">{on_work_today}</div></div>
              <div class="mb mg-purple"><div class="ml ml-purple">On Leave Today</div><div class="mv">{on_leave_today}</div></div>
              <div class="mb mg-red"><div class="ml ml-red">Absent Today</div><div class="mv">{absent_today}</div></div>
              <div class="mb mg-cyan"><div class="ml ml-cyan">Day-Off Today</div><div class="mv">{dayoff_today}</div></div>
              <div class="mb mg-gold"><div class="ml ml-gold">Total in Scope</div><div class="mv">{total_in_scope}</div></div>
            </div>""",unsafe_allow_html=True)
            st.markdown("<hr>",unsafe_allow_html=True)

        st.markdown('<div class="ey">Who We Are</div>',unsafe_allow_html=True)
        st.markdown('<div class="tl">Engineering Workforce Excellence</div>',unsafe_allow_html=True)
        st.markdown("""<div class="card card-gold"><p style="font-size:14px;line-height:1.85;color:#C8D8F0;margin:0">
        Yetebaberut General Service Provider specializes in engineering high-capacity workforce deployment models.
        We eliminate operational downtime for tier-one organizations by sourcing, verifying, and routing vetted manpower
        assets across key regional stations and global industrial hubs.</p></div>""",unsafe_allow_html=True)
        c1,c2,c3=st.columns(3)
        for col,ic,ttl,col_hex,txt in [
            (c1,"","Our Mission","#60C8F8","To seamlessly bridge the gap between heavy industrial demand and verified talent arrays through real-time matching pipelines."),
            (c2,"","Our Vision","#34D399","To set the gold standard for global service provisioning, building stable talent supply infrastructures that power Africa's economy."),
            (c3,"","Core Values","#F0C96B","<ul style='padding-left:14px;line-height:1.7'><li><b style='color:#F0C96B'>Integrity:</b> Full verification protocols.</li><li><b style='color:#F0C96B'>Speed:</b> Zero grid lag dispatch.</li><li><b style='color:#F0C96B'>Compliance:</b> Total risk containment.</li></ul>")]:
            with col: st.markdown(f"""<div class="card" style="border-top:2px solid {col_hex};height:100%">
              <span style="font-size:22px;display:block;margin-bottom:8px">{ic}</span>
              <div style="font-family:'Cinzel',serif;color:{col_hex};font-size:14px;margin-bottom:8px">{ttl}</div>
              <div style="font-size:12px;color:#94A8C8;line-height:1.7">{txt}</div></div>""",unsafe_allow_html=True)
        st.write("")

        # ── APPLICANT GATE CONTROL ──
        applications_open = get_setting("applications_open","1")=="1"

        bc,ac=st.columns([2,1])
        with bc:
            if applications_open:
                st.info("All new applications are routed to intake verification. Processing: 3–5 working days.")
            else:
                st.warning("Applications are currently **closed**. All vacancies are filled. Please check back later.")
        with ac:
            st.write("")
            if applications_open:
                if st.button("APPLY NOW",use_container_width=True): st.session_state.show_apply=True
            else:
                st.button("Applications Closed",use_container_width=True,disabled=True)

            if applications_open:
                @st.dialog("Application Form",width="large")
                def apply_dlg():
                    st.markdown('<div class="gt">Personnel Intake Registration</div>',unsafe_allow_html=True)
                    with st.form("af",clear_on_submit=True):
                        st.markdown('<div class="gl">Personal</div>',unsafe_allow_html=True)
                        a1,a2,a3=st.columns(3)
                        with a1: fn=st.text_input("First Name")
                        with a2: fa=st.text_input("Father Name")
                        with a3: gf=st.text_input("Grand Father")
                        a4,a5,a6=st.columns(3)
                        with a4: ph=st.text_input("Phone")
                        with a5: em=st.text_input("Email")
                        with a6: sx=st.selectbox("Sex",["Male","Female"])
                        st.markdown('<div class="gl">Address</div>',unsafe_allow_html=True)
                        a7,a8,a9,a10=st.columns(4)
                        with a7: ha=st.text_input("House No.")
                        with a8: wo=st.text_input("Woreda")
                        with a9: sc=st.text_input("Subcity")
                        with a10: ke=st.text_input("Kebele")
                        a11,a12,a13=st.columns(3)
                        with a11: ri=st.text_input("Resident ID")
                        with a12: pb=st.text_input("Place of Birth")
                        with a13: ag=st.number_input("Age",18,80,step=1)
                        st.markdown('<div class="gl">Education</div>',unsafe_allow_html=True)
                        a14,a15=st.columns(2)
                        with a14: ed=st.selectbox("Level",["High School Graduate","TVET Diploma","BSc/BA Degree","MSc/MA Post-Graduate"])
                        with a15: fi=st.text_input("Field of Study")
                        up=st.file_uploader("Upload Documents (PDF/ZIP/JPG)",type=["pdf","zip","rar","jpg","png"])
                        if st.form_submit_button("Submit",use_container_width=True):
                            if not(fn and fa and ph and ri and up): st.error("Fill all mandatory fields.")
                            else:
                                name=f"{fn} {fa} {gf}".strip(); fb=up.read()
                                conn=get_conn(); cur=conn.cursor()
                                cur.execute("SELECT COUNT(*) FROM employees"); cnt=cur.fetchone()[0]
                                nid=f"YGS-{2000+cnt}"; div=random.choice(["Catering","MRO","Appearance","Ramp","Cargo"])
                                conn.execute("""INSERT INTO employees(emp_id,full_name,division,contact,email,house_address,
                                    woreda,subcity,kebele,resident_id,place_of_birth,age,sex,edu_background,field_of_graduate,
                                    current_status,registration_date,edu_doc_name,edu_doc_data)
                                    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,'Pending Screening',?,?,?)""",
                                    (nid,name,div,ph,em,ha,wo,sc,ke,ri,pb,int(ag),sx,ed,fi,
                                     datetime.now().strftime("%Y-%m-%d"),up.name,sqlite3.Binary(fb)))
                                conn.commit(); conn.close()
                                get_emp_list_cached.clear(); get_stats.clear()
                                st.success(f"Application received! ID: **{nid}**  {div}"); st.balloons()
                if st.session_state.get("show_apply"): st.session_state.show_apply=False; apply_dlg()

    # ════════════════════════════════════════════════════════
    # APPLICANT INTAKE
    # ════════════════════════════════════════════════════════
    elif V=="Applicant Intake":
        st.markdown('<div class="ey">HR Operations</div>',unsafe_allow_html=True)
        st.markdown('<div class="tl">Applicant Intake Screening</div>',unsafe_allow_html=True)

        if st.session_state.role=="Manager":
            applications_open = get_setting("applications_open","1")=="1"
            gate_cls = "gate-open" if applications_open else ""
            gate_icon = "" if applications_open else ""
            gate_text = "Applications are OPEN — candidates can submit new applications" if applications_open else "Applications are CLOSED — public Apply button is hidden"
            st.markdown(f"""<div class="gate-banner {gate_cls}">
              <div style="font-size:24px">{gate_icon}</div>
              <div style="flex:1">
                <div style="font-size:13px;font-weight:600;color:{'#34D399' if applications_open else '#FCA5A5'}">{gate_text}</div>
                <div style="font-size:11px;color:#6B7FA3;margin-top:2px">Toggle this when all vacancies are filled or when hiring re-opens.</div>
              </div>
            </div>""",unsafe_allow_html=True)
            gc1,gc2=st.columns(2)
            with gc1:
                if not applications_open and st.button("Open Applications",use_container_width=True):
                    set_setting("applications_open","1",st.session_state.uid); st.success("Applications opened."); st.rerun()
            with gc2:
                if applications_open and st.button("Close Applications",use_container_width=True):
                    set_setting("applications_open","0",st.session_state.uid); st.warning("Applications closed."); st.rerun()
            st.markdown("<hr>",unsafe_allow_html=True)

        conn=get_conn()
        sdf=pd.read_sql_query("SELECT emp_id,full_name,division,contact,age,sex,edu_background,current_status,registration_date FROM employees WHERE current_status='Pending Screening' ORDER BY emp_id",conn)
        conn.close()
        if len(sdf)==0: st.info("No new applications waiting.")
        else:
            st.markdown(f'<div class="card"><span style="color:#D4A847;font-weight:600">⏳ Backlog: <b style="color:#E8EEF7">{len(sdf)}</b> profiles awaiting validation</span></div>',unsafe_allow_html=True)
            st.dataframe(sdf,use_container_width=True,hide_index=True)
            st.markdown("<hr>",unsafe_allow_html=True)
            p1,p2=st.columns(2)
            with p1:
                tid=st.selectbox("Applicant ID",sdf["emp_id"].unique())
                div_list=get_division_list()
                da=st.selectbox("Assign Division",div_list)
                ccs=get_cost_centers(da)
                cc_opts=["Unassigned"]+ccs['code'].tolist() if len(ccs)>0 else ["Unassigned"]
                dcc=st.selectbox("Assign Cost Center",cc_opts)
            with p2:
                st.write(""); st.write(""); st.write("")
                if st.button("Transfer to Pre-Employment",use_container_width=True):
                    conn=get_conn()
                    conn.execute("UPDATE employees SET current_status='Pre-Employment Process',division=?,cost_center=? WHERE emp_id=?",
                        (da, None if dcc=="Unassigned" else dcc, tid))
                    conn.commit(); conn.close()
                    st.cache_data.clear()
                    st.success(f"{tid}  {da} ({dcc}) — Pre-Employment"); st.rerun()

    # ════════════════════════════════════════════════════════
    # EMPLOYEE DIRECTORY (formerly Master Records)
    # ════════════════════════════════════════════════════════
    elif V=="Employee Directory":
        st.markdown('<div class="ey">Enterprise Database</div>',unsafe_allow_html=True)
        st.markdown('<div class="tl">Employee Directory</div>',unsafe_allow_html=True)
        stats=get_stats()
        total=sum(stats.values())
        active=stats.get("Active Deployment",0); pend=stats.get("Pending Screening",0)
        pre=stats.get("Pre-Employment Process",0); onlv=stats.get("On Leave",0); term=stats.get("Terminated",0)
        not_term = total - term
        st.markdown(f"""<div class="mg">
          <div class="mb mg-gold"><div class="ml ml-gold">Total</div><div class="mv">{total}</div></div>
          <div class="mb mg-teal"><div class="ml ml-teal">All Active*</div><div class="mv">{not_term}</div></div>
          <div class="mb mg-green"><div class="ml ml-green">Deployed</div><div class="mv">{active}</div></div>
          <div class="mb mg-cyan"><div class="ml ml-cyan">Applicants</div><div class="mv">{pend}</div></div>
          <div class="mb mg-amber"><div class="ml ml-amber">Pre-Employ</div><div class="mv">{pre}</div></div>
          <div class="mb mg-purple"><div class="ml ml-purple">On Leave</div><div class="mv">{onlv}</div></div>
          <div class="mb mg-red"><div class="ml ml-red">Terminated</div><div class="mv">{term}</div></div>
        </div>""",unsafe_allow_html=True)
        st.markdown('<div style="font-size:10px;color:#6B7FA3;margin:-10px 0 12px">*All Active = every employee except Terminated (includes Deployed, On Leave, Sick, Maternity, Mourning, Pre-Employment)</div>',unsafe_allow_html=True)

        with st.expander("Export Full Employee Data to Excel"):
            ex1,ex2,ex3=st.columns(3)
            with ex1: xst=st.selectbox("Status",["All","Active Deployment","Pending Screening","Pre-Employment Process","On Leave","Terminated"])
            with ex2: xdp=st.selectbox("Division",["All"]+get_division_list())
            with ex3:
                xcc_list=get_cost_centers(xdp if xdp!="All" else None)
                xcc_opts=["All"]+xcc_list['code'].tolist() if len(xcc_list)>0 else ["All"]
                xcc=st.selectbox("Cost Center",xcc_opts)
            if st.button("Prepare Export"):
                conn=get_conn()
                xq="SELECT * FROM employees WHERE 1=1"; xp=[]
                if xst!="All": xq+=" AND current_status=?"; xp.append(xst)
                if xdp!="All": xq+=" AND division=?"; xp.append(xdp)
                if xcc!="All": xq+=" AND cost_center=?"; xp.append(xcc)
                xdf=pd.read_sql_query(xq,conn,params=xp); conn.close()
                st.download_button(f"Download ({len(xdf)} employees)",export_excel(xdf),
                    file_name=f"YGSP_{xst}_{xdp}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
        st.markdown("<hr>",unsafe_allow_html=True)
        f1,f2,f3,f4=st.columns(4)
        with f1: sf=st.selectbox("Status Filter",["All","Active Deployment","Pending Screening","Pre-Employment Process","On Leave","Terminated"],key="sf_rec")
        with f2: df2=st.selectbox("Division Filter",["All"]+get_division_list(),key="df_rec")
        with f3:
            cc_list2=get_cost_centers(df2 if df2!="All" else None)
            cc_opts2=["All"]+cc_list2['code'].tolist() if len(cc_list2)>0 else ["All"]
            ccf=st.selectbox("Cost Center Filter",cc_opts2,key="cc_rec")
        with f4: sv2=st.text_input(" Search",placeholder="Name or ID...",key="sv_rec")
        total_count = count_records(sf if sf!="All" else None, df2 if df2!="All" else None, ccf if ccf!="All" else None, sv2)
        PAGE_SIZE = 50
        total_pages = max((total_count + PAGE_SIZE - 1) // PAGE_SIZE, 1)
        if "dir_page" not in st.session_state: st.session_state.dir_page = 1
        if st.session_state.dir_page > total_pages: st.session_state.dir_page = total_pages

        pg1,pg2,pg3 = st.columns([2,1,1])
        with pg1:
            st.markdown(f'<div style="color:#6B7FA3;font-size:12px;padding-top:8px">Showing page <b style="color:#E8EEF7">{st.session_state.dir_page}</b> of <b style="color:#E8EEF7">{total_pages}</b> ({total_count} total records, {PAGE_SIZE}/page) — <small style="color:#38BDF8">Click a row, then go to Employee Profile for full details</small></div>',unsafe_allow_html=True)
        with pg2:
            if st.button("Previous Page",use_container_width=True,disabled=(st.session_state.dir_page<=1)):
                st.session_state.dir_page -= 1; st.rerun()
        with pg3:
            if st.button("Next Page",use_container_width=True,disabled=(st.session_state.dir_page>=total_pages)):
                st.session_state.dir_page += 1; st.rerun()

        vdf=query_records(sf if sf!="All" else None, df2 if df2!="All" else None, ccf if ccf!="All" else None, sv2, page=st.session_state.dir_page, page_size=PAGE_SIZE)
        sel=st.dataframe(vdf,use_container_width=True,hide_index=True,on_select="rerun",selection_mode="single-row")
        if sel and "rows" in sel["selection"] and len(sel["selection"]["rows"])>0:
            st.session_state.eid=vdf.iloc[sel["selection"]["rows"][0]]["emp_id"]
            st.toast(f"Loaded {st.session_state.eid} Employee Profile")
        if st.session_state.role=="Manager":
            st.markdown("<hr>",unsafe_allow_html=True)
            t1,t2,t3=st.tabs(["Quick Edit","Add New Employee","Remove Employee"])
            with t1:
                if not st.session_state.eid: st.info("Select a row above to edit.")
                else:
                    r=get_employee(st.session_state.eid)
                    if r:
                        with st.form("qe"):
                            q1,q2,q3=st.columns(3)
                            with q1: qn=st.text_input("Full Name",value=r.get("full_name","") or "")
                            with q2: qc=st.text_input("Contact",value=r.get("contact","") or "")
                            with q3: qe=st.text_input("Email",value=r.get("email","") or "")
                            q4,q5,q6,q7=st.columns(4)
                            dlist=get_division_list()
                            so=["Pending Screening","Pre-Employment Process","Active Deployment","On Leave","Terminated"]
                            with q4: qdiv=st.selectbox("Division",dlist,index=dlist.index(r.get("division","Catering")) if r.get("division") in dlist else 0)
                            qcc_list=get_cost_centers(qdiv)
                            qcc_opts=["Unassigned"]+qcc_list['code'].tolist() if len(qcc_list)>0 else ["Unassigned"]
                            with q5:
                                cur_cc=r.get("cost_center") or "Unassigned"
                                qcc=st.selectbox("Cost Center",qcc_opts,index=qcc_opts.index(cur_cc) if cur_cc in qcc_opts else 0)
                            with q6: qst=st.selectbox("Status",so,index=so.index(r.get("current_status","Pending Screening")) if r.get("current_status") in so else 0)
                            with q7: qsal=st.number_input("Basic Salary",min_value=0.0,value=float(r.get("basic_salary") or 0),step=100.0)
                            if st.form_submit_button("Save",use_container_width=True):
                                conn=get_conn()
                                conn.execute("UPDATE employees SET full_name=?,contact=?,email=?,division=?,cost_center=?,current_status=?,basic_salary=? WHERE emp_id=?",
                                    (qn,qc,qe,qdiv,None if qcc=="Unassigned" else qcc,qst,qsal,st.session_state.eid))
                                conn.commit(); conn.close()
                                st.cache_data.clear()
                                st.success("Saved"); st.rerun()
            with t2:
                with st.form("af2"):
                    a1,a2,a3,a4=st.columns(4)
                    with a1: aid=st.text_input("Employee ID")
                    with a2: anam=st.text_input("Full Name")
                    with a3: apho=st.text_input("Phone")
                    with a4: asal=st.number_input("Basic Salary",min_value=0.0,step=100.0)
                    a5,a6=st.columns(2)
                    with a5: adiv=st.selectbox("Division",get_division_list())
                    with a6:
                        acc_list=get_cost_centers(adiv)
                        acc_opts=["Unassigned"]+acc_list['code'].tolist() if len(acc_list)>0 else ["Unassigned"]
                        acc=st.selectbox("Cost Center",acc_opts)
                    if st.form_submit_button("Add",use_container_width=True):
                        if not(aid and anam): st.error("ID and Name required.")
                        else:
                            conn=get_conn()
                            try:
                                conn.execute("INSERT INTO employees(emp_id,full_name,division,cost_center,contact,basic_salary,current_status,registration_date)VALUES(?,?,?,?,?,?,'Active Deployment',?)",
                                    (aid,anam,adiv,None if acc=="Unassigned" else acc,apho,asal,datetime.now().strftime("%Y-%m-%d")))
                                conn.commit()
                                st.cache_data.clear()
                                st.success(f"Employee {aid} added.")
                            except sqlite3.IntegrityError: st.error(f"ID {aid} already exists.")
                            finally: conn.close()
            with t3:
                if st.session_state.eid:
                    did=st.session_state.eid
                    st.markdown(f'<div class="card" style="border-color:rgba(239,68,68,0.3)"><span style="color:#EF4444;font-weight:600">This will move <b style="color:#F0C96B">{did}</b> to the Recycle Bin. It can be restored later.</span></div>',unsafe_allow_html=True)
                    if st.button(f"Remove {did}",use_container_width=True):
                        emp_full = get_employee(did)
                        conn=get_conn(); conn.execute("DELETE FROM employees WHERE emp_id=?",(did,)); conn.commit(); conn.close()
                        if emp_full:
                            soft_delete("Employee", did, f"{did} — {emp_full.get('full_name','')}", emp_full, st.session_state.uid)
                        st.cache_data.clear(); get_employee.clear()
                        st.session_state.eid=None; st.success("Moved to Recycle Bin."); st.rerun()
                else: st.info("Select a row above.")

    # ════════════════════════════════════════════════════════
    # EMPLOYEE PROFILE (Full View + Documents with Delete)
    # ════════════════════════════════════════════════════════
    elif V=="Employee Profile":
        st.markdown('<div class="ey">Full Profile</div>',unsafe_allow_html=True)
        st.markdown('<div class="tl">Employee Record & Documents</div>',unsafe_allow_html=True)
        emp_df=get_emp_list_cached()
        opts={f"{r['emp_id']} — {r['full_name']} ({r['division']})":r['emp_id'] for _,r in emp_df.iterrows()}
        def_idx=list(opts.values()).index(st.session_state.eid) if st.session_state.eid in list(opts.values()) else 0
        lbl=st.selectbox("Select Employee",list(opts.keys()),index=def_idx)
        eid2=opts[lbl]; st.session_state.eid=eid2
        r=get_employee(eid2)
        if not r: st.warning("Not found."); st.stop()
        st.markdown("<hr>",unsafe_allow_html=True)
        cc,pc=st.columns([2.5,1])
        with cc:
            sc=sclass(r.get("current_status",""))
            bs=float(r.get("basic_salary") or 0)
            cc_tag=f'<span class="cc-tag">{r.get("cost_center")}</span>' if r.get("cost_center") else ""
            st.markdown(f"""<div class="ic">
              <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:14px">
                <div><div style="font-size:8px;color:#D4A847;letter-spacing:.12em;text-transform:uppercase;margin-bottom:2px">Employee ID</div>
                <div class="ifv-g">{r.get("emp_id","")}</div></div>
                <span class="sp {sc}">{r.get("current_status","")}</span>
              </div>
              <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:9px">
                <div><div class="ifl">Full Name</div><div class="ifv">{r.get("full_name","")}</div></div>
                <div><div class="ifl">Job Title</div><div class="ifv">{r.get("job_title","—")}</div></div>
                <div><div class="ifl">Division</div><div class="ifv">{r.get("division","")}{cc_tag}</div></div>
                <div><div class="ifl">Type</div><div class="ifv">{r.get("employment_type","—")}</div></div>
                <div><div class="ifl">Start Date</div><div class="ifv">{r.get("start_date","—")}</div></div>
                <div><div class="ifl">Contract End</div><div class="ifv">{r.get("contract_end_date","—")}</div></div>
                <div><div class="ifl">Contact</div><div class="ifv">{r.get("contact","—")}</div></div>
                <div><div class="ifl">Email</div><div class="ifv">{r.get("email","—")}</div></div>
                <div><div class="ifl">Salary</div><div class="ifv" style="color:#10B981">ETB {bs:,.2f}</div></div>
                <div><div class="ifl">Bank/Account</div><div class="ifv">{r.get("bank_name","—")}/{r.get("bank_account","—")}</div></div>
                <div><div class="ifl">TIN</div><div class="ifv">{r.get("tin_number","—")}</div></div>
                <div><div class="ifl">Weekly Day-Off</div><div class="ifv" style="color:#38BDF8">{r.get("weekly_dayoff","Sunday")}</div></div>
              </div></div>""",unsafe_allow_html=True)
        with pc:
            st.markdown('<div style="font-size:8px;color:#D4A847;letter-spacing:.1em;text-transform:uppercase;margin-bottom:5px">Profile Photo</div>',unsafe_allow_html=True)
            if r.get("photo_data"):
                b64p,extp=b64file(r["photo_data"],r.get("photo_name","p.jpg"))
                if b64p and extp in ["jpg","jpeg","png"]:
                    st.markdown(f'<div class="pb"><img src="data:image/{extp};base64,{b64p}" style="max-width:100%;max-height:190px;border-radius:8px;object-fit:contain"/></div>',unsafe_allow_html=True)
            else:
                st.markdown('<div class="pb" style="display:flex;align-items:center;justify-content:center;min-height:170px;flex-direction:column;gap:6px"><div style="font-size:36px;opacity:0.25"></div><div style="font-size:10px;color:#6B7FA3">No photo</div></div>',unsafe_allow_html=True)
        st.markdown("<hr>",unsafe_allow_html=True)
        t1,t2,t3,t4,t5,t6=st.tabs(["Personal","Edu & Work","Financial","Documents","Edit","History"])
        with t1:
            c1,c2=st.columns(2)
            with c1:
                st.markdown(f"""<div class="card">
                  <div class="fs">Personal</div>
                  <div><div class="ifl">Sex</div><div class="ifv">{r.get("sex","—")}</div></div>
                  <div><div class="ifl">Age</div><div class="ifv">{r.get("age","—")}</div></div>
                  <div><div class="ifl">Marital</div><div class="ifv">{r.get("marital_status","—")}</div></div>
                  <div><div class="ifl">Nationality</div><div class="ifv">{r.get("nationality","—")}</div></div>
                  <div><div class="ifl">Religion</div><div class="ifv">{r.get("religion","—")}</div></div>
                  <div><div class="ifl">Blood Type</div><div class="ifv" style="color:#EF4444">{r.get("blood_type","—")}</div></div>
                  <div><div class="ifl">Place of Birth</div><div class="ifv">{r.get("place_of_birth","—")}</div></div>
                  <div><div class="ifl">Resident ID</div><div class="ifv">{r.get("resident_id","—")}</div></div>
                </div>""",unsafe_allow_html=True)
            with c2:
                st.markdown(f"""<div class="card">
                  <div class="fs">Address</div>
                  <div><div class="ifl">House No.</div><div class="ifv">{r.get("house_address","—")}</div></div>
                  <div><div class="ifl">Woreda</div><div class="ifv">{r.get("woreda","—")}</div></div>
                  <div><div class="ifl">Subcity</div><div class="ifv">{r.get("subcity","—")}</div></div>
                  <div><div class="ifl">Kebele</div><div class="ifv">{r.get("kebele","—")}</div></div>
                  <div class="fs" style="margin-top:10px">Emergency Contact</div>
                  <div><div class="ifl">Name</div><div class="ifv">{r.get("emergency_contact_name","—")}</div></div>
                  <div><div class="ifl">Phone</div><div class="ifv">{r.get("emergency_contact_phone","—")}</div></div>
                </div>""",unsafe_allow_html=True)
        with t2:
            c1,c2=st.columns(2)
            with c1:
                st.markdown(f"""<div class="card">
                  <div class="fs">Education</div>
                  <div><div class="ifl">Level</div><div class="ifv">{r.get("edu_background","—")}</div></div>
                  <div><div class="ifl">Field</div><div class="ifv">{r.get("field_of_graduate","—")}</div></div>
                  <div><div class="ifl">Grad Year</div><div class="ifv">{r.get("graduation_year","—")}</div></div>
                  <div><div class="ifl">Institution</div><div class="ifv">{r.get("institution_name","—")}</div></div>
                </div>""",unsafe_allow_html=True)
            with c2:
                st.markdown(f"""<div class="card">
                  <div class="fs">Employment</div>
                  <div><div class="ifl">Job Title</div><div class="ifv">{r.get("job_title","—")}</div></div>
                  <div><div class="ifl">Division</div><div class="ifv">{r.get("division","—")}</div></div>
                  <div><div class="ifl">Cost Center</div><div class="ifv">{r.get("cost_center","—")}</div></div>
                  <div><div class="ifl">Type</div><div class="ifv">{r.get("employment_type","—")}</div></div>
                  <div><div class="ifl">Weekly Day-Off</div><div class="ifv">{r.get("weekly_dayoff","Sunday")}</div></div>
                  <div><div class="ifl">Reg. Date</div><div class="ifv">{r.get("registration_date","—")}</div></div>
                </div>""",unsafe_allow_html=True)
        with t3:
            bs2=float(r.get("basic_salary") or 0)
            st.markdown(f"""<div class="card">
              <div class="fs">Financial</div>
              <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:10px">
                <div><div class="ifl">Basic Salary</div><div class="ifv" style="color:#10B981;font-size:15px;font-family:'Cinzel',serif">ETB {bs2:,.2f}</div></div>
                <div><div class="ifl">Daily Rate (÷26)</div><div class="ifv">ETB {bs2/26:,.2f}</div></div>
                <div><div class="ifl">Hourly (÷8)</div><div class="ifv">ETB {bs2/26/8:,.2f}</div></div>
                <div><div class="ifl">Income Tax (est.)</div><div class="ifv" style="color:#FCA5A5">ETB {eth_tax(bs2):,.2f}</div></div>
                <div><div class="ifl">Pension Emp 7%</div><div class="ifv" style="color:#FCA5A5">ETB {bs2*0.07:,.2f}</div></div>
                <div><div class="ifl">Pension Er 11%</div><div class="ifv" style="color:#38BDF8">ETB {bs2*0.11:,.2f}</div></div>
                <div><div class="ifl">TIN</div><div class="ifv">{r.get("tin_number","—")}</div></div>
                <div><div class="ifl">Pension No.</div><div class="ifv">{r.get("pension_number","—")}</div></div>
                <div><div class="ifl">Bank/Account</div><div class="ifv">{r.get("bank_name","—")} {r.get("bank_account","—")}</div></div>
              </div></div>""",unsafe_allow_html=True)
            conn=get_conn()
            ph=pd.read_sql_query("SELECT month,COALESCE(gross_salary,basic_salary) as gross_salary,net_salary,COALESCE(fine_amount,0) as fine_amount,COALESCE(absent_days,0) as absent_days,COALESCE(dayoff_days,4) as dayoff_days,COALESCE(holiday_days,0) as holiday_days,payment_status FROM payroll WHERE emp_id=? ORDER BY created_at DESC LIMIT 12",conn,params=(eid2,)); conn.close()
            if len(ph)>0:
                st.markdown('<div class="ey" style="margin-top:10px">Payroll History</div>',unsafe_allow_html=True)
                st.dataframe(ph,use_container_width=True,hide_index=True)
        with t4:
            doc_defs=[
                (" Photo","photo_name","photo_data",["jpg","jpeg","png"]),
                (" Education Docs","edu_doc_name","edu_doc_data",["pdf","zip","jpg","png"]),
                (" Forensic/Medical Scan","forensic_doc_name","forensic_doc_data",["pdf","jpg","png"]),
                (" National ID/Passport","id_scan_name","id_scan_data",["pdf","jpg","png"]),
                (" Medical Report","medical_doc_name","medical_doc_data",["pdf","jpg","png"]),
                (" Guarantee Letter","guarantee_letter_name","guarantee_letter_data",["pdf"]),
                (" Police Clearance","police_clearance_name","police_clearance_data",["pdf","jpg"]),
                (" Contract","contract_doc_name","contract_doc_data",["pdf"]),
                (" First Document","first_doc_name","first_doc_data",["pdf","jpg","png","zip"]),
            ]
            bh='<div style="margin-bottom:10px">'
            for lbl2,nk,dk,_ in doc_defs:
                bh+=f'<span class="db {"db-up" if r.get(nk) else "db-mis"}">{"" if r.get(nk) else ""} {lbl2}</span>'
            bh+='</div>'
            st.markdown(bh,unsafe_allow_html=True)
            st.markdown("""<div class="scan">
              <div style="font-size:28px;margin-bottom:5px"></div>
              <div style="font-size:12px;color:#38BDF8;font-weight:500">Scanner / Printer / Camera Upload</div>
              <div style="font-size:10px;color:#6B7FA3;margin-top:2px">Scan  save as PDF/JPG  upload below. Or take a phone photo and upload directly.</div>
            </div>""",unsafe_allow_html=True)
            st.markdown("<hr>",unsafe_allow_html=True)
            for lbl2,nk,dk,allowed in doc_defs:
                fname=r.get(nk); fdata=r.get(dk)
                with st.expander(f"{lbl2}  {' Preview available' if fname else ' Not uploaded'}",expanded=False):
                    st.markdown(f'<div class="pb">{preview_html(fdata,fname,lbl2)}</div>',unsafe_allow_html=True)
                    if fname and fdata:
                        dcol1,dcol2=st.columns(2)
                        with dcol1:
                            st.download_button(f"Download",data=bytes(fdata),file_name=fname,use_container_width=True,key=f"dl_{dk}_{eid2}")
                        with dcol2:
                            if st.session_state.role=="Manager":
                                if st.button(f"Delete {lbl2}",key=f"del_{dk}_{eid2}",use_container_width=True):
                                    conn=get_conn()
                                    conn.execute(f"UPDATE employees SET {nk}=NULL,{dk}=NULL WHERE emp_id=?",(eid2,))
                                    conn.commit(); conn.close()
                                    get_employee.clear()
                                    st.success(f"{lbl2} deleted."); st.rerun()
                    if st.session_state.role=="Manager":
                        st.markdown('<div style="font-size:10px;color:#6B7FA3;margin:8px 0 5px"> Upload new / replace — from scanner, camera or file</div>',unsafe_allow_html=True)
                        upl=st.file_uploader(f"Upload {lbl2}",type=allowed,key=f"up_{dk}_{eid2}")
                        if upl and st.button("Save",key=f"sv_{dk}_{eid2}",use_container_width=True):
                            fb=upl.read()
                            conn=get_conn()
                            conn.execute(f"UPDATE employees SET {nk}=?,{dk}=? WHERE emp_id=?",(upl.name,sqlite3.Binary(fb),eid2))
                            conn.commit(); conn.close()
                            get_employee.clear(); st.success(f"{lbl2} saved!"); st.rerun()
        with t5:
            if st.session_state.role=="Manager":
                with st.form(f"fe_{eid2}"):
                    st.markdown('<div class="fs">Personal Information</div>',unsafe_allow_html=True)
                    p1,p2,p3=st.columns(3)
                    with p1: en=st.text_input("Full Name",value=r.get("full_name","") or "")
                    with p2: ec=st.text_input("Contact",value=r.get("contact","") or "")
                    with p3: ee=st.text_input("Email",value=r.get("email","") or "")
                    p4,p5,p6,p7=st.columns(4)
                    with p4: esx=st.selectbox("Sex",["Male","Female"],index=0 if r.get("sex")=="Male" else 1)
                    with p5:
                        mo=["Single","Married","Divorced","Widowed"]
                        emar=st.selectbox("Marital",mo,index=mo.index(r.get("marital_status","Single")) if r.get("marital_status") in mo else 0)
                    with p6: enat=st.text_input("Nationality",value=r.get("nationality","Ethiopian") or "Ethiopian")
                    with p7:
                        ro=["Orthodox","Muslim","Protestant","Catholic","Other"]
                        erel=st.selectbox("Religion",ro,index=ro.index(r.get("religion","Orthodox")) if r.get("religion") in ro else 0)
                    p8,p9,p10,p11=st.columns(4)
                    with p8: eage=st.number_input("Age",18,80,value=int(r.get("age") or 25))
                    with p9: epob=st.text_input("Place of Birth",value=r.get("place_of_birth","") or "")
                    with p10:
                        bo=["A+","A-","B+","B-","O+","O-","AB+","AB-"]
                        ebt=st.selectbox("Blood",bo,index=bo.index(r.get("blood_type","O+")) if r.get("blood_type") in bo else 0)
                    with p11: eres=st.text_input("Resident ID",value=r.get("resident_id","") or "")
                    st.markdown('<div class="fs">Address</div>',unsafe_allow_html=True)
                    a1,a2,a3,a4=st.columns(4)
                    with a1: eha=st.text_input("House",value=r.get("house_address","") or "")
                    with a2: ewo=st.text_input("Woreda",value=r.get("woreda","") or "")
                    with a3: esc2=st.text_input("Subcity",value=r.get("subcity","") or "")
                    with a4: eke=st.text_input("Kebele",value=r.get("kebele","") or "")
                    st.markdown('<div class="fs">Emergency Contact</div>',unsafe_allow_html=True)
                    ec1,ec2=st.columns(2)
                    with ec1: eecn=st.text_input("Name",value=r.get("emergency_contact_name","") or "")
                    with ec2: eecp=st.text_input("Phone",value=r.get("emergency_contact_phone","") or "")
                    st.markdown('<div class="fs">Financial & IDs</div>',unsafe_allow_html=True)
                    f1,f2,f3,f4=st.columns(4)
                    with f1: etin=st.text_input("TIN",value=r.get("tin_number","") or "")
                    with f2: epen2=st.text_input("Pension No.",value=r.get("pension_number","") or "")
                    with f3: ebnk=st.text_input("Bank",value=r.get("bank_name","") or "")
                    with f4: eacc=st.text_input("Account",value=r.get("bank_account","") or "")
                    st.markdown('<div class="fs">Education</div>',unsafe_allow_html=True)
                    ed1,ed2,ed3,ed4=st.columns(4)
                    eo=["High School Graduate","TVET Diploma","BSc/BA Degree","MSc/MA Post-Graduate"]
                    with ed1: eedu=st.selectbox("Level",eo,index=eo.index(r.get("edu_background","High School Graduate")) if r.get("edu_background") in eo else 0)
                    with ed2: efld=st.text_input("Field",value=r.get("field_of_graduate","") or "")
                    with ed3: egry=st.text_input("Grad Year",value=r.get("graduation_year","") or "")
                    with ed4: eins=st.text_input("Institution",value=r.get("institution_name","") or "")
                    st.markdown('<div class="fs">Employment & Division</div>',unsafe_allow_html=True)
                    em1,em2b,em3=st.columns(3)
                    with em1: ejob=st.text_input("Job Title",value=r.get("job_title","") or "")
                    with em2b:
                        et=["Permanent","Contract","Temporary","Part-Time"]
                        eety=st.selectbox("Type",et,index=et.index(r.get("employment_type","Permanent")) if r.get("employment_type") in et else 0)
                    with em3:
                        dlist2=get_division_list()
                        edep=st.selectbox("Division",dlist2,index=dlist2.index(r.get("division","Catering")) if r.get("division") in dlist2 else 0)
                    em_cc_list=get_cost_centers(edep)
                    em_cc_opts=["Unassigned"]+em_cc_list['code'].tolist() if len(em_cc_list)>0 else ["Unassigned"]
                    em4,em5,em6=st.columns(3)
                    with em4:
                        cur_ecc=r.get("cost_center") or "Unassigned"
                        eccsel=st.selectbox("Cost Center",em_cc_opts,index=em_cc_opts.index(cur_ecc) if cur_ecc in em_cc_opts else 0)
                    with em5: esal=st.number_input("Basic Salary",min_value=0.0,value=float(r.get("basic_salary") or 0),step=100.0)
                    with em6:
                        wd_opts=["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"]
                        cur_wd=r.get("weekly_dayoff") or "Sunday"
                        ewd=st.selectbox("Weekly Day-Off",wd_opts,index=wd_opts.index(cur_wd) if cur_wd in wd_opts else 0,
                            help="System automatically calculates every matching date each month for payroll")
                    em7,em8=st.columns(2)
                    with em7: esd=st.text_input("Start Date (YYYY-MM-DD)",value=r.get("start_date","") or "")
                    with em8: ecd=st.text_input("Contract End (YYYY-MM-DD)",value=r.get("contract_end_date","") or "")
                    so2=["Pending Screening","Pre-Employment Process","Active Deployment","On Leave","Terminated"]
                    est=st.selectbox("Status",so2,index=so2.index(r.get("current_status","Pending Screening")) if r.get("current_status") in so2 else 0)
                    enotes=st.text_area("Internal Notes",value=r.get("notes","") or "")
                    if st.form_submit_button("Save All Changes",use_container_width=True):
                        conn=get_conn()
                        conn.execute("""UPDATE employees SET full_name=?,contact=?,email=?,sex=?,marital_status=?,nationality=?,religion=?,
                            age=?,place_of_birth=?,blood_type=?,resident_id=?,house_address=?,woreda=?,subcity=?,kebele=?,
                            emergency_contact_name=?,emergency_contact_phone=?,tin_number=?,pension_number=?,bank_name=?,bank_account=?,
                            edu_background=?,field_of_graduate=?,graduation_year=?,institution_name=?,job_title=?,employment_type=?,
                            division=?,cost_center=?,basic_salary=?,weekly_dayoff=?,start_date=?,contract_end_date=?,current_status=?,notes=? WHERE emp_id=?""",
                            (en,ec,ee,esx,emar,enat,erel,eage,epob,ebt,eres,eha,ewo,esc2,eke,eecn,eecp,etin,epen2,ebnk,eacc,
                             eedu,efld,egry,eins,ejob,eety,edep,None if eccsel=="Unassigned" else eccsel,esal,ewd,esd,ecd,est,enotes,eid2))
                        conn.commit(); conn.close()
                        st.cache_data.clear()
                        st.success("Profile saved!"); st.rerun()
            else: st.info("Manager role required.")
        with t6:
            st.markdown(f'<div class="card"><div class="fs">Notes</div><div style="color:#C8D8F0;font-size:12px;line-height:1.7;white-space:pre-wrap">{r.get("notes","No notes.") or "No notes."}</div></div>',unsafe_allow_html=True)
            conn=get_conn()
            lv=pd.read_sql_query("SELECT leave_type,start_date,end_date,days_taken,CASE WHEN is_paid=1 THEN 'Paid' ELSE 'Unpaid' END as paid,deduction_amount,status FROM leave_records WHERE emp_id=? ORDER BY created_at DESC",conn,params=(eid2,))
            fl=pd.read_sql_query("SELECT COALESCE(month,'—') as month,issue_date,COALESCE(fine_type,'Disciplinary') as fine_type,COALESCE(fine_reason,'') as fine_reason,fine_days,fine_amount,applied_to_payroll FROM fine_letters WHERE emp_id=? ORDER BY created_at DESC",conn,params=(eid2,))
            ab=pd.read_sql_query("SELECT absent_date,reason,CASE WHEN is_excused=1 THEN 'Excused' ELSE 'Unexcused' END as type FROM absent_records WHERE emp_id=? ORDER BY absent_date DESC",conn,params=(eid2,))
            conn.close()
            if len(lv)>0:
                st.markdown('<div class="ey" style="margin-top:8px">Leave History</div>',unsafe_allow_html=True); st.dataframe(lv,use_container_width=True,hide_index=True)
            if len(fl)>0:
                st.markdown('<div class="ey" style="margin-top:8px">Fine History</div>',unsafe_allow_html=True); st.dataframe(fl,use_container_width=True,hide_index=True)
            if len(ab)>0:
                st.markdown('<div class="ey" style="margin-top:8px">Absent Records</div>',unsafe_allow_html=True); st.dataframe(ab,use_container_width=True,hide_index=True)


    # ════════════════════════════════════════════════════════
    # SUPERVISOR CONSOLE
    # Supervisors control attendance, absence and leave only
    # for employees in their assigned division. They compile
    # the monthly sheet and submit it to the Payroll Section
    # for approval before any salary is paid.
    # ════════════════════════════════════════════════════════
    elif V=="Supervisor Console":
        my_division = st.session_state.assigned_division
        st.markdown('<div class="ey">Division Operations</div>',unsafe_allow_html=True)
        st.markdown('<div class="tl">Supervisor Console</div>',unsafe_allow_html=True)

        if not my_division:
            st.markdown("""<div class="card" style="border-color:rgba(239,68,68,0.3)">
              <span style="color:#EF4444;font-weight:600">No division assigned.</span>
              <span style="color:#6B7FA3"> Contact the Manager to assign you to a division in Administration.</span>
            </div>""",unsafe_allow_html=True)
            st.stop()

        st.markdown(f'<div class="card card-gold">Supervising division: <b style="color:#F0C96B">{my_division}</b></div>',unsafe_allow_html=True)

        conn=get_conn()
        my_emps=pd.read_sql_query(
            "SELECT emp_id,full_name,job_title,cost_center,weekly_dayoff,basic_salary,current_status FROM employees WHERE division=? AND current_status != 'Terminated' ORDER BY emp_id",
            conn, params=(my_division,))
        conn.close()

        sc1,sc2,sc3=st.columns(3)
        with sc1: total_div=len(my_emps)
        with sc2: active_div=len(my_emps[my_emps['current_status']=='Active Deployment'])
        with sc3: leave_div=len(my_emps[my_emps['current_status']=='On Leave'])
        st.markdown(f"""<div class="mg" style="grid-template-columns:repeat(3,1fr)">
          <div class="mb mg-gold"><div class="ml ml-gold">Division Employees</div><div class="mv">{total_div}</div></div>
          <div class="mb mg-green"><div class="ml ml-green">Active Deployment</div><div class="mv">{active_div}</div></div>
          <div class="mb mg-purple"><div class="ml ml-purple">On Leave</div><div class="mv">{leave_div}</div></div>
        </div>""",unsafe_allow_html=True)

        sup1,sup2,sup3,sup4,sup5=st.tabs(["Record Absence","Record Leave","Issue Fine","Movement Log","Submit Monthly Sheet"])

        elo_sup={f"{r['emp_id']} — {r['full_name']}":r['emp_id'] for _,r in my_emps.iterrows()}

        with sup1:
            st.markdown('<div style="color:#6B7FA3;font-size:12px;margin-bottom:10px">Record an absence for any employee in your division. Unexcused absences are automatically deducted from salary.</div>',unsafe_allow_html=True)
            if len(elo_sup)==0:
                st.info("No employees currently assigned to your division.")
            else:
                with st.form("sup_absent_form"):
                    a1,a2,a3=st.columns(3)
                    with a1: ab_emp=st.selectbox("Employee",list(elo_sup.keys()))
                    with a2: ab_date=st.date_input("Date",value=date.today())
                    with a3: ab_type=st.selectbox("Type",["Unexcused (Deducted)","Excused (Not Deducted)"])
                    ab_reason=st.text_input("Reason")
                    if st.form_submit_button("Record Absence",use_container_width=True):
                        ab_eid_sup=elo_sup[ab_emp]
                        conn=get_conn(); cur=conn.cursor()
                        cur.execute("""SELECT COUNT(*) FROM absent_records WHERE emp_id=? AND absent_date=?
                            AND COALESCE(record_status,'Active')='Active'""",(ab_eid_sup,str(ab_date)))
                        dup_count_sup=cur.fetchone()[0]
                        if dup_count_sup>0:
                            conn.close()
                            st.error(f"{ab_emp} already has an absence recorded for {ab_date}. The system blocks duplicate entries automatically.")
                        else:
                            conn.execute("INSERT INTO absent_records(emp_id,absent_date,reason,is_excused,record_status,created_at)VALUES(?,?,?,?,'Active',?)",
                                (ab_eid_sup,str(ab_date),f"Supervisor {st.session_state.full_name or st.session_state.uid}: {ab_reason}",1 if "Excused" in ab_type else 0,datetime.now().strftime("%Y-%m-%d")))
                            conn.commit(); conn.close()
                            st.success(f"Absence recorded for {ab_date}"); st.rerun()
                conn=get_conn()
                div_absences=pd.read_sql_query("""SELECT ar.emp_id,e.full_name,ar.absent_date,ar.reason,
                    CASE WHEN ar.is_excused=1 THEN 'Excused' ELSE 'Unexcused' END as type
                    FROM absent_records ar JOIN employees e ON ar.emp_id=e.emp_id
                    WHERE e.division=? ORDER BY ar.absent_date DESC LIMIT 100""",conn,params=(my_division,))
                conn.close()
                if len(div_absences)>0:
                    st.markdown('<div class="ey" style="margin-top:10px">Recent Absences — Your Division</div>',unsafe_allow_html=True)
                    st.dataframe(div_absences,use_container_width=True,hide_index=True)

        with sup2:
            st.markdown('<div style="color:#6B7FA3;font-size:12px;margin-bottom:10px">Record leave for any employee in your division. Sick, Annual, Maternity, Paternity and Mourning leave are paid per Ethiopian Labour Law.</div>',unsafe_allow_html=True)
            if len(elo_sup)==0:
                st.info("No employees currently assigned to your division.")
            else:
                st.markdown('<div style="font-size:11px;color:#F59E0B;margin-bottom:8px">Leave records you submit are sent to the Department Head for approval before they take effect on payroll.</div>',unsafe_allow_html=True)
                with st.form("sup_leave_form"):
                    l1,l2=st.columns(2)
                    with l1: l_emp=st.selectbox("Employee",list(elo_sup.keys()),key="sup_l_emp")
                    with l2:
                        ltypes=["Sick Leave","Annual Leave","Maternity Leave","Paternity Leave","Mourning Leave","Unpaid Leave","Emergency Leave","Study Leave"]
                        l_type=st.selectbox("Leave Type",ltypes,key="sup_l_type")
                    l4,l5=st.columns(2)
                    with l4: l_start=st.date_input("Start",value=date.today(),key="sup_l_start")
                    with l5: l_end=st.date_input("End",value=date.today(),key="sup_l_end")
                    l_notes=st.text_area("Notes")
                    if st.form_submit_button("Submit Leave for Approval",use_container_width=True):
                        l_eid=elo_sup[l_emp]
                        conn=get_conn(); cur=conn.cursor()
                        cur.execute("""SELECT COUNT(*) FROM leave_records WHERE emp_id=? AND status != 'Cancelled'
                            AND start_date<=? AND end_date>=?""",(l_eid,str(l_end),str(l_start)))
                        overlap_count_sup=cur.fetchone()[0]
                        if overlap_count_sup>0:
                            conn.close()
                            st.error(f"{l_emp} already has a leave record overlapping {l_start} to {l_end}.")
                            st.stop()
                        cur.execute("SELECT basic_salary FROM employees WHERE emp_id=?",(l_eid,))
                        sr=cur.fetchone(); dr=float(sr[0])/26 if sr and sr[0] else 0
                        days=max((l_end-l_start).days+1,0)
                        is_paid=0 if l_type=="Unpaid Leave" else 1
                        ded=0.0 if is_paid else round(dr*days,2)
                        conn.execute("INSERT INTO leave_records(emp_id,leave_type,start_date,end_date,days_taken,is_paid,daily_rate,deduction_amount,approved_by,status,notes,created_at)VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",
                            (l_eid,l_type,str(l_start),str(l_end),days,is_paid,round(dr,2),ded,
                             f"Supervisor: {st.session_state.full_name or st.session_state.uid}","Pending Dept Head Approval",l_notes,datetime.now().strftime("%Y-%m-%d")))
                        conn.commit(); conn.close()
                        st.success(f"{l_type} submitted for Department Head approval: {days} days"); st.rerun()
                conn=get_conn()
                div_leave=pd.read_sql_query("""SELECT lr.emp_id,e.full_name,lr.leave_type,lr.start_date,lr.end_date,lr.days_taken,lr.status
                    FROM leave_records lr JOIN employees e ON lr.emp_id=e.emp_id
                    WHERE e.division=? ORDER BY lr.created_at DESC LIMIT 100""",conn,params=(my_division,))
                conn.close()
                if len(div_leave)>0:
                    st.markdown('<div class="ey" style="margin-top:10px">Recent Leave — Your Division</div>',unsafe_allow_html=True)
                    st.dataframe(div_leave,use_container_width=True,hide_index=True)

        with sup3:
            st.markdown('<div style="color:#6B7FA3;font-size:12px;margin-bottom:10px">Issue a fine letter for an employee in your division. Submitted fines are sent to the Department Head for approval before they apply to payroll.</div>',unsafe_allow_html=True)
            if len(elo_sup)==0:
                st.info("No employees currently assigned to your division.")
            else:
                with st.form("sup_fine_form"):
                    f1,f2=st.columns(2)
                    with f1: fn_emp=st.selectbox("Employee",list(elo_sup.keys()),key="sup_fn_emp")
                    with f2:
                        fine_month_sup=st.selectbox("Fine Month",[f"{datetime.now().year}-{m:02d}" for m in range(1,13)],index=datetime.now().month-1,key="sup_fn_month")
                    f3,f4=st.columns(2)
                    with f3:
                        ftypes_sup=["Disciplinary","Misconduct","Late Arrival","Unauthorized Absence","Policy Violation","Performance","Other"]
                        fn_type=st.selectbox("Fine Type",ftypes_sup,key="sup_fn_type")
                    with f4: fn_days=st.number_input("Fine Days",min_value=0,max_value=30,step=1,key="sup_fn_days")
                    fn_reason=st.text_area("Reason",placeholder="Describe the violation...",key="sup_fn_reason")
                    fn_letter=st.file_uploader("Attach Scanned Fine Letter (optional)",type=["pdf","jpg","jpeg","png"],key="sup_fn_letter")
                    if st.form_submit_button("Submit Fine for Approval",use_container_width=True):
                        fn_eid=elo_sup[fn_emp]
                        conn=get_conn(); cur=conn.cursor()
                        cur.execute("SELECT basic_salary FROM employees WHERE emp_id=?",(fn_eid,))
                        sr=cur.fetchone(); dr=float(sr[0])/26 if sr and sr[0] else 0
                        fine_amt=round(dr*fn_days,2)
                        ln=fn_letter.name if fn_letter else None
                        ld=sqlite3.Binary(fn_letter.read()) if fn_letter else None
                        conn.execute("""INSERT INTO fine_letters(emp_id,month,issue_date,fine_reason,fine_type,fine_days,fine_amount,letter_name,letter_data,applied_to_payroll,created_at)
                            VALUES(?,?,?,?,?,?,?,?,?,'Pending Dept Head Approval',?)""",
                            (fn_eid,fine_month_sup,datetime.now().strftime("%Y-%m-%d"),
                             f"[Supervisor: {st.session_state.full_name or st.session_state.uid}] {fn_reason}",fn_type,fn_days,fine_amt,ln,ld,datetime.now().strftime("%Y-%m-%d")))
                        conn.commit(); conn.close()
                        st.success(f"Fine submitted for Department Head approval: {fn_days} days = ETB {fine_amt:,.2f}"); st.rerun()
                conn=get_conn()
                div_fines=pd.read_sql_query("""SELECT fl.emp_id,e.full_name,fl.fine_type,fl.fine_days,fl.fine_amount,fl.applied_to_payroll
                    FROM fine_letters fl JOIN employees e ON fl.emp_id=e.emp_id
                    WHERE e.division=? ORDER BY fl.created_at DESC LIMIT 100""",conn,params=(my_division,))
                conn.close()
                if len(div_fines)>0:
                    st.markdown('<div class="ey" style="margin-top:10px">Recent Fines — Your Division</div>',unsafe_allow_html=True)
                    st.dataframe(div_fines,use_container_width=True,hide_index=True)

        with sup4:
            st.markdown('<div style="color:#6B7FA3;font-size:12px;margin-bottom:10px">Work movement log — transfers, status changes and notes for employees in your division.</div>',unsafe_allow_html=True)
            st.dataframe(my_emps[['emp_id','full_name','job_title','cost_center','current_status']],use_container_width=True,hide_index=True)
            if len(elo_sup)>0:
                with st.form("sup_status_form"):
                    ms1,ms2=st.columns(2)
                    with ms1: mv_emp=st.selectbox("Employee",list(elo_sup.keys()),key="sup_mv_emp")
                    with ms2:
                        mv_status=st.selectbox("New Status",["Active Deployment","On Leave"])
                    mv_notes=st.text_input("Notes",placeholder="Reason for status change...")
                    if st.form_submit_button("Update Status",use_container_width=True):
                        mv_eid=elo_sup[mv_emp]
                        conn=get_conn()
                        conn.execute("UPDATE employees SET current_status=?,notes=COALESCE(notes,'') || ? WHERE emp_id=?",
                            (mv_status, f"\\n[{datetime.now().strftime('%Y-%m-%d')}] Supervisor {st.session_state.uid}: {mv_notes}", mv_eid))
                        conn.commit(); conn.close()
                        st.cache_data.clear()
                        st.success(f"{mv_eid} status updated to {mv_status}"); st.rerun()

        with sup5:
            st.markdown('<div style="color:#6B7FA3;font-size:13px;margin-bottom:14px">Compile the full month attendance for one of your cost centers and submit it to the Payroll Section for review and approval. Salaries are only released after Payroll Section approves.</div>',unsafe_allow_html=True)
            my_ccs=get_cost_centers(my_division)
            if len(my_ccs)==0:
                st.warning("No cost centers exist for your division yet. Ask the Manager to create one in Cost Centers.")
            else:
                sub1,sub2,sub3=st.columns(3)
                with sub1: submit_cc=st.selectbox("Cost Center",my_ccs['code'].tolist())
                with sub2: submit_yr=st.selectbox("Year",[datetime.now().year,datetime.now().year+1],key="sub_yr")
                with sub3: submit_mo=st.selectbox("Month",list(range(1,13)),index=datetime.now().month-1,format_func=lambda m: calendar.month_name[m],key="sub_mo")
                submit_month_str=f"{submit_yr}-{submit_mo:02d}"

                conn=get_conn()
                cc_emps=pd.read_sql_query("SELECT emp_id,full_name,basic_salary,weekly_dayoff FROM employees WHERE cost_center=? AND current_status != 'Terminated'",conn,params=(submit_cc,))
                existing_sub=pd.read_sql_query("SELECT * FROM payroll_submissions WHERE cost_center=? AND month=?",conn,params=(submit_cc,submit_month_str))
                conn.close()

                if len(existing_sub)>0:
                    st_row=existing_sub.iloc[0]
                    badge_cls = {"Pending Approval":"spe","Approved":"sa","Rejected":"st"}.get(st_row['status'],"spe")
                    st.markdown(f"""<div class="card">
                      <div style="display:flex;justify-content:space-between;align-items:center">
                        <div>
                          <div style="font-size:13px;color:#E8EEF7;font-weight:600">{submit_cc} — {submit_month_str}</div>
                          <div style="font-size:11px;color:#6B7FA3">Submitted by {st_row['submitted_by']} on {st_row['submitted_at']}</div>
                        </div>
                        <span class="sp {badge_cls}">{st_row['status']}</span>
                      </div>
                      {"<div style='font-size:11px;color:#94A8C8;margin-top:8px'>Review notes: " + str(st_row['review_notes']) + "</div>" if st_row['review_notes'] else ""}
                    </div>""",unsafe_allow_html=True)
                    if st_row['status']=="Rejected":
                        st.warning("This submission was rejected. You may resubmit after correcting the issue.")
                    elif st_row['status']=="Pending Approval":
                        st.info("Awaiting Payroll Section review. You cannot resubmit until reviewed.")
                        st.stop()
                    elif st_row['status']=="Approved":
                        st.success("Already approved and processed for this month.")
                        st.stop()

                st.markdown(f'<div style="color:#6B7FA3;font-size:12px;margin:10px 0">Employees in {submit_cc}: <b style="color:#F0C96B">{len(cc_emps)}</b></div>',unsafe_allow_html=True)
                if len(cc_emps)>0:
                    st.dataframe(cc_emps,use_container_width=True,hide_index=True)
                    est_total = cc_emps['basic_salary'].fillna(0).sum()
                    st.markdown(f'<div style="font-size:12px;color:#94A8C8">Estimated total basic salary: <b style="color:#10B981">ETB {est_total:,.2f}</b> (final net calculated by Payroll Section after review)</div>',unsafe_allow_html=True)
                    if st.button("Submit Monthly Sheet for Approval",use_container_width=True):
                        conn=get_conn()
                        try:
                            conn.execute("""INSERT INTO payroll_submissions(cost_center,division,month,submitted_by,submitted_at,status,employee_count,total_net_amount)
                                VALUES(?,?,?,?,?,'Pending Approval',?,?)""",
                                (submit_cc,my_division,submit_month_str,st.session_state.uid,datetime.now().strftime("%Y-%m-%d %H:%M:%S"),len(cc_emps),float(est_total)))
                            conn.commit()
                            st.success(f"Submitted {submit_cc} — {submit_month_str} to Payroll Section for approval.")
                            st.rerun()
                        except sqlite3.IntegrityError:
                            st.error("A submission for this cost center and month already exists.")
                        finally: conn.close()
                else:
                    st.info("No employees currently in this cost center.")

    # ════════════════════════════════════════════════════════
    # PAYROLL — with autonomous day-off date calculation
    # ════════════════════════════════════════════════════════
    elif V=="Payroll":
        st.markdown('<div class="ey">Human Resources</div>',unsafe_allow_html=True)
        st.markdown('<div class="tl">Payroll Processing System</div>',unsafe_allow_html=True)
        conn=get_conn()
        el=pd.read_sql_query("SELECT emp_id,full_name,division,cost_center,basic_salary,weekly_dayoff FROM employees WHERE current_status='Active Deployment' ORDER BY emp_id LIMIT 5000",conn); conn.close()
        if len(el)==0: st.warning("No active employees."); st.stop()
        pt1,pt2,pt3,pt4=st.tabs(["Process Payroll","History","Day-Off Calendar","Cost Center Sheet"])
        with pt1:
            ps1,ps2=st.columns(2)
            with ps1:
                eo2={f"{r['emp_id']} — {r['full_name']}":r['emp_id'] for _,r in el.iterrows()}
                slbl=st.selectbox("Employee",list(eo2.keys()))
                seid=eo2[slbl]
            with ps2:
                yr=datetime.now().year
                pay_month=st.selectbox("Month",[f"{yr}-{m:02d}" for m in range(1,13)]+[f"{yr+1}-{m:02d}" for m in range(1,4)],index=datetime.now().month-1)
            conn=get_conn()
            er=pd.read_sql_query("SELECT * FROM employees WHERE emp_id=?",conn,params=(seid,))
            fines_df=pd.read_sql_query("""SELECT id,emp_id,COALESCE(month,'—') as month,issue_date,
                COALESCE(fine_type,'Disciplinary') as fine_type,COALESCE(fine_reason,'') as fine_reason,
                fine_days,fine_amount,letter_name,applied_to_payroll
                FROM fine_letters WHERE emp_id=? AND applied_to_payroll='No'
                AND COALESCE(record_status,'Active')='Active'""",conn,params=(seid,))
            # Absences for this month — excluding Cancelled and Compensated records automatically
            ab_count=pd.read_sql_query("""SELECT COUNT(*) as c FROM absent_records
                WHERE emp_id=? AND is_excused=0 AND substr(absent_date,1,7)=?
                AND COALESCE(record_status,'Active')='Active'""",conn,params=(seid,pay_month)).iloc[0]['c']
            # Leave days for this month — auto-collected from approved leave records, excluding Cancelled
            leave_this_month=pd.read_sql_query("""SELECT leave_type, SUM(days_taken) as total_days FROM leave_records
                WHERE emp_id=? AND status IN ('Approved') AND
                ((substr(start_date,1,7)=?) OR (substr(end_date,1,7)=?))
                GROUP BY leave_type""",conn,params=(seid,pay_month,pay_month))
            conn.close()
            leave_lookup_auto = dict(zip(leave_this_month['leave_type'], leave_this_month['total_days'])) if len(leave_this_month)>0 else {}
            auto_sick = int(leave_lookup_auto.get("Sick Leave",0))
            auto_annual = int(leave_lookup_auto.get("Annual Leave",0))
            auto_maternity = int(leave_lookup_auto.get("Maternity Leave",0))
            auto_mourning = int(leave_lookup_auto.get("Mourning Leave",0))
            auto_unpaid = int(leave_lookup_auto.get("Unpaid Leave",0))
            if len(er)==0: st.warning("Not found."); st.stop()
            er=er.iloc[0]; base=float(er.get("basic_salary",0) or 0)
            weekly_dayoff = er.get("weekly_dayoff") or "Sunday"
            pay_yr=int(pay_month.split("-")[0]); pay_mo=int(pay_month.split("-")[1])
            holidays=get_holidays(pay_yr)
            month_hols=[d for d in holidays if d.month==pay_mo]
            hol_count=len(month_hols)

            # ── AUTONOMOUS DAY-OFF CALCULATION ──
            # System knows the employee's chosen weekday and automatically finds
            # every matching calendar date in the selected month — no manual entry needed.
            dayoff_dates = get_dayoff_dates(weekly_dayoff, pay_yr, pay_mo)
            dayoff_count = len(dayoff_dates)
            absent_count=int(ab_count)

            st.markdown("<hr>",unsafe_allow_html=True)
            st.markdown('<div class="fs">Employee Weekly Day-Off Setting</div>',unsafe_allow_html=True)
            wd_disp_col,wd_edit_col=st.columns([2,1])
            with wd_disp_col:
                st.markdown(f"""<div style="background:#0D1526;border:1px solid rgba(56,189,248,0.25);border-radius:10px;padding:12px 16px">
                  <div style="font-size:9px;color:#6B7FA3;text-transform:uppercase;letter-spacing:.08em;margin-bottom:4px">Current Setting</div>
                  <div style="font-size:14px;color:#38BDF8;font-weight:600">Every {weekly_dayoff}</div>
                  <div style="font-size:11px;color:#94A8C8;margin-top:6px">System automatically found <b style="color:#F0C96B">{dayoff_count}</b> {weekly_dayoff}s in {pay_month}: {", ".join([d.strftime("%b %d") for d in dayoff_dates])}</div>
                </div>""",unsafe_allow_html=True)
            with wd_edit_col:
                st.write("")
                with st.popover("Change Day-Off",use_container_width=True):
                    wd_opts=["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"]
                    new_wd=st.selectbox("New weekly day-off",wd_opts,index=wd_opts.index(weekly_dayoff))
                    if st.button("Apply Change",use_container_width=True):
                        conn=get_conn()
                        conn.execute("UPDATE employees SET weekly_dayoff=? WHERE emp_id=?",(new_wd,seid))
                        conn.commit(); conn.close()
                        get_employee.clear()
                        st.success(f"Day-off changed to {new_wd}. System will recalculate dates automatically.")
                        st.rerun()

            st.markdown('<div class="fs">Allowances</div>',unsafe_allow_html=True)
            al1,al2,al3,al4=st.columns(4)
            with al1: transport=st.number_input("Transport Allowance",min_value=0.0,value=500.0,step=50.0)
            with al2: housing=st.number_input("Housing Allowance",min_value=0.0,value=300.0,step=50.0)
            with al3: other_al=st.number_input("Other Allowance",min_value=0.0,value=0.0,step=50.0)
            with al4: st.markdown(f'<div style="background:#0D1526;border:1px solid rgba(212,168,71,0.2);border-radius:8px;padding:10px;margin-top:20px"><div style="color:#6B7FA3;font-size:9px">BASIC SALARY</div><div style="color:#F0C96B;font-size:17px;font-family:Cinzel,serif;font-weight:700">ETB {base:,.2f}</div></div>',unsafe_allow_html=True)
            st.markdown('<div class="fs">Leave This Month — Auto-Collected from Leave Records (editable)</div>',unsafe_allow_html=True)
            lv1,lv2,lv3,lv4,lv5=st.columns(5)
            with lv1: sick_days=st.number_input("Sick Leave",min_value=0,max_value=30,value=auto_sick,step=1)
            with lv2: annual_days=st.number_input("Annual Leave",min_value=0,max_value=30,value=auto_annual,step=1)
            with lv3: mat_days=st.number_input("Maternity",min_value=0,max_value=90,value=auto_maternity,step=1)
            with lv4: mourning_days=st.number_input("Mourning",min_value=0,max_value=5,value=auto_mourning,step=1)
            with lv5: unpaid_days=st.number_input("Unpaid Leave",min_value=0,max_value=30,value=auto_unpaid,step=1)
            st.markdown('<div class="fs">Absence & Public Holidays (Auto-Detected — Cancelled and Compensated days excluded)</div>',unsafe_allow_html=True)
            ab1,ab2,ab3=st.columns(3)
            with ab1: absent_input=st.number_input("Absent Days (unexcused, deducted)",min_value=0,max_value=30,value=absent_count,step=1)
            with ab2: st.markdown(f'<div style="background:#0D1526;border:1px solid rgba(56,189,248,0.2);border-radius:8px;padding:10px;margin-top:20px"><div style="color:#6B7FA3;font-size:9px">DAY-OFF ({weekly_dayoff}s) — AUTO</div><div style="color:#38BDF8;font-size:17px;font-family:Cinzel,serif;font-weight:700">{dayoff_count} days  Paid</div></div>',unsafe_allow_html=True)
            with ab3: st.markdown(f'<div style="background:#0D1526;border:1px solid rgba(212,168,71,0.2);border-radius:8px;padding:10px;margin-top:20px"><div style="color:#6B7FA3;font-size:9px">PAID HOLIDAYS — AUTO</div><div style="color:#F0C96B;font-size:17px;font-family:Cinzel,serif;font-weight:700">{hol_count} days  Paid</div><div style="font-size:9px;color:#6B7FA3">{"  ".join([d.strftime("%b %d") for d in month_hols]) or "None this month"}</div></div>',unsafe_allow_html=True)
            total_fine_days=0; total_fine_amt=0.0; apply_all=False
            if len(fines_df)>0:
                st.markdown('<div class="fs">Pending Fine Letters (Auto-Collected)</div>',unsafe_allow_html=True)
                st.dataframe(fines_df[["id","month","issue_date","fine_reason","fine_type","fine_days","fine_amount","letter_name"]],use_container_width=True,hide_index=True)
                apply_all=st.checkbox("Apply all pending fines to this payroll (Cancelled and Compensated fines are excluded automatically)",value=True)
                if apply_all:
                    total_fine_days=int(fines_df["fine_days"].sum()); total_fine_amt=float(fines_df["fine_amount"].sum())
                    st.markdown(f'<div style="color:#FCA5A5;font-size:12px">Fine: <b>ETB {total_fine_amt:,.2f}</b> ({total_fine_days} days)</div>',unsafe_allow_html=True)
            else:
                st.success("No pending fines.")
            other_ded=st.number_input("Other Deductions (ETB)",min_value=0.0,value=0.0,step=50.0)
            notes_pay=st.text_area("Payroll Notes",placeholder="Optional...")
            net,tax,pen_emp,pen_er,daily_rate,gross=calc_pay(base,transport,housing,other_al,total_fine_amt,unpaid_days,absent_input,other_ded)
            st.markdown(f"""<div class="ps">
              <div style="font-family:'Cinzel',serif;font-size:12px;color:#D4A847;margin-bottom:10px;letter-spacing:.05em">AUTO PAYROLL — {pay_month}</div>
              <div class="pr"><span class="pl">Basic Salary</span><span class="pv">ETB {base:,.2f}</span></div>
              <div class="pr"><span class="pl">Transport</span><span class="pv">ETB {transport:,.2f}</span></div>
              <div class="pr"><span class="pl">Housing</span><span class="pv">ETB {housing:,.2f}</span></div>
              <div class="pr"><span class="pl">Other Allow.</span><span class="pv">ETB {other_al:,.2f}</span></div>
              <div class="pr"><span class="pl" style="font-weight:600">GROSS</span><span class="pv" style="color:#F0C96B;font-weight:600">ETB {gross:,.2f}</span></div>
              <div class="pr"><span class="pl">Income Tax</span><span class="pd">- ETB {tax:,.2f}</span></div>
              <div class="pr"><span class="pl">Pension Emp 7%</span><span class="pd">- ETB {pen_emp:,.2f}</span></div>
              <div class="pr"><span class="pl">Pension Er 11%</span><span class="pv" style="color:#38BDF8">ETB {pen_er:,.2f} (company)</span></div>
              <div class="pr"><span class="pl">Fines ({total_fine_days} days)</span><span class="pd">- ETB {total_fine_amt:,.2f}</span></div>
              <div class="pr"><span class="pl">Unpaid Leave ({unpaid_days} × ETB {daily_rate:.2f})</span><span class="pd">- ETB {daily_rate*unpaid_days:,.2f}</span></div>
              <div class="pr"><span class="pl">Absent ({absent_input} × ETB {daily_rate:.2f})</span><span class="pd">- ETB {daily_rate*absent_input:,.2f}</span></div>
              <div class="pr"><span class="pl">Other Deductions</span><span class="pd">- ETB {other_ded:,.2f}</span></div>
              <div class="pr"><span class="pl" style="color:#10B981">Paid Leave (Sick {sick_days}+Annual {annual_days}+Mat {mat_days}+Mourning {mourning_days})</span><span class="pv" style="color:#10B981"> Paid</span></div>
              <div class="pr"><span class="pl" style="color:#10B981">Day-Off {dayoff_count} days ({weekly_dayoff}) + Holidays {hol_count}</span><span class="pv" style="color:#10B981"> Paid</span></div>
              <div class="pr" style="border-top:1px solid rgba(16,185,129,0.3);margin-top:6px;padding-top:10px"><span class="pl" style="font-size:14px;font-weight:600;color:#E8EEF7">NET SALARY</span><span class="pn">ETB {net:,.2f}</span></div>
            </div>""",unsafe_allow_html=True)
            st.write("")
            sc1,sc2=st.columns(2)
            with sc1:
                if st.button("Save Payroll",use_container_width=True):
                    conn=get_conn(); cur=conn.cursor()
                    cur.execute("SELECT COUNT(*) FROM payroll WHERE emp_id=? AND month=?",(seid,pay_month))
                    existing_payroll_count=cur.fetchone()[0]
                    if existing_payroll_count>0:
                        conn.close()
                        st.error(f"Payroll for {seid} — {pay_month} has already been processed. The system blocks duplicate payroll entries automatically. Check Payroll History to review or correct the existing record.")
                    else:
                        conn.execute("""INSERT INTO payroll(emp_id,month,basic_salary,transport_allowance,housing_allowance,
                            other_allowance,income_tax,pension_employee,pension_employer,other_deductions,fine_amount,fine_days,
                            sick_leave_days,annual_leave_days,maternity_leave_days,mourning_leave_days,unpaid_leave_days,
                            absent_days,holiday_days,dayoff_days,gross_salary,net_salary,payment_status,notes,created_at)
                            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,'Processed',?,?)""",
                            (seid,pay_month,base,transport,housing,other_al,tax,pen_emp,pen_er,other_ded,
                             total_fine_amt,total_fine_days,sick_days,annual_days,mat_days,mourning_days,
                             unpaid_days,absent_input,hol_count,dayoff_count,gross,net,notes_pay,datetime.now().strftime("%Y-%m-%d")))
                        if apply_all and len(fines_df)>0:
                            ids=tuple(fines_df["id"].tolist())
                            if len(ids)==1: conn.execute("UPDATE fine_letters SET applied_to_payroll='Yes' WHERE id=?",(ids[0],))
                            else: conn.execute(f"UPDATE fine_letters SET applied_to_payroll='Yes' WHERE id IN {ids}")
                        conn.commit(); conn.close()
                        st.success(f"Payroll saved — Net: ETB {net:,.2f}"); st.balloons()
            with sc2:
                er_dict=er.to_dict()
                pay_row={"month":pay_month,"basic_salary":base,"transport_allowance":transport,"housing_allowance":housing,
                    "other_allowance":other_al,"gross_salary":gross,"income_tax":tax,"pension_employee":pen_emp,
                    "pension_employer":pen_er,"fine_days":total_fine_days,"fine_amount":total_fine_amt,
                    "unpaid_leave_days":unpaid_days,"absent_days":absent_input,"sick_leave_days":sick_days,
                    "annual_leave_days":annual_days,"maternity_leave_days":mat_days,"mourning_leave_days":mourning_days,
                    "holiday_days":hol_count,"dayoff_days":dayoff_count,"dayoff_weekday":weekly_dayoff,
                    "other_deductions":other_ded,"net_salary":net}
                html_slip=print_slip(er_dict,pay_row)
                b64_slip=base64.b64encode(html_slip.encode()).decode()
                st.markdown(f'<a href="data:text/html;base64,{b64_slip}" download="Payroll_{seid}_{pay_month}.html" target="_blank"><button style="width:100%;background:linear-gradient(135deg,#7B2FBE,#9333EA);color:#fff;border:none;border-radius:8px;padding:10px;font-weight:600;font-size:12px;cursor:pointer"> Print / Download Payroll Statement</button></a>',unsafe_allow_html=True)
        with pt2:
            conn=get_conn()
            hist=pd.read_sql_query("""SELECT p.emp_id,e.full_name,e.division,p.month,p.basic_salary,
                COALESCE(p.gross_salary,p.basic_salary) as gross_salary,p.net_salary,
                COALESCE(p.fine_amount,0) as fine_amount,COALESCE(p.absent_days,0) as absent_days,
                COALESCE(p.holiday_days,0) as holiday_days,COALESCE(p.dayoff_days,4) as dayoff_days,
                p.income_tax,p.pension_employee,p.payment_status,p.created_at
                FROM payroll p LEFT JOIN employees e ON p.emp_id=e.emp_id ORDER BY p.created_at DESC LIMIT 1000""",conn)
            conn.close()
            if len(hist)==0: st.info("No payroll records yet.")
            else:
                hf1,hf2=st.columns(2)
                with hf1: hist_month=st.selectbox("Filter by Month",["All"]+sorted(hist['month'].unique().tolist(),reverse=True))
                with hf2: hist_div=st.selectbox("Filter by Division",["All"]+get_division_list())
                fhist=hist.copy()
                if hist_month!="All": fhist=fhist[fhist['month']==hist_month]
                if hist_div!="All": fhist=fhist[fhist['division']==hist_div]
                hbuf=io.BytesIO()
                with pd.ExcelWriter(hbuf,engine="xlsxwriter") as w: fhist.to_excel(w,index=False,sheet_name="Payroll")
                st.download_button("Export Payroll History",hbuf.getvalue(),file_name=f"Payroll_{datetime.now().strftime('%Y%m%d')}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                st.dataframe(fhist,use_container_width=True,hide_index=True)
                # Bulk print all
                if hist_month!="All" and st.button(f"Print All Payslips for {hist_month}",use_container_width=True):
                    st.info(f"Use the individual employee payslip print button in 'Process Payroll' tab for each employee, or export to Excel above for a bulk register.")
        with pt3:
            st.markdown('<div style="color:#6B7FA3;font-size:12px;margin-bottom:12px">Each employee has one fixed weekly day-off (set in Employee Profile  Edit). The system <b style="color:#F0C96B">automatically</b> calculates which calendar dates match that weekday every month — no manual date entry required. Change the day below and dates recalculate instantly.</div>',unsafe_allow_html=True)
            dc1,dc2,dc3=st.columns(3)
            with dc1:
                cal_div=st.selectbox("Filter Division",["All"]+get_division_list(),key="cal_div")
            with dc2:
                cal_yr=st.selectbox("Year",[datetime.now().year,datetime.now().year+1],key="cal_yr")
            with dc3:
                cal_mo=st.selectbox("Month",list(range(1,13)),index=datetime.now().month-1,format_func=lambda m: calendar.month_name[m],key="cal_mo")
            conn=get_conn()
            if cal_div!="All":
                cal_emps=pd.read_sql_query("SELECT emp_id,full_name,division,weekly_dayoff FROM employees WHERE current_status='Active Deployment' AND division=? ORDER BY emp_id",conn,params=(cal_div,))
            else:
                cal_emps=pd.read_sql_query("SELECT emp_id,full_name,division,weekly_dayoff FROM employees WHERE current_status='Active Deployment' ORDER BY emp_id LIMIT 500",conn)
            conn.close()
            if len(cal_emps)>0:
                cal_emps['Day-Off Dates This Month']=cal_emps['weekly_dayoff'].apply(
                    lambda wd: ", ".join([d.strftime("%b %d (%a)") for d in get_dayoff_dates(wd or "Sunday",cal_yr,cal_mo)]))
                cal_emps['Total Days']=cal_emps['weekly_dayoff'].apply(lambda wd: count_dayoffs_in_month(wd or "Sunday",cal_yr,cal_mo))
                st.dataframe(cal_emps[['emp_id','full_name','division','weekly_dayoff','Total Days','Day-Off Dates This Month']],use_container_width=True,hide_index=True)
                cbuf=io.BytesIO()
                with pd.ExcelWriter(cbuf,engine="xlsxwriter") as w: cal_emps.to_excel(w,index=False,sheet_name="DayOff_Calendar")
                st.download_button("Export Day-Off Calendar",cbuf.getvalue(),file_name=f"DayOff_{cal_yr}_{cal_mo:02d}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.info("No active employees found for this filter.")


        # ════════════════════════════════════════════════════════
        # COST CENTER MONTHLY ATTENDANCE & PAYROLL SHEET
        # Official format: one row per employee, days of month as
        # columns, letter codes for status, grouped by cost center.
        # ════════════════════════════════════════════════════════
        def build_attendance_code(emp_id, dt, weekly_dayoff, holidays_dict, leave_lookup, absent_lookup):
            """Return single/double letter code for one employee on one date."""
            if dt in holidays_dict:
                return "H"
            wd_name = dt.strftime("%A")
            if wd_name == (weekly_dayoff or "Sunday"):
                return "D"
            key = (emp_id, dt.isoformat())
            if key in absent_lookup:
                return "X"
            if key in leave_lookup:
                lt = leave_lookup[key]
                if lt == "Sick Leave": return "S"
                if lt == "Maternity Leave": return "MT"
                if lt == "Mourning Leave": return "ML"
                if lt == "Unpaid Leave": return "EA"
                if lt == "Annual Leave": return "L"
                if lt == "Paternity Leave": return "PT"
                if lt == "Study Leave": return "SB"
                if lt == "Emergency Leave": return "EM"
            return "O"

        CODE_LEGEND = [
            ("O","Day Worked"), ("X","Lost Time (Absent)"), ("D","Day Off"),
            ("M","Morning Shift"), ("N","Night Shift"), ("S","Sick Leave"),
            ("H","Holiday"), ("MT","Maternity Leave"), ("ML","Mourning Leave (indicate relationship)"),
            ("EA","Excused Absence Without Pay"), ("CC","Court Case (indicate under Misc.)"),
            ("LB","Labour Business"), ("M","Miscellaneous (explain)"), ("K","Kebele"),
            ("H8","Holiday Overtime"),
        ]

        CODE_COLOR = {
            "O":"#0D1526","D":"#D4A847","H":"#A0522D","S":"#EF4444","MT":"#EC4899",
            "ML":"#A855F7","EA":"#F59E0B","X":"#DC2626","L":"#10B981","PT":"#06B6D4",
            "SB":"#6366F1","EM":"#F97316","H8":"#B45309"
        }

        with pt4:
            st.markdown('<div style="color:#6B7FA3;font-size:12px;margin-bottom:12px">Official monthly attendance and payroll register. Select a cost center and month — every employee in that cost center appears as one row with the full month as columns. All leave, absence, holiday and day-off records are automatically pulled in.</div>',unsafe_allow_html=True)

            sheet_c1,sheet_c2,sheet_c3=st.columns(3)
            with sheet_c1:
                all_ccs=get_cost_centers()
                cc_pick_opts = all_ccs['code'].tolist() if len(all_ccs)>0 else []
                if not cc_pick_opts:
                    st.warning("No cost centers exist yet. Create one in the Cost Centers module first.")
                    st.stop()
                sheet_cc=st.selectbox("Select Cost Center",cc_pick_opts,key="sheet_cc")
            with sheet_c2:
                sheet_yr=st.selectbox("Year",[datetime.now().year,datetime.now().year+1],key="sheet_yr")
            with sheet_c3:
                sheet_mo=st.selectbox("Month",list(range(1,13)),index=datetime.now().month-1,format_func=lambda m: calendar.month_name[m],key="sheet_mo")

            conn=get_conn()
            cc_info=pd.read_sql_query("SELECT * FROM cost_centers WHERE code=?",conn,params=(sheet_cc,))
            sheet_emps=pd.read_sql_query(
                "SELECT emp_id,full_name,job_title,weekly_dayoff,basic_salary FROM employees WHERE cost_center=? AND current_status != 'Terminated' ORDER BY emp_id",
                conn, params=(sheet_cc,))
            cc_division = cc_info.iloc[0]['division'] if len(cc_info)>0 else None
            conn.close()

            if len(sheet_emps)==0:
                st.info(f"No active employees are currently assigned to cost center {sheet_cc}.")

            if len(sheet_emps)>0:
                _,days_in_month=calendar.monthrange(sheet_yr,sheet_mo)
                month_dates=[date(sheet_yr,sheet_mo,d) for d in range(1,days_in_month+1)]
                holidays_dict=get_holidays(sheet_yr)

                emp_ids_tuple=tuple(sheet_emps['emp_id'].tolist())
                month_str=f"{sheet_yr}-{sheet_mo:02d}"

                conn=get_conn()
                if len(emp_ids_tuple)==1:
                    leave_df=pd.read_sql_query("SELECT emp_id,leave_type,start_date,end_date FROM leave_records WHERE emp_id=? AND status='Approved'",conn,params=(emp_ids_tuple[0],))
                    absent_df=pd.read_sql_query("SELECT emp_id,absent_date FROM absent_records WHERE emp_id=? AND is_excused=0 AND COALESCE(record_status,'Active')='Active'",conn,params=(emp_ids_tuple[0],))
                else:
                    leave_df=pd.read_sql_query(f"SELECT emp_id,leave_type,start_date,end_date FROM leave_records WHERE emp_id IN {emp_ids_tuple} AND status='Approved'",conn)
                    absent_df=pd.read_sql_query(f"SELECT emp_id,absent_date FROM absent_records WHERE emp_id IN {emp_ids_tuple} AND is_excused=0 AND COALESCE(record_status,'Active')='Active'",conn)
                conn.close()

                leave_lookup={}
                for _,lr in leave_df.iterrows():
                    try:
                        sd=datetime.strptime(lr['start_date'],"%Y-%m-%d").date()
                        ed=datetime.strptime(lr['end_date'],"%Y-%m-%d").date()
                        cur=sd
                        while cur<=ed:
                            leave_lookup[(lr['emp_id'],cur.isoformat())]=lr['leave_type']
                            cur+=timedelta(days=1)
                    except: pass
                absent_lookup={(ar['emp_id'],ar['absent_date']):True for _,ar in absent_df.iterrows()}

                # ── Build the grid ──
                rows_html=[]
                summary_rows=[]
                for _,emp in sheet_emps.iterrows():
                    eid=emp['emp_id']; wd=emp['weekly_dayoff'] or "Sunday"
                    day_codes=[]
                    counts={"D":0,"H":0,"S":0,"X":0,"O":0,"MT":0,"ML":0,"EA":0,"L":0,"PT":0,"SB":0,"EM":0}
                    for dt in month_dates:
                        code=build_attendance_code(eid,dt,wd,holidays_dict,leave_lookup,absent_lookup)
                        day_codes.append(code)
                        counts[code]=counts.get(code,0)+1
                    summary_rows.append({
                        "Employee ID":eid,"Full Name":emp['full_name'],"Job Title":emp['job_title'] or "",
                        "Cost Center":sheet_cc,"Day Worked (O)":counts["O"],"Day Off (D)":counts["D"],
                        "Holiday (H)":counts["H"],"Sick (S)":counts["S"],"Absent (X)":counts["X"],
                        "Maternity (MT)":counts["MT"],"Mourning (ML)":counts["ML"],"Unpaid (EA)":counts["EA"],
                        "Annual (L)":counts["L"],"Basic Salary":emp['basic_salary']
                    })
                    cells_html="".join([
                        f'<td style="background:{CODE_COLOR.get(c,"#0D1526")};color:#fff;text-align:center;font-size:10px;font-weight:600;padding:4px 2px;border:1px solid rgba(255,255,255,0.08);min-width:26px">{c}</td>'
                        for c in day_codes])
                    rows_html.append(f'''<tr>
                        <td style="padding:5px 8px;font-size:11px;color:#E8EEF7;font-weight:500;white-space:nowrap;border:1px solid rgba(255,255,255,0.08);background:#0D1526;position:sticky;left:0">{eid}</td>
                        <td style="padding:5px 8px;font-size:11px;color:#E8EEF7;white-space:nowrap;border:1px solid rgba(255,255,255,0.08);background:#0D1526">{emp['full_name']}</td>
                        {cells_html}
                    </tr>''')

                header_cells="".join([
                    f'<th style="background:#D4A847;color:#0D1526;text-align:center;font-size:9px;padding:4px 2px;border:1px solid rgba(255,255,255,0.15);min-width:26px">{dt.day}<br>{dt.strftime("%a")[:2].upper()}</th>'
                    for dt in month_dates])

                cc_name = cc_info.iloc[0]['name'] if len(cc_info)>0 else sheet_cc

                sheet_html = f'''
                <div style="background:#0A0F1E;border:1px solid rgba(212,168,71,0.25);border-radius:12px;padding:18px;overflow-x:auto">
                  <div style="font-family:'Cinzel',serif;font-size:16px;color:#F0C96B;text-align:center;margin-bottom:4px">YETEBABERUT GENERAL SERVICE PROVIDER</div>
                  <div style="text-align:center;font-size:11px;color:#94A8C8;margin-bottom:14px">Cost Center: <b style="color:#D4A847">{sheet_cc}</b> — {cc_name} &nbsp;|&nbsp; Period: {month_dates[0].strftime("%b %d, %Y")} to {month_dates[-1].strftime("%b %d, %Y")}</div>
                  <table style="border-collapse:collapse;width:100%;font-family:Inter,sans-serif">
                    <thead><tr>
                      <th style="background:#D4A847;color:#0D1526;padding:5px 8px;font-size:10px;position:sticky;left:0;min-width:80px">EMP ID</th>
                      <th style="background:#D4A847;color:#0D1526;padding:5px 8px;font-size:10px;position:sticky;left:0;min-width:140px">NAME</th>
                      {header_cells}
                    </tr></thead>
                    <tbody>{"".join(rows_html)}</tbody>
                  </table>
                </div>'''
                st.markdown(sheet_html, unsafe_allow_html=True)

                # ── Legend ──
                legend_html='<div style="margin-top:14px;display:grid;grid-template-columns:repeat(3,1fr);gap:6px;font-size:11px;color:#94A8C8">'
                for code,desc in CODE_LEGEND:
                    legend_html+=f'<div><b style="color:#D4A847">{code}</b> - {desc}</div>'
                legend_html+='</div>'
                st.markdown(legend_html, unsafe_allow_html=True)

                st.markdown("<hr>",unsafe_allow_html=True)

                # Summary data kept for export and net payroll calc, but not displayed here
                # to avoid duplicating the Payroll History page.
                summary_df=pd.DataFrame(summary_rows)

                # ── Auto-calculate net pay per employee in this cost center ──
                st.markdown('<div class="fs">Auto-Calculated Net Payroll — All Employees in This Cost Center</div>',unsafe_allow_html=True)
                payroll_rows=[]
                for sr in summary_rows:
                    basic=float(sr["Basic Salary"] or 0)
                    unpaid_d=sr["Unpaid (EA)"]
                    absent_d=sr["Absent (X)"]
                    net,tax,pen,pen_er,daily,gross=calc_pay(basic,0,0,0,0,unpaid_d,absent_d,0)
                    payroll_rows.append({
                        "Employee ID":sr["Employee ID"],"Full Name":sr["Full Name"],
                        "Basic Salary":basic,"Days Worked":sr["Day Worked (O)"],
                        "Day Off":sr["Day Off (D)"],"Holiday":sr["Holiday (H)"],
                        "Sick (Paid)":sr["Sick (S)"],"Maternity (Paid)":sr["Maternity (MT)"],
                        "Mourning (Paid)":sr["Mourning (ML)"],"Annual (Paid)":sr["Annual (L)"],
                        "Unpaid Leave":unpaid_d,"Absent (Deducted)":absent_d,
                        "Income Tax":tax,"Pension (7%)":pen,"Net Salary":net
                    })
                payroll_df=pd.DataFrame(payroll_rows)
                st.dataframe(payroll_df,use_container_width=True,hide_index=True)

                total_net=payroll_df["Net Salary"].sum() if len(payroll_df)>0 else 0
                st.markdown(f'<div class="ps"><div class="pr"><span class="pl" style="font-size:14px;font-weight:600;color:#E8EEF7">TOTAL NET PAYROLL — {sheet_cc} ({month_str})</span><span class="pn">ETB {total_net:,.2f}</span></div></div>',unsafe_allow_html=True)

                st.markdown("<hr>",unsafe_allow_html=True)
                st.markdown('<div class="fs">Export, Save, and Print</div>',unsafe_allow_html=True)

                # Build Excel bytes immediately — do not wrap in a column or button click
                xbuf=io.BytesIO()
                with pd.ExcelWriter(xbuf,engine="xlsxwriter") as w:
                    wb=w.book
                    # Sheet 1: Attendance grid codes per day
                    att_rows=[]
                    for _,emp in sheet_emps.iterrows():
                        row_d={"EMP ID":emp['emp_id'],"NAME":emp['full_name']}
                        for dt in month_dates:
                            code_val=build_attendance_code(emp['emp_id'],dt,
                                emp['weekly_dayoff'] or "Sunday",
                                holidays_dict,leave_lookup,absent_lookup)
                            row_d[dt.strftime("%d %a")]=code_val
                        att_rows.append(row_d)
                    att_df=pd.DataFrame(att_rows)
                    att_df.to_excel(w,index=False,sheet_name="Attendance_Sheet")
                    ws_att=w.sheets["Attendance_Sheet"]
                    hdr_fmt=wb.add_format({"bold":True,"bg_color":"#D4A847","font_color":"#0D1526","border":1,"align":"center"})
                    emp_fmt=wb.add_format({"bold":True,"bg_color":"#0D1526","font_color":"#F0C96B","border":1})
                    o_fmt=wb.add_format({"align":"center","border":1,"bg_color":"#0D1526","font_color":"#E8EEF7"})
                    x_fmt=wb.add_format({"align":"center","border":1,"bg_color":"#DC2626","font_color":"#fff","bold":True})
                    d_fmt=wb.add_format({"align":"center","border":1,"bg_color":"#D4A847","font_color":"#0D1526","bold":True})
                    h_fmt=wb.add_format({"align":"center","border":1,"bg_color":"#92400E","font_color":"#fff"})
                    for ci in range(len(att_df.columns)):
                        ws_att.write(0,ci,att_df.columns[ci],hdr_fmt)
                    for ri in range(len(att_df)):
                        for ci,col in enumerate(att_df.columns):
                            val=str(att_df.iloc[ri,ci])
                            if ci<2: ws_att.write(ri+1,ci,val,emp_fmt)
                            elif val=="X": ws_att.write(ri+1,ci,val,x_fmt)
                            elif val=="D": ws_att.write(ri+1,ci,val,d_fmt)
                            elif val=="H": ws_att.write(ri+1,ci,val,h_fmt)
                            else: ws_att.write(ri+1,ci,val,o_fmt)
                    ws_att.set_column(0,0,10); ws_att.set_column(1,1,22)
                    for ci in range(2,len(att_df.columns)): ws_att.set_column(ci,ci,5)
                    # Sheet 2: Payroll summary
                    payroll_df.to_excel(w,index=False,sheet_name="Payroll_Summary")
                    ws_p=w.sheets["Payroll_Summary"]
                    ph_fmt=wb.add_format({"bold":True,"bg_color":"#0D1526","font_color":"#D4A847","border":1,"font_size":10})
                    pd_fmt=wb.add_format({"bg_color":"#060B18","font_color":"#E8EEF7","border":1,"font_size":10})
                    for ci,col in enumerate(payroll_df.columns):
                        ws_p.write(0,ci,col,ph_fmt); ws_p.set_column(ci,ci,max(len(col)+2,12))
                    for ri in range(len(payroll_df)):
                        for ci,col in enumerate(payroll_df.columns):
                            ws_p.write(ri+1,ci,payroll_df.iloc[ri,ci],pd_fmt)
                xbuf_val = xbuf.getvalue()

                ec1,ec2,ec3=st.columns(3)
                with ec1:
                    st.download_button(
                        "Export Attendance & Payroll to Excel",
                        data=xbuf_val,
                        file_name=f"{sheet_cc}_{month_str}_AttendancePayroll.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True)
                with ec2:
                    if st.button("Save All to Payroll History",use_container_width=True):
                        conn=get_conn(); cur=conn.cursor()
                        saved_count=0; skipped_count=0
                        for sr in payroll_rows:
                            cur.execute("SELECT COUNT(*) FROM payroll WHERE emp_id=? AND month=?",(sr["Employee ID"],month_str))
                            if cur.fetchone()[0]>0:
                                skipped_count+=1
                                continue
                            conn.execute("""INSERT INTO payroll(emp_id,month,basic_salary,transport_allowance,housing_allowance,
                                other_allowance,income_tax,pension_employee,pension_employer,other_deductions,fine_amount,fine_days,
                                sick_leave_days,annual_leave_days,maternity_leave_days,mourning_leave_days,unpaid_leave_days,
                                absent_days,holiday_days,dayoff_days,gross_salary,net_salary,payment_status,notes,created_at)
                                VALUES(?,?,?,0,0,0,?,?,?,0,0,0,?,?,?,?,?,?,?,?,?,?,'Processed',?,?)""",
                                (sr["Employee ID"],month_str,sr["Basic Salary"],sr["Income Tax"],sr["Pension (7%)"],
                                 sr["Basic Salary"]*0.11,
                                 next((s["Sick (S)"] for s in summary_rows if s["Employee ID"]==sr["Employee ID"]),0),
                                 next((s["Annual (L)"] for s in summary_rows if s["Employee ID"]==sr["Employee ID"]),0),
                                 next((s["Maternity (MT)"] for s in summary_rows if s["Employee ID"]==sr["Employee ID"]),0),
                                 next((s["Mourning (ML)"] for s in summary_rows if s["Employee ID"]==sr["Employee ID"]),0),
                                 sr["Unpaid Leave"],sr["Absent (Deducted)"],
                                 next((s["Holiday (H)"] for s in summary_rows if s["Employee ID"]==sr["Employee ID"]),0),
                                 next((s["Day Off (D)"] for s in summary_rows if s["Employee ID"]==sr["Employee ID"]),0),
                                 sr["Basic Salary"],sr["Net Salary"],
                                 f"Bulk cost center sheet — {sheet_cc}",datetime.now().strftime("%Y-%m-%d")))
                            saved_count+=1
                        conn.commit(); conn.close()
                        if skipped_count>0:
                            st.warning(f"Saved {saved_count} new payroll record(s) for {sheet_cc} ({month_str}). Skipped {skipped_count} employee(s) already processed this month — duplicates blocked automatically.")
                        else:
                            st.success(f"Saved {saved_count} payroll records for {sheet_cc} ({month_str}) to history.")
                with ec3:
                    print_rows="".join([f'''<tr>
                        <td style="padding:5px 8px;font-size:11px;border:1px solid #ddd">{sr["Employee ID"]}</td>
                        <td style="padding:5px 8px;font-size:11px;border:1px solid #ddd">{sr["Full Name"]}</td>
                        <td style="padding:5px 8px;font-size:11px;border:1px solid #ddd;text-align:right">{sr["Basic Salary"]:,.2f}</td>
                        <td style="padding:5px 8px;font-size:11px;border:1px solid #ddd;text-align:right;color:#c0392b">-{sr["Income Tax"]:,.2f}</td>
                        <td style="padding:5px 8px;font-size:11px;border:1px solid #ddd;text-align:right;color:#c0392b">-{sr["Pension (7%)"]:,.2f}</td>
                        <td style="padding:5px 8px;font-size:11px;border:1px solid #ddd;text-align:right;font-weight:bold">{sr["Net Salary"]:,.2f}</td>
                    </tr>''' for sr in payroll_rows])
                    print_html=f'''<!DOCTYPE html><html><head><meta charset="UTF-8"><style>
                    body{{font-family:Arial,sans-serif;margin:0;padding:20px;color:#000}}
                    h2{{text-align:center;color:#0D1526;margin-bottom:2px}}
                    .sub{{text-align:center;color:#666;font-size:12px;margin-bottom:16px}}
                    table{{width:100%;border-collapse:collapse}}
                    th{{background:#0D1526;color:#D4A847;padding:6px 8px;font-size:11px;text-align:left;border:1px solid #ddd}}
                    tfoot td{{font-weight:bold;background:#f5f5f5;padding:8px}}
                    @media print{{body{{margin:0}}}}
                    </style></head><body>
                    <h2>YETEBABERUT GENERAL SERVICE PROVIDER</h2>
                    <div class="sub">Cost Center Payroll Register — {sheet_cc} ({cc_name}) — {month_str}</div>
                    <table>
                      <tr><th>Employee ID</th><th>Full Name</th><th style="text-align:right">Basic Salary</th><th style="text-align:right">Tax</th><th style="text-align:right">Pension</th><th style="text-align:right">Net Salary</th></tr>
                      {print_rows}
                      <tfoot><tr><td colspan="5" style="text-align:right;padding:8px">TOTAL NET PAYROLL</td><td style="text-align:right;padding:8px">ETB {total_net:,.2f}</td></tr></tfoot>
                    </table>
                    <script>window.onload=function(){{window.print()}}</script>
                    </body></html>'''
                    b64_print=base64.b64encode(print_html.encode()).decode()
                    st.markdown(f'<a href="data:text/html;base64,{b64_print}" download="{sheet_cc}_{month_str}_Register.html" target="_blank"><button style="width:100%;background:linear-gradient(135deg,#7B2FBE,#9333EA);color:#fff;border:none;border-radius:8px;padding:10px;font-weight:600;font-size:12px;cursor:pointer">Print Cost Center Payroll Register</button></a>',unsafe_allow_html=True)


    # ════════════════════════════════════════════════════════
    # PAYROLL APPROVALS — Payroll Section reviews submissions
    # from division Supervisors before salaries are released.
    # ════════════════════════════════════════════════════════
    elif V=="Payroll Approvals":
        st.markdown('<div class="ey">Payroll Section</div>',unsafe_allow_html=True)
        st.markdown('<div class="tl">Payroll Approval Workflow</div>',unsafe_allow_html=True)
        st.markdown("""<div class="card card-gold">
          <div style="font-size:12px;color:#C8D8F0;line-height:1.7">
            Division supervisors submit each cost center's monthly attendance sheet here for review.
            Salaries are released only after the Payroll Section approves the submission.
          </div></div>""",unsafe_allow_html=True)

        conn=get_conn()
        pending=pd.read_sql_query("SELECT * FROM payroll_submissions WHERE status='Pending Approval' ORDER BY submitted_at ASC",conn)
        approved=pd.read_sql_query("SELECT * FROM payroll_submissions WHERE status='Approved' ORDER BY reviewed_at DESC LIMIT 200",conn)
        rejected=pd.read_sql_query("SELECT * FROM payroll_submissions WHERE status='Rejected' ORDER BY reviewed_at DESC LIMIT 200",conn)
        conn.close()

        st.markdown(f"""<div class="mg" style="grid-template-columns:repeat(3,1fr)">
          <div class="mb mg-amber"><div class="ml ml-amber">Pending Review</div><div class="mv">{len(pending)}</div></div>
          <div class="mb mg-green"><div class="ml ml-green">Approved</div><div class="mv">{len(approved)}</div></div>
          <div class="mb mg-red"><div class="ml ml-red">Rejected</div><div class="mv">{len(rejected)}</div></div>
        </div>""",unsafe_allow_html=True)

        pa1,pa2,pa3=st.tabs(["Pending Review","Approved History","Rejected History"])

        with pa1:
            if len(pending)==0:
                st.info("No submissions waiting for review.")
            else:
                for _,sub in pending.iterrows():
                    with st.expander(f"{sub['cost_center']} — {sub['month']} — submitted by {sub['submitted_by']} ({sub['employee_count']} employees)",expanded=False):
                        st.markdown(f"""<div style="font-size:12px;color:#94A8C8;margin-bottom:10px">
                          Division: <b style="color:#F0C96B">{sub['division']}</b> &nbsp;|&nbsp;
                          Submitted: {sub['submitted_at']} &nbsp;|&nbsp;
                          Estimated total: <b style="color:#10B981">ETB {sub['total_net_amount']:,.2f}</b>
                        </div>""",unsafe_allow_html=True)

                        conn=get_conn()
                        review_emps=pd.read_sql_query(
                            "SELECT emp_id,full_name,job_title,basic_salary,weekly_dayoff,current_status FROM employees WHERE cost_center=?",
                            conn,params=(sub['cost_center'],))
                        conn.close()
                        st.dataframe(review_emps,use_container_width=True,hide_index=True)

                        review_notes=st.text_area("Review Notes",key=f"rn_{sub['id']}",placeholder="Comments about this submission...")
                        rc1,rc2=st.columns(2)
                        with rc1:
                            if st.button("Approve and Process Payroll",key=f"appr_{sub['id']}",use_container_width=True):
                                conn=get_conn()
                                month_str=sub['month']
                                for _,emp in review_emps.iterrows():
                                    basic=float(emp['basic_salary'] or 0)
                                    net,tax,pen,pen_er,daily,gross=calc_pay(basic,0,0,0,0,0,0,0)
                                    conn.execute("""INSERT INTO payroll(emp_id,month,basic_salary,transport_allowance,housing_allowance,
                                        other_allowance,income_tax,pension_employee,pension_employer,other_deductions,fine_amount,fine_days,
                                        sick_leave_days,annual_leave_days,maternity_leave_days,mourning_leave_days,unpaid_leave_days,
                                        absent_days,holiday_days,dayoff_days,gross_salary,net_salary,payment_status,notes,created_at)
                                        VALUES(?,?,?,0,0,0,?,?,?,0,0,0,0,0,0,0,0,0,0,4,?,?,'Processed',?,?)""",
                                        (emp['emp_id'],month_str,basic,tax,pen,pen_er,gross,net,
                                         f"Approved via Payroll Approvals workflow — Cost Center {sub['cost_center']}",
                                         datetime.now().strftime("%Y-%m-%d")))
                                conn.execute("UPDATE payroll_submissions SET status='Approved',reviewed_by=?,reviewed_at=?,review_notes=? WHERE id=?",
                                    (st.session_state.uid,datetime.now().strftime("%Y-%m-%d %H:%M:%S"),review_notes,sub['id']))
                                conn.commit(); conn.close()
                                st.success(f"Approved. Payroll processed for {len(review_emps)} employees in {sub['cost_center']}.")
                                st.rerun()
                        with rc2:
                            if st.button("Reject Submission",key=f"rej_{sub['id']}",use_container_width=True):
                                if not review_notes:
                                    st.error("Please provide review notes explaining the rejection.")
                                else:
                                    conn=get_conn()
                                    conn.execute("UPDATE payroll_submissions SET status='Rejected',reviewed_by=?,reviewed_at=?,review_notes=? WHERE id=?",
                                        (st.session_state.uid,datetime.now().strftime("%Y-%m-%d %H:%M:%S"),review_notes,sub['id']))
                                    conn.commit(); conn.close()
                                    st.warning(f"Submission rejected. Supervisor will be notified to resubmit.")
                                    st.rerun()

        with pa2:
            if len(approved)==0:
                st.info("No approved submissions yet.")
            else:
                st.dataframe(approved[['cost_center','division','month','submitted_by','reviewed_by','reviewed_at','employee_count','total_net_amount']],use_container_width=True,hide_index=True)
                abuf=io.BytesIO()
                with pd.ExcelWriter(abuf,engine="xlsxwriter") as w: approved.to_excel(w,index=False,sheet_name="Approved")
                st.download_button("Export Approved History",abuf.getvalue(),file_name=f"Approved_Payroll_{datetime.now().strftime('%Y%m%d')}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        with pa3:
            if len(rejected)==0:
                st.info("No rejected submissions.")
            else:
                st.dataframe(rejected[['cost_center','division','month','submitted_by','reviewed_by','reviewed_at','review_notes']],use_container_width=True,hide_index=True)

    # ════════════════════════════════════════════════════════
    # LEAVE & DISCIPLINE
    # ════════════════════════════════════════════════════════
    elif V=="Leave & Discipline":
        st.markdown('<div class="ey">HR Management</div>',unsafe_allow_html=True)
        st.markdown('<div class="tl">Leave Records, Fine Letters & Absences</div>',unsafe_allow_html=True)
        conn=get_conn()
        emp_opts=pd.read_sql_query("SELECT emp_id,full_name FROM employees WHERE current_status='Active Deployment' ORDER BY emp_id LIMIT 5000",conn); conn.close()
        elo={f"{r['emp_id']} — {r['full_name']}":r['emp_id'] for _,r in emp_opts.iterrows()}
        lf1,lf2,lf3,lf4=st.tabs(["Leave Records","Fine Letters","Absent Records","Pending Dept Head Approval"])
        with lf1:
            st.markdown('<div class="fs">Submit New Leave</div>',unsafe_allow_html=True)
            with st.form("lf"):
                lc1,lc2,lc3=st.columns(3)
                with lc1: l_emp=st.selectbox("Employee",list(elo.keys()))
                with lc2:
                    ltypes=["Sick Leave","Annual Leave","Maternity Leave","Paternity Leave","Mourning Leave","Unpaid Leave","Emergency Leave","Study Leave"]
                    l_type=st.selectbox("Leave Type",ltypes)
                with lc3: l_status=st.selectbox("Status",["Approved","Pending","Rejected"])
                lc4,lc5=st.columns(2)
                with lc4: l_start=st.date_input("Start",value=date.today())
                with lc5: l_end=st.date_input("End",value=date.today())
                l_by = f"{st.session_state.role}: {st.session_state.full_name or st.session_state.uid}"
                st.markdown(f'<div style="font-size:11px;color:#6B7FA3;margin:-4px 0 8px">Recorded by: <b style="color:#10B981">{l_by}</b></div>',unsafe_allow_html=True)
                l_notes=st.text_area("Notes",placeholder="Reason...")
                if st.form_submit_button("Save Leave",use_container_width=True):
                    l_eid=elo[l_emp]
                    conn=get_conn(); cur=conn.cursor()
                    # Block true duplicates automatically: same employee, overlapping dates, not cancelled.
                    cur.execute("""SELECT COUNT(*) FROM leave_records WHERE emp_id=? AND status != 'Cancelled'
                        AND start_date<=? AND end_date>=?""",(l_eid,str(l_end),str(l_start)))
                    overlap_count=cur.fetchone()[0]
                    if overlap_count>0:
                        conn.close()
                        st.error(f"{l_emp} already has a leave record overlapping {l_start} to {l_end}. Edit the existing record below instead of creating a duplicate.")
                        st.stop()
                    cur.execute("SELECT basic_salary FROM employees WHERE emp_id=?",(l_eid,))
                    sr=cur.fetchone(); dr=float(sr[0])/26 if sr and sr[0] else 0
                    days=max((l_end-l_start).days+1,0)
                    is_paid=0 if l_type=="Unpaid Leave" else 1; ded=0.0 if is_paid else round(dr*days,2)
                    conn.execute("INSERT INTO leave_records(emp_id,leave_type,start_date,end_date,days_taken,is_paid,daily_rate,deduction_amount,approved_by,status,notes,created_at)VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",
                        (l_eid,l_type,str(l_start),str(l_end),days,is_paid,round(dr,2),ded,l_by,l_status,l_notes,datetime.now().strftime("%Y-%m-%d")))
                    conn.commit(); conn.close()
                    if l_type in ["Maternity Leave","Sick Leave"] and days>=7:
                        conn=get_conn(); conn.execute("UPDATE employees SET current_status='On Leave' WHERE emp_id=?",(l_eid,)); conn.commit(); conn.close()
                        st.cache_data.clear()
                    st.success(f"{l_type}: {days} days {'(Paid)' if is_paid else f'(Unpaid — ETB {ded:,.2f})'}"); st.rerun()

            st.markdown("<hr>",unsafe_allow_html=True)
            st.markdown('<div class="fs">Manage Existing Leave Records</div>',unsafe_allow_html=True)
            st.markdown('<div style="font-size:11px;color:#6B7FA3;margin-bottom:8px">Edit dates, change status, or cancel a leave record. Cancelling removes its effect from payroll.</div>',unsafe_allow_html=True)
            conn=get_conn()
            lv_full=pd.read_sql_query("""SELECT lr.id,lr.emp_id,e.full_name,e.division,lr.leave_type,lr.start_date,lr.end_date,
                lr.days_taken,lr.is_paid,lr.deduction_amount,lr.status,lr.approved_by,lr.notes
                FROM leave_records lr LEFT JOIN employees e ON lr.emp_id=e.emp_id
                WHERE lr.status != 'Cancelled' ORDER BY lr.created_at DESC LIMIT 500""",conn); conn.close()

            if len(lv_full)>0:
                lv_opts={f"ID {r['id']} — {r['emp_id']} {r['full_name']} — {r['leave_type']} ({r['start_date']} to {r['end_date']}) [{r['status']}]":r['id'] for _,r in lv_full.iterrows()}
                sel_lv_label=st.selectbox("Select Leave Record to Edit or Cancel",list(lv_opts.keys()),key="sel_lv_edit")
                sel_lv_id=lv_opts[sel_lv_label]
                lv_row=lv_full[lv_full['id']==sel_lv_id].iloc[0]

                with st.form(f"edit_leave_{sel_lv_id}"):
                    el1,el2,el3=st.columns(3)
                    with el1:
                        ltypes2=["Sick Leave","Annual Leave","Maternity Leave","Paternity Leave","Mourning Leave","Unpaid Leave","Emergency Leave","Study Leave"]
                        e_ltype=st.selectbox("Leave Type",ltypes2,index=ltypes2.index(lv_row['leave_type']) if lv_row['leave_type'] in ltypes2 else 0,key=f"elt_{sel_lv_id}")
                    with el2:
                        e_lstart=st.date_input("Start Date",value=datetime.strptime(lv_row['start_date'],"%Y-%m-%d").date(),key=f"els_{sel_lv_id}")
                    with el3:
                        e_lend=st.date_input("End Date",value=datetime.strptime(lv_row['end_date'],"%Y-%m-%d").date(),key=f"ele_{sel_lv_id}")
                    el4,el5=st.columns(2)
                    with el4:
                        status_opts3=["Approved","Pending","Rejected","Pending Dept Head Approval"]
                        e_lstatus=st.selectbox("Status",status_opts3,index=status_opts3.index(lv_row['status']) if lv_row['status'] in status_opts3 else 0,key=f"elst_{sel_lv_id}")
                    with el5:
                        e_lnotes=st.text_input("Notes",value=lv_row['notes'] or "",key=f"eln_{sel_lv_id}")
                    elc1,elc2=st.columns(2)
                    with elc1:
                        if st.form_submit_button("Save Changes",use_container_width=True):
                            new_days=max((e_lend-e_lstart).days+1,0)
                            conn=get_conn(); cur=conn.cursor()
                            cur.execute("SELECT basic_salary FROM employees WHERE emp_id=?",(lv_row['emp_id'],))
                            sr=cur.fetchone(); dr2=float(sr[0])/26 if sr and sr[0] else 0
                            new_is_paid = 0 if e_ltype=="Unpaid Leave" else 1
                            new_ded = 0.0 if new_is_paid else round(dr2*new_days,2)
                            conn.execute("""UPDATE leave_records SET leave_type=?,start_date=?,end_date=?,days_taken=?,
                                is_paid=?,deduction_amount=?,status=?,notes=?,edited_by=?,edited_at=? WHERE id=?""",
                                (e_ltype,str(e_lstart),str(e_lend),new_days,new_is_paid,new_ded,e_lstatus,e_lnotes,
                                 st.session_state.uid,datetime.now().strftime("%Y-%m-%d %H:%M:%S"),sel_lv_id))
                            conn.commit(); conn.close()
                            st.success("Leave record updated."); st.rerun()
                    with elc2:
                        cancel_reason_lv = st.text_input("Cancellation Reason (required to cancel)",key=f"cr_lv_{sel_lv_id}")
                        if st.form_submit_button("Cancel This Leave",use_container_width=True):
                            if not cancel_reason_lv:
                                st.error("Please provide a cancellation reason.")
                            else:
                                conn=get_conn()
                                conn.execute("UPDATE leave_records SET status='Cancelled',cancelled_by=?,cancelled_at=?,cancel_reason=? WHERE id=?",
                                    (st.session_state.uid,datetime.now().strftime("%Y-%m-%d %H:%M:%S"),cancel_reason_lv,sel_lv_id))
                                conn.commit(); conn.close()
                                st.warning("Leave record cancelled. It will no longer affect payroll."); st.rerun()
            else:
                st.info("No leave records to manage yet.")

            st.markdown("<hr>",unsafe_allow_html=True)
            conn=get_conn()
            lv=pd.read_sql_query("""SELECT lr.emp_id,e.full_name,e.division,lr.leave_type,lr.start_date,lr.end_date,lr.days_taken,
                CASE WHEN lr.is_paid=1 THEN 'Paid' ELSE 'Unpaid' END as paid,lr.deduction_amount,lr.status
                FROM leave_records lr LEFT JOIN employees e ON lr.emp_id=e.emp_id ORDER BY lr.created_at DESC LIMIT 1000""",conn); conn.close()
            if len(lv)>0:
                buf3=io.BytesIO()
                with pd.ExcelWriter(buf3,engine="xlsxwriter") as w: lv.to_excel(w,index=False,sheet_name="Leave")
                st.download_button("Export",buf3.getvalue(),file_name=f"Leave_{datetime.now().strftime('%Y%m%d')}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                st.dataframe(lv,use_container_width=True,hide_index=True)

        with lf2:
            st.markdown('<div class="fs">Issue New Fine Letter</div>',unsafe_allow_html=True)
            with st.form("ff"):
                fc1,fc2,fc3,fc4=st.columns(4)
                with fc1: f_emp=st.selectbox("Employee",list(elo.keys()))
                with fc2: fine_month=st.selectbox("Fine Month",[f"{datetime.now().year}-{m:02d}" for m in range(1,13)],index=datetime.now().month-1)
                with fc3: f_date=st.date_input("Issue Date",value=date.today())
                with fc4: f_days=st.number_input("Fine Days",min_value=0,max_value=30,step=1)
                fc5,fc6=st.columns(2)
                with fc5:
                    ftypes=["Disciplinary","Misconduct","Late Arrival","Unauthorized Absence","Policy Violation","Performance","Other"]
                    f_type=st.selectbox("Fine Type",ftypes)
                with fc6: f_reason=st.text_input("Reason",placeholder="Brief description")
                f_details=st.text_area("Full Details",placeholder="Detailed explanation...")
                st.markdown('<div style="font-size:10px;color:#38BDF8;margin-bottom:5px">Attach scanned fine letter — PDF or JPG</div>',unsafe_allow_html=True)
                f_letter=st.file_uploader("Upload Signed Fine Letter",type=["pdf","jpg","jpeg","png"])
                if st.form_submit_button("Issue Fine Letter",use_container_width=True):
                    f_eid=elo[f_emp]
                    conn=get_conn(); cur=conn.cursor()
                    cur.execute("SELECT basic_salary FROM employees WHERE emp_id=?",(f_eid,))
                    sr=cur.fetchone(); dr=float(sr[0])/26 if sr and sr[0] else 0
                    fine_amt=round(dr*f_days,2)
                    full_reason=f"{f_type}: {f_reason}\n{f_details}".strip()
                    ln=f_letter.name if f_letter else None
                    ld=sqlite3.Binary(f_letter.read()) if f_letter else None
                    conn.execute("INSERT INTO fine_letters(emp_id,month,issue_date,fine_reason,fine_type,fine_days,fine_amount,letter_name,letter_data,applied_to_payroll,record_status,created_at)VALUES(?,?,?,?,?,?,?,?,?,'No','Active',?)",
                        (f_eid,fine_month,str(f_date),full_reason,f_type,f_days,fine_amt,ln,ld,datetime.now().strftime("%Y-%m-%d")))
                    conn.commit(); conn.close()
                    st.success(f"Fine issued: {f_days} days = ETB {fine_amt:,.2f} — auto-deducts from {fine_month} payroll"); st.rerun()

            st.markdown("<hr>",unsafe_allow_html=True)
            st.markdown('<div class="fs">Manage Fines — Cancel or Compensate</div>',unsafe_allow_html=True)
            st.markdown('<div style="font-size:11px;color:#6B7FA3;margin-bottom:8px"><b>Cancel</b> voids the fine completely (e.g. issued by mistake). <b>Compensate</b> lets the employee work off the fine days instead of having pay deducted.</div>',unsafe_allow_html=True)
            conn=get_conn()
            fl_manage=pd.read_sql_query("""SELECT fl.id,fl.emp_id,e.full_name,fl.month,fl.fine_type,fl.fine_days,fl.fine_amount,
                COALESCE(fl.record_status,'Active') as record_status,fl.applied_to_payroll,COALESCE(fl.compensated_days,0) as compensated_days
                FROM fine_letters fl LEFT JOIN employees e ON fl.emp_id=e.emp_id
                WHERE COALESCE(fl.record_status,'Active') != 'Cancelled' ORDER BY fl.created_at DESC LIMIT 500""",conn); conn.close()

            if len(fl_manage)>0:
                fl_opts={f"ID {r['id']} — {r['emp_id']} {r['full_name']} — {r['fine_type']} ({r['fine_days']} days, ETB {r['fine_amount']:,.2f}) [{r['record_status']}]":r['id'] for _,r in fl_manage.iterrows()}
                sel_fl_label=st.selectbox("Select Fine to Manage",list(fl_opts.keys()),key="sel_fl_manage")
                sel_fl_id=fl_opts[sel_fl_label]
                fl_row=fl_manage[fl_manage['id']==sel_fl_id].iloc[0]

                fmc1,fmc2,fmc3=st.columns(3)
                with fmc1:
                    cancel_reason_fl=st.text_input("Cancellation Reason",key=f"cr_fl_{sel_fl_id}")
                    if st.button("Cancel This Fine",key=f"cancel_fl_{sel_fl_id}",use_container_width=True):
                        if not cancel_reason_fl:
                            st.error("Please provide a cancellation reason.")
                        else:
                            conn=get_conn()
                            conn.execute("UPDATE fine_letters SET record_status='Cancelled',applied_to_payroll='Cancelled',cancelled_by=?,cancelled_at=?,cancel_reason=? WHERE id=?",
                                (st.session_state.uid,datetime.now().strftime("%Y-%m-%d %H:%M:%S"),cancel_reason_fl,sel_fl_id))
                            conn.commit(); conn.close()
                            st.warning("Fine cancelled — will not affect payroll."); st.rerun()
                with fmc2:
                    comp_days_fl=st.number_input("Days to Compensate (work off)",min_value=0,max_value=int(fl_row['fine_days']),step=1,key=f"compd_fl_{sel_fl_id}")
                    comp_notes_fl=st.text_input("Compensation Notes",key=f"compn_fl_{sel_fl_id}")
                    if st.button("Record Compensation",key=f"comp_fl_{sel_fl_id}",use_container_width=True):
                        conn=get_conn()
                        new_remaining_days = max(int(fl_row['fine_days']) - comp_days_fl, 0)
                        new_status_fl = 'Compensated' if new_remaining_days==0 else 'Active'
                        conn.execute("""UPDATE fine_letters SET compensated_days=COALESCE(compensated_days,0)+?,
                            fine_days=?,compensation_notes=?,record_status=? WHERE id=?""",
                            (comp_days_fl,new_remaining_days,comp_notes_fl,new_status_fl,sel_fl_id))
                        conn.commit(); conn.close()
                        st.success(f"Recorded {comp_days_fl} compensated day(s). Remaining fine days: {new_remaining_days}."); st.rerun()
                with fmc3:
                    st.markdown(f"""<div style="background:#0D1526;border:1px solid rgba(255,255,255,0.07);border-radius:8px;padding:10px;margin-top:0">
                      <div style="font-size:9px;color:#6B7FA3;text-transform:uppercase">Current Status</div>
                      <div style="font-size:13px;color:#F0C96B;font-weight:600">{fl_row['record_status']}</div>
                      <div style="font-size:10px;color:#6B7FA3;margin-top:4px">Compensated so far: {fl_row['compensated_days']} day(s)</div>
                    </div>""",unsafe_allow_html=True)
            else:
                st.info("No active fines to manage.")

            st.markdown("<hr>",unsafe_allow_html=True)
            conn=get_conn()
            af=pd.read_sql_query("""SELECT fl.id,fl.emp_id,e.full_name,e.division,
                COALESCE(fl.month,'—') as month,fl.issue_date,
                COALESCE(fl.fine_type,'Disciplinary') as fine_type,
                COALESCE(fl.fine_reason,'') as fine_reason,
                fl.fine_days,fl.fine_amount,fl.letter_name,fl.applied_to_payroll,
                COALESCE(fl.record_status,'Active') as record_status
                FROM fine_letters fl LEFT JOIN employees e ON fl.emp_id=e.emp_id ORDER BY fl.created_at DESC LIMIT 1000""",conn); conn.close()
            if len(af)>0:
                buf4=io.BytesIO()
                with pd.ExcelWriter(buf4,engine="xlsxwriter") as w: af.to_excel(w,index=False,sheet_name="Fines")
                st.download_button("Export Fine Letters",buf4.getvalue(),file_name=f"Fines_{datetime.now().strftime('%Y%m%d')}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                prev_opts=[f"ID {r['id']} — {r['emp_id']} — {r['fine_type']} — {r['issue_date']}" for _,r in af.iterrows() if r['letter_name']]
                if prev_opts:
                    sel_fine=st.selectbox("Preview Fine Letter Document",prev_opts)
                    if sel_fine:
                        fid=int(sel_fine.split(" ")[1])
                        conn=get_conn(); cur=conn.cursor()
                        cur.execute("SELECT letter_name,letter_data FROM fine_letters WHERE id=?",(fid,))
                        frow=cur.fetchone(); conn.close()
                        if frow and frow[1]:
                            st.markdown(f'<div class="pb">{preview_html(frow[1],frow[0],"Fine Letter")}</div>',unsafe_allow_html=True)
                            fcol1,fcol2=st.columns(2)
                            with fcol1:
                                st.download_button("Download Fine Letter",data=bytes(frow[1]),file_name=frow[0],use_container_width=True)
                            with fcol2:
                                if st.session_state.role=="Manager" and st.button("Delete Fine Letter Document",use_container_width=True):
                                    conn=get_conn(); conn.execute("UPDATE fine_letters SET letter_name=NULL,letter_data=NULL WHERE id=?",(fid,)); conn.commit(); conn.close()
                                    st.success("Fine letter document deleted."); st.rerun()
                st.dataframe(af.drop(columns=["id","letter_name"]),use_container_width=True,hide_index=True)
        with lf3:
            st.markdown('<div class="fs">Record New Absence</div>',unsafe_allow_html=True)
            with st.form("abf"):
                ab1,ab2,ab3=st.columns(3)
                with ab1: ab_emp=st.selectbox("Employee",list(elo.keys()))
                with ab2: ab_date=st.date_input("Absent Date",value=date.today())
                with ab3: ab_ex=st.selectbox("Type",["Unexcused (Deducted)","Excused (Not Deducted)"])
                ab_reason=st.text_input("Reason",placeholder="Reason for absence...")
                if st.form_submit_button("Record Absent",use_container_width=True):
                    ab_eid_chosen=elo[ab_emp]
                    conn=get_conn(); cur=conn.cursor()
                    cur.execute("""SELECT COUNT(*) FROM absent_records WHERE emp_id=? AND absent_date=?
                        AND COALESCE(record_status,'Active')='Active'""",(ab_eid_chosen,str(ab_date)))
                    dup_count=cur.fetchone()[0]
                    if dup_count>0:
                        conn.close()
                        st.error(f"{ab_emp} already has an absence recorded for {ab_date}. The system blocks duplicate entries automatically — edit or cancel the existing record below instead.")
                    else:
                        conn.execute("INSERT INTO absent_records(emp_id,absent_date,reason,is_excused,record_status,created_at)VALUES(?,?,?,?,'Active',?)",
                            (ab_eid_chosen,str(ab_date),ab_reason,1 if "Excused" in ab_ex else 0,datetime.now().strftime("%Y-%m-%d")))
                        conn.commit(); conn.close(); st.success(f"Absent recorded: {ab_date}"); st.rerun()

            st.markdown("<hr>",unsafe_allow_html=True)
            st.markdown('<div class="fs">Manage Absences — Cancel or Compensate</div>',unsafe_allow_html=True)
            st.markdown('<div style="font-size:11px;color:#6B7FA3;margin-bottom:8px"><b>Cancel</b> removes an absence recorded by mistake. <b>Compensate</b> marks the absent day as worked off, so it stops counting against the employee pay.</div>',unsafe_allow_html=True)
            conn=get_conn()
            ab_manage=pd.read_sql_query("""SELECT ar.id,ar.emp_id,e.full_name,ar.absent_date,ar.reason,
                CASE WHEN ar.is_excused=1 THEN 'Excused' ELSE 'Unexcused' END as type,
                COALESCE(ar.record_status,'Active') as record_status, COALESCE(ar.is_compensated,0) as is_compensated
                FROM absent_records ar LEFT JOIN employees e ON ar.emp_id=e.emp_id
                WHERE COALESCE(ar.record_status,'Active') != 'Cancelled' ORDER BY ar.absent_date DESC LIMIT 500""",conn); conn.close()

            if len(ab_manage)>0:
                ab_opts={f"ID {r['id']} — {r['emp_id']} {r['full_name']} — {r['absent_date']} ({r['type']}) [{r['record_status']}{', Compensated' if r['is_compensated'] else ''}]":r['id'] for _,r in ab_manage.iterrows()}
                sel_ab_label=st.selectbox("Select Absence to Manage",list(ab_opts.keys()),key="sel_ab_manage")
                sel_ab_id=ab_opts[sel_ab_label]

                amc1,amc2=st.columns(2)
                with amc1:
                    cancel_reason_ab=st.text_input("Cancellation Reason",key=f"cr_ab_{sel_ab_id}")
                    if st.button("Cancel This Absence",key=f"cancel_ab_{sel_ab_id}",use_container_width=True):
                        if not cancel_reason_ab:
                            st.error("Please provide a cancellation reason.")
                        else:
                            conn=get_conn()
                            conn.execute("UPDATE absent_records SET record_status='Cancelled',cancelled_by=?,cancelled_at=?,cancel_reason=? WHERE id=?",
                                (st.session_state.uid,datetime.now().strftime("%Y-%m-%d %H:%M:%S"),cancel_reason_ab,sel_ab_id))
                            conn.commit(); conn.close()
                            st.warning("Absence cancelled — will not affect payroll."); st.rerun()
                with amc2:
                    comp_date_ab=st.date_input("Compensation Date (extra day worked)",value=date.today(),key=f"compdate_ab_{sel_ab_id}")
                    comp_notes_ab=st.text_input("Compensation Notes",key=f"compn_ab_{sel_ab_id}")
                    if st.button("Mark as Compensated",key=f"comp_ab_{sel_ab_id}",use_container_width=True):
                        conn=get_conn()
                        conn.execute("UPDATE absent_records SET is_compensated=1,compensation_date=?,compensation_notes=?,record_status='Compensated' WHERE id=?",
                            (str(comp_date_ab),comp_notes_ab,sel_ab_id))
                        conn.commit(); conn.close()
                        st.success("Absence marked as compensated — will not affect payroll deduction."); st.rerun()
            else:
                st.info("No active absences to manage.")

            st.markdown("<hr>",unsafe_allow_html=True)
            conn=get_conn()
            aab=pd.read_sql_query("""SELECT ar.emp_id,e.full_name,e.division,ar.absent_date,ar.reason,
                CASE WHEN ar.is_excused=1 THEN 'Excused' ELSE 'Unexcused' END as type,
                COALESCE(ar.record_status,'Active') as record_status
                FROM absent_records ar LEFT JOIN employees e ON ar.emp_id=e.emp_id ORDER BY ar.absent_date DESC LIMIT 1000""",conn); conn.close()
            if len(aab)>0:
                buf5=io.BytesIO()
                with pd.ExcelWriter(buf5,engine="xlsxwriter") as w: aab.to_excel(w,index=False,sheet_name="Absences")
                st.download_button("Export",buf5.getvalue(),file_name=f"Absences_{datetime.now().strftime('%Y%m%d')}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                st.dataframe(aab,use_container_width=True,hide_index=True)
            else: st.info("No absent records yet.")

        with lf4:
            if st.session_state.role not in ("Manager","Department Head"):
                st.info("Only Department Head or Manager can review and approve submissions here.")
            else:
                st.markdown('<div style="color:#6B7FA3;font-size:12px;margin-bottom:12px">Leave records and fine letters submitted by Supervisors wait here until reviewed. Approving makes them effective; rejecting sends them back.</div>',unsafe_allow_html=True)
                conn=get_conn()
                pending_leave=pd.read_sql_query("""SELECT lr.id,lr.emp_id,e.full_name,e.division,lr.leave_type,lr.start_date,lr.end_date,
                    lr.days_taken,lr.approved_by,lr.notes FROM leave_records lr
                    JOIN employees e ON lr.emp_id=e.emp_id
                    WHERE lr.status='Pending Dept Head Approval' ORDER BY lr.created_at ASC""",conn)
                pending_fines=pd.read_sql_query("""SELECT fl.id,fl.emp_id,e.full_name,e.division,fl.fine_type,fl.fine_reason,
                    fl.fine_days,fl.fine_amount FROM fine_letters fl
                    JOIN employees e ON fl.emp_id=e.emp_id
                    WHERE fl.applied_to_payroll='Pending Dept Head Approval' ORDER BY fl.created_at ASC""",conn)
                conn.close()

                st.markdown(f"""<div class="mg" style="grid-template-columns:repeat(2,1fr)">
                  <div class="mb mg-amber"><div class="ml ml-amber">Pending Leave</div><div class="mv">{len(pending_leave)}</div></div>
                  <div class="mb mg-red"><div class="ml ml-red">Pending Fines</div><div class="mv">{len(pending_fines)}</div></div>
                </div>""",unsafe_allow_html=True)

                st.markdown('<div class="fs">Leave Records Awaiting Approval</div>',unsafe_allow_html=True)
                if len(pending_leave)==0:
                    st.info("No leave records waiting for approval.")
                else:
                    for _,lv in pending_leave.iterrows():
                        with st.expander(f"{lv['emp_id']} — {lv['full_name']} — {lv['leave_type']} — {lv['days_taken']} days"):
                            st.markdown(f"""<div style="font-size:12px;color:#94A8C8">
                              Division: <b style="color:#F0C96B">{lv['division']}</b> &nbsp;|&nbsp;
                              Period: {lv['start_date']} to {lv['end_date']} &nbsp;|&nbsp;
                              Submitted by: {lv['approved_by']}
                            </div>
                            <div style="font-size:12px;color:#C8D8F0;margin-top:6px">Notes: {lv['notes'] or '—'}</div>""",unsafe_allow_html=True)
                            lvc1,lvc2=st.columns(2)
                            with lvc1:
                                if st.button("Approve",key=f"appr_lv_{lv['id']}",use_container_width=True):
                                    conn=get_conn()
                                    conn.execute("UPDATE leave_records SET status='Approved',approved_by=? WHERE id=?",
                                        (f"Dept Head: {st.session_state.uid}",lv['id']))
                                    conn.commit(); conn.close()
                                    st.success("Leave approved."); st.rerun()
                            with lvc2:
                                if st.button("Reject",key=f"rej_lv_{lv['id']}",use_container_width=True):
                                    conn=get_conn()
                                    conn.execute("UPDATE leave_records SET status='Rejected',approved_by=? WHERE id=?",
                                        (f"Dept Head: {st.session_state.uid}",lv['id']))
                                    conn.commit(); conn.close()
                                    st.warning("Leave rejected."); st.rerun()

                st.markdown('<div class="fs" style="margin-top:14px">Fine Letters Awaiting Approval</div>',unsafe_allow_html=True)
                if len(pending_fines)==0:
                    st.info("No fine letters waiting for approval.")
                else:
                    for _,fn in pending_fines.iterrows():
                        with st.expander(f"{fn['emp_id']} — {fn['full_name']} — {fn['fine_type']} — ETB {fn['fine_amount']:,.2f}"):
                            st.markdown(f"""<div style="font-size:12px;color:#94A8C8">Division: <b style="color:#F0C96B">{fn['division']}</b> &nbsp;|&nbsp; Days: {fn['fine_days']}</div>
                            <div style="font-size:12px;color:#C8D8F0;margin-top:6px">Reason: {fn['fine_reason'] or '—'}</div>""",unsafe_allow_html=True)
                            fnc1,fnc2=st.columns(2)
                            with fnc1:
                                if st.button("Approve",key=f"appr_fn_{fn['id']}",use_container_width=True):
                                    conn=get_conn()
                                    conn.execute("UPDATE fine_letters SET applied_to_payroll='No' WHERE id=?",(fn['id'],))
                                    conn.commit(); conn.close()
                                    st.success("Fine approved — will apply to next payroll."); st.rerun()
                            with fnc2:
                                if st.button("Reject",key=f"rej_fn_{fn['id']}",use_container_width=True):
                                    conn=get_conn()
                                    conn.execute("UPDATE fine_letters SET applied_to_payroll='Rejected' WHERE id=?",(fn['id'],))
                                    conn.commit(); conn.close()
                                    st.warning("Fine rejected."); st.rerun()

    # ════════════════════════════════════════════════════════
    # PUBLIC HOLIDAYS
    # ════════════════════════════════════════════════════════
    elif V=="Public Holidays":
        st.markdown('<div class="ey">Ethiopian Labour Law</div>',unsafe_allow_html=True)
        st.markdown('<div class="tl">Ethiopian Public Holidays</div>',unsafe_allow_html=True)
        yr=st.selectbox("Year",[2024,2025,2026,2027],index=2)
        holidays=get_holidays(yr)
        st.markdown(f"""<div class="card card-gold">
          <div style="font-size:12px;color:#C8D8F0;line-height:1.8;margin-bottom:12px">All holidays below are <b style="color:#F0C96B">paid public holidays</b> per Ethiopian Labour Proclamation No. 1156/2019. Automatically counted as paid days in the payroll system.</div>
          <div>{''.join([f'<span class="hch"> {d.strftime("%b %d")} — {n}</span>' for d,n in sorted(holidays.items())])}</div>
        </div>""",unsafe_allow_html=True)
        hdf=pd.DataFrame([(d.strftime("%Y-%m-%d"),d.strftime("%A"),n,"Paid","Auto-included in payroll") for d,n in sorted(holidays.items())],
            columns=["Date","Day","Holiday Name","Type","Payroll"])
        st.dataframe(hdf,use_container_width=True,hide_index=True)
        st.markdown('<div class="fs" style="margin-top:14px">Ethiopian Leave Policy (Proc. 1156/2019)</div>',unsafe_allow_html=True)

        if st.session_state.role=="Manager":
            st.markdown('<div style="font-size:11px;color:#6B7FA3;margin-bottom:10px">Manager can adjust these values. Changes apply immediately to payroll calculations.</div>',unsafe_allow_html=True)
            with st.form("policy_edit_form"):
                pc1,pc2=st.columns(2)
                with pc1:
                    annual_days=st.number_input("Annual Leave (days/year)",min_value=0,max_value=60,value=int(get_setting("policy_annual_leave_days","20")),step=1)
                    sick_full=st.number_input("Sick Leave — Full Pay (months)",min_value=0,max_value=12,value=int(get_setting("policy_sick_leave_full_months","1")),step=1)
                    sick_half=st.number_input("Sick Leave — Half Pay (months)",min_value=0,max_value=12,value=int(get_setting("policy_sick_leave_half_months","2")),step=1)
                    maternity_days=st.number_input("Maternity Leave (days)",min_value=0,max_value=180,value=int(get_setting("policy_maternity_leave_days","90")),step=1)
                    paternity_days=st.number_input("Paternity Leave (days)",min_value=0,max_value=30,value=int(get_setting("policy_paternity_leave_days","3")),step=1)
                    mourning_days=st.number_input("Mourning Leave (days)",min_value=0,max_value=30,value=int(get_setting("policy_mourning_leave_days","3")),step=1)
                with pc2:
                    working_days=st.number_input("Working Days per Month (for daily rate)",min_value=20,max_value=31,value=int(get_setting("policy_working_days_per_month","26")),step=1)
                    ot_weekday=st.number_input("Overtime Multiplier — Weekday",min_value=1.0,max_value=5.0,value=float(get_setting("policy_overtime_weekday","1.25")),step=0.05,format="%.2f")
                    ot_weekend=st.number_input("Overtime Multiplier — Weekend",min_value=1.0,max_value=5.0,value=float(get_setting("policy_overtime_weekend","1.5")),step=0.05,format="%.2f")
                    ot_holiday=st.number_input("Overtime Multiplier — Holiday",min_value=1.0,max_value=5.0,value=float(get_setting("policy_overtime_holiday","2.0")),step=0.05,format="%.2f")
                    holiday_status=st.selectbox("Holiday Payment Status",["Paid","Unpaid"],index=0 if get_setting("policy_holiday_payment_status","Paid")=="Paid" else 1)
                    dayoff_status=st.selectbox("Weekly Day-Off Payment Status",["Paid","Unpaid"],index=0 if get_setting("policy_dayoff_payment_status","Paid")=="Paid" else 1)
                if st.form_submit_button("Save Policy Changes",use_container_width=True):
                    for k,v in {
                        "policy_annual_leave_days":annual_days,"policy_sick_leave_full_months":sick_full,
                        "policy_sick_leave_half_months":sick_half,"policy_maternity_leave_days":maternity_days,
                        "policy_paternity_leave_days":paternity_days,"policy_mourning_leave_days":mourning_days,
                        "policy_working_days_per_month":working_days,"policy_overtime_weekday":ot_weekday,
                        "policy_overtime_weekend":ot_weekend,"policy_overtime_holiday":ot_holiday,
                        "policy_holiday_payment_status":holiday_status,"policy_dayoff_payment_status":dayoff_status,
                    }.items():
                        set_setting(k,str(v),st.session_state.uid)
                    st.success("Policy updated. New values apply to all future payroll calculations.")
                    st.rerun()
        else:
            pol_annual=get_setting("policy_annual_leave_days","20")
            pol_sick_f=get_setting("policy_sick_leave_full_months","1")
            pol_sick_h=get_setting("policy_sick_leave_half_months","2")
            pol_mat=get_setting("policy_maternity_leave_days","90")
            pol_pat=get_setting("policy_paternity_leave_days","3")
            pol_mourn=get_setting("policy_mourning_leave_days","3")
            pol_wd=get_setting("policy_working_days_per_month","26")
            pol_otw=get_setting("policy_overtime_weekday","1.25")
            pol_ote=get_setting("policy_overtime_weekend","1.5")
            pol_oth=get_setting("policy_overtime_holiday","2.0")
            st.markdown(f"""<div class="card">
              <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;font-size:12px">
                <div><div style="color:#F0C96B;font-weight:600;margin-bottom:2px">Annual Leave</div><div style="color:#94A8C8">{pol_annual} working days/year. Paid.</div></div>
                <div><div style="color:#F0C96B;font-weight:600;margin-bottom:2px">Sick Leave</div><div style="color:#94A8C8">{pol_sick_f} month(s) full pay, then {pol_sick_h} month(s) at 50%.</div></div>
                <div><div style="color:#F0C96B;font-weight:600;margin-bottom:2px">Maternity Leave</div><div style="color:#94A8C8">{pol_mat} calendar days. Fully paid.</div></div>
                <div><div style="color:#F0C96B;font-weight:600;margin-bottom:2px">Paternity Leave</div><div style="color:#94A8C8">{pol_pat} working days. Paid.</div></div>
                <div><div style="color:#F0C96B;font-weight:600;margin-bottom:2px">Mourning Leave</div><div style="color:#94A8C8">{pol_mourn} days immediate family. Paid.</div></div>
                <div><div style="color:#F0C96B;font-weight:600;margin-bottom:2px">Weekly Day-Off</div><div style="color:#94A8C8">1 day/week, employee-specific. Auto-calculated.</div></div>
                <div><div style="color:#F0C96B;font-weight:600;margin-bottom:2px">Daily Rate</div><div style="color:#94A8C8">Basic salary ÷ {pol_wd} working days.</div></div>
                <div><div style="color:#F0C96B;font-weight:600;margin-bottom:2px">Overtime</div><div style="color:#94A8C8">{pol_otw}× weekdays, {pol_ote}× weekends, {pol_oth}× holidays.</div></div>
              </div></div>""",unsafe_allow_html=True)

        hbuf2=io.BytesIO()
        with pd.ExcelWriter(hbuf2,engine="xlsxwriter") as w: hdf.to_excel(w,index=False,sheet_name="Holidays")
        st.download_button("Export Holiday Calendar",hbuf2.getvalue(),file_name=f"Holidays_{yr}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ════════════════════════════════════════════════════════
    # COST CENTERS — manually created per division
    # ════════════════════════════════════════════════════════
    elif V=="Cost Centers":
        st.markdown('<div class="ey">Financial Structure</div>',unsafe_allow_html=True)
        st.markdown('<div class="tl">Division Cost Centers</div>',unsafe_allow_html=True)
        conn=get_conn()
        cc_all=pd.read_sql_query("SELECT * FROM cost_centers ORDER BY division,code",conn); conn.close()
        st.markdown(f'<div class="card"><span style="color:#D4A847;font-weight:600"> Total Cost Centers:</span> <b style="color:#E8EEF7">{len(cc_all)}</b></div>',unsafe_allow_html=True)
        if len(cc_all)>0:
            for div in cc_all['division'].unique():
                div_ccs=cc_all[cc_all['division']==div]
                st.markdown(f'<div class="fs">{div} Division</div>',unsafe_allow_html=True)
                for _,ccr in div_ccs.iterrows():
                    status_label="Active" if ccr['is_active']==1 else "Inactive"
                    st.markdown(f"""<div class="card" style="padding:12px 16px;margin-bottom:8px">
                      <div style="display:flex;justify-content:space-between;align-items:center">
                        <div>
                          <span style="font-family:'Cinzel',serif;color:#F0C96B;font-weight:700;font-size:14px">{ccr['code']}</span>
                          <span style="color:#94A8C8;font-size:12px;margin-left:8px">{ccr['name']}</span>
                        </div>
                        <div style="text-align:right">
                          <div style="color:#10B981;font-size:13px;font-weight:600">ETB {ccr['budget']:,.2f}</div>
                          <div style="font-size:10px;color:#6B7FA3">{status_label}</div>
                        </div>
                      </div>
                    </div>""",unsafe_allow_html=True)
        st.markdown("<hr>",unsafe_allow_html=True)
        if st.session_state.role=="Manager":
            t1,t2=st.tabs(["Create Cost Center","Manage Cost Centers"])
            with t1:
                with st.form("cc_form"):
                    c1,c2=st.columns(2)
                    with c1: cc_code=st.text_input("Cost Center Code *",placeholder="e.g. CC-CAT-03")
                    with c2: cc_division=st.selectbox("Division *",get_division_list())
                    cc_name=st.text_input("Cost Center Name *",placeholder="e.g. Catering Operations Phase 2")
                    c3,c4=st.columns(2)
                    with c3: cc_budget=st.number_input("Annual Budget (ETB)",min_value=0.0,step=10000.0)
                    with c4: cc_active=st.selectbox("Status",["Active","Inactive"])
                    cc_desc=st.text_area("Description",placeholder="Purpose of this cost center...")
                    if st.form_submit_button("Create Cost Center",use_container_width=True):
                        if not(cc_code and cc_name): st.error("Code and Name required.")
                        else:
                            conn=get_conn()
                            try:
                                conn.execute("INSERT INTO cost_centers(code,name,division,budget,description,is_active,created_by,created_at)VALUES(?,?,?,?,?,?,?,?)",
                                    (cc_code,cc_name,cc_division,cc_budget,cc_desc,1 if cc_active=="Active" else 0,st.session_state.uid,datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                                conn.commit()
                                get_cost_centers.clear()
                                st.success(f"Cost center {cc_code} created for {cc_division} division.")
                            except sqlite3.IntegrityError:
                                st.error(f"Code '{cc_code}' already exists.")
                            finally: conn.close()
            with t2:
                if len(cc_all)>0:
                    cc_opts={f"{r['code']} — {r['name']} ({r['division']})":r['code'] for _,r in cc_all.iterrows()}
                    sel_cc=st.selectbox("Select Cost Center",list(cc_opts.keys()))
                    sel_code=cc_opts[sel_cc]
                    mc1,mc2,mc3=st.columns(3)
                    with mc1:
                        if st.button("Activate",use_container_width=True):
                            conn=get_conn(); conn.execute("UPDATE cost_centers SET is_active=1 WHERE code=?",(sel_code,)); conn.commit(); conn.close()
                            get_cost_centers.clear(); st.success("Activated."); st.rerun()
                    with mc2:
                        if st.button("Deactivate",use_container_width=True):
                            conn=get_conn(); conn.execute("UPDATE cost_centers SET is_active=0 WHERE code=?",(sel_code,)); conn.commit(); conn.close()
                            get_cost_centers.clear(); st.warning("Deactivated."); st.rerun()
                    with mc3:
                        if st.button("Delete",use_container_width=True):
                            conn=get_conn(); cur=conn.cursor()
                            cur.execute("SELECT * FROM cost_centers WHERE code=?",(sel_code,))
                            cc_cols=[d[0] for d in cur.description]; cc_row=cur.fetchone()
                            cc_dict=dict(zip(cc_cols,cc_row)) if cc_row else {}
                            conn.execute("DELETE FROM cost_centers WHERE code=?",(sel_code,)); conn.commit(); conn.close()
                            soft_delete("Cost Center", sel_code, f"{sel_code} — {cc_dict.get('name','')}", cc_dict, st.session_state.uid)
                            get_cost_centers.clear(); st.error("Moved to Recycle Bin."); st.rerun()
                else:
                    st.info("No cost centers yet. Create one in the first tab.")
        cc_buf=io.BytesIO()
        with pd.ExcelWriter(cc_buf,engine="xlsxwriter") as w: cc_all.to_excel(w,index=False,sheet_name="CostCenters")
        st.download_button("Export Cost Center List",cc_buf.getvalue(),file_name=f"CostCenters_{datetime.now().strftime('%Y%m%d')}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ════════════════════════════════════════════════════════
    # RECYCLE BIN — restore accidentally deleted records
    # ════════════════════════════════════════════════════════
    elif V=="Recycle Bin":
        st.markdown('<div class="ey">Data Recovery</div>',unsafe_allow_html=True)
        st.markdown('<div class="tl">Recycle Bin</div>',unsafe_allow_html=True)
        st.markdown("""<div class="card card-gold">
          <div style="font-size:12px;color:#C8D8F0;line-height:1.7">
            Deleted employees, cost centers, and user accounts are kept here before being
            permanently erased. A Manager can restore any item or permanently purge it.
          </div></div>""",unsafe_allow_html=True)

        conn=get_conn()
        bin_items=pd.read_sql_query("SELECT * FROM recycle_bin WHERE restored=0 ORDER BY deleted_at DESC LIMIT 300",conn)
        conn.close()

        type_counts = bin_items['record_type'].value_counts().to_dict() if len(bin_items)>0 else {}
        st.markdown(f"""<div class="mg" style="grid-template-columns:repeat(3,1fr)">
          <div class="mb mg-amber"><div class="ml ml-amber">Deleted Employees</div><div class="mv">{type_counts.get("Employee",0)}</div></div>
          <div class="mb mg-purple"><div class="ml ml-purple">Deleted Cost Centers</div><div class="mv">{type_counts.get("Cost Center",0)}</div></div>
          <div class="mb mg-cyan"><div class="ml ml-cyan">Deleted Users</div><div class="mv">{type_counts.get("User Account",0)}</div></div>
        </div>""",unsafe_allow_html=True)

        if len(bin_items)==0:
            st.info("Recycle Bin is empty.")
        else:
            bin_filter=st.selectbox("Filter by Type",["All"]+list(bin_items['record_type'].unique()))
            display_items = bin_items if bin_filter=="All" else bin_items[bin_items['record_type']==bin_filter]

            for _,item in display_items.iterrows():
                with st.expander(f"{item['record_type']} — {item['record_label']} — deleted {item['deleted_at']} by {item['deleted_by']}"):
                    try:
                        data_preview = json.loads(item['record_data']) if item['record_data'] else {}
                        safe_preview = {k:v for k,v in data_preview.items() if not str(k).endswith('_data')}
                        st.json(safe_preview)
                    except: st.write("No preview available.")

                    rc1,rc2=st.columns(2)
                    with rc1:
                        if st.button("Restore",key=f"restore_{item['id']}",use_container_width=True):
                            try:
                                data_dict = json.loads(item['record_data']) if item['record_data'] else {}
                            except: data_dict = {}
                            conn=get_conn()
                            restored_ok=False
                            if item['record_type']=="Employee" and data_dict:
                                cols_list = [k for k in data_dict.keys()]
                                placeholders = ",".join(["?"]*len(cols_list))
                                col_names = ",".join(cols_list)
                                vals = [data_dict[c] for c in cols_list]
                                try:
                                    conn.execute(f"INSERT INTO employees({col_names}) VALUES({placeholders})", vals)
                                    restored_ok=True
                                except Exception as ex:
                                    st.error(f"Could not restore — a record with this ID may already exist. ({ex})")
                            elif item['record_type']=="Cost Center" and data_dict:
                                try:
                                    conn.execute("""INSERT INTO cost_centers(code,name,division,budget,description,is_active,created_by,created_at)
                                        VALUES(?,?,?,?,?,?,?,?)""",
                                        (data_dict.get('code'),data_dict.get('name'),data_dict.get('division'),
                                         data_dict.get('budget',0),data_dict.get('description'),1,
                                         data_dict.get('created_by'),data_dict.get('created_at')))
                                    restored_ok=True
                                except Exception as ex:
                                    st.error(f"Could not restore — code may already exist. ({ex})")
                            elif item['record_type']=="User Account" and data_dict:
                                st.warning("User accounts cannot be restored with their original password for security. Please create a new account in Administration instead.")
                            if restored_ok:
                                conn.execute("UPDATE recycle_bin SET restored=1 WHERE id=?",(item['id'],))
                                conn.commit()
                                st.cache_data.clear(); get_employee.clear(); get_cost_centers.clear()
                                st.success("Restored successfully.")
                            conn.close()
                            st.rerun()
                    with rc2:
                        if st.button("Permanently Purge",key=f"purge_{item['id']}",use_container_width=True):
                            conn=get_conn()
                            conn.execute("DELETE FROM recycle_bin WHERE id=?",(item['id'],))
                            conn.commit(); conn.close()
                            st.warning("Permanently purged — this cannot be undone.")
                            st.rerun()

            st.markdown("<hr>",unsafe_allow_html=True)
            if st.button("Empty Entire Recycle Bin (Permanent)",use_container_width=True):
                conn=get_conn()
                conn.execute("DELETE FROM recycle_bin WHERE restored=0")
                conn.commit(); conn.close()
                st.warning("Recycle Bin emptied permanently.")
                st.rerun()

    # ════════════════════════════════════════════════════════
    # ADMINISTRATION (Manager only)
    # ════════════════════════════════════════════════════════
    elif V=="Administration":
        if st.session_state.role != "Manager":
            st.markdown("""<div style="background:rgba(239,68,68,0.08);border:1px solid rgba(239,68,68,0.3);
              border-radius:14px;padding:48px;text-align:center">
              <div style="font-size:48px;margin-bottom:14px"></div>
              <div style="font-family:'Cinzel',serif;font-size:20px;color:#EF4444;margin-bottom:8px">Access Denied</div>
              <div style="color:#6B7FA3;font-size:13px">Administration is restricted to Manager role only.</div>
            </div>""",unsafe_allow_html=True)
            st.stop()

        st.markdown("""<style>
        .admin-hero{background:linear-gradient(135deg,#0D1526,#0A1020,#0D1830);
          border:1px solid rgba(212,168,71,0.3);border-radius:16px;padding:24px 28px;
          margin-bottom:20px;position:relative;overflow:hidden}
        .admin-hero::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;
          background:linear-gradient(90deg,#D4A847,#F0C96B,#D4A847)}
        .user-card{background:#0D1526;border:1px solid rgba(255,255,255,0.07);border-radius:12px;
          padding:16px 18px;margin-bottom:10px;display:flex;align-items:center;gap:14px;
          transition:border-color .2s;position:relative;overflow:hidden}
        .user-card:hover{border-color:rgba(212,168,71,0.25)}
        .user-card::before{content:'';position:absolute;left:0;top:0;bottom:0;width:3px}
        .uc-manager::before{background:#D4A847}.uc-officer::before{background:#38BDF8}
        .uc-hr::before{background:#10B981}.uc-payroll::before{background:#A855F7}.uc-dept::before{background:#F59E0B}
        .user-avatar{width:40px;height:40px;border-radius:50%;display:flex;align-items:center;
          justify-content:center;font-family:'Cinzel',serif;font-size:14px;font-weight:700;flex-shrink:0}
        .ua-manager{background:rgba(212,168,71,0.15);color:#F0C96B;border:1px solid rgba(212,168,71,0.3)}
        .ua-officer{background:rgba(56,189,248,0.15);color:#38BDF8;border:1px solid rgba(56,189,248,0.3)}
        .ua-hr{background:rgba(16,185,129,0.15);color:#10B981;border:1px solid rgba(16,185,129,0.3)}
        .ua-payroll{background:rgba(168,85,247,0.15);color:#A855F7;border:1px solid rgba(168,85,247,0.3)}
        .ua-dept{background:rgba(245,158,11,0.15);color:#F59E0B;border:1px solid rgba(245,158,11,0.3)}
        .user-info{flex:1;min-width:0}
        .user-name{font-size:13px;font-weight:600;color:#E8EEF7;margin-bottom:2px}
        .user-uid{font-size:11px;color:#6B7FA3;font-family:monospace}
        .role-badge{display:inline-block;padding:2px 10px;border-radius:20px;font-size:10px;font-weight:600;letter-spacing:.04em}
        .rb-manager{background:rgba(212,168,71,0.15);color:#F0C96B;border:1px solid rgba(212,168,71,0.3)}
        .rb-officer{background:rgba(56,189,248,0.1);color:#7DD3FC;border:1px solid rgba(56,189,248,0.2)}
        .rb-hr{background:rgba(16,185,129,0.15);color:#34D399;border:1px solid rgba(16,185,129,0.3)}
        .rb-payroll{background:rgba(168,85,247,0.1);color:#C4B5FD;border:1px solid rgba(168,85,247,0.2)}
        .rb-dept{background:rgba(245,158,11,0.1);color:#FCD34D;border:1px solid rgba(245,158,11,0.2)}
        .active-dot{display:inline-block;width:7px;height:7px;border-radius:50%;margin-right:5px}
        .dot-active{background:#10B981;box-shadow:0 0 6px rgba(16,185,129,0.5)}.dot-inactive{background:#EF4444}
        .cred-box{background:linear-gradient(135deg,#0A1A10,#0D1526);border:1px solid rgba(16,185,129,0.3);
          border-radius:10px;padding:16px 20px;margin-top:14px}
        .cred-row{display:flex;align-items:center;justify-content:space-between;padding:6px 0;
          border-bottom:1px solid rgba(255,255,255,0.05);font-size:13px}
        .cred-row:last-child{border-bottom:none}
        .cred-label{color:#6B7FA3;font-size:10px;text-transform:uppercase;letter-spacing:.08em}
        .cred-val{color:#F0C96B;font-family:monospace;font-size:14px;font-weight:600}
        .perm-matrix{display:grid;grid-template-columns:repeat(5,1fr);gap:6px;margin-top:10px}
        .pm-cell{background:#0A1020;border:1px solid rgba(255,255,255,0.06);border-radius:8px;padding:10px 8px;text-align:center}
        .pm-title{font-size:9px;color:#6B7FA3;text-transform:uppercase;letter-spacing:.07em;margin-bottom:6px}
        .pm-perms{font-size:10px}
        .perm-chip{display:inline-flex;align-items:center;gap:4px;font-size:10px;padding:2px 8px;border-radius:8px;background:rgba(255,255,255,0.05);color:#6B7FA3;margin:1px}
        .perm-chip.on{background:rgba(16,185,129,0.1);color:#34D399}
        .perm-chip.off{background:rgba(239,68,68,0.08);color:#FCA5A5}
        </style>""",unsafe_allow_html=True)

        conn=get_conn()
        total_users=pd.read_sql_query("SELECT COUNT(*) as c FROM system_users",conn).iloc[0]['c']
        active_users=pd.read_sql_query("SELECT COUNT(*) as c FROM system_users WHERE is_active=1",conn).iloc[0]['c']
        conn.close()

        st.markdown(f"""<div class="admin-hero">
          <div style="display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:12px">
            <div>
              <div style="font-size:9px;color:#D4A847;letter-spacing:.14em;text-transform:uppercase;margin-bottom:4px">System Administration</div>
              <div style="font-family:'Cinzel',serif;font-size:22px;font-weight:700;color:#F0C96B">Administration Panel</div>
              <div style="font-size:12px;color:#6B7FA3;margin-top:4px">Manage users, roles, permissions and system-wide settings</div>
            </div>
            <div style="display:flex;gap:10px;flex-wrap:wrap">
              <div style="background:rgba(212,168,71,0.08);border:1px solid rgba(212,168,71,0.2);border-radius:10px;padding:10px 16px;text-align:center">
                <div style="font-family:'Cinzel',serif;font-size:20px;font-weight:700;color:#F0C96B">{total_users}</div>
                <div style="font-size:9px;color:#6B7FA3;text-transform:uppercase;letter-spacing:.08em">Total Users</div>
              </div>
              <div style="background:rgba(16,185,129,0.08);border:1px solid rgba(16,185,129,0.2);border-radius:10px;padding:10px 16px;text-align:center">
                <div style="font-family:'Cinzel',serif;font-size:20px;font-weight:700;color:#10B981">{active_users}</div>
                <div style="font-size:9px;color:#6B7FA3;text-transform:uppercase;letter-spacing:.08em">Active</div>
              </div>
            </div>
          </div>
        </div>""",unsafe_allow_html=True)

        with st.expander("Permission Matrix — Role Reference"):
            st.markdown("""<div class="perm-matrix" style="grid-template-columns:repeat(4,1fr)">
              <div class="pm-cell"><div class="pm-title" style="color:#D4A847">Manager</div><div class="pm-perms">
                <div class="perm-chip on">All Modules</div><div class="perm-chip on">Admin</div>
                <div class="perm-chip on">Edit/Delete</div></div></div>
              <div class="pm-cell"><div class="pm-title" style="color:#06B6D4">Supervisor</div><div class="pm-perms">
                <div class="perm-chip on">One Division</div><div class="perm-chip on">Absence/Leave</div>
                <div class="perm-chip on">Submit Sheet</div></div></div>
              <div class="pm-cell"><div class="pm-title" style="color:#EAB308">Payroll Section</div><div class="pm-perms">
                <div class="perm-chip on">Review Submissions</div><div class="perm-chip on">Approve/Reject</div>
                <div class="perm-chip on">Process Payroll</div></div></div>
              <div class="pm-cell"><div class="pm-title" style="color:#10B981">HR Staff</div><div class="pm-perms">
                <div class="perm-chip on">Edit Profiles</div><div class="perm-chip on">Documents</div>
                <div class="perm-chip off">Payroll</div></div></div>
              <div class="pm-cell"><div class="pm-title" style="color:#A855F7">Payroll Officer</div><div class="pm-perms">
                <div class="perm-chip on">Payroll</div><div class="perm-chip on">Fines/Leave</div>
                <div class="perm-chip off">Edit Profile</div></div></div>
              <div class="pm-cell"><div class="pm-title" style="color:#F59E0B">Dept Head</div><div class="pm-perms">
                <div class="perm-chip on">View All</div><div class="perm-chip off">Payroll</div>
                <div class="perm-chip off">Edit</div></div></div>
              <div class="pm-cell"><div class="pm-title" style="color:#38BDF8">Data Officer</div><div class="pm-perms">
                <div class="perm-chip on">View/Search</div><div class="perm-chip off">Edit</div>
                <div class="perm-chip off">Admin</div></div></div>
            </div>""",unsafe_allow_html=True)

        st.markdown("<hr>",unsafe_allow_html=True)
        at1,at2,at3=st.tabs(["All Users","Create User","Edit & Reset Password"])

        def role_cls(role):
            return {"Manager":"manager","HR Staff":"hr","Payroll Officer":"payroll",
                    "Department Head":"dept","Data Officer":"officer",
                    "Supervisor":"hr","Payroll Section":"payroll"}.get(role,"officer")
        def role_init(name):
            return (name or "?")[0].upper()

        with at1:
            conn=get_conn()
            all_users_df=pd.read_sql_query(
                "SELECT id,username,full_name,role,permissions,is_active,email,created_by,created_at,last_login,assigned_division FROM system_users ORDER BY created_at DESC",conn)
            conn.close()
            for _,u in all_users_df.iterrows():
                rc=role_cls(u['role']); is_active=int(u['is_active'])==1
                _ll=u['last_login']; last="Never logged in" if (_ll is None or (isinstance(_ll,float))) else str(_ll); perm=u['permissions'] or "view_only"
                div_badge = f'<span style="font-size:10px;color:#38BDF8">Division: {u["assigned_division"]}</span>' if u.get("assigned_division") else ''
                st.markdown(f"""<div class="user-card uc-{rc}">
                  <div class="user-avatar ua-{rc}">{role_init(u["full_name"])}</div>
                  <div class="user-info">
                    <div class="user-name">{u["full_name"] or u["username"]}</div>
                    <div class="user-uid">@{u["username"]} &nbsp;·&nbsp; {u["email"] or "No email"}</div>
                    <div style="margin-top:5px;display:flex;align-items:center;gap:8px;flex-wrap:wrap">
                      <span class="role-badge rb-{rc}">{u["role"]}</span>
                      <span style="font-size:10px;color:#6B7FA3"><span class="active-dot {"dot-active" if is_active else "dot-inactive"}"></span>{"Active" if is_active else "Disabled"}</span>
                      <span style="font-size:10px;color:#6B7FA3"> {last[:16] if last!="Never logged in" else "Never"}</span>
                      {div_badge}
                    </div>
                  </div>
                  <div style="text-align:right;flex-shrink:0">
                    <div style="font-size:9px;color:#6B7FA3;text-transform:uppercase;letter-spacing:.07em">Permissions</div>
                    <div style="font-size:11px;color:{"#10B981" if perm=="full" else "#94A8C8"};font-weight:500">{perm}</div>
                  </div>
                </div>""",unsafe_allow_html=True)
            st.markdown("<hr>",unsafe_allow_html=True)
            st.markdown('<div class="fs">Enable / Disable Account</div>',unsafe_allow_html=True)
            conn=get_conn()
            toggle_df=pd.read_sql_query("SELECT username,full_name,role,is_active FROM system_users WHERE username != ? ORDER BY username",conn,params=(st.session_state.uid,))
            conn.close()
            if len(toggle_df)>0:
                tog_opts={f"@{r['username']} — {r['full_name'] or ''} ({r['role']}) [{' Active' if r['is_active'] else ' Disabled'}]":r['username'] for _,r in toggle_df.iterrows()}
                sel_tog=st.selectbox("Select Account",list(tog_opts.keys()))
                tog_uid=tog_opts[sel_tog]
                tc1,tc2,tc3=st.columns(3)
                with tc1:
                    if st.button("Enable",use_container_width=True):
                        conn=get_conn(); conn.execute("UPDATE system_users SET is_active=1 WHERE username=?",(tog_uid,)); conn.commit(); conn.close()
                        st.success(f"@{tog_uid} enabled."); st.rerun()
                with tc2:
                    if st.button("Disable",use_container_width=True):
                        conn=get_conn(); conn.execute("UPDATE system_users SET is_active=0 WHERE username=?",(tog_uid,)); conn.commit(); conn.close()
                        st.warning(f"@{tog_uid} disabled."); st.rerun()
                with tc3:
                    if st.button("Delete",use_container_width=True):
                        conn=get_conn(); cur=conn.cursor()
                        cur.execute("SELECT * FROM system_users WHERE username=?",(tog_uid,))
                        u_cols=[d[0] for d in cur.description]; u_row=cur.fetchone()
                        u_dict=dict(zip(u_cols,u_row)) if u_row else {}
                        u_dict.pop("password",None)
                        conn.execute("DELETE FROM system_users WHERE username=?",(tog_uid,)); conn.commit(); conn.close()
                        soft_delete("User Account", tog_uid, f"@{tog_uid} — {u_dict.get('full_name','')}", u_dict, st.session_state.uid)
                        st.error(f"@{tog_uid} moved to Recycle Bin."); st.rerun()
            st.markdown("<hr>",unsafe_allow_html=True)
            conn=get_conn()
            exp_u=pd.read_sql_query("SELECT id,username,full_name,role,permissions,is_active,email,created_by,created_at,last_login FROM system_users",conn); conn.close()
            buf_u=io.BytesIO()
            with pd.ExcelWriter(buf_u,engine="xlsxwriter") as w: exp_u.to_excel(w,index=False,sheet_name="Users")
            st.download_button("Export User Registry",buf_u.getvalue(),file_name=f"Users_{datetime.now().strftime('%Y%m%d')}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)

        with at2:
            st.markdown("""<div class="card" style="border-color:rgba(212,168,71,0.2)">
              <div class="fs">New User Registration</div>
              <div style="font-size:12px;color:#6B7FA3;line-height:1.7">User can log in immediately after creation.</div>
            </div>""",unsafe_allow_html=True)

            with st.form("create_user_form",clear_on_submit=True):
                cu1,cu2,cu3=st.columns(3)
                with cu1: new_uname=st.text_input("Username *",placeholder="e.g. ygs_hr_staff")
                with cu2: new_pw=st.text_input("Password *",placeholder="Min. 6 characters")
                with cu3: new_fullname=st.text_input("Full Name *",placeholder="Employee full name")
                cu4,cu5,cu6=st.columns(3)
                with cu4: new_email=st.text_input("Email",placeholder="user@company.com")
                with cu5: new_role=st.selectbox("Role *",["Supervisor","Payroll Section","Data Officer","HR Staff","Payroll Officer","Department Head","Manager"])
                with cu6:
                    if new_role=="Supervisor":
                        dept_assign=st.selectbox("Assigned Division *",get_division_list())
                    else:
                        dept_assign=st.selectbox("Division Scope",["All Divisions"]+get_division_list())
                PERM_DETAILS={"Manager":("full","Full access — all modules, admin, edit, delete, payroll","#D4A847"),
                    "Supervisor":("division_control","Controls attendance, leave and absence for one division. Submits monthly sheets for approval.","#06B6D4"),
                    "Payroll Section":("payroll_approve","Reviews and approves monthly submissions from supervisors. Releases salary payments.","#EAB308"),
                    "HR Staff":("hr_edit","Employee profiles, documents, leave records.","#10B981"),
                    "Payroll Officer":("payroll_edit","Payroll, fines, leave, absences.","#A855F7"),
                    "Department Head":("dept_view","View records in assigned division. Read-only.","#F59E0B"),
                    "Data Officer":("view_only","Browse and search only.","#38BDF8")}
                pd_key,pd_desc,pd_col=PERM_DETAILS.get(new_role,("view_only","Read-only access","#38BDF8"))
                st.markdown(f"""<div style="background:#0A1020;border:1px solid rgba(255,255,255,0.07);border-radius:10px;padding:14px 16px">
                  <div style="font-size:9px;color:#6B7FA3;text-transform:uppercase;margin-bottom:4px">Permission Level (Role Default)</div>
                  <span style="font-size:11px;font-weight:600;color:{pd_col}">{pd_key}</span> — <span style="font-size:12px;color:#94A8C8">{pd_desc}</span>
                </div>""",unsafe_allow_html=True)

                st.markdown('<div class="fs" style="margin-top:14px">Navigation Access — Set Authority Per Module</div>',unsafe_allow_html=True)
                st.markdown('<div style="font-size:11px;color:#6B7FA3;margin-bottom:8px">For every module, choose the authority this user has: Role Default (use the permission level above), View (read-only), Edit (change data, no delete), Both (view and edit), or Full Control (view, edit and delete).</div>',unsafe_allow_html=True)
                custom_nav_selection = {}
                PERM_LEVEL_OPTIONS = ["Role Default","View","Edit","Both","Full Control"]
                for nv in ALL_NAV_VIEWS:
                    nvc1,nvc2 = st.columns([1,1])
                    with nvc1:
                        st.markdown(f'<div style="padding-top:9px;font-size:13px;color:#E8EEF7">{nv}</div>',unsafe_allow_html=True)
                    with nvc2:
                        nv_perm = st.selectbox(" ",PERM_LEVEL_OPTIONS,key=f"cu_perm_{nv}",label_visibility="collapsed")
                        if nv_perm != "Role Default":
                            custom_nav_selection[nv] = nv_perm.lower().replace(" ","_")

                if st.form_submit_button("Create User Account",use_container_width=True):
                    if not(new_uname and new_pw and new_fullname): st.error("Username, password and full name required.")
                    elif len(new_pw)<6: st.error("Password must be at least 6 characters.")
                    elif " " in new_uname: st.error("Username cannot contain spaces.")
                    elif new_role=="Supervisor" and not dept_assign: st.error("Supervisors must be assigned a division.")
                    else:
                        conn=get_conn()
                        division_to_save = dept_assign if dept_assign!="All Divisions" else None
                        nav_access_to_save = json.dumps(custom_nav_selection) if custom_nav_selection else None
                        try:
                            conn.execute("""INSERT INTO system_users(username,password,role,full_name,email,permissions,is_active,assigned_division,nav_access,created_by,created_at)
                                VALUES(?,?,?,?,?,?,1,?,?,?,?)""",(new_uname,new_pw,new_role,new_fullname,new_email,pd_key,division_to_save,nav_access_to_save,st.session_state.uid,datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                            conn.commit(); conn.close()
                            st.success(f"User **@{new_uname}** created as **{new_role}**.")
                            st.markdown(f"""<div class="cred-box">
                              <div style="font-family:'Cinzel',serif;font-size:12px;color:#D4A847;margin-bottom:12px">LOGIN CREDENTIALS</div>
                              <div class="cred-row"><span class="cred-label">Username</span><span class="cred-val">{new_uname}</span></div>
                              <div class="cred-row"><span class="cred-label">Password</span><span class="cred-val">{new_pw}</span></div>
                              <div class="cred-row"><span class="cred-label">Role</span><span style="color:#94A8C8;font-size:13px">{new_role}</span></div>
                              <div class="cred-row"><span class="cred-label">Division</span><span style="color:#94A8C8;font-size:13px">{division_to_save or "All Divisions"}</span></div>
                              <div class="cred-row"><span class="cred-label">Navigation Access</span><span style="color:#94A8C8;font-size:13px">{", ".join(custom_nav_selection.keys()) if custom_nav_selection else "Role Default"}</span></div>
                            </div>""",unsafe_allow_html=True)
                        except sqlite3.IntegrityError:
                            conn.close(); st.error(f"Username '@{new_uname}' already exists.")

        with at3:
            conn=get_conn()
            edit_df=pd.read_sql_query("SELECT username,full_name,role,permissions,email,is_active FROM system_users WHERE username != ? ORDER BY username",conn,params=(st.session_state.uid,))
            conn.close()
            if len(edit_df)==0: st.info("No other users to edit.")
            else:
                edit_opts={f"@{r['username']} — {r['full_name'] or ''} ({r['role']})":r['username'] for _,r in edit_df.iterrows()}
                sel_edit=st.selectbox("Select User",list(edit_opts.keys()))
                edit_uid=edit_opts[sel_edit]
                conn=get_conn(); cur=conn.cursor()
                cur.execute("SELECT username,full_name,role,email,permissions,is_active,assigned_division,nav_access FROM system_users WHERE username=?",(edit_uid,))
                eu=cur.fetchone(); conn.close()
                if eu:
                    rc2=role_cls(eu[2])
                    st.markdown(f"""<div class="user-card uc-{rc2}">
                      <div class="user-avatar ua-{rc2}">{role_init(eu[1])}</div>
                      <div class="user-info"><div class="user-name">{eu[1] or eu[0]}</div>
                      <div class="user-uid">@{eu[0]}</div></div></div>""",unsafe_allow_html=True)
                    try:
                        existing_nav_access = json.loads(eu[7]) if eu[7] else {}
                    except: existing_nav_access = {}

                    with st.form("edit_user_form"):
                        e1,e2,e3=st.columns(3)
                        with e1: e_fullname=st.text_input("Full Name",value=eu[1] or "")
                        with e2: e_email=st.text_input("Email",value=eu[3] or "")
                        with e3:
                            roles=["Supervisor","Payroll Section","Data Officer","HR Staff","Payroll Officer","Department Head","Manager"]
                            e_role=st.selectbox("Role",roles,index=roles.index(eu[2]) if eu[2] in roles else 0)
                        div_list_edit=get_division_list()
                        cur_div_edit=eu[6] or "All Divisions"
                        e_division=st.selectbox("Assigned Division",["All Divisions"]+div_list_edit,
                            index=(["All Divisions"]+div_list_edit).index(cur_div_edit) if cur_div_edit in ["All Divisions"]+div_list_edit else 0,
                            help="Required for Supervisor role — controls which division's attendance they manage.")
                        ep1,ep2=st.columns(2)
                        with ep1: e_newpw=st.text_input("New Password",placeholder="Blank = keep current",type="password")
                        with ep2: e_confirmpw=st.text_input("Confirm Password",type="password")

                        st.markdown('<div class="fs" style="margin-top:14px">Navigation Access — Set Authority Per Module</div>',unsafe_allow_html=True)
                        st.markdown('<div style="font-size:11px;color:#6B7FA3;margin-bottom:8px">For every module, choose this user\'s authority: Role Default, View (read-only), Edit (change data, no delete), Both (view and edit), or Full Control (view, edit and delete).</div>',unsafe_allow_html=True)
                        edit_nav_selection = {}
                        PERM_LEVEL_OPTIONS_EDIT = ["Role Default","View","Edit","Both","Full Control"]
                        perm_levels_map = {"view":"View","edit":"Edit","both":"Both","full_control":"Full Control","delete":"Full Control"}
                        for ev in ALL_NAV_VIEWS:
                            evc1,evc2 = st.columns([1,1])
                            with evc1:
                                st.markdown(f'<div style="padding-top:9px;font-size:13px;color:#E8EEF7">{ev}</div>',unsafe_allow_html=True)
                            with evc2:
                                existing_raw = existing_nav_access.get(ev)
                                existing_display = perm_levels_map.get(existing_raw,"Role Default") if existing_raw else "Role Default"
                                default_perm_idx = PERM_LEVEL_OPTIONS_EDIT.index(existing_display) if existing_display in PERM_LEVEL_OPTIONS_EDIT else 0
                                ev_perm = st.selectbox(" ",PERM_LEVEL_OPTIONS_EDIT,index=default_perm_idx,key=f"eu_perm_{ev}",label_visibility="collapsed")
                                if ev_perm != "Role Default":
                                    edit_nav_selection[ev] = ev_perm.lower().replace(" ","_")

                        perm_map={"Data Officer":"view_only","HR Staff":"hr_edit","Payroll Officer":"payroll_edit",
                            "Department Head":"dept_view","Manager":"full","Supervisor":"division_control","Payroll Section":"payroll_approve"}
                        new_perm2=perm_map.get(e_role,"view_only")
                        if st.form_submit_button("Save Changes",use_container_width=True):
                            if e_newpw and e_newpw!=e_confirmpw: st.error("Passwords do not match.")
                            elif e_newpw and len(e_newpw)<6: st.error("Min 6 characters.")
                            elif e_role=="Supervisor" and e_division=="All Divisions": st.error("Supervisors must be assigned a specific division.")
                            else:
                                conn=get_conn()
                                division_to_save = e_division if e_division!="All Divisions" else None
                                nav_access_to_save = json.dumps(edit_nav_selection) if edit_nav_selection else None
                                if e_newpw:
                                    conn.execute("UPDATE system_users SET full_name=?,email=?,role=?,permissions=?,assigned_division=?,nav_access=?,password=? WHERE username=?",(e_fullname,e_email,e_role,new_perm2,division_to_save,nav_access_to_save,e_newpw,edit_uid))
                                else:
                                    conn.execute("UPDATE system_users SET full_name=?,email=?,role=?,permissions=?,assigned_division=?,nav_access=? WHERE username=?",(e_fullname,e_email,e_role,new_perm2,division_to_save,nav_access_to_save,edit_uid))
                                conn.commit(); conn.close()
                                st.success(f"@{edit_uid} updated."); st.rerun()