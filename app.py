import streamlit as st
import pandas as pd
import sqlite3
import hashlib
import io
import re
import json
import os
import math
import zlib
import base64
import urllib.parse
import smtplib
from email.mime.text import MIMEText
import random
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import psycopg2
from psycopg2.extras import RealDictCursor

# ==========================================
# DIRECT SUPABASE POSTGRESQL ENGINE (UNIFIED SINGLE DB)
# ==========================================
def get_pg_url():
    pg_url = None
    if hasattr(st, "secrets") and "postgres" in st.secrets:
        pg_url = st.secrets["postgres"].get("url")
    if not pg_url:
        pg_url = "postgresql://postgres.mcmesfmsshwfxvtkhnoh:rcPiJjPU5wFrt7mV@aws-0-ap-southeast-1.pooler.supabase.com:5432/postgres"
    return pg_url

def _translate_sql_for_pg(sql):
    s = str(sql)
    # 1. Replace ? placeholders with %s
    s = s.replace("?", "%s")
    # 2. Translate SQLite rowid ordering to PostgreSQL ctid
    s = re.sub(r"\browid\b", "ctid", s, flags=re.IGNORECASE)
    # 3. Translate SQLite 'Alias' single-quoted column aliases to double-quoted "Alias"
    s = re.sub(r"\bAS\s+'([^']+)'", r'AS "\1"', s, flags=re.IGNORECASE)
    # 4. Translate INSERT OR IGNORE INTO students (...) VALUES (...)
    if re.search(r"INSERT\s+OR\s+IGNORE\s+INTO\s+students", s, flags=re.IGNORECASE):
        s = re.sub(r"INSERT\s+OR\s+IGNORE\s+INTO", "INSERT INTO", s, flags=re.IGNORECASE)
        if "ON CONFLICT" not in s.upper():
            s = s.rstrip().rstrip(";") + " ON CONFLICT (cnic) DO NOTHING"
    # 5. Translate INSERT OR REPLACE INTO students (cols) VALUES (...)
    m_rep = re.search(r"INSERT\s+OR\s+REPLACE\s+INTO\s+students\s*\(([^)]+)\)", s, flags=re.IGNORECASE)
    if m_rep:
        cols = [c.strip() for c in m_rep.group(1).split(",")]
        non_pk = [c for c in cols if c.lower() != "cnic"]
        set_clause = ", ".join([f"{c} = EXCLUDED.{c}" for c in non_pk])
        s = re.sub(r"INSERT\s+OR\s+REPLACE\s+INTO", "INSERT INTO", s, flags=re.IGNORECASE)
        if "ON CONFLICT" not in s.upper():
            s = s.rstrip().rstrip(";") + f" ON CONFLICT (cnic) DO UPDATE SET {set_clause}"
    return s

class PgCompatCursor:
    def __init__(self, raw_cursor):
        self._cur = raw_cursor

    def execute(self, query, params=None):
        pg_q = _translate_sql_for_pg(query)
        if params is not None:
            clean_params = tuple(
                psycopg2.Binary(p) if isinstance(p, (bytes, bytearray, memoryview)) else p
                for p in params
            )
            return self._cur.execute(pg_q, clean_params)
        return self._cur.execute(pg_q)

    def executemany(self, query, seq_of_params):
        pg_q = _translate_sql_for_pg(query)
        clean_seq = [
            tuple(psycopg2.Binary(p) if isinstance(p, (bytes, bytearray, memoryview)) else p for p in row)
            for row in seq_of_params
        ]
        return self._cur.executemany(pg_q, clean_seq)

    def fetchone(self):
        row = self._cur.fetchone()
        if row is None:
            return None
        return tuple(bytes(v) if isinstance(v, memoryview) else v for v in row)

    def fetchall(self):
        rows = self._cur.fetchall()
        return [
            tuple(bytes(v) if isinstance(v, memoryview) else v for v in r)
            for r in rows
        ]

    @property
    def description(self):
        return self._cur.description

    @property
    def rowcount(self):
        return self._cur.rowcount

    def close(self):
        return self._cur.close()

class PgCompatConnection:
    def __init__(self, pg_url):
        self._pg_url = pg_url
        self._conn = psycopg2.connect(pg_url, connect_timeout=10)

    def _ensure_open(self):
        try:
            if self._conn is None or self._conn.closed != 0:
                self._conn = psycopg2.connect(self._pg_url, connect_timeout=10)
        except Exception:
            self._conn = psycopg2.connect(self._pg_url, connect_timeout=10)

    def cursor(self):
        self._ensure_open()
        return PgCompatCursor(self._conn.cursor())

    def commit(self):
        self._ensure_open()
        res = self._conn.commit()
        try:
            st.cache_data.clear()
        except Exception:
            pass
        return res

    def rollback(self):
        try:
            if self._conn and self._conn.closed == 0:
                return self._conn.rollback()
        except Exception:
            pass

    def close(self):
        # Keep persistent pooled connection open for ultra-fast local & cloud performance
        pass

    def force_close(self):
        try:
            if self._conn and self._conn.closed == 0:
                self._conn.close()
        except Exception:
            pass

@st.cache_resource
def get_db_connection():
    pg_url = get_pg_url()
    try:
        conn = PgCompatConnection(pg_url)
        return conn, "postgres"
    except Exception as e:
        st.error(f"⚠️ Supabase Cloud Connection Error: {e}")
        conn = sqlite3.connect("ipo_management.db", check_same_thread=False)
        return conn, "sqlite"

def run_db_query(query, params=(), fetchone=False, fetchall=False, commit=False):
    conn, db_type = get_db_connection()
    c = conn.cursor()
    try:
        c.execute(query, params)
    except Exception:
        if hasattr(conn, "rollback"):
            conn.rollback()
        if hasattr(conn, "force_close"):
            conn.force_close()
        c = conn.cursor()
        c.execute(query, params)
    res = None
    if fetchone:
        res = c.fetchone()
    elif fetchall:
        res = c.fetchall()
    if commit:
        conn.commit()
    return res
# ==========================================
# EMAIL SENDER SETUP (OTP)
# ==========================================
SENDER_EMAIL = "placement.gctfsd@gmail.com"
APP_PASSWORD = "jryv lxng zhmb fpvv"

def send_otp_email(to_email, otp_code):
    try:
        msg = MIMEText(f"Assalam-o-Alaikum,\n\nIPO Portal password reset OTP: {otp_code}\n\nYeh code kisi ko mat dein.")
        msg["Subject"] = "IPO Portal - Password Reset OTP"
        msg["From"] = SENDER_EMAIL
        msg["To"] = to_email

        server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
        server.login(SENDER_EMAIL, APP_PASSWORD.replace(" ", ""))
        server.sendmail(SENDER_EMAIL, to_email, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        print("Email error:", e)
        return False

# ==========================================
# AUTHENTICATION & LOGIN LOGIC
# ==========================================
def check_user_login(username, password):
    p_hash = hashlib.sha256(str.encode(password)).hexdigest()
    query = "SELECT username, name, role FROM system_users WHERE username = ? AND password_hash = ?"
    res = run_db_query(query, (username.strip().lower(), p_hash), fetchone=True)
    if res:
        return {"username": res[0], "name": res[1], "role": res[2]}
    return None

if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
    st.session_state["user_info"] = None

if not st.session_state["logged_in"]:
    st.title("🔐 GCT Samanabad - IPO Management System")
    st.info("System access karne ke liye apna username aur password enter karein.")
    
    # 1. LOGIN FORM
    with st.form("login_box"):
        u_in = st.text_input("Username")
        p_in = st.text_input("Password", type="password")
        btn = st.form_submit_button("Login")
        
        if btn:
            user = check_user_login(u_in, p_in)
            if user:
                st.session_state["logged_in"] = True
                st.session_state["user_info"] = user
                st.success(f"Khush Amdeed, {user['name']}!")
                st.rerun()
            else:
                st.error("Ghalat Username ya Password! Dobara koshish karein.")

    # 2. FORGOT PASSWORD
    with st.expander("🔑 Forgot Password?"):
        reset_user = st.text_input("Apna Username likhein", key="reset_u_input")
        
        if st.button("Send Recovery OTP"):
            query = "SELECT email FROM system_users WHERE username = ?"
            res = run_db_query(query, (reset_user.strip().lower(),), fetchone=True)
            
            if res and res[0]:
                otp = str(random.randint(100000, 999999))
                upd_query = "UPDATE system_users SET otp = ? WHERE username = ?"
                run_db_query(upd_query, (otp, reset_user.strip().lower()), commit=True)
                
                if send_otp_email(res[0], otp):
                    st.success(f"OTP code aapki registered email ({res[0]}) par bhej diya gaya hai!")
                    st.session_state["otp_sent_to"] = reset_user.strip().lower()
                else:
                    st.error("Email send karne mein masla aaya. Internet ya Gmail app password check karein.")
            else:
                st.error("Yeh username mojood nahi ya email link nahi hai.")

        if st.session_state.get("otp_sent_to") == reset_user.strip().lower():
            entered_otp = st.text_input("Email par aaya hua 6-digit OTP", key="entered_otp")
            new_password_set = st.text_input("Naya Password", type="password", key="new_password_set")
            
            if st.button("Reset Password"):
                q_verify = "SELECT otp FROM system_users WHERE username = ?"
                db_otp = run_db_query(q_verify, (reset_user.strip().lower(),), fetchone=True)
                
                if db_otp and str(db_otp[0]) == str(entered_otp.strip()):
                    new_hash = hashlib.sha256(str.encode(new_password_set)).hexdigest()
                    q_update_pass = "UPDATE system_users SET password_hash = ?, otp = NULL WHERE username = ?"
                    run_db_query(q_update_pass, (new_hash, reset_user.strip().lower()), commit=True)
                    st.success("Password successfully update ho gaya! Ab naye password se login karein.")
                    del st.session_state["otp_sent_to"]
                    st.rerun()
                else:
                    st.error("Ghalat OTP! Dobara check karein.")

    st.stop()

# Sidebar par user profile aur logout button
st.sidebar.markdown(f"👤 **{st.session_state['user_info']['name']}**")
st.sidebar.markdown(f"Role: `{st.session_state['user_info']['role']}`")
if st.sidebar.button("Logout"):
    st.session_state["logged_in"] = False
    st.session_state["user_info"] = None
    st.rerun()
st.sidebar.markdown("---")

# ==========================================
# LIVE SUPABASE CLOUD STATUS & ONE-TIME MIGRATION
# ==========================================
st.sidebar.success("☁️ Live Connected: Supabase Cloud DB")
with st.sidebar.expander("🧹 Remove Duplicate Cloud Records"):
    st.caption("Agar Supabase Cloud tables mein ek hi ID ke duplicate rows aa gayi hain toh yahan se foran clean karein:")
    if st.button("Clean Duplicates in Supabase"):
        with st.spinner("Cleaning duplicate records in Supabase..."):
            try:
                pg_conn = psycopg2.connect(get_pg_url())
                pg_cursor = pg_conn.cursor()
                _DEDUP_TABLES = [
                    ("system_users", "username"),
                    ("students", "cnic"),
                    ("traceability_forms", "id"),
                    ("activities", "id"),
                    ("guest_lectures", "id"),
                    ("industry_visits", "id"),
                    ("recruitment_drives", "id"),
                    ("monthly_activity_plans", "id")
                ]
                removed_total = 0
                for tbl, pk in _DEDUP_TABLES:
                    # Step A: Delete duplicate rows inside its own savepoint
                    pg_cursor.execute("SAVEPOINT sp_del_dup")
                    try:
                        pg_cursor.execute(f"""
                            DELETE FROM {tbl} a
                            USING {tbl} b
                            WHERE a.ctid < b.ctid
                              AND a.{pk} = b.{pk}
                        """)
                        removed_total += pg_cursor.rowcount
                        pg_cursor.execute("RELEASE SAVEPOINT sp_del_dup")
                    except Exception as del_err:
                        pg_cursor.execute("ROLLBACK TO SAVEPOINT sp_del_dup")
                        print(f"Dedup delete skipped on {tbl}: {del_err}")

                    # Step B: Separately add PRIMARY KEY if not already present (so failure never undoes Step A)
                    pg_cursor.execute("SAVEPOINT sp_add_pk")
                    try:
                        pg_cursor.execute(f"ALTER TABLE {tbl} ADD PRIMARY KEY ({pk})")
                        pg_cursor.execute("RELEASE SAVEPOINT sp_add_pk")
                    except Exception:
                        pg_cursor.execute("ROLLBACK TO SAVEPOINT sp_add_pk")
                pg_conn.commit()
                pg_cursor.close()
                pg_conn.close()
                st.cache_data.clear()
                st.sidebar.success(f"✅ {removed_total} duplicate rows delete ho gayi!")
                st.rerun()
            except Exception as e:
                st.sidebar.error(f"❌ Cleanup Error: {e}")

if os.path.exists("ipo_management.db"):
    with st.sidebar.expander("📤 One-Time Local DB Upload"):
        st.caption("Agar aapke paas local `ipo_management.db` mein koi data hai jo Supabase par bhejna hai toh ek dafa click karein:")
        if st.button("Upload Local DB to Supabase"):
            with st.spinner("Uploading local records to Supabase Cloud..."):
                try:
                    pg_conn = psycopg2.connect(get_pg_url())
                    pg_cursor = pg_conn.cursor()
                    sqlite_conn = sqlite3.connect("ipo_management.db", check_same_thread=False)
                    sqlite_cursor = sqlite_conn.cursor()
                    _MIG_TABLES = [
                        ("system_users", "username"),
                        ("students", "cnic"),
                        ("traceability_forms", "id"),
                        ("activities", "id"),
                        ("guest_lectures", "id"),
                        ("industry_visits", "id"),
                        ("recruitment_drives", "id"),
                        ("monthly_activity_plans", "id")
                    ]
                    total_up = 0
                    for table, pk_col in _MIG_TABLES:
                        try:
                            sqlite_cursor.execute(f"SELECT * FROM {table}")
                            rows = sqlite_cursor.fetchall()
                            col_names = [desc[0] for desc in sqlite_cursor.description]
                            if rows:
                                pg_cursor.execute("SAVEPOINT sp_mig")
                                try:
                                    # Delete existing matching keys first so no duplicate can ever be inserted even without PK constraint
                                    pk_idx = col_names.index(pk_col)
                                    for row in rows:
                                        pg_cursor.execute(f"DELETE FROM {table} WHERE {pk_col} = %s", (row[pk_idx],))
                                    placeholders = ", ".join(["%s" for _ in col_names])
                                    cols_str = ", ".join(col_names)
                                    for row in rows:
                                        clean_row = [psycopg2.Binary(val) if isinstance(val, (bytes, bytearray)) else val for val in row]
                                        pg_cursor.execute(
                                            f"INSERT INTO {table} ({cols_str}) VALUES ({placeholders})",
                                            clean_row
                                        )
                                    pg_cursor.execute("RELEASE SAVEPOINT sp_mig")
                                    total_up += len(rows)
                                except Exception as inner_e:
                                    pg_cursor.execute("ROLLBACK TO SAVEPOINT sp_mig")
                                    print(f"Migration skipped for {table}: {inner_e}")
                        except Exception as t_err:
                            print(f"Local table {table} read skipped: {t_err}")
                    pg_conn.commit()
                    pg_cursor.close()
                    pg_conn.close()
                    sqlite_conn.close()
                    st.sidebar.success(f"✅ {total_up} records Supabase Cloud par save ho gaye!")
                    st.rerun()
                except Exception as e:
                    st.sidebar.error(f"❌ Upload Failed: {e}")
st.sidebar.markdown("---")

# ==========================================
# GLOBAL DELETE PROTECTION (Admin Only)
# ==========================================
_original_st_button = st.button

def protected_button(label, *args, **kwargs):
    # Agar button ke naam mein delete/remove ho aur user Admin na ho
    if any(w in str(label).lower() for w in ["delete", "remove", "clear table", "drop"]):
        if st.session_state.get("user_info", {}).get("role") != "Administrator":
            st.caption(f"🔒 `{label}` (Sirf Admin ke liye)")
            return False
    return _original_st_button(label, *args, **kwargs)

st.button = protected_button
# ==========================================

# ==========================================
# ADMIN USER MANAGEMENT PANEL
# ==========================================
if st.session_state["user_info"]["role"] == "Administrator":
    with st.sidebar.expander("⚙️ User Management (Admin Only)"):
        st.markdown("#### Naya User Banayein")
        new_u = st.text_input("Username", key="new_u_key")
        new_n = st.text_input("Pura Naam", key="new_n_key")
        new_p = st.text_input("Password", type="password", key="new_p_key")
        new_role = st.selectbox("Role", ["Operator", "Staff"], key="new_role_key")
        
        if st.button("Create User"):
            if new_u and new_p:
                conn_u, _ = get_db_connection()
                c_u = conn_u.cursor()
                p_hash = hashlib.sha256(str.encode(new_p)).hexdigest()
                try:
                    c_u.execute("INSERT INTO system_users (username, name, password_hash, role) VALUES (?, ?, ?, ?)", (new_u.strip().lower(), new_n, p_hash, new_role))
                    conn_u.commit()
                    st.success(f"User '{new_u}' ban gaya!")
                except Exception:
                    conn_u.rollback()
                    st.error("Yeh username pehle se mojood hai.")
                conn_u.close()
                st.rerun()
            else:
                st.warning("Username aur Password likhna zaroori hai.")

        st.markdown("---")
        st.markdown("#### Mojooda Users")
        conn_u, _ = get_db_connection()
        c_u = conn_u.cursor()
        c_u.execute("SELECT username, name, role FROM system_users WHERE role != 'Administrator'")
        all_users = c_u.fetchall()
        conn_u.close()
        
        for u in all_users:
            col_a, col_b = st.columns([3, 1])
            col_a.write(f"**{u[1]}** (`{u[0]}`)\n*{u[2]}*")
            if col_b.button("❌", key=f"del_user_{u[0]}"):
                conn_u, _ = get_db_connection()
                c_u = conn_u.cursor()
                c_u.execute("DELETE FROM system_users WHERE username = ?", (u[0],))
                conn_u.commit()
                conn_u.close()
                st.success(f"{u[0]} delete ho gaya.")
                st.rerun()
# ==========================================
# ==========================================

# PDF aur Image Parsing Support
try:
    import pypdf
except ImportError:
    pypdf = None

# Optional PDF rendering support for scanned/image-only Traceability Forms
try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

try:
    from PIL import Image
    import pytesseract
except ImportError:
    Image = None
    pytesseract = None

# ----------------- Helper Functions: WhatsApp Link Builder -----------------


def safe_read_file(uploaded_file):
    """
    Safely reads CSV or Excel files handling various Windows/Excel encodings
    such as UTF-8, UTF-8-SIG, Windows-1252 (cp1252), and Latin-1.
    """
    fname = uploaded_file.name.lower()
    if fname.endswith('.csv'):
        if hasattr(uploaded_file, 'seek'):
            uploaded_file.seek(0)
        encodings = ['utf-8', 'utf-8-sig', 'cp1252', 'latin1', 'iso-8859-1']
        for enc in encodings:
            try:
                if hasattr(uploaded_file, 'seek'):
                    uploaded_file.seek(0)
                return pd.read_csv(uploaded_file, dtype=str, encoding=enc)
            except (UnicodeDecodeError, UnicodeError):
                continue
        if hasattr(uploaded_file, 'seek'):
            uploaded_file.seek(0)
        return pd.read_csv(uploaded_file, dtype=str, encoding='utf-8', errors='replace')
    else:
        if hasattr(uploaded_file, 'seek'):
            uploaded_file.seek(0)
        return pd.read_excel(uploaded_file, dtype=str)

def format_whatsapp_number(raw_num):
    if not raw_num or pd.isna(raw_num):
        return ""
    digits = re.sub(r"[^\d]", "", str(raw_num).strip())
    if digits.startswith("0092"):
        digits = digits[2:]
    elif digits.startswith("0"):
        digits = "92" + digits[1:]
    elif digits.startswith("3") and len(digits) == 10:
        digits = "92" + digits
    return digits

def generate_whatsapp_url(phone_num, message_text):
    clean_p = format_whatsapp_number(phone_num)
    if not clean_p or len(clean_p) < 11:
        return None
    encoded_msg = urllib.parse.quote(message_text)
    return f"https://wa.me/{clean_p}?text={encoded_msg}"

# ----------------- Helper Functions: Compression & File Storage -----------------

def compress_file_bytes(raw_bytes):
    if not raw_bytes:
        return None
    return zlib.compress(raw_bytes, level=9)

def decompress_file_bytes(compressed_bytes):
    if not compressed_bytes:
        return None
    try:
        return zlib.decompress(compressed_bytes)
    except Exception:
        return compressed_bytes

# ----------------- Helper Functions: Section-Specific CV Extractor -----------------

def extract_text_from_bytes(file_bytes, fname):
    fname = fname.lower()
    text = ""

    if fname.endswith(".docx"):
        try:
            doc = Document(io.BytesIO(file_bytes))
            for p in doc.paragraphs:
                if p.text.strip():
                    text += p.text.strip() + "\n"
            for tbl in doc.tables:
                for row in tbl.rows:
                    r_text = " ".join([c.text.strip() for c in row.cells if c.text.strip()])
                    if r_text:
                        text += r_text + "\n"
        except Exception as e:
            st.error(f"Word file reading error: {e}")

    elif fname.endswith(".pdf"):
        if pypdf:
            try:
                pdf_reader = pypdf.PdfReader(io.BytesIO(file_bytes))
                for page in pdf_reader.pages:
                    page_text = page.extract_text() or ""
                    if page_text.strip():
                        text += page_text + "\n"
                    
                    if len(page_text.strip()) < 40 and pytesseract:
                        try:
                            for img_obj in page.images:
                                img = Image.open(io.BytesIO(img_obj.data))
                                ocr_t = pytesseract.image_to_string(img)
                                if ocr_t.strip():
                                    text += ocr_t + "\n"
                        except Exception:
                            pass
            except Exception as e:
                st.error(f"PDF read karne mein issue aaya: {e}")
        else:
            st.error("PyPDF library install nahi hai. Terminal mein 'pip install pypdf' run karein.")

    elif any(fname.endswith(ext) for ext in [".png", ".jpg", ".jpeg", ".webp"]):
        if pytesseract:
            try:
                img = Image.open(io.BytesIO(file_bytes))
                text = pytesseract.image_to_string(img)
            except Exception as e:
                st.error(f"Image OCR error: {e}")
        else:
            st.info("Image se text auto-extract karne ke liye Tesseract OCR darkar hai. Aap CV text manually paste bhi kar sakte hain.")

    return text.strip()

def extract_cv_sections(cv_text):
    sections = {
        'education': '',
        'experience': '',
        'general': cv_text
    }
    
    edu_regex = r"(?:education|academic|qualification|diploma|degrees?)"
    exp_regex = r"(?:experience|employment|work history|job history|internships?|career history)"
    
    edu_match = re.search(rf"{edu_regex}[\s\S]*?(?={exp_regex}|$|\n[A-Z\s]{{4,}}\n)", cv_text, re.IGNORECASE)
    if edu_match:
        sections['education'] = edu_match.group(0)
    else:
        sections['education'] = cv_text

    exp_match = re.search(rf"{exp_regex}[\s\S]*?(?={edu_regex}|$|\n[A-Z\s]{{4,}}\n)", cv_text, re.IGNORECASE)
    if exp_match:
        sections['experience'] = exp_match.group(0)
    else:
        sections['experience'] = cv_text

    return sections

def parse_cv_details(cv_text, all_techs):
    data = {
        'name': '',
        'father_name': '',
        'cnic': '',
        'phone': '',
        'email': '',
        'technology': 'Electrical',
        'session': '2023-2026',
        'company': '',
        'roll_no': ''
    }

    if not cv_text:
        return data

    sections = extract_cv_sections(cv_text)
    edu_text = sections['education']
    exp_text = sections['experience']

    cnic_match = re.search(r"\b(\d{5}-\d{7}-\d{1})\b", cv_text)
    if cnic_match:
        data['cnic'] = cnic_match.group(1)
    else:
        cnic_plain = re.search(r"\b(3\d{12})\b", cv_text)
        if cnic_plain:
            d = cnic_plain.group(1)
            data['cnic'] = f"{d[:5]}-{d[5:12]}-{d[12]}"

    phone_match = re.search(r"(?:(?:\+92|0092|92|0)?\s*3\d{2}[-\s]?\d{7})", cv_text)
    if phone_match:
        raw_ph = phone_match.group(0)
        cl_ph = re.sub(r"[^\d]", "", raw_ph)
        if cl_ph.startswith("92"):
            cl_ph = "0" + cl_ph[2:]
        if len(cl_ph) == 11 and cl_ph.startswith("03"):
            data['phone'] = f"{cl_ph[:4]}-{cl_ph[4:]}"
        else:
            data['phone'] = raw_ph.strip()

    email_match = re.search(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", cv_text)
    if email_match:
        data['email'] = email_match.group(0).strip()

    roll_match = re.search(r"(?:roll\s*(?:no|number)?[:\s\-]*)(\b[A-Za-z0-9\-]+\b)", cv_text, re.IGNORECASE)
    if roll_match:
        data['roll_no'] = roll_match.group(1).strip()

    found_tech = False
    for tech in all_techs:
        tech_core = tech.split()[0]
        if re.search(rf"\b{tech_core}\b", edu_text, re.IGNORECASE):
            data['technology'] = tech
            found_tech = True
            break
    
    if not found_tech:
        for tech in all_techs:
            tech_core = tech.split()[0]
            if re.search(rf"\b{tech_core}\b", cv_text, re.IGNORECASE):
                data['technology'] = tech
                break

    session_match = re.search(r"\b(202\d\s*[-–to]\s*202\d)\b", edu_text, re.IGNORECASE)
    if not session_match:
        session_match = re.search(r"\b(202\d\s*[-–to]\s*202\d)\b", cv_text, re.IGNORECASE)
    if session_match:
        s_clean = re.sub(r"[–\s]|to", "-", session_match.group(1))
        s_clean = re.sub(r"-+", "-", s_clean)
        data['session'] = s_clean

    lines = [line.strip() for line in cv_text.splitlines() if line.strip()]
    name_match = re.search(r"(?:name|candidate name|student name)[:\s]+([A-Za-z\s.]{3,35})", cv_text, re.IGNORECASE)
    if name_match:
        data['name'] = name_match.group(1).strip().title()
    else:
        for l in lines[:5]:
            if not any(k in l.lower() for k in ["curriculum", "resume", "cv", "bio", "data", "profile", "contact", "email", "phone", "address"]):
                if 3 < len(l) < 35 and re.match(r"^[A-Za-z\s.]+$", l):
                    data['name'] = l.strip().title()
                    break

    fname_match = re.search(r"(?:father(?:'s)?\s*name|s/o|d/o)[:\s]+([A-Za-z\s.]{3,35})", cv_text, re.IGNORECASE)
    if fname_match:
        data['father_name'] = fname_match.group(1).strip().title()

    org_match = re.search(r"(?:organization|company|worked at|internship at|employer)[:\s]+([A-Za-z0-9\s&,.-]{3,40})", exp_text, re.IGNORECASE)
    if org_match:
        data['company'] = org_match.group(1).strip().title()
    else:
        for exp_line in exp_text.splitlines():
            if any(term in exp_line.lower() for term in ["limited", "ltd", "textile", "mills", "motors", "industries", "pvt", "plant"]):
                clean_org = re.sub(r"(?:at|as|in|from|to|dated?|experience|worked)[:\s]*", "", exp_line, flags=re.IGNORECASE).strip()
                if 3 < len(clean_org) < 45:
                    data['company'] = clean_org.title()
                    break


    return data

# ----------------- Traceability Form OCR / Field Parser -----------------

def normalize_cnic(raw_value):
    """Return a standard XXXXX-XXXXXXX-X CNIC, or empty string."""
    if raw_value is None:
        return ""
    digits = re.sub(r"[^\d]", "", str(raw_value))
    if len(digits) == 13:
        return f"{digits[:5]}-{digits[5:12]}-{digits[12]}"
    return ""

def _clean_ocr_line(value):
    value = str(value or "").replace("\r", " ").replace("\t", " ")
    value = re.sub(r"[ ]{2,}", " ", value)
    return value.strip(" :-")

def _extract_labeled_value(text, labels):
    for label in labels:
        m = re.search(rf"{label}\s*[:\-]?\s*(.+)", text, re.IGNORECASE)
        if m:
            value = _clean_ocr_line(m.group(1))
            if value:
                return value
    return ""

def extract_traceability_text(file_bytes, fname):
    """Extract text from text PDFs/images and OCR scanned PDF pages."""
    fname_l = str(fname or "").lower()
    parts = []

    try:
        base_text = extract_text_from_bytes(file_bytes, fname)
        if base_text.strip():
            parts.append(base_text.strip())
    except Exception:
        pass

    if any(fname_l.endswith(ext) for ext in [".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff"]):
        if pytesseract and Image:
            try:
                img = Image.open(io.BytesIO(file_bytes))
                ocr = pytesseract.image_to_string(img, config="--psm 6")
                if ocr.strip():
                    parts.append(ocr.strip())
            except Exception:
                pass

    if fname_l.endswith(".pdf") and fitz and pytesseract and Image:
        try:
            pdf = fitz.open(stream=file_bytes, filetype="pdf")
            for page in pdf:
                pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0), alpha=False)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                ocr = pytesseract.image_to_string(img, config="--psm 6")
                if ocr.strip():
                    parts.append(ocr.strip())
            pdf.close()
        except Exception:
            pass

    unique = []
    seen = set()
    for part in parts:
        key = re.sub(r"\s+", " ", part).strip().lower()
        if key and key not in seen:
            seen.add(key)
            unique.append(part)
    return "\n".join(unique).strip()

def parse_traceability_form(text, all_techs):
    raw = str(text or "")
    data = {
        "name": "", "cnic": "", "phone": "", "technology": "",
        "professional_status": "", "session_start_date": "",
        "designation_role": "", "monthly_salary": "", "course_degree": "",
        "admission_status": "", "joining_date": "", "business_trade_work": "",
        "estimated_income": "", "preferred_work_location": "",
        "company_name": "", "work_address_city": "", "university_name": "",
        "campus_location_city": "", "business_name": "",
        "business_address_city": "", "roll_no": ""
    }

    cnic_raw = _extract_labeled_value(raw, [r"CNIC\s+NUMBER", r"CNIC\s+NO", r"CNIC"])
    cnic_match = re.search(r"\b\d{5}\s*[-–]?\s*\d{7}\s*[-–]?\s*\d\b", cnic_raw)
    if not cnic_match:
        cnic_match = re.search(r"\b\d{5}\s*[-–]?\s*\d{7}\s*[-–]?\s*\d\b", raw)
    if cnic_match:
        digits = re.sub(r"[^\d]", "", cnic_match.group(0))
        data["cnic"] = normalize_cnic(digits)

    data["name"] = _extract_labeled_value(raw, [
        r"STUDENT\s+FULL\s+NAME", r"FULL\s+NAME", r"STUDENT\s+NAME"
    ])

    phone_raw = _extract_labeled_value(raw, [
        r"MOBILE\s*/\s*WHATSAPP\s+NUMBER", r"MOBILE\s*/\s*WHATSAPP", r"MOBILE\s+NUMBER"
    ])
    phone_match = re.search(r"(?:0?3\d{2}\s*[-]?\s*\d{7}|92\s*3\d{2}\s*[-]?\s*\d{7})", phone_raw)
    if not phone_match:
        phone_match = re.search(r"(?:0?3\d{2}\s*[-]?\s*\d{7}|92\s*3\d{2}\s*[-]?\s*\d{7})", raw)
    if phone_match:
        ph = re.sub(r"[^\d]", "", phone_match.group(0))
        if ph.startswith("92"):
            ph = "0" + ph[2:]
        if len(ph) == 11 and ph.startswith("03"):
            data["phone"] = f"{ph[:4]}-{ph[4:]}"
    else:
        data["phone"] = phone_raw

    for tech in all_techs:
        if re.search(rf"\b{re.escape(tech)}\b", raw, re.IGNORECASE):
            data["technology"] = tech
            break

    status_patterns = [
        (r"EMPLOYED\s*/\s*ON\s*JOB", "Employed / On Job"),
        (r"HIGHER\s+STUDIES", "Higher Studies"),
        (r"SELF[-\s]?EMPLOYED", "Self-Employed"),
        (r"SEEKING\s+EMPLOYMENT", "Seeking Employment"),
    ]
    for pat, label in status_patterns:
        m = re.search(rf".{{0,100}}{pat}.{{0,40}}", raw, re.IGNORECASE | re.DOTALL)
        if m and any(x in m.group(0) for x in ["✓", "✔", "☑", "X", "x"]):
            data["professional_status"] = label
            break

    data["session_start_date"] = _extract_labeled_value(raw, [r"SESSION\s+START\s+DATE", r"START\s+DATE"])
    data["session_start_date"] = re.sub(r"^\(DD\s*-\s*MM\s*-\s*YYYY\)\s*:?\s*", "", data["session_start_date"], flags=re.IGNORECASE)
    data["designation_role"] = _extract_labeled_value(raw, [r"DESIGNATION\s*/\s*ROLE", r"DESIGNATION", r"ROLE"])
    data["monthly_salary"] = _extract_labeled_value(raw, [r"MONTHLY\s+SALARY"])
    data["monthly_salary"] = re.sub(r"^\(PKR\)\s*:?\s*(?:Rs\.?\s*)?", "", data["monthly_salary"], flags=re.IGNORECASE)
    data["course_degree"] = _extract_labeled_value(raw, [r"COURSE\s*/\s*DEGREE\s+PROGRAM", r"COURSE", r"DEGREE\s+PROGRAM"])
    data["admission_status"] = _extract_labeled_value(raw, [r"ADMISSION\s+STATUS"])
    data["joining_date"] = _extract_labeled_value(raw, [r"DATE\s+OF\s+JOINING", r"JOINING\s+DATE"])
    data["business_trade_work"] = _extract_labeled_value(raw, [r"NATURE\s+OF\s+BUSINESS\s*/\s*TRADE\s+WORK"])
    data["estimated_income"] = _extract_labeled_value(raw, [r"ESTIMATED\s+MONTHLY\s+INCOME"])
    data["preferred_work_location"] = _extract_labeled_value(raw, [r"PREFERRED\s+WORK\s+LOCATION"])
    data["company_name"] = _extract_labeled_value(raw, [
        r"COMPANY\s*/\s*INDUSTRY\s*/\s*EMPLOYER\s+NAME", r"EMPLOYER\s+NAME", r"COMPANY\s+NAME"
    ])
    data["work_address_city"] = _extract_labeled_value(raw, [r"WORK\s+ADDRESS\s*/\s*CITY", r"WORK\s+ADDRESS"])
    data["university_name"] = _extract_labeled_value(raw, [r"UNIVERSITY\s*/\s*INSTITUTE\s+NAME", r"UNIVERSITY\s+NAME", r"INSTITUTE\s+NAME"])
    data["campus_location_city"] = _extract_labeled_value(raw, [r"CAMPUS\s+LOCATION\s*/\s*CITY", r"CAMPUS\s+LOCATION"])
    data["business_name"] = _extract_labeled_value(raw, [r"BUSINESS\s*/\s*COMMERCIAL\s+SETUP\s*/\s*ENTERPRISE\s+NAME", r"BUSINESS\s+NAME"])
    data["business_address_city"] = _extract_labeled_value(raw, [r"SHOP\s*/\s*SETUP\s+ADDRESS\s*&\s*CITY", r"SETUP\s+ADDRESS\s*&\s*CITY"])
    data["roll_no"] = _extract_labeled_value(raw, [r"ROLL\s+NO", r"ROLL\s+NUMBER"])

    if not data["roll_no"]:
        roll_match = re.search(r"\b[A-Za-z0-9]{1,4}\s*-\s*[A-Za-z0-9]{1,4}\s*-\s*[A-Za-z0-9]{1,4}\b", raw)
        if roll_match:
            data["roll_no"] = re.sub(r"\s+", "", roll_match.group(0))

    return data

# ----------------- Helper Functions: Guest Speaker Docs -----------------

def generate_principal_application(data):
    doc = Document()
    
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_date = title.add_run(f"Date: {data['date']}\n")
    r_date.font.name = 'Calibri'
    r_date.font.size = Pt(11)

    to_p = doc.add_paragraph()
    to_run = to_p.add_run("To,\n    The Worthy Principal,\n    Government College of Technology,\n    Samanabad, Faisalabad.\n")
    to_run.bold = True
    to_run.font.name = 'Calibri'
    to_run.font.size = Pt(12)

    subj = doc.add_paragraph()
    subj_run = subj.add_run(f"SUBJECT: PERMISSION & ADMINISTRATIVE APPROVAL FOR GUEST SPEAKER LECTURE ({data['tech'].upper()})")
    subj_run.bold = True
    subj_run.font.name = 'Calibri'
    subj_run.font.size = Pt(11)

    body = doc.add_paragraph()
    body.paragraph_format.line_spacing = 1.15
    body_text = (
        f"Respected Sir,\n\n"
        f"It is submitted that the Placement Cell (IPO Office) has planned a guest speaker lecture "
        f"to bridge industry-academia linkages and provide practical insights to students. Details are as follows:\n\n"
        f"• Speaker Name: {data['spk_name']} ({data['designation']})\n"
        f"• Background/Affiliation: {data['speaker_type']} - {data['org']}\n"
        f"• Qualification: {data['qualification']}\n"
        f"• Topic: {data['topic']}\n"
        f"• Target Class/Technology: DAE {data['tech']} (Session: {data['session']})\n"
        f"• Proposed Date & Venue: {data['date']} at {data['venue']}\n\n"
        f"To organize this event successfully, following administrative/protocol arrangements are requested:\n"
        f"  1. Refreshment: For approximately {data['refreshment_heads']} persons (Guest & Faculty staff).\n"
        f"  2. Stage/Display: {data['backdrop_details']}.\n"
        f"  3. Protocol/Welcome: {data['bouquet_details']}.\n\n"
        f"Submitted for your kind approval and administrative sanction, please."
    )
    b_run = body.add_run(body_text)
    b_run.font.name = 'Calibri'
    b_run.font.size = Pt(11)

    sig = doc.add_paragraph()
    sig.paragraph_format.space_before = Pt(35)
    sig_run = sig.add_run(
        "Institute Placement Officer (IPO)\n"
        "GCT Samanabad, Faisalabad\n\n\n"
        "Approved / Not Approved:\n\n"
        "_________________________\n"
        "PRINCIPAL\n"
        "GCT Samanabad, Faisalabad"
    )
    sig_run.bold = True
    sig_run.font.name = 'Calibri'

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io


EMBEDDED_ANNEXURE_E_B64 = "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"
EMBEDDED_ANNEXURE_F_B64 = "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"

def generate_annexure_e(data):
    template_raw = get_official_template_bytes('E')
    doc = Document(io.BytesIO(template_raw))

    district_val = str(data.get('district', 'Faisalabad'))
    name_val = str(data.get('speaker_name', ''))
    gender_val = str(data.get('gender', 'Male'))
    profession_val = str(data.get('profession', 'Industrialist / Professional'))
    qual_val = str(data.get('qualification', ''))
    org_val = str(data.get('company_org', ''))
    org_addr = str(data.get('org_address', 'Faisalabad'))
    desig_val = str(data.get('designation', ''))
    exp_val = str(data.get('experience', '5'))
    cnic_val = str(data.get('cnic', ''))
    contact_val = str(data.get('contact_no', ''))
    email_val = str(data.get('email', ''))
    is_reg = str(data.get('is_registered_employer', 'Yes'))
    emp_code = str(data.get('emp_code', 'N/A'))

    for p in doc.paragraphs:
        txt = p.text
        if "District of Guest:" in txt:
            p.text = f"District of Guest:                {district_val}"
        elif "Name:" in txt and "Organization" not in txt and "Gender" not in txt:
            p.text = f"Name:                                {name_val}"
        elif "Gender:" in txt:
            p.text = f"Gender:                             {gender_val}"
        elif "Profession:" in txt:
            p.text = f"Profession:                        {profession_val}"
        elif "Qualification:" in txt:
            p.text = f"Qualification:                     {qual_val}"
        elif "Organization:" in txt and "Address" not in txt:
            p.text = f"Organization:                    {org_val}"
        elif "Organization Address:" in txt:
            p.text = f"Organization Address:       {org_addr}"
        elif "Designation:" in txt:
            p.text = f"Designation:                      {desig_val}"
        elif "Experience in Years:" in txt:
            p.text = f"Experience in Years:        {exp_val}"
        elif "CNIC:" in txt:
            p.text = f"CNIC:                                {cnic_val}"
        elif "Contact Number:" in txt:
            p.text = f"Contact Number:              {contact_val}"
        elif "Email:" in txt:
            p.text = f"Email:                               {email_val}"
        elif "Is guest speaker from register employer:" in txt:
            p.text = f"Is guest speaker from register employer:                    {is_reg}"
        elif "If yes, Then employer career dashboard code:" in txt:
            p.text = f"If yes, Then employer career dashboard code: {emp_code}"
            
        for r in p.runs:
            r.font.name = 'Calibri'
            r.font.size = Pt(11)

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io.getvalue()

def generate_annexure_f(data):
    template_raw = get_official_template_bytes('F')
    doc = Document(io.BytesIO(template_raw))

    spk_name = str(data.get('speaker_name', ''))
    qual = str(data.get('qualification', ''))
    ind_name = str(data.get('company_org', ''))
    emp_code = str(data.get('emp_code', 'N/A'))
    topic = str(data.get('topic', ''))
    trade_class = f"DAE {str(data.get('target_tech', ''))} ({str(data.get('target_session', ''))})"
    attendees = str(data.get('total_attendees', ''))

    for p in doc.paragraphs:
        txt = p.text
        if "Name of The Speaker:" in txt:
            p.text = f"Name of The Speaker:\t\t\t{spk_name}"
        elif "Qualification:" in txt:
            p.text = f"Qualification:\t\t\t\t{qual}"
        elif "Name of Industry:" in txt:
            p.text = f"Name of Industry:\t\t\t{ind_name}"
        elif "Speaker Employer Code:" in txt:
            p.text = f"Speaker Employer Code:\t\t\t{emp_code}"
        elif "Topic of Guest Speaker Lecture" in txt:
            p.text = f"Topic of Guest Speaker Lecture\t\t\t{topic}"
        elif "Trade/Class/Section" in txt:
            p.text = f"Trade/Class/Section\t\t\t{trade_class}"
        elif "No. of Students Attended Lecture" in txt:
            p.text = f"No. of Students Attended Lecture\t\t{attendees}"
            
        for r in p.runs:
            r.font.name = 'Calibri'
            r.font.size = Pt(11)

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io.getvalue()


EMBEDDED_TEVTA_LOGO_B64 = "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"
EMBEDDED_PUNJAB_LOGO_B64 = "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"
EMBEDDED_PRINCIPAL_SIG_B64 = "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"
EMBEDDED_IPO_SIG_B64 = "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"

def _extract_signature_rgba(raw_bytes):
    if not raw_bytes:
        return None
    from PIL import Image as PilImage
    data_bytes = decompress_file_bytes(raw_bytes)
    im_s = PilImage.open(io.BytesIO(data_bytes)).convert("RGBA")
    px = im_s.load()
    sw, sh = im_s.size
    for y in range(sh):
        for x in range(sw):
            r, g, b, a = px[x, y]
            if a == 0:
                continue
            lum = 0.299 * r + 0.587 * g + 0.114 * b
            if lum > 215:
                px[x, y] = (r, g, b, 0)
            elif lum > 155:
                alpha = int(a * ((215.0 - lum) / 60.0))
                px[x, y] = (max(0, r - 25), max(0, g - 25), max(0, b - 15), max(0, min(255, alpha)))
            else:
                px[x, y] = (max(0, r - 20), max(0, g - 20), max(0, b - 10), a)
    bbox = im_s.getbbox()
    if bbox:
        im_s = im_s.crop(bbox)
    return im_s

def _get_cert_font(size, bold=False, italic=False, sans=False):
    from PIL import ImageFont
    candidates = []
    if sans:
        if bold:
            candidates += ["arialbd.ttf", "calibrib.ttf", "LiberationSans-Bold.ttf", "DejaVuSans-Bold.ttf"]
        else:
            candidates += ["arial.ttf", "calibri.ttf", "LiberationSans-Regular.ttf", "DejaVuSans.ttf"]
    else:
        if bold and italic:
            candidates += ["georgiaz.ttf", "timesbi.ttf", "LiberationSerif-BoldItalic.ttf", "DejaVuSerif-BoldItalic.ttf"]
        elif bold:
            candidates += ["georgiab.ttf", "timesbd.ttf", "LiberationSerif-Bold.ttf", "DejaVuSerif-Bold.ttf"]
        elif italic:
            candidates += ["georgiai.ttf", "timesi.ttf", "LiberationSerif-Italic.ttf", "DejaVuSerif-Italic.ttf"]
        else:
            candidates += ["georgia.ttf", "times.ttf", "LiberationSerif-Regular.ttf", "DejaVuSerif.ttf"]
    search_dirs = ["", "C:/Windows/Fonts", "/usr/share/fonts/truetype/liberation", "/usr/share/fonts/truetype/dejavu", "/usr/share/fonts"]
    for d in search_dirs:
        for c in candidates:
            p = os.path.join(d, c) if d else c
            try:
                return ImageFont.truetype(p, size)
            except Exception:
                continue
    return ImageFont.load_default()

def _draw_cert_laurel_wreath(draw, cx, cy, radius, color_main=(206, 155, 42), color_hi=(238, 198, 92)):
    for side in (-1, 1):
        pts_stem = []
        for deg in range(-58, 66, 4):
            rad = math.radians(deg)
            sx = cx + side * (radius * math.cos(rad))
            sy = cy + (radius * math.sin(rad))
            pts_stem.append((sx, sy))
        if len(pts_stem) > 1:
            draw.line(pts_stem, fill=color_main, width=4)
        for deg in range(-52, 60, 12):
            rad = math.radians(deg)
            bx = cx + side * (radius * math.cos(rad))
            by = cy + (radius * math.sin(rad))
            for spread in (0.42, -0.35):
                ang = rad + side * spread
                lx = bx + side * 24 * math.cos(ang)
                ly = by + 24 * math.sin(ang)
                mx1 = (bx + lx) / 2 - 6 * math.sin(ang)
                my1 = (by + ly) / 2 + side * 6 * math.cos(ang)
                mx2 = (bx + lx) / 2 + 6 * math.sin(ang)
                my2 = (by + ly) / 2 - side * 6 * math.cos(ang)
                draw.polygon([(bx, by), (mx1, my1), (lx, ly), (mx2, my2)], fill=color_main, outline=color_hi)

def _draw_cert_star(draw, cx, cy, r_outer, r_inner, fill=(206, 155, 42)):
    pts = []
    for i in range(10):
        ang = math.radians(-90 + i * 36)
        r = r_outer if i % 2 == 0 else r_inner
        pts.append((cx + r * math.cos(ang), cy + r * math.sin(ang)))
    draw.polygon(pts, fill=fill)

def _paste_cert_signature(img, sig_rgba, center_x, line_y, max_w=240, max_h=95):
    if sig_rgba is None:
        return
    from PIL import Image as PilImage
    sw, sh = sig_rgba.size
    if sw <= 0 or sh <= 0:
        return
    scale = min(max_w / float(sw), max_h / float(sh))
    nw = max(1, int(sw * scale))
    nh = max(1, int(sh * scale))
    resized = sig_rgba.resize((nw, nh), PilImage.LANCZOS)
    px_pos = int(center_x - nw / 2)
    py_pos = int(line_y - nh - 5)
    img.alpha_composite(resized, (px_pos, py_pos))

def generate_guest_speaker_certificate(data, speaker_photo_bytes=None, pri_sig_bytes=None, ipo_sig_bytes=None, output_format="PDF"):
    import math
    from PIL import Image as PilImage, ImageDraw as PilDraw
    W, H = 1600, 1130
    img = PilImage.new("RGBA", (W, H), (251, 248, 239, 255))
    draw = PilDraw.Draw(img)
    for y in range(H):
        t = abs(y - H / 2) / (H / 2)
        draw.line([(0, y), (W, y)], fill=(int(253 - 10 * t), int(250 - 12 * t), int(241 - 16 * t), 255))
    EMERALD_DARK = (8, 72, 34, 255)
    EMERALD_MID = (18, 108, 52, 255)
    GOLD_DARK = (176, 128, 30, 255)
    GOLD_BRIGHT = (224, 178, 68, 255)
    draw.polygon([(0, 0), (362, 0), (275, 58), (168, 122), (86, 210), (0, 312)], fill=GOLD_BRIGHT)
    draw.polygon([(0, 0), (340, 0), (260, 48), (155, 110), (75, 195), (0, 290)], fill=EMERALD_DARK)
    draw.polygon([(W, 0), (W - 238, 0), (W - 152, 50), (W - 74, 116), (W, 202)], fill=GOLD_BRIGHT)
    draw.polygon([(W, 0), (W - 220, 0), (W - 140, 42), (W - 65, 105), (W, 185)], fill=EMERALD_DARK)
    wave_gold_outer, wave_green_main, wave_green_light = [], [], []
    for x in range(0, W + 20, 20):
        nx = x / W
        y_base = H - 155 + 58 * math.sin(nx * math.pi * 1.35 - 0.35) - 25 * math.cos(nx * math.pi * 2.0)
        wave_gold_outer.append((x, y_base - 14))
        wave_green_main.append((x, y_base))
        wave_green_light.append((x, y_base + 34))
    for arr in (wave_gold_outer, wave_green_main, wave_green_light):
        arr += [(W, H), (0, H)]
    draw.polygon(wave_gold_outer, fill=GOLD_BRIGHT)
    draw.polygon(wave_green_main, fill=EMERALD_DARK)
    draw.polygon(wave_green_light, fill=EMERALD_MID)
    draw.rectangle([28, 24, W - 28, H - 24], outline=(212, 175, 78, 180), width=2)
    try:
        t_logo = PilImage.open(io.BytesIO(base64.b64decode(EMBEDDED_TEVTA_LOGO_B64))).convert("RGBA").resize((168, 172), PilImage.LANCZOS)
        img.alpha_composite(t_logo, (118, 34))
    except Exception:
        pass
    try:
        p_logo = PilImage.open(io.BytesIO(base64.b64decode(EMBEDDED_PUNJAB_LOGO_B64))).convert("RGBA").resize((152, 175), PilImage.LANCZOS)
        img.alpha_composite(p_logo, (W - 272, 32))
    except Exception:
        pass
    f_inst = _get_cert_font(44, bold=True)
    f_sub_inst = _get_cert_font(38, bold=True)
    inst_txt = "GOVERNMENT COLLEGE OF TECHNOLOGY"
    bbox = draw.textbbox((0, 0), inst_txt, font=f_inst)
    draw.text(((W - (bbox[2] - bbox[0])) / 2, 54), inst_txt, fill=(18, 28, 22, 255), font=f_inst)
    sub_txt = "Samanabad, Faisalabad"
    bbox2 = draw.textbbox((0, 0), sub_txt, font=f_sub_inst)
    sub_w = bbox2[2] - bbox2[0]
    sub_x = (W - sub_w) / 2
    draw.text((sub_x, 112), sub_txt, fill=(18, 28, 22, 255), font=f_sub_inst)
    draw.line([(sub_x - 135, 134), (sub_x - 18, 134)], fill=GOLD_DARK, width=3)
    draw.line([(sub_x + sub_w + 18, 134), (sub_x + sub_w + 135, 134)], fill=GOLD_DARK, width=3)
    bx1, by1, bx2, by2 = 365, 182, W - 365, 258
    draw.polygon([(bx1 - 55, by1 + 14), (bx1 + 15, by1 + 14), (bx1 + 15, by2 + 12), (bx1 - 55, by2 + 12), (bx1 - 30, (by1 + by2) // 2 + 13)], fill=(6, 56, 26, 255))
    draw.polygon([(bx2 + 55, by1 + 14), (bx2 - 15, by1 + 14), (bx2 - 15, by2 + 12), (bx2 + 55, by2 + 12), (bx2 + 30, (by1 + by2) // 2 + 13)], fill=(6, 56, 26, 255))
    draw.polygon([(bx1 - 12, by1), (bx2 + 12, by1), (bx2 - 6, by2), (bx1 + 6, by2)], fill=EMERALD_DARK, outline=GOLD_BRIGHT)
    draw.line([(bx1 - 12, by1), (bx2 + 12, by1)], fill=GOLD_BRIGHT, width=4)
    draw.line([(bx1 + 6, by2), (bx2 - 6, by2)], fill=GOLD_BRIGHT, width=4)
    f_banner = _get_cert_font(40, bold=True)
    b_txt = "CERTIFICATE OF APPRECIATION"
    bb = draw.textbbox((0, 0), b_txt, font=f_banner)
    draw.text(((W - (bb[2] - bb[0])) / 2, by1 + 14), b_txt, fill=(255, 255, 255, 255), font=f_banner)
    _draw_cert_star(draw, W // 2, 286, 15, 6, fill=GOLD_DARK)
    draw.line([(W // 2 - 190, 286), (W // 2 - 26, 286)], fill=GOLD_DARK, width=2)
    draw.line([(W // 2 + 26, 286), (W // 2 + 190, 286)], fill=GOLD_DARK, width=2)
    px_c, py_c, p_rad = 265, 435, 132
    _draw_cert_laurel_wreath(draw, px_c, py_c + 5, p_rad + 32, color_main=GOLD_DARK, color_hi=GOLD_BRIGHT)
    draw.ellipse([px_c - p_rad - 10, py_c - p_rad - 10, px_c + p_rad + 10, py_c + p_rad + 10], fill=GOLD_BRIGHT, outline=GOLD_DARK, width=3)
    draw.ellipse([px_c - p_rad - 3, py_c - p_rad - 3, px_c + p_rad + 3, py_c + p_rad + 3], fill=(255, 255, 255, 255))
    portrait = None
    if speaker_photo_bytes:
        try:
            raw_pic = decompress_file_bytes(speaker_photo_bytes)
            sp_im = PilImage.open(io.BytesIO(raw_pic)).convert("RGBA")
            sw, sh = sp_im.size
            side_min = min(sw, sh)
            left = (sw - side_min) // 2
            top = max(0, (sh - side_min) // 4)
            sp_im = sp_im.crop((left, top, left + side_min, top + side_min)).resize((p_rad * 2, p_rad * 2), PilImage.LANCZOS)
            portrait = sp_im
        except Exception:
            portrait = None
    if portrait is None:
        portrait = PilImage.new("RGBA", (p_rad * 2, p_rad * 2), (232, 238, 234, 255))
        p_draw = PilDraw.Draw(portrait)
        p_draw.ellipse([p_rad - 44, 48, p_rad + 44, 136], fill=(165, 182, 172, 255))
        p_draw.pieslice([p_rad - 88, 148, p_rad + 88, 290], 180, 360, fill=(165, 182, 172, 255))
    mask = PilImage.new("L", (p_rad * 2, p_rad * 2), 0)
    PilDraw.Draw(mask).ellipse([0, 0, p_rad * 2, p_rad * 2], fill=255)
    img.paste(portrait, (px_c - p_rad, py_c - p_rad), mask)
    ry1, ry2 = py_c + p_rad - 6, py_c + p_rad + 40
    draw.polygon([(px_c - 175, ry1 + 10), (px_c - 125, ry1 + 10), (px_c - 125, ry2 + 10), (px_c - 175, ry2 + 10), (px_c - 155, (ry1 + ry2) // 2 + 10)], fill=GOLD_DARK)
    draw.polygon([(px_c + 175, ry1 + 10), (px_c + 125, ry1 + 10), (px_c + 125, ry2 + 10), (px_c + 175, ry2 + 10), (px_c + 155, (ry1 + ry2) // 2 + 10)], fill=GOLD_DARK)
    draw.rounded_rectangle([px_c - 142, ry1, px_c + 142, ry2], radius=10, fill=GOLD_BRIGHT, outline=GOLD_DARK, width=2)
    f_rib = _get_cert_font(22, bold=True, sans=True)
    r_txt = "GUEST SPEAKER"
    rb = draw.textbbox((0, 0), r_txt, font=f_rib)
    draw.text((px_c - (rb[2] - rb[0]) / 2, ry1 + 10), r_txt, fill=(20, 28, 20, 255), font=f_rib)
    _draw_cert_star(draw, px_c, ry2 + 24, 11, 5, fill=GOLD_DARK)
    draw.line([(px_c - 75, ry2 + 24), (px_c - 18, ry2 + 24)], fill=GOLD_DARK, width=2)
    draw.line([(px_c + 18, ry2 + 24), (px_c + 75, ry2 + 24)], fill=GOLD_DARK, width=2)
    cx_text = 980
    def draw_centered_text(y_pos, text, font, fill=(20, 24, 22, 255), max_width=1020):
        words = str(text or "").split()
        if not words:
            return y_pos
        lines_out, curr = [], words[0]
        for w_tok in words[1:]:
            test = curr + " " + w_tok
            tb = draw.textbbox((0, 0), test, font=font)
            if (tb[2] - tb[0]) <= max_width:
                curr = test
            else:
                lines_out.append(curr)
                curr = w_tok
        lines_out.append(curr)
        cur_y = y_pos
        for ln in lines_out:
            lb = draw.textbbox((0, 0), ln, font=font)
            draw.text((cx_text - (lb[2] - lb[0]) / 2, cur_y), ln, fill=fill, font=font)
            cur_y += (lb[3] - lb[1]) + 10
        return cur_y
    f_ital = _get_cert_font(28, italic=True)
    y_cur = draw_centered_text(318, "This certificate is proudly presented to", f_ital, fill=(30, 35, 32, 255))
    spk_name = str(data.get("speaker_name") or data.get("spk_name") or "Guest Speaker Name").strip()
    f_name = _get_cert_font(52 if len(spk_name) <= 26 else 42, bold=True)
    y_cur = draw_centered_text(y_cur + 6, spk_name, f_name, fill=(14, 22, 18, 255), max_width=1040)
    pill_y1, pill_y2 = y_cur + 6, y_cur + 60
    draw.polygon([(cx_text - 210, pill_y1), (cx_text + 210, pill_y1), (cx_text + 235, (pill_y1 + pill_y2) // 2), (cx_text + 210, pill_y2), (cx_text - 210, pill_y2), (cx_text - 235, (pill_y1 + pill_y2) // 2)], fill=EMERALD_DARK, outline=GOLD_BRIGHT)
    draw.line([(cx_text - 380, (pill_y1 + pill_y2) // 2), (cx_text - 245, (pill_y1 + pill_y2) // 2)], fill=GOLD_DARK, width=2)
    draw.line([(cx_text + 245, (pill_y1 + pill_y2) // 2), (cx_text + 380, (pill_y1 + pill_y2) // 2)], fill=GOLD_DARK, width=2)
    f_pill = _get_cert_font(34, bold=True, italic=True)
    pb = draw.textbbox((0, 0), "Guest Speaker", font=f_pill)
    draw.text((cx_text - (pb[2] - pb[0]) / 2, pill_y1 + 8), "Guest Speaker", fill=(255, 255, 255, 255), font=f_pill)
    desig = str(data.get("designation") or "").strip()
    org = str(data.get("company_org") or data.get("org") or "").strip()
    topic = str(data.get("topic") or "").strip()
    lec_date = str(data.get("lecture_date") or data.get("date") or "").strip()
    g_low = str(data.get("gender") or "Male").strip().lower()
    p_word = "her" if g_low == "female" else "his"
    y_cur = pill_y2 + 14
    if desig:
        y_cur = draw_centered_text(y_cur, desig, _get_cert_font(30, bold=True), fill=(18, 26, 22, 255), max_width=1020)
    if org:
        y_cur = draw_centered_text(y_cur - 2, org, _get_cert_font(27, bold=True), fill=(25, 35, 28, 255), max_width=1020)
    y_cur += 10
    f_body = _get_cert_font(24, bold=True, italic=True)
    y_cur = draw_centered_text(y_cur, "in recognition of " + p_word + " valuable contribution as a", f_body, fill=(25, 30, 28, 255), max_width=1040)
    y_cur = draw_centered_text(y_cur - 4, "Guest Speaker at Government College of Technology, Samanabad, Faisalabad.", f_body, fill=(25, 30, 28, 255), max_width=1040)
    if topic:
        if lec_date:
            topic_line = 'Topic: "' + topic + '"  |  Date: ' + lec_date
        else:
            topic_line = 'Topic: "' + topic + '"'
        y_cur = draw_centered_text(y_cur + 2, topic_line, _get_cert_font(22, bold=True), fill=EMERALD_DARK, max_width=1040)
    y_cur = draw_centered_text(y_cur + 2, "We deeply appreciate " + p_word + " time, expertise and inspiring insights", f_body, fill=(25, 30, 28, 255), max_width=1040)
    y_cur = draw_centered_text(y_cur - 4, "for the benefit of our students and faculty.", f_body, fill=(25, 30, 28, 255), max_width=1040)
    _draw_cert_star(draw, W // 2 + 120, 775, 13, 5, fill=GOLD_DARK)
    seal_x, seal_y, seal_r = W // 2, 875, 84
    _draw_cert_laurel_wreath(draw, seal_x, seal_y + 4, seal_r - 16, color_main=GOLD_BRIGHT, color_hi=(255, 232, 140))
    draw.ellipse([seal_x - seal_r, seal_y - seal_r, seal_x + seal_r, seal_y + seal_r], fill=GOLD_DARK, outline=GOLD_BRIGHT, width=4)
    draw.ellipse([seal_x - seal_r + 10, seal_y - seal_r + 10, seal_x + seal_r - 10, seal_y + seal_r - 10], fill=EMERALD_DARK, outline=GOLD_BRIGHT, width=3)
    _draw_cert_laurel_wreath(draw, seal_x, seal_y + 2, seal_r - 24, color_main=GOLD_BRIGHT, color_hi=(255, 235, 150))
    f_seal_big = _get_cert_font(34, bold=True)
    sb1 = draw.textbbox((0, 0), "GCT", font=f_seal_big)
    draw.text((seal_x - (sb1[2] - sb1[0]) / 2, seal_y - 34), "GCT", fill=GOLD_BRIGHT, font=f_seal_big)
    f_seal_sm = _get_cert_font(18, bold=True)
    sb2 = draw.textbbox((0, 0), "Faisalabad", font=f_seal_sm)
    draw.text((seal_x - (sb2[2] - sb2[0]) / 2, seal_y + 6), "Faisalabad", fill=(255, 255, 255, 255), font=f_seal_sm)
    _draw_cert_star(draw, seal_x, seal_y + 38, 8, 3, fill=GOLD_BRIGHT)
    f_sig_title = _get_cert_font(25, bold=True)
    f_sig_sub = _get_cert_font(20, bold=False)
    lx = 370
    draw.line([(lx - 165, 855), (lx + 165, 855)], fill=(120, 120, 110, 255), width=2)
    ib1 = draw.textbbox((0, 0), "Institute Placement Officer", font=f_sig_title)
    draw.text((lx - (ib1[2] - ib1[0]) / 2, 864), "Institute Placement Officer", fill=(18, 22, 20, 255), font=f_sig_title)
    ib2 = draw.textbbox((0, 0), "GCT Samanabad, Faisalabad", font=f_sig_sub)
    draw.text((lx - (ib2[2] - ib2[0]) / 2, 896), "GCT Samanabad, Faisalabad", fill=(35, 40, 38, 255), font=f_sig_sub)
    rx = W - 370
    draw.line([(rx - 165, 855), (rx + 165, 855)], fill=(120, 120, 110, 255), width=2)
    pb1 = draw.textbbox((0, 0), "Principal", font=f_sig_title)
    draw.text((rx - (pb1[2] - pb1[0]) / 2, 864), "Principal", fill=(18, 22, 20, 255), font=f_sig_title)
    draw.text((rx - (ib2[2] - ib2[0]) / 2, 896), "GCT Samanabad, Faisalabad", fill=(35, 40, 38, 255), font=f_sig_sub)
    try:
        effective_ipo_bytes = ipo_sig_bytes if ipo_sig_bytes else base64.b64decode(EMBEDDED_IPO_SIG_B64)
        ipo_rgba = _extract_signature_rgba(effective_ipo_bytes)
        _paste_cert_signature(img, ipo_rgba, lx, 853, max_w=240, max_h=85)
    except Exception:
        pass
    try:
        effective_pri_bytes = pri_sig_bytes if pri_sig_bytes else base64.b64decode(EMBEDDED_PRINCIPAL_SIG_B64)
        pri_rgba = _extract_signature_rgba(effective_pri_bytes)
        _paste_cert_signature(img, pri_rgba, rx, 853, max_w=230, max_h=95)
    except Exception:
        pass
    final_rgb = img.convert("RGB")
    out_buf = io.BytesIO()
    if output_format.upper() == "PDF":
        final_rgb.save(out_buf, format="PDF", resolution=150.0)
    else:
        final_rgb.save(out_buf, format="PNG", optimize=True)
    out_buf.seek(0)
    return out_buf.getvalue()

def generate_tour_approval_application(data):
    doc = Document()
    
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_date = title.add_run(f"Date: {data.get('tour_date', '')}\n")
    r_date.font.name = 'Calibri'
    r_date.font.size = Pt(11)

    to_p = doc.add_paragraph()
    to_run = to_p.add_run("To,\n    The Worthy Principal,\n    Government College of Technology,\n    Samanabad, Faisalabad.\n")
    to_run.bold = True
    to_run.font.name = 'Calibri'
    to_run.font.size = Pt(12)

    techs_display = str(data.get('technology', '')).upper()
    subj = doc.add_paragraph()
    subj_run = subj.add_run(f"SUBJECT: PERMISSION & ADMINISTRATIVE APPROVAL FOR INDUSTRIAL STUDY TOUR (DAE {techs_display})")
    subj_run.bold = True
    subj_run.font.name = 'Calibri'
    subj_run.font.size = Pt(11)

    body = doc.add_paragraph()
    body.paragraph_format.line_spacing = 1.15
    body_text = (
        f"Respected Sir,\n\n"
        f"It is submitted that an industrial study tour has been scheduled for the students of "
        f"DAE {data.get('technology', '')} (Session: {data.get('session', '')}) to visit {data.get('industry_name', '')}.\n\n"
        f"The objective of this visit is to bridge academic training with industrial operations and provide "
        f"practical orientation to trainees regarding live plant machinery and industrial standards.\n\n"
        f"Details of the Visit & Accompanying Staff are as under:\n"
        f"  • Destination / Industry: {data.get('industry_name', '')}\n"
        f"  • Date & Departure Time: {data.get('tour_date', '')} at {data.get('dep_time', '')}\n"
        f"  • Mode of Travel / Vehicle: {data.get('travel_mode', 'Bus')}\n"
        f"  • Total Students: {data.get('students_count', '')}\n"
        f"  • Institute Bus Driver: {data.get('driver_name', '')}\n"
        f"  • Institute Bus Conductor: {data.get('conductor_name', '')}\n"
        f"  • Accompanying AIPO / Trade Teacher: {data.get('aipo_name', '')}\n"
        f"  • Accompanying Shop Assistant / Attendant: {data.get('assistant_name', '')}\n\n"
        f"It is requested that permission and administrative approval for official vehicle movement, "
        f"fuel allocation, and tour facilitation may kindly be accorded.\n"
    )
    b_run = body.add_run(body_text)
    b_run.font.name = 'Calibri'
    b_run.font.size = Pt(11)

    sig = doc.add_paragraph()
    sig.paragraph_format.space_before = Pt(30)
    sig_run = sig.add_run(
        "Institute Placement Officer (IPO)\n"
        "GCT Samanabad, Faisalabad\n\n\n"
        "Approved / Sanctioned:\n\n\n"
        "_________________________\n"
        "PRINCIPAL\n"
        "GCT Samanabad, Faisalabad"
    )
    sig_run.bold = True
    sig_run.font.name = 'Calibri'

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

def generate_annexure_g(data, student_df=None):
    doc = Document()
    
    # 0.75 inch clean page margins
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    # Document Heading
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(2)
    p_title.paragraph_format.space_before = Pt(0)
    
    r_sub = p_title.add_run("INDUSTRIAL VISIT FORM\n")
    r_sub.bold = True
    r_sub.font.name = 'Calibri'
    r_sub.font.size = Pt(14)
    
    r_annex = p_title.add_run("Annexure ‘G’")
    r_annex.bold = True
    r_annex.font.name = 'Calibri'
    r_annex.font.size = Pt(12)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # 4-Column Header Table
    t_meta = doc.add_table(rows=2, cols=4)
    t_meta.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers = ["Name of Institution", "Name of IPO/Career Counsellor", "Name of Technology", "Date"]
    for i, h in enumerate(headers):
        c = t_meta.cell(0, i)
        c.text = h
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].bold = True
        p.runs[0].font.name = 'Calibri'
        p.runs[0].font.size = Pt(9.5)
    
    inst_val = "GCT FSD"
    ipo_val = str(data.get('ipo_name', 'IKHLAS AHMAD'))
    tech_val = str(data.get('technology', 'CIVIL')).upper()
    date_val = str(data.get('tour_date', ''))
    
    vals = [inst_val, ipo_val, tech_val, date_val]
    for i, v in enumerate(vals):
        c = t_meta.cell(1, i)
        c.text = str(v)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].font.name = 'Calibri'
        p.runs[0].font.size = Pt(9.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    st_total_cnt = len(student_df) if student_df is not None and not student_df.empty else data.get('students_count', 0)

    fields_list = [
        ("1. Class/Trade/Section", f"3RD YEAR {tech_val} ({data.get('session', '')})"),
        ("2. Date & Time of Departure", f"{data.get('tour_date', '')} at {data.get('dep_time', '09:00 AM')}"),
        ("3. Details of Industry Focal Person:", ""),
        ("    Name", str(data.get('focal_person', ''))),
        ("    Contact no.", str(data.get('focal_contact', ''))),
        ("4. Mode of Travel", f"{data.get('travel_mode', 'Bus')} (Bus / Van / Car / Train / Other Mode)"),
        ("5. If Other Mode, Then Specify", str(data.get('other_mode', '') or 'N/A')),
        ("6. Copy of Approval letter from Industry", "Yes" if data.get('approval_copy') == "Yes" else "No"),
        ("7. List of Students Male/Female", f"Total: {st_total_cnt} Students (Enclosed on below format)")
    ]

    p_fields = doc.add_paragraph()
    p_fields.paragraph_format.line_spacing = 1.35
    p_fields.paragraph_format.space_after = Pt(12)

    for lbl, val in fields_list:
        r_lbl = p_fields.add_run(f"{lbl:<42}")
        r_lbl.bold = True
        r_lbl.font.name = 'Calibri'
        r_lbl.font.size = Pt(10.5)
        
        if val:
            r_val = p_fields.add_run(f"{val}\n")
            r_val.font.name = 'Calibri'
            r_val.font.size = Pt(10.5)
        else:
            p_fields.add_run("\n")

    p_tbl_lbl = doc.add_paragraph()
    p_tbl_lbl.paragraph_format.space_before = Pt(4)
    p_tbl_lbl.paragraph_format.space_after = Pt(4)
    r_tbl_lbl = p_tbl_lbl.add_run("List of Students Details:")
    r_tbl_lbl.bold = True
    r_tbl_lbl.font.name = 'Calibri'
    r_tbl_lbl.font.size = Pt(10.5)

    # 9-Column Students Details Table
    st_table = doc.add_table(rows=1, cols=9)
    st_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    st_headers = ["Session", "Trade", "Roll#", "Name", "Gender", "CNIC", "Contact", "Email", "Type"]
    for i, h in enumerate(st_headers):
        c = st_table.cell(0, i)
        c.text = h
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].bold = True
        p.runs[0].font.name = 'Calibri'
        p.runs[0].font.size = Pt(8.5)

    if student_df is not None and not student_df.empty:
        for _, row in student_df.iterrows():
            r_cells = st_table.add_row().cells
            r_cells[0].text = str(data.get('session', ''))
            r_cells[1].text = str(row.get('Technology', data.get('technology', '')))
            r_cells[2].text = str(row.get('Roll No', ''))
            r_cells[3].text = str(row.get('Name', ''))
            r_cells[4].text = "Male"
            r_cells[5].text = str(row.get('CNIC', ''))
            r_cells[6].text = str(row.get('Mobile Number', ''))
            r_cells[7].text = str(row.get('Email', ''))
            r_cells[8].text = "DAE"
            for cell in r_cells:
                cell.paragraphs[0].runs[0].font.name = 'Calibri'
                cell.paragraphs[0].runs[0].font.size = Pt(8)
    else:
        r_cells = st_table.add_row().cells
        for cell in r_cells:
            cell.text = "-"
            cell.paragraphs[0].runs[0].font.name = 'Calibri'
            cell.paragraphs[0].runs[0].font.size = Pt(8)

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

def generate_annexure_h(data):
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.7)
        section.bottom_margin = Inches(0.7)
        section.left_margin = Inches(0.7)
        section.right_margin = Inches(0.7)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(2)
    r1 = p_title.add_run("Annexure ‘H’\n")
    r1.bold = True
    r1.font.name = 'Calibri'
    r1.font.size = Pt(13)
    r2 = p_title.add_run("Industrial Visit Feedback Performa\n")
    r2.bold = True
    r2.font.name = 'Calibri'
    r2.font.size = Pt(12)
    r3 = p_title.add_run("(To be filled by Industrial Representative)")
    r3.italic = True
    r3.font.name = 'Calibri'
    r3.font.size = Pt(10)

    rep_person = data.get('rep_name') or data.get('focal_person') or ""
    rep_desig = data.get('rep_designation') or "Focal Person / HR Representative"

    p_info = doc.add_paragraph()
    p_info.paragraph_format.line_spacing = 1.25
    p_info.paragraph_format.space_before = Pt(8)
    p_info.paragraph_format.space_after = Pt(8)

    def add_field(label, val):
        r_lbl = p_info.add_run(f"{label}: ")
        r_lbl.bold = True
        r_lbl.font.name = 'Calibri'
        r_lbl.font.size = Pt(10.5)
        r_val = p_info.add_run(f"{val}\n")
        r_val.underline = True
        r_val.font.name = 'Calibri'
        r_val.font.size = Pt(10.5)

    add_field("Name of Organization/Industry", str(data.get('industry_name', '')))
    add_field("Name of Industrial Representative", str(rep_person))
    add_field("Designation", str(rep_desig))
    add_field("Date of Visit", str(data.get('tour_date', '')))
    add_field("Number of Students Attended", str(data.get('students_count', '')))
    add_field("Institution Name (TEVTA Institute)", "Government College of Technology, Samanabad, Faisalabad")

    p_lbl = doc.add_paragraph()
    p_lbl.paragraph_format.space_after = Pt(4)
    r_tbl = p_lbl.add_run("Feedback on Student Visit")
    r_tbl.bold = True
    r_tbl.font.name = 'Calibri'
    r_tbl.font.size = Pt(11)

    table = doc.add_table(rows=1, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_names = ["Criteria", "Excellent", "Good", "Needs Improvement"]
    for i, c in enumerate(col_names):
        cell = table.cell(0, i)
        cell.text = c
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].bold = True
        p.runs[0].font.name = 'Calibri'
        p.runs[0].font.size = Pt(9.5)

    criteria_items = [
        ("Punctuality and Time Management", data.get('h_punctuality', 'Excellent')),
        ("Discipline and Behavior", data.get('h_discipline', 'Excellent')),
        ("Interest and Participation Level", data.get('h_interest', 'Excellent')),
        ("Technical Knowledge and Understanding", data.get('h_tech_know', 'Good')),
        ("Communication and Interpersonal Skills", data.get('h_comm', 'Good')),
        ("Safety Awareness and Compliance", data.get('h_safety', 'Excellent'))
    ]

    for crit, choice_val in criteria_items:
        row = table.add_row().cells
        row[0].text = crit
        row[0].paragraphs[0].runs[0].font.name = 'Calibri'
        row[0].paragraphs[0].runs[0].font.size = Pt(9.5)
        
        row[1].text = "☑" if choice_val == "Excellent" else "☐"
        row[2].text = "☑" if choice_val == "Good" else "☐"
        row[3].text = "☑" if choice_val == "Needs Improvement" else "☐"
        for idx in [1, 2, 3]:
            p = row[idx].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.runs[0].font.name = 'Calibri'
            p.runs[0].font.size = Pt(11)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    p_comm = doc.add_paragraph()
    p_comm.paragraph_format.line_spacing = 1.15
    p_comm.paragraph_format.space_after = Pt(8)
    r_c_lbl = p_comm.add_run("General Comments / Suggestions:\n")
    r_c_lbl.bold = True
    r_c_lbl.font.name = 'Calibri'
    r_c_lbl.font.size = Pt(10)
    comments_text = data.get('h_comments', 'The students demonstrated strong disciplinary conduct and actively engaged with industrial operations.')
    r_c_val = p_comm.add_run(comments_text)
    r_c_val.underline = True
    r_c_val.font.name = 'Calibri'
    r_c_val.font.size = Pt(10)

    rec = data.get('h_recommendation', 'Yes')
    p_rec = doc.add_paragraph()
    p_rec.paragraph_format.space_after = Pt(18)
    r_r_lbl = p_rec.add_run("Would you recommend future visits from TEVTA students?\n")
    r_r_lbl.bold = True
    r_r_lbl.font.name = 'Calibri'
    r_r_lbl.font.size = Pt(10)
    c_yes = "☑ Yes" if rec == "Yes" else "☐ Yes"
    c_no = "☑ No" if rec == "No" else "☐ No"
    c_imp = "☑ With Improvements" if rec == "With Improvements" else "☐ With Improvements"
    r_r_val = p_rec.add_run(f"{c_yes}      {c_no}      {c_imp}")
    r_r_val.font.name = 'Calibri'
    r_r_val.font.size = Pt(10)

    p_sig = doc.add_paragraph()
    p_sig.paragraph_format.line_spacing = 1.2
    sig_text = (
        "Signature & Stamp: _______________________\n"
        f"(Industrial Representative: {rep_person})\n"
        f"Date: {data.get('tour_date', '')}"
    )
    r_sig = p_sig.add_run(sig_text)
    r_sig.bold = True
    r_sig.font.name = 'Calibri'
    r_sig.font.size = Pt(10)

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

def generate_annexure_i(data):
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.7)
        section.bottom_margin = Inches(0.7)
        section.left_margin = Inches(0.7)
        section.right_margin = Inches(0.7)

    # Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(2)
    r1 = p_title.add_run("Annexure ‘I’\n")
    r1.bold = True
    r1.font.name = 'Calibri'
    r1.font.size = Pt(13)
    r2 = p_title.add_run("Feedback about Industry Visit\n")
    r2.bold = True
    r2.font.name = 'Calibri'
    r2.font.size = Pt(12)
    r3 = p_title.add_run("(To be filled by IPO/Career Counsellor)")
    r3.italic = True
    r3.font.name = 'Calibri'
    r3.font.size = Pt(10)

    # Meta Lines
    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.line_spacing = 1.25
    p_meta.paragraph_format.space_before = Pt(8)
    p_meta.paragraph_format.space_after = Pt(8)

    def add_meta(label, val):
        r_lbl = p_meta.add_run(f"{label}: ")
        r_lbl.bold = True
        r_lbl.font.name = 'Calibri'
        r_lbl.font.size = Pt(10.5)
        r_val = p_meta.add_run(f"{val}\n")
        r_val.underline = True
        r_val.font.name = 'Calibri'
        r_val.font.size = Pt(10.5)

    add_meta("Name of IPO/Career Counsellor", str(data.get('ipo_name', 'Engr. Muhammad Wasif Ali (IPO)')))
    add_meta("Institute", "Government College of Technology, Samanabad, Faisalabad")
    add_meta("District", "Faisalabad")
    add_meta("Industry Visited", str(data.get('industry_name', '')))

    p_q = doc.add_paragraph()
    p_q.paragraph_format.line_spacing = 1.25
    p_q.paragraph_format.space_after = Pt(6)

    # 1. Relevant to field of study
    q1_lbl = p_q.add_run("1. The Industrial visit was relevant to field of study?\n")
    q1_lbl.bold = True
    q1_lbl.font.name = 'Calibri'
    q1_lbl.font.size = Pt(10.5)
    q1_ans = data.get('q_relevant', 'YES')
    q1_tick = "☑ YES    ☐ NO" if q1_ans == "YES" else "☐ YES    ☑ NO"
    p_q.add_run(f"{q1_tick}\n\n")

    # 2. Pre visit orientation
    q2_lbl = p_q.add_run("2. Pre visit orientation conducted for the trainees?\n")
    q2_lbl.bold = True
    q2_lbl.font.name = 'Calibri'
    q2_lbl.font.size = Pt(10.5)
    q2_ans = data.get('q1_orientation', 'YES')
    q2_tick = "☑ YES    ☐ NO" if q2_ans == "YES" else "☐ YES    ☑ NO"
    p_q.add_run(f"{q2_tick}\n\n")

    # 3. Objectives achieved
    q3_lbl = p_q.add_run("3. Objectives of industry visit were achieved?\n")
    q3_lbl.bold = True
    q3_lbl.font.name = 'Calibri'
    q3_lbl.font.size = Pt(10.5)
    q3_ans = data.get('q_objectives', 'YES')
    q3_tick = "☑ YES    ☐ NO" if q3_ans == "YES" else "☐ YES    ☑ NO"
    p_q.add_run(f"{q3_tick}\n\n")

    # 4. Machinery explained
    q4_lbl = p_q.add_run("4. The working on the machinery/equipment was explained by the industry representative or the IPO or the trade teacher of class?\n")
    q4_lbl.bold = True
    q4_lbl.font.name = 'Calibri'
    q4_lbl.font.size = Pt(10.5)
    q4_ans = data.get('q2_machinery', 'YES')
    q4_tick = "☑ YES    ☐ NO" if q4_ans == "YES" else "☐ YES    ☑ NO"
    p_q.add_run(f"{q4_tick}\n\n")

    r_hdr = p_q.add_run("Rating (5 is Best):\n")
    r_hdr.bold = True
    r_hdr.font.name = 'Calibri'
    r_hdr.font.size = Pt(10.5)

    ratings_dict = [
        ("5. Students Discipline during industry", data.get('r_discipline', 5)),
        ("6. Students gain technical exposure?", data.get('r_tech', 5)),
        ("7. Students gain Professional Exposure", data.get('r_prof', 5)),
        ("8. Overall satisfaction level after industrial visit", data.get('r_satisfaction', 5))
    ]

    for label, val in ratings_dict:
        scale_parts = []
        for x in range(1, 6):
            if str(x) == str(val):
                scale_parts.append(f"[{x}]")
            else:
                scale_parts.append(f"{x}")
        scale_str = " – ".join(scale_parts)
        
        r_lbl = p_q.add_run(f"{label}\n")
        r_lbl.font.name = 'Calibri'
        r_lbl.font.size = Pt(10)

        r_scale = p_q.add_run(f"{scale_str}\n\n")
        r_scale.bold = True
        r_scale.font.name = 'Calibri'
        r_scale.font.size = Pt(10)

    p_sig = doc.add_paragraph()
    p_sig.paragraph_format.space_before = Pt(25)
    p_sig.paragraph_format.line_spacing = 1.2
    sig_text = (
        "_______________________________\n"
        f"Signature with Date: {data.get('tour_date', '')}\n"
        "Institute Placement Officer"
    )
    r_sig = p_sig.add_run(sig_text)
    r_sig.bold = True
    r_sig.font.name = 'Calibri'
    r_sig.font.size = Pt(10.5)

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# ----------------- Helper Functions: Recruitment Drive Docs -----------------

def generate_drive_principal_application(data):
    doc = Document()
    
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_date = title.add_run(f"Date: {data['drive_date']}\n")
    r_date.font.name = 'Calibri'
    r_date.font.size = Pt(11)

    to_p = doc.add_paragraph()
    to_run = to_p.add_run("To,\n    The Worthy Principal,\n    Government College of Technology,\n    Samanabad, Faisalabad.\n")
    to_run.bold = True
    to_run.font.name = 'Calibri'
    to_run.font.size = Pt(12)

    techs_str = ", ".join(data.get('target_techs', []))
    subj = doc.add_paragraph()
    subj_run = subj.add_run(f"SUBJECT: PERMISSION & ADMINISTRATIVE APPROVAL FOR ON-CAMPUS RECRUITMENT DRIVE ({data['company_name'].upper()})")
    subj_run.bold = True
    subj_run.font.name = 'Calibri'
    subj_run.font.size = Pt(11)

    body = doc.add_paragraph()
    body.paragraph_format.line_spacing = 1.15
    body_text = (
        f"Respected Sir,\n\n"
        f"It is submitted with utmost respect that the Placement Cell (IPO Office) has coordinated an "
        f"On-Campus Recruitment Drive in collaboration with renowned industrial organization '{data['company_name']}'.\n\n"
        f"The recruitment drive is scheduled to hire fresh DAE diploma holders / pass-outs. Details of the drive are as follows:\n\n"
        f"  • Hiring Company / Group: {data['company_name']} ({data['company_address']})\n"
        f"  • Scheduled Date & Time: {data['drive_date']} at {data['drive_time']}\n"
        f"  • Venue / Hall: {data['venue']}\n"
        f"  • Written Test / Screening: {data['written_test']}\n"
        f"  • Targeted Technologies / Trades: {techs_str} (Session: {data['session']})\n"
        f"  • Expected Salary Range: PKR {data['salary_range']} / month\n\n"
        f"To conduct this mega on-campus recruitment event in an organized and professional manner, "
        f"the following administrative sanctions and protocol arrangements are requested:\n"
        f"  1. Refreshment Sanction: For approximately {data['refreshment_heads']} persons (Company HR delegation & Placement committee).\n"
        f"  2. Stage / Reception Display: {data['backdrop_details']}.\n"
        f"  3. Protocol / Welcome: {data['bouquet_details']}.\n"
        f"  4. Venue & Seating Arrangements: Main Hall / Conference Room with multimedia setup.\n\n"
        f"Submitted for your kind administrative approval and sanction, please."
    )
    b_run = body.add_run(body_text)
    b_run.font.name = 'Calibri'
    b_run.font.size = Pt(11)

    sig = doc.add_paragraph()
    sig.paragraph_format.space_before = Pt(30)
    sig_run = sig.add_run(
        "Institute Placement Officer (IPO)\n"
        "GCT Samanabad, Faisalabad\n\n\n"
        "Sanctioned / Approved:\n\n\n"
        "_________________________\n"
        "PRINCIPAL\n"
        "GCT Samanabad, Faisalabad"
    )
    sig_run.bold = True
    sig_run.font.name = 'Calibri'

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# ==============================================================================
# OFFICIAL TEVTA ANNEXURES GENERATOR (STRICT 1:1 ZERO STYLING INJECTION)
# ==============================================================================


# ==============================================================================
# OFFICIAL TEVTA ANNEXURES GENERATOR (STRICT 1:1 ZERO STYLING INJECTION)
# With Cloud & Offline Embedded In-Memory Template Fallback
# ==============================================================================

EMBEDDED_ANNEXURE_J_B64 = "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"
EMBEDDED_ANNEXURE_K_B64 = "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"
EMBEDDED_ANNEXURE_L_B64 = "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"

EMBEDDED_ANNEXURE_G_B64 = "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"
EMBEDDED_ANNEXURE_H_B64 = "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"
EMBEDDED_ANNEXURE_I_B64 = "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"

def generate_annexure_g(data, student_df=None):
    template_raw = get_official_template_bytes('G')
    doc = Document(io.BytesIO(template_raw))

    # 1. Update Table 0 (Meta Header)
    if len(doc.tables) > 0:
        t_meta = doc.tables[0]
        if len(t_meta.rows) >= 2:
            inst_val = "GCT FSD"
            ipo_val = str(data.get('ipo_name', 'Engr. Muhammad Wasif Ali (IPO)')).upper()
            tech_val = str(data.get('technology', '')).upper()
            date_val = str(data.get('tour_date', ''))
            vals = [inst_val, ipo_val, tech_val, date_val]
            for i, v in enumerate(vals):
                cell = t_meta.cell(1, i)
                cell.text = str(v)
                for p in cell.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for r in p.runs:
                        r.font.name = 'Calibri'
                        r.font.size = Pt(9.5)

    # 2. Update Paragraph fields
    tech_val_full = str(data.get('technology', '')).upper()
    session_val = str(data.get('session', ''))
    tour_date_val = str(data.get('tour_date', ''))
    dep_time_val = str(data.get('dep_time', '09:00 AM'))
    focal_name = str(data.get('focal_person', ''))
    focal_contact = str(data.get('focal_contact', ''))
    travel_mode = str(data.get('travel_mode', 'Bus'))
    other_mode = str(data.get('other_mode', '') or 'N/A')
    approval_copy = str(data.get('approval_copy', 'Yes'))
    st_total_cnt = len(student_df) if student_df is not None and not student_df.empty else data.get('students_count', 0)

    for p in doc.paragraphs:
        txt = p.text
        if "1. Class/Trade/Section" in txt:
            p.text = f"1. Class/Trade/Section\t\t\t\t3RD YEAR {tech_val_full} ({session_val})"
        elif "2. Date & Time of Departure" in txt:
            p.text = f"2. Date & Time of Departure\t\t\t\t{tour_date_val} at {dep_time_val}"
        elif "Name" in txt and "Contact no." not in txt and len(txt.strip()) < 40 and "Institution" not in txt and "Technology" not in txt and "IPO" not in txt:
            p.text = f"Name                                 \t\t\t {focal_name}"
        elif "Contact no." in txt:
            p.text = f"Contact no.                         \t\t\t {focal_contact}"
        elif "4. Mode of Travel" in txt:
            p.text = f"4. Mode of Travel\t\t\t\t\t{travel_mode} (Bus / Van / Car / Train / Other Mode)"
        elif "5. If Other Mode, Then Specify" in txt:
            p.text = f"5. If Other Mode, Then Specify            \t\t{other_mode}"
        elif "6. Copy of Approval letter from Industry" in txt:
            p.text = f"6. Copy of Approval letter from Industry\t\t{approval_copy}"
        elif "7. List of Students Male/Female" in txt:
            p.text = f"7. List of Students Male/Female\t\t\tTotal: {st_total_cnt} Students (Enclosed on below format)"
        
        for r in p.runs:
            r.font.name = 'Calibri'
            r.font.size = Pt(10)

    # 3. Update Table 1 (Students Table)
    if len(doc.tables) > 1:
        st_table = doc.tables[1]
        while len(st_table.rows) > 1:
            r = st_table.rows[1]
            st_table._tbl.remove(r._tr)

        if student_df is not None and not student_df.empty:
            for _, row in student_df.iterrows():
                r_cells = st_table.add_row().cells
                r_cells[0].text = str(session_val)
                r_cells[1].text = str(row.get('Technology', data.get('technology', '')))
                r_cells[2].text = str(row.get('Roll No', ''))
                r_cells[3].text = str(row.get('Name', ''))
                r_cells[4].text = "Male"
                r_cells[5].text = str(row.get('CNIC', ''))
                r_cells[6].text = str(row.get('Mobile Number', ''))
                r_cells[7].text = str(row.get('Email', ''))
                r_cells[8].text = "DAE"
                for cell in r_cells:
                    for cp in cell.paragraphs:
                        for r in cp.runs:
                            r.font.name = 'Calibri'
                            r.font.size = Pt(8.5)
                        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io.getvalue()

def generate_annexure_h(data):
    template_raw = get_official_template_bytes('H')
    doc = Document(io.BytesIO(template_raw))

    rep_person = data.get('rep_name') or data.get('focal_person') or ""
    rep_desig = data.get('rep_designation') or "Focal Person / HR Representative"

    for p in doc.paragraphs:
        txt = p.text
        if "Name of Organization/Industry:" in txt:
            p.text = f"Name of Organization/Industry: \t\t\t\t{data.get('industry_name', '')}"
        elif "Name of Industrial Representative:" in txt:
            p.text = f"Name of Industrial Representative: \t\t\t{rep_person}"
        elif "Designation:" in txt:
            p.text = f"Designation: \t\t\t\t\t\t{rep_desig}"
        elif "Date of Visit:" in txt:
            p.text = f"Date of Visit: \t\t\t\t\t\t{data.get('tour_date', '')}"
        elif "Number of Students Attended:" in txt:
            p.text = f"Number of Students Attended: \t\t\t\t{data.get('students_count', '')}"
        elif "Institution Name (TEVTA Institute):" in txt:
            p.text = f"Institution Name (TEVTA Institute): \t\t\tGovernment College of Technology, Samanabad, Faisalabad"
        elif "Would you recommend future visits from TEVTA students?" in txt:
            rec = data.get('h_recommendation', 'Yes')
            c_yes = "☑ Yes" if rec == "Yes" else "☐ Yes"
            c_no = "☑ No" if rec == "No" else "☐ No"
            c_imp = "☑ With Improvements" if rec == "With Improvements" else "☐ With Improvements"
            p.text = f"Would you recommend future visits from TEVTA students?\n{c_yes}      {c_no}      {c_imp}"
        elif "Signature & Stamp:" in txt:
            p.text = f"Signature & Stamp: _______________________\n(Industrial Representative: {rep_person})\nDate: {data.get('tour_date', '')}"
        elif txt.strip() == "Good professional behavior.":
            p.text = str(data.get('h_comments', ''))
            
        for r in p.runs:
            r.font.name = 'Calibri'
            r.font.size = Pt(11)

    if len(doc.tables) > 0:
        table = doc.tables[0]
        criteria_vals = [
            data.get('h_punctuality', 'Excellent'),
            data.get('h_discipline', 'Excellent'),
            data.get('h_interest', 'Excellent'),
            data.get('h_tech_know', 'Good'),
            data.get('h_comm', 'Good'),
            data.get('h_safety', 'Excellent')
        ]
        if len(table.rows) >= 7:
            for i, choice_val in enumerate(criteria_vals):
                row = table.rows[i+1].cells
                row[1].text = "☑" if choice_val == "Excellent" else "☐"
                row[2].text = "☑" if choice_val == "Good" else "☐"
                row[3].text = "☑" if choice_val == "Needs Improvement" else "☐"
                for idx in [1, 2, 3]:
                    for p in row[idx].paragraphs:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        for r in p.runs:
                            r.font.name = 'Calibri'
                            r.font.size = Pt(11)

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io.getvalue()

def generate_annexure_i(data):
    template_raw = get_official_template_bytes('I')
    doc = Document(io.BytesIO(template_raw))

    for p in doc.paragraphs:
        txt = p.text
        if "Name of IPO/Career Counsellor:" in txt:
            p.text = f"Name of IPO/Career Counsellor: \tEngr. Muhammad Wasif Ali (IPO)\nInstitute: \t\t\t\tGovernment College of Technology, Samanabad, FSD\nDistrict: \t\t\t\tFaisalabad\nIndustry Visited: \t\t\t{data.get('industry_name', '')}"
            for r in p.runs:
                r.font.name = 'Calibri'
                r.font.size = Pt(10.5)

        elif "1. The Industrial visit was relevant to field of study?" in txt:
            q1_ans = data.get('q_relevant', 'YES')
            q1_tick = "☑ YES    ☐ NO" if q1_ans == "YES" else "☐ YES    ☑ NO"
            
            q2_ans = data.get('q1_orientation', 'YES')
            q2_tick = "☑ YES    ☐ NO" if q2_ans == "YES" else "☐ YES    ☑ NO"
            
            q3_ans = data.get('q_objectives', 'YES')
            q3_tick = "☑ YES    ☐ NO" if q3_ans == "YES" else "☐ YES    ☑ NO"
            
            q4_ans = data.get('q2_machinery', 'YES')
            q4_tick = "☑ YES    ☐ NO" if q4_ans == "YES" else "☐ YES    ☑ NO"
            
            p.text = f"1. The Industrial visit was relevant to field of study?\t\t\t\t{q1_tick}\n\n2. Pre visit orientation conducted for the trainees?\t\t\t\t{q2_tick}\n\n3. Objectives of industry visit were achieved?\t\t\t\t{q3_tick}\n\n4. The work on the machinery/equipment was explained by the industry representative or the IPO or the trade teacher of class?\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t{q4_tick}\nRating (5 is Best):"
            for r in p.runs:
                r.font.name = 'Calibri'
                r.font.size = Pt(10.5)

        elif "5. Students Discipline during industry" in txt:
            ratings = [
                int(data.get('r_discipline', 5)),
                int(data.get('r_tech', 5)),
                int(data.get('r_prof', 5)),
                int(data.get('r_satisfaction', 5))
            ]
            
            s_parts = []
            for val in ratings:
                scale = []
                for x in range(1, 6):
                    if x == val:
                        scale.append(f"[{x}]")
                    else:
                        scale.append(f"{x}")
                s_parts.append(" – ".join(scale))
            
            p.text = f"5. Students Discipline during industry\t\t\t\t\t{s_parts[0]}\n\n6. Students gain technical exposure?\t\t\t\t\t\t{s_parts[1]}\n\n7. Students gain Professional Exposure\t\t\t\t\t{s_parts[2]}\n\n8. Overall satisfaction level after industrial visit\t\t\t\t{s_parts[3]}"
            for r in p.runs:
                r.font.name = 'Calibri'
                r.font.size = Pt(10)

        elif "Signature with Date:" in txt:
            p.text = f"______________________________\nSignature with Date: {data.get('tour_date', '')}\nInstitute Placement Officer"
            for r in p.runs:
                r.font.name = 'Calibri'
                r.font.size = Pt(10.5)

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io.getvalue()

def generate_drive_principal_application(data):
    doc = Document()
    
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_date = title.add_run(f"Date: {data['drive_date']}\n")
    r_date.font.name = 'Calibri'
    r_date.font.size = Pt(11)

    to_p = doc.add_paragraph()
    to_run = to_p.add_run("To,\n    The Worthy Principal,\n    Government College of Technology,\n    Samanabad, Faisalabad.\n")
    to_run.bold = True
    to_run.font.name = 'Calibri'
    to_run.font.size = Pt(12)

    techs_str = ", ".join(data.get('target_techs', []))
    subj = doc.add_paragraph()
    subj_run = subj.add_run(f"SUBJECT: PERMISSION & ADMINISTRATIVE APPROVAL FOR ON-CAMPUS RECRUITMENT DRIVE ({data['company_name'].upper()})")
    subj_run.bold = True
    subj_run.font.name = 'Calibri'
    subj_run.font.size = Pt(11)

    body = doc.add_paragraph()
    body.paragraph_format.line_spacing = 1.15
    body_text = (
        f"Respected Sir,\n\n"
        f"It is submitted with utmost respect that the Placement Cell (IPO Office) has coordinated an "
        f"On-Campus Recruitment Drive in collaboration with renowned industrial organization '{data['company_name']}'.\n\n"
        f"The recruitment drive is scheduled to hire fresh DAE diploma holders / pass-outs. Details of the drive are as follows:\n\n"
        f"  • Hiring Company / Group: {data['company_name']} ({data['company_address']})\n"
        f"  • Scheduled Date & Time: {data['drive_date']} at {data['drive_time']}\n"
        f"  • Venue / Hall: {data['venue']}\n"
        f"  • Written Test / Screening: {data['written_test']}\n"
        f"  • Targeted Technologies / Trades: {techs_str} (Session: {data['session']})\n"
        f"  • Expected Salary Range: PKR {data['salary_range']} / month\n\n"
        f"To conduct this mega on-campus recruitment event in an organized and professional manner, "
        f"the following administrative sanctions and protocol arrangements are requested:\n"
        f"  1. Refreshment Sanction: For approximately {data['refreshment_heads']} persons (Company HR delegation & Placement committee).\n"
        f"  2. Stage / Reception Display: {data['backdrop_details']}.\n"
        f"  3. Protocol / Welcome: {data['bouquet_details']}.\n"
        f"  4. Venue & Seating Arrangements: Main Hall / Conference Room with multimedia setup.\n\n"
        f"Submitted for your kind administrative approval and sanction, please."
    )
    b_run = body.add_run(body_text)
    b_run.font.name = 'Calibri'
    b_run.font.size = Pt(11)

    sig = doc.add_paragraph()
    sig.paragraph_format.space_before = Pt(30)
    sig_run = sig.add_run(
        "Institute Placement Officer (IPO)\n"
        "GCT Samanabad, Faisalabad\n\n\n"
        "Sanctioned / Approved:\n\n\n"
        "_________________________\n"
        "PRINCIPAL\n"
        "GCT Samanabad, Faisalabad"
    )
    sig_run.bold = True
    sig_run.font.name = 'Calibri'

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# ==============================================================================
# OFFICIAL TEVTA ANNEXURES GENERATOR (STRICT 1:1 ZERO STYLING INJECTION)
# ==============================================================================


# ==============================================================================
# OFFICIAL TEVTA ANNEXURES GENERATOR (STRICT 1:1 ZERO STYLING INJECTION)
# With Cloud & Offline Embedded In-Memory Template Fallback
# ==============================================================================

EMBEDDED_ANNEXURE_J_B64 = "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"
EMBEDDED_ANNEXURE_K_B64 = "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"
EMBEDDED_ANNEXURE_L_B64 = "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"

def get_official_template_bytes(template_type):
    """
    Loads the template from disk if available, otherwise returns the embedded exact
    official blank template bytes (guarantees zero missing file errors on Streamlit Cloud).
    """
    possible_names = {
        'J': [
            "Annexure_J_template.docx",
            "Annexure J (Campus Recruitement approval).docx",
            "Annexure J (Campus Recruitement approval)_2.docx",
            "Annexure J.docx"
        ],
        'K': [
            "Annexure_K_template.xlsx",
            "Annexure K (Campus Recruitment assessment sheet).xlsx",
            "Annexure K (Campus Recruitment assessment sheet)_2.xlsx",
            "Annexure K.xlsx"
        ],
        'L': [
            "Annexure_L_template.docx",
            "Annexure L (Employer's Feed Back recruitment drive).docx",
            "Annexure L (Employer's Feed Back recruitment drive)_2.docx",
            "Annexure L.docx"
        ],
        'G': ["Annex_G.docx", "Annexure_G.docx", "Annexure G.docx"],
        'H': ["Annex_H.docx", "Annexure_H.docx", "Annexure H.docx"],
        'I': ["Annex_I.docx", "Annexure_I.docx", "Annexure I.docx"],
        'E': ["Annexure E Guest speaker profile).docx", "Annexure_E.docx", "Annexure E.docx"],
        'F': ["Annexure F (Feedback Guest lecture).docx", "Annexure_F.docx", "Annexure F.docx"]
    }
    
    search_dirs = []
    if '__file__' in locals() or '__file__' in globals():
        try:
            search_dirs.append(os.path.dirname(os.path.abspath(__file__)))
        except:
            pass
    search_dirs.append(os.getcwd())
    for d in list(search_dirs):
        search_dirs.append(os.path.join(d, "templates"))
        search_dirs.append(os.path.join(d, "Templates"))

    names = possible_names.get(template_type, [])
    for folder in search_dirs:
        for fname in names:
            p = os.path.join(folder, fname)
            if os.path.exists(p) and os.path.isfile(p):
                if os.path.getsize(p) > 2000:
                    try:
                        with open(p, "rb") as f:
                            return f.read()
                    except:
                        pass

    # Embedded In-Memory Fallback (100% Reliable on Streamlit Cloud & Local)
    if template_type == 'J':
        return zlib.decompress(base64.b64decode(EMBEDDED_ANNEXURE_J_B64))
    elif template_type == 'K':
        return zlib.decompress(base64.b64decode(EMBEDDED_ANNEXURE_K_B64))
    elif template_type == 'L':
        return zlib.decompress(base64.b64decode(EMBEDDED_ANNEXURE_L_B64))
    elif template_type == 'G':
        return zlib.decompress(base64.b64decode(EMBEDDED_ANNEXURE_G_B64))
    elif template_type == 'H':
        return zlib.decompress(base64.b64decode(EMBEDDED_ANNEXURE_H_B64))
    elif template_type == 'I':
        return zlib.decompress(base64.b64decode(EMBEDDED_ANNEXURE_I_B64))
    elif template_type == 'E':
        return zlib.decompress(base64.b64decode(EMBEDDED_ANNEXURE_E_B64))
    elif template_type == 'F':
        return zlib.decompress(base64.b64decode(EMBEDDED_ANNEXURE_F_B64))
    
    return None


def generate_annexure_j(data, student_df=None):
    template_raw = get_official_template_bytes('J')
    doc = Document(io.BytesIO(template_raw))

    # 1. Company Profile (Table 0)
    if len(doc.tables) > 0:
        c_prof = doc.tables[0].rows[0].cells[0]
        if len(c_prof.paragraphs) >= 2:
            p0 = c_prof.paragraphs[0]
            p0.text = "NAME OF THE COMPANY:  " + str(data.get('company_name', '')).upper()
            p0.paragraph_format.line_spacing = 1.0
            if p0.runs:
                p0.runs[0].font.name = 'Calibri'
                p0.runs[0].font.size = Pt(11)

            p1 = c_prof.paragraphs[1]
            p1.text = "ADDRESS OF COMPANY:  " + str(data.get('company_address', 'FAISALABAD')).upper()
            p1.paragraph_format.line_spacing = 1.0
            if p1.runs:
                p1.runs[0].font.name = 'Calibri'
                p1.runs[0].font.size = Pt(11)

    # 2. Job Site Location & Nested Job Details (Table 1)
    if len(doc.tables) > 1:
        c_job = doc.tables[1].rows[0].cells[0]
        if len(c_job.paragraphs) >= 1:
            p_loc = c_job.paragraphs[0]
            loc_val = data.get('job_location') or data.get('company_address') or 'FAISALABAD'
            p_loc.text = f"JOB SITE LOCATION:	{str(loc_val).upper()}"
            p_loc.paragraph_format.line_spacing = 1.0
            if p_loc.runs:
                p_loc.runs[0].font.name = 'Calibri'
                p_loc.runs[0].font.size = Pt(11)

        if len(c_job.tables) > 0:
            nt_job = c_job.tables[0]
            if len(nt_job.rows) >= 3:
                r2 = nt_job.rows[2]
                
                def fill_box(cell, txt):
                    p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
                    p.text = str(txt)
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p.paragraph_format.line_spacing = 1.0
                    if p.runs:
                        p.runs[0].font.name = 'Calibri'
                        p.runs[0].font.size = Pt(9)
                        p.runs[0].bold = False

                fill_box(r2.cells[0], data.get('experience', 'FRESH / RESULT AWAITED'))
                fill_box(r2.cells[1], data.get('job_title', 'ASSOCIATE / TRAINEE'))
                fill_box(r2.cells[2], str(data.get('vacancies_male', '')))
                fill_box(r2.cells[3], str(data.get('vacancies_female', '0')))

                techs = data.get('target_techs', [])
                if isinstance(techs, list):
                    tech_txt = "\n".join([f"DAE {t.upper()}" for t in techs])
                else:
                    tech_txt = str(techs).upper()
                fill_box(r2.cells[4], tech_txt)

                fill_box(r2.cells[5], data.get('marks_div', '1ST DIVISION'))
                fill_box(r2.cells[6], data.get('salary_range', ''))
                fill_box(r2.cells[7], data.get('job_desc', ''))

    # 3. Recruitment Drive Details (Table 2)
    if len(doc.tables) > 2:
        c_dr = doc.tables[2].rows[0].cells[0]
        if len(c_dr.tables) > 0:
            nt_dr = c_dr.tables[0]
            if len(nt_dr.rows) >= 3:
                r2_dr = nt_dr.rows[2]
                fill_box(r2_dr.cells[0], data.get('venue', 'GCT FSD'))
                fill_box(r2_dr.cells[1], f"{data.get('drive_date', '')} @ {data.get('drive_time', '10:00 AM')}")
                fill_box(r2_dr.cells[2], f"{data.get('dpo_name', 'MR. ASGHAR (DPO)')}\n{data.get('ipo_name', 'WASIF ALI (IPO)')}")
                fill_box(r2_dr.cells[3], str(data.get('tevta_contact', '')))
                fill_box(r2_dr.cells[4], f"{data.get('emp_focal', '')}\n{data.get('emp_desig', 'HR MANAGER')}")
                fill_box(r2_dr.cells[5], str(data.get('emp_contact', '')))

    # 4. Enclosed Candidate List (Page 2)
    if student_df is not None and not student_df.empty:
        doc.add_page_break()
        p_list = doc.add_paragraph()
        p_list.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_list.paragraph_format.line_spacing = 1.0
        r_head = p_list.add_run(
            f"LIST OF ELIGIBLE CANDIDATES ENCLOSED FOR RECRUITMENT DRIVE\n"
            f"Company: {str(data.get('company_name', '')).upper()} | Session: {data.get('session', '')}\n"
        )
        r_head.bold = True
        r_head.font.name = 'Calibri'
        r_head.font.size = Pt(11)

        s_table = doc.add_table(rows=1, cols=7)
        s_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        s_headers = ["Sr#", "Roll No", "Student Name", "Trade / Technology", "CNIC", "Mobile Number", "Email"]
        for i, h in enumerate(s_headers):
            c = s_table.cell(0, i)
            c.text = h
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.line_spacing = 1.0
            if p.runs:
                p.runs[0].bold = True
                p.runs[0].font.name = 'Calibri'
                p.runs[0].font.size = Pt(9)

        for idx, (_, row) in enumerate(student_df.iterrows(), start=1):
            rc = s_table.add_row().cells
            vals = [
                str(idx),
                str(row.get('Roll No', '')),
                str(row.get('Name', '')),
                str(row.get('Technology', '')),
                str(row.get('CNIC', '')),
                str(row.get('Mobile Number', '')),
                str(row.get('Email', ''))
            ]
            for c_i, val in enumerate(vals):
                rc[c_i].text = val
                p = rc[c_i].paragraphs[0]
                p.paragraph_format.line_spacing = 1.0
                if p.runs:
                    p.runs[0].font.name = 'Calibri'
                    p.runs[0].font.size = Pt(8.5)

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io.getvalue()


def generate_candidate_roster_excel(data, student_df):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Candidate Roster"

    ws.merge_cells("A1:G1")
    t = ws["A1"]
    t.value = f"RECRUITMENT DRIVE CANDIDATE ROSTER - {str(data.get('company_name', '')).upper()}"
    t.font = Font(name="Calibri", size=13, bold=True)
    t.alignment = Alignment(horizontal="center", vertical="center")

    ws["A2"] = f"Drive Date: {data.get('drive_date', '')}"
    ws["D2"] = f"Target Session: {data.get('session', '')}"
    ws["F2"] = f"Venue: {data.get('venue', 'GCT Samanabad')}"
    for c in ["A2", "D2", "F2"]:
        ws[c].font = Font(name="Calibri", size=11, bold=True)

    headers = ["Sr#", "Roll No", "Student Name", "Technology", "CNIC", "Mobile Number", "Email Address"]
    for col_idx, h in enumerate(headers, start=1):
        c = ws.cell(row=4, column=col_idx, value=h)
        c.font = Font(name="Calibri", size=10, bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center")

    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )

    current_row = 5
    if student_df is not None and not student_df.empty:
        for idx, (_, row) in enumerate(student_df.iterrows(), start=1):
            ws.cell(row=current_row, column=1, value=idx).alignment = Alignment(horizontal="center")
            ws.cell(row=current_row, column=2, value=str(row.get('Roll No', '')))
            ws.cell(row=current_row, column=3, value=str(row.get('Name', '')))
            ws.cell(row=current_row, column=4, value=str(row.get('Technology', '')))
            ws.cell(row=current_row, column=5, value=str(row.get('CNIC', '')))
            ws.cell(row=current_row, column=6, value=str(row.get('Mobile Number', '')))
            ws.cell(row=current_row, column=7, value=str(row.get('Email', '')))
            for c_idx in range(1, 8):
                ws.cell(row=current_row, column=c_idx).border = thin_border
                ws.cell(row=current_row, column=c_idx).font = Font(name="Calibri", size=10)
            current_row += 1

    col_widths = {1: 8, 2: 14, 3: 25, 4: 20, 5: 18, 6: 16, 7: 25}
    for c_idx, width in col_widths.items():
        ws.column_dimensions[openpyxl.utils.get_column_letter(c_idx)].width = width

    excel_io = io.BytesIO()
    wb.save(excel_io)
    excel_io.seek(0)
    return excel_io.getvalue()



def generate_annexure_k_excel(data, student_df=None):
    """
    Official TEVTA Annexure 'K' Generator:
    100% Strict compliance with official Excel assessment sheet.
    Unmerges any pre-existing ranges from row 6 onwards to prevent MergedCell read-only errors,
    applies plain white styling, thin borders, and clean candidate rows.
    """
    template_path = get_official_template_bytes('K')
    if not template_path:
        st.error("⚠️ Annexure K template file project folder mein nahi mila!")
        return io.BytesIO().getvalue()

    wb = openpyxl.load_workbook(io.BytesIO(template_path))
    ws = wb.active

    # Unmerge any pre-existing merged cell ranges from row 6 onwards to avoid MergedCell read-only errors
    for m in list(ws.merged_cells.ranges):
        if m.min_row >= 6:
            ws.unmerge_cells(str(m))

    # Exact Top Headers (White, Calibri)
    ws["A2"] = f"Interviewer: {data.get('company_name', '')}"
    ws["H2"] = data.get("venue", "GCT FSD")
    ws["L2"] = str(data.get("drive_date", ""))

    ws["A3"] = f"Time: {data.get('drive_time', '10:00 am')}"
    techs = data.get('target_techs', [])
    trades_str = ", ".join(techs) if isinstance(techs, list) else str(techs)
    ws["H3"] = trades_str
    ws["L3"] = str(data.get("salary_range", ""))

    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )

    start_row = 6
    if student_df is not None and not student_df.empty:
        for r_clear in range(6, max(25, start_row + len(student_df) + 5)):
            for c_clear in range(1, 14):
                cell = ws.cell(row=r_clear, column=c_clear)
                cell.value = None
                cell.fill = PatternFill(fill_type=None)

        for idx, (_, row) in enumerate(student_df.iterrows(), start=1):
            curr_r = start_row + idx - 1
            ws.cell(row=curr_r, column=1, value=idx).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(row=curr_r, column=2, value=str(row.get('Name', '')).title()).alignment = Alignment(horizontal="left", vertical="center")
            
            tech_val = str(row.get('Technology', ''))
            if not tech_val.upper().startswith("DAE"):
                tech_val = f"DAE {tech_val}".upper()
            ws.cell(row=curr_r, column=3, value=tech_val).alignment = Alignment(horizontal="left", vertical="center")
            ws.cell(row=curr_r, column=4, value=str(row.get('Session', data.get('session', '')))).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(row=curr_r, column=5, value="GCT Samanabad").alignment = Alignment(horizontal="left", vertical="center")

            for c_idx in range(1, 14):
                cell = ws.cell(row=curr_r, column=c_idx)
                cell.border = thin_border
                cell.fill = PatternFill(fill_type=None)
                cell.font = Font(name="Calibri", size=11, bold=(c_idx in [1, 2]))

        sig_r = start_row + len(student_df) + 2
        ws.cell(row=sig_r, column=1, value=f"Name of DPO: {data.get('dpo_name', 'Mr. Muhammad Asghar Ch (DPO)')}").font = Font(name="Calibri", size=11, bold=True)
        ws.cell(row=sig_r, column=8, value="Signature of DPO with date: ______________________").font = Font(name="Calibri", size=11, bold=True)

    excel_io = io.BytesIO()
    wb.save(excel_io)
    excel_io.seek(0)
    return excel_io.getvalue()


def generate_annexure_l(data):
    template_raw = get_official_template_bytes('L')
    doc = Document(io.BytesIO(template_raw))

    venue_str = str(data.get('venue', 'GCT Samanabad, Faisalabad'))
    date_str = str(data.get('drive_date', ''))
    app_count = str(data.get('students_appeared', ''))
    short_count = str(data.get('students_shortlisted', ''))
    sugg_str = str(data.get('l_suggestions', ''))
    comp_str = str(data.get('company_name', ''))
    focal_str = str(data.get('emp_focal', ''))
    desig_str = str(data.get('emp_desig', 'HR MANAGER'))
    contact_str = str(data.get('emp_contact', ''))
    email_str = str(data.get('emp_email', ''))

    for p in doc.paragraphs:
        if "Thank you for scheduling On-Campus Recruiting at" in p.text:
            p.text = f"Thank you for scheduling On-Campus Recruiting at (venue) {venue_str} Date: {date_str}."
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            if p.runs:
                p.runs[0].font.name = 'Calibri'
                p.runs[0].font.size = Pt(12)

        elif "No of students appeared" in p.text:
            p.text = f"No of students appeared   {app_count}"
            p.paragraph_format.line_spacing = 1.5
            if p.runs:
                p.runs[0].font.name = 'Calibri'
                p.runs[0].font.size = Pt(12)

        elif "No of students shortlisted" in p.text:
            p.text = f"No of students shortlisted {short_count}"
            p.paragraph_format.line_spacing = 1.5
            if p.runs:
                p.runs[0].font.name = 'Calibri'
                p.runs[0].font.size = Pt(12)

        elif "Suggestions (if any):" in p.text:
            if sugg_str.strip():
                p.text = f"Suggestions (if any): {sugg_str}"
            p.paragraph_format.line_spacing = 1.5
            if p.runs:
                p.runs[0].font.name = 'Calibri'
                p.runs[0].font.size = Pt(12)

        elif "Industry/Organization:" in p.text:
            p.text = f"Industry/Organization: {comp_str}\t\tName: {focal_str}\t\tDesignation: {desig_str}"
            p.paragraph_format.line_spacing = 1.5
            if p.runs:
                p.runs[0].font.name = 'Calibri'
                p.runs[0].font.size = Pt(12)

        elif "Office Phone:" in p.text:
            p.text = f"Office Phone: {contact_str}\t\tMobile: {contact_str}\t\tEmail: {email_str}"
            p.paragraph_format.line_spacing = 1.5
            if p.runs:
                p.runs[0].font.name = 'Calibri'
                p.runs[0].font.size = Pt(12)

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io.getvalue()


def render_printable_html(title, content_html):
    html_page = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <title>{title}</title>
        <style>
            @media print {{
                .no-print {{ display: none !important; }}
                body {{ margin: 0; padding: 15px; font-family: Arial, sans-serif; font-size: 11pt; }}
            }}
            body {{ font-family: Arial, sans-serif; padding: 25px; line-height: 1.4; color: #111; max-width: 980px; margin: auto; }}
            .header-banner {{ text-align: center; border-bottom: 2px solid #1F4E79; padding-bottom: 10px; margin-bottom: 20px; }}
            .inst-name {{ font-size: 15pt; font-weight: bold; color: #1F4E79; text-transform: uppercase; margin: 0; }}
            .dept-name {{ font-size: 12pt; font-weight: bold; color: #444; margin: 4px 0 0 0; }}
            .print-btn {{ background: #1F4E79; color: #fff; padding: 9px 20px; border: none; border-radius: 4px; font-size: 11pt; cursor: pointer; font-weight: bold; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 9.5pt; }}
            th, td {{ border: 1px solid #555; padding: 6px 6px; text-align: center; }}
            th {{ background-color: #e9eef4; font-weight: bold; }}
            .left-align {{ text-align: left; }}
            .right-align {{ text-align: right; }}
        </style>
    </head>
    <body>
        <div class="no-print" style="margin-bottom: 20px; text-align: right;">
            <button class="print-btn" onclick="window.print()">🖨️ Click to Print / Save as PDF</button>
        </div>
        <div class="header-banner">
            <div class="inst-name">GOVERNMENT COLLEGE OF TECHNOLOGY, SAMANABAD, FAISALABAD</div>
            <div class="dept-name">Placement & Industrial Liaison Cell (IPO Office)</div>
        </div>
        {content_html}
    </body>
    </html>
    """
    return html_page

def generate_annual_fiscal_report_excel(fiscal_year):
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "Executive Summary"
    
    ws1.merge_cells("A1:G1")
    t = ws1["A1"]
    t.value = f"TEVTA ANNUAL PLACEMENT & LINKAGES REPORT (FY: {fiscal_year})"
    t.font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    t.fill = PatternFill("solid", fgColor="1F4E79")
    t.alignment = Alignment(horizontal="center", vertical="center")

    cursor.execute("SELECT COUNT(*) FROM students")
    tot_std = cursor.fetchone()[0] or 0
    cursor.execute("SELECT COUNT(*) FROM students WHERE status = 'Placed'")
    tot_placed = cursor.fetchone()[0] or 0
    cursor.execute("SELECT COUNT(*) FROM guest_lectures")
    tot_lec = cursor.fetchone()[0] or 0
    cursor.execute("SELECT COUNT(*) FROM industry_visits")
    tot_vis = cursor.fetchone()[0] or 0
    cursor.execute("SELECT COUNT(*) FROM recruitment_drives")
    tot_dri = cursor.fetchone()[0] or 0
    
    cursor.execute("SELECT SUM(budget_utilized) FROM guest_lectures")
    b_lec = cursor.fetchone()[0] or 0.0
    cursor.execute("SELECT SUM(budget_utilized) FROM industry_visits")
    b_vis = cursor.fetchone()[0] or 0.0
    cursor.execute("SELECT SUM(budget_utilized) FROM recruitment_drives")
    b_dri = cursor.fetchone()[0] or 0.0
    tot_bud = b_lec + b_vis + b_dri

    metrics = [
        ("Total Enrolled Candidates in Directory", tot_std),
        ("Total Candidates Successfully Placed", tot_placed),
        ("Institutional Placement Efficiency Rate (%)", f"{(tot_placed/tot_std*100):.1f}%" if tot_std > 0 else "0.0%"),
        ("Total Guest Lectures Organized", tot_lec),
        ("Total Industrial Study Visits Arranged", tot_vis),
        ("Total On-Campus Recruitment Drives", tot_dri),
        ("Consolidated Financial Budget Utilized (PKR)", f"PKR {tot_bud:,.2f}")
    ]

    for idx, (label, val) in enumerate(metrics, start=3):
        ws1.cell(row=idx, column=1, value=label).font = Font(bold=True)
        ws1.cell(row=idx, column=3, value=val)

    ws1.column_dimensions["A"].width = 45
    ws1.column_dimensions["C"].width = 25

    df_l = pd.read_sql("SELECT speaker_name, topic, target_tech, target_session, lecture_date, total_attendees, budget_utilized, status FROM guest_lectures", conn)
    df_v = pd.read_sql("SELECT industry_name, technology, session, tour_date, students_count, budget_utilized, status FROM industry_visits", conn)
    df_d = pd.read_sql("SELECT company_name, job_title, target_techs, session, drive_date, students_appeared, students_shortlisted, budget_utilized, status FROM recruitment_drives", conn)
    
    excel_io = io.BytesIO()
    with pd.ExcelWriter(excel_io, engine='openpyxl') as writer:
        ws1_df = pd.DataFrame(metrics, columns=["Performance Indicator", "Value"])
        ws1_df.to_excel(writer, sheet_name="Executive_Summary", index=False)
        df_l.to_excel(writer, sheet_name="Guest_Lectures", index=False)
        df_v.to_excel(writer, sheet_name="Industrial_Visits", index=False)
        df_d.to_excel(writer, sheet_name="Recruitment_Drives", index=False)
        
    excel_io.seek(0)
    return excel_io

# ----------------- Helper Functions: Monthly Activity Plan -----------------

def generate_monthly_plan_excel(plan_data):
    wb = openpyxl.Workbook()
    ws = wb.active
    month_name = plan_data.get('month_year', 'Monthly Plan')
    ws.title = month_name[:31]

    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )

    ws.merge_cells("A1:M1")
    t = ws["A1"]
    t.value = f"Govt. College of Technology Samanabad Faisalabad\nMonthly Activity Plan (Placement) {month_name}"
    t.font = Font(name="Calibri", size=13, bold=True)
    t.alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A2:I2")
    ws["A2"] = "Career Counseling Lectures"
    ws["A2"].font = Font(name="Calibri", size=11, bold=True)
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("J2:L2")
    ws["J2"] = "Linkage Activities"
    ws["J2"].font = Font(name="Calibri", size=11, bold=True)
    ws["J2"].alignment = Alignment(horizontal="center", vertical="center")

    ws["M2"] = "Budget"
    ws["M2"].font = Font(name="Calibri", size=11, bold=True)
    ws["M2"].alignment = Alignment(horizontal="center", vertical="center")

    tech_headers = [
        "Electrical", "Mechanical", "Civil", "Food", 
        "Instrument", "Auto & Farm", "Textile", "B.Sc", "Short Courses",
        "Guest speaker lecture", "Students\nIndustrial Visit", "Recruitment\nDrive", "Budget Required"
    ]
    for col_i, h_text in enumerate(tech_headers, start=1):
        cell = ws.cell(row=3, column=col_i, value=h_text)
        cell.font = Font(name="Calibri", size=10, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    tech_vals = [
        plan_data.get('lec_electrical', 0),
        plan_data.get('lec_mechanical', 0),
        plan_data.get('lec_civil', 0),
        plan_data.get('lec_food', 0),
        plan_data.get('lec_instrument', 0),
        plan_data.get('lec_autofarm', 0),
        plan_data.get('lec_textile', 0),
        plan_data.get('lec_bsc', 0),
        plan_data.get('lec_shortcourses', 0),
        f"{plan_data.get('guest_lectures_expected', 1):02d} Expected",
        f"{plan_data.get('industrial_visits_expected', 1):02d} Expected",
        f"{plan_data.get('recruitment_drives_expected', 1):02d} Expected",
        plan_data.get('budget_breakdown', '')
    ]
    for col_i, val in enumerate(tech_vals, start=1):
        cell = ws.cell(row=4, column=col_i, value=str(val))
        cell.font = Font(name="Calibri", size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.merge_cells("A5:I5")
    tot_lectures = sum([
        int(plan_data.get('lec_electrical', 0)), int(plan_data.get('lec_mechanical', 0)),
        int(plan_data.get('lec_civil', 0)), int(plan_data.get('lec_food', 0)),
        int(plan_data.get('lec_instrument', 0)), int(plan_data.get('lec_autofarm', 0)),
        int(plan_data.get('lec_textile', 0)), int(plan_data.get('lec_bsc', 0)),
        int(plan_data.get('lec_shortcourses', 0))
    ])
    ws["A5"] = f"Total = {tot_lectures} Lectures Planned"
    ws["A5"].font = Font(name="Calibri", size=10, bold=True)
    ws["A5"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("J5:L5")
    tot_events = plan_data.get('guest_lectures_expected', 1) + plan_data.get('industrial_visits_expected', 1) + plan_data.get('recruitment_drives_expected', 1)
    ws["J5"] = f"{tot_events} Events Total"
    ws["J5"].font = Font(name="Calibri", size=10, bold=True)
    ws["J5"].alignment = Alignment(horizontal="center", vertical="center")

    ws["M5"] = f"PKR {plan_data.get('total_budget', 0):,}/-"
    ws["M5"].font = Font(name="Calibri", size=10, bold=True)
    ws["M5"].alignment = Alignment(horizontal="center", vertical="center")

    for r in range(1, 6):
        for c in range(1, 14):
            ws.cell(row=r, column=c).border = thin_border

    ws.merge_cells("A9:D13")
    sig1 = ws["A9"]
    sig1.value = "IPO GCT\n\n\n______________________\nInstitute Placement Officer\nGCT Samanabad, Faisalabad"
    sig1.font = Font(name="Calibri", size=10, bold=True)
    sig1.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.merge_cells("F9:I13")
    sig2 = ws["F9"]
    sig2.value = "Account Officer\n\n\n______________________\nAccounts Department\nGCT Samanabad, Faisalabad"
    sig2.font = Font(name="Calibri", size=10, bold=True)
    sig2.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.merge_cells("K9:M13")
    sig3 = ws["K9"]
    sig3.value = "Principal\n\n\n______________________\nPrincipal / Head of Institute\nGCT Samanabad, Faisalabad"
    sig3.font = Font(name="Calibri", size=10, bold=True)
    sig3.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for rng_str in ["A9:D13", "F9:I13", "K9:M13"]:
        min_c, min_r, max_c, max_r = openpyxl.utils.range_boundaries(rng_str)
        for r_idx in range(min_r, max_r + 1):
            for c_idx in range(min_c, max_c + 1):
                ws.cell(row=r_idx, column=c_idx).border = thin_border

    col_w = {1: 13, 2: 13, 3: 12, 4: 12, 5: 13, 6: 13, 7: 12, 8: 12, 9: 14, 10: 20, 11: 20, 12: 18, 13: 25}
    for c_i, w in col_w.items():
        ws.column_dimensions[openpyxl.utils.get_column_letter(c_i)].width = w

    ws.row_dimensions[1].height = 42
    ws.row_dimensions[4].height = 65

    excel_io = io.BytesIO()
    wb.save(excel_io)
    excel_io.seek(0)
    return excel_io

def render_monthly_plan_html(plan_data):
    month_name = plan_data.get('month_year', 'Monthly Plan')
    tot_lectures = sum([
        int(plan_data.get('lec_electrical', 0)), int(plan_data.get('lec_mechanical', 0)),
        int(plan_data.get('lec_civil', 0)), int(plan_data.get('lec_food', 0)),
        int(plan_data.get('lec_instrument', 0)), int(plan_data.get('lec_autofarm', 0)),
        int(plan_data.get('lec_textile', 0)), int(plan_data.get('lec_bsc', 0)),
        int(plan_data.get('lec_shortcourses', 0))
    ])
    tot_events = plan_data.get('guest_lectures_expected', 1) + plan_data.get('industrial_visits_expected', 1) + plan_data.get('recruitment_drives_expected', 1)
    breakdown_formatted = plan_data.get('budget_breakdown', '').replace('\n', '<br>')

    body_html = f"""
    <div style="text-align: center; margin-bottom: 15px;">
        <h3 style="margin: 0; text-transform: uppercase; color: #1F4E79;">Monthly Activity Plan (Placement) - {month_name}</h3>
        <p style="margin: 3px 0 0 0; font-weight: bold; font-size: 11pt; color: #333;">For Submission to District Placement Officer (DPO), Faisalabad</p>
    </div>

    <table>
        <thead>
            <tr>
                <th colspan="9" style="background: #1F4E79; color: white; font-size: 11pt;">Career Counseling Lectures</th>
                <th colspan="3" style="background: #2F5597; color: white; font-size: 11pt;">Linkage Activities</th>
                <th style="background: #1F4E79; color: white; font-size: 11pt;">Budget</th>
            </tr>
            <tr style="background: #f2f5f9;">
                <th>Electrical</th>
                <th>Mechanical</th>
                <th>Civil</th>
                <th>Food</th>
                <th>Instrument</th>
                <th>Auto & Farm</th>
                <th>Textile</th>
                <th>B.Sc</th>
                <th>Short Courses</th>
                <th>Guest Speaker</th>
                <th>Industrial Visit</th>
                <th>Recruitment Drive</th>
                <th>Budget Required</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td><b>{plan_data.get('lec_electrical', 0)}</b></td>
                <td><b>{plan_data.get('lec_mechanical', 0)}</b></td>
                <td><b>{plan_data.get('lec_civil', 0)}</b></td>
                <td><b>{plan_data.get('lec_food', 0)}</b></td>
                <td><b>{plan_data.get('lec_instrument', 0)}</b></td>
                <td><b>{plan_data.get('lec_autofarm', 0)}</b></td>
                <td><b>{plan_data.get('lec_textile', 0)}</b></td>
                <td><b>{plan_data.get('lec_bsc', 0)}</b></td>
                <td><b>{plan_data.get('lec_shortcourses', 0)}</b></td>
                <td><b>{plan_data.get('guest_lectures_expected', 1):02d} Expected</b></td>
                <td><b>{plan_data.get('industrial_visits_expected', 1):02d} Expected</b></td>
                <td><b>{plan_data.get('recruitment_drives_expected', 1):02d} Expected</b></td>
                <td style="font-size: 9pt; text-align: left; padding: 6px 10px;">{breakdown_formatted}</td>
            </tr>
            <tr style="background: #fafafa; font-weight: bold;">
                <td colspan="9" style="font-size: 10.5pt; text-align: center;">Total = {tot_lectures} Lectures Planned</td>
                <td colspan="3" style="font-size: 10.5pt; text-align: center;">{tot_events} Events Total</td>
                <td style="font-size: 11pt; color: #1F4E79;">PKR {plan_data.get('total_budget', 0):,}/-</td>
            </tr>
        </tbody>
    </table>

    <br><br><br><br>
    <div style="display: flex; justify-content: space-between; text-align: center; margin-top: 40px; padding: 0 10px;">
        <div style="width: 28%; border-top: 2px solid #333; padding-top: 6px;">
            <b>IPO GCT</b><br>
            <span style="font-size: 9.5pt;">Institute Placement Officer<br>GCT Samanabad, Faisalabad</span>
        </div>
        <div style="width: 28%; border-top: 2px solid #333; padding-top: 6px;">
            <b>Account Officer</b><br>
            <span style="font-size: 9.5pt;">Accounts Department<br>GCT Samanabad, Faisalabad</span>
        </div>
        <div style="width: 28%; border-top: 2px solid #333; padding-top: 6px;">
            <b>Principal</b><br>
            <span style="font-size: 9.5pt;">Head of Institution<br>GCT Samanabad, Faisalabad</span>
        </div>
    </div>
    """
    return render_printable_html(f"Monthly_Activity_Plan_{month_name.replace(' ', '_')}", body_html)

# --- Helper: Communication Charges Application Generator ---

def generate_comm_charges_application_docx(app_data):
    doc = Document()
    
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_date = title.add_run(f"Dated: {app_data['app_date']}\n")
    r_date.font.name = 'Calibri'
    r_date.font.size = Pt(11)

    to_p = doc.add_paragraph()
    to_run = to_p.add_run("The Principal\nGCT Samanabad Faisalabad\n")
    to_run.bold = True
    to_run.font.name = 'Calibri'
    to_run.font.size = Pt(12)

    subj = doc.add_paragraph()
    subj_run = subj.add_run(f"Subject: Request for Payment of Communication Charges for the Month of {app_data['target_month']}")
    subj_run.bold = True
    subj_run.font.name = 'Calibri'
    subj_run.font.size = Pt(11)

    body = doc.add_paragraph()
    body.paragraph_format.line_spacing = 1.25
    body_text = (
        f"With due respect, it is stated that I am performing my duties as Institute Placement Officer (IPO) "
        f"at GCT Samanabad Faisalabad. As per the TEVTA notification No. TEVTA/DIR-HR/CA-69/2022, dated 23-02-2022, "
        f"Communication Charges of Rs. 2,500/- per month have been sanctioned for Institute Placement Officers.\n\n"
        f"It is therefore requested that the Communication Charges for the month of {app_data['target_month']} amounting to "
        f"Rs. 2,500/- may kindly be released/paid to me as per the approved entitlement.\n\n"
        f"Kindly accord approval for the payment of the above-mentioned Communication Charges."
    )
    b_run = body.add_run(body_text)
    b_run.font.name = 'Calibri'
    b_run.font.size = Pt(11)

    sig = doc.add_paragraph()
    sig.paragraph_format.space_before = Pt(45)
    sig_run = sig.add_run(
        "Regards,\n\n"
        "Muhammad Wasif Ali\n"
        "Institute Placement Officer"
    )
    sig_run.bold = True
    sig_run.font.name = 'Calibri'

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

def render_comm_charges_html(app_data):
    html_page = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <title>Communication Charges Application - {app_data['target_month']}</title>
        <style>
            @media print {{
                .no-print {{ display: none !important; }}
                body {{ margin: 0; padding: 30px 40px; font-family: 'Calibri', Arial, sans-serif; font-size: 12pt; }}
            }}
            body {{ font-family: 'Calibri', Arial, sans-serif; padding: 40px 50px; line-height: 1.5; color: #000; max-width: 800px; margin: auto; }}
            .print-btn {{ background: #1F4E79; color: #fff; padding: 9px 20px; border: none; border-radius: 4px; font-size: 11pt; cursor: pointer; font-weight: bold; }}
            .right-align {{ text-align: right; }}
            .bold {{ font-weight: bold; }}
        </style>
    </head>
    <body>
        <div class="no-print" style="margin-bottom: 25px; text-align: right;">
            <button class="print-btn" onclick="window.print()">🖨️ Click to Print / Save as PDF</button>
        </div>

        <div class="right-align">
            <b>Dated:</b> {app_data['app_date']}
        </div>

        <p style="margin-top: 25px;">
            <b>The Principal</b><br>
            GCT Samanabad Faisalabad
        </p>

        <p style="margin-top: 25px;">
            <b>Subject: <u>Request for Payment of Communication Charges for the Month of {app_data['target_month']}</u></b>
        </p>

        <p style="text-align: justify; margin-top: 25px; line-height: 1.6;">
            With due respect, it is stated that I am performing my duties as Institute Placement Officer (IPO) at GCT Samanabad Faisalabad. As per the TEVTA notification No. TEVTA/DIR-HR/CA-69/2022, dated 23-02-2022, Communication Charges of Rs. 2,500/- per month have been sanctioned for Institute Placement Officers.
        </p>

        <p style="text-align: justify; margin-top: 18px; line-height: 1.6;">
            It is therefore requested that the Communication Charges for the month of <b>{app_data['target_month']}</b> amounting to <b>Rs. 2,500/-</b> may kindly be released/paid to me as per the approved entitlement.
        </p>

        <p style="text-align: justify; margin-top: 18px; line-height: 1.6;">
            Kindly accord approval for the payment of the above-mentioned Communication Charges.
        </p>

        <div style="margin-top: 60px;">
            <b>Regards,</b><br><br><br>
            <b>Muhammad Wasif Ali</b><br>
            Institute Placement Officer
        </div>
    </body>
    </html>
    """
    return html_page

# ----------------- Database Setup & Migrations (DIRECT SUPABASE - CACHED) -----------------

conn, _active_db_type = get_db_connection()
cursor = conn.cursor()

@st.cache_data(ttl=60, show_spinner=False)
def _cached_pg_read_sql(pg_sql, params_tuple):
    c_conn, _ = get_db_connection()
    c_conn._ensure_open()
    try:
        if params_tuple:
            return _orig_pd_read_sql(pg_sql, c_conn._conn, params=params_tuple)
        return _orig_pd_read_sql(pg_sql, c_conn._conn)
    except Exception:
        c_conn.force_close()
        c_conn._ensure_open()
        if params_tuple:
            return _orig_pd_read_sql(pg_sql, c_conn._conn, params=params_tuple)
        return _orig_pd_read_sql(pg_sql, c_conn._conn)

# Wrap pd.read_sql so every DataFrame query uses RAM cache (60s TTL, auto-invalidated on writes)
_orig_pd_read_sql = pd.read_sql
def _pg_safe_read_sql(sql, con, *args, **kwargs):
    if isinstance(con, PgCompatConnection) and not args:
        pg_sql = _translate_sql_for_pg(sql)
        params = kwargs.get("params")
        params_tuple = tuple(params) if params is not None else None
        return _cached_pg_read_sql(pg_sql, params_tuple).copy()
    return _orig_pd_read_sql(sql, con, *args, **kwargs)
pd.read_sql = _pg_safe_read_sql

@st.cache_resource(show_spinner=False)
def _ensure_database_schema_once():
    s_conn, s_db_type = get_db_connection()
    s_cur = s_conn.cursor()
    _is_pg = (s_db_type == "postgres")
    _pk_auto = "SERIAL PRIMARY KEY" if _is_pg else "INTEGER PRIMARY KEY AUTOINCREMENT"
    _blob_type = "BYTEA" if _is_pg else "BLOB"
    _run_schema_migrations(s_conn, s_cur, _pk_auto, _blob_type)
    return True

def _run_schema_migrations(conn, cursor, pk_auto, blob_type):
        # 1. Students Table
    cursor.execute(f"""
    CREATE TABLE IF NOT EXISTS students (
        cnic TEXT PRIMARY KEY,
        roll_no TEXT,
        name TEXT,
        technology TEXT,
        session TEXT,
        phone TEXT,
        email TEXT,
        status TEXT DEFAULT 'Enrolled',
        placed_company TEXT,
        placed_salary TEXT,
        placed_date TEXT,
        cv_verified INTEGER DEFAULT 0,
        cv_blob {blob_type},
        cv_filename TEXT
    )
    """)
    conn.commit()

    # Safe Column Additions for students table
    for col_def in [
        ("email", "TEXT"),
        ("placed_company", "TEXT"),
        ("placed_salary", "TEXT"),
        ("placed_date", "TEXT"),
        ("cv_verified", "INTEGER DEFAULT 0"),
        ("cv_blob", blob_type),
        ("cv_filename", "TEXT")
    ]:
        try:
            cursor.execute(f"ALTER TABLE students ADD COLUMN IF NOT EXISTS {col_def[0]} {col_def[1]}")
            conn.commit()
        except Exception:
            try:
                if hasattr(conn, "rollback"):
                    conn.rollback()
                cursor.execute(f"ALTER TABLE students ADD COLUMN {col_def[0]} {col_def[1]}")
                conn.commit()
            except Exception:
                if hasattr(conn, "rollback"):
                    conn.rollback()

    # 2. Traceability Forms Table
    cursor.execute(f"""
    CREATE TABLE IF NOT EXISTS traceability_forms (
        id {pk_auto},
        cnic TEXT NOT NULL,
        student_name TEXT,
        original_filename TEXT,
        uploaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        extracted_data_blob {blob_type},
        form_blob {blob_type} NOT NULL
    )
    """)
    conn.commit()

    # 3. Activities Table
    cursor.execute(f"""
    CREATE TABLE IF NOT EXISTS activities (
        id {pk_auto},
        activity_type TEXT,
        title TEXT,
        company_or_speaker TEXT,
        activity_date TEXT,
        budget_utilized REAL DEFAULT 0.0
    )
    """)
    conn.commit()

    # 4. Guest Lectures Table
    cursor.execute(f"""
    CREATE TABLE IF NOT EXISTS guest_lectures (
        id {pk_auto},
        speaker_name TEXT,
        speaker_type TEXT,
        qualification TEXT,
        designation TEXT,
        company_org TEXT,
        contact_no TEXT,
        topic TEXT,
        target_tech TEXT,
        target_session TEXT,
        lecture_date TEXT,
        venue TEXT,
        total_attendees INTEGER,
        refreshment_heads INTEGER,
        backdrop_details TEXT,
        bouquet_details TEXT,
        speaker_rating TEXT,
        student_rating TEXT,
        remarks TEXT,
        budget_utilized REAL DEFAULT 0.0,
        status TEXT DEFAULT 'Planned'
    )
    """)
    conn.commit()

    for col_gl in [
        ("district", "TEXT DEFAULT 'Faisalabad'"),
        ("gender", "TEXT DEFAULT 'Male'"),
        ("profession", "TEXT DEFAULT 'Industrialist / Professional'"),
        ("org_address", "TEXT DEFAULT 'Faisalabad'"),
        ("experience", "TEXT DEFAULT '5'"),
        ("cnic", "TEXT"),
        ("email", "TEXT"),
        ("is_registered_employer", "TEXT DEFAULT 'Yes'"),
        ("emp_code", "TEXT DEFAULT 'N/A'"),
        ("speaker_photo_blob", blob_type),
        ("approval_doc_blob", blob_type)
    ]:
        try:
            cursor.execute(f"ALTER TABLE guest_lectures ADD COLUMN IF NOT EXISTS {col_gl[0]} {col_gl[1]}")
            conn.commit()
        except Exception:
            try:
                if hasattr(conn, "rollback"):
                    conn.rollback()
                cursor.execute(f"ALTER TABLE guest_lectures ADD COLUMN {col_gl[0]} {col_gl[1]}")
                conn.commit()
            except Exception:
                if hasattr(conn, "rollback"):
                    conn.rollback()

    # 5. Industry Visits Table
    cursor.execute(f"""
    CREATE TABLE IF NOT EXISTS industry_visits (
        id {pk_auto},
        industry_name TEXT,
        focal_person TEXT,
        focal_contact TEXT,
        technology TEXT,
        session TEXT,
        tour_date TEXT,
        dep_time TEXT,
        travel_mode TEXT,
        other_mode TEXT,
        students_count INTEGER,
        driver_name TEXT,
        conductor_name TEXT,
        aipo_name TEXT,
        assistant_name TEXT,
        approval_copy TEXT,
        rep_name TEXT,
        rep_designation TEXT,
        h_punctuality TEXT,
        h_discipline TEXT,
        h_interest TEXT,
        h_tech_know TEXT,
        h_comm TEXT,
        h_safety TEXT,
        h_recommendation TEXT,
        h_comments TEXT,
        q1_orientation TEXT,
        q2_machinery TEXT,
        q_relevant TEXT DEFAULT 'YES',
        q_objectives TEXT DEFAULT 'YES',
        r_discipline INTEGER,
        r_tech INTEGER,
        r_prof INTEGER,
        r_satisfaction INTEGER,
        candidate_limits_json TEXT,
        budget_utilized REAL DEFAULT 0.0,
        status TEXT DEFAULT 'Planned'
    )
    """)
    conn.commit()

    for col_def in [
        ("q_relevant", "TEXT DEFAULT 'YES'"),
        ("q_objectives", "TEXT DEFAULT 'YES'"),
        ("candidate_limits_json", "TEXT")
    ]:
        try:
            cursor.execute(f"ALTER TABLE industry_visits ADD COLUMN IF NOT EXISTS {col_def[0]} {col_def[1]}")
            conn.commit()
        except Exception:
            try:
                if hasattr(conn, "rollback"):
                    conn.rollback()
                cursor.execute(f"ALTER TABLE industry_visits ADD COLUMN {col_def[0]} {col_def[1]}")
                conn.commit()
            except Exception:
                if hasattr(conn, "rollback"):
                    conn.rollback()

    # 6. Recruitment Drives Table
    cursor.execute(f"""
    CREATE TABLE IF NOT EXISTS recruitment_drives (
        id {pk_auto},
        company_name TEXT,
        company_address TEXT,
        job_location TEXT,
        job_title TEXT,
        experience TEXT,
        vacancies_male INTEGER,
        vacancies_female INTEGER,
        marks_div TEXT,
        salary_range TEXT,
        job_desc TEXT,
        recruitment_mode TEXT,
        eobi INTEGER,
        social_sec INTEGER,
        bonuses INTEGER,
        accommodation TEXT,
        venue TEXT,
        drive_date TEXT,
        drive_time TEXT,
        written_test TEXT,
        target_techs TEXT,
        session TEXT,
        refreshment_heads INTEGER,
        backdrop_details TEXT,
        bouquet_details TEXT,
        dpo_name TEXT,
        ipo_name TEXT,
        tevta_contact TEXT,
        emp_focal TEXT,
        emp_desig TEXT,
        emp_contact TEXT,
        emp_email TEXT,
        students_appeared INTEGER,
        students_shortlisted INTEGER,
        l_coordination TEXT,
        l_staff TEXT,
        l_venue_comfort TEXT,
        l_tech_skills TEXT,
        l_resumes TEXT,
        l_comm_skills TEXT,
        l_dress TEXT,
        l_overall TEXT,
        l_suggestions TEXT,
        candidate_limits_json TEXT,
        budget_utilized REAL DEFAULT 0.0,
        status TEXT DEFAULT 'Planned'
    )
    """)
    conn.commit()

    for col_rd in [
        ("participated_cnics_json", "TEXT"),
        ("selected_cnics_json", "TEXT")
    ]:
        try:
            cursor.execute(f"ALTER TABLE recruitment_drives ADD COLUMN IF NOT EXISTS {col_rd[0]} {col_rd[1]}")
            conn.commit()
        except Exception:
            try:
                if hasattr(conn, "rollback"):
                    conn.rollback()
                cursor.execute(f"ALTER TABLE recruitment_drives ADD COLUMN {col_rd[0]} {col_rd[1]}")
                conn.commit()
            except Exception:
                if hasattr(conn, "rollback"):
                    conn.rollback()

    # 7. Monthly Activity Plans Table
    cursor.execute(f"""
    CREATE TABLE IF NOT EXISTS monthly_activity_plans (
        id {pk_auto},
        month_year TEXT,
        lec_electrical INTEGER,
        lec_mechanical INTEGER,
        lec_civil INTEGER,
        lec_food INTEGER,
        lec_instrument INTEGER,
        lec_autofarm INTEGER,
        lec_textile INTEGER,
        lec_bsc INTEGER,
        lec_shortcourses INTEGER DEFAULT 0,
        guest_lectures_expected INTEGER,
        industrial_visits_expected INTEGER,
        recruitment_drives_expected INTEGER DEFAULT 0,
        budget_breakdown TEXT,
        total_budget REAL,
        created_date DATE DEFAULT CURRENT_DATE
    )
    """)
    conn.commit()

    for col_plan in [
        ("lec_shortcourses", "INTEGER DEFAULT 0"),
        ("recruitment_drives_expected", "INTEGER DEFAULT 0")
    ]:
        try:
            cursor.execute(f"ALTER TABLE monthly_activity_plans ADD COLUMN IF NOT EXISTS {col_plan[0]} {col_plan[1]}")
            conn.commit()
        except Exception:
            try:
                if hasattr(conn, "rollback"):
                    conn.rollback()
                cursor.execute(f"ALTER TABLE monthly_activity_plans ADD COLUMN {col_plan[0]} {col_plan[1]}")
                conn.commit()
            except Exception:
                if hasattr(conn, "rollback"):
                    conn.rollback()

    conn.commit()


_ensure_database_schema_once()

# ----------------- App Config & Header Setup -----------------

st.set_page_config(page_title="GCT IPO Management & Placement System", layout="wide")

st.markdown("""
<style>
    .header-box {
        text-align: center;
        padding-top: 10px;
        padding-bottom: 5px;
    }
    .inst-title {
        font-size: 24px;
        font-weight: 800;
        letter-spacing: 0.5px;
        color: #1F4E79;
        margin-bottom: 2px;
        text-transform: uppercase;
    }
    .sub-title {
        font-size: 17px;
        font-weight: 600;
        color: #555555;
        margin-top: 0px;
    }
    .wa-btn {
        display: inline-block;
        background-color: #25D366;
        color: white !important;
        font-weight: bold;
        padding: 6px 14px;
        border-radius: 5px;
        text-decoration: none;
        font-size: 13px;
        margin-top: 2px;
    }
    .wa-btn:hover {
        background-color: #128C7E;
    }
    .cv-badge {
        background-color: #e6f4ea;
        color: #137333;
        font-weight: bold;
        padding: 3px 8px;
        border-radius: 4px;
        border: 1px solid #ceead6;
        font-size: 12px;
    }
</style>
""", unsafe_allow_html=True)

def find_logo_image(base_names):
    exts = [".png", ".jpg", ".jpeg", ".PNG", ".JPG", ".JPEG", ".webp"]
    script_dir = os.path.dirname(os.path.abspath(__file__)) if '__file__' in locals() else os.getcwd()
    search_dirs = [os.getcwd(), script_dir]
    user_home = os.path.expanduser("~")
    downloads_path = os.path.join(user_home, "Downloads")
    if os.path.exists(downloads_path):
        search_dirs.append(downloads_path)

    for s_dir in search_dirs:
        for b in base_names:
            for ext in exts:
                candidate = os.path.join(s_dir, f"{b}{ext}")
                if os.path.exists(candidate):
                    return candidate
            candidate_direct = os.path.join(s_dir, b)
            if os.path.exists(candidate_direct):
                return candidate_direct
    return None

head_col1, head_col2, head_col3 = st.columns([1.5, 7, 1.5])

with head_col1:
    tevta_img = find_logo_image(["tevta_logo", "TEVTA Logo", "tevta", "TEVTA", "tevta_logo.png"])
    if tevta_img:
        st.image(tevta_img, width=120)
    else:
        st.markdown("<div style='border:1px dashed #aaa; padding:12px; text-align:center; font-size:11px; border-radius:6px;'><b>TEVTA LOGO</b><br>tevta_logo.png</div>", unsafe_allow_html=True)

with head_col2:
    st.markdown("""
    <div class="header-box">
        <div class="inst-title">GOVERNMENT COLLEGE OF TECHNOLOGY, SAMANABAD, FAISALABAD</div>
        <div class="sub-title">IPO Management & Placement System</div>
    </div>
    """, unsafe_allow_html=True)

with head_col3:
    placement_img = find_logo_image(["Placement Logo", "placement_logo", "Placement Logo.png", "placement_logo.png", "placement", "Placement"])
    if placement_img:
        st.image(placement_img, width=120)
    else:
        st.markdown("<div style='border:1px dashed #aaa; padding:12px; text-align:center; font-size:11px; border-radius:6px;'><b>PLACEMENT LOGO</b><br>placement_logo.png</div>", unsafe_allow_html=True)

st.markdown("<hr style='margin-top:5px; margin-bottom:20px; border-top: 2px solid #1F4E79;'>", unsafe_allow_html=True)

TECHNOLOGIES_LIST = [
    "Electrical",
    "Mechanical",
    "Civil",
    "Auto and Farm",
    "Textile",
    "Modern Textile",
    "Instrumentation",
    "Food",
    "CIT",
    "Auto and Diesel",
    "Short Courses"
]

def extract_digits(val):
    if pd.isna(val):
        return ""
    return re.sub(r"[^\d]", "", str(val).strip())

def parse_cnic_input(raw_items):
    parsed_map = {}
    for item in raw_items:
        d = extract_digits(item)
        if len(d) == 13:
            dashed = f"{d[:5]}-{d[5:12]}-{d[12]}"
            parsed_map[d] = dashed
        elif str(item).strip():
            parsed_map[str(item).strip()] = str(item).strip()
    dashed_cnics = list(set(parsed_map.values()))
    plain_cnics = list(set(parsed_map.keys()))
    lookup_cnics = list(set(dashed_cnics + plain_cnics))
    return dashed_cnics, lookup_cnics

# --- Left Sidebar Navigation ---
st.sidebar.markdown("### 📌 Navigation Menu")
menu_items = [
    "📊 Executive Dashboard",
    "📄 CV Parser & Auto-Extractor",
    "📅 Monthly Activity Plan",
    "Student Entry", 
    "Student Directory", 
    "Placement Verification & CNIC Match",
    "Guest Speaker Sessions", 
    "Student Industry Visit",
    "Recruitment Drive",
    "Log Activity", 
    "📑 Traceability Form Upload",
    "Traceability & Export"
]
choice = st.sidebar.radio("Go to Module:", menu_items, index=0)

# =========================================================================
# --- 0. Executive Dashboard & KPI Metrics ---
# =========================================================================
if choice == "📊 Executive Dashboard":
    st.subheader("🏛️ Institutional Placement Executive Dashboard")
    st.caption("Real-time placement rate metrics, budget utilization overview, and industrial linkages tracking.")

    cursor.execute("SELECT COUNT(*) FROM students")
    tot_students = cursor.fetchone()[0] or 0

    cursor.execute("SELECT COUNT(*) FROM students WHERE status = 'Placed'")
    tot_placed = cursor.fetchone()[0] or 0

    placement_pct = (tot_placed / tot_students * 100) if tot_students > 0 else 0.0

    cursor.execute("SELECT SUM(budget_utilized) FROM guest_lectures")
    b_lec = cursor.fetchone()[0] or 0.0
    cursor.execute("SELECT SUM(budget_utilized) FROM industry_visits")
    b_vis = cursor.fetchone()[0] or 0.0
    cursor.execute("SELECT SUM(budget_utilized) FROM recruitment_drives")
    b_dri = cursor.fetchone()[0] or 0.0
    tot_budget = b_lec + b_vis + b_dri

    cursor.execute("SELECT COUNT(*) FROM guest_lectures")
    c_lec = cursor.fetchone()[0] or 0
    cursor.execute("SELECT COUNT(*) FROM industry_visits")
    c_vis = cursor.fetchone()[0] or 0
    cursor.execute("SELECT COUNT(*) FROM recruitment_drives")
    c_dri = cursor.fetchone()[0] or 0
    tot_events = c_lec + c_vis + c_dri

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Registered Students", f"{tot_students:,}")
    m2.metric("Verified Placed", f"{tot_placed:,}")
    m3.metric("Placement Rate", f"{placement_pct:.1f}%")
    m4.metric("Budget Expended", f"PKR {tot_budget:,.0f}")
    m5.metric("Linkage Events", f"{tot_events}")

    st.markdown("---")

    dash_c1, dash_c2 = st.columns(2)
    with dash_c1:
        st.markdown("###### 📈 Placed Students by Technology")
        df_placed_tech = pd.read_sql("""
            SELECT technology AS 'Technology', COUNT(*) AS 'Placed Count' 
            FROM students 
            WHERE status = 'Placed' 
            GROUP BY technology 
            ORDER BY COUNT(*) DESC
        """, conn)
        if not df_placed_tech.empty:
            st.bar_chart(df_placed_tech.set_index("Technology"))
        else:
            st.info("Abhi tak koi student 'Placed' mark nahi hua. Placement verification tab se mark karein.")

    with dash_c2:
        st.markdown("###### 👥 Total Student Enrollment by Technology")
        df_all_tech = pd.read_sql("""
            SELECT technology AS 'Technology', COUNT(*) AS 'Total Enrolled' 
            FROM students 
            GROUP BY technology 
            ORDER BY COUNT(*) DESC
        """, conn)
        if not df_all_tech.empty:
            st.bar_chart(df_all_tech.set_index("Technology"))

    st.markdown("---")
    st.markdown("###### 💼 Recent On-Campus Placements")
    recent_placed_df = pd.read_sql("""
        SELECT name AS 'Student Name', roll_no AS 'Roll No', technology AS 'Trade', 
               session AS 'Session', placed_company AS 'Hired Company', placed_salary AS 'Salary / Package', placed_date AS 'Placement Date'
        FROM students 
        WHERE status = 'Placed'
        ORDER BY rowid DESC LIMIT 10
    """, conn)
    if not recent_placed_df.empty:
        st.dataframe(recent_placed_df, use_container_width=True)
    else:
        st.caption("No recent placed records available yet.")

# =========================================================================
# --- 📄 CV Parser & Auto-Extractor ---
# =========================================================================
elif choice == "📄 CV Parser & Auto-Extractor":
    st.subheader("📄 Student CV Upload & Intelligent Data Extractor")
    st.caption("CV mein se Education section se Trade aur Experience section se Company auto-detect karein. Original CV compression ke sath minimal space mein database mein save ho jati hai.")

    cv_input_method = st.radio(
        "Choose Input Method:",
        ["📁 Upload CV File (.pdf, .docx, .png, .jpg, .jpeg)", "✍️ Paste CV Text Manually (Instant Backup)"],
        horizontal=True
    )

    extracted_text = ""
    compressed_cv_bytes = None
    cv_orig_filename = ""

    if cv_input_method == "📁 Upload CV File (.pdf, .docx, .png, .jpg, .jpeg)":
        uploaded_cv = st.file_uploader(
            "Select Student CV File:",
            type=["pdf", "docx", "png", "jpg", "jpeg", "webp"]
        )

        if uploaded_cv is not None:
            cv_orig_filename = uploaded_cv.name
            file_raw_bytes = uploaded_cv.read()
            compressed_cv_bytes = compress_file_bytes(file_raw_bytes)
            
            with st.spinner("Analyzing document structure and extracting text..."):
                extracted_text = extract_text_from_bytes(file_raw_bytes, cv_orig_filename)
                
            orig_kb = len(file_raw_bytes) / 1024
            comp_kb = len(compressed_cv_bytes) / 1024
            st.caption(f"💾 **Storage Compression:** Original `{orig_kb:.1f} KB` ➔ Compressed to `{comp_kb:.1f} KB` (Saved {(1 - comp_kb/orig_kb)*100:.1f}% space)")
    else:
        pasted_text = st.text_area(
            "Paste CV Text Here (Copy from PDF/Word/WhatsApp and paste):",
            height=180,
            placeholder="Muhammad Ahmad\nS/O Muhammad Aslam\nCNIC: 33100-1234567-1\nPhone: 0300-1234567\nEDUCATION:\nDAE Electrical Technology (2023-2026)\nEXPERIENCE:\nTrainee Technician at Hyundai Nishat Motors"
        )
        if pasted_text.strip():
            extracted_text = pasted_text.strip()
            cv_orig_filename = "Pasted_CV_Text.txt"
            compressed_cv_bytes = compress_file_bytes(extracted_text.encode('utf-8'))

    if extracted_text:
        parsed_data = parse_cv_details(extracted_text, TECHNOLOGIES_LIST)

        st.success("✅ Document Successfully Read & Parsed!")

        detected_cnic = parsed_data['cnic']
        is_existing = False
        existing_student_info = None

        if detected_cnic:
            cursor.execute("SELECT * FROM students WHERE cnic = ?", (detected_cnic,))
            existing_row = cursor.fetchone()
            if existing_row:
                is_existing = True
                desc_cols = [d[0] for d in cursor.description]
                existing_student_info = dict(zip(desc_cols, existing_row))

        st.markdown("---")
        
        if is_existing:
            st.warning(f"⚠️ **Attention: CNIC `{detected_cnic}` pehle se Database mein mojood hai!**")
            st.markdown("##### 🔍 Side-by-Side Data Comparison (Database vs Fresh CV)")
            
            cmp_data = [
                {"Field": "Student Name", "Previously in Database": str(existing_student_info.get('name') or ''), "Extracted from Fresh CV": str(parsed_data.get('name') or '')},
                {"Field": "Roll No", "Previously in Database": str(existing_student_info.get('roll_no') or ''), "Extracted from Fresh CV": str(parsed_data.get('roll_no') or '')},
                {"Field": "Technology (Education)", "Previously in Database": str(existing_student_info.get('technology') or ''), "Extracted from Fresh CV": str(parsed_data.get('technology') or '')},
                {"Field": "Session", "Previously in Database": str(existing_student_info.get('session') or ''), "Extracted from Fresh CV": str(parsed_data.get('session') or '')},
                {"Field": "Mobile / WhatsApp", "Previously in Database": str(existing_student_info.get('phone') or ''), "Extracted from Fresh CV": str(parsed_data.get('phone') or '')},
                {"Field": "Email Address", "Previously in Database": str(existing_student_info.get('email') or ''), "Extracted from Fresh CV": str(parsed_data.get('email') or '')},
                {"Field": "Current Status", "Previously in Database": str(existing_student_info.get('status') or ''), "Extracted from Fresh CV": "Placed" if parsed_data.get('company') else "Enrolled"},
                {"Field": "Industry Name (Experience)", "Previously in Database": str(existing_student_info.get('placed_company') or ''), "Extracted from Fresh CV": str(parsed_data.get('company') or '')},
                {"Field": "Salary / Package", "Previously in Database": str(existing_student_info.get('placed_salary') or ''), "Extracted from Fresh CV": "35,000" if parsed_data.get('company') else ""}
            ]
            cmp_df = pd.DataFrame(cmp_data)
            st.table(cmp_df)
            st.info("💡 **Tip:** CV ka data form mein pre-fill ho chuka hai. Jo cheez CV mein na mili ho, aap upar Database column se dekh kar neeche form mein fill kar sakte hain.")
        else:
            if detected_cnic:
                st.success(f"✨ **Naya Student Hai!** CNIC `{detected_cnic}` pehle se database mein mojood nahi hai.")
            else:
                st.info("ℹ️ CV mein se standard CNIC detect nahi hua. Barah-e-karam neeche CNIC lazmi verify ya enter karein.")

        with st.form("cv_parsed_form"):
            st.markdown("##### 📝 Review & Confirm Final Values before Saving")
            
            c_col1, c_col2 = st.columns(2)
            with c_col1:
                default_name = parsed_data['name'] if parsed_data['name'] else (str(existing_student_info.get('name') or '') if is_existing else '')
                cv_name = st.text_input("Student Name*", value=default_name)
                
                cv_fname = st.text_input("Father Name (Optional)", value=parsed_data['father_name'])
                cv_cnic = st.text_input("CNIC (Unique Key)*", value=parsed_data['cnic'], placeholder="33100-1234567-1")
                
                default_roll = parsed_data['roll_no'] if parsed_data['roll_no'] else (str(existing_student_info.get('roll_no') or '') if is_existing else '')
                cv_roll = st.text_input("Roll No", value=default_roll)
                
                selected_tech = parsed_data['technology'] if parsed_data['technology'] else (str(existing_student_info.get('technology') or 'Electrical') if is_existing else 'Electrical')
                tech_idx = TECHNOLOGIES_LIST.index(selected_tech) if selected_tech in TECHNOLOGIES_LIST else 0
                cv_tech = st.selectbox("Technology / Trade (From Education)*", TECHNOLOGIES_LIST, index=tech_idx)
            
            with c_col2:
                default_phone = parsed_data['phone'] if parsed_data['phone'] else (str(existing_student_info.get('phone') or '') if is_existing else '')
                cv_phone = st.text_input("Mobile / WhatsApp Number*", value=default_phone)
                
                default_email = parsed_data['email'] if parsed_data['email'] else (str(existing_student_info.get('email') or '') if is_existing else '')
                cv_email = st.text_input("Email Address", value=default_email)
                
                default_session = parsed_data['session'] if parsed_data['session'] else (str(existing_student_info.get('session') or '2023-2026') if is_existing else '2023-2026')
                cv_session = st.text_input("Session*", value=default_session)
                
                status_choices = ["Enrolled", "Placed", "Pass Out", "Dropped Out"]
                default_stat_idx = 1 if (parsed_data['company'] or (is_existing and existing_student_info.get('status') == 'Placed')) else 0
                cv_status = st.selectbox("Placement Status", status_choices, index=default_stat_idx)
                
                default_comp = parsed_data['company'] if parsed_data['company'] else (str(existing_student_info.get('placed_company') or '') if is_existing else '')
                cv_company = st.text_input("Industry / Company Name (From Experience)*", value=default_comp)
                
                default_sal = (str(existing_student_info.get('placed_salary') or '35,000') if is_existing and existing_student_info.get('placed_salary') else "35,000") if cv_company else ""
                cv_salary = st.text_input("Salary / Package (PKR)", value=default_sal)

            st.markdown("---")
            save_c1, save_c2 = st.columns([2, 1])
            
            if is_existing:
                submit_save = save_c1.form_submit_button("🔄 Confirm & Overwrite / Update Existing Student (Save CV File)")
            else:
                submit_save = save_c1.form_submit_button("✅ Confirm & Save New Student into Directory (Save CV File)")

            if submit_save:
                clean_cnic_val = str(cv_cnic or "").strip()
                clean_name_val = str(cv_name or "").strip()
                clean_phone_val = str(cv_phone or "").strip()
                clean_roll_val = str(cv_roll or "").strip()
                clean_session_val = str(cv_session or "").strip()
                clean_email_val = str(cv_email or "").strip()
                clean_comp_val = str(cv_company or "").strip()
                clean_sal_val = str(cv_salary or "").strip()

                if not re.match(r"^\d{5}-\d{7}-\d{1}$", clean_cnic_val):
                    st.error("Ghalat CNIC format! CNIC lazmi dashes ke sath hona chahiye (e.g. 33100-1234567-1).")
                elif not clean_name_val or not clean_phone_val:
                    st.warning("Student Name aur Mobile Number likhna lazmi hai.")
                else:
                    cursor.execute("""
                        INSERT INTO students (cnic, roll_no, name, technology, session, phone, email, status, placed_company, placed_salary, cv_verified, cv_blob, cv_filename)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?, ?)
                        ON CONFLICT(cnic) DO UPDATE SET
                            roll_no = excluded.roll_no,
                            name = excluded.name,
                            technology = excluded.technology,
                            session = excluded.session,
                            phone = excluded.phone,
                            email = excluded.email,
                            status = excluded.status,
                            placed_company = excluded.placed_company,
                            placed_salary = excluded.placed_salary,
                            cv_verified = 1,
                            cv_blob = CASE WHEN excluded.cv_blob IS NOT NULL THEN excluded.cv_blob ELSE students.cv_blob END,
                            cv_filename = CASE WHEN excluded.cv_filename IS NOT NULL THEN excluded.cv_filename ELSE students.cv_filename END
                    """, (
                        clean_cnic_val, clean_roll_val, clean_name_val, cv_tech, clean_session_val,
                        clean_phone_val, clean_email_val, cv_status, clean_comp_val, clean_sal_val,
                        compressed_cv_bytes, cv_orig_filename
                    ))
                    conn.commit()
                    if is_existing:
                        st.success(f"Student '{clean_name_val}' (CNIC: {clean_cnic_val}) ka record kamyabi se UPDATE ho gaya aur Original Compressed CV save ho gayi! 🟢")
                    else:
                        st.success(f"Naya Student '{clean_name_val}' (CNIC: {clean_cnic_val}) Directory mein SAVE ho gaya aur Original Compressed CV save ho gayi! 🟢")

        with st.expander("🔍 View Raw Extracted CV Text (For Reference)"):
            st.text_area("Raw Text Extracted from File:", value=extracted_text, height=180)

# =========================================================================
# --- 📅 Monthly Activity Plan ---
# =========================================================================
elif choice == "📅 Monthly Activity Plan":
    st.subheader("📅 Monthly Placement Activity Plan & Allowances")
    st.caption("Monthly Career Counseling & Linkages Plan (DPO Submission) aur Principal Communication Charges Application.")

    m_tab1, m_tab2, m_tab3 = st.tabs([
        "📝 Create & Download Monthly Plan", 
        "📩 Principal Application (Communication Allowance)",
        "📂 Archived Plans & History"
    ])

    with m_tab1:
        st.markdown("##### Plan Details & Month Selection")
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            plan_month = st.text_input("Plan Month & Year*", value="October 2026", placeholder="e.g. October 2026, November 2026")
        with col_m2:
            st.info("💡 **TEVTA Standard Sanctions:** Guest Speaker = **PKR 10,000**, Industrial Tour = **PKR 30,000**, Recruitment Drive = **PKR 30,000**.")

        st.markdown("---")
        st.markdown("###### 🎓 Career Counseling Lectures (Technology-wise Target)")
        t_col1, t_col2, t_col3 = st.columns(3)
        with t_col1:
            lec_elec = st.number_input("Electrical", min_value=0, value=12, step=1)
            lec_mech = st.number_input("Mechanical", min_value=0, value=12, step=1)
            lec_civil = st.number_input("Civil", min_value=0, value=9, step=1)
        with t_col2:
            lec_food = st.number_input("Food", min_value=0, value=3, step=1)
            lec_inst = st.number_input("Instrumentation", min_value=0, value=6, step=1)
            lec_auto = st.number_input("Auto & Farm", min_value=0, value=6, step=1)
        with t_col3:
            lec_tex = st.number_input("Textile", min_value=0, value=3, step=1)
            lec_bsc = st.number_input("B.Sc", min_value=0, value=3, step=1)
            lec_short = st.number_input("Short Courses", min_value=0, value=5, step=1)

        total_lec_calc = lec_elec + lec_mech + lec_civil + lec_food + lec_inst + lec_auto + lec_tex + lec_bsc + lec_short
        st.markdown(f"**Total Career Counseling Lectures Planned:** `{total_lec_calc} Lectures`")

        st.markdown("---")
        st.markdown("###### 🏭 Linkages & Activities (Expected in Month)")
        act_col1, act_col2, act_col3 = st.columns(3)
        with act_col1:
            exp_spk = st.number_input("Expected Guest Speaker Lectures", min_value=0, value=1, step=1)
            st.caption("Allow: PKR 10,000 / lecture")
        with act_col2:
            exp_vis = st.number_input("Expected Industrial Study Visits", min_value=0, value=1, step=1)
            st.caption("Allow: PKR 30,000 / visit")
        with act_col3:
            exp_dri = st.number_input("Expected Recruitment Drives", min_value=0, value=1, step=1)
            st.caption("Allow: PKR 30,000 / drive")

        calc_guest_budget = exp_spk * 10000
        calc_tour_budget = exp_vis * 30000
        calc_drive_budget = exp_dri * 30000
        calc_total_budget = float(calc_guest_budget + calc_tour_budget + calc_drive_budget)

        auto_breakdown_lines = []
        if exp_spk > 0:
            auto_breakdown_lines.append(f"{exp_spk}x10,000 = {calc_guest_budget:,} (Guest Speaker Lecture)")
        if exp_vis > 0:
            auto_breakdown_lines.append(f"{exp_vis}x30,000 = {calc_tour_budget:,} (Students Industrial Visit)")
        if exp_dri > 0:
            auto_breakdown_lines.append(f"{exp_dri}x30,000 = {calc_drive_budget:,} (Recruitment Drive Facilitation)")

        default_breakdown_str = "\n".join(auto_breakdown_lines) if auto_breakdown_lines else "No budget allocated"

        st.markdown("---")
        st.markdown("###### 💰 Budget Required Breakdown")
        b_col1, b_col2 = st.columns(2)
        with b_col1:
            b_breakdown = st.text_area("Budget Calculation Breakdown*", value=default_breakdown_str, height=100)
        with b_col2:
            b_total = st.number_input("Total Budget Required (PKR)*", min_value=0.0, value=calc_total_budget, step=1000.0)

        plan_dict = {
            'month_year': plan_month,
            'lec_electrical': lec_elec,
            'lec_mechanical': lec_mech,
            'lec_civil': lec_civil,
            'lec_food': lec_food,
            'lec_instrument': lec_inst,
            'lec_autofarm': lec_auto,
            'lec_textile': lec_tex,
            'lec_bsc': lec_bsc,
            'lec_shortcourses': lec_short,
            'guest_lectures_expected': exp_spk,
            'industrial_visits_expected': exp_vis,
            'recruitment_drives_expected': exp_dri,
            'budget_breakdown': b_breakdown,
            'total_budget': b_total
        }

        save_plan_col, print_col, excel_col = st.columns(3)
        with save_plan_col:
            if st.button("💾 Save Plan to Archive"):
                cursor.execute("""
                    INSERT INTO monthly_activity_plans (
                        month_year, lec_electrical, lec_mechanical, lec_civil, lec_food,
                        lec_instrument, lec_autofarm, lec_textile, lec_bsc, lec_shortcourses,
                        guest_lectures_expected, industrial_visits_expected, recruitment_drives_expected,
                        budget_breakdown, total_budget
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    plan_month, lec_elec, lec_mech, lec_civil, lec_food,
                    lec_inst, lec_auto, lec_tex, lec_bsc, lec_short,
                    exp_spk, exp_vis, exp_dri, b_breakdown, b_total
                ))
                conn.commit()
                st.success(f"Monthly Plan for '{plan_month}' save ho gaya!")

        with print_col:
            plan_html = render_monthly_plan_html(plan_dict)
            st.download_button(
                label="🖨️ Direct Printable A4 / PDF (.html)",
                data=plan_html,
                file_name=f"Monthly_Plan_{plan_month.replace(' ', '_')}.html",
                mime="text/html"
            )

        with excel_col:
            plan_excel = generate_monthly_plan_excel(plan_dict)
            st.download_button(
                label="📊 Download Plan as Excel (.xlsx)",
                data=plan_excel,
                file_name=f"Monthly_Activity_Plan_{plan_month.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    with m_tab2:
        st.markdown("##### 📩 Request for Payment of Communication Charges")
        st.caption("Rs. 2,500/- allowance, TEVTA notification aur sign details permanently configured hain. Sirf Month aur Date chunein:")

        with st.form("comm_charges_simple_form"):
            cc1, cc2 = st.columns(2)
            with cc1:
                app_month = st.text_input("Application for Month & Year*", value="October 2026")
            with cc2:
                app_date_val = st.date_input("Application Date")

            st.form_submit_button("Preview / Update Application")

        comm_data = {
            'target_month': app_month,
            'app_date': str(app_date_val)
        }

        st.markdown("---")
        st.markdown("###### 📥 Download Header-Free Clean Application")
        d_col_w, d_col_p = st.columns(2)
        
        with d_col_w:
            app_docx = generate_comm_charges_application_docx(comm_data)
            st.download_button(
                label="📄 Download Clean Application (.docx)",
                data=app_docx,
                file_name=f"Communication_Charges_{app_month.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        
        with d_col_p:
            app_html = render_comm_charges_html(comm_data)
            st.download_button(
                label="🖨️ Direct Print / Save as PDF (.html)",
                data=app_html,
                file_name=f"Communication_Charges_{app_month.replace(' ', '_')}.html",
                mime="text/html"
            )

        st.markdown("---")
        with st.expander("👀 Application Text Preview (No Header)", expanded=True):
            st.markdown(f"""
<div style="text-align: right;"><b>Dated:</b> {comm_data['app_date']}</div>

**The Principal**  
GCT Samanabad Faisalabad  

**Subject: Request for Payment of Communication Charges for the Month of {comm_data['target_month']}**  

With due respect, it is stated that I am performing my duties as Institute Placement Officer (IPO) at GCT Samanabad Faisalabad. As per the TEVTA notification No. TEVTA/DIR-HR/CA-69/2022, dated 23-02-2022, Communication Charges of Rs. 2,500/- per month have been sanctioned for Institute Placement Officers.  

It is therefore requested that the Communication Charges for the month of **{comm_data['target_month']}** amounting to **Rs. 2,500/-** may kindly be released/paid to me as per the approved entitlement.  

Kindly accord approval for the payment of the above-mentioned Communication Charges.  

<br>
**Regards,**  
<br>
**Muhammad Wasif Ali**  
Institute Placement Officer  
            """, unsafe_allow_html=True)

    with m_tab3:
        st.markdown("##### History of Submitted Monthly Plans")
        df_plans = pd.read_sql("SELECT id, month_year AS 'Month & Year', total_budget AS 'Total Budget (PKR)', created_date AS 'Submission Date' FROM monthly_activity_plans ORDER BY id DESC", conn)
        if not df_plans.empty:
            st.dataframe(df_plans, use_container_width=True)
            
            selected_pid = st.selectbox("Select Past Plan to Re-Download or Delete:", df_plans['id'].tolist(), format_func=lambda x: f"ID {x} - {df_plans.loc[df_plans['id']==x, 'Month & Year'].values[0]}")
            
            cursor.execute("SELECT * FROM monthly_activity_plans WHERE id = ?", (selected_pid,))
            p_row = cursor.fetchone()
            p_cols = [d[0] for d in cursor.description]
            p_arch_data = dict(zip(p_cols, p_row))

            down_p1, down_p2, del_p = st.columns(3)
            with down_p1:
                st.download_button(
                    "🖨️ Print A4 PDF (.html)",
                    render_monthly_plan_html(p_arch_data),
                    file_name=f"Archived_Plan_{p_arch_data['month_year'].replace(' ', '_')}.html",
                    mime="text/html"
                )
            with down_p2:
                st.download_button(
                    "📊 Download Excel (.xlsx)",
                    generate_monthly_plan_excel(p_arch_data),
                    file_name=f"Archived_Plan_{p_arch_data['month_year'].replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            with del_p:
                if st.button("🗑️ Delete Selected Plan", key="del_plan_btn"):
                    cursor.execute("DELETE FROM monthly_activity_plans WHERE id = ?", (selected_pid,))
                    conn.commit()
                    st.success("Plan delete ho gaya!")
                    st.rerun()
        else:
            st.info("Abhi tak koi monthly plan save nahi kiya gaya.")

# =========================================================================
# --- 1. Student Entry Form ---
# =========================================================================
elif choice == "Student Entry":
    st.subheader("Single Student Entry")
    with st.form("student_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            name = st.text_input("Student Name*")
            cnic = st.text_input("CNIC (Format: 33100-1234567-1)*")
            roll_no = st.text_input("Roll No")
            tech = st.selectbox("Technology*", TECHNOLOGIES_LIST)
        with col2:
            session = st.text_input("Session (e.g., 2023-2026, 2024-2027)*", value="2023-2026")
            phone = st.text_input("Mobile Number*", placeholder="0300-1234567")
            email = st.text_input("Email Address", placeholder="student@example.com")
            status = st.selectbox("Current Status", ["Enrolled", "Placed", "Pass Out", "Dropped Out"])

        submitted = st.form_submit_button("Save Student Record")
        if submitted:
            cnic_clean = str(cnic or "").strip()
            if not re.match(r"^\d{5}-\d{7}-\d{1}$", cnic_clean):
                st.error("Ghalat CNIC format! CNIC lazmi dashes k sath hona chahiye: 33100-1234567-1")
            elif not name or not phone:
                st.warning("Student Name, Mobile Number aur Session likhna lazmi hai.")
            else:
                try:
                    cursor.execute(
                        "INSERT INTO students (cnic, roll_no, name, technology, session, phone, email, status) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                        (cnic_clean, str(roll_no or "").strip(), str(name or "").strip(), tech, str(session or "").strip(), str(phone or "").strip(), str(email or "").strip(), status)
                    )
                    conn.commit()
                    st.success(f"Student '{name}' kamyabi se save ho gaya!")
                except (sqlite3.IntegrityError, psycopg2.IntegrityError):
                    conn.rollback()
                    st.error("Yeh CNIC pehle se database mein mojood hai!")

# =========================================================================
# --- 2. Student Directory ---
# =========================================================================
elif choice == "Student Directory":
    st.subheader("Student Directory & Data Management")
    dir_tab1, dir_tab2 = st.tabs(["Browse, Search & Edit Records", "Bulk Upload (CSV / Excel)"])
    
    with dir_tab1:
        f1, f2, f3 = st.columns(3)
        with f1:
            filter_tech = st.selectbox("Filter by Technology", ["All Technologies"] + TECHNOLOGIES_LIST)
        with f2:
            filter_session = st.text_input("Filter by Session (e.g. 2023-2026)")
        with f3:
            filter_status = st.selectbox("Filter Status", ["All Statuses", "Placed", "Enrolled", "Pass Out", "Dropped Out"])

        f4, f5 = st.columns(2)
        with f4:
            search_query = st.text_input("Quick Search (CNIC, Roll No or Name)", placeholder="e.g. 33100-1234567-1 or Ahmad")
        with f5:
            cursor.execute("SELECT DISTINCT placed_company FROM students WHERE placed_company IS NOT NULL AND TRIM(placed_company) != ''")
            existing_comps = [r[0] for r in cursor.fetchall() if r[0]]
            search_industry = st.text_input("🏢 Search by Industry / Company (Placed Students)", placeholder="e.g. Hyundai, Interloop, Sitara")
            if existing_comps:
                st.caption(f"Registered Placed Companies: {', '.join(existing_comps[:6])}{'...' if len(existing_comps)>6 else ''}")

        query = """
            SELECT 
                cnic AS 'CNIC', 
                roll_no AS 'Roll No', 
                name AS 'Name', 
                technology AS 'Technology', 
                session AS 'Session', 
                phone AS 'Mobile Number', 
                email AS 'Email', 
                status AS 'Status', 
                placed_company AS 'Hired Company', 
                placed_salary AS 'Salary', 
                placed_date AS 'Placement Date',
                cv_verified,
                cv_filename
            FROM students 
            WHERE 1=1
        """
        params = []
        
        if filter_tech != "All Technologies":
            query += " AND technology = ?"
            params.append(filter_tech)
        if filter_session:
            query += " AND session LIKE ?"
            params.append(f"%{filter_session.strip()}%")
        if filter_status != "All Statuses":
            query += " AND status = ?"
            params.append(filter_status)
        if search_industry:
            query += " AND placed_company LIKE ?"
            params.append(f"%{search_industry.strip()}%")
        if search_query:
            clean_digits = extract_digits(search_query)
            if len(clean_digits) == 13:
                dashed_search = f"{clean_digits[:5]}-{clean_digits[5:12]}-{clean_digits[12]}"
                query += " AND (cnic = ? OR cnic = ?)"
                params.extend([dashed_search, clean_digits])
            else:
                query += " AND (name LIKE ? OR cnic LIKE ? OR roll_no LIKE ?)"
                params.extend([f"%{search_query.strip()}%", f"%{search_query.strip()}%", f"%{search_query.strip()}%"])

        query += " ORDER BY session DESC, technology ASC, rowid DESC LIMIT 1000"
        df_display = pd.read_sql(query, conn, params=params)

        if search_industry.strip() and not df_display.empty:
            st.markdown(f"#### 📊 Placed Candidates Summary in '{search_industry.strip()}'")
            ind_c1, ind_c2 = st.columns(2)
            with ind_c1:
                st.markdown("###### 🔧 By Technology (Trade)")
                tech_grp = df_display.groupby('Technology')['Name'].count().reset_index(name='Students Count')
                st.dataframe(tech_grp, use_container_width=True)
            with ind_c2:
                st.markdown("###### 🎓 By Session")
                sess_grp = df_display.groupby('Session')['Name'].count().reset_index(name='Students Count')
                st.dataframe(sess_grp, use_container_width=True)

        if not df_display.empty:
            df_display_view = df_display.copy()
            df_display_view['CV Verified'] = df_display_view['cv_verified'].apply(lambda x: "🟢 Verified by CV" if x == 1 else "⚪ Standard")
            df_display_view['CV Stored'] = df_display_view['cv_filename'].apply(lambda x: f"📁 {x}" if (x and str(x).strip()) else "No CV")
            df_display_view = df_display_view.drop(columns=['cv_verified', 'cv_filename'])
            st.write(f"**Showing Records:** {len(df_display_view)} (Matching students)")
            st.dataframe(df_display_view, use_container_width=True)
        else:
            st.write("**Showing Records:** 0")
            st.dataframe(df_display, use_container_width=True)

        if not df_display.empty:
            st.markdown("---")
            st.markdown("#### 💬 Direct WhatsApp Click-to-Chat & CV Downloads")
            
            with st.expander("📲 Send 1-Click WhatsApp Message or Download Student CV", expanded=True):
                wa_st_sel = st.selectbox(
                    "Select Student:",
                    df_display['CNIC'].tolist(),
                    format_func=lambda x: f"{df_display.loc[df_display['CNIC']==x, 'Name'].values[0]} - Roll# {df_display.loc[df_display['CNIC']==x, 'Roll No'].values[0]} ({df_display.loc[df_display['CNIC']==x, 'Technology'].values[0]} | {df_display.loc[df_display['CNIC']==x, 'Session'].values[0]}) {'[🟢 Verified by CV]' if df_display.loc[df_display['CNIC']==x, 'cv_verified'].values[0] == 1 else ''}",
                    key="dir_wa_student_sel"
                )
                
                sel_row = df_display.loc[df_display['CNIC'] == wa_st_sel].iloc[0]
                s_name = str(sel_row['Name'])
                s_roll = str(sel_row['Roll No'])
                s_tech = str(sel_row['Technology'])
                s_phone = str(sel_row['Mobile Number'])
                s_status = str(sel_row['Status'])
                s_comp = str(sel_row['Hired Company']) if pd.notna(sel_row['Hired Company']) else ""
                
                cursor.execute("SELECT cv_blob, cv_filename FROM students WHERE cnic = ?", (wa_st_sel,))
                cv_res = cursor.fetchone()
                
                down_col, msg_type_col = st.columns([1, 2])
                with down_col:
                    if cv_res and cv_res[0]:
                        decomp_data = decompress_file_bytes(cv_res[0])
                        cv_fname = cv_res[1] or f"CV_{s_name.replace(' ', '_')}.pdf"
                        st.download_button(
                            label=f"📥 Download Original CV ({cv_fname})",
                            data=decomp_data,
                            file_name=cv_fname,
                            mime="application/octet-stream"
                        )
                    else:
                        st.caption("ℹ️ Is student ki CV upload nahi hui.")

                with msg_type_col:
                    msg_type = st.radio(
                        "Select Message Purpose:",
                        ["📋 Academic & Placement Status Verification", "🎉 Job Placement Congratulations", "📢 General IPO Announcement"],
                        horizontal=True,
                        key=f"msg_type_{wa_st_sel}"
                    )

                if msg_type == "🎉 Job Placement Congratulations":
                    default_dir_msg = f"Assalam-o-Alaikum {s_name} (Roll No: {s_roll}, Trade: {s_tech}),\nPlacement Cell GCT Samanabad ki taraf se aapko '{s_comp}' mein placement par dilli mubarakbaad pesh ki jati hai. Apni joining formalities Placement Office confirm karein.\nRegards,\nIPO Cell, GCT Samanabad"
                elif msg_type == "📢 General IPO Announcement":
                    default_dir_msg = f"Assalam-o-Alaikum {s_name} (Roll No: {s_roll}, Trade: {s_tech}),\nBarah-e-karam kal subah 9:00 AM Placement Office (IPO Cell) GCT Samanabad Faisalabad mein report karein.\nRegards,\nInstitute Placement Officer"
                else:
                    default_dir_msg = f"Assalam-o-Alaikum {s_name} (Roll No: {s_roll}, Trade: {s_tech}),\nYeh message Placement Cell GCT Samanabad Faisalabad ki taraf se bheja ja raha hai regarding your academic and career placement record. Barah-e-karam apna updated placement status confirm karein.\nRegards,\nInstitute Placement Officer"

                custom_wa_msg = st.text_area(
                    f"Message Preview for {s_name} (Editable):", 
                    value=default_dir_msg, 
                    height=90, 
                    key=f"dir_wa_msg_box_{wa_st_sel}_{msg_type}"
                )
                
                wa_link = generate_whatsapp_url(s_phone, custom_wa_msg)
                if wa_link:
                    st.markdown(f"<a href='{wa_link}' target='_blank' class='wa-btn'>💬 Send WhatsApp to {s_name} ({s_phone})</a>", unsafe_allow_html=True)
                else:
                    st.warning(f"⚠️ {s_name} ka mobile number ({s_phone}) valid international format mein nahi hai.")

            st.markdown("---")
            st.markdown("#### ✏️ Edit / Update Student Record")
            st.caption("Jab aap oper search karte hain, toh neeche foran us student ki details khul jati hain jinhein aap update ya delete kar saktay hain:")

            edit_cnic_options = df_display['CNIC'].tolist()
            selected_edit_cnic = st.selectbox(
                "Select Student to Edit / Correct:", 
                edit_cnic_options, 
                format_func=lambda x: f"{x} - {df_display.loc[df_display['CNIC']==x, 'Name'].values[0]} ({df_display.loc[df_display['CNIC']==x, 'Technology'].values[0]} | {df_display.loc[df_display['CNIC']==x, 'Session'].values[0]}) {'[🟢 Verified by CV]' if df_display.loc[df_display['CNIC']==x, 'cv_verified'].values[0] == 1 else ''}"
            )

            cursor.execute("SELECT * FROM students WHERE cnic = ?", (selected_edit_cnic,))
            st_row = cursor.fetchone()
            st_cols = [d[0] for d in cursor.description]
            st_data = dict(zip(st_cols, st_row))

            with st.expander(f"📝 Click to Edit Details for: {st_data['name']} (CNIC: {st_data['cnic']})", expanded=True):
                if st_data.get('cv_verified') == 1:
                    st.markdown("<span class='cv-badge'>🟢 Authenticated & Verified by CV</span>", unsafe_allow_html=True)
                    st.write("")

                with st.form(f"edit_student_form_{selected_edit_cnic}"):
                    ed_c1, ed_c2 = st.columns(2)
                    with ed_c1:
                        st.text_input("CNIC (Unique Key - Fixed)", value=st_data['cnic'], disabled=True)
                        ed_name = st.text_input("Student Name*", value=st_data['name'])
                        ed_roll = st.text_input("Roll No", value=st_data['roll_no'] or "")
                        ed_tech_idx = TECHNOLOGIES_LIST.index(st_data['technology']) if st_data['technology'] in TECHNOLOGIES_LIST else 0
                        ed_tech = st.selectbox("Technology*", TECHNOLOGIES_LIST, index=ed_tech_idx)
                        ed_session = st.text_input("Session*", value=st_data['session'] or "2023-2026")

                    with ed_c2:
                        ed_phone = st.text_input("Mobile Number*", value=st_data['phone'] or "")
                        ed_email = st.text_input("Email Address", value=st_data['email'] or "")
                        all_statuses = ["Enrolled", "Placed", "Pass Out", "Dropped Out"]
                        ed_stat_idx = all_statuses.index(st_data['status']) if st_data['status'] in all_statuses else 0
                        ed_status = st.selectbox("Student Status", all_statuses, index=ed_stat_idx)
                        ed_comp = st.text_input("Placed Company (Industry Name)", value=st_data['placed_company'] or "")
                        ed_salary = st.text_input("Salary / Package (PKR)", value=st_data['placed_salary'] or "")

                    ed_btn_col, del_btn_col = st.columns([2, 1])
                    with ed_btn_col:
                        save_student_changes = st.form_submit_button("💾 Save Changes to Student Record")
                    with del_btn_col:
                        delete_student_record = st.form_submit_button("🗑️ Delete This Student")

                    if save_student_changes:
                        cursor.execute("""
                            UPDATE students 
                            SET name = ?, roll_no = ?, technology = ?, session = ?, phone = ?, email = ?, status = ?, placed_company = ?, placed_salary = ?
                            WHERE cnic = ?
                        """, (
                            str(ed_name or "").strip(), str(ed_roll or "").strip(), ed_tech, str(ed_session or "").strip(),
                            str(ed_phone or "").strip(), str(ed_email or "").strip(), ed_status, str(ed_comp or "").strip(), str(ed_salary or "").strip(),
                            selected_edit_cnic
                        ))
                        conn.commit()
                        st.success(f"Student '{ed_name}' ka record kamyabi se update ho gaya!")
                        st.rerun()

                    if delete_student_record:
                        cursor.execute("DELETE FROM students WHERE cnic = ?", (selected_edit_cnic,))
                        conn.commit()
                        st.warning(f"Student '{st_data['name']}' (CNIC: {selected_edit_cnic}) database se delete ho chuka hai.")
                        st.rerun()

            st.markdown("---")
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                df_display.to_excel(writer, index=False, sheet_name="Students_List")
            
            file_exp_name = f"Students_Placed_{search_industry.replace(' ', '_')}.xlsx" if search_industry else "GCT_Students_Export.xlsx"
            st.download_button(
                label=f"📥 Download Current Filtered List as Excel ({len(df_display)} Students) (.xlsx)",
                data=excel_buffer.getvalue(),
                file_name=file_exp_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    with dir_tab2:
        st.markdown("##### Upload Student Data (CSV ya Excel)")
        st.info("Bari file (10,000 se 100,000 records) ke liye .csv format sab se tez aur behtareen hai.")

        sample_df = pd.DataFrame([{
            "Name": "Muhammad Ahmad",
            "Roll No": "EL-101",
            "CNIC": "33100-1234567-1",
            "Mobile number": "0300-1234567",
            "Email": "ahmad@gmail.com",
            "Technology": "Electrical",
            "Session": "2023-2026",
            "Status": "Placed",
            "Placed Company": "Hyundai Nishat",
            "Salary": "40,000"
        }])
        
        sample_csv = sample_df.to_csv(index=False).encode('utf-8')
        st.download_button("Download Sample File (CSV)", sample_csv, "Sample_Students_Template.csv", "text/csv")
        
        uploaded_file = st.file_uploader("Upload Student Sheet (.csv ya .xlsx)", type=["csv", "xlsx"])
        if uploaded_file is not None:
            try:
                uploaded_df = safe_read_file(uploaded_file)

                st.write(f"Found **{len(uploaded_df):,}** rows in the file.")
                st.dataframe(uploaded_df.head(5), use_container_width=True)

                col_map = {str(c).strip().lower(): c for c in uploaded_df.columns}
                required_cols = ['name', 'cnic', 'mobile number', 'email', 'session']
                missing_cols = [c for c in required_cols if c not in col_map]

                if missing_cols:
                    st.error(f"File mein yeh zaroori columns missing hain: {', '.join(missing_cols)}")
                else:
                    st.markdown("---")
                    st.markdown("###### ⚙️ Duplicate CNIC Handling Mode (Aik Option Select Karein):")
                    upload_mode = st.radio(
                        "Agar CNIC pehle se Database mein mojood ho toh kya kiya jaye?",
                        [
                            "🛡️ Sirf Naye Students Add Karein (Purana Data Mehfooz / Skip Duplicates)",
                            "🔄 Purana Data Naye Data se Update / Replace Karein (Overwrite Existing)"
                        ],
                        index=0
                    )

                    if st.button("🚀 Start Fast Import into Database"):
                        with st.spinner(f"Importing {len(uploaded_df):,} records... please wait."):
                            records_to_insert = []
                            uploaded_cnics = []
                            for _, row in uploaded_df.iterrows():
                                r_name = str(row[col_map['name']]).strip() if pd.notna(row[col_map['name']]) else ""
                                r_cnic = str(row[col_map['cnic']]).strip() if pd.notna(row[col_map['cnic']]) else ""
                                r_phone = str(row[col_map['mobile number']]).strip() if pd.notna(row[col_map['mobile number']]) else ""
                                r_email = str(row[col_map['email']]).strip() if pd.notna(row[col_map['email']]) else ""
                                r_session = str(row[col_map['session']]).strip() if pd.notna(row[col_map['session']]) else ""
                                r_tech = str(row[col_map.get('technology', '')]).strip() if ('technology' in col_map and pd.notna(row[col_map['technology']])) else "General"
                                r_roll = str(row[col_map.get('roll no', '')]).strip() if ('roll no' in col_map and pd.notna(row[col_map['roll no']])) else ""
                                r_status = str(row[col_map.get('status', '')]).strip() if ('status' in col_map and pd.notna(row[col_map['status']])) else "Enrolled"
                                r_comp = str(row[col_map.get('placed company', '')]).strip() if ('placed company' in col_map and pd.notna(row[col_map['placed company']])) else ""
                                r_sal = str(row[col_map.get('salary', '')]).strip() if ('salary' in col_map and pd.notna(row[col_map['salary']])) else ""

                                if r_cnic and r_name:
                                    records_to_insert.append((r_cnic, r_roll, r_name, r_tech, r_session, r_phone, r_email, r_status, r_comp, r_sal))
                                    uploaded_cnics.append(r_cnic)

                            cursor.execute("SELECT COUNT(*) FROM students")
                            b_cnt = cursor.fetchone()[0]

                            if upload_mode == "🛡️ Sirf Naye Students Add Karein (Purana Data Mehfooz / Skip Duplicates)":
                                cursor.executemany("""
                                    INSERT OR IGNORE INTO students 
                                    (cnic, roll_no, name, technology, session, phone, email, status, placed_company, placed_salary) 
                                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                                """, records_to_insert)
                                conn.commit()

                                cursor.execute("SELECT COUNT(*) FROM students")
                                a_cnt = cursor.fetchone()[0]
                                added = a_cnt - b_cnt
                                skipped = len(records_to_insert) - added
                                st.success(f"Mukammal! Kul Database: **{a_cnt:,}** | Naye Add Huay: **{added:,}** | Pehle se mojood (Skipped): **{skipped:,}**")

                            else:
                                existing_matched = 0
                                if uploaded_cnics:
                                    chunk_size = 500
                                    for i in range(0, len(uploaded_cnics), chunk_size):
                                        c_chunk = uploaded_cnics[i:i+chunk_size]
                                        pl_h = ",".join("?" * len(c_chunk))
                                        cursor.execute(f"SELECT COUNT(*) FROM students WHERE cnic IN ({pl_h})", c_chunk)
                                        existing_matched += cursor.fetchone()[0]

                                cursor.executemany("""
                                    INSERT OR REPLACE INTO students 
                                    (cnic, roll_no, name, technology, session, phone, email, status, placed_company, placed_salary) 
                                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                                """, records_to_insert)
                                conn.commit()

                                cursor.execute("SELECT COUNT(*) FROM students")
                                a_cnt = cursor.fetchone()[0]
                                added = a_cnt - b_cnt
                                updated_cnt = existing_matched

                                st.success(f"Mukammal! Naye Add Huay: **{added:,}** | Pehle se mojood jo Update/Replace huay: **{updated_cnt:,}** | Kul Database: **{a_cnt:,}**")
            except Exception as e:
                st.error(f"File process karne mein masla aaya: {e}")

# =========================================================================
# --- 3. Placement Verification & CNIC Match ---
# =========================================================================
elif choice == "Placement Verification & CNIC Match":
    st.subheader("🔍 Placement Extraction & Status Lifecycle Update")
    st.markdown(
        "Company se aayi hui recruited CNIC list yahan upload karein ya paste karein. "
        "System **sirf aur sirf inhi CNIC** ka record database se extract karega aur aapko inhein **'Placed'** status mark karne ki sahulat deta hai."
    )

    v_tab1, v_tab2 = st.tabs(["📁 Upload File (Excel / CSV)", "✍️ Direct Paste CNICs"])
    raw_input_cnics = []

    with v_tab1:
        st.markdown("##### Upload Company File")
        comp_file = st.file_uploader("Upload Excel / CSV File (Company ki taraf se aayi hui list)", type=["xlsx", "csv"], key="placed_file_upload")
        if comp_file is not None:
            try:
                df_up = safe_read_file(comp_file)

                st.write(f"File parhi gayi: **{len(df_up)} rows**")
                
                cnic_col = None
                for c in df_up.columns:
                    if "cnic" in str(c).lower():
                        cnic_col = c
                        break
                if not cnic_col:
                    cnic_col = st.selectbox("Sheet mein se CNIC wala column select karein:", df_up.columns)

                if cnic_col:
                    raw_input_cnics = df_up[cnic_col].dropna().tolist()
            except Exception as e:
                st.error(f"File process karne mein masla aaya: {e}")

    with v_tab2:
        st.markdown("##### Direct Paste")
        text_cnics = st.text_area(
            "CNICs Paste karein (Har line par aik CNIC):", 
            placeholder="33100-1234567-1\n33100-7654321-2\n3310212345678"
        )
        if text_cnics.strip():
            raw_input_cnics = re.split(r"[\n,\s;]+", text_cnics.strip())

    if raw_input_cnics:
        target_dashed_cnics, all_lookup_cnics = parse_cnic_input(raw_input_cnics)

        st.markdown("---")
        st.info(f"Targeted CNICs to Extract: **{len(target_dashed_cnics)} Unique Students**")

        if st.button("🚀 Extract Only These Students"):
            with st.spinner("Extracting from student records..."):
                placeholders = ",".join("?" * len(all_lookup_cnics))
                sql_extract = f"""
                    SELECT 
                        name AS 'Name',
                        cnic AS 'CNIC',
                        roll_no AS 'Roll No',
                        technology AS 'Technology',
                        session AS 'Session',
                        phone AS 'Mobile Number',
                        email AS 'Email',
                        status AS 'Current Status',
                        placed_company AS 'Hired Company'
                    FROM students 
                    WHERE cnic IN ({placeholders})
                    ORDER BY technology, name
                """
                extracted_df = pd.read_sql(sql_extract, conn, params=all_lookup_cnics)

                found_in_db = set(extracted_df['CNIC'].tolist())
                unmatched = [c for c in target_dashed_cnics if c not in found_in_db]

            st.session_state['last_extracted_df'] = extracted_df
            st.session_state['last_unmatched'] = unmatched
            st.session_state['last_lookup_cnics'] = all_lookup_cnics

    if 'last_extracted_df' in st.session_state and not st.session_state['last_extracted_df'].empty:
        extracted_df = st.session_state['last_extracted_df']
        unmatched = st.session_state['last_unmatched']
        all_lookup_cnics = st.session_state['last_lookup_cnics']

        st.success(f"Kamyabi! Total **{len(extracted_df)}** target students extract ho gaye.")
        
        if unmatched:
            st.warning(f"⚠️ **{len(unmatched)}** CNICs aapke database mein mojood nahi hain.")
            with st.expander("Dekhein kon se CNIC database mein nahi miley"):
                st.write(unmatched)

        st.dataframe(extracted_df, use_container_width=True)

        st.markdown("##### 📥 Download Extracted Target Students")
        d_col1, d_col2 = st.columns(2)
        with d_col1:
            excel_out = io.BytesIO()
            with pd.ExcelWriter(excel_out, engine='openpyxl') as writer:
                extracted_df.to_excel(writer, index=False, sheet_name="Target_Recruited")
            st.download_button(
                label=f"📊 Download Extracted List ({len(extracted_df)} Students) (.xlsx)",
                data=excel_out.getvalue(),
                file_name="Company_Recruited_Extracted_Data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with d_col2:
            csv_out = extracted_df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label=f"📄 Download Extracted List ({len(extracted_df)} Students) (.csv)",
                data=csv_out,
                file_name="Company_Recruited_Extracted_Data.csv",
                mime="text/csv"
            )

        st.markdown("---")
        with st.expander("💼 Update Status of These Candidates to 'Placed' (One-Click)"):
            st.caption("In tamam extracted students ka status Student Directory mein 'Placed' ho jayega aur Dashboard par count hoga.")
            with st.form("mark_placed_form"):
                pl_c1, pl_c2, pl_c3 = st.columns(3)
                with pl_c1:
                    p_comp_name = st.text_input("Recruiting Company Name*", value="Hyundai Nishat Group")
                with pl_c2:
                    p_salary = st.text_input("Salary / Package (PKR)", value="38,000")
                with pl_c3:
                    p_date = st.date_input("Placement Date")

                update_status_btn = st.form_submit_button("✅ Mark All Extracted Students as 'Placed'")
                if update_status_btn:
                    placeholders = ",".join("?" * len(all_lookup_cnics))
                    cursor.execute(f"""
                        UPDATE students 
                        SET status = 'Placed', placed_company = ?, placed_salary = ?, placed_date = ?
                        WHERE cnic IN ({placeholders})
                    """, [str(p_comp_name or "").strip(), str(p_salary or "").strip(), str(p_date)] + all_lookup_cnics)
                    conn.commit()
                    st.success(f"Mubarak ho! In **{len(extracted_df)}** students ka status kamyabi se 'Placed' mark ho gaya hai.")

# =========================================================================
# --- 4. Guest Speaker Sessions ---
# =========================================================================
elif choice == "Guest Speaker Sessions":
    st.subheader("Guest Speaker Management & Documentation")
    tab1, tab2, tab3 = st.tabs([
        "1. Plan Lecture & Principal Approval", 
        "2. Post-Session Feedback & Final Archive", 
        "3. Permanent Archive, Edit & Delete"
    ])

    with tab1:
        st.markdown("##### Pre-Session: Event Planning, Speaker Photo, Annexure 'E' & Administrative Approval")
        with st.form("plan_lecture_form"):
            c1, c2 = st.columns(2)
            with c1:
                spk_name = st.text_input("Speaker Name*")
                gender = st.selectbox("Gender", ["Male", "Female"])
                spk_type = st.selectbox("Speaker Category / Profession*", [
                    "Industrialist / Professional", "University / Academia", "Entrepreneur", "Self Employed / Freelancer"
                ])
                designation = st.text_input("Designation / Title", value="Associate Professor / Manager")
                qualification = st.text_input("Speaker Qualification", value="Ph.D / B.Sc Engineering")
                experience = st.text_input("Experience (in Years)", value="5")
                cnic = st.text_input("Speaker CNIC", placeholder="33100-1234567-1")
                contact_no = st.text_input("Speaker Contact No.", placeholder="0300-1234567")
                email = st.text_input("Speaker Email", placeholder="speaker@company.com")
            with c2:
                org = st.text_input("Industry / Institute Name*")
                org_address = st.text_input("Organization Address", value="Faisalabad")
                district = st.text_input("District of Guest", value="Faisalabad")
                is_reg_emp = st.selectbox("Is Guest Speaker from Registered Employer?", ["Yes", "No"])
                emp_code = st.text_input("Employer Career Dashboard Code (If Yes)", value="N/A")
                tech = st.selectbox("Target Technology*", TECHNOLOGIES_LIST)
                session = st.text_input("Target Session", value="2023-2026")
                topic = st.text_input("Lecture Topic*")
                date = st.date_input("Proposed Date")
                venue = st.text_input("Venue", value="Main Auditorium, GCT Samanabad")

            st.markdown("##### 🖼️ Guest Speaker Portrait Picture (For Certificate of Appreciation)")
            spk_photo_file = st.file_uploader(
                "Upload Guest Speaker Picture (JPG / PNG / WEBP) - IPO & Principal Signatures Already Auto-Applied",
                type=["jpg", "jpeg", "png", "webp"],
                key="plan_spk_photo_uploader"
            )

            st.markdown("##### Administrative & Protocol Requirements")
            c3, c4, c5 = st.columns(3)
            with c3:
                refreshment = st.number_input("Refreshment Required (Heads)", min_value=0, value=10, step=1)
            with c4:
                backdrop = st.selectbox("Stage Backdrop / Flex", ["Backdrop Flex Required", "Digital Presentation Only", "Not Required"])
            with c5:
                bouquet = st.selectbox("Welcome Bouquet / Souvenir", ["Bouquet & Souvenir Required", "Bouquet Only", "Shield / Souvenir Only", "Not Required"])

            save_btn = st.form_submit_button("Save Session Plan")
            if save_btn:
                if spk_name and org and topic:
                    raw_photo = spk_photo_file.read() if spk_photo_file is not None else None
                    comp_photo = compress_file_bytes(raw_photo) if raw_photo else None

                    app_payload = {
                        'spk_name': spk_name, 'speaker_name': spk_name, 'speaker_type': spk_type,
                        'designation': designation, 'org': org, 'qualification': qualification,
                        'tech': tech, 'session': session, 'topic': topic, 'date': str(date),
                        'venue': venue, 'refreshment_heads': refreshment,
                        'backdrop_details': backdrop, 'bouquet_details': bouquet
                    }
                    saved_app_bytes = generate_principal_application(app_payload).getvalue()
                    comp_app_blob = compress_file_bytes(saved_app_bytes)

                    cursor.execute("""
                    INSERT INTO guest_lectures (
                        speaker_name, speaker_type, qualification, designation, company_org,
                        contact_no, topic, target_tech, target_session, lecture_date, venue,
                        refreshment_heads, backdrop_details, bouquet_details,
                        district, gender, profession, org_address, experience,
                        cnic, email, is_registered_employer, emp_code,
                        speaker_photo_blob, approval_doc_blob,
                        speaker_rating, student_rating, remarks, budget_utilized, status
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        spk_name, spk_type, qualification, designation, org,
                        contact_no, topic, tech, session, str(date), venue,
                        refreshment, backdrop, bouquet,
                        district, gender, spk_type, org_address, experience,
                        cnic, email, is_reg_emp, emp_code,
                        comp_photo, comp_app_blob,
                        "Pending", "Pending", "Planned", 0.0, "Planned"
                    ))
                    conn.commit()
                    st.success("✅ Session plan, Principal Approval Letter aur Speaker Photo permanently save ho gaye! Neeche buttons se download karein.")
                else:
                    st.warning("Speaker Name, Organization aur Topic likhna lazmi hai.")

        if spk_name and org and topic:
            live_photo_bytes = None
            if spk_photo_file is not None:
                try:
                    spk_photo_file.seek(0)
                    live_photo_bytes = spk_photo_file.read()
                except Exception:
                    pass

            app_data = {
                'spk_name': spk_name, 'speaker_name': spk_name, 'speaker_type': spk_type,
                'profession': spk_type, 'gender': gender, 'designation': designation,
                'org': org, 'company_org': org, 'org_address': org_address, 'district': district,
                'qualification': qualification, 'experience': experience, 'cnic': cnic,
                'contact_no': contact_no, 'email': email, 'is_registered_employer': is_reg_emp,
                'emp_code': emp_code, 'tech': tech, 'session': session,
                'topic': topic, 'date': str(date), 'lecture_date': str(date), 'venue': venue,
                'refreshment_heads': refreshment, 'backdrop_details': backdrop, 'bouquet_details': bouquet
            }
            app_file = generate_principal_application(app_data)
            annex_e_docx = generate_annexure_e(app_data)
            cert_pdf_bytes = generate_guest_speaker_certificate(
                app_data, speaker_photo_bytes=live_photo_bytes, output_format="PDF"
            )
            
            p_c1, p_c2, p_c3, p_c4 = st.columns(4)
            with p_c1:
                st.download_button(
                    label="📥 Principal Approval (.docx)",
                    data=app_file,
                    file_name=f"Approval_Letter_{spk_name.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            with p_c2:
                st.download_button(
                    label="📄 Annexure 'E' Profile (.docx)",
                    data=annex_e_docx,
                    file_name=f"Annexure_E_{spk_name.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            with p_c3:
                st.download_button(
                    label="🏆 Honor Certificate (.pdf)",
                    data=cert_pdf_bytes,
                    file_name=f"Honor_Certificate_{spk_name.replace(' ', '_')}.pdf",
                    mime="application/pdf"
                )
            with p_c4:
                html_body = f"""
                <p class='right-align'><b>Date:</b> {app_data['date']}</p>
                <p><b>To,</b><br>&nbsp;&nbsp;&nbsp;&nbsp;The Worthy Principal,<br>&nbsp;&nbsp;&nbsp;&nbsp;Government College of Technology,<br>&nbsp;&nbsp;&nbsp;&nbsp;Samanabad, Faisalabad.</p>
                <p><b>SUBJECT: PERMISSION & ADMINISTRATIVE APPROVAL FOR GUEST SPEAKER LECTURE ({app_data['tech'].upper()})</b></p>
                <p>Respected Sir,<br><br>It is submitted that the Placement Cell (IPO Office) has planned a guest speaker lecture to bridge industry-academia linkages and provide practical orientation to students:</p>
                <ul>
                    <li><b>Speaker:</b> {app_data['spk_name']} ({app_data['designation']})</li>
                    <li><b>Organization:</b> {app_data['org']}</li>
                    <li><b>Topic:</b> {app_data['topic']}</li>
                    <li><b>Target Class:</b> DAE {app_data['tech']} ({app_data['session']})</li>
                    <li><b>Date & Venue:</b> {app_data['date']} at {app_data['venue']}</li>
                </ul>
                <p>Administrative arrangements requested: Refreshment for {app_data['refreshment_heads']} persons, {app_data['backdrop_details']}, and {app_data['bouquet_details']}.</p>
                <br><br><br>
                <p><b>Institute Placement Officer (IPO)</b><br>GCT Samanabad, Faisalabad</p>
                """
                full_html = render_printable_html(f"Approval_{spk_name}", html_body)
                st.download_button("🖨️ Printable Approval (.html)", full_html, f"Print_Approval_{spk_name}.html", "text/html")

    with tab2:
        st.markdown("##### Post-Session: Annexure 'E', Annexure 'F', Honor Certificate & Final Archive")
        df_lectures = pd.read_sql("SELECT id, speaker_name, topic, lecture_date, status FROM guest_lectures ORDER BY id DESC", conn)
        
        if not df_lectures.empty:
            session_dict = {f"ID {row['id']} - {row['speaker_name']} ({row['topic']}) [{row['status']}]": row['id'] for _, row in df_lectures.iterrows()}
            selected_label = st.selectbox("Select Session to Complete:", list(session_dict.keys()))
            selected_id = session_dict[selected_label]

            cursor.execute("SELECT * FROM guest_lectures WHERE id = ?", (selected_id,))
            s_row = cursor.fetchone()
            c_names = [d[0] for d in cursor.description]
            s_data = dict(zip(c_names, s_row))

            with st.form("annexure_f_form"):
                st.write(f"**Speaker:** {s_data['speaker_name']} ({s_data.get('designation', '')}) | **Date:** {s_data['lecture_date']}")
                f1, f2 = st.columns(2)
                with f1:
                    attendees = st.number_input("Total Students Attended", min_value=1, value=int(s_data['total_attendees'] or 45))
                    methodology = st.selectbox("Teaching Methodology", ["Presentation / Slides", "Interactive & Discussion", "Board Lecture", "Verbal Communication"])
                    spk_rating = st.select_slider("Speaker Rating (1 to 5)", options=["1", "2", "3", "4", "5"], value=str(s_data['speaker_rating'] if s_data['speaker_rating'] in ["1","2","3","4","5"] else "5"))
                with f2:
                    discipline = st.selectbox("Student Discipline", ["Discipline Maintained & Attentive", "Good", "Satisfactory"])
                    engagement = st.selectbox("Student Engagement", ["Highly Interactive Q&A", "Moderate Engagement", "Passive"])
                    std_rating = st.select_slider("Student Rating (1 to 5)", options=["1", "2", "3", "4", "5"], value=str(s_data['student_rating'] if s_data['student_rating'] in ["1","2","3","4","5"] else "5"))
                
                takeaways = st.text_input("Key Takeaways", value="Practical industrial exposure & career roadmap")
                upd_photo_tab2 = st.file_uploader("Upload / Replace Speaker Photo for Certificate (Optional)", type=["jpg", "jpeg", "png", "webp"], key=f"upd_photo_t2_{selected_id}")
                
                st.markdown("---")
                st.markdown("###### 💰 Final Activity Expense & Submission")
                budget_val = st.number_input("Budget Utilized for this Activity (PKR)*", min_value=0.0, value=float(s_data['budget_utilized'] or 10000.0), step=500.0)

                submit_archive = st.form_submit_button("✅ Final Submit & Save Permanently to Archive")
                if submit_archive:
                    new_photo_blob = s_data.get('speaker_photo_blob')
                    if upd_photo_tab2 is not None:
                        new_photo_blob = compress_file_bytes(upd_photo_tab2.read())
                    cursor.execute("""
                    UPDATE guest_lectures 
                    SET total_attendees = ?, speaker_rating = ?, student_rating = ?, remarks = ?, budget_utilized = ?, speaker_photo_blob = ?, status = 'Completed'
                    WHERE id = ?
                    """, (attendees, spk_rating, std_rating, f"Method: {methodology} | Takeaways: {takeaways}", budget_val, new_photo_blob, selected_id))
                    conn.commit()
                    st.success(f"Activity final submit ho kar permanent record mein archive ho gayi! (Budget Utilized: PKR {budget_val:,.2f})")
                    st.rerun()

            s_data.update({
                'total_attendees': attendees, 'methodology': methodology,
                'speaker_rating': spk_rating, 'discipline': discipline,
                'engagement': engagement, 'student_rating': std_rating, 'takeaways': takeaways,
                'budget_utilized': budget_val
            })
            active_photo_t2 = s_data.get('speaker_photo_blob')
            if upd_photo_tab2 is not None:
                try:
                    upd_photo_tab2.seek(0)
                    active_photo_t2 = upd_photo_tab2.read()
                except Exception:
                    pass

            annex_e_docx2 = generate_annexure_e(s_data)
            annex_f_docx = generate_annexure_f(s_data)
            cert_t2_pdf = generate_guest_speaker_certificate(s_data, speaker_photo_bytes=active_photo_t2, output_format="PDF")
            cert_t2_png = generate_guest_speaker_certificate(s_data, speaker_photo_bytes=active_photo_t2, output_format="PNG")

            dl_c1, dl_c2, dl_c3, dl_c4 = st.columns(4)
            with dl_c1:
                st.download_button(
                    label="📄 Annexure 'E' (.docx)",
                    data=annex_e_docx2,
                    file_name=f"Annexure_E_{s_data['speaker_name'].replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            with dl_c2:
                st.download_button(
                    label="📄 Annexure 'F' (.docx)",
                    data=annex_f_docx,
                    file_name=f"Annexure_F_{s_data['speaker_name'].replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            with dl_c3:
                st.download_button(
                    label="🏆 Certificate (.pdf)",
                    data=cert_t2_pdf,
                    file_name=f"Certificate_{s_data['speaker_name'].replace(' ', '_')}.pdf",
                    mime="application/pdf"
                )
            with dl_c4:
                st.download_button(
                    label="🖼️ Certificate Image (.png)",
                    data=cert_t2_png,
                    file_name=f"Certificate_{s_data['speaker_name'].replace(' ', '_')}.png",
                    mime="image/png"
                )
        else:
            st.info("Pehle Tab 1 se koi session plan karein.")

    with tab3:
        st.markdown("##### 📂 Permanent Archived Records, Re-Download Documents, Edit & Delete")
        df_all = pd.read_sql("SELECT id, speaker_name, designation, company_org, topic, target_tech, target_session, lecture_date, budget_utilized, status FROM guest_lectures ORDER BY id DESC", conn)
        if not df_all.empty:
            st.dataframe(df_all, use_container_width=True)
            
            st.markdown("---")
            st.markdown("###### ⚙️ Manage Archived Session (Re-download Approval, Annexures, Certificate / Edit / Delete)")
            re_id = st.selectbox("Session Select Karein:", df_all['id'].tolist(), format_func=lambda x: f"ID {x} - {df_all.loc[df_all['id']==x, 'speaker_name'].values[0]} ({df_all.loc[df_all['id']==x, 'lecture_date'].values[0]})")
            
            cursor.execute("SELECT * FROM guest_lectures WHERE id = ?", (re_id,))
            re_data = dict(zip([d[0] for d in cursor.description], cursor.fetchone()))
            
            if re_data.get('approval_doc_blob'):
                re_app_docx = decompress_file_bytes(re_data['approval_doc_blob'])
            else:
                re_app_payload = {
                    'spk_name': re_data.get('speaker_name', ''),
                    'speaker_type': re_data.get('speaker_type', 'Industrialist / Professional'),
                    'designation': re_data.get('designation', ''),
                    'org': re_data.get('company_org', ''),
                    'qualification': re_data.get('qualification', ''),
                    'tech': re_data.get('target_tech', ''),
                    'session': re_data.get('target_session', ''),
                    'topic': re_data.get('topic', ''),
                    'date': str(re_data.get('lecture_date', '')),
                    'venue': re_data.get('venue', 'Main Auditorium, GCT Samanabad'),
                    'refreshment_heads': re_data.get('refreshment_heads') or 10,
                    'backdrop_details': re_data.get('backdrop_details') or 'Backdrop Flex Required',
                    'bouquet_details': re_data.get('bouquet_details') or 'Bouquet & Souvenir Required'
                }
                re_app_docx = generate_principal_application(re_app_payload).getvalue()

            re_e_docx = generate_annexure_e(re_data)
            re_f_docx = generate_annexure_f(re_data)
            re_cert_pdf = generate_guest_speaker_certificate(re_data, speaker_photo_bytes=re_data.get('speaker_photo_blob'), output_format="PDF")
            re_cert_png = generate_guest_speaker_certificate(re_data, speaker_photo_bytes=re_data.get('speaker_photo_blob'), output_format="PNG")

            arc_c1, arc_c2, arc_c3, arc_c4, arc_c5 = st.columns(5)
            with arc_c1:
                st.download_button(
                    label="📥 Principal Approval",
                    data=re_app_docx,
                    file_name=f"Archived_Approval_Letter_ID_{re_id}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            with arc_c2:
                st.download_button(
                    label="📥 Annexure 'E'",
                    data=re_e_docx,
                    file_name=f"Archived_Annexure_E_ID_{re_id}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            with arc_c3:
                st.download_button(
                    label="📥 Annexure 'F'",
                    data=re_f_docx,
                    file_name=f"Archived_Annexure_F_ID_{re_id}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            with arc_c4:
                st.download_button(
                    label="🏆 Certificate (PDF)",
                    data=re_cert_pdf,
                    file_name=f"Archived_Certificate_ID_{re_id}.pdf",
                    mime="application/pdf"
                )
            with arc_c5:
                st.download_button(
                    label="🖼️ Certificate (PNG)",
                    data=re_cert_png,
                    file_name=f"Archived_Certificate_ID_{re_id}.png",
                    mime="image/png"
                )

            with st.expander("👁️ Preview Guest Speaker Honor Certificate"):
                st.image(re_cert_png, caption=f"Certificate of Appreciation - {re_data.get('speaker_name', '')}", use_container_width=True)

            with st.expander(f"✏️ Edit / Correct Entry & Speaker Photo (ID: {re_id})"):
                with st.form(f"edit_guest_{re_id}"):
                    e_spk = st.text_input("Speaker Name", value=re_data['speaker_name'])
                    e_desig = st.text_input("Designation", value=str(re_data.get('designation') or ''))
                    e_org = st.text_input("Organization / Institute", value=str(re_data.get('company_org') or ''))
                    e_topic = st.text_input("Lecture Topic", value=re_data['topic'])
                    e_c1, e_c2 = st.columns(2)
                    with e_c1:
                        e_att = st.number_input("Attendees", min_value=0, value=int(re_data['total_attendees'] or 0))
                        e_bud = st.number_input("Budget Utilized (PKR)", min_value=0.0, value=float(re_data['budget_utilized'] or 0.0), step=500.0)
                    with e_c2:
                        e_stat = st.selectbox("Status", ["Completed", "Planned"], index=0 if re_data['status'] == 'Completed' else 1)
                        e_rem = st.text_input("Remarks", value=str(re_data['remarks'] or ''))
                    e_photo_new = st.file_uploader("Upload / Replace Speaker Portrait Photo (Optional)", type=["jpg", "jpeg", "png", "webp"], key=f"edit_photo_g_{re_id}")
                    
                    if st.form_submit_button("💾 Save Corrected Data"):
                        upd_photo_blob = re_data.get('speaker_photo_blob')
                        if e_photo_new is not None:
                            upd_photo_blob = compress_file_bytes(e_photo_new.read())
                        upd_app_payload = {
                            'spk_name': e_spk,
                            'speaker_type': re_data.get('speaker_type', 'Industrialist / Professional'),
                            'designation': e_desig,
                            'org': e_org,
                            'qualification': re_data.get('qualification', ''),
                            'tech': re_data.get('target_tech', ''),
                            'session': re_data.get('target_session', ''),
                            'topic': e_topic,
                            'date': str(re_data.get('lecture_date', '')),
                            'venue': re_data.get('venue', 'Main Auditorium, GCT Samanabad'),
                            'refreshment_heads': re_data.get('refreshment_heads') or 10,
                            'backdrop_details': re_data.get('backdrop_details') or 'Backdrop Flex Required',
                            'bouquet_details': re_data.get('bouquet_details') or 'Bouquet & Souvenir Required'
                        }
                        upd_app_blob = compress_file_bytes(generate_principal_application(upd_app_payload).getvalue())
                        cursor.execute("""
                            UPDATE guest_lectures 
                            SET speaker_name = ?, designation = ?, company_org = ?, topic = ?, total_attendees = ?, budget_utilized = ?, status = ?, remarks = ?, speaker_photo_blob = ?, approval_doc_blob = ?
                            WHERE id = ?
                        """, (e_spk, e_desig, e_org, e_topic, e_att, e_bud, e_stat, e_rem, upd_photo_blob, upd_app_blob, re_id))
                        conn.commit()
                        st.success("Record kamyabi se update ho gaya!")
                        st.rerun()

            with st.expander(f"🗑️ Delete Record (ID: {re_id})"):
                st.warning("⚠️ Kya aap waqai is record ko hamesha ke liye delete karna chahte hain?")
                confirm_del_g = st.checkbox("Haan, main is record ko permanently delete karna chahta hoon", key=f"del_chk_g_{re_id}")
                if st.button("🔴 Delete This Guest Session", key=f"del_btn_g_{re_id}"):
                    if confirm_del_g:
                        cursor.execute("DELETE FROM guest_lectures WHERE id = ?", (re_id,))
                        conn.commit()
                        st.success(f"Record ID {re_id} delete ho chuka hai.")
                        st.rerun()
                    else:
                        st.error("Pehle confirmation checkbox par click karein.")
        else:
            st.write("Abhi tak koi session record nahi mila.")

# =========================================================================
# --- 5. Student Industry Visit (With Dynamic Sync & Multi-Tech Student Selection) ---
# =========================================================================
elif choice == "Student Industry Visit":
    st.subheader("Student Industry Study Tour & Documentation")
    
    tour_tab1, tour_tab2, tour_tab3 = st.tabs([
        "1. Plan Tour & Approval Application", 
        "2. Feedback, Budget & Final Submit", 
        "3. Permanent Archive, Edit & Delete"
    ])

    with tour_tab1:
        st.markdown("##### Pre-Tour: Visit Planning & Accompanying Staff")
        
        st.markdown("###### 🎯 Targeted Technologies & Session (Student Selection for Tour)")
        t_sel_c1, t_sel_c2 = st.columns(2)
        with t_sel_c1:
            tour_target_techs = st.multiselect(
                "Select Eligible Technologies / Trades for this Tour*",
                TECHNOLOGIES_LIST,
                default=["Civil"]
            )
        with t_sel_c2:
            tour_target_session = st.text_input("Target Session (e.g. 2023-2026)*", value="2023-2026", key="tour_session_input")

        if not tour_target_techs:
            st.warning("⚠️ Barah-e-karam kam az kam aik Technology lazmi select karein.")

        st.markdown("###### 👥 Technology-Wise Student Allocation")
        tour_selection_mode = st.radio(
            "Student Selection Mode:",
            ["Tamam Matching Students (All in Selected Technologies)", "Custom Tadaad (Specify Count Per Technology)"],
            horizontal=True,
            key="tour_sel_mode"
        )

        tour_custom_counts = {}
        all_tour_students = []

        if tour_selection_mode == "Custom Tadaad (Specify Count Per Technology)" and tour_target_techs:
            t_num_cols = min(len(tour_target_techs), 4) if len(tour_target_techs) > 0 else 1
            t_cols_dyn = st.columns(t_num_cols)
            
            for idx, t in enumerate(tour_target_techs):
                with t_cols_dyn[idx % t_num_cols]:
                    cursor.execute("SELECT COUNT(*) FROM students WHERE session = ? AND technology = ?", (tour_target_session, t))
                    avail = cursor.fetchone()[0]
                    st.markdown(f"**🔧 {t}**")
                    cnt = st.number_input(
                        f"DB Available: {avail}",
                        min_value=0,
                        max_value=max(avail, 1000),
                        value=min(avail, 40),
                        step=5,
                        key=f"tour_count_{t}"
                    )
                    tour_custom_counts[t] = cnt
                    if cnt > 0:
                        q_t = f"SELECT roll_no AS 'Roll No', name AS 'Name', technology AS 'Technology', cnic AS 'CNIC', phone AS 'Mobile Number', email AS 'Email' FROM students WHERE session = ? AND technology = ? ORDER BY roll_no LIMIT {cnt}"
                        df_sub = pd.read_sql(q_t, conn, params=[tour_target_session, t])
                        if not df_sub.empty:
                            all_tour_students.append(df_sub)
        else:
            for t in tour_target_techs:
                cursor.execute("SELECT COUNT(*) FROM students WHERE session = ? AND technology = ?", (tour_target_session, t))
                avail = cursor.fetchone()[0]
                tour_custom_counts[t] = avail
                q_t = "SELECT roll_no AS 'Roll No', name AS 'Name', technology AS 'Technology', cnic AS 'CNIC', phone AS 'Mobile Number', email AS 'Email' FROM students WHERE session = ? AND technology = ? ORDER BY roll_no"
                df_sub = pd.read_sql(q_t, conn, params=[tour_target_session, t])
                if not df_sub.empty:
                    all_tour_students.append(df_sub)

        if all_tour_students:
            tour_students_allocated_df = pd.concat(all_tour_students, ignore_index=True)
        else:
            tour_students_allocated_df = pd.DataFrame(columns=['Roll No', 'Name', 'Technology', 'CNIC', 'Mobile Number', 'Email'])

        st.info(f"📊 Total Trainees Selected for this Industrial Tour: **{len(tour_students_allocated_df)} Students**")

        st.markdown("---")
        with st.form("tour_plan_form"):
            t1, t2 = st.columns(2)
            with t1:
                ind_name = st.text_input("Name of Organization / Industry*", value="  ")
                focal_person = st.text_input("Industry Focal Person Name*", value=" ")
                focal_contact = st.text_input("Focal Person Contact No*", value=" ")
                students_count = st.number_input("Number of Visiting Students*", min_value=1, value=max(len(tour_students_allocated_df), 40), step=1)
            with t2:
                tour_date = st.date_input("Date of Visit")
                dep_time = st.text_input("Time of Departure*", value="09:00 AM")
                travel_mode = st.selectbox("Mode of Travel*", ["Bus", "Van", "Car", "Train", "Other Mode"])
                other_mode = st.text_input("If Other Mode, Specify", value="")
                approval_copy = st.selectbox("Copy of Approval letter from Industry", ["Yes", "No"])

            st.markdown("##### Accompanying Staff (Staff Deputed for Tour)")
            s1, s2 = st.columns(2)
            with s1:
                driver_name = st.text_input("College Bus Driver Name*", value="Muhammad Tahir")
                conductor_name = st.text_input("College Bus Conductor Name*", value="Shahbaz Sultan")
            with s2:
                aipo_name = st.text_input("Accompanying AIPO / Trade Teacher*", value=" ")
                assistant_name = st.text_input("Accompanying Shop Assistant / Attendant*", value="Shop Assistant (Trade)")

            tour_limits_json = json.dumps(tour_custom_counts)
            save_tour = st.form_submit_button("Save Industrial Tour Plan")
            
            if save_tour:
                if ind_name and focal_person and tour_target_techs:
                    primary_tech = ", ".join(tour_target_techs)
                    primary_tech = ", ".join(tour_target_techs)
                    tour_app_payload = {
                        'industry_name': ind_name.strip(), 'technology': primary_tech,
                        'session': tour_target_session.strip(), 'tour_date': str(tour_date),
                        'dep_time': dep_time.strip(), 'travel_mode': travel_mode,
                        'students_count': students_count, 'driver_name': driver_name.strip(),
                        'conductor_name': conductor_name.strip(), 'aipo_name': aipo_name.strip(),
                        'assistant_name': assistant_name.strip()
                    }
                    tour_app_blob = compress_file_bytes(generate_tour_approval_application(tour_app_payload).getvalue())
                    cursor.execute("""
                    INSERT INTO industry_visits (
                        industry_name, focal_person, focal_contact, technology, session,
                        tour_date, dep_time, travel_mode, other_mode, students_count,
                        driver_name, conductor_name, aipo_name, assistant_name, approval_copy,
                        rep_name, rep_designation, h_punctuality, h_discipline, h_interest,
                        h_tech_know, h_comm, h_safety, h_recommendation, h_comments,
                        q1_orientation, q2_machinery, q_relevant, q_objectives,
                        r_discipline, r_tech, r_prof, r_satisfaction,
                        candidate_limits_json, approval_doc_blob, budget_utilized, status
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        ind_name.strip(), focal_person.strip(), focal_contact.strip(), primary_tech, tour_target_session.strip(),
                        str(tour_date), dep_time.strip(), travel_mode, other_mode.strip(), students_count,
                        driver_name.strip(), conductor_name.strip(), aipo_name.strip(), assistant_name.strip(), approval_copy,
                        focal_person.strip(), "Industrial Representative", "Excellent", "Excellent", "Excellent", "Good", "Good", "Excellent",
                        "Yes", "Good professional behavior.", "YES", "YES", "YES", "YES", 5, 5, 5, 5,
                        tour_limits_json, tour_app_blob, 0.0, "Planned"
                    ))
                    conn.commit()
                    st.success("Industrial Tour plan save ho gaya!")
                else:
                    st.warning("Industry Name, Focal Person aur Target Technologies select karna lazmi hain.")

        if ind_name and focal_person and tour_target_techs:
            tour_data = {
                'industry_name': ind_name, 'technology': ", ".join(tour_target_techs), 'session': tour_target_session,
                'tour_date': str(tour_date), 'dep_time': dep_time, 'travel_mode': travel_mode,
                'students_count': students_count, 'driver_name': driver_name, 'conductor_name': conductor_name,
                'aipo_name': aipo_name, 'assistant_name': assistant_name, 'focal_person': focal_person,
                'focal_contact': focal_contact, 'other_mode': other_mode, 'approval_copy': approval_copy
            }
            c_app, c_ann_g = st.columns(2)
            with c_app:
                tour_app_docx = generate_tour_approval_application(tour_data)
                st.download_button(
                    label="Download Principal Tour Approval Application (.docx)",
                    data=tour_app_docx,
                    file_name=f"Tour_Approval_{ind_name.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            with c_ann_g:
                ann_g_live = generate_annexure_g(tour_data, tour_students_allocated_df)
                st.download_button(
                    label=f"📄 Download Planned Annexure ‘G’ ({len(tour_students_allocated_df)} Students) (.docx)",
                    data=ann_g_live,
                    file_name=f"Annexure_G_{ind_name.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

    with tour_tab2:
        df_visits = pd.read_sql("SELECT id, industry_name, technology, session, tour_date, status FROM industry_visits ORDER BY id DESC", conn)
        
        if not df_visits.empty:
            visit_map = {f"ID {r['id']} - {r['industry_name']} ({r['technology']}) [{r['status']}]": r['id'] for _, r in df_visits.iterrows()}
            chosen_label = st.selectbox("Select Planned Visit for Feedback & Official Forms:", list(visit_map.keys()))
            chosen_id = visit_map[chosen_label]

            cursor.execute("SELECT * FROM industry_visits WHERE id = ?", (chosen_id,))
            v_row = cursor.fetchone()
            v_cols = [d[0] for d in cursor.description]
            v_data = dict(zip(v_cols, v_row))

            st.markdown(f"##### 🏭 Active Tour: **{v_data['industry_name']}** | Class: **DAE {v_data['technology']} ({v_data['session']})** | Date: **{v_data['tour_date']}**")
            st.info(f"Focal Person Planned in Tab 1: **{v_data['focal_person']}** ({v_data['focal_contact']}) | Mode: **{v_data['travel_mode']}**")

            # Load actual students belonging to the technologies stored in this visit
            stored_limits = {}
            if v_data.get('candidate_limits_json'):
                try:
                    stored_limits = json.loads(v_data['candidate_limits_json'])
                except:
                    pass

            v_techs = [x.strip() for x in v_data['technology'].split(",") if x.strip()]
            active_tour_stds = []
            for vt in v_techs:
                cnt_l = stored_limits.get(vt)
                q_vt = "SELECT roll_no AS 'Roll No', name AS 'Name', technology AS 'Technology', cnic AS 'CNIC', phone AS 'Mobile Number', email AS 'Email' FROM students WHERE session = ? AND technology = ? ORDER BY roll_no"
                if cnt_l and cnt_l > 0:
                    q_vt += f" LIMIT {cnt_l}"
                df_vt_sub = pd.read_sql(q_vt, conn, params=[v_data['session'], vt])
                if not df_vt_sub.empty:
                    active_tour_stds.append(df_vt_sub)

            tour_students_df = pd.concat(active_tour_stds, ignore_index=True) if active_tour_stds else pd.DataFrame(columns=['Roll No', 'Name', 'Technology', 'CNIC', 'Mobile Number', 'Email'])
            
            annex_g_docx = generate_annexure_g(v_data, tour_students_df)
            st.download_button(
                label=f"📄 Download Exact Annexure ‘G’ ({len(tour_students_df)} Trainees) (.docx)",
                data=annex_g_docx,
                file_name=f"Annexure_G_{v_data['industry_name'].replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

            st.markdown("---")
            st.markdown("#### 📝 Live Evaluation & Feedback Controls (Annexure ‘H’ & ‘I’)")
            st.caption("Aap neeche jo bhi select karenge, wahi fresh inputs download hone wali Word files mein aayenge:")

            with st.container():
                st.markdown("##### 🏢 Section A: Annexure ‘H’ (Industry Representative Feedback)")
                h_c1, h_c2 = st.columns(2)
                with h_c1:
                    default_rep_name = v_data['rep_name'] if (v_data.get('rep_name') and str(v_data['rep_name']).strip()) else (v_data['focal_person'] or "Industrial Representative")
                    rep_n = st.text_input("Industrial Representative Name*", value=default_rep_name, key=f"dyn_rep_name_{chosen_id}")
                    rep_d = st.text_input("Designation*", value=v_data['rep_designation'] or "Plant / HR Manager", key=f"dyn_rep_desig_{chosen_id}")
                    
                    punc_opts = ["Excellent", "Good", "Needs Improvement"]
                    punc_idx = punc_opts.index(v_data['h_punctuality']) if v_data.get('h_punctuality') in punc_opts else 0
                    punc = st.radio("Punctuality and Time Management", punc_opts, index=punc_idx, key=f"dyn_punc_{chosen_id}")
                    
                    disc_opts = ["Excellent", "Good", "Needs Improvement"]
                    disc_idx = disc_opts.index(v_data['h_discipline']) if v_data.get('h_discipline') in disc_opts else 0
                    disc = st.radio("Discipline and Behavior", disc_opts, index=disc_idx, key=f"dyn_disc_{chosen_id}")
                    
                    inte_opts = ["Excellent", "Good", "Needs Improvement"]
                    inte_idx = inte_opts.index(v_data['h_interest']) if v_data.get('h_interest') in inte_opts else 0
                    inte = st.radio("Interest and Participation Level", inte_opts, index=inte_idx, key=f"dyn_inte_{chosen_id}")

                with h_c2:
                    tech_k_opts = ["Excellent", "Good", "Needs Improvement"]
                    tech_k_idx = tech_k_opts.index(v_data['h_tech_know']) if v_data.get('h_tech_know') in tech_k_opts else 1
                    tech_k = st.radio("Technical Knowledge and Understanding", tech_k_opts, index=tech_k_idx, key=f"dyn_tech_k_{chosen_id}")
                    
                    comm_s_opts = ["Excellent", "Good", "Needs Improvement"]
                    comm_s_idx = comm_s_opts.index(v_data['h_comm']) if v_data.get('h_comm') in comm_s_opts else 1
                    comm_s = st.radio("Communication and Interpersonal Skills", comm_s_opts, index=comm_s_idx, key=f"dyn_comm_s_{chosen_id}")
                    
                    safe_a_opts = ["Excellent", "Good", "Needs Improvement"]
                    safe_a_idx = safe_a_opts.index(v_data['h_safety']) if v_data.get('h_safety') in safe_a_opts else 0
                    safe_a = st.radio("Safety Awareness and Compliance", safe_a_opts, index=safe_a_idx, key=f"dyn_safe_a_{chosen_id}")
                    
                    recom_opts = ["Yes", "No", "With Improvements"]
                    recom_idx = recom_opts.index(v_data['h_recommendation']) if v_data.get('h_recommendation') in recom_opts else 0
                    recom = st.radio("Would you recommend future visits from TEVTA students?", recom_opts, index=recom_idx, key=f"dyn_recom_{chosen_id}")
                    
                    comm = st.text_area("General Comments / Suggestions", value=v_data['h_comments'] or "The students demonstrated strong disciplinary conduct and actively engaged with industrial operations.", key=f"dyn_comm_{chosen_id}")

                st.markdown("---")
                st.markdown("##### 📋 Section B: Annexure ‘I’ (All 8 Feedback Points by IPO / Career Counsellor)")
                i_c1, i_c2 = st.columns(2)
                with i_c1:
                    q_rel = st.selectbox("1. The Industrial visit was relevant to field of study?", ["YES", "NO"], index=0 if v_data.get('q_relevant', 'YES') == 'YES' else 1, key=f"dyn_q_rel_{chosen_id}")
                    q1 = st.selectbox("2. Pre visit orientation conducted for the trainees?", ["YES", "NO"], index=0 if v_data.get('q1_orientation', 'YES') == 'YES' else 1, key=f"dyn_q1_{chosen_id}")
                    q_obj = st.selectbox("3. Objectives of industry visit were achieved?", ["YES", "NO"], index=0 if v_data.get('q_objectives', 'YES') == 'YES' else 1, key=f"dyn_q_obj_{chosen_id}")
                    q2 = st.selectbox("4. The working on machinery explained by industry rep / IPO / teacher?", ["YES", "NO"], index=0 if v_data.get('q2_machinery', 'YES') == 'YES' else 1, key=f"dyn_q2_{chosen_id}")

                with i_c2:
                    r_disc = st.select_slider("5. Students Discipline during industry (1-5, 5 is best)", options=[1, 2, 3, 4, 5], value=int(v_data.get('r_discipline') or 5), key=f"dyn_r_disc_{chosen_id}")
                    r_tech = st.select_slider("6. Students gain technical exposure? (1-5, 5 is best)", options=[1, 2, 3, 4, 5], value=int(v_data.get('r_tech') or 5), key=f"dyn_r_tech_{chosen_id}")
                    r_prof = st.select_slider("7. Students gain Professional Exposure (1-5, 5 is best)", options=[1, 2, 3, 4, 5], value=int(v_data.get('r_prof') or 5), key=f"dyn_r_prof_{chosen_id}")
                    r_sat = st.select_slider("8. Overall satisfaction level after visit (1-5, 5 is best)", options=[1, 2, 3, 4, 5], value=int(v_data.get('r_satisfaction') or 5), key=f"dyn_r_sat_{chosen_id}")

                st.markdown("---")
                tour_budget = st.number_input("Budget Utilized for this Activity (PKR)*", min_value=0.0, value=float(v_data['budget_utilized'] or 30000.0), step=1000.0, key=f"dyn_tour_budget_{chosen_id}")

                if st.button("💾 Permanent Archive & Update Tour Record in Database", key=f"save_btn_tour_{chosen_id}"):
                    cursor.execute("""
                    UPDATE industry_visits SET
                        rep_name = ?, rep_designation = ?, h_punctuality = ?, h_discipline = ?,
                        h_interest = ?, h_tech_know = ?, h_comm = ?, h_safety = ?,
                        h_recommendation = ?, h_comments = ?, 
                        q1_orientation = ?, q2_machinery = ?, q_relevant = ?, q_objectives = ?,
                        r_discipline = ?, r_tech = ?, r_prof = ?, r_satisfaction = ?,
                        budget_utilized = ?, status = 'Completed'
                    WHERE id = ?
                    """, (
                        rep_n.strip(), rep_d.strip(), punc, disc, inte, tech_k, comm_s, safe_a, recom, comm.strip(),
                        q1, q2, q_rel, q_obj,
                        r_disc, r_tech, r_prof, r_sat, tour_budget, chosen_id
                    ))
                    conn.commit()
                    st.success(f"Industrial Tour Mukammal Save ho kar Archive ho gaya! (Budget Utilized: PKR {tour_budget:,.2f})")

            live_tour_payload = v_data.copy()
            live_tour_payload.update({
                'rep_name': rep_n,
                'rep_designation': rep_d,
                'h_punctuality': punc,
                'h_discipline': disc,
                'h_interest': inte,
                'h_tech_know': tech_k,
                'h_comm': comm_s,
                'h_safety': safe_a,
                'h_recommendation': recom,
                'h_comments': comm,
                'q_relevant': q_rel,
                'q1_orientation': q1,
                'q_objectives': q_obj,
                'q2_machinery': q2,
                'r_discipline': r_disc,
                'r_tech': r_tech,
                'r_prof': r_prof,
                'r_satisfaction': r_sat,
                'budget_utilized': tour_budget
            })

            st.markdown("##### 📥 Download Dynamic Evaluated Forms")
            d1, d2 = st.columns(2)
            with d1:
                annex_h_docx = generate_annexure_h(live_tour_payload)
                st.download_button("📄 Download Annexure ‘H’ (With Selected Ratings) (.docx)", annex_h_docx, f"Annexure_H_{v_data['industry_name'].replace(' ', '_')}.docx")
            with d2:
                annex_i_docx = generate_annexure_i(live_tour_payload)
                st.download_button("📄 Download 8-Point Annexure ‘I’ (With Selected Ratings) (.docx)", annex_i_docx, f"Annexure_I_{v_data['industry_name'].replace(' ', '_')}.docx")
        else:
            st.info("Pehle Tab 1 se industrial tour plan karein.")

    with tour_tab3:
        st.markdown("##### 📂 Permanent Archived Tour Records, Edit & Delete")
        df_all_tours = pd.read_sql("SELECT id, industry_name, technology, session, tour_date, students_count, budget_utilized, status FROM industry_visits ORDER BY id DESC", conn)
        if not df_all_tours.empty:
            st.dataframe(df_all_tours, use_container_width=True)
            
            st.markdown("---")
            st.markdown("###### ⚙️ Manage Archived Tour (Re-download / Edit / Delete)")
            re_tid = st.selectbox("Archived Tour Select Karein:", df_all_tours['id'].tolist(), format_func=lambda x: f"ID {x} - {df_all_tours.loc[df_all_tours['id']==x, 'industry_name'].values[0]} ({df_all_tours.loc[df_all_tours['id']==x, 'tour_date'].values[0]})")
            cursor.execute("SELECT * FROM industry_visits WHERE id = ?", (re_tid,))
            archive_v = dict(zip([d[0] for d in cursor.description], cursor.fetchone()))

            arch_limits = {}
            if archive_v.get('candidate_limits_json'):
                try:
                    arch_limits = json.loads(archive_v['candidate_limits_json'])
                except:
                    pass

            a_techs = [x.strip() for x in archive_v['technology'].split(",") if x.strip()]
            arch_tour_stds = []
            for at in a_techs:
                cnt_al = arch_limits.get(at)
                q_at = "SELECT roll_no AS 'Roll No', name AS 'Name', technology AS 'Technology', cnic AS 'CNIC', phone AS 'Mobile Number', email AS 'Email' FROM students WHERE session = ? AND technology = ? ORDER BY roll_no"
                if cnt_al and cnt_al > 0:
                    q_at += f" LIMIT {cnt_al}"
                df_at_sub = pd.read_sql(q_at, conn, params=[archive_v['session'], at])
                if not df_at_sub.empty:
                    arch_tour_stds.append(df_at_sub)

            std_df = pd.concat(arch_tour_stds, ignore_index=True) if arch_tour_stds else pd.DataFrame(columns=['Roll No', 'Name', 'Technology', 'CNIC', 'Mobile Number', 'Email'])

            if archive_v.get('approval_doc_blob'):
                re_tour_app_bytes = decompress_file_bytes(archive_v['approval_doc_blob'])
            else:
                re_tour_app_bytes = generate_tour_approval_application(archive_v).getvalue()

            c_app_t, c_g, c_h, c_i = st.columns(4)
            with c_app_t:
                st.download_button("📥 Principal Approval", re_tour_app_bytes, f"Archived_Tour_Approval_{re_tid}.docx")
            with c_g:
                st.download_button("📥 Annexure G", generate_annexure_g(archive_v, std_df), f"Archived_Annex_G_{re_tid}.docx")
            with c_h:
                st.download_button("📥 Annexure H", generate_annexure_h(archive_v), f"Archived_Annex_H_{re_tid}.docx")
            with c_i:
                st.download_button("📥 Annexure I", generate_annexure_i(archive_v), f"Archived_Annex_I_{re_tid}.docx")

            with st.expander(f"✏️ Edit / Correct Tour Entry (ID: {re_tid})"):
                with st.form(f"edit_tour_{re_tid}"):
                    e_ind = st.text_input("Industry Name", value=archive_v['industry_name'])
                    e_c1, e_c2 = st.columns(2)
                    with e_c1:
                        e_cnt = st.number_input("Students Count", min_value=0, value=int(archive_v['students_count'] or 0))
                        e_bud_t = st.number_input("Budget Utilized (PKR)", min_value=0.0, value=float(archive_v['budget_utilized'] or 0.0), step=1000.0)
                    with e_c2:
                        e_stat_t = st.selectbox("Status", ["Completed", "Planned"], index=0 if archive_v['status'] == 'Completed' else 1)
                        e_rep = st.text_input("Industry Rep Name", value=str(archive_v['rep_name'] or ''))
                    
                    if st.form_submit_button("💾 Save Corrected Tour Data"):
                        cursor.execute("""
                            UPDATE industry_visits 
                            SET industry_name = ?, students_count = ?, budget_utilized = ?, status = ?, rep_name = ?
                            WHERE id = ?
                        """, (e_ind, e_cnt, e_bud_t, e_stat_t, e_rep, re_tid))
                        conn.commit()
                        st.success("Tour record update ho gaya! Refresh karein.")
                        st.rerun()

            with st.expander(f"🗑️ Delete Record (ID: {re_tid})"):
                st.warning("⚠️ Kya aap waqai is tour record ko permanently delete karna chahte hain?")
                confirm_del_t = st.checkbox("Haan, main is tour ko delete karna chahta hoon", key=f"del_chk_t_{re_tid}")
                if st.button("🔴 Delete This Industrial Tour", key=f"del_btn_t_{re_tid}"):
                    if confirm_del_t:
                        cursor.execute("DELETE FROM industry_visits WHERE id = ?", (re_tid,))
                        conn.commit()
                        st.success(f"Tour Record ID {re_tid} delete ho chuka hai.")
                        st.rerun()
                    else:
                        st.error("Pehle confirmation check karein.")
        else:
            st.write("Abhi tak koi industrial visit record nahi mila.")

            # =========================================================================
# =========================================================================
# --- 6. Recruitment Drive ---
# =========================================================================
elif choice == "Recruitment Drive":
    st.subheader("On-Campus Recruitment Drive Management & Documentation")
    
    # Safe Database Column Alterations
    for col_rd in [("participated_cnics_json", "TEXT"), ("selected_cnics_json", "TEXT")]:
        try:
            cursor.execute(f"ALTER TABLE recruitment_drives ADD COLUMN IF NOT EXISTS {col_rd[0]} {col_rd[1]}")
            conn.commit()
        except Exception:
            try:
                if hasattr(conn, "rollback"):
                    conn.rollback()
                cursor.execute(f"ALTER TABLE recruitment_drives ADD COLUMN {col_rd[0]} {col_rd[1]}")
                conn.commit()
            except Exception:
                if hasattr(conn, "rollback"):
                    conn.rollback()

    rd_tab1, rd_tab2, rd_tab3 = st.tabs([
        "1. Plan Drive, Candidate Selection & Annexure 'J'", 
        "2. Assessment, Feedback, Budget & Final Submit", 
        "3. Permanent Archive, Edit & Delete"
    ])

    pl = "?"

    # -------------------------------------------------------------------------
    # TAB 1: PLAN DRIVE, CANDIDATE SELECTION & ANNEXURE J
    # -------------------------------------------------------------------------
    with rd_tab1:
        st.markdown("##### Pre-Drive: Company Profile, Job Details & Administrative Approvals")
        
        st.markdown("###### 🎯 Targeted Technologies & Session (Live Selection)")
        sel_c1, sel_c2 = st.columns(2)
        with sel_c1:
            target_techs = st.multiselect(
                "Select Eligible Technologies / Trades*",
                TECHNOLOGIES_LIST,
                default=["Electrical", "Mechanical"]
            )
        with sel_c2:
            target_session = st.text_input("Target Session (e.g. 2023-2026)*", value="2023-2026")

        if not target_techs:
            st.warning("⚠️ Barah-e-karam kam az kam aik Technology lazmi select karein.")

        st.markdown("###### 👥 Technology-Wise Candidate Count & Live Availability")
        selection_mode = st.radio(
            "Selection Type:",
            ["Tamam Matching Students (All in Selected Technologies)", "Custom Tadaad (Specify Count Per Technology)"],
            horizontal=True
        )

        custom_counts = {}
        all_eligible_list = []

        if selection_mode == "Custom Tadaad (Specify Count Per Technology)" and target_techs:
            st.caption("Neeche har select shuda technology ka live count check karein aur jitne students chahiyein utne enter karein:")
            num_cols = min(len(target_techs), 4) if len(target_techs) > 0 else 1
            t_cols = st.columns(num_cols)
            
            for idx, t in enumerate(target_techs):
                with t_cols[idx % num_cols]:
                    cursor.execute(f"SELECT COUNT(*) FROM students WHERE session = {pl} AND technology = {pl}", (target_session, t))
                    avail = cursor.fetchone()[0] or 0
                    st.markdown(f"**🔧 {t}**")
                    cnt = st.number_input(
                        f"DB Available: {avail}",
                        min_value=0,
                        max_value=max(avail, 1000),
                        value=min(avail, 20),
                        step=5,
                        key=f"tech_count_{t}"
                    )
                    custom_counts[t] = cnt
                    if cnt > 0:
                        q_t = f"SELECT roll_no AS 'Roll No', name AS 'Name', technology AS 'Technology', cnic AS 'CNIC', phone AS 'Mobile Number', email AS 'Email' FROM students WHERE session = {pl} AND technology = {pl} ORDER BY roll_no LIMIT {cnt}"
                        df_sub = pd.read_sql(q_t, conn, params=[target_session, t])
                        if not df_sub.empty:
                            all_eligible_list.append(df_sub)
        else:
            for t in target_techs:
                cursor.execute(f"SELECT COUNT(*) FROM students WHERE session = {pl} AND technology = {pl}", (target_session, t))
                avail = cursor.fetchone()[0] or 0
                custom_counts[t] = avail
                q_t = f"SELECT roll_no AS 'Roll No', name AS 'Name', technology AS 'Technology', cnic AS 'CNIC', phone AS 'Mobile Number', email AS 'Email' FROM students WHERE session = {pl} AND technology = {pl} ORDER BY roll_no"
                df_sub = pd.read_sql(q_t, conn, params=[target_session, t])
                if not df_sub.empty:
                    all_eligible_list.append(df_sub)

        if all_eligible_list:
            drive_students_df = pd.concat(all_eligible_list, ignore_index=True)
        else:
            drive_students_df = pd.DataFrame(columns=['Roll No', 'Name', 'Technology', 'CNIC', 'Mobile Number', 'Email'])

        st.info(f"📊 Total Candidates Selected: **{len(drive_students_df)} Students**")

        st.markdown("---")
        with st.form("plan_recruitment_form"):
            st.markdown("###### 🏢 Company & Job Details")
            r1, r2 = st.columns(2)
            with r1:
                comp_name = st.text_input("Company Name*", value="   ")
                comp_address = st.text_input("Company Address*", value=" ")
                job_loc = st.text_input("Job Site Location*", value=" ")
                job_title = st.text_input("Job Title / Designation*", value="Associate / Trainee Technician")
                exp_req = st.selectbox("Required Experience*", ["Fresh / Result Awaited", "Experienced", "Result Awaited"])
                vac_m = st.number_input("No. of Vacancies (Male)", min_value=0, value=0, step=5)
                vac_f = st.number_input("No. of Vacancies (Female)", min_value=0, value=0, step=1)
            with r2:
                marks_div = st.text_input("Required Marks / Division", value="1st Division (Minimum 60%)")
                salary_range = st.text_input("Salary Range (PKR / month)*", value="35,000 - 45,000")
                written_test = st.selectbox("Written Test / Screening to be conducted?", ["Yes", "No"])
                job_desc = st.text_area("Job Description / Responsibilities", value="Assembly line operations, troubleshooting, and production shift maintenance.")

            st.markdown("###### 📋 Terms, Perks & Facilities")
            p1, p2, p3 = st.columns(3)
            with p1:
                rec_mode = st.selectbox("Initial Recruitment Mode*", ["Permanent", "Contractual", "Apprentices", "Internship / OJT"])
            with p2:
                eobi_val = st.checkbox("EOBI Act 1976 Included", value=True)
                soc_val = st.checkbox("Social Security Ordinance 1965", value=True)
                bon_val = st.checkbox("Bonuses / Performance Incentives", value=True)
            with p3:
                acc_mode = st.selectbox("Accommodation Available*", [
                    "Company maintained accommodation",
                    "Company paid / subsidized accommodation",
                    "Any other benefit (Food etc.)"
                ])

            st.markdown("###### 📅 Event Schedule & Focal Persons")
            e1, e2 = st.columns(2)
            with e1:
                d_venue = st.text_input("Venue*", value="Main Auditorium & Lab, GCT Samanabad")
                d_date = st.date_input("Recruitment Drive Date")
                d_time = st.text_input("Drive Time*", value="10:00 AM")
                dpo_name = st.text_input("District Placement Officer (DPO) Name*", value="Mr. Muhammad Asghar Ch (DPO)")
                ipo_name = st.text_input("Institute Placement Officer (IPO) Name*", value="Engr. Muhammad Wasif Ali (IPO)")
                tevta_contact = st.text_input("TEVTA Focal Contact Number*", value="+92 346 4843168")
            with e2:
                emp_focal = st.text_input("Employer Focal Person Name*", value=" ")
                emp_desig = st.text_input("Employer Focal Designation*", value="HR Manager")
                emp_contact = st.text_input("Employer Focal Contact No*", value=" ")
                emp_email = st.text_input("Employer Email Address", value=" ")

            st.markdown("###### ☕ Administrative & Protocol Sanctions")
            ap1, ap2, ap3 = st.columns(3)
            with ap1:
                ref_heads = st.number_input("Refreshment Sanction (Persons)*", min_value=0, value=15, step=1)
            with ap2:
                backdrop_choice = st.selectbox("Stage Backdrop Flex", ["Stage Backdrop Flex Required", "Digital Presentation Screen", "Not Required"])
            with ap3:
                bouquet_choice = st.selectbox("Welcome Bouquet / Souvenir", ["Bouquets & Souvenir Shields Required", "Bouquet Only", "Shield Only", "Not Required"])

            limits_json_str = json.dumps(custom_counts)
            save_drive_btn = st.form_submit_button("Save Recruitment Drive Plan")
            
            if save_drive_btn:
                if comp_name and emp_focal and target_techs:
                    drive_app_payload = {
                        'company_name': comp_name, 'company_address': comp_address,
                        'drive_date': str(d_date), 'drive_time': d_time, 'venue': d_venue,
                        'written_test': written_test, 'target_techs': target_techs,
                        'session': target_session, 'salary_range': salary_range,
                        'refreshment_heads': ref_heads, 'backdrop_details': backdrop_choice,
                        'bouquet_details': bouquet_choice
                    }
                    drive_app_blob = compress_file_bytes(generate_drive_principal_application(drive_app_payload).getvalue())
                    ph_insert = ",".join([pl] * 46)
                    insert_sql = f"""
                    INSERT INTO recruitment_drives (
                        company_name, company_address, job_location, job_title, experience,
                        vacancies_male, vacancies_female, marks_div, salary_range, job_desc,
                        recruitment_mode, eobi, social_sec, bonuses, accommodation,
                        venue, drive_date, drive_time, written_test, target_techs, session,
                        refreshment_heads, backdrop_details, bouquet_details,
                        dpo_name, ipo_name, tevta_contact, emp_focal, emp_desig, emp_contact, emp_email,
                        students_appeared, students_shortlisted,
                        l_coordination, l_staff, l_venue_comfort, l_tech_skills, l_resumes, l_comm_skills, l_dress, l_overall, l_suggestions,
                        candidate_limits_json, approval_doc_blob, budget_utilized, status
                    ) VALUES ({ph_insert})
                    """
                    cursor.execute(insert_sql, (
                        comp_name, comp_address, job_loc, job_title, exp_req,
                        vac_m, vac_f, marks_div, salary_range, job_desc,
                        rec_mode, 1 if eobi_val else 0, 1 if soc_val else 0, 1 if bon_val else 0, acc_mode,
                        d_venue, str(d_date), d_time, written_test, ",".join(target_techs), target_session,
                        ref_heads, backdrop_choice, bouquet_choice,
                        dpo_name, ipo_name, tevta_contact, emp_focal, emp_desig, emp_contact, emp_email,
                        0, 0,
                        "E", "E", "E", "E", "G", "G", "E", "E", "Drive conducted smoothly.",
                        limits_json_str, drive_app_blob, 0.0, "Planned"
                    ))
                    conn.commit()
                    st.success("Recruitment Drive plan kamyabi se save ho gaya!")
                    st.rerun()
                else:
                    st.warning("Company Name, Employer Focal Person aur Target Technologies select karna lazmi hai.")

        if comp_name and emp_focal and target_techs:
            d_data = {
                'company_name': comp_name, 'company_address': comp_address, 'job_location': job_loc,
                'job_title': job_title, 'experience': exp_req, 'vacancies_male': vac_m, 'vacancies_female': vac_f,
                'marks_div': marks_div, 'salary_range': salary_range, 'job_desc': job_desc,
                'recruitment_mode': rec_mode, 'eobi': eobi_val, 'social_sec': soc_val, 'bonuses': bon_val,
                'accommodation': acc_mode, 'venue': d_venue, 'drive_date': str(d_date), 'drive_time': d_time,
                'written_test': written_test, 'target_techs': target_techs, 'session': target_session,
                'refreshment_heads': ref_heads, 'backdrop_details': backdrop_choice, 'bouquet_details': bouquet_choice,
                'dpo_name': dpo_name, 'ipo_name': ipo_name, 'tevta_contact': tevta_contact,
                'emp_focal': emp_focal, 'emp_desig': emp_desig, 'emp_contact': emp_contact, 'emp_email': emp_email
            }

            st.markdown("---")
            st.markdown("##### 📥 Download All Approved Drive Documents (Annexures J, K, L & Roster)")
            b1, b2, b3 = st.columns(3)
            with b1:
                app_drive_docx = generate_drive_principal_application(d_data)
                st.download_button("📄 Principal Approval Letter (.docx)", app_drive_docx, f"Drive_Approval_{comp_name.replace(' ', '_')}.docx")

            with b2:
                annex_j_docx = generate_annexure_j(d_data, drive_students_df)
                st.download_button(f"📄 Annexure “J” ({len(drive_students_df)} Students) (.docx)", annex_j_docx, f"Annexure_J_{comp_name.replace(' ', '_')}.docx")

            with b3:
                roster_excel = generate_candidate_roster_excel(d_data, drive_students_df)
                st.download_button(f"📊 Candidates Roster ({len(drive_students_df)} Students) (.xlsx)", roster_excel, f"Candidates_List_{comp_name.replace(' ', '_')}.xlsx")

            b4, b5 = st.columns(2)
            with b4:
                annex_k_tab1 = generate_annexure_k_excel(d_data, drive_students_df)
                st.download_button(f"📊 Annexure “K” Assessment Sheet ({len(drive_students_df)} Candidates) (.xlsx)", annex_k_tab1, f"Annexure_K_{comp_name.replace(' ', '_')}.xlsx")

            with b5:
                annex_l_tab1 = generate_annexure_l(d_data)
                st.download_button("📄 Annexure “L” Employer Feedback (.docx)", annex_l_tab1, f"Annexure_L_{comp_name.replace(' ', '_')}.docx")

    # -------------------------------------------------------------------------
    # TAB 2: ASSESSMENT, FEEDBACK, BUDGET & FINAL SUBMIT (ANNEX K OPTIONS)
    # -------------------------------------------------------------------------
    with rd_tab2:
        df_drives = pd.read_sql("SELECT id, company_name, drive_date, session, target_techs, venue, drive_time, status FROM recruitment_drives ORDER BY id DESC", conn)
        
        if not df_drives.empty:
            drive_map = {f"ID {r['id']} - {r['company_name']} ({r['drive_date']}) [{r['status']}]": r['id'] for _, r in df_drives.iterrows()}
            selected_drive_label = st.selectbox("Select Recruitment Drive to Finalize:", list(drive_map.keys()))
            selected_drive_id = drive_map[selected_drive_label]

            cursor.execute(f"SELECT * FROM recruitment_drives WHERE id = {pl}", (selected_drive_id,))
            dr_row = cursor.fetchone()
            dr_cols = [d[0] for d in cursor.description]
            dr_data = dict(zip(dr_cols, dr_row))
            dr_data['target_techs'] = dr_data['target_techs'].split(",") if dr_data['target_techs'] else []

            st.write(f"**Company:** {dr_data['company_name']} | **Date:** {dr_data['drive_date']} | **Trades:** {', '.join(dr_data['target_techs'])}")

            st.markdown("---")
            st.markdown("#### 📝 Annexure “K” Assessment Sheet Setup")
            st.caption("Participated candidates ka source aur input method select karein:")

            saved_part_cnics = []
            if dr_data.get('participated_cnics_json'):
                try:
                    saved_part_cnics = json.loads(dr_data['participated_cnics_json'])
                except Exception:
                    saved_part_cnics = []

            # --- MAIN 2 CONDITIONS ---
            k_mode_choice = st.radio(
                "Assessment Sheet Candidate Data Source:",
                [
                    "🔍 Option 1: Strictly Search From Database",
                    "📂 Option 2: Search Database + Fallback to Temporary Data File (Combined/Merge)"
                ],
                horizontal=False,
                key=f"k_main_mode_choice_{selected_drive_id}"
            )

            target_raw_cnics = []
            cand_df = pd.DataFrame(columns=['Roll No', 'Name', 'Technology', 'Session', 'CNIC'])

            # =========================================================================
            # OPTION 1: STRICTLY DATABASE SEARCH (PASTE YA EXCEL UPLOAD)
            # =========================================================================
            if k_mode_choice == "🔍 Option 1: Strictly Search From Database":
                st.info("💡 Yeh mode sirf unhi students ko layega jo pehle se Database mein mojood hain.")
                
                in_tab1, in_tab2 = st.tabs(["✍️ Direct Paste CNICs", "📁 Upload Target CNICs File (.xlsx / .csv)"])
                
                with in_tab1:
                    default_paste = "\n".join(saved_part_cnics) if saved_part_cnics else ""
                    pasted_val = st.text_area(
                        "CNICs Paste karein (Har line par aik CNIC):",
                        value=default_paste,
                        placeholder="33100-1234567-1\n33100-7654321-2",
                        key=f"opt1_paste_area_{selected_drive_id}",
                        height=120
                    )
                    if pasted_val.strip():
                        target_raw_cnics = re.split(r"[\n,\s;]+", pasted_val.strip())
                    elif saved_part_cnics:
                        target_raw_cnics = saved_part_cnics

                with in_tab2:
                    up_file1 = st.file_uploader(
                        "CNICs wali File Upload Karein (.xlsx / .csv):",
                        type=["xlsx", "csv"],
                        key=f"opt1_file_up_{selected_drive_id}"
                    )
                    if up_file1 is not None:
                        try:
                            df_in1 = pd.read_csv(up_file1, dtype=str) if up_file1.name.lower().endswith(".csv") else pd.read_excel(up_file1, dtype=str)
                            col_c1 = next((c for c in df_in1.columns if "cnic" in str(c).lower()), None)
                            if not col_c1:
                                col_c1 = st.selectbox("Sheet mein se CNIC column select karein:", df_in1.columns, key=f"opt1_col_sel_{selected_drive_id}")
                            if col_c1:
                                target_raw_cnics = df_in1[col_c1].dropna().tolist()
                        except Exception as e:
                            st.error(f"File parhne mein masla aya: {e}")

                if target_raw_cnics:
                    target_dashed, lookup_cnics = parse_cnic_input(target_raw_cnics)
                    if lookup_cnics:
                        placeholders_k = ",".join([pl] * len(lookup_cnics))
                        sql_k = f"""
                            SELECT roll_no AS 'Roll No', name AS 'Name', technology AS 'Technology', session AS 'Session', cnic AS 'CNIC'
                            FROM students 
                            WHERE cnic IN ({placeholders_k})
                            ORDER BY technology, roll_no
                        """
                        cand_df = pd.read_sql(sql_k, conn, params=lookup_cnics)
                        found_k = set(cand_df['CNIC'].tolist())
                        unmatched_k = [c for c in target_dashed if c not in found_k]

                        st.success(f"✅ Database se **{len(cand_df)}** Valid Candidates mil gaye.")
                        if unmatched_k:
                            st.warning(f"⚠️ **{len(unmatched_k)}** CNICs Database mein nahi milay.")
                            with st.expander("Dekhein kon se CNICs missing hain"):
                                st.write(unmatched_k)

            # =========================================================================
            # OPTION 2: COMBINED SEARCH (DATABASE + TEMPORARY REFERENCE FILE)
            # =========================================================================
            else:
                st.info("💡 **Combined Mode:** Target CNICs upload ya paste karein, aur sath Temporary Reference File upload karein jahan se missing students ki details uthayi jayengi.")
                
                # Part A: Target CNICs input (Paste ya File)
                st.markdown("###### 1️⃣ Target CNICs Input (Kisko Search Karna Hai):")
                comb_tab1, comb_tab2 = st.tabs(["✍️ Direct Paste CNICs", "📁 Upload Target CNICs File"])
                
                with comb_tab1:
                    default_comb_paste = "\n".join(saved_part_cnics) if saved_part_cnics else ""
                    comb_pasted_val = st.text_area(
                        "Target CNICs Paste karein:",
                        value=default_comb_paste,
                        placeholder="33100-1234567-1\n33100-7654321-2",
                        key=f"opt2_paste_area_{selected_drive_id}",
                        height=100
                    )
                    if comb_pasted_val.strip():
                        target_raw_cnics = re.split(r"[\n,\s;]+", comb_pasted_val.strip())
                    elif saved_part_cnics:
                        target_raw_cnics = saved_part_cnics

                with comb_tab2:
                    up_file2 = st.file_uploader(
                        "Target CNICs Sheet Upload Karein (.xlsx / .csv):",
                        type=["xlsx", "csv"],
                        key=f"opt2_file_up_{selected_drive_id}"
                    )
                    if up_file2 is not None:
                        try:
                            df_in2 = pd.read_csv(up_file2, dtype=str) if up_file2.name.lower().endswith(".csv") else pd.read_excel(up_file2, dtype=str)
                            col_c2 = next((c for c in df_in2.columns if "cnic" in str(c).lower()), None)
                            if not col_c2:
                                col_c2 = st.selectbox("Sheet se CNIC column select karein:", df_in2.columns, key=f"opt2_col_sel_{selected_drive_id}")
                            if col_c2:
                                target_raw_cnics = df_in2[col_c2].dropna().tolist()
                        except Exception as e:
                            st.error(f"Target file error: {e}")

                # Part B: Temporary Reference File (Jis mein se details leni hain)
                st.markdown("###### 2️⃣ Temporary Reference Data File (Jahan se missing students ka Name, Trade wagera uthana hai):")
                ref_file = st.file_uploader(
                    "Temporary Data File Upload Karein (.xlsx / .csv):",
                    type=["xlsx", "csv"],
                    key=f"opt2_reference_file_{selected_drive_id}"
                )

                if target_raw_cnics:
                    target_dashed, lookup_cnics = parse_cnic_input(target_raw_cnics)
                    
                    # 1. Pehle Database Query
                    db_matched_df = pd.DataFrame()
                    if lookup_cnics:
                        placeholders_k = ",".join([pl] * len(lookup_cnics))
                        sql_k = f"""
                            SELECT roll_no AS 'Roll No', name AS 'Name', technology AS 'Technology', session AS 'Session', cnic AS 'CNIC'
                            FROM students 
                            WHERE cnic IN ({placeholders_k})
                        """
                        db_matched_df = pd.read_sql(sql_k, conn, params=lookup_cnics)

                    found_in_db = set(db_matched_df['CNIC'].tolist()) if not db_matched_df.empty else set()
                    missing_from_db = [c for c in target_dashed if c not in found_in_db]

                    # 2. Jo Database mein nahi hain, unhein Temporary Reference File se dhoondein
                    excel_found_rows = []
                    if ref_file is not None and missing_from_db:
                        try:
                            ref_df = pd.read_csv(ref_file, dtype=str) if ref_file.name.lower().endswith(".csv") else pd.read_excel(ref_file, dtype=str)
                            
                            ref_cnic_col = next((c for c in ref_df.columns if "cnic" in str(c).lower()), None)
                            ref_name_col = next((c for c in ref_df.columns if any(k in str(c).lower() for k in ["name", "student"])), None)
                            ref_roll_col = next((c for c in ref_df.columns if "roll" in str(c).lower()), None)
                            ref_tech_col = next((c for c in ref_df.columns if any(k in str(c).lower() for k in ["tech", "trade"])), None)

                            if ref_cnic_col:
                                for _, r in ref_df.iterrows():
                                    raw_rc = str(r[ref_cnic_col]).strip()
                                    d_rc = extract_digits(raw_rc)
                                    norm_rc = f"{d_rc[:5]}-{d_rc[5:12]}-{d_rc[12]}" if len(d_rc) == 13 else raw_rc

                                    if norm_rc in missing_from_db or d_rc in missing_from_db:
                                        s_name = str(r[ref_name_col]).strip() if ref_name_col and pd.notna(r[ref_name_col]) else "Trainee"
                                        s_roll = str(r[ref_roll_col]).strip() if ref_roll_col and pd.notna(r[ref_roll_col]) else "-"
                                        s_tech = str(r[ref_tech_col]).strip() if ref_tech_col and pd.notna(r[ref_tech_col]) else (dr_data['target_techs'][0] if dr_data['target_techs'] else "General")

                                        excel_found_rows.append({
                                            'Roll No': s_roll,
                                            'Name': s_name,
                                            'Technology': s_tech,
                                            'Session': dr_data['session'],
                                            'CNIC': norm_rc
                                        })
                        except Exception as e:
                            st.error(f"Reference file read karne mein masla: {e}")

                    excel_found_df = pd.DataFrame(excel_found_rows)
                    if not excel_found_df.empty:
                        excel_found_df = excel_found_df.drop_duplicates(subset=['CNIC'])

                    # 3. Merge both dataframes
                    cand_df = pd.concat([db_matched_df, excel_found_df], ignore_index=True)

                    st.success(
                        f"🎉 Combined Results: Kul **{len(cand_df)} Students** shamil huay "
                        f"(🏛️ Database: **{len(db_matched_df)}** | 📂 Temporary File: **{len(excel_found_df)}**)"
                    )

                    all_found_set = set(cand_df['CNIC'].tolist())
                    still_missing = [c for c in target_dashed if c not in all_found_set]
                    if still_missing:
                        st.warning(f"⚠️ **{len(still_missing)}** CNICs na to Database mein milay aur na hi Temporary File mein!")
                        with st.expander("Dekhein dono jagah se missing CNICs"):
                            st.write(still_missing)

            # =========================================================================
            # SAVE, EDIT & DOWNLOAD (COMMON ACTIONS)
            # =========================================================================
            if not cand_df.empty:
                col_save_k, col_clr_k = st.columns([1, 1])
                with col_save_k:
                    if st.button("💾 Save Participated Candidates to Database", key="btn_save_part_cands"):
                        all_part_cnics = cand_df['CNIC'].dropna().tolist()
                        cursor.execute(f"""
                            UPDATE recruitment_drives 
                            SET participated_cnics_json = {pl}, students_appeared = {pl}
                            WHERE id = {pl}
                        """, (json.dumps(all_part_cnics), len(cand_df), selected_drive_id))
                        conn.commit()
                        dr_data['participated_cnics_json'] = json.dumps(all_part_cnics)
                        dr_data['students_appeared'] = len(cand_df)
                        st.success(f"✅ {len(cand_df)} Candidates drive record mein successfully save ho gaye!")
                        st.rerun()

                with col_clr_k:
                    if saved_part_cnics:
                        if st.button("🔄 Clear / Reset Participated List", key="btn_clr_part_cands"):
                            cursor.execute(f"UPDATE recruitment_drives SET participated_cnics_json = '[]', students_appeared = 0 WHERE id = {pl}", (selected_drive_id,))
                            conn.commit()
                            st.warning("Participated list reset ho gayi hai.")
                            st.rerun()

                # Single Candidate Management (Remove / Add)
                with st.expander("🛠️ Candidates Management (Galti Durust / Add Karein)"):
                    cand_remove_options = {
                        f"{r['Roll No']} - {r['Name']} ({r['Technology']}) [{r['CNIC']}]": r['CNIC']
                        for _, r in cand_df.iterrows()
                    }
                    to_remove_label = st.selectbox("Galat candidate select karein jise hatana hai:", list(cand_remove_options.keys()), key="sel_cand_remove")
                    if st.button("❌ Remove Selected Candidate From List", key="btn_remove_single_cand"):
                        cnic_to_del = cand_remove_options[to_remove_label]
                        updated_cnics = [c for c in cand_df['CNIC'].tolist() if c != cnic_to_del]
                        cursor.execute(f"""
                            UPDATE recruitment_drives 
                            SET participated_cnics_json = {pl}, students_appeared = {pl}
                            WHERE id = {pl}
                        """, (json.dumps(updated_cnics), len(updated_cnics), selected_drive_id))
                        conn.commit()
                        st.success(f"Candidate {to_remove_label} list se hata diya gaya!")
                        st.rerun()

                    st.markdown("---")
                    new_add_cnic = st.text_input("New Student CNIC to Add:", placeholder="33100-1234567-1", key="add_new_part_cnic_input")
                    if st.button("➕ Add This CNIC to List", key="btn_add_part_single_cnic"):
                        if new_add_cnic.strip():
                            _, lookup_new = parse_cnic_input([new_add_cnic.strip()])
                            if lookup_new:
                                cursor.execute(f"SELECT cnic FROM students WHERE cnic = {pl}", (lookup_new[0],))
                                row_verify = cursor.fetchone()
                                if row_verify:
                                    curr_cnics = cand_df['CNIC'].tolist()
                                    if row_verify[0] not in curr_cnics:
                                        curr_cnics.append(row_verify[0])
                                        cursor.execute(f"UPDATE recruitment_drives SET participated_cnics_json = {pl}, students_appeared = {pl} WHERE id = {pl}", (json.dumps(curr_cnics), len(curr_cnics), selected_drive_id))
                                        conn.commit()
                                        st.success("Candidate kamyabi se shamil ho gaya!")
                                        st.rerun()
                                    else:
                                        st.info("Yeh candidate pehle se list mein mojood hai.")
                                else:
                                    st.error("Diya gaya CNIC student directory mein nahi mila.")

                # Download Annexure K Excel
                annex_k_excel = generate_annexure_k_excel(dr_data, cand_df)
                st.download_button(
                    label=f"📊 Download Official Annexure “K” Assessment Sheet ({len(cand_df)} Candidates) (.xlsx)",
                    data=annex_k_excel,
                    file_name=f"Annexure_K_{dr_data['company_name'].replace(' ', '_')}.xlsx",
                    key="btn_dl_annex_k_active"
                )

            st.markdown("---")
            st.markdown("#### 🎯 Post-Interview / Test Finalized Candidates (Selected Students)")
            st.caption("Test aur interview ke baad company ki taraf se **finalize/select kiye gaye** students ke CNICs yahan provide karein:")

            saved_sel_cnics = []
            if dr_data.get('selected_cnics_json'):
                try:
                    saved_sel_cnics = json.loads(dr_data['selected_cnics_json'])
                except Exception:
                    saved_sel_cnics = []

            fn_t1, fn_t2 = st.tabs(["📁 Upload Finalized Students File", "✍️ Paste Finalized CNICs"])
            finalized_raw_cnics = []

            with fn_t1:
                fin_file = st.file_uploader("Upload Company Final Selected / Hired List", type=["xlsx", "csv"], key="final_hired_file")
                if fin_file is not None:
                    try:
                        df_fin = safe_read_file(fin_file)
                        
                        fin_cnic_col = None
                        for c in df_fin.columns:
                            if "cnic" in str(c).lower():
                                fin_cnic_col = c
                                break
                        if not fin_cnic_col:
                            fin_cnic_col = st.selectbox("Select CNIC column:", df_fin.columns, key="fin_cnic_sel")
                        
                        if fin_cnic_col:
                            finalized_raw_cnics = df_fin[fin_cnic_col].dropna().tolist()
                    except Exception as e:
                        st.error(f"File error: {e}")

            with fn_t2:
                default_fin_text = "\n".join(saved_sel_cnics) if saved_sel_cnics else ""
                fin_text = st.text_area(
                    "Finalized / Hired students ke CNICs paste karein:",
                    value=default_fin_text,
                    placeholder="33100-1234567-1\n33100-7654321-2",
                    key="fin_text_area"
                )
                if fin_text.strip():
                    finalized_raw_cnics = re.split(r"[\n,\s;]+", fin_text.strip())

            if not finalized_raw_cnics and saved_sel_cnics:
                finalized_raw_cnics = saved_sel_cnics

            finalized_students_df = pd.DataFrame()
            if finalized_raw_cnics:
                _, fin_lookup = parse_cnic_input(finalized_raw_cnics)
                if fin_lookup:
                    fin_placeholders = ",".join([pl] * len(fin_lookup))
                    sql_fin = f"""
                        SELECT roll_no AS 'Roll No', name AS 'Name', technology AS 'Technology', session AS 'Session', phone AS 'Mobile Number', cnic AS 'CNIC', email AS 'Email'
                        FROM students 
                        WHERE cnic IN ({fin_placeholders})
                        ORDER BY technology, name
                    """
                    finalized_students_df = pd.read_sql(sql_fin, conn, params=fin_lookup)

                if not finalized_students_df.empty:
                    st.success(f"🏆 Found **{len(finalized_students_df)}** Finalized / Selected Candidate(s) in Database.")
                    
                    c_save_sel, c_dl_sel = st.columns(2)
                    with c_save_sel:
                        if st.button("💾 Save Selected Candidates to Database", key="btn_save_sel_db"):
                            cursor.execute(f"""
                                UPDATE recruitment_drives 
                                SET selected_cnics_json = {pl}, students_shortlisted = {pl}
                                WHERE id = {pl}
                            """, (json.dumps(finalized_students_df['CNIC'].tolist()), len(finalized_students_df), selected_drive_id))
                            conn.commit()
                            dr_data['selected_cnics_json'] = json.dumps(finalized_students_df['CNIC'].tolist())
                            dr_data['students_shortlisted'] = len(finalized_students_df)
                            st.success(f"✅ {len(finalized_students_df)} Selected Students permanently save ho gaye!")
                            st.rerun()

                    with c_dl_sel:
                        fin_excel_buf = io.BytesIO()
                        with pd.ExcelWriter(fin_excel_buf, engine='openpyxl') as wr:
                            finalized_students_df.to_excel(wr, index=False, sheet_name="Selected_Students")
                        st.download_button(
                            label=f"📥 Download Selected Candidates Excel ({len(finalized_students_df)} Students) (.xlsx)",
                            data=fin_excel_buf.getvalue(),
                            file_name=f"Selected_Candidates_{dr_data['company_name'].replace(' ', '_')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="btn_dl_sel_active_excel"
                        )

                    with st.expander("🛠️ Edit Selected / Shortlisted List (Remove Incorrect Candidate)"):
                        sel_remove_options = {
                            f"{r['Roll No']} - {r['Name']} ({r['Technology']}) [{r['CNIC']}]": r['CNIC']
                            for _, r in finalized_students_df.iterrows()
                        }
                        to_remove_sel_label = st.selectbox("Galat selected student jise hatana hai:", list(sel_remove_options.keys()), key="sel_to_remove_box")
                        if st.button("❌ Remove From Selected List", key="btn_remove_single_sel"):
                            cnic_sel_del = sel_remove_options[to_remove_sel_label]
                            updated_sel_cnics = [c for c in finalized_students_df['CNIC'].tolist() if c != cnic_sel_del]
                            cursor.execute(f"""
                                UPDATE recruitment_drives 
                                SET selected_cnics_json = {pl}, students_shortlisted = {pl}
                                WHERE id = {pl}
                            """, (json.dumps(updated_sel_cnics), len(updated_sel_cnics), selected_drive_id))
                            conn.commit()
                            st.success(f"{to_remove_sel_label} ko selected list se hata diya gaya!")
                            st.rerun()

                    st.markdown("##### 💬 Individual Candidate WhatsApp Message & 1-Click Dispatch")
                    
                    sel_fin_cnic = st.selectbox(
                        "Select Candidate to Send Personalized WhatsApp Message:",
                        finalized_students_df['CNIC'].tolist(),
                        format_func=lambda x: f"{finalized_students_df.loc[finalized_students_df['CNIC']==x, 'Name'].values[0]} - Roll# {finalized_students_df.loc[finalized_students_df['CNIC']==x, 'Roll No'].values[0]} ({finalized_students_df.loc[finalized_students_df['CNIC']==x, 'Technology'].values[0]})",
                        key="rd_fin_candidate_sel"
                    )

                    row_sel_fin = finalized_students_df.loc[finalized_students_df['CNIC'] == sel_fin_cnic].iloc[0]
                    f_name = str(row_sel_fin['Name'])
                    f_roll = str(row_sel_fin['Roll No'])
                    f_tech = str(row_sel_fin['Technology'])
                    f_phone = str(row_sel_fin['Mobile Number'])

                    template_mode = st.radio(
                        "Choose Message Purpose:",
                        ["🎉 Job Offer / Selection Notification", "📋 Drive / Test Interview Call", "✍️ Custom Message"],
                        horizontal=True,
                        key=f"rd_tmpl_mode_{sel_fin_cnic}"
                    )

                    c_name_val = dr_data['company_name']
                    d_date_val = dr_data.get('drive_date', '')
                    d_venue_val = dr_data.get('venue', 'GCT Samanabad Faisalabad')

                    if template_mode == "🎉 Job Offer / Selection Notification":
                        sample_tmpl = (
                            f"Assalam-o-Alaikum {f_name} (Roll No: {f_roll}, Trade: {f_tech}),\n"
                            f"Mubarak ho! Placement Cell GCT Samanabad ki taraf se aapko inform kiya jata hai ke aap '{c_name_val}' mein recruit/finalize ho chuke hain.\n"
                            f"Barah-e-karam apna original CNIC, updated CV aur documents le kar foran Placement Office (IPO Cell) tashreef layein taake offer letter formalities complete ki ja sakein.\n"
                            f"Regards,\nEngr. Muhammad Wasif Ali\nInstitute Placement Officer (IPO), GCT Samanabad"
                        )
                    elif template_mode == "📋 Drive / Test Interview Call":
                        sample_tmpl = (
                            f"Assalam-o-Alaikum {f_name} (Roll No: {f_roll}, Trade: {f_tech}),\n"
                            f"Aapko GCT Samanabad Placement Cell ki taraf se inform kiya jata hai ke '{c_name_val}' ki On-Campus Recruitment Drive {d_date_val} ko {d_venue_val} mein scheduled hai.\n"
                            f"Barah-e-karam formal dress, 2 copies of CV aur original CNIC ke sath theek time par pohanchein.\n"
                            f"Regards,\nInstitute Placement Officer, GCT Samanabad"
                        )
                    else:
                        sample_tmpl = (
                            f"Assalam-o-Alaikum {f_name} (Roll No: {f_roll}),\n"
                            f"Yeh message Placement Cell GCT Samanabad ki taraf se bheja ja raha hai regarding {c_name_val} placement process. Barah-e-karam IPO Office se rabta karein."
                        )

                    custom_fin_msg = st.text_area(
                        f"Message for {f_name} (Editable):", 
                        value=sample_tmpl, 
                        height=120, 
                        key=f"rd_fin_msg_box_{sel_fin_cnic}_{template_mode}"
                    )

                    wa_url_fin = generate_whatsapp_url(f_phone, custom_fin_msg)
                    if wa_url_fin:
                        st.markdown(f"<a href='{wa_url_fin}' target='_blank' class='wa-btn'>💬 Send WhatsApp to {f_name} ({f_phone})</a>", unsafe_allow_html=True)
                    else:
                        st.warning(f"⚠️ {f_name} ka mobile number ({f_phone}) valid international format mein nahi hai.")

                    st.markdown("---")
                    st.markdown("###### 👥 All Selected Candidates Quick Chat Links:")
                    for idx, row_st in finalized_students_df.iterrows():
                        p_raw = row_st['Mobile Number']
                        st_n = row_st['Name']
                        st_r = row_st['Roll No']
                        st_t = row_st['Technology']

                        quick_msg = (
                            f"Assalam-o-Alaikum {st_n} (Roll No: {st_r}),\n"
                            f"Mubarak ho! Aap '{c_name_val}' mein recruit/finalize ho chuke hain. Barah-e-karam Placement Office (IPO Cell) GCT Samanabad report karein."
                        )
                        wa_url = generate_whatsapp_url(p_raw, quick_msg)

                        col_a, col_b, col_c = st.columns([3, 2, 2])
                        with col_a:
                            st.write(f"**{idx+1}. {st_n}** ({st_t} - {st_r})")
                        with col_b:
                            st.write(f"📞 `{p_raw}`")
                        with col_c:
                            if wa_url:
                                st.markdown(f"<a href='{wa_url}' target='_blank' class='wa-btn'>💬 Chat with {st_n.split()[0]}</a>", unsafe_allow_html=True)
                            else:
                                st.caption("Invalid Phone")

                    fin_phones = finalized_students_df['Mobile Number'].dropna().tolist()
                    clean_fin_nums = [format_whatsapp_number(p) for p in fin_phones if format_whatsapp_number(p)]
                    st.text_area(
                        "Finalized Candidates Comma-Separated Numbers (For Bulk Group Dispatch):",
                        value=",".join(clean_fin_nums),
                        height=70
                    )
                else:
                    st.warning("Diye gaye CNICs database mein nahi milay.")

            default_app_val = len(cand_df) if not cand_df.empty else int(dr_data.get('students_appeared') or 40)
            default_short_val = len(finalized_students_df) if not finalized_students_df.empty else int(dr_data.get('students_shortlisted') or 12)

            with st.form("annex_l_final_form"):
                st.markdown("---")
                st.markdown("#### Annexure “L” Details (Employer Feedback)")
                l_col1, l_col2 = st.columns(2)
                with l_col1:
                    app_count = st.number_input("Total Students Appeared in Drive*", min_value=0, value=int(default_app_val))
                    coord = st.selectbox("Coordination of recruiting schedule by Placement Officer", ["E", "G", "UN"], index=0)
                    staff_h = st.selectbox("Staff of institute / venue was friendly & helpful", ["E", "G", "UN"], index=0)
                    ven_comf = st.selectbox("Interview venue was comfortable", ["E", "G", "UN"], index=0)
                    tech_sk = st.selectbox("Students Technical skills met your expectations", ["E", "G", "UN"], index=0)
                with l_col2:
                    short_count = st.number_input("Total Students Shortlisted / Finalized*", min_value=0, value=int(default_short_val))
                    res_prof = st.selectbox("Students resumes were professional", ["E", "G", "UN"], index=1)
                    comm_sk = st.selectbox("Students communication skills were commendable", ["E", "G", "UN"], index=1)
                    dress_f = st.selectbox("Students were formally dressed", ["E", "G", "UN"], index=0)
                    over_imp = st.selectbox("Overall impression of recruitment drive", ["E", "G", "UN"], index=0)

                sugg = st.text_area("Suggestions (if any)", value=dr_data.get('l_suggestions') or "Practical lab exposure was highly appreciated. Future drives are strongly recommended.")

                st.markdown("---")
                st.markdown("###### 💰 Final Activity Expense & Submission")
                rd_budget = st.number_input("Budget Utilized for this Activity (PKR)*", min_value=0.0, value=float(dr_data.get('budget_utilized') or 30000.0), step=500.0)

                final_drive_btn = st.form_submit_button("✅ Final Submit & Save Permanently to Archive")
                if final_drive_btn:
                    part_json_val = json.dumps(cand_df['CNIC'].tolist()) if not cand_df.empty and 'CNIC' in cand_df.columns else dr_data.get('participated_cnics_json')
                    sel_json_val = json.dumps(finalized_students_df['CNIC'].tolist()) if not finalized_students_df.empty and 'CNIC' in finalized_students_df.columns else dr_data.get('selected_cnics_json')

                    cursor.execute(f"""
                    UPDATE recruitment_drives SET
                        students_appeared = {pl}, students_shortlisted = {pl},
                        l_coordination = {pl}, l_staff = {pl}, l_venue_comfort = {pl},
                        l_tech_skills = {pl}, l_resumes = {pl}, l_comm_skills = {pl}, l_dress = {pl},
                        l_overall = {pl}, l_suggestions = {pl}, budget_utilized = {pl}, 
                        participated_cnics_json = {pl}, selected_cnics_json = {pl},
                        status = 'Completed'
                    WHERE id = {pl}
                    """, (
                        app_count, short_count, coord, staff_h, ven_comf,
                        tech_sk, res_prof, comm_sk, dress_f, over_imp, sugg,
                        rd_budget, part_json_val, sel_json_val, selected_drive_id
                    ))
                    conn.commit()
                    st.success(f"Recruitment Drive mukammal taur par archive ho gayi! (Budget Utilized: PKR {rd_budget:,.2f})")
                    st.rerun()

            dr_data.update({
                'students_appeared': default_app_val, 'students_shortlisted': default_short_val,
                'l_coordination': dr_data.get('l_coordination', 'E'), 'l_staff': dr_data.get('l_staff', 'E'),
                'l_venue_comfort': dr_data.get('l_venue_comfort', 'E'), 'l_tech_skills': dr_data.get('l_tech_skills', 'E'),
                'l_resumes': dr_data.get('l_resumes', 'G'), 'l_comm_skills': dr_data.get('l_comm_skills', 'G'),
                'l_dress': dr_data.get('l_dress', 'E'), 'l_overall': dr_data.get('l_overall', 'E'),
                'l_suggestions': dr_data.get('l_suggestions', ''), 'budget_utilized': dr_data.get('budget_utilized', 0.0)
            })
            st.markdown("##### 📥 Download All Evaluated Official Formats")
            dl_col1, dl_col2, dl_col3 = st.columns(3)
            with dl_col1:
                tab2_annex_j = generate_annexure_j(dr_data, cand_df)
                st.download_button(
                    "📄 Download Annexure “J” (.docx)",
                    data=tab2_annex_j,
                    file_name=f"Annexure_J_{dr_data['company_name'].replace(' ', '_')}.docx"
                )
            with dl_col2:
                tab2_annex_k = generate_annexure_k_excel(dr_data, cand_df)
                st.download_button(
                    "📊 Download Annexure “K” (.xlsx)",
                    data=tab2_annex_k,
                    file_name=f"Annexure_K_{dr_data['company_name'].replace(' ', '_')}.xlsx"
                )
            with dl_col3:
                annex_l_docx = generate_annexure_l(dr_data)
                st.download_button(
                    "📝 Download Annexure “L” (.docx)",
                    data=annex_l_docx,
                    file_name=f"Annexure_L_{dr_data['company_name'].replace(' ', '_')}.docx"
                )
        else:
            st.info("Pehle Tab 1 se Recruitment Drive plan karein.")

    # -------------------------------------------------------------------------
    # TAB 3: PERMANENT ARCHIVED RECRUITMENT DRIVES, EDIT & DELETE
    # -------------------------------------------------------------------------
    with rd_tab3:
        st.markdown("##### 📂 Permanent Archived Recruitment Drives, Edit & Delete")
        df_all_drives = pd.read_sql("SELECT id, company_name, drive_date, job_title, target_techs, session, students_appeared, students_shortlisted, budget_utilized, status FROM recruitment_drives ORDER BY id DESC", conn)
        if not df_all_drives.empty:
            st.dataframe(df_all_drives, use_container_width=True)
            
            st.markdown("---")
            st.markdown("###### ⚙️ Manage Archived Drive (Re-download / Edit / Delete)")
            re_did = st.selectbox("Archived Drive Select Karein:", df_all_drives['id'].tolist(), format_func=lambda x: f"ID {x} - {df_all_drives.loc[df_all_drives['id']==x, 'company_name'].values[0]} ({df_all_drives.loc[df_all_drives['id']==x, 'drive_date'].values[0]})")
            
            cursor.execute(f"SELECT * FROM recruitment_drives WHERE id = {pl}", (re_did,))
            archive_dr = dict(zip([d[0] for d in cursor.description], cursor.fetchone()))
            archive_dr['target_techs'] = archive_dr['target_techs'].split(",") if archive_dr['target_techs'] else []

            # 1. Roster Candidates
            re_saved_limits = {}
            if archive_dr.get('candidate_limits_json'):
                try:
                    re_saved_limits = json.loads(archive_dr['candidate_limits_json'])
                except Exception:
                    pass

            re_cands = []
            for t in archive_dr['target_techs']:
                cnt_lim = re_saved_limits.get(t)
                q_k = f"SELECT roll_no AS 'Roll No', name AS 'Name', technology AS 'Technology', session AS 'Session', cnic AS 'CNIC', phone AS 'Mobile Number', email AS 'Email' FROM students WHERE session = {pl} AND technology = {pl} ORDER BY roll_no"
                if cnt_lim and cnt_lim > 0:
                    q_k += f" LIMIT {cnt_lim}"
                df_re_sub = pd.read_sql(q_k, conn, params=[archive_dr['session'], t])
                if not df_re_sub.empty:
                    re_cands.append(df_re_sub)

            re_cand_df = pd.concat(re_cands, ignore_index=True) if re_cands else pd.DataFrame(columns=['Roll No', 'Name', 'Technology', 'Session', 'CNIC', 'Mobile Number', 'Email'])

            # 2. Actual Participated Candidates
            re_part_df = re_cand_df.copy()
            if archive_dr.get('participated_cnics_json'):
                try:
                    p_cnics = json.loads(archive_dr['participated_cnics_json'])
                    if p_cnics:
                        p_holders = ",".join([pl] * len(p_cnics))
                        q_part = f"SELECT roll_no AS 'Roll No', name AS 'Name', technology AS 'Technology', session AS 'Session', cnic AS 'CNIC' FROM students WHERE cnic IN ({p_holders}) ORDER BY technology, roll_no"
                        df_part_loaded = pd.read_sql(q_part, conn, params=p_cnics)
                        if not df_part_loaded.empty:
                            re_part_df = df_part_loaded
                except Exception:
                    pass

            # 3. Final Selected Candidates
            re_sel_df = pd.DataFrame()
            if archive_dr.get('selected_cnics_json'):
                try:
                    s_cnics = json.loads(archive_dr['selected_cnics_json'])
                    if s_cnics:
                        s_holders = ",".join([pl] * len(s_cnics))
                        q_sel = f"SELECT roll_no AS 'Roll No', name AS 'Name', technology AS 'Technology', session AS 'Session', phone AS 'Mobile Number', cnic AS 'CNIC', email AS 'Email' FROM students WHERE cnic IN ({s_holders}) ORDER BY technology, name"
                        re_sel_df = pd.read_sql(q_sel, conn, params=s_cnics)
                except Exception:
                    pass

            if archive_dr.get('approval_doc_blob'):
                re_drive_app_bytes = decompress_file_bytes(archive_dr['approval_doc_blob'])
            else:
                re_drive_app_payload = {
                    'company_name': archive_dr.get('company_name', ''),
                    'company_address': archive_dr.get('company_address', 'Faisalabad'),
                    'drive_date': str(archive_dr.get('drive_date', '')),
                    'drive_time': archive_dr.get('drive_time', '09:30 AM'),
                    'venue': archive_dr.get('venue', 'Main Hall, GCT Samanabad'),
                    'written_test': archive_dr.get('written_test', 'Yes - Screening & Technical Test'),
                    'target_techs': archive_dr.get('target_techs', []),
                    'session': archive_dr.get('session', ''),
                    'salary_range': archive_dr.get('salary_range', '22,000 - 30,000'),
                    'refreshment_heads': archive_dr.get('refreshment_heads') or 10,
                    'backdrop_details': archive_dr.get('backdrop_details') or 'Drive Backdrop Flex Required',
                    'bouquet_details': archive_dr.get('bouquet_details') or 'Welcome Bouquet & Protocol Required'
                }
                re_drive_app_bytes = generate_drive_principal_application(re_drive_app_payload).getvalue()

            c_app, c_j, c_k, c_l, c_xl = st.columns(5)
            with c_app:
                st.download_button("📥 Re-Download Approval", re_drive_app_bytes, f"Archived_Drive_Approval_{re_did}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=f"dl_arch_dr_app_{re_did}")
            with c_j:
                st.download_button("📥 Re-Download Annexure J", generate_annexure_j(archive_dr, re_cand_df), f"Archived_Annex_J_{re_did}.docx", key=f"dl_arch_j_{re_did}")
            with c_k:
                st.download_button(f"📥 Re-Download Annexure K ({len(re_part_df)} Cands)", generate_annexure_k_excel(archive_dr, re_part_df), f"Archived_Annex_K_{re_did}.xlsx", key=f"dl_arch_k_{re_did}")
            with c_l:
                st.download_button("📥 Re-Download Annexure L", generate_annexure_l(archive_dr), f"Archived_Annex_L_{re_did}.docx", key=f"dl_arch_l_{re_did}")
            with c_xl:
                st.download_button("📥 Re-Download Roster Excel", generate_candidate_roster_excel(archive_dr, re_cand_df), f"Archived_Candidates_{re_did}.xlsx", key=f"dl_arch_xl_{re_did}")

            if not re_sel_df.empty:
                s_buf_arch = io.BytesIO()
                with pd.ExcelWriter(s_buf_arch, engine='openpyxl') as wr:
                    re_sel_df.to_excel(wr, index=False, sheet_name="Selected_Students")
                st.download_button(
                    f"🏆 Re-Download Selected Candidates Excel ({len(re_sel_df)} Students) (.xlsx)",
                    s_buf_arch.getvalue(),
                    f"Archived_Selected_{archive_dr['company_name'].replace(' ', '_')}.xlsx",
                    key=f"dl_arch_sel_{re_did}"
                )

            with st.expander(f"✏️ Edit / Correct Drive Entry (ID: {re_did})"):
                with st.form(f"edit_drive_{re_did}"):
                    e_comp = st.text_input("Company Name", value=archive_dr['company_name'])
                    e_jtitle = st.text_input("Job Title", value=archive_dr['job_title'])
                    e_c1, e_c2 = st.columns(2)
                    with e_c1:
                        e_app = st.number_input("Students Appeared", min_value=0, value=int(archive_dr['students_appeared'] or 0))
                        e_short = st.number_input("Students Shortlisted", min_value=0, value=int(archive_dr['students_shortlisted'] or 0))
                    with e_c2:
                        e_bud_d = st.number_input("Budget Utilized (PKR)", min_value=0.0, value=float(archive_dr['budget_utilized'] or 0.0), step=500.0)
                        e_stat_d = st.selectbox("Status", ["Completed", "Planned"], index=0 if archive_dr['status'] == 'Completed' else 1)
                    
                    if st.form_submit_button("💾 Save Corrected Drive Data"):
                        cursor.execute(f"""
                            UPDATE recruitment_drives 
                            SET company_name = {pl}, job_title = {pl}, students_appeared = {pl}, students_shortlisted = {pl}, budget_utilized = {pl}, status = {pl}
                            WHERE id = {pl}
                        """, (e_comp, e_jtitle, e_app, e_short, e_bud_d, e_stat_d, re_did))
                        conn.commit()
                        st.success("Drive data update ho gaya! Refresh karein.")
                        st.rerun()

            with st.expander(f"🗑️ Delete Record (ID: {re_did})"):
                st.warning("⚠️ Kya aap waqai is Recruitment Drive ko permanently delete karna chahte hain?")
                confirm_del_d = st.checkbox("Haan, main is recruitment drive ko delete karna chahta hoon", key=f"del_chk_d_{re_did}")
                if st.button("🔴 Delete This Recruitment Drive", key=f"del_btn_d_{re_did}"):
                    if confirm_del_d:
                        cursor.execute(f"DELETE FROM recruitment_drives WHERE id = {pl}", (re_did,))
                        conn.commit()
                        st.success(f"Drive Record ID {re_did} delete ho chuka hai.")
                        st.rerun()
                    else:
                        st.error("Pehle confirmation check karein.")
        else:
            st.write("Abhi tak koi Recruitment Drive record nahi mili.")
            
# =========================================================================
# --- 7. Log General Activity ---
# =========================================================================
elif choice == "Log Activity":
    st.subheader("Record General Activity")
    act_type = st.selectbox("Activity Type", ["Industrial Tour", "Recruitment Drive", "Workshop/Seminar"])
    title = st.text_input("Activity Title / Topic")
    org = st.text_input("Company Name / Organizer")
    date = st.date_input("Date")
    budget_gen = st.number_input("Budget Utilized (PKR)", min_value=0.0, value=0.0, step=500.0)
    if st.button("Create Activity Event"):
        cursor.execute(
            "INSERT INTO activities (activity_type, title, company_or_speaker, activity_date, budget_utilized) VALUES (?, ?, ?, ?, ?)",
            (act_type, title, org, str(date), budget_gen)
        )
        conn.commit()
        st.success("Activity logged successfully!")

# =========================================================================
# --- 8. Traceability Form Upload & CNIC-linked Record ---
# =========================================================================
elif choice == "📑 Traceability Form Upload":
    st.subheader("📑 DAE Pass-Out Graduate Traceability Form")
    st.caption(
        "Filled Traceability Form upload karein. System CNIC detect karke student record "
        "samne layega. Aap chahein to student directory update karein, warna sirf form ko "
        "usi CNIC ke against compressed record ke taur par save karein."
    )

    tr_tab1, tr_tab2 = st.tabs(["📤 Upload & Save Form", "🗂️ Saved Traceability Forms"])

    with tr_tab1:
        uploaded_trace = st.file_uploader(
            "Filled Traceability Form Upload Karein:",
            type=["pdf", "png", "jpg", "jpeg", "webp", "tif", "tiff"],
            key="traceability_form_uploader"
        )

        if uploaded_trace is not None:
            trace_filename = uploaded_trace.name
            trace_raw = uploaded_trace.getvalue()
            trace_compressed = compress_file_bytes(trace_raw)

            with st.spinner("Traceability Form read/OCR ho raha hai..."):
                trace_text = extract_traceability_text(trace_raw, trace_filename)
                trace_data = parse_traceability_form(trace_text, TECHNOLOGIES_LIST)

            st.markdown("---")
            orig_kb = len(trace_raw) / 1024
            comp_kb = len(trace_compressed or b"") / 1024
            if orig_kb > 0:
                st.caption(
                    f"💾 Storage: Original `{orig_kb:.1f} KB` → Compressed `{comp_kb:.1f} KB` "
                    f"({(1 - comp_kb / orig_kb) * 100:.1f}% smaller)"
                )

            detected_trace_cnic = trace_data.get("cnic", "").strip()

            if detected_trace_cnic:
                cursor.execute("SELECT * FROM students WHERE cnic = ?", (detected_trace_cnic,))
                existing_trace_row = cursor.fetchone()
                if existing_trace_row:
                    trace_cols = [d[0] for d in cursor.description]
                    existing_trace_student = dict(zip(trace_cols, existing_trace_row))
                    trace_student_found = True
                else:
                    existing_trace_student = None
                    trace_student_found = False
            else:
                existing_trace_student = None
                trace_student_found = False

            if detected_trace_cnic:
                if trace_student_found:
                    st.success(
                        f"✅ CNIC `{detected_trace_cnic}` mil gaya — "
                        f"Student **{existing_trace_student.get('name', '')}** ka record database mein mojood hai."
                    )
                else:
                    st.warning(
                        f"⚠️ CNIC `{detected_trace_cnic}` form se detect hua, lekin Student Directory "
                        "mein is CNIC ka record nahi mila. Form phir bhi CNIC ke against save kiya ja sakta hai."
                    )
            else:
                st.error(
                    "❌ Form se CNIC detect nahi hua. Neeche CNIC manually enter karein; "
                    "save karne ke liye valid 13-digit CNIC zaroori hai."
                )

            with st.expander("🔍 OCR / Extracted Form Data", expanded=False):
                if trace_text:
                    st.text_area(
                        "Extracted Text (Reference Only)",
                        value=trace_text,
                        height=160,
                        key="trace_ocr_preview"
                    )
                else:
                    st.warning(
                        "Text/OCR detect nahi hua. Agar form scanned image hai to "
                        "PyMuPDF + Tesseract install honay chahiye."
                    )

            st.markdown("#### 👤 CNIC Match & Student Record")
            left, right = st.columns(2)

            with left:
                edit_trace_cnic = st.text_input(
                    "CNIC (Unique Key)*",
                    value=detected_trace_cnic,
                    placeholder="33100-1234567-1",
                    key="trace_edit_cnic"
                )
                edit_trace_name = st.text_input(
                    "Student Full Name",
                    value=trace_data.get("name") or (
                        str(existing_trace_student.get("name") or "") if trace_student_found else ""
                    ),
                    key="trace_edit_name"
                )
                edit_trace_roll = st.text_input(
                    "Roll No",
                    value=trace_data.get("roll_no") or (
                        str(existing_trace_student.get("roll_no") or "") if trace_student_found else ""
                    ),
                    key="trace_edit_roll"
                )
                selected_trace_tech = trace_data.get("technology") or (
                    str(existing_trace_student.get("technology") or "") if trace_student_found else ""
                )
                if selected_trace_tech not in TECHNOLOGIES_LIST:
                    selected_trace_tech = "Electrical"
                edit_trace_tech = st.selectbox(
                    "DAE Technology / Trade",
                    TECHNOLOGIES_LIST,
                    index=TECHNOLOGIES_LIST.index(selected_trace_tech),
                    key="trace_edit_tech"
                )

            with right:
                edit_trace_phone = st.text_input(
                    "Mobile / WhatsApp",
                    value=trace_data.get("phone") or (
                        str(existing_trace_student.get("phone") or "") if trace_student_found else ""
                    ),
                    key="trace_edit_phone"
                )

                status_choices_trace = ["Enrolled", "Placed", "Pass Out", "Dropped Out"]
                existing_status_trace = (
                    str(existing_trace_student.get("status") or "")
                    if trace_student_found else ""
                )

                trace_status_guess = existing_status_trace or "Pass Out"
                if trace_data.get("professional_status") == "Employed / On Job":
                    trace_status_guess = "Placed"
                elif trace_data.get("professional_status") == "Seeking Employment":
                    trace_status_guess = "Pass Out"

                if trace_status_guess not in status_choices_trace:
                    trace_status_guess = "Enrolled"

                edit_trace_status = st.selectbox(
                    "Student Directory Status",
                    status_choices_trace,
                    index=status_choices_trace.index(trace_status_guess),
                    key="trace_edit_status"
                )

                default_company_trace = (
                    trace_data.get("company_name")
                    or trace_data.get("business_name")
                    or (str(existing_trace_student.get("placed_company") or "") if trace_student_found else "")
                )
                edit_trace_company = st.text_input(
                    "Company / Industry / Business",
                    value=default_company_trace,
                    key="trace_edit_company"
                )

                default_salary_trace = trace_data.get("monthly_salary") or (
                    str(existing_trace_student.get("placed_salary") or "")
                    if trace_student_found else ""
                )
                edit_trace_salary = st.text_input(
                    "Salary / Monthly Income (PKR)",
                    value=default_salary_trace,
                    key="trace_edit_salary"
                )

            st.markdown("---")
            st.markdown("#### 📋 Traceability Form Details")
            st.caption(
                "Ye fields form se read huay data ka reference record hain. Original filled form "
                "bhi CNIC ke against save hoga; Student Directory mein sirf core fields update honge."
            )

            detail_cols = st.columns(3)
            detail_items = [
                ("Professional Status", trace_data.get("professional_status", "")),
                ("Session Start Date", trace_data.get("session_start_date", "")),
                ("Designation / Role", trace_data.get("designation_role", "")),
                ("Course / Degree", trace_data.get("course_degree", "")),
                ("Admission Status", trace_data.get("admission_status", "")),
                ("Date of Joining", trace_data.get("joining_date", "")),
                ("Nature of Business / Trade", trace_data.get("business_trade_work", "")),
                ("Estimated Monthly Income", trace_data.get("estimated_income", "")),
                ("Preferred Work Location", trace_data.get("preferred_work_location", "")),
                ("Work Address / City", trace_data.get("work_address_city", "")),
                ("University / Institute", trace_data.get("university_name", "")),
                ("Campus Location / City", trace_data.get("campus_location_city", "")),
            ]
            for idx, (label, val) in enumerate(detail_items):
                with detail_cols[idx % 3]:
                    st.text_input(
                        label,
                        value=val,
                        disabled=True,
                        key=f"trace_detail_{idx}"
                    )

            st.markdown("---")
            action1, action2 = st.columns(2)

            valid_trace_cnic = bool(re.match(r"^\d{5}-\d{7}-\d{1}$", str(edit_trace_cnic).strip()))

            with action1:
                update_and_save = st.button(
                    "🔄 Update Student + Save Traceability Form",
                    type="primary",
                    use_container_width=True,
                    disabled=not valid_trace_cnic,
                    key="trace_update_and_save"
                )

            with action2:
                save_form_only = st.button(
                    "📌 No Change — Save Form Against CNIC Only",
                    use_container_width=True,
                    disabled=not valid_trace_cnic,
                    key="trace_save_only"
                )

            if not valid_trace_cnic:
                st.warning(
                    "CNIC valid format mein enter karein: `33100-1234567-1`. 13 digits honi chahiye."
                )

            if update_and_save or save_form_only:
                final_cnic = str(edit_trace_cnic).strip()
                extracted_data_json = json.dumps(
                    trace_data,
                    ensure_ascii=False,
                    separators=(",", ":")
                ).encode("utf-8")
                compressed_extracted_data = compress_file_bytes(extracted_data_json)

                cursor.execute(
                    "INSERT INTO traceability_forms "
                    "(cnic, student_name, original_filename, extracted_data_blob, form_blob) "
                    "VALUES (?, ?, ?, ?, ?)",
                    (
                        final_cnic,
                        str(edit_trace_name or "").strip(),
                        trace_filename,
                        compressed_extracted_data,
                        trace_compressed
                    )
                )

                if update_and_save:
                    clean_name = str(edit_trace_name or "").strip()
                    clean_roll = str(edit_trace_roll or "").strip()
                    clean_phone = str(edit_trace_phone or "").strip()
                    clean_company = str(edit_trace_company or "").strip()
                    clean_salary = str(edit_trace_salary or "").strip()

                    if trace_student_found and detected_trace_cnic == final_cnic:
                        cursor.execute("""
                            UPDATE students
                            SET name = ?,
                                roll_no = ?,
                                technology = ?,
                                phone = ?,
                                status = ?,
                                placed_company = ?,
                                placed_salary = ?
                            WHERE cnic = ?
                        """, (
                            clean_name,
                            clean_roll,
                            edit_trace_tech,
                            clean_phone,
                            edit_trace_status,
                            clean_company,
                            clean_salary,
                            final_cnic
                        ))
                        update_message = (
                            f"Student **{clean_name}** ka core record update ho gaya aur "
                            f"Traceability Form CNIC `{final_cnic}` ke against save ho gaya."
                        )
                    elif trace_student_found and detected_trace_cnic != final_cnic:
                        update_message = (
                            f"Traceability Form CNIC `{final_cnic}` ke against save ho gaya. "
                            "Safety ke liye Student Directory update nahi kiya gaya kyun ke CNIC change hua."
                        )
                    else:
                        update_message = (
                            f"Student Directory mein CNIC `{final_cnic}` nahi mila, is liye "
                            "sirf Traceability Form record save kiya gaya."
                        )
                else:
                    update_message = (
                        "Student Directory mein koi change nahi ki gayi. "
                        f"Filled Traceability Form CNIC `{final_cnic}` ke against save ho gaya."
                    )

                conn.commit()
                st.success("✅ " + update_message)
                st.rerun()

    with tr_tab2:
        st.markdown("##### 🗂️ Saved Traceability Forms")
        st.caption("Har uploaded form compressed form mein CNIC ke against stored hai.")

        saved_cnic = st.text_input(
            "CNIC se saved forms search karein:",
            placeholder="33100-1234567-1",
            key="saved_trace_cnic_search"
        )

        if saved_cnic.strip():
            search_cnic = normalize_cnic(saved_cnic) or saved_cnic.strip()
            saved_df = pd.read_sql("""
                SELECT id AS 'ID',
                       cnic AS 'CNIC',
                       student_name AS 'Student Name',
                       original_filename AS 'File Name',
                       uploaded_at AS 'Saved At'
                FROM traceability_forms
                WHERE cnic = ?
                ORDER BY id DESC
            """, conn, params=(search_cnic,))
        else:
            saved_df = pd.read_sql("""
                SELECT id AS 'ID',
                       cnic AS 'CNIC',
                       student_name AS 'Student Name',
                       original_filename AS 'File Name',
                       uploaded_at AS 'Saved At'
                FROM traceability_forms
                ORDER BY id DESC
                LIMIT 100
            """, conn)

        if not saved_df.empty:
            st.dataframe(saved_df, use_container_width=True, hide_index=True)

            selected_form_id = st.selectbox(
                "Saved Form open/download karne ke liye ID select karein:",
                saved_df["ID"].tolist(),
                key="trace_saved_id"
            )

            cursor.execute(
                "SELECT cnic, original_filename, form_blob FROM traceability_forms WHERE id = ?",
                (int(selected_form_id),)
            )
            saved_row = cursor.fetchone()

            if saved_row:
                saved_cnic_value, saved_filename, saved_blob = saved_row
                original_form_bytes = decompress_file_bytes(saved_blob)
                st.download_button(
                    "📥 Download Original Filled Traceability Form",
                    data=original_form_bytes,
                    file_name=saved_filename or f"Traceability_Form_{saved_cnic_value}.pdf",
                    mime="application/pdf" if str(saved_filename).lower().endswith(".pdf") else "application/octet-stream",
                    use_container_width=True,
                    key=f"download_trace_{selected_form_id}"
                )
        else:
            st.info("Abhi koi Traceability Form saved nahi hai.")

# =========================================================================
# --- 9. Traceability, Export & Consolidated Annual TEVTA Report ---
# =========================================================================
elif choice == "Traceability & Export":
    st.subheader("Data Traceability & Official Reports")
    
    exp_tab1, exp_tab2 = st.tabs(["Full Student Directory Export", "📑 Consolidated TEVTA Annual Report"])

    with exp_tab1:
        st.markdown("##### Download Complete Student Directory")
        st.caption("Aapke tamam enrolled, placed aur archived students ka complete CSV data backup.")
        df_all_stds = pd.read_sql("SELECT cnic AS 'CNIC', roll_no AS 'Roll No', name AS 'Name', technology AS 'Technology', session AS 'Session', phone AS 'Mobile Number', email AS 'Email', status AS 'Status', placed_company AS 'Placed Company', placed_salary AS 'Salary', placed_date AS 'Placement Date', cv_verified AS 'CV Verified', cv_filename AS 'CV Stored' FROM students", conn)
        df_all_stds['CV Verified'] = df_all_stds['CV Verified'].apply(lambda x: "Yes (Verified)" if x == 1 else "No")
        csv = df_all_stds.to_csv(index=False).encode('utf-8')
        st.download_button(f"📥 Download Full Student Directory ({len(df_all_stds):,} Records) (CSV)", data=csv, file_name="GCT_Full_Student_Records.csv", mime="text/csv")

    with exp_tab2:
        st.markdown("##### 🏛️ TEVTA Official Fiscal Year Comprehensive Dossier")
        st.caption("Consolidated dossier containing executive metrics, placement ratios, all guest sessions, industrial tours, and financial expenditures.")
        
        rep_c1, rep_c2 = st.columns(2)
        with rep_c1:
            fiscal_year_sel = st.selectbox("Select Reporting Fiscal Year", ["2024-2025", "2025-2026", "2026-2027", "2027-2028"])
        with rep_c2:
            st.write("")
            st.write("")
            annual_rep_excel = generate_annual_fiscal_report_excel(fiscal_year_sel)
            st.download_button(
                label=f"📊 Generate Consolidated FY {fiscal_year_sel} Dossier (.xlsx)",
                data=annual_rep_excel,
                file_name=f"TEVTA_Consolidated_Annual_Placement_Report_FY_{fiscal_year_sel}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
