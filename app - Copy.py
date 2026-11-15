import streamlit as st
import pandas as pd
import sqlite3
import hashlib
import io
import re
import json
import os
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
# UNIVERSAL DATABASE CONNECTION (SUPABASE / SQLITE)
# ==========================================
def get_db_connection():
    pg_url = None

    # 1. Pehle Streamlit Secrets se URL uthayein
    if hasattr(st, "secrets") and "postgres" in st.secrets:
        pg_url = st.secrets["postgres"].get("url")

    # 2. Agar Secrets na milein toh Local Hardcoded Pooler URL use karein
    if not pg_url:
        pg_url = "postgresql://postgres.mcmesfmsshwfxvtkhnoh:APNA_PASSWORD@aws-0-ap-southeast-1.pooler.supabase.com:6543/postgres"

    # 3. Supabase PostgreSQL connect karne ki koshish karein
    if pg_url:
        try:
            conn = psycopg2.connect(pg_url)
            return conn, "postgres"
        except Exception as e:
            st.warning(f"⚠️ Supabase Connect nahi ho saka, Local SQLite use ho raha hai: {e}")

    # 4. Aakhri Fallback (Sirf internet band hone par)
    conn = sqlite3.connect("ipo_management.db", check_same_thread=False)
    return conn, "sqlite"


def run_db_query(query, params=(), fetchone=False, fetchall=False, commit=False):
    conn, db_type = get_db_connection()
    c = conn.cursor()
    if db_type == "postgres":
        query = query.replace("?", "%s")

    c.execute(query, params)
    res = None
    if fetchone:
        res = c.fetchone()
    elif fetchall:
        res = c.fetchall()

    if commit:
        conn.commit()

    conn.close()
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
                conn_u = sqlite3.connect("ipo_management.db")
                c_u = conn_u.cursor()
                p_hash = hashlib.sha256(str.encode(new_p)).hexdigest()
                try:
                    c_u.execute("INSERT INTO system_users VALUES (?, ?, ?, ?)", (new_u.strip().lower(), new_n, p_hash, new_role))
                    conn_u.commit()
                    st.success(f"User '{new_u}' ban gaya!")
                except:
                    st.error("Yeh username pehle se mojood hai.")
                conn_u.close()
                st.rerun()
            else:
                st.warning("Username aur Password likhna zaroori hai.")

        st.markdown("---")
        st.markdown("#### Mojooda Users")
        conn_u = sqlite3.connect("ipo_management.db")
        c_u = conn_u.cursor()
        c_u.execute("SELECT username, name, role FROM system_users WHERE role != 'Administrator'")
        all_users = c_u.fetchall()
        conn_u.close()
        
        for u in all_users:
            col_a, col_b = st.columns([3, 1])
            col_a.write(f"**{u[1]}** (`{u[0]}`)\n*{u[2]}*")
            if col_b.button("❌", key=f"del_user_{u[0]}"):
                conn_u = sqlite3.connect("ipo_management.db")
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

def generate_annexure_f(data):
    doc = Document()

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_annex = title.add_run('Annexure "F"\nFEEDBACK OF GUEST SPEAKER LECTURE\n')
    r_annex.bold = True
    r_annex.font.size = Pt(13)
    
    sub = title.add_run('(To be filled by IPO/Career Counsellor)\n')
    sub.font.italic = True
    sub.font.size = Pt(10)

    fields = [
        ("Name of The Speaker", data.get('speaker_name', '')),
        ("Qualification", data.get('qualification', '')),
        ("Name of Industry / Affiliation", data.get('company_org', '')),
        ("Speaker Employer Code (Career Dashboard)", data.get('emp_code', 'N/A')),
        ("Topic of Guest Speaker Lecture", data.get('topic', '')),
        ("Trade/Class/Section", f"DAE {data.get('target_tech', '')} ({data.get('target_session', '')})"),
        ("No. of Students Attended Lecture", str(data.get('total_attendees', ''))),
    ]

    for label, val in fields:
        row = doc.add_paragraph()
        r1 = row.add_run(f"{label}: ")
        r1.bold = True
        r1.font.size = Pt(10)
        r2 = row.add_run(f"\t{val}")
        r2.underline = True
        r2.font.size = Pt(10)

    doc.add_paragraph().add_run("\nA. Feedback Regarding Guest Speaker:").bold = True
    speaker_items = (
        f"• Teaching Methodology: {data.get('methodology', 'Interactive / Presentation')}\n"
        f"• Question Answering Session: Conducted (Yes)\n"
        f"• Overall Rating (5 is Best): {data.get('speaker_rating', '5')} / 5"
    )
    doc.add_paragraph(speaker_items)

    doc.add_paragraph().add_run("B. Feedback Regarding Students:").bold = True
    student_items = (
        f"• Behavior & Discipline: {data.get('discipline', 'Satisfactory')}\n"
        f"• Student Engagement: {data.get('engagement', 'Active Participation')}\n"
        f"• Key Takeaways: {data.get('takeaways', 'Practical Exposure')}\n"
        f"• Overall Rating (5 is Best): {data.get('student_rating', '5')} / 5"
    )
    doc.add_paragraph(student_items)

    footer_p = doc.add_paragraph()
    footer_p.paragraph_format.space_before = Pt(35)
    footer_p.add_run(
        "\n\n______________________________\n"
        "Signature of IPO / Career Counsellor\n"
        "GCT Samanabad, Faisalabad"
    ).bold = True

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# ----------------- Helper Functions: Exact Approved Annexures G, H & I -----------------

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
EMBEDDED_ANNEXURE_H_B64 = "eNrsuWN4ps3SLhw7E9sTJxPbtm07E01s25zYtjOZ2LZt27qDL/Ms7LX2u753H/v3fvtHHX1311lddV53d3Vdl7wUKBgqEAQQFBAQEBHQHkmYtBcwEBAlBBAQChAUiJaQrY2jiY2jnrKbnYmDDq2rtRVhHhgIRQ4QCND/tP+nW0PiqGUvPaLIY9B7wHRTn2QLRWHfcF6MCY45zDrVDTvSPKxFxuCmyxlJESo5rUwdwxDT40Y+8LhdGyunhmDoyHIk/AyvVJdkMjXX5NLgXaHLZHxoz4EgFyzWKr/x7vT9vF3LNSKK4mpKLXiezxb3CDgmktJht2KhXTdUgVGKA91ecgpo5/njZbMjerchAjUCKzavAJFk1PBtmretBMjRUPygtVFSEsN2u6oi4fzzzEGnontrYGKik5w2BYFtid/1QmCZPbt+sUf+NFiSECMSsF5QLm2HPdHh74RCIhP3taKgWbSg+KkjajaCjEBMiwPRYpl+KJPHV2vcusijseJpbaIUCTveL7IDFtLUfFAlOanb3RpLM2OEiG/PIdPqLoLCa+ahh+RLL86DoF+haYFrHC6E1bvP3gxR+ca4BGBGMsNTg2+NmBcyEcXLoZo2Z/VErB2+xXOeQoqX07zurgDWVUo93uXoqiVIThMuIec1x4hshw/u8wi8QIGAPj5AgeT/bdfix2H/uvzsyX5uStjPXatnb2LlQEf7R/7Pbv2f9qdVxHd/7lY4/suAt4DpdoWVbsx8WTqjqVXMZUgoLQlplNUly8X3N+6lJRCdYwOx0ZnZhCl3hwxRp61EnSvsnyrw4e3QC8+6vVqANlegpKC2QgbXBbsduPSDB9O3EWIJypGACA208ElsiLin0JTxLpR9KuGH/WJYi6+Uk3NGJPX5WT9UFh9TvJlyUhnmFyzFtTwgy3hEOoLxWiawvGFM+WqTmWhiY3o26nT6lE8SmhmjLCnvv2Fb0dORZsQZVsBgX7VNn7jndxb3hwV9g86FDqMDeJpqI2u+H4YfYPGxXhJpNOht2+Jk6AjcAv+n/VGU68FEAwME9KQBBIT0OeJia29MZ2xr5GT9mdr+5LQz7dXo1WasDropF0RPBWM/mYrb7vXUoOjhnHRHxEmyIxUBqYZKUWTwsJvIR0uCXI9fXu3cEhKSdfuzkWdDE7jcoOrIlN58zwS8iiSF2d4dEcQ2EPXcwsMicq69j+wpTxRQCcMil4s4l40bu/Zw8UDCFH4Y7RxjHnx67/d0fOtfQPYwGKTpxQjjPe3SXzDaKQeOczYn3ygjBd7JvrzB7ZtHbj5RXORfs98ZfKCbCOjtgWdlFmLdGGShuLGKmxJeqti963A1EIL+KqAHTwS8oxpjbtAwBLYLJHZML7L59YH6isOKGKBIruBIns9qWgMNQwLHaVoTUoYQWVtkuGPGN4cV4AWTWH89/hxoVmo73qIIgh7Gzf8GkaJtmuANV0cn8pIGVAwUkk8Pc52hf2MJD9qzNwrM8rUQHaoW9Swq594ItUscxhLVcZq6zlg+yaedQq6cl2JULUQEGrMRRaoPy8s1JB930aSetYU6ggutlr7yngkTiSePJDv42B9vR2ha/KABInW92FcDd/5xvUhwVH5VxFrWIphUP/1ErSkiv+7gkRjGE/a7YqhIGKuKepLGIgkZfhuqZ6wl9giaZrwcPKENUAmv1bLGHKZ70PRj8rTMj56fBiq7iaq57NuKtYJSX7SrpUEJKR78EQtb8O/4I9hqboAgZ9+2LHHuLTR8u1H3InIokQZpwnJEE3FCCwNEhvSusPoDAZQyGkATu8Hws8O/wE9W4dCOkWleWRvKSU02Z7pn8dVVutgkKD6aRuCKCUml8QEB5RL+jPTh4lHiT/YkIBymwC9SuKn4S5RtRsdzPmUJP4Tl/P3zE5daSQK6kGAO7lLi4dxi/eG04zBnHDcQHeVOapVnggx/ar7FzQFj2UEQ0LlzX0qGb0FFkEBGCyAZpBGX55VgS0u6Z04LZYFkNwSuwQ5gxuOLHHh4pV0TZBeMA/2QzKbnGTO39FGoeC2d4Encyvi65y0+9j2AAteEKImPJaEtrsdX7IztozIYYx6mxb832TS+q/VrwNri5ofn4/3jx7S09oa98XJ+RcuZZnTbxkA5CYV9RhDhDm+QWygHQo8+2CvdGsuYOjtB7fz3xuOy+e0W3APJcdXOFoE1KPtd+nD9rqgDbSF4AoO07ownbKy0sS0E6vHiaDGkBUFBRtzMxzZ0olzXuq+/AcF9UMhSsbHoylnolV8LWsIFR/iQO5uf1KeuIMFFCzm/KBkKeQmmKs5BJ9p/XZo7/gESP2ydEjVDOgtjgV+vwBrQPopAGIyL7fAxRkYI7WMEOuJtMvnsq+WyWcSirGLvGHf8Kq4CcRddf3ZJAbAT8BDpkn1sGN6Fn4DJbiFlg2PYWwMiWyMXLzHrOMlcXRZhQRVI/FYI61xqX4XDxuRn1sPha+f8gGU24wZXJ7cXNjd1GUEZaoRoZ0zICx+ALkBT5m/KEez55UQMee7LWK9EDwTC4RXPKK/SQRzlrBDvMdswk6lssIjXxfpgI28DeMhI0ZZb1tqwkVqSw8GUs019gBCwKSgxa5yd6XT9GVMEfqqIiK6y3u3AdwXr7uSyHfO03IznutL+y+FkgnBXMKaRFCFtsFhJZoZgCaAAJo8Ox+UQ2JVskUT0M6KrNnLsr/thEB5QZXkRaccVNYMrLZNOD3uU+96rXnJffkva4+74j9PWhZWqsV0U4ySUz24jmFF9RoWA0eDC3inClEJqBciKyoOSXcp5hXiZBRMxZ/yIw/HHJNEH8yfgQAjkCD16bknOqBiIMuk8nyO+Lipt+VIYmk/LcXrPacA2vdYwbKiSRAuSmOn1nNKQ5HZNUfZF5VnhLhWqhfQ8B+SaW2e7wVpqr2KKDsqrtN512Tho0ODnbkyoPgOLSE6LpilyMaVhi1HDRDXKEg20u5kfibq2F1W15aXtwrM1OsSdo2A0WL6VXwmFVW1rwp6iPg1ktfBTD1qaEoT/SRijXqGNqyK04oLPYEcAJw5g39ZLZTlExcogb+vLXH8ZzuKLskgjyeIMl/NZgnUqYqTTRW19LBScGht8os14w1TqibEaknbcr+yYkzWu7CposlWKXK11iTJDdv9FU5HrsKMLGsFVVZ0WlfQjzpaiY8cKgPTcg9+dJIcLWjEWFuMUcK8NHuOJUSE+mcbwlZXIcmEx8p4jSXnsCFSdhJSjYZNtzZBXUzMtMgGGrw3loXjMyHqgMf1HFtwsGYJ6EfWlKo3KldWthauLkoPex+vuUyLl7yor7qQLUxWuD54WZ1QZJeFut694kt/Pz7rdygxvg60v3IUpBWQH3cqE32XymQJwK04BHJPQ0MKSxRaY+xp2MKZEI9EX1WkHy00UjanQPmHaxZp1pxKZTimTkafKyzVWfrxvHfJVs1b9X2MQHWkfGewyzYfKhEQzYjJOq8wBBtFYszoBcCGBL0vexBcY7AHT03jJ3LTUZSQXjKuw5DzditLr9qtWQx5+BcUaP/M0Hhxw3i+E3MOtQ3+2SYGSDh2LyyGZdNtX6qCkxgTdLeUJHQWEAsT8bYgxpUfiNqolgJgezv3oU3HXPUjl7JU3XL/OVhRHW4ex8ayr9Dl01rBMwcnVMQ+7/x4rPmqiGZ2pCsv9pHOMHLbldMXZrNuzpopAHUrgh7b8iREx94ibFVBKHXx1IEKIXg/FSMxFRx5Jc3A+96/U1OiG/JMWyxLFxUtGPpQTYmtnsPFTq8jru5KYR2Y11X2daSX/jOgXYTk0NN3TimJe9cVHpU3ESKIshwOaIX2Q24xagu1eZhiBHdvV5c0fyjg/z0ZeUTSKuZyUqX9uL/XU40w5FuAt78usMrRsWN6Js8fgj1SKfI95LDK0tBz3hmxqGbFoePGDVzWBZ3tVYpARi12VVpKoM2521F7pAcFxysURxUrCE/XNZ8HWfhsQhmzyEgdjudfqqcbCkp7pJiuY8rJRlyajNyOYx7a6DoEf1Wohhtl5acDAuqT2lyLn+nL7WNXYXhV4WVPoI/Os41VuXP2TIZ0Y1c4ieOBulu9MwVCZJr3dwshW134P3jsBT4WkW2SQw2OpdcQApwlPQPnMXrhhFLzARmerL/cBC02ha1yh22TKbBvdbhdkthp0/Ih3IgPHjOJgZHzR7Z5zcJIa3S3+vXml/7X7DfYAF0D8Bv8ITDPJF+K6sibZcL7sbFRQpNVirdlahJ62YY6ZZzUTdxnMVnXAjWy7huRVlB+/pZXeusUjgsJj9GzTt22l1fFOeKLc83ORQBaVWuiBjAc+qtOP/Y2LrW+Pfuu9/e2yj8G+9VQmXA3Se7KXAbybJvNcHeYtS7I3uTy7izMsaKm8D/TCbxHH4Wj3FE2rQnQ8dm9WmVFeGbwxMmU4FoZ8EtyfYsQKFpB40btwOjlemPDoK2iYWOeSpU5ism0cnpxMHMgQnryIhMtmX4wIlmjAXLumGINInHFir6jvrBF2qyDcqMZB3B5CCjlUnL5STtjxpZZPR16104adi4ZvZlHPu+755gPoP90tFw+gsnA+b53CYEBAuEBQwH/dLf9WgP3rDfMfxRjwZzEG/P98MZIwLttLjxh4Kd4W0OiyU8AvUCJLB1WqBoWsxrb8GroUfQYf+2swzzI+nOqkexDvuxmvR3v6eYzlMobmIiYNYZihfomDEnML8lVW+xo3iAOJE5FWvkKX75em4/v5AzQtiY7heZnDFJdpWX6ZOZ4fFzfYSZ+3AnlAwow66XBy007mGJ9JkCfa4HebQ/zYLNc6HaKui3At/eWpbdKNom/puNkyG9RzjmT+XNiaPsA/JbSIoNnn4h6BvXnfiYy5n3UzZj5GAlK+MVQIMh7pnk+/XMvi5jeAxQUGDnmIiXo9NSB9DSdLYgu0vJZYILKQzKdMADizpilo+jcjQF9HAdhIbpRS3iZTm8HsMRzF3hVwk4hB+Jbf7jD+05+LSDMv0BACCGgVDwgI7R+Fi6O5ibXJ3yTDX8WLhpQtihhup9qJC+Y53msoXjwUZH80vDWUcFjuDvsG3uWX1HWN8ciEJ9XuWsXCsoaTXPH2e/lGEZBSjBeLGcVG5R3qOz6m0/sqHjII/3BNz5rDa0AnAVNrh+c+ielRKJSE28yNJUY5/tewSdC0QJpL1U4Xg3EhxW2Srkjygv2osLsqK++nbpO23K/IKGB3YWtdNHSKVOslygUIo+q6e5QS+rZDIpFo0iLq0GC8EB0M5XPl8/OSkRY7uWCDXQ8HduJe/i9phb84DDZMrpIoSjE8yb4wf1fMo1/aqiny2VN3TuEXS2nkjkW52uXlUMA3VKdaMndTub2mSk+m5kBSgjWHdeS5HT5i9N+qsTkSZh+C/t4+8FJ4ntEJTM8Sc+rcmmlDc2gGtPTYrVguhbR6VDm/gpIFHos3cXuhnD4viqW31ckYomf7a3X8HnnD96yGIUSvvIq16YS0DXjZwEH340G1uF6zBycWr2pi2yN5QQ3HpA1YJJIsxMEcb5KbkcN4aJmF3BN+63sCWEoeVdBZ3V2YZuJ3ho5REQQgsmb1abDSH5LBS5DakEyciMqwXThgGM5n8XSHPxu+kXgJKM3f8MoyKPeh3uwcx9Bl8YoWP2ZlbAD2797fZj34uEqqtkm+KfJ8vEd0EniHv7u/rt9C6v2+rgv9eCue9AbcS7HJo/lNdnjfDSaITXY8v04S4LNvaYEghF/xRoUJsw0xrBwyYB4g8kPszFO8IBWJhFeSJLJ3pw34XZSVBHYwDMBx2VkrPmxZA9N0gbdxtmgpPhEiEyf1t4Kthe0g2yU7LBz0WMxG/1S566GAXu/pLwldHXwaIm1lQtm8oAjSLqL9mRmaOLdmhGiPGx9IUfD4RET0pY4r8edri5RXNykBlWNIHmuRbPLdcU6ijFLed+bg5kgdGME2qY7l9qiVjN9Ltgn5jMhlIEwYUj2SNm3hYjCDslE/xsKHTtVFaEkRBnroOLSpQh3FMcXWfcWP835TDpA6ueQ4Nfd2jA78wKJE5XEkqjaaiBiy9l4Ks8vgdCqSUO7zJkEgEmFY2R0pBaNTC80SwrWro1/7RUkZBDhb6RDpjZecBZvmIt6mgu4ZvWc3FZKGux2Rro62uKkbwBPanIDj6xyAr7xZ/oWwnnmfabWstQWmNj7GeiLSvPywtv0159LbZWMg3ziGX4A95zL9FnDwmm07EXHy/Kp10wCbTv20l2e91Adruy40suSHzWSvk/ZO092WAxlhahh6I+WN3U2PotYTBJLOgKztJajiiLiN2jcXGShvWiEUpki7Pd3zjcvtFJ6dL1Hnuyw+FXDUPC3C2Y3ktQGTmZFprltOJpTmNNxPtCb74/HA+W92EhYs+SttDWvGxcy3CW4oJxAGb1S5TZq/8ogU6In9HlpN6KFHteAUTBtZSHqAf3zDOCvHcslqZYsIs29CPtzdgw4oTXBaHldGkRphcly7WUi/LwrxzGhqI1r7qY9Z3CV3IhugkdLAracBOlnVjxWxuK6hUi67cAhHGyYepLyqLvpzxMpE+FvqiZ4tawPBxCCQGJOGbERS40pnSdAzZfBPuJA+NVjVolmrmBhUJEgFu0gIK5IHhqP1mmU0eejMjb20LKgh63rQevtt00YGwAhZOdbYDJTZcpXeuPS3BjumCSf6en/3ZyvRx0VmEbVXt1RF6d2PVEpOh23vFuSjXITi5i7xQ/PRXUER3BZ5bIVO8q01tjp+OjCTXBVB6U7fhZb4fIvNOm5pge4gofNZRGiYTFkR/1iepGU3iHrNBbT6kSFMEy8yNbWiBxf+aKjT0xOd+QcPijTHhDEkCDachhtuhwHCuaCXKl1tq5W2k24mNOPjZDRV85egfP23cWX1AyDa0Hq877La7b+v9lJ5LJVcR/hfYLHZ0HjTdeUAngR5MTcO+AJfdad0ZMwMOtiBMs/zkoo0NKBtP9ATRzN/ONeyT0wvP5NaDizTYTy+TEyIkuph3A5tn4x4hqnCTBKZsj6KoRKsgS+vXrkyGTqLFNHXB/AYwOFf6WfApzbQdHMazsxo1MkFh0KQJ9RMuDIDisn5xh4rgaw9ZJvtWQ2CF9lXM+13Xpb7OILi4Gi4yFnlftBCWFEt71gn3qKneoST2u7UkSa6tI6lto+2WIXxtwePzg2P6vx0K9uR5LbPA7uQfm/xf22MlB6xef+Pb5sr3XiCaj5vO0Bf/uVtmoOJo6OFjZnDn4TUoLZqs8uK6EU35UN/pz3Yssq4oPFUZONfgDOncZcab54eRQFDYjKodbfJVSAsJTEn0nbe0TeXkcrUuz0W+3x1nXoQAHrgX2NoyqHWxNIlMsKoeGa6kDzoyl4LIZj3VYiD+mqM08DAq+39uRhhsLRL0T8NE0iEqIg5U8ODcvTYnGSGp0hcuxYZA8gGq1i8h3wzimJHToVDGosmjgR6OIw7Snv1S2Vw8HhFOMZssKG4xEiqvlapdJFVKymljXckuE/M5opx/0z4jqBlDHG0Q6gfmqGcpMHS0WMdcDMbSYit9Q6/Q76L45mL1eOQq3OTq5zrSM7et37BLuDpNTDx8IXaj+CjMwCMJ5vWnLL+xZL7fBR3c2SXlxBi0dMzY7RELwPIKC3CMO0u9CR3gUwFircePZ5/RUtoz7007olYqkuzMFn9m4njQaToCLpEg9iOdhzLLmtn5PQl/QL4z/q70tEwOilNOzLu6Qe4/tUa97KW4AvxHqRqM1GlGebMjTOYAYV4IXb5FdCSCb9o2bcgIVU+97DsRu39Jg0rh0bfOb1eccAIJh5uMekB5R3LmNBikyjM0BW6ym9P4WHROmTlL1CG+Djs68TRimHPRcKUDWvxaPH1a0cmZ2xkyfQ37I7GRrmQOZoqIIOnXI6GSWeE6e7NIEJ+FGcn/V1tZljSfPT7e2/HahCc2IOXwVcrSAGjvglwfGVWNnuUgwxWpkIdO25x66qCyPpzRrd8wleVQHa53D7Btz7cSewD4qqxTQR70tnSGQk3WXYTTzJx3b8/KMzsZnx8GRv8HqH2bxr0FRlfjpUvDCY/hETF1uv8cXhQTK+5jGZKVh56CGLai+zRVr42HYbILkUEOh/RABDJawbf8zBSl7bzFgeiBOk074mycAYUbK2yLkVO3Qh2QqpwRL3EEQF+vy6Eeqa0ZNZ4ScK1fTYixvKv7XmORQ4mLGdas30m7+2BBZgrn7ChVXWrs5N2IUZ/0DG6ZDjR1T2AmM3P7UuR6AbHYbAlzc7BZpa5jIGyUc+/gmpqfUOSYRcI8LifwaKMCuSjETe4AuttHtAxViKFwP3ynV8+8ZAhsRPY0f6ImThPe7tS9/Ax9FJzQzTgUdKAcTPr99zcwKvaRFmdaAxukpMSDtrgI46U4hCXYJS28vCAOppdt2NO2RAvQfOEUnrWrrKgw8m2a4hjwncdbbhOK2UtQ5I9bUTm+nfh3lVIUklxiPaoJ7vZxB9pt1GWPzxjRuqyUIcnUjU4ls8L86VJPE2ujBcIhies6dWMbKKOERLL2gOc681NquKc7riJbzVTbY29qWEjv5mb5MYFSTaFYI07KHJ3TzFKkX6esmZiGV6W12x5nNDuYJPUw2Pgz/yKPDcsVA8lfeUvD0E3xo05v6P7t+j7J+3kdiPFn/gKSbBEroocrdQQakW5FpNPYXWCKPdvSAPcWVINXnrUfw2Vq8+VE5LKT6SjyKyMvpdtqMqCaQORtsGS7EIARIyiEdJlbefK+6r3q2xDqvAWdMq2VobhRbjaqCTKj5bvv/3wZxKRJ/ZZfuf+sljLpi0L4nLEpcse+ZULP5qYLk2s2/Ws5WFTpj7dZ73q4z9+LHg/yYUo++yBAf/5BE4BZOTk4GhrrW5tRWfhaGLN8LdP4CRAFDmE/+c6J1oApoceVbha8X2jqAWRUqaQsEo4kKJZZmaTNXYvmQTNw4NYt3u/zQB/FndzAp0Objr2sIaGmmLq+0kUAQf3xeE4cHhKwWUhaQyuaqSdFY4KrKb/gPpee0zjDn0h82KkCE9uQm0BltcFBIxsEW7vkDlm5ADKcX+unJWaikdRZF3wAOGhRVD2GR+VH7pTj27B/MFmRavTvpK2/TW42jHjGJw64D9S0KUJabX92VP5nMX8LxTI29vaOfxf8ZAR220ZSI8YfBnUETDGg5Ok0xjxo7a8R1cK1mN8oAQ4IY7k588ZF9uC4vPa4u10gwNA6s1d7uGRjmbxFL24M4PwDNRZEhZWW7tzge+Qm407b34cHFxPqbsnhYfL2e/DifwkDDonVGyqOanerfLzmTaK+MiGwwpipya+2GJ2W/GNwpAWHj7y8Y73YIoSGfZdTOkNOjHJzJSCA5uQCGYD9pz5NcgG4n4DeUFWClWaqGILXDPlCoHqavdgy9Hb1revkC4xWg3zZR+aNies1Lg0ekVyOSoSHyWsHR4i9/8xQWIVaJrNfiZIxs+qDfkfCdLGydrQxP4zRf4hbizl3KacFXOwU26FR1gcsOTPcMl5rrBCMpsaXSRxTtsFXCzQCOwnkVbdWqqz8DtoGYLcAiqSv6WRwPJqdDImuzN7kEyvo0OBcXmrLXR2Txb0VzBrt0SK6I8b++K357bm6waJbnlkoniyJ9nVbUGA+8etnF6TgOIM6YTbILQZiUjNeA+X9bg6+pE7UhWDSD9NVC83zQxaK7xSWXG5Cq6hV/pBnA12V6ycRlZwMwX7IRIpjrI5C63nqViEfcN+APjUTC/LCJY3F31XP6LWo1M1NCOk7xOMzJRwd4RCNG2FpW+FAKGNSoAxAnGmRgZTpDGrxAU9slV8m0HBgHPeOHyWKft2VDlzcej14qtIMYUPTT3OlwnEiGSGsOVoJngxA/mMti81oTErKBSRtGTNnDLE0voFFMeZqgcS8jGdyzEYRZDzTy1LMV0Fmw2enHqDtpGo0lHjgKcqbF6p3vidQInqYw6ytRgLxYpdq2y8I3lsW23qoaJRYBxJ5U5rCPdgTS3LdBVGXF+vserG/IK6tnSoSHxLmD5vYQKUjcnSb7mekARk+HtpxIf+WcNiwGMQXdA/JnsFGdrSckNCjtAzms3UKWULw7f6IAf9sziu5lKI+ZV1A+lz64RQDZ+QykPVGzrgipln+tX7prDRADmloeVZYfxIlcSGjd3i2Be3qfo7oxmml35c7gUGD4y08j7js4xcUzs/GmxDEtvRfhXDqjt1EZDnGmYwyH1PvJaW7d1alYo/MJdoFuH2JOcGIVcEvs62NVYYEnslO+1/LyVQFeYW3hcuED7P2gSZrtPngFm+JM1s8wAMCIVtknFwDzF0Ch7k1bXQy6rBHSJwiFbnReSQBKWwLhF2VKLfTHrNOU4LSvXsXLoGDN/S9pbFqCN017WcDJiLKu8Sgw28TjOWo9chwTPQzRnPYM5jKbc/mTgLPCYhfTEKt2QmeB6LYZx/gG1eX/xAL9P0V6som3gfKhRYQ3MqfbniaObz/7YvCDJJL+zzNsUkGwgLVmJzYPs9sW2onN1R2WV4cf+ujtv+ufMEWPSIAadicI2RO7vYV64yxMTphCitujqEQzDmWZ8HHaHL0ZO9PR5/GWU17m1sczdPF+pV/DdctRLttKfX7y2o3+gB+/5ZJsCDHaqc3xNOuzo7tKuWC3+at3eoL9jVtq1EgPaB0Kbw9hyHav/6Bt3d6CLX93S9uPnhU7YK4abzbSYQredaBPM2YmY2mlbpa4jsb7lNlT5olFW7oGax7ykFcHML+LnW7ZMhezLs2bSmty+GjKqrpri8Qwu3FreSMzFS6dkC7maBd43fQ/a2sFLo0Oupg5zS7w6VrpwKCJL8+QKuLqC227t7JiNLjXiVoW9BETbre/QOz3jE3WRdp1eIXvkr25A0nwMaGb0WgbWrYuBnECKnFRTfxBjSxNsT8i8lrhMEEJNuIMGfFYxWyWp0jiSKO1exLKcZ8Z3CLOj32aYpZ3fFEi5xeK3MOPGk3X7wmYCMhgC1IPo7l6wDJ2GMKZXsLr82XWh+zoU20O/3HyRarCI+sTJzol860++HLau3GWlG73NMWYSM7NaBU9X/ALXkDEawwnk6l8fKpJFPZcUuf4QfpV/1Y4zV8n1/LtT6acZRtFrriKaBrJmkAtP6AfmfjsBfa72Gd6xAQKDC4EAI/6wRHN2sTP6qEM68tBJ2WeMvAGFLH/LTcmYE3JWA/loenJYfK7VNg6cXqW6qEhW0suUFnPEj1oBFAuQ2FCBokvhHR5HmcgoG4L03KeSwj5fLTT1ZhM7lxXQHDh9XttfRrfUAL17t6Inn0RCXjycrr1+uj3atl5ODb8f2TXxtgN6OdsBl5/vgx3O17ebmhU/G2WqA1JYn4H3/anv99u395Q3w+Ku3/q2u823/xjJ7HSGbdjH5bf3t9mX/bfT95f78vrV3f9/JxwU/M8HhY/fNnD6lHXZIW6VEOysrI6C9HXdD9+w2c/3j3TJbOzz59nzZrT1orqNUHuD8G1+vOjnAq/V3hDfH853uwwTHV6YMu0DdjN57+1bu9Y8dWrIUJOL9DV7cmk1dWQa/Unw/JDF95xqfxfxaKPM2l0vpmR/R7o6mSHiHKTmg6dL+XdcyEX3muBO9gMFWHoFT5zaC7PuhIEb3y/XWzsvsx1+HzP3FCfj9Ev551frFo9vf44sqBxsshK9iMOSC6UIufZ7sO7QfY8iT48uTEKoMeeh7Nj50nl0/avTON5edYmi66eTILzGSv/s4yTy6clidF7+tHnvfvre928O/w70bfqgn3749XJ0D3s632LPXn1sCOtgvngjUDtNvJ976NjYvytPr5bybYavcsvk89FA33JrXUZIXJ6bdHjdw7h6KAdu2Hsm3+QquJrsyX53OG/h8oKuXDy/fryk5UvsKxZIvQlwApy7ZqhtRlwGNzx9kqq9pbxHwry187FYZb7PrzngJpxjceeSDtwH5z99edxfZ5HABvwCkmIFIVx9RQx2HKRcv1IseXjeZJYDSt7v8ctqH69n1t61eLtafHtcEfK/2TrRqiwrerr+0ITD4lukmxA5dGybD6LKyM/Pfqs0ArzHpXpHVH9X9jgzNds1tmeubgxd3yyQ2g+uXkIB5h8yL4sN28oTsdhzZjV+bGa52zoAP2DB4UInQC5+mR91XV0xTiZrj73GYbi5ryOVos56jFZleUt9zfqiAypRhm2vO848hlmGa687zR5RBGBqP4A5xEKzLd3VMIvQq0OQR14pacyIbzRYR15J+dpSmPztgnx2RxSLiaaPPDst8jkMnrrFmivxDrfDq7TH6IBTm+k4sUL4WLrC+YwQeb6RNrU/INI3zIX1L/waO6K8ljdPZe6XKEAJtJq7NqMqzY/vf2XQugG2bDxtZax9nCwb5KKFJWPcHNJM4KGPPSXQwQyPLEB5Aq7pQSlzP7vN6rzdCcnEI+eHo6nwHltwa74dpNm31Q++7q+MxEXLbyuOv8kBixMyOFbmHeV4CK+n+AeITou6Q29VRYn+bjovOnOm6STHjhyBUhQzBgilab++HyD4B+ZPkt213huwVk5gDAm9Hm4O2m1Duik1ZcpywhUl2qgHzVCNQ0QFcbHrGgyHfOAZICOhUg1qKnOC+Ul9s5F4QS3HP23uUCFtgFpIO8LnYJ3R5psNVs0NvQ1reVCe4+e8u6bP9HnCG3DFyE8m3GzJAY70y8kxeTyY+6z6/L/SyCdp6GlKs7MdbdivoBvnweY36GkKRVXUnlY13q1Y4OlOZmEXgR5u/uosfoE0LMdaEZv0Oe9srGtb1f0CCnPR092iDjfqpoMi4IgeYTc+eStNafbepMqgFVp+4DzaFMKyDwn6tT55w3+m78rz+Nk6OPsvIVl6hJpipy49s8HODSQmX7fU6Rg1w0lNXPZBYMXrs+GxKcE3h3aJH11FFKuPdJ5xnfNOavHj4Nm6gTikiyDMXQifP/drgw6I7RVfc7jAQt3rRmtzu4N41Ssjr/X48GDKZrIZg6h7ZTdTcsXaZjMGcJkB69fE+6v0qlwFdpY3o+Q5q0UhdcINI3PQD0T9S7theOG+vKHViWSAsN3LjY8mHW/dhkfhLo15V04fk8Lpo3pPFx80JOR5U0/5Lv23Vh8BQMGa7z0V1tBpGD2PZJEUfsknv6upu53nA73LA842vItYLYDVznkCjbPRigNgsi6xgKLUYSjrNxvXqDsOgltL+IEImrlfz7KChcx33mFiZnqdiP/iuZvH7l3cQ9+TZ/Zc75x4jfXHEqAAiAWZ0XVlbukOcg3Asnj5wJvelUxFlLFGEXJLBo7fzE9zpwO5QAwl/pF5kUOJgUlBROBWgA9sjktDUzqIEQtP3qmg5OEpwccTeGaa5q5i7RHIBOKlJTIZvPEZ9NXeRSL0z9dcP3d7u9gn6bYgJhGGFDtDxrn4Hnr/0dOkW3+4RSOJbE95MYuQJ70vf7Nu53dI1aM0+SG3A4nWjXgJTrAB7NYGvdAAka0t8mjE3ImXhx+LZfX9lrQjQL61MCuwZtLC2RlwckcVXoxY+Ndl0xfleSRUS3tURyXzPq4x1s2az8lQe70eebh18nu7RZW8SjPUC25lE+nLrqtw5oYf1dnwRR03ZZ/m3D8iyiZvSDDjlm2PMnE63BU8H7Ani6xwqbQkEDqqeZhmkxRcZ1BvDMb/l2xKWLU4muD2WLqr1gsLpJ2ezoXYlEm2zikWmx3sMblEpW5QK/G+5aSRTR3pP7K6y0K9mFOqxnt8sB4g79QbD95QdxyRkRZar9khu4M7UAGyDEWecsglRGO08cXg2Y/HmpJPPnbYYnTbZ7lsFHaWnkGw/ODfPIIkya18CyhOT424VlteHJCgAeBzQJpAk5Ll7YQW/vv8m4UAz2fz8dVawR9LQjkaAdNeG5tOZ9WvnMsSVOgumTwvLGNgeeeTYNyOwKfeGpAF+nlApNmGKUfnCOTb3RrkBPok817qEG9paaB+Fnkdo/3KRSfndO/RzijB4RGmaVTnr77h/6oEfPMysCvxvk/sREMEjm9VmJa9viEL7pfxwScV5fzcrYY5DJSe4Ku+GvBI88h9Xvv+F3QOf1Kk3Kf8MIPo07IdJJSfyT8C/uWBSsnZO+LnAv9rER8fGpPqMOffmUehzXRAG5azw/83nvzsxKv/73z2/uWNyFab0GPcu0btP7Kitfj+iSpmrHq5ck1A0clLcrthrmHv8WkGXXI3QExViFMw2F8oeQ37YIzkXP3gi0XKWBFUJme2zuXd5JaHIZtXS1zDhYzxS+COBGyC1xu51Fcng8cg8+YrsMpn3CM0BE46SPUq5p6Hr33PVHcPs9/xDO2vv9xwYnpu1V/yhAgrqXuG9RJaQTLiCrIBUliLji39SwszPmgxU6TltFnmtOW5wHYVbVEeRcSWgsDRiUAwWOZhvlll/Znkk/ptZT4pZNBY5AbW/NHAgjCmxmkvcFPUVbgv08w7oK5tK3GBa8g4oj8LTioGQWcQgaD9R6nPclCklHMoFNJZZ/lV/1/AVHo/zi0gjB0dnkaOSm4NjKRauJN3uO3eXUN8RbffNaFaY5CC3Acsp7Fvnb0jMARKWaqsQTD3NLSyY2kNVLiQvzqES5QCncZMcPoJVlZXU6rFXpFc4zSWNK9kOpirOobbkaSBBa8TXHPmniI/QRKAJE64M+SmcF8xlxgEeW3I/pVj1DxGvPeduLkHyB4f8qY0ctkJPKDXcJfKpb/9H7OUtCEl9/dsSn9rif7SpPrUTqf8hSDD+wv2vVT4Bip+AQu1/iuw1LdzAn3zZn4c2AFTXpyYGsHrrZbsReUFwWO7VmuFdH4PQLPWm8D3psrMwmuvGvH52d+1yUyfLDa8aZVUp2ntw0pcQsIMweKT+9dLUlgyhG1i+24R1e+saF/Dm7sDgdJXZ3bUy9tOIOgWWWIM+zHxpq3Q7Foed9yA06rHZZRHgEiO31jChG8Hd7SSn0nO6TZ9++Db/way6c8Xll1pWwcg5cFJu3fl+Fj1Z5ebTHF9ve1b9FFguVt2aeP4buQmWnP2Kiv4wYrPjcHBTdpofmCCAyB4BbvWD67CL/lLo2M+glGsXtLdt9PSoq0bs4wPHRwRyEiuwHneyaCG+FLE1iSMc4fJhTKYvud0FTbvu1Pd7GOgQyK80Tt9IHakMYojpgjjlBBC5oHwJYnAZSZzvwyNoGdmqBdq5veZN9jIV/XbtSZRqqkn+34gKJc/ScqmiCstQ21b98yV3f5Yh1xc/lMfnxksXSCK3aYbkk6CAk4bVBw2cwDFlkJ1E7n50foFPk2ShcK1HFgsoWp8iNFUZn8pJ/BlJZvR9QDUnFmG8V3tZ2PaVSVTEHQVMGLc91UnINoWGRcjW3lPmXpM52VJBnrN7Os+NrYFCmjq9UiTB7zFh/uulcKl6CC/XhkQ/Lz2Tcm6QchBIbEi+YIRNCSJTOQQEPc8Y2je6oN89PUFhtDSFLajimGLJP9enCmv0Z2hcAgvNIdUGAEJTbDqKD3ATfOHYTzWzMXVUFBwkEnm8ubWoJTGeGy7VT1lMZqoZk5WLL/cepd7U1E9v13vliYBRBDO5zk5wsXkAE897x6PE72zbs9nj4vaP89Hy13JywHC94uPstm/j6Jjn+1H2K2H7xwg+lgR0m5zcwAeTntxrPSGez2V4A/cFDwZSTbp1AeM+c59Gi2VmZtQh1XlX5f7TcAVfQ/XvDXldf0lA27pVMxQ8u514FVn9INJF5/Zjm92O1t7p/gHPjhYYetbTF5Qfd0n7S5u1x2GMORhcY0NCvDVOblWzLiwbBJ0dmWuwf275x+RoeAPb5wcn2+SXsBom4vQ01CAf+SEpuVgcUG9RwlQiyHG1emrp3z7GNs8U/fN9nRqGi4VXCgploItMRLZ8jMtk/DypZ4qy6QxQps8skd9ESB2gBYtVdeeB9LYC5dBFNrrSVxF7iVzzxYMo4g3iNe5Qp+W73kH5Un3FxEG+ZCk4i3Ul8QzPohZ/TMmQTTDEswccBvIgk6O+A3DW6m+uqsP7ciIAWIVtfcaYJF80aw+dvIw/RFOCvOa5dtjYkU7lEPI9OpsFRh7wdV2tHlOLekZieTF2B7y9cghH3IwyeTLemvjA49hyVmtWZu0uIECiMTbGlke3ewgRawMVoPCTvBqr48L7uoleFLd0E977O4MP9QxFUKjqS40k1AO1Q9KPVVfIftTmh39BnQhnewHq1Qnqs+dBJ9Hi7Wx/6mQnfy/M2eo9wEBFiw9vqIMDtZbYw+4+xSrVJvZsjtyDVVRMPI4oFRM/zdSWSTtGxf5mUvvmsd/RvO6hWnK2f/y0+1WtS+o7rrT6aAmeato9TLnuaREh1F6Q97qIh5hC/axw3WnURAmttqWiiZA9RVDevGWzjpXS69ZMocHdEBD0fbEqoUl4qYWYFd1WXCSRPMPt5RRZjE6IOOeHk4aVgsdAF+PuVtQp4ynfMdrBFbwU1h6pf3yDnw88q67KdURNHule57Vm3rZYLW00opDAeHdncDJkwWJmr8w3vLhFPBhHsFSBG0SqyIygkcBymuGaYx40uw/XjezOZpi9j/GOCCzRdpabjqR9uH6bef3z3ekxTYjTuAK4pIrMRxKr837ZgrZGBEjvUAqf69Nfs8Pop29CVxqNVqE5OZ/bZFqNSb3mG6/4WNdHjf13xPqvFSiJYzBhKmAVJSvsfRVQm6t0UWh89S70fRlQvN9BCHp+hUodBPRxviKMNkZP27PqFu0WpKAZwMgzUpMsINXkKZ/Km4o8SZpZVFzJ9jtN2HpOx11/fWnQ1B2Kt+atNcQY3QnrU+2Jof6FztKUqFgkUSdiyIakVdaEy9JU4wiXScd3jrAYUeN9v+O1g5DSpzANT8Gv0gXfXS6ND9qmUf8rsegWyGm4gx43OIa1YDMlJEbIeBMzClKgUeElKmjLq2yObSDBolua3MYR/jW/80Q+PL6nflH//fzkKScNLXGX2ioHWffLRPeFhcrQCQIAtXAB49Tcr2XYqVzWORx+ki+CeZT5wzIOoOAmPUCxv3xEoZ+mycGjrz/BFAYMmNQyeqgt3FlFKLPEqJYe4lGylgrln2hQpadq5OMcLb/yyoKeAXD7tSZ2TCa4TbbCVz8mtl5rKIW/45ukohaLK8JlITdphe/b4swnZnbKqZO1YsWf+Rvnr9eVu4PWFJG8tbcwSkzEkWBZ0V1UkFyzbVuGT2/SUaPsgkhEAg8cxTJVWhFXFHG/8ZQzutHG7WH5azJKrY2XGPo2LNc/+GQoI3Q7xbGr8Mdv9eddrVOs6Yz/0BD7MShTgBnzvhLJBg8H7sH/+j0NHRN6f/bNHclKD4b2VCMFSAfcCYdGPM0Y3VFjGsltOUiEAJNO5hTPHTg0ME3BbiOreQ0RfyBEzVn8S7KaF2ZyY7HvOTH/qhKhDSc9jJTP8MuZ/RhnVx1N0rtef9E6XODw8Dc1GekAC0ki2Aah0SdqE7MzpOXpFEXIeJ6J2bkYuLhy5wi/x3zkW4FQqnIjRd4fIUlpjJYFjeA5BieoBNFU08VPsEWe2EJPjg5eH8MyU2AZy3BFAWGBTNuA1G0X0EnrKlV77+vBcWymMWYJBpNs13qEr0rtRk796V9QIC+w6xlcR8IHRJG0ciRmDzttQcYGawndsqpDIY/Pj7AJheG+yghAbWKGRaEckNoPCGc6Gxa5qrWB9rpm1fx8OiNGkoqXcSMpz0bnVYvjoV60sLiimW808/qNf17+UZHqtC+RFvPrSpp1OBRy2bgtivloGbRFjSUGbRPPGQC8fpahI7tKMG7zQX728GjxAPA5sSTmwu+lfgiPnOPjIn5FLUnYXuUwOTXeWlU0ILzEMnBlKHEsliqwQDM7fq9Otj19e34afmbakdvK7LFt3tvflGdI2BO9KOiDxXc60AKKLVFu3bl/37leVqsgP3RvbcU1pOobZgr0kbiCRLZhtDMR3r99RtTmSBy5KKjn5PI7SCYMiHQIDx5xzRy7H8oEW+SZ7n/V9wHX/YpeIFxRzCPR6KFoMnAyER7G0vlaXEAkXoCvs8nxWux1v5lLjdgpr/7k6bQPwBWp7B3eAIjkCBM6f3jay5Ij2Lvyxc3RODPgWbZ990brQvDSVDQ6e3b9/b7s8cB/m375Slz0lOCr6CbriqiCqamVdAXUosPCtK/gyg/6Y5wKvLijiAfmmxhsHIyRwolQPI1IXeI7xa0KcYftoNb3XxF7F4OPVlvHDiCBI4S9qOIN+RKxIGFsrCzc4JqyA/UM5Zlhkh/b++gF/DCT9IRfWmKtYpQIwX/mg0sG1VKAfx3++VVD0i/xuQxp2+Unzu4uYkLk1XG5l4yfPJZxCkiYMx2sIuE12UFwMf/tsj3zRijeFHvLmj+q+W/d6mKpZcGu/3+DKNtmrevO2MM8WNu7XbOnqFAJAsGUiD41zfTgME0Fv6jB0zQT9j1Jnfl//CgBp5wypE+v2Dljr0esCaAlWd1FFffbN/dTNPgm6SdeVyj+U5HQpMNS+KmyUecsRdEOifcLd7s3YvDwZ8TmI+rqKSDg6UKaD7RTdgwaVf1fKCb4V69op9wbl2t0cvOCfPEciZbxi/lhvg16GBW4/qvJE+GnaRuTQj3LbVC5aiVCaQisIn6mgZ4eVHP0Zj/VQv7WayumWGU/w845f9ShNDozlDByeFJQu6kgY48MEMoaXBp5qkgRSkSJJtRhRZiCMbUoKFtrlwQaJTDansGVqUNUqCSBFcrYPV/6GIIZox5l+0y7EKHh37IJiiiVXo3iJKObSRGC9dfLSmQ6TUs+y/vcFwh6ndp0tlWt0FCSu3V192aYh1o+4bZ4o9LRyb+mkN6OM63A5qa9VxVjmHDtR6PkSSsLIoxVG8f43fQFtICcS1TkwWDdRMNyJe/1bBcZsKepnGmIFQrBGExVOQXGamihEQCZX+NVKDXQCsP3StJzFfeLIkm4SJZTC6LrnGGG55A2pARSOG0WApa/uUY1eJTmH9hYM+epnD2I5kmWGo1VOJ6kGOi4J/84QnhJ4fybnei/24n+hx1W1D92AvJCPtG0Ln/Q8WL/inZx/u/QD3h/0K35n2hTF6BPtCMG/7+i1Tz+O3QT+R/0bcEftBr4H7TJ39FjrdEWC4WtwIHvOTGlxBpJ8SoKJ+LRLVhUKjetn3Dcv1HA8AdOXfQHzgbzB17/d7g0RhDz6YwOqqDP1OBcXIXkcBmFE4nYOlwSFZsTyTLRsCyX0+vqPAdj3A0lMWDILMG7CLPjfjlfOm4FdG9qrKFSsBCbpV0wd2XI6sFSZ+IxXWzP2S5jcqONr81j+QNg/9Zs77alULvO397CAdgLad1Ke3yufxTZzgzBMRK2czvHzeT/DfYDf/ORbdwqWKYJR6OVjXkcsghyXj9BVRO48edJBFKovpqRMnroPcXV3fSetN/XRplyLdNNR1ZjDz6kdeP6wb1GS87teNEGYAIW1UnRArP1OaY5fzUBKYFL6gMacA914HwbMRhlgSR6k5Y01qUJ9FA8CFIZZiILkQpKVmNVRIr6Vtz5frhTQ1TFFuMvykT5+sKZYuVuCWnGxeStttiHQtG/FmwwFGRdOmV1rc3HIx3zVbnn9TOpq71FORcRM6nz48lOazIqCrq5vj3WxIuOE6CV+FNZjDiyVNZk4uVywemmmQvOSGoqWQ1VOMqb1uuUFWCpzlAUOA7jlsSLdodi16opShAvW4xcpLurGJlZD2r90Oq4mw9WnRFns+xXgl/4fUNlUeAAP/2bkcqiDpQ8Pon5yQ+FaPS4K8x3mf/7sSqdlT6pTGpVEVrTrW/kl4BksrTYhvgFVBxGarxEtwcglRk4aou/hSbxGdrhjPI/lhNVMWpatS5G2r9bNh8uh68siLOZCy/GrnDYV5kRP8TOczBRnjE+tCWltri6/eWuLpGfL4D6v0I7DHLCbGOeDixIaihaitU4zhnLTymJN/1Ui50hdciRxAYn5iyqjpfEHpZVct5kdl7RWO2Phh+lKI4Jl/nkkTTqnzzG/5NHrmhmDYkAArGVXknszMQifrQFcpmkqgv8U1X87JX9BWZ4ZiK37IbgpVgDoxJkuaICyTQD0ihByeQxPyoL1MijuR3Ll4K4szvjR5n8MoaVygpHl7IUuxtLJcCf0ML+hEb6IEGc+T14IklBIiFzIuvzqcmZXX4a+RKtkgNTVkA5O9/Q45xFtmlwNSOMEq86k2YoZFoYwIhwUvTva7WYeEZ9EnbGbPIo015m8OdkxAlXkFVl2bnG1pYgXi9Bk2NRPxwpL4ibYBE+pKey4FsWOzzfn719P5qI5Wue/Xgielrbmpuwo7bO3hDFceXTsChCOLiX8RyXxlX2XcAHVY5ZwAdXzFrAB3sgSz+gj/qWzvOSZ5xDpK4xBB7xlqyzuft+dT96f3Nw8BW4uw5/OWFGreyuDkbADi+jwPhL3ENC4h3WDKd4q2YCRXWoyMm6/j0rrX6YI5S6lH88FRUpBc6V6dbbZ5nqVomtrRU3W2Y4PtJcBicTbPbQHXCzA8fnqS1b2hGcnV6tFhBV0M2iq+8bCQINH/sDqBuYEY5if5EgNN5b0afyx4nCR92O+cHeqHFQ3eYxH6U7LGqmPr9DzFzVkB/LmfyhqCDwu5UBFirUZCkRRc7WOeuCJ0+y41HkKnxKMNQQRSwiMBYk4lfCXt8c33d+MHDxt8aBUrG3UerSwugLSdGZpz3aTLBSanlf1HbwcFn5nN9m/qp7XdjOVCbEuUA+7HlwMIFRD9AwFK6t9671GuKakf2H7P+t4fCC1Hqx+DhmXndhcoETGFRCEVu/PBRU84tcUxUBP1O88nlxiq7sC5IAEINifX8QmA0bCxrV7xLbxDbC1thfCfP5BcGMcflDfeFA/UFeQJFQCe8XsYvx7+mo403TismJLda4dP7dbuk3lXbCuQoHAWxo8KddSmN5QisncXn+7tVxjPRuVtW6xF8R3aUWoZ/0uoeg0n/Zw1PsoiekyOFDFAD+O78zDzAFrZ1kn7ZfVT5tj+kCK94dUHMK+B00JVXDwDCX17p9krGppwUBbuoX8RfLFfvu9MmurewFaX+4yNHV8v1BOP4L6D+Q8eA0LD7SI2ibC/TXo8vhdwDsqO53pc2nNirI59i10nsm27myu0Iy+Rz74Wzf3ejYbZ8n/hBczoqsjjPBGv6Kss94oNlya6kzOJBoSNKhooi8YDBtVjGCwWOr7Hoso6KKxxBZgZbB95OKOle7pWcpz6plO2OJti0btzr2tW90x+oSHzPZXnGlaPzRyvYswZamXY7uDvORxjCDwl/JUo1NUOn8mpPx8loyNooQZ7nRItDGIFsGxFq9h3xK9EGyphJ1AdlO6QAZmA3EJnGLBbQZjc+Q6XJIqivK9+/aasogNp/aW5/aXf9V+0Eu2henzYfxz4jxYhFSSxCPFXSzAXT2J5wa+JeiulSXTKoU2BpUs4L/ItyW6tI5J4jNlc3in5E4nXTRI3RnXLKtYFLPT3g12JqkAQ0wmwktkh3xlii0LpOvloHtDVRzt0cCyp8RZ2vsXK0lgjRfnJ9X0M1x7ch21EGNYGuJv4X/uEvqOg5kE5S5R7wFcUki97cAflxsdNhsgK1NBcp1VT7+lwjmwda8HW1cLKglLoUCHuK6hX7XIaWJ5zK++YEGSbJCyZYpKlp856ao0J95dociOaFGhIKBwG3TTL59O4C+/CtNExJstJ99eP321v5ddhZl1eo2cHJwlkBw4f9943ByAyPIBd/GaqF7v9HBMSeLfzgFPgXKMVJiSDnDH75y2B7ftldujHgtyPaIKL0ojb9qwc5O7tJNN/Ha8i7tyjuP8uV3Bl/vXuPMBEmZUV5Xl84c7ArC1kFUDmX8qYZG0aQJ/bFwvlTKNXp8oXn0wIDN4C0nWsMkQ3uQp1jE6er1wsp5ngTZ6XeZiLMSJZaI8xnn3OOlL5dokrU6kwuMlVOKTeiSwPo9QcaUTfm1tau17bLg5w5WtkSb04aACXuhItdfmi+Wb7V4TmWmtTorCrR4q9vnexRNp8uaRVqyIwuClEDSzbHf+p2ZWZ25mhZLc7yiz2n2uHStss+dVy2Yh427cIurWAmbjGTnvzVJpX8pNOJq20Xsd+bRGmVWFl/WpAD3hcaIL1z6y/OHP57Xfno+WD6s9ZnnKn1hUxLm5iHm9hiKLdFXC70oZ2rxNhL9DVuzKk0e1oqsaBJB9mo1kTTCz82tzp5uH7DRGhYnV1cHXdqWeD8T193uiWv0Z04OPU/5K/TK0c/QN39J/xX6at1/E3pr47pMUy25HMTY3ZeFFyOuE2WDWp4PPM6ncge/nyqoI9+hyHWGCGDIMjduwoTrAjM5NmdWxRXM04rZf/H1QkeBJHW+/b52ZP4xTkE+vRHDZdjPb/SuuuEJ1IomxujWE92mSUYapynkzpu4AlyI64Xh8P+1d93vcG5tVzd674SIIHqbERy9RO8jeu89gkgijk6i16jR2wwJZrRRogTRok8YXaJ30RIk3zjnvOfkePP+B9/tB1x71nrWfT/bs+/tup61SdV8PIwecdS/pMSWhOzyez3UPzUy0FKdB55wIEIzMIjdS9pWtKBy9i6Wt7roS3ba9Ny6R8goXm9wvNabUi9nndxE5QHUHXmmSOfYA4sG7BTuBvQo0+ObrGkJxYduA1ZSacgAa2K1uXY7ZntM8b79TuQ2LF65KOOyxdG4x+fYnZL87ZKEZtBjOzOL8zp5NS3cnZum4X6+jzWN8YcbAVtsWy0ESL1EFpvc2mVMm1xO29xvZklizFupw2ydkvOSdTXaQVIZgK32mQ5ypAsmUs8ZWMA6s7lphanuyuP8cpy9Ei2pUxItyfhKUvWVpK0rSVtilFsT81kVG8ZlvngUyANkRZzf4Ywd+mPzN2KqtYmq4xW2WC6njO7LsurqG8ifV9PasFjcoEJ2iAp0Sr6X9Ar1o7a/Z1xWGSBnXJa0Vz05D3m9ZFwW+A8VrgqaSiDxTypekfbzaeCxOPZtIA3cwH08iEXrf3EZl50gc7V6JUtEo6vjG7Q9KlMWZICdkvsnFeF+psF+1Kaim731iG/SoclZhKiBesDW+mFGNemrBxRIeeRjqsoiNqlOSYdudF213kkeN4h43kXeyZMP92vx47HJfWuRWwT/kXVaID9T99g67rZiNidmRF/7ed+3fd8Yeauqj5pV8kDhHWG5BE34vZBU9Tr8vrNyn6jn5pYSXZ/lOS2cHMZQIhuKLYIaFKiiN1r5t2ioVXG37vvEuDwnFlbW33ZcISaiUsGJvJmkIFVoI7wfRA10j8+bnBDWs/6Nf9Ntsxf8pa1Umrj6y7FrdRLImypTjOOdgkoIjsn7u8FCRAx41M5L2C6RppchQmFVJo+ABadbKo0RzyOwXusOGH2pI0pLq6v8kBgjCYbl83pUvyRF7OM9lSZGMLm1vdLw/9gjWJiNZ4FeKyP90WtlVSV6rQwVCGQbyofKkjuOAN2+3Dr6fQ6AEKrRHaEp80e6tqm/LPgCr/yQ/yfhzlEqmlBbB024deAIi8bNFNkKePPWZSQC3TYh0G2TPB26bVrvxB8jRnzDLsCzKAqiUvjxMS6fiIoI16Ig3aN+Zse1qNLBD89ioiIL18JlIMl3JKKBecNmxMI0ZtI0WaM3W1srTMCeWldKsf1uoopjsI2cPJ1MouPIJCtAdD3LI5D221tPHIsMY+51wjp3j6qo5I/VSqZGTiPKuRCTkbTdZzr4FmfzpzIy4xbueDPGiXdg8TVRf+h7UkpEj3+oSSdEvJ0pEMA254RA/aGQJQCtMMkJwXaVs+u0KyJ39x2ZMJpzS9Ri8ErkBueg70iu3gj4tB7pI52hMSgZVG7N6HuzKBgkjdUT4v4IixibELEfIhb2+rtchvxHd97okR9+t9Y73ynI36902BriGRWqUR3YJEaUG0Tgzl3QlS/uafFDnteEjOkVhZn33voihZsSWYwRwXHzoC2Rm62RbgxH2UcZkmXTbNw0Yuu7IhUTuw1Y9X8kunfvQVxXiX8gtyrHvlFlhn81wM71u7TnO90Y4qGnknBLQOG2yHJAe1Sjg6Z3xOfOJt+TPIhGQnz5uEZUc/WaUPf87P5wKP6TM3x9XF6V6YvXb4Wxyanu7KzHIjo6HvlSqrpiqmYZxyyZIgbiiHCdleQ9mF/0KXwDURtEAF11S6AzmQXRApzWEJtHziyJLm5xfAca6jhZADEKZ9dYcoBYwXuH5nmt24LRYiSgybcPZ5nvZpZiWSOsce/kbXlE6eAmpA+erlIFumIGoq7+vfkOC3sUOwufmgr3mJCbk8iFRFeTBkkgh94I9fxgR2TsV0Q++w5s53Krmdm1TaetSl40IBJRCtWxoMbq4GJuiq3f2jGdq/oIyNvcw7QlzdM0wBSyqA5NHT6u7yq7EOClCAx9ZHLzbALrDG7+5qNBH+eU3FEwy0a3kjzTSEyfwkk2nQFuB5cDAtHcF7Ne2Fm2KVU/gZlXKhBM+RGlBOU4m4gkTlo0wI+Q2ydQFphf/VzVaMnhuwXsLPvWLJHZxTVuiKFJV8m28rB1Z//OWysjnLzFy1qATB2GDD8eTM02yqIfC1t4o4sLt5D8rFiOFiMPvaVcTBGx0DozVKWjkpkceQolPRNDmaQ3IjlNb3eU0dpH4+QhuHFCdbh9prUakNo5LERzT+L2K763Ba7u94nRYIzj8qp+NfeSCuzfC6Tv4ApYUwzV4Q/WSTl2hT0X7U3FxSvE337Gw0mIOjfUpK09rcjgAcig93lz54ZKdDQy4AXlGrwz+IypKO/rZZp/F5O6yf1+c25aUGziFnJr9tzwaRQ4vFUnsgIjQ8BYvn/tycySQRDZDECm4+kU5dXdq10dmFayDQlZx9Tyj3ke+mSTNiUCPsJfShiq0+/rx/SZdrQ7XAd41FWFKWMi5cZ+p+eFkdp+BS4De0eZAGHIIlW8eX3zipG5ICXimHHJ4LK1dmgpxYUHo5TofuJY07zRWerbNYK2yecr1CtFJSo2CfB9rOMhVbJw2QlAmFA0owuAVK+I+KGcJNwvN5HQQ1NTX/GHF1eOSaLe3a8SkTwaxLAzEGoyurVlBKsrucQy2i606CFnlJZpsCe9AE7pkOEzA2BhhgLEinjdbfJyzFDsXoLmEIT97MX2jywotrIynvkGsEY/hHuzMFnpUcrrevwbBrddJypIik3quuypUXHqMeRAZmNaA8lEW/cgNtzJvveY9FLBJs9J26NBjrmV6Xas5fqRDiaGdxSjYzJXN12Y3hOEKuyUmTJ0p83zyIIHRkig7Ecxu9jZUFxDQUUaEHjwm6mw21SIdLFZneLq7uSgrt9spWv676R6hsCwPEUlsKgPZo+AtVEhkVBhnAkQugJnRT8YTGjBQ+TsdquacXCSzzjyBScGk7HL6l5peRDH9UgwBZbKacjQJBYkq4AwG+L4rCoLuJ0DGqtC3Xx4CnEsl/Q0AvJOvXAJumAGD73HpW/GcOTNfZM1+eA9E3iomi8G4ghIEn6Yqo/C8wnN1z2vcFV5UNLQrwMW9e2CADCLzttXXTHBQ7oRWmBYm9EQo/9vXBtGO4lNT8Gihizi0KyVPElG8NBiWbz+ZKx5cA1RRto6VqyzndGQq8RMCsTROF0ZLBr+PkNCXtADW0Jet/zSEFjUOGa2UgZhQNiAw4NHn1k/rAtG6+wshZhYQR1JpNIyfaH+72dbIQyQ+Gzw0CFZdjzEcb+rFK1MoH0Kd1DdF1MFrgd19GPe0ASL8g0UQhgU7AeZc7LqQ3LYuCeCIFl65htVhwaTxHPVPArt6zLVTaZyaLU2BpOVb8Cww4QI9LXRGTmLK/YOQRxFHXEiwLB9zQBIVhJRPsQEhWsNHqrc8WSvKCkb54FkAUdry10jyRlE1xzKPNi4AIR16LJT338CvTAEKopXviAXQpC5v6bVnSQGTxkCt9wmB029jJLv0ypPhUAcc6MG56BZBZuN2pPuzW+DLkebvloCTXi+yARY552pF1uwZOPfrN/eNCaPi7v1eEy8nchdBUzZ/ak0QqaynDk0vUGJ7stFP/l4jWFuAxvr/g1CXMb7mo8r2FocP8dIBOpYnh+1yKLy7+Z1xq1khGSkNGJ5FN9FfltWXRY3FHLoXzxhwpdj1A128GSZs7+NoXOUnZYIBz/ZfZraxpj3Vftxpf8NaCBx8cgT7pQSzfqVzG1nkXpHk7YovMQW9SZ3IBbDcClSv8UtqnaakbgY6/ERV17Vi8zW984PMl0Sxw/Lan0k+M+UjPpd6h6arihmhZ8oSvCvwPmHXZyKRu1TPekYmeQOeVbgkwB7CYbhng1t6iasgpz648N1HW0zGtGnSePLIuNsZSvDmp98TPcfu6SOu67B++9xA3aifEz7XcxejWv7w5ai+Y2sMfqgqeNGtxK/ukswQnMM5SQ87n8ezcup53E8YS/rdiiV1x7Ic+5cgS+Oa3KXTX5wSWNiJvqbJL7qwJsnKDhNAFNX2yxOrhRpL2FVQRTug0IOFMlKeIAHt5W1YwY0w3LS98Q/oDksP/0pZAc/4u+E7oVt2IJe0v+RDk1I7JJtrTn0L4p7e+NG7waDWnRi+v/gwOtfgcOC41LHv/UbO2pDZ5lVnWVTx/W1Zw5z/CXyBvFFsFUtS3LqGUO2e2BDcubPKAj/S0puTvr03jjTrW1JjJ4Ne3WmoO/kQR1rBUEKTI6cQA54ZjiHt0lfLI5smsOCrppO8TNJ0fOuUJNzdYrt8ZRmL4ZRLKPPTiXWhRIFzJqidpmuhnWHYe8RFC7n6YwbERSL00MtealPyevUD5guSKZHKT065Tgj9aCiw9yhgnQhN4J1E+X4APsWaqkuB0LdosLY7u6ZLbThfQ6mG3UUpEZ+OHP6R7W0DxOYLDUP6vo7MeO9cS9bZwRmanMyXvGpy6lvOpCdWC1pXgrd1XiJJawb3RNTPJ+wrITZnQlrrdHxWTEaDE1T+WbUqYpXroyC819YRlsOPm2giCWQfUPVY/+toySaEYMBtxaDu0gurX3Q86rr6JEIFcVxegGsSfN2hGe6ry4/O8Jx1sTUimxieMtkKUA3ZshmjQXUSsZjrDQRBGXn9ptv4YOgX1Go3dQy9REVVtR+63nnnkMBSUsEq3z2/hfJgjgsPQB8dEhktJwSXkPbvCPgqD8mHa9/nwBFDt8ghmcO38tOOkuLhCOp4KOnCQyrjsS8LeYC8Aq5dhyPLA4SMWdY9ms6t4NVXj25SiqUe9VkYNqzXgUZcCjo20ik1xJ9kCCRx2IuMwEqPI/X0TOzVLwCg6pakxVlj1AsjY6sxZVjQmfojM4wPw7o1xHXeJiymzkp4VxkzYbjcZXniCMiW6e3RuOlm7Q807MfGTaePXuRLwciQDSAjHytlOIQPI962FxA1ZhSbZAzbNqzgfe4YjmXfNg3Mwwz8OPA/lMTegyvIDUhtgtz4QE91dBjm2CFTZwQTAQDAyM4FBPp7GSyzKP2jcB/zTpY4ISDfoSFXmg69iPMl0a38hgH++jS7j+1xUPXNhTzqrbEU+6gsKozZuarnPEqq469eVXuF1aMUZE+c+4SlyV5+zICtH+jHMejc5hGtEGzudwwmADlj3lSUnwOuUo6YlvjDifR9Et9zbT8IU+3Ib4arrzS1/6kKG56gHy2rZb7bTjDkBO0qRGUgXyCLuR34a7Z2AUSwZctgiWBmFRwl2Sx/ZOSwfqZ3ICiUm42JELQ4ap1j9YT+7PseyFMf5S9yvuvshM+b5WB66ni3xxRER9QKlIf6UGZpOTe3SzOW/ElRa0vu4SC5kJAW9O+IPI1QZcqLueiPVbRwaGzuX2fppFSrD8vOZIfVzevjdyIuvEYs0bb2a5KnLevAF5YQQnXCwHN97G4Myr81iVzvKPJ2hXebPE01Cb2cKz/HJW+2fqFcuImbLHfIH/Fs4Ux5pkPqgRH5jEMU1nVL96JM1WdaZnBwckuI0g908pdiKRvU/mxY/KApLZx4pic2vEacfhQ+7nekaHdBTa57/y9Bz8McrmfPtAD6Tlb8++OZv52lzdalNplIA4+KW5RbKxRA+rge1FH7TbJgWDx4u8z+lSqPLPzRLxuLKU/1zXzAFT2wFxkvrO5R69lcl998hFWUYDGwRWkgdoLDYHNRvO+FooJBOceIBk9Ps7UfTMciG2QmItNPqCd749eroq0NVx4Vep1gAQ3zylsc4Y316onNGQSlCzzNzxIXPZQf3NnHknv5dQ8B7o0dX31BiI6v2uE9OaNFO6sSGfweoiK5PjE3+IE6DbLVao8lCF/FtPwZc7nk5tByA6YxkuUSr3b7LUVXfdMqqVBS1XNK3cPuuRnfOxWYLeWXE/+vhjUnywntVyf+ANvSXWbVa01gL20d81a6lS+HMjN3rc4++66cQASLTBsWbGsKQ3IHnN70QDsmUnCZ0p7Bq4a9rpqdz+UPiwzdWwZmrbOjG3Yn0uj9gK26B9LrX6j9epN6fa/LRBW8sboTzUzZyLdZi8aZleenCZ6PqJSsxJPegM5/rrj35DqO0ClTSM9/31nk3maPfxB+rPNARPZlkgt66eolvwWi5O2qaKC2GVE39aBwrjQNP6b7xCqc50HyOChF08e6OvMZDUtH4p78goCu7Sn9cNa7Kr2WjUWumZzSgo9yC2JW1iVL7G3m5RjyqOpInrGly9l5hsGT891gz3kg4FPb1k686aFTx3bwqafS2H5gt6nv13Fi5g+cHflB1PDiELqQ7sLI1CTgTudCg78laJThdqTQrOr7qdqkWZn/qyEDfyABkfBd0RVJcneZt87RoaIGlbeNMkv+V+tzcuz7K1TAuEzhLQxtJ4DsmvMext8q4YkXR4fPPPsGyP2Tgq8Z0K93U5kWkZMeQgaVliNUC1oliwmas812IpzZe3fYL3h+TSdOqpeHK5Ojw+OV2CvK3Ckt9uR5H/AhFP/AZehF2O6t7wCs0Rzm95mHbajnR5fncwLI8wmKhgjZ3JaZfq8iRa2FJfqSYY94+Yj6Uw1gMOeGrLlkBD0N6uHKa/mSOlSSs7uRO2DD/xlm/l/rwm4Gn9EsjfN/6d+I7QEB5DGEui7SkNn1fTAvoJbBcds+u8B6l/Q13Lz1G2ci3Niys0dqSJtoOGj5UcP0ZLPa7MOVvQsgS52Tj7+IPvVC3LpzXOJ1pgYGLQ4GBjU/3lB7pGDrcFPPho54/JenWzUGIvnyZdRkc36uYzjU09M6rdr+aZPsMkt+SkCSBYhP4J0odAD20d3Eur3BrBvQI/loZ/xK1OcfUQYyjgwu/RUYUN9px8HnNUo1zIpNAtCcpg7kvwOzqpda/FAeYvPzsZvywsLjP/GlXbqp1LxicZ5tcc4jIBVA571VajODJI5t7NFcc9CVFBcX+334nExi8T4mlNu4/TkLtHDVrxPM4Gu8rgUcVD7nQ+mDxOQ+TR1yqQfIo6d/Q7FUlY7NR3FE9NnC7vGErpsES9MeYp1xkBjm2QZZNil4u2hXTYZNdvEjCLn8dRdOd3dSYVrF28ybrCdhEVvjDfF33139GolnG5SaDdQRWxVZJhnFeJWY9j1zkTnN2CGQttWtpD7+CrdfvknBtdXrR4kpvV68oQWmeE96qZ3Zbq+visgWRppS5gZwy15IT9Vxz5rVUoNYt17Azcehc6YX2D96o70m6tLn6FHHhP/9Na2o5enr6GNrbvD1f2YTZvzRAlRB+9x7bAUzjZ8vXkBgbu6ZPrXWFbOYdP33jSyk6rVCpk5eVxq4T7m+Q30m5yVTa92MFv+7VxDl1uhGZ/OYi6/5nz+7N4jsvZKVRl3r7mjZk8y3I7CQGa3ecEz6h7LRz3GA3qH0JQ2Ka5+aanLsyV7lAbOx/6QA/Nj+t3mCuTj4S/U2gSKm8afwxWd2DkUGb/31/QmtmR/vmdczewiWcgw+VDw6DLdzPbzTeokDuoUjRCuWG+P3qfJ9nLU5iVi9zVQR+Qow/d3X2JRlVgP8fEGEOu1deKM0kbr8UVQPpeJ7E4VreWmv3FW310LzDchI/fENM0Y3BSWvN0VlgQmjdt6YMpLGseb06BVihXNVUduYydB9lCki5ePZutuVZTZbpJTD3ucfvGYlm3y5C0VGtsk68z2R2DOCq6qJqwPmdOKkU/zes1VxfqDcUw25GwlaN2DbH4f7AP1DfbEpdcm2ZdiIrUH+v1KpePpgz1X3TlEmIYp6F7gjPglWB8HUcW3izPMnZh5e37z+ChKt0ytpkxc1CjE10lgK5YtWKQv9C4mPsc1R+DTqao+LLJQY1TYV3PiPb721wcxqhhkxUOqjX5adqKh/GUwBgKej+3SrQrt6dFm/dE+r/smvro6tQ1E9+WkL2BQUQQqOis3x01JXbYm8WY8DAAyv44jZhvbFROsj/W1ErH9gYe75n0L+52KEEyaMzakM1yYsi2M4C0YSNGRIhlrRGHHbv7m3IFqLouzbfauFXdvRcBNXoR154Sqk5eRrz/Cmp3T52Vs+AfVll4lY+CoJCs86mWXq/YUJaDzvmbwGwTnB5swf1l+hR9YfPT4A0XaKuV3wJt5kLhYhQfbwsJwHfytfH+Plg4luEMnUuXcQlPOI5BSsmQKIZfZFrnvMzass5CXFyQZE2UVE7UQ09vb3ZXg4yIFKyPyyiqMPJE4Z3V1ZuNMpmzieoRKFLXIluKMKjRtCuM0nVDP9amPq4veZCf7fSU9ZJGGwkEx9/svzSFyvQ/cs9Ajt7Gv7H8AmPZedn84Qgjaefk4/HVAxP+7HKIjMLlD+5MQudwe746sUltZPFWdbdEEwPiBaXEaN9bO0NT4ROJUuvTtoYVPxR7qvPAGhuHsl31Tk8evPZLJ9uXNYBzu6Uk4sXD17Cdqr4xON2jUsAiLlOxI+NTIQsy/5MnybKi7SF8cI0Pukdx8zkmlvpFuH7aurh42K8fbyUfJSfsR2YBlt+m2GR5IBxz5bYAWjkutnp8fJuSXksScqq6t3nA8xK2brIDNT0+cOkQvXQNoI1kAxaYJ2kNrHu3zP+HTPhJYJ9hGeCJpbjGuOR2tLA1/OoNWgLpec5Bb3jEp7MELUNaxie08OSbh2E0o4+gZC5Pbatp6YKcFqZyNGzwInDqTOso95nm3C+OzJQw9Qtm03LMMwxQkNzxEcb1se7O46K7UHhWw5EayWb82zL7u5zO+/lmPqW463gphHQkiY30SIHbxUVXTxa8x5HL2QXhjk8/bNbuVpsbzVgTXb8PGSbW/DcWALVPdLR2n39q8g31PqZYm3reya1Nl0JgO7KZdudicDaCO/TLA/aZq7xTa8v7DDdAPnF9NaEQj++46erYSokfJf57QNt7e/z+ff3JxMRjwnBEib6FblwUgelPjOLRIsuS38t8XuBAbq30wc3WgKlENtY8ttloooYb5zhj4ZU3fjoKRHWy9lL+PfyRvAoOswUpfj6jO3IGZTZLtvIRBEYQUmvgoRvPJN/SbbgffWRYCtkbuRRSwE1AZHcuoNNjzvZcmJXdMBaVuVLGyUR5LOZ5bqrXUHD6Wo5oudBFgaFMSY/F3EiBCpiwfkv+4D4eJfCVK2IHeku4wY86JjF42Q1RIDGDOowAIMPYLIkKnTvHqLrbVHK1gWZMFCtkf76dxvj4hu2FfvPocSoQw9/7eJ1HFoeQKbTQYm7E+N55LsIT65sEOAVw/zFKHLbvvj3ob6X8HM4jsh8AIEkXJpMg8Wd+xTw/Eet12NjdYKcRnSYQtYAbFJJYiMNMuWJYPg0VMrXCPU1YR5dA9cqFjNzBrO69Qzat4Lft3qIh1ZWgh587gPeMbvDWSOtI57RsdSbE3AHkIJivAqfQk7Q8bf8ssFmlNIyKpEtgnpt/L50MK+kqVVNyk7gpN+KRTHF/cC/CualWbOUQix61rADo8uw2k9qicUpNUKb3is/TCqdHELzk+wQv6YMtx7LyKdxOZBq9c2FxgXbvcB4/Y+Zc+HbnFfF2pLPCaRdjF+nP9eLhM90t/W19ZiSed6J+40fP/Bvqv4x8voD9Nbv82Rfp/h9ufIna45kWXEDHe4uXifIezV86sbj9FmcYdXyl1UgTjPTVNuwFjV5eFttVSXt917uc/zvWaWyOULYRhZc75Fa669CwkMtp4QeeXrtgvuns6BKhH21kY6ufW7j19161OcLBNBS1UAgnFFHtVTQ/R891PFUofUJUwvmP3yrd2151rwsqvFfn1LeZcSgVpWEggPaPupfSSFCiu6FHjrUcE9CvKEf50tIRl1Rf2PHGQhscKwou4XQuQihcECLfj8d/n/rY0wsTix/hfJ0Ndj1+cE3Wd4N+H1Pwc9ehL/nRkzXXg9dM7/gkk3i/O8rgO/7dB888hQvEX/H/YNV9nuu7G+08kUP0Pb97rFNe9E/8JDsZfOCleh//bm+znwGbF+JVT2XWCfzt7/Rx4bBj/0+frOst1g6t/Qvnmr+yuruOvu8P8E6qc/+UVcx18fef8T7hY/HoffZ3h+k7vnzi3/NW+7zr+3x3wz3HLFgPjv/rh6/B/9xs/x5H9T/C/uo/r6H8/j38OSuef7+Cvns66Grh4V58kQ39NoMXTu1799n+wECJ6"
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
EMBEDDED_ANNEXURE_K_B64 = "eNrte3VUXMuy9+AWgrsO7i4T3N0tQJDgENzd3d3dnQSH4O6uQYIHCO4OX845792ck5u71vv/u79Za0/33lO1e/au6qrqqlaQhoBEB0ADYAEAABAgqI46EQIGANxBAQBoAFjwd8LWVg5GVg66Kq42RvbaDC6WFsR5kOBUOQBwwH/x/zUqEkasupmQRW8S75FiGzGYHRRVMXxFiUYw4e7BKAhRw3sbeHoOTriZkVQxyMPex/dW72ZxG7gHbVlqk5vDmmaM9RIfhUNaHNdVi1K0mxC1mcQSU5er7tAZ6sNQ9ec/tGxY7IKhRTWkhMYNdkiEusLVd8+6fuKptclV6+1RkPh4XkgjwHf9IM/inNSpFI2LyIzb3iVUSSP1tfHR/cTQqkJC5+YAV4Ui4SZxCKNasbnhHqlgFxMh3t/prsx4H90TUFgNmfNZypNJ5pUy8WavwmDtKADDyZ+wRuoNC27PtIvH+uwLkI31xcQx5GHzCFZMmEXRl091rHnEkpyTgW53nH5mhERbGk06YX54Mf8E4bEJ28gD0cqiYIgRvvGWmM+AeOvCDUGxCCqkrhxgrX1NtSTGoArgJedr6xtDfJy3Osv1X2LD3N7Yuv9MFMBge0cj73cxuEoicPOd6KMx2uoFHaTvqwyG3WZhLxK78ma3ltx9I1yHt07tQnZXc7dtozhNEZHulxAAwMsLBEDhHwqMptNuBP5Dgdd+6CfCDwXWtTOysGdk+OP4X8X9L/5U3DhZaQhm5KATvxfgaBv3Mg3JtAWj6J6s4RMYc+rTe87UmILPTzdigg65DYRuGR+ds1xTr4xr8l3ykd2dUJNE4fwyjGda1Lqln5qaYFDi7cQiDyW6fUksyg7Kr1gjaXZyF6/4wio28wIYyF98sB3F1fr8eJu1zEDqiXbYmpdSSGXkWok7BEFruApjbH3jErDc9hB5fLetd60cnLIfzUbcrIcKSWljikTGYzRuY74Ifb9gz1PLLSTr+uQ1vB8y/EBiAEFIqfmUvXdiPqWkJEYemAaR+e3uEARV11ogao/mvZIq5rHL1MqCVKxuxTiUd8NczWsw/naCE3X8a8094Hcqo/MVL+35h1YUQAMASD/OuFgwOlvbmetbW5v/Ye8KVLXll5nQPbPrXmI0WotCMwsSQtDrZcuBkXCI0NGDslr6G4upR5MkdS/tBwX0TNedxB+eHLe27E94F97wfZntbjFZsiFnDYABnqfQqQztWAZ56TA8opxHFNYUCEkm5cIYTw7dXhRv1U87Sxce8wBiKIQT62VxIYtK1GXHBNhl/IwKHWTpZa4s3+OOmZrR9FvQg6MG9eQUaeURAuesuxTKMRjVq8IajKevidPgPvi2l7oYOJk6nmE2pX1vaThDUOUEVG/7LQe/Rnd1jM36KlW4r+SzjJfhgofYAXy0hUX8kglZ9FZ8rnVyuCSRAOAwGe+1w/8aIkI34P0FF7S1cBjE2HHzyEZWImF7/JxNNLsHeM9G8BSTRPmMFHyXO6Vmza7Y5K1LH1TL98KDqaI7OfNkal8ewcjM5JUaFuFgJlXz67oYYvnGsbod8ArFlBhh9IgbO74FnSG1Ooxr1IBW/+BhYmag8A0FrAEJ3UycJofytN/rRi9PFXuC6wqlHJsTmIc5u7GcTwK667R1dRL70xnK71wrMT0hERO03lCp6kH7rdX6GnJIDKJoEMbO1RToRZCibo+uaMxBBs1Hvu9i3ZheUD0nmi4MDBcHgO1L3mx1+EOsOH6O322eI4yQv7COyUNwQ4LrgdNiXhd8O1UcgxZhLcYsd06VGmFTR0zq14PT5ZK2RSaHyjOSjxP6ehfkuyIkWhSYV1nZIGXTiRP2YsSVn54JFWw2MkKJGn1hFnGBk/pAWpPycEaVdpvP0+3FEinWA9Vfq8HydLTi6PwVw8T8ogCH84pjQSvpQyaQCE3hvmdk/zOrzsW0Gg6JuN6rj5fmujgdlT7MLlvXgpYgYxy3V43RSJ7HaZO6nTPs3O8GqFzESlu0fTUX3GvKQxCct7od5yqbAT3qxcevPDU3GQJbLIk9GTT2aeAEssYaFDjxgAddzd+mv8qsczjXqypMZa7HCLMdyz/YMb3p3O/6AhYa5W8CIo5NNuk8Bv8w7qh9vGUo4SbXgjR6i/RbVfHlS0i++NH6/ENdcAGwYD9U5S8D83eF+V9jA/bD2ID9fz/ZKslK9zIhtyE8BzM5pzthoHxxYCQ5VD1Agh2qo16mALJ6Xa7h0XzWVp+XZoSwDnI9GRtwumF/vQmBdtVFWxKGCumGZnHCDWU9djEGAat/VEnoJw6Fhnr4/aD8COF1mj1aM1/cIF0WOujbLRU5TnB+EYnIZfN0ENuHcSZHlg/2kwTDshpfR+q2fe4DlA8r/eoaG4nTqI13XpqFl6+hs/1ciaYucOAkeRBdAVN7s0Da2y7Ra9jz1kdjWddx8678+1c7+OXShGXsvB5BDsTBg+JE3tRx9jt5qS2qJxsOQg/D1B3L53QjlyfG20MDc2u/FRmGugrlrR9NVAoAAPvn7GpvamTkYM/45xfzH1JTkfbOPr8JdWBNfu7uNfoJGp6HpFDfl7RerKoW5fTMVRM/W5k+Dj6BuNQvXGtPdQWYqB/mt+KHVShA/I2amx+TREhXH88tO2p1LsoYhntGnRiCGhKJauy08bOMrxf4k7KrLMqtpnu4VhODPYy9J8wfkpe9Xw6+VVZkHVhmcWK6gGZZ3ZfZxK2ZSmot2+UPws9TiNbTZTWOZqCfK7Kgr6qHZ9gZ69YQVzZsuBaL+7owCVu/4xPtGCsb3rWO6lzf5aLIlXkkc0w+u73DN9+3zM7csspWm2AckNdo5vcLH9B/y7bWrLy8ZKJz6kBVfmkLc/M0hlOcelXpR4T1HG3InhGfdi2qlujYbMbQeSFQl/r1+WrdLca4ktLtxaTZtt+fxpteHYbm7fFD2raMpm6VHYdbca3ZQcSQqitjn5vxC2+vd/b690+6wlfSXje6D9dBA0PJ36+/pfDtILleXN0W1nCIXoj/iOh6y3ib2nDq1A7Awf1C36veGRt7O7X2OqbCNd7WDq59OKqdK4hGumOYSMNf2l6cp8UetVCbI4Dcv8CiMHs2fsc/bPcSU0gHWUWxlN1R0/R17FAam6lyDAIPyfa+zcNMPTAPFo8q2uhuVAKQiPil7axMw2Cmglrn4sEDpn2TOqL+5fTbgDKjU9H31xp6ro33bXR3p+B683cUnl/YO5KtvXR4LPGaqE0ZouWsnzyzFnBeYZU8dmcyfdRohniz3ljPuNgUbCPfcXbwrtqqgdXdbUNHNmMIfbRyoNUqCcre7N0X6dWEh44AlBO/eTiRV87gbx1Y5tsyqOA2uqS7MaQ/x4cP7dhsf/ouj45vINg07EdauUEp5aerQ66lmmjgbOd6VcuVG+8i2aQvaWKnk5msjlplORTfCGPP+BG+qgYg4bBZVVUN0NrRRmSsoog/u/x+Xhf3urcaFuw2Z5hrC4jki+bIZXnaCaMx/ab8AR+H5dD/e6e929dY5h6+wbzjIXeCkJzdiy4ey+nl9DKP2vfzgU4LAwmVelEDK9VTrH58lhjmwQmW7Msk7u20fe/3cHRIsVdx7O9TXTO/MDhAYS/46FqvboytYwTZBnE4dxuzbQvHJX6mNJQb9IshB9IKeQiBCwkJkQjpC6GS6QtBy77Ty31b+qkn/IBP76GYhprigq12ADldSmYBGc2yt6C3j+xPykJiJonlsowRK1Vas25jJbX4aJ1+S3HogIAAEb/OPTV22Rb2vm7D3rcsOqp+hr1YMNqWISu1c3z8EalvFQg/y1DzsySI47XrnIl9RWOhNBQcVEL27AVF5baymSWKMyZLuf4xJMHFguICzk1qTA7SuWI41Ha/1w77ek/A6WjGB5OYGd9xHt+VySHLnirE+qkAGB+F9Jg3QAShJIn1PLYUbLtGjsqvAZB/jCegN4DM7xOcer24bApBGxC39OvLbSspJYoafN3cY/pF9ae9uhzaQd56pm5jOfqIjZoczC0remm9XBhfe6PLV+zM9+PFmP1smAoIBvdw7rj6AQYisloGk/2zvU3pIuogjwpHtzRe1HTzC/SgShfiKcxvMz1hzXrajD23r7sSNV1jtRoNGQZFKscRBGcH38HnGO8mwnmgQVLAyiKn+uxCnVJNNdChYJ5u9ykRe5L4qR6QbutbcdnBMU3Nq5FJZfz4R4JszZCT2yy+1n6dkLlwJGi9FBKQe7HhbSu3Pnt//BQ4mYcQiqnttzdFHkKRDEjTojiVRU4hZhM09FuEDAXw2nUNP6LYMXMSNCOb2sGxKDbyiefNasDmDZH7UCT/ayjEHIIkAtFah2ebpJebXlx+knnUQRLhxW9EEUTSTCXY5mDP4OBg4eBCYUPhJGMuQxeSE9sMvpnsm3EukU31oYkRnMFqvhfU5tErCDdBJZKmTveJiYkJiUmJgZTMdvFMXXDv7ckQrUynUzax+74Jx6ZYfoBfeTmEGMuQRVX8J8X+ZioHsPone4k2cR4ZbrazmL1NuFkUpNjrlS7CMEFHbR5BhzOZTQNqbZ4ZCjIOFCAzEC6O7+xeKfcmJPjOl4IQVZvHl2sTON3X0yu0/ykARvbZruiqnaSXkoyYjIyMhIyCTHC6WEdX/4LwyrBshu+kprHrWxOGhK7hVBmrlVEDiTb9lAQS753PmnFoL8VPEsqvMDvwp19XtZhNG4gj62C4RAjOJLp9xD67kBkRc4ABgUAUoAVELt7OH6Pho7K3kfjuIzJBQDYdBxOGGXIpGb2NZJq/ffmPZ3gj6Brg/8aXY3+jtunab5HMnq7QrhYKKC2kFnD7U3Iler2Ssqq4iBIUbfhZ31lYVmPVOEvUBST+JewBFf94MQobK7XK4Rdrl3bZBCixe9azGm1NN8NrvOOhY/+u7LwkadT98tfENtkQxqpXftMeGMd5sW5EntGwmAzNPOPkStajWOj01JqSBSsPIYui7CMuudgSpOyKQaild+xKVYbSBppMsSGNuRfXcBwrZt0ILfkWJ2cvIcIrGwpIbs4rzDwvZXPXhCgy9FKqisjQkSIIZoWTcXkYJLDo9bn+ZkHt+f4SVwZUCGVeEWRKUAahKoJObL6MY61XQsK2CYIWMscX30OJ4EsqKzz08G5OR13cZ/Rt+4E47qeb6GkWNiGg6FW0J2i/PoPIojWBtREItCBOpzi/CQB0tDbV5BY1oesnRSHz21fsn6IOH/r9LZsI+a7fL+edMKxT3B+KE1rQXC1G8yyC53m8uX6PB34rsI84IFu2OfnotA0TVmGpgcE/LBx/fLzxKKIzNEPiOoduQyMSvB5MQCQ5/b0UNDvPsPbpArNBNXQL26K8lMfuXJ3DJsjPe6XbtwR/nvbWSTlJAewhJGt9dGRkiqtooNWNyqsTDL1BPSyfCm1w61YR+w4cf1hYOjf0kXh29GtXKP0n3m8KHQNRqypdac6bn6XW1i2S87QGBd3TDq69f+sYdQcvn3v+CDkX8P5Yav3TMXIwNbI0+uv4p090qCFtjceJ3I615018nbWnJ9kQMehLLm2bDmK2IBeNmg1bMurVZpAYjrIStaNTyDHVvFTKaSQMawBDzUm+XGiSnSoSeBYK3pbS1HSwoIS+r3nd0PsEOsEZt/mo7Z3C5CAQUMBjCZVYvgOI/WFS0sQq1606nI3oBZQ2HEJw5CVpU1BxtNZcXtrMXbT7yICw18V2BI1+6qUWbSpxQDoD10QTskATUgbUivXiEClhXyRfq26pObtmRSkmWYiETcjBrziRphIKazEBxeV78GVJtDvvDfpLRnzelRXowKYuLliEtDzctm805U3Z0XDL5ECMaoKdEp338LmWgPd+UjnNwUgYo6Sy1KLduPJAU6Y8pD88SJVvZ1xjwn+LKI+C+E2Gtk3T3tFcbg2Qotnv2LIq18hSbAIAvRI0p5WONa5sVDObMMUUuD5gOZJr1qlqvgV/SiRPTT9yavk6alQtgjjoVUm/sTAcuWQ814UX07eQfFp/bVw8Y3MKbXgWSe5hXf+lpNkFIgxHdkDa9GhSh7GYLzgCtHzdvQBoVxvfQEAX6eGZZgtHkZ/T2GToQmraA1O31lefQi81ExMDBUicrBsSU82PnnJ1BVI6USdIS6cxxuXmpH6AIjLGrjEWPLkKP9HtcAnS6XBxu9y7m8rPjiLk53O7zC/csVjpN1NZuf2+PPH8eG9wOVBrrcN71x2dLPr4kd/jfGQi+8E3RhYhVMfr7rJxV17X62r/ZI33rJMTihB9wzMNOdQqXrhuAG1yDLs/MeNV753TpYFMJgEIeolwWhuDu3NZedSjicM20lleAHgGZi1DCKo5HtDPKwm+GAPZ+nBn1HJg6UCszLbCYx7LTw9OaMmENfFGIHm0coMctZYzDRVkTQT3F9vR0MRHtYOUA0njUAMvWfkoevNXPaUIWo4nFWyGxOFXi0siSxfIPLlswM2qzdIp1pNL5U1s6g3SM91LxlLlD+jt1kmAQKYL2DXKUeYtRB5F0Fns4VYABqIXQvki5fazSoCo1luaUwcRbPFVYOx0424UwBNBriM018l1tb2nKDQcPSzKk81abzx8kJ7vuBc01Moyw8agzx9jAKsY4dR83uXclA6857wFK6b8/EBIHtUizwjpbLYdUgeL/GRIwtPeKtiUHq3M0d2rgkN1IlcoZ3pfVDNGNXs0UsPLPcCxgXxBPnFdre+Jeqyrntd4SrZKHoMxIj5rFNheMdnl3S6UnTCCpZuMVYhddpl3wqHtYldE4SGCs/xWWcxKciZYGhQRxKvO1uOXSrgLoYnZs9D65GwT22mMhKrlKaLaSrT1yu02DEeaNDU+imy4KeDNNF06WAp+vmW/Orxla6rQgRtftaqIfn1PsJNrISh4fCillI3F/iRiKozCRs2Ytvykv0K/kGvKouoDBmkadUzUB/vEruyRsq7NlbTJHeK9yM5hROy9+kOmh/diQUvHTmLZzaK17QpYQ9hIkseklTc10TT9krXqqYGxA+JuRTx94lBWsrLXBE2ZfZKiTvVoi2OtNww6aYjTCIz7+HxJUXT0FJTSU3q0pvPlV4BAkDkUKXb6hnmc70KhFR1z+LoXjUp30FPpkEW348ZYET8drQ2t/lF7wSdlGBbgHGZSNSLiBRUSWFIacG1Qt9Jfg1EKDG/fGtg3jAvZs0YBtTjlSI8TKVESNs5QkcKNDDEoTZ1NRXK9tQHE/RS62LPuRVCWMxEj1fTEhCH44mPObWl5MUg/7jqNwTjyarzKIKkPxNjIJWqSarFBEb3LkqKrpklKuPctcZXK2UtVZlg3hqaP6Tm7NG8CrSlKHecDUoBNOh3T9N0SfJMCczBNQh/bIVyTCGs9RB5UldcGMyTGiln6ix/+3gvHLh1i/DyTahkvLlBwg0epmtgQazgutF05RPW8diTO8A0QTRHK1si4jzAGovv+MzTH5RfCHGyhVZX41lfYK2e39t67OGwhDvEX8kceFizZjbyTq76W1YfPrNck9UJTjPhQsidfTNzKXN+49HR5jefPq1WQR5ZfsRLV42jRAJ9gxvlphMRksA8g6iduchpIXD7qLdNCKnr2cRBKYEm8I4SZBH0RxZNJS7BUPz++7l6majob5159ozs43bjvMGF26Gr/rcbe7mxsdfcuIvPdYrVj9cxVCm7ctUvY92tTjlnsgRRO8z4sHY9TAH4b1EdkChnJh+/rQC1gd4KeroTA+bvdtEo8b7DfWStjiGS4SCgAAJ8LAED8y1rZO7haGNn/aafevbPOB6F76Zq9QGzy7kzHSObFTBXQT0vU5icuJX6bA1JU9wsYZrxOnXu+4f0WN0nOJuATM3Wd29WP783HSUS0XBudK134dhR8072oYrW0JoB3WRfF2nLYu03dv5oOCt+VVsQTKf2cEOZ84O6kYwy8vfjW1Z852ScAK+U86OJIRRHc3V8bfeXu3WEX3oAJw6aCoTn5IkuEmu/tYGFYVRlpn2hvvAG88gKDcK/o4YMV9ydP53CQBqrULV8euXQV71PrtUdQEnpO0BLyF6WSyBSGQLojsrKxqy3U9Dvfk1PunPLbKGGLjGBuRQr4mugEHzteqbTckvv06uAKisLc7lNalM34XdaBGxfiL8/jW0WPYD7wmDId7s9IE6gfBboly+/obqnbSODZ0ypGTRrkPSlskQlsGMrOYsK3MJFC4shJ81L9tYRubOsZaEjMKI+0wFuxEsM6UfJTJ5Ly5tjvw04d2apsg3lGI8Irk4ssyL3BhZSFuGErpCnEPLX5/CQewBX4hfa7sB1ybxSYs30X1lNcMdELtu+MOqAqbzq7+pbw6rln3MYmdWYiR2GkZoFUrRNq6wrKrLMnuQdNVnz5ixvEFGLzeGcjT/SPuvcZ/gupE2qQlZ5IinpC+1hfGLvQ/GfYjdrBilOX6M7UQL20yoSIpEx54QXBSkekdCgJSmNgdmFNKg8uNrsD10m01GZYHwfH6fHNoWW8F98S2tO+ohv63LqIt+PoG5cM+VbNzO0DsglIY46MRvpyXVlT/yb1TPvYtyxdQZjuRCLXysrDBKC7KeO2iaqJrxdYWhqcdGX1SXHuR8jT7jzVgNEye4Vb6MiwO4mHITVnU1NTqqVK1abcTCuvWil9IrUVnqzcpg56NDFaOH2VHDTP/ta5MuAOo7qOIKfheKoMxqae1ykOzicgKJLaExPHx3oc2EJ87ppL6sq9YHdnMgguAP7WcG1aXMKILNqWBMgQAimcpCnDLTzP9GEqUWh3UhCFBhGQJQR1tb+AumS1r2NIJ8rAwkCfVyEVl24R6bSyQCk1nxW+dbzLKuHdL2q2xXqO9j60tRTxhLkgEeg6HppZedc1eQ0VjOLStT9JnTTVt2IKbmsZc2Nk5IRH5YuUyD5iKmB5BR6/l9CmhFF/BLVp/+EdpN1s4TxH97wcysI1dM/Xc1WZsMJvGVt09qrRqm4MhxMejv7dwZVHjSjijdOMSlI279oCoFweQqTTL9HeMjDthTrteEjLe+iM44Zvvpi7fcuwvEY1HFd6CGGxbvQU0eBq7ShTc8JlnI1E4o5E/u6HHlwUh9//oBHCsI140BGvZ/5RywhlcKXOQ/aCAbz0ali1m+WkoPUhHQaQzheqfuKQx9PRj3kuX5NXWAuu63yMykrL1eE2KtxQ2he5DV8np61OPhg4NJ4OW25F+sUVrTRPx5fODlYTT5N6XEz7s8mk0Xk9XOtbDnfQFZBzs+izLSBx45mrBUYTzytXNvyByFr5ISdQ7vbpKgXUp2sbja12SSk8z0rrW1fM5R5w1W8zRxAyvuhF4ago3Ck6IysJ5tx1nTfrTaift7bviwjKi30Dc/4NfXCDJ5isajLl/G21TS7OGwbnALvNTBAYONvbjTOZUJrjWBvbUAYuS7s9qMnKT/5OCR34OmnnZ3fRezXQ6nbT3+XPv1+EBFATZXeEtDCu7+cigUjPyH1GaE5pj7CZdzqyMYkoouWT+LDLw8/eu0RHdvWq87yO2OjZPJtKK/l8fzpxBv711C/cY2pMePF0IWjGp9uRP+kVB9iLRwXrjgwwiGi1DmxR4xDNs8CLhc3x/Wr50nuQ0FaXi4P1KzrhGSLxnr54PRycQ7ExBO+yPLOzOUGwryl2m/pSrSpqpQE19cRkVO2ssOSJAXat829esH43byua7743BgcAHiABAPT/mbdN9eyMDJUd7MysTP6cviNUlq1mBJA9P856Y99BxaSB6KFyqWwxyKFEJSxLZlqt59Hch3RW2rbP+61P3T3gal+jG97ketWYfDvMeGM/n5/uH/dJ+CBgamu7hl8CwM4ZOLkDctXyOvaU6nQtQqCI7Le1E0leMdw1Id3lQVZrsVVy1UWamT2Dx7D3VM4vGu1X7+PxXGJxrBdVa4Kfok3VQYQmUu4LkPNJz3kCNX+WIzY9Q7e5xx8UkzMkzd5pQHI2Tri3SOgfpOYil8p7T7gRRW1hkpyEkfk+ikHuS1Wm+W4gUrtyh1CLMWqRGPVmf+qJXz648FtWueHJkNdHr+MtqniX0tjY007v1CT2c+aujbY7StNYY9isZtqKjZip9ar0e5Jh8q65fd07CeJzNxBdZwlTRE+h5yW6yzXzBsDnxU8xumVA1lA3n2LYooSONrxqUfAtRdSLGmVxXsofPhLxE/B89zsfdIHCoE+RsQer7eNtGobUEKB2mo1+7S0GBkGArkSq1I/Krn29jMYXE34479gjL6hOskLXGN/OOzu4xfMmfpRweW8PlSl90GpxfCdvdhRm+psKwH2DNuSb29y0XRpUsMCEUTQ6EFynCXHr6+IWlOhHt8zJB0KZTULpCKgF3lt9GRRd/sv1607F4RLv9F26pUwqjq8hCFZYFQrBUZEdUBSB/GLrDnNidBVMmNf9VqWMGmqTOwMg+M7lO2jC04K7Fg6rTfzP/WSUBlNGwdTKjjYBGxywa2QmHFwCpuqQT5mVeq4tmnJfQWnrE5kHFbahrWHJslnZnnPGjNVUuo2BJoqTr7mH6QO7X+azpidoHxB/J5dEg/qSchAAQCvWv9IChnZ6zn+I5P82/gqB1S2t05mQn/nMXrrObxKSAXCKknMFRmls9aq6Xj4IRYKsHCxWQtxXHScWPfc1WJUzCQxoHkzBzdz87bVEd63ek2qvKkstpLrQDo+asRWzQlu3OeHaP8uIK0fgirrIpayzdTnXdcFc3JxcL2uNc4zFiqz7cLLeypxKlbp9A8k25GGYXYvIvnmkD6gI8ZN33dqOs+h07MNfcyQtLiUnd63NUyH4wt90If7c+U5UKmoXGue9tlL/a9cG1GknDCUVtWHI9OCHqCxfzEgw2GJRZucOxwPGC/GwMZ5Pq5qHD1mmwo1THIREwbmVQhoIgpnQ8dXBik7J7/nhuwf7ohxnZktIUtHh+iV8uWjvTGKO5NvbSya4Dr8M3cvkscjbBKBQYpBtol5i5YkCS+JxPDw+qFDmvtOguJVOuw8/BSGKUUvkd2pD5vfZIkcrefPwXlPuTxOJIy0thRvjbAt2od6m3Jzd5J2wKCT61KMELo2I7ITkUbEumaCf70NFDfl03fUdZaG0Fn+2Fa8qxnsSud71fx1gMHljzYqQjsu7y/CGkgJTop7JALaEeWbfHTe/XyR8dElIKGvo1SlQIWpziA1eT0P4smcSSSBAcCylY/C1z/en2RpFopKSGExUOrMimcEh3E+aj21h1yOimixFD+wCCMMS36MUKztv39ATi/mMsUelyXzVYkwV+eQPoDzkNv3oaN6GB2Gp4bP1dsmJwrhbFGxS10zenZS1fT36WnUo8PC0x6nYDdy8G7VCM52o0/zqQWERsKe+x/PNeWmHvVXPCAKsVNkUnoIDblxdwwAS7+bZaE0sjqsL34lTTKm60OeMWL3Xwh57abMUyRrDMFc6iqBbpF0RCzdgyNtSv/uClWTojcB7EVbCnfxyhPet7upGnzRbXvlUDRlHt0hbNoI0JJu5SMIuJV/A7h4bPdhch8SnGM6nDN092M7yFvFuYWFd1vN9o5/v76i8hsRntLfzXcI92BD9/zy7bvjZ3WO7z+o2Ps3ycWLWghBfMRvL8owhbnrjWsLsFfkH2AFr7TrXDbKUpfxbhfonnT6QiuW92sZRGcdaZA9BDnlTvJuctIAKt0tq3SIJI1Q7XxpiK5uKtbWjJouz/P2m5zPG73TwDXN97PKP1uQPj5/031JzfyV2fybo/kzrVsSOfehheuVz4t/uP9XmNN2DnS/HaDB5sMTrY0IbqkGCJ/1uoTHbmYqu8NOI1b1gNl/u0bF51iwlZlQoXd1nsMBt5RA2dRdevo9EAtUgeCupAMzxwU8Has7LWBkZs0BCA2NV6l5MsdhH01cPKCYDXEYFS/vKqZmzda8ouF+OodnSi5o5dzghPkMpH/kH1iDhKQzYV6W/GBe3RxNmx72kxG73w6OFcc+N0EwngvX5zSfj6vqpvTViarRceqkbIo97/yZHInEzcE51Qmun0iv8ktQM03uv/bfxTsg3CcWmH1fYfjw5yr+ejc0Pg+lgZKds5ODw5zT1S5+ZQd/M6jBhTGpLAX3gRVKscksoEAzHBshQGJjU/3E1oODKvx/ToJC2cBEWTFYKF4VAK8RcV8Ajhc8PaX04kPYMbGXvpX+YSakgBG9jxeb4rpXozvYNA/ghUCzgtRERnLFCPhGxzFGAIoatglhH7hNTkPXUrdE22x5231xKqplgDwj3s9zqJOI45mh7OT/WrNtIwjy1FL99RYXrC8FWheWsia1/RWxm63LzJR+XJZXs83P3u2G1tcjO5Rbc7a+cdOcrlY6q3L4MzyTR5O3nLI6+tYH4UuoCsdX4fV0kHfjmYRHJpjv76ZcIsxxtJN+UTyK4rZJ4ipUEeV/pd+RXaNYeDejYfO9hUYyE4JdJk0QOLFHHizJ5Vz9dZcm4W73vnOefzyZcu6SmBubkgNxnNXuuW+n1kpwJ2SA74Glg81TLyz3WDRMm6DpdhnPxvNjmgq+R+X08u/zl46Xd4OOy4hyNd1Vx4ktwTeUjXIxwr5d9gPCtWXBp7cntISc7m5mjzahHBGb42lu1VASYivrmAbilw/eEb0suVMSlseOBkAE+yMjgQgKwsGBA4p8dX0HY1t8uzUpI3pgF/Xj57T8cJxQALJihtYGCnbWNPaOBtZ3R/9TA/re64Qc84rSle5iwA0+gXoQHE3XqJMucUD8PqwT5+XfRoJ0InVshNJhJsuw+t5k1zt35Ju9kXNh8jTfmJFIN/drFpZ0ONGbug42oSOQYAXox9E+xi40j5Ito4GFv0oM84rOzDGs2ahqY6h0h5s1tqGIitmpwNhNvECl7hzBUuvu69PyaSizJPW7rpKdh7yiSBFCllJV8saRcBDXfz16KtIrTfCFNr6fn9omDlCrpdBFXHx1MXvCvXMn2lXtjp7ZknW4z037M/G4oc7GFP+M4Mn2M4quWA7uQt5gOhXzt8axEuhMtHGgiYXva77R+jGh+HCcXEbetsvXTnsV2RvSUSAqTr6izMot7Gs/4OKsd+gcYhsCL+FNMBMTvTukTgTMIbDTsaWb5agcRj9f4+srlkTCJlGySpMN5ZNSL5CYQcvELfM5MgTWrHC39OQpNSScNb6guwlse9CPe6UyfQUPlZxTx5PrekK7pdECoPVy/oM2EVIdM+b0GSd/NzaGcAQZbptdemh7Rfie8jArY8Tk/JDP4h3eF/Hfh1bOx+a/s/q0yJ37EavmP+mWoeySaG2YjTlh4f33kbV+6e6YPlBqYyh8MUXIlfDNgnTraC0JAUsqGKIacjFo6WcC+/SrWPh8+UxNajuzECoFyrcqjBn3O9PJ8xgpYyAboLTo1i2jfzOFoot7hFkJes1eiCZ9y9ekVO7Z5SjC5YWsXG6pnQGtgfbPWzea4lafdpcGv6c9ENbJsVlCWcA/yn4kHazbtchrD5uXIvTecak/y9ajH77+aQqFbU3eY7WjbIGIkcweKnZopizibqyDpzOgk+ckrV6u9P+WBH4GWEejKht/EBQq+870WDPNNgqCxa4e9l+91Gg1bQ2b6UroI0PEvNKWN3DGMkCUaHOHePREOlxA8BcWEG4AyM/vv1ccHfLFbuVLq24WDHc/QDjHXNkK2mdfYs54yusXiu07zNXyLXwQ0ZHIj5lniDu0ykemKYF8z2R3MDJtLd4zzdCMokD7s7RLB6Na3IXE+9WsDNrZQs7H6PkNFO6486A0F+zVFPVxQhGY0wuk+gS5qYrNWLJ/f2nmx0PInn7HLfvDjfyK++20c+xbrQfDzD5k9+mtjwt+mZEd7B2vL/wr2T9QlKJovM6EHrsE8IImdlJFrxBmgFbPZwgc6oscd6TEKJm33v5fuV6oJrH7RMXqHx4L9CqkDlgFVfK0Zf/zx+t2CmDhvU38a+DoeJIllgWqkK0Nn8RFvhu+gq16CaeUUFMgMxgrrZG0sG3r5oItzSVSwuwn8sO3kAuRk95F7w6q7Ho6znoBrowfEM9BQhVcqTYuUycFHAXumTpjRSHC4tt3J2Pa4NyBFQzvhcHxxPQuz0vfYeEtdxr1WAu7B+KWMDkc/pqIWoxeWUS/JkFqENWAwruK4tuUKCflqdwBSRMyh637r24bZjaDK51VqbYQwZcjYbhwRKqpA/26pcWPqIjijsMN5MZVKtY2X/Inl5/KLnm4kGN5oxgseL6MH+WFDuh2xByUXT1FJ5A1kPYlBzBn84jfS7TJCsgNK/B86GhfKVx9l9nMowIwMBPmD9rSOrnXzWEOh+LiAkQtICy1JabWaLPDKgSgFZmI574Ev4G9ZJGmkVPopmGOj2G0pPU2b+C1UmqTMx83V5zCC2TY0SsC6gHpkyjzYcWzpvWJUc7aZNTX9wd+1OJ5XzJMJBb17Ihh521z+5YKCgdMD/tNunF/xm705vzL4526Av6Prh3X4296AXwl/rYn+iUvof6uQ/pX4n1WifwcR/J/E/6Fm9Fc+v5YO/oTSq/9cSPgrl1/z7D9hjvHbrPuvDH5NffwEDv4viZBfSX9dffuJKeDv1+J+5fDrOslPmJH951WTX7n8Gun9xBHl/ynu+5Xhr+HRT8BR/5+DpV+Z/tPt/jugaQGAf3PCfyX/p+Pzd/TQ/438f9ygX6n/aWD+jlSmv9/8X+ZGQRoK+o/Lr398Kn/cMoL1j97/A8UZg4g="
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
        'I': ["Annex_I.docx", "Annexure_I.docx", "Annexure I.docx"]
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

# ----------------- Database Setup & Migrations -----------------

conn = sqlite3.connect("ipo_management.db", check_same_thread=False)
cursor = conn.cursor()

# Database Type Detection (PostgreSQL / Supabase vs SQLite)
is_postgres = "psycopg2" in str(type(cursor)).lower() or "postgres" in str(type(conn)).lower()
pk_auto = "SERIAL PRIMARY KEY" if is_postgres else "INTEGER PRIMARY KEY AUTOINCREMENT"
blob_type = "BYTEA" if is_postgres else "BLOB"

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
                except sqlite3.IntegrityError:
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
        st.markdown("##### Pre-Session: Event Planning & Administrative Approval")
        with st.form("plan_lecture_form"):
            c1, c2 = st.columns(2)
            with c1:
                spk_name = st.text_input("Speaker Name*")
                spk_type = st.selectbox("Speaker Category*", [
                    "Industry Professional", "University / Academia", "Entrepreneur", "Self Employed / Freelancer"
                ])
                designation = st.text_input("Designation / Title", value="Manager / Specialist")
                org = st.text_input("Industry / Institute Name*")
                qualification = st.text_input("Speaker Qualification", value="B.Sc Engineering")
            with c2:
                tech = st.selectbox("Target Technology*", TECHNOLOGIES_LIST)
                session = st.text_input("Target Session", value="2023-2026")
                topic = st.text_input("Lecture Topic*")
                date = st.date_input("Proposed Date")
                venue = st.text_input("Venue", value="Main Auditorium, GCT Samanabad")

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
                    cursor.execute("""
                    INSERT INTO guest_lectures (
                        speaker_name, speaker_type, qualification, designation, company_org,
                        topic, target_tech, target_session, lecture_date, venue,
                        refreshment_heads, backdrop_details, bouquet_details,
                        speaker_rating, student_rating, remarks, budget_utilized, status
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        spk_name, spk_type, qualification, designation, org,
                        topic, tech, session, str(date), venue,
                        refreshment, backdrop, bouquet,
                        "Pending", "Pending", "Planned", 0.0, "Planned"
                    ))
                    conn.commit()
                    st.success("Session plan save ho gaya! Neeche button se Principal Approval download karein.")
                else:
                    st.warning("Speaker Name, Organization aur Topic likhna lazmi hai.")

        if spk_name and org and topic:
            app_data = {
                'spk_name': spk_name, 'speaker_type': spk_type, 'designation': designation,
                'org': org, 'qualification': qualification, 'tech': tech, 'session': session, 
                'topic': topic, 'date': str(date), 'venue': venue, 
                'refreshment_heads': refreshment, 'backdrop_details': backdrop, 'bouquet_details': bouquet
            }
            app_file = generate_principal_application(app_data)
            
            p_c1, p_c2 = st.columns(2)
            with p_c1:
                st.download_button(
                    label="Download Principal Approval Application (.docx)",
                    data=app_file,
                    file_name=f"Approval_Letter_{spk_name.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            with p_c2:
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
                st.download_button("🖨️ Direct Printable A4 / PDF (.html)", full_html, f"Print_Approval_{spk_name}.html", "text/html")

    with tab2:
        st.markdown("##### Post-Session: Annexure 'F' Official Feedback & Final Archive")
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
                st.write(f"**Speaker:** {s_data['speaker_name']} | **Date:** {s_data['lecture_date']}")
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
                
                st.markdown("---")
                st.markdown("###### 💰 Final Activity Expense & Submission")
                budget_val = st.number_input("Budget Utilized for this Activity (PKR)*", min_value=0.0, value=float(s_data['budget_utilized'] or 10000.0), step=500.0)

                submit_archive = st.form_submit_button("✅ Final Submit & Save Permanently to Archive")
                if submit_archive:
                    cursor.execute("""
                    UPDATE guest_lectures 
                    SET total_attendees = ?, speaker_rating = ?, student_rating = ?, remarks = ?, budget_utilized = ?, status = 'Completed'
                    WHERE id = ?
                    """, (attendees, spk_rating, std_rating, f"Method: {methodology} | Takeaways: {takeaways}", budget_val, selected_id))
                    conn.commit()
                    st.success(f"Activity final submit ho kar permanent record mein archive ho gayi! (Budget Utilized: PKR {budget_val:,.2f})")

            s_data.update({
                'total_attendees': attendees, 'methodology': methodology,
                'speaker_rating': spk_rating, 'discipline': discipline,
                'engagement': engagement, 'student_rating': std_rating, 'takeaways': takeaways,
                'budget_utilized': budget_val
            })
            annex_docx = generate_annexure_f(s_data)
            st.download_button(
                label="📄 Download Ready-to-Print Annexure 'F' (.docx)",
                data=annex_docx,
                file_name=f"Annexure_F_{s_data['speaker_name'].replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        else:
            st.info("Pehle Tab 1 se koi session plan karein.")

    with tab3:
        st.markdown("##### 📂 Permanent Archived Records, Edit & Delete")
        df_all = pd.read_sql("SELECT id, speaker_name, topic, target_tech, target_session, lecture_date, budget_utilized, status FROM guest_lectures ORDER BY id DESC", conn)
        if not df_all.empty:
            st.dataframe(df_all, use_container_width=True)
            
            st.markdown("---")
            st.markdown("###### ⚙️ Manage Archived Session (Re-download / Edit / Delete)")
            re_id = st.selectbox("Session Select Karein:", df_all['id'].tolist(), format_func=lambda x: f"ID {x} - {df_all.loc[df_all['id']==x, 'speaker_name'].values[0]} ({df_all.loc[df_all['id']==x, 'lecture_date'].values[0]})")
            
            cursor.execute("SELECT * FROM guest_lectures WHERE id = ?", (re_id,))
            re_data = dict(zip([d[0] for d in cursor.description], cursor.fetchone()))
            
            re_docx = generate_annexure_f(re_data)
            st.download_button(
                label=f"📥 Re-Download Annexure 'F' (ID {re_id}) (.docx)",
                data=re_docx,
                file_name=f"Archived_Annexure_F_ID_{re_id}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

            with st.expander(f"✏️ Edit / Correct Entry (ID: {re_id})"):
                with st.form(f"edit_guest_{re_id}"):
                    e_spk = st.text_input("Speaker Name", value=re_data['speaker_name'])
                    e_topic = st.text_input("Lecture Topic", value=re_data['topic'])
                    e_c1, e_c2 = st.columns(2)
                    with e_c1:
                        e_att = st.number_input("Attendees", min_value=0, value=int(re_data['total_attendees'] or 0))
                        e_bud = st.number_input("Budget Utilized (PKR)", min_value=0.0, value=float(re_data['budget_utilized'] or 0.0), step=500.0)
                    with e_c2:
                        e_stat = st.selectbox("Status", ["Completed", "Planned"], index=0 if re_data['status'] == 'Completed' else 1)
                        e_rem = st.text_input("Remarks", value=str(re_data['remarks'] or ''))
                    
                    if st.form_submit_button("💾 Save Corrected Data"):
                        cursor.execute("""
                            UPDATE guest_lectures 
                            SET speaker_name = ?, topic = ?, total_attendees = ?, budget_utilized = ?, status = ?, remarks = ?
                            WHERE id = ?
                        """, (e_spk, e_topic, e_att, e_bud, e_stat, e_rem, re_id))
                        conn.commit()
                        st.success("Record kamyabi se update ho gaya! Refresh karein.")
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
                    cursor.execute("""
                    INSERT INTO industry_visits (
                        industry_name, focal_person, focal_contact, technology, session,
                        tour_date, dep_time, travel_mode, other_mode, students_count,
                        driver_name, conductor_name, aipo_name, assistant_name, approval_copy,
                        rep_name, rep_designation, h_punctuality, h_discipline, h_interest,
                        h_tech_know, h_comm, h_safety, h_recommendation, h_comments,
                        q1_orientation, q2_machinery, q_relevant, q_objectives,
                        r_discipline, r_tech, r_prof, r_satisfaction,
                        candidate_limits_json, budget_utilized, status
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        ind_name.strip(), focal_person.strip(), focal_contact.strip(), primary_tech, tour_target_session.strip(),
                        str(tour_date), dep_time.strip(), travel_mode, other_mode.strip(), students_count,
                        driver_name.strip(), conductor_name.strip(), aipo_name.strip(), assistant_name.strip(), approval_copy,
                        focal_person.strip(), "Industrial Representative", "Excellent", "Excellent", "Excellent", "Good", "Good", "Excellent",
                        "Yes", "Good professional behavior.", "YES", "YES", "YES", "YES", 5, 5, 5, 5,
                        tour_limits_json, 0.0, "Planned"
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

            c_g, c_h, c_i = st.columns(3)
            with c_g:
                st.download_button("📥 Re-Download Annexure G", generate_annexure_g(archive_v, std_df), f"Archived_Annex_G_{re_tid}.docx")
            with c_h:
                st.download_button("📥 Re-Download Annexure H", generate_annexure_h(archive_v), f"Archived_Annex_H_{re_tid}.docx")
            with c_i:
                st.download_button("📥 Re-Download 8-Point Annexure I", generate_annexure_i(archive_v), f"Archived_Annex_I_{re_tid}.docx")

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

    pl = "%s" if is_postgres else "?"

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
                    ph_insert = ",".join([pl] * 45)
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
                        candidate_limits_json, budget_utilized, status
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
                        limits_json_str, 0.0, "Planned"
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

            c_j, c_k, c_l, c_xl = st.columns(4)
            with c_j:
                st.download_button("📥 Re-Download Annexure J", generate_annexure_j(archive_dr, re_cand_df), f"Archived_Annex_J_{re_did}.docx")
            with c_k:
                st.download_button(f"📥 Re-Download Annexure K ({len(re_part_df)} Cands)", generate_annexure_k_excel(archive_dr, re_part_df), f"Archived_Annex_K_{re_did}.xlsx")
            with c_l:
                st.download_button("📥 Re-Download Annexure L", generate_annexure_l(archive_dr), f"Archived_Annex_L_{re_did}.docx")
            with c_xl:
                st.download_button("📥 Re-Download Roster Excel", generate_candidate_roster_excel(archive_dr, re_cand_df), f"Archived_Candidates_{re_did}.xlsx")

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
