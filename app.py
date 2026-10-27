import streamlit as st
import pandas as pd
import sqlite3
import hashlib
import io
import re
import json
import os
import zlib
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
    # 1. Agar Streamlit Cloud Secrets mein Supabase URL ho toh Supabase use karega
    if "postgres" in st.secrets:
        try:
            pg_url = st.secrets["postgres"]["url"]
            conn = psycopg2.connect(pg_url)
            return conn, "postgres"
        except Exception as e:
            st.error(f"Supabase connection error: {e}")
            
    # 2. Local Fallback (Laptop ke liye)
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

def generate_annexure_j(data, student_df=None):
    doc = Document()
    
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    t_run = title.add_run('Annexure “J”\nRECRUITMENT DRIVE APPROVAL FORM\n')
    t_run.bold = True
    t_run.font.size = Pt(13)

    p_comp = doc.add_paragraph()
    p_comp.paragraph_format.line_spacing = 1.15
    p_comp.add_run("Company Profile:\n").bold = True
    p_comp.add_run(f"NAME OF THE COMPANY:  {data.get('company_name', '').upper()}\n")
    p_comp.add_run(f"ADDRESS OF COMPANY:  {data.get('company_address', '').upper()}\n\n")
    
    p_comp.add_run("JOB DETAILS:\n").bold = True
    p_comp.add_run(f"JOB SITE LOCATION:  {data.get('job_location', '').upper()}\n")

    table = doc.add_table(rows=2, cols=8)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers = [
        "Required Experience", "Job Title", "No. of Vacancies (M)", 
        "No. of Vacancies (F)", "Trade / Technology", "Marks / Division", 
        "Salary (PKR)", "Job Description"
    ]
    for i, h in enumerate(headers):
        c = table.cell(0, i)
        c.text = h
        c.paragraphs[0].runs[0].bold = True
        c.paragraphs[0].runs[0].font.size = Pt(8)

    techs_str = ", ".join(data.get('target_techs', []))
    row_vals = [
        data.get('experience', 'Fresh / Result Awaited'),
        data.get('job_title', 'Trainee Engineer / Associate'),
        str(data.get('vacancies_male', 20)),
        str(data.get('vacancies_female', 0)),
        techs_str,
        data.get('marks_div', '1st Division'),
        str(data.get('salary_range', '35,000 - 45,000')),
        data.get('job_desc', 'Operational & Technical responsibilities')
    ]
    for i, v in enumerate(row_vals):
        c = table.cell(1, i)
        c.text = str(v)
        c.paragraphs[0].runs[0].font.size = Pt(8)

    doc.add_paragraph()

    p_opts = doc.add_paragraph()
    p_opts.paragraph_format.line_spacing = 1.2
    
    m_rec = data.get('recruitment_mode', 'Permanent')
    p_opts.add_run("Please mention the mode of initial recruitment that candidate shall attain under above said drive?\n").bold = True
    p_opts.add_run(f"{'☑' if m_rec == 'Permanent' else '☐'} Permanent      {'☑' if m_rec == 'Contractual' else '☐'} Contractual      {'☑' if m_rec == 'Apprentices' else '☐'} Apprentices      {'☑' if m_rec == 'Internship / OJT' else '☐'} Internship / OJT\n\n")

    p_opts.add_run("Perks and Privileges:\n").bold = True
    eobi = "☑" if data.get('eobi', True) else "☐"
    soc = "☑" if data.get('social_sec', True) else "☐"
    bon = "☑" if data.get('bonuses', True) else "☐"
    p_opts.add_run(f"{eobi} Employees’ Old Age Benefits Act, 1976.\n{soc} Provincial Employees’ Social Security Ordinance, 1965.\n{bon} Bonuses / Annual Increments\n\n")

    p_opts.add_run("Accommodation options available for selected candidate’s on/close vicinity of site?\n").bold = True
    acc = data.get('accommodation', 'Company maintained accommodation')
    p_opts.add_run(f"{'☑' if acc == 'Company maintained accommodation' else '☐'} Company maintained accommodation\n{'☑' if acc == 'Company paid / subsidized accommodation' else '☐'} Company paid / subsidized accommodation\n{'☑' if acc == 'Any other benefit (Food etc.)' else '☐'} Any other benefit (Food / Transport etc.)\n\n")

    d_table = doc.add_table(rows=2, cols=6)
    d_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    d_headers = [
        "Venue", "Date & Time", "DPO & IPO Name", 
        "TEVTA Focal Contact", "Employer Focal Person", "Employer Focal Contact"
    ]
    for i, h in enumerate(d_headers):
        c = d_table.cell(0, i)
        c.text = h
        c.paragraphs[0].runs[0].bold = True
        c.paragraphs[0].runs[0].font.size = Pt(8)

    d_vals = [
        data.get('venue', 'GCT Samanabad FSD'),
        f"{data.get('drive_date', '')} @ {data.get('drive_time', '')}",
        f"{data.get('dpo_name', 'DPO')}\n{data.get('ipo_name', 'IPO')}",
        data.get('tevta_contact', '0300-1234567'),
        f"{data.get('emp_focal', '')}\n({data.get('emp_desig', '')})",
        data.get('emp_contact', '')
    ]
    for i, v in enumerate(d_vals):
        c = d_table.cell(1, i)
        c.text = str(v)
        c.paragraphs[0].runs[0].font.size = Pt(8)

    doc.add_paragraph()
    p_sign = doc.add_paragraph()
    p_sign.add_run("OFFICE USE ONLY\n").bold = True
    p_sign.add_run("Initiated by (Director concerned / Head of concerned zone): __________________   Dated: _________\n")
    p_sign.add_run("Approved by: _____________________________________________________   Dated: _________\n\n")

    doc_page_break = doc.add_page_break()
    p_list = doc.add_paragraph()
    p_list.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_list.add_run(f"LIST OF ELIGIBLE CANDIDATES ENCLOSED FOR RECRUITMENT DRIVE\nCompany: {data.get('company_name', '').upper()} | Session: {data.get('session', '')}\n").bold = True

    s_table = doc.add_table(rows=1, cols=7)
    s_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    s_headers = ["Sr#", "Roll No", "Student Name", "Trade / Technology", "CNIC", "Mobile Number", "Email"]
    for i, h in enumerate(s_headers):
        c = s_table.cell(0, i)
        c.text = h
        c.paragraphs[0].runs[0].bold = True
        c.paragraphs[0].runs[0].font.size = Pt(8)

    if student_df is not None and not student_df.empty:
        for idx, (_, row) in enumerate(student_df.iterrows(), start=1):
            rc = s_table.add_row().cells
            rc[0].text = str(idx)
            rc[1].text = str(row.get('Roll No', ''))
            rc[2].text = str(row.get('Name', ''))
            rc[3].text = str(row.get('Technology', ''))
            rc[4].text = str(row.get('CNIC', ''))
            rc[5].text = str(row.get('Mobile Number', ''))
            rc[6].text = str(row.get('Email', ''))
            for cell in rc:
                cell.paragraphs[0].runs[0].font.size = Pt(8)
    else:
        rc = s_table.add_row().cells
        for cell in rc:
            cell.text = "-"

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

def generate_candidate_roster_excel(data, student_df):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Candidate Roster"

    ws.merge_cells("A1:G1")
    t = ws["A1"]
    t.value = f"RECRUITMENT DRIVE CANDIDATE ROSTER - {data.get('company_name', '').upper()}"
    t.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    t.fill = PatternFill("solid", fgColor="1F4E79")
    t.alignment = Alignment(horizontal="center", vertical="center")

    ws["A2"] = f"Drive Date: {data.get('drive_date', '')}"
    ws["D2"] = f"Target Session: {data.get('session', '')}"
    ws["F2"] = f"Venue: {data.get('venue', 'GCT Samanabad')}"
    for c in ["A2", "D2", "F2"]:
        ws[c].font = Font(bold=True)

    headers = ["Sr#", "Roll No", "Student Name", "Technology", "CNIC", "Mobile Number", "Email Address"]
    for col_idx, h in enumerate(headers, start=1):
        c = ws.cell(row=4, column=col_idx)
        c.value = h
        c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="2F5597")
        c.alignment = Alignment(horizontal="center", vertical="center")

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
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
                ws.cell(row=current_row, column=c_idx).font = Font(name="Calibri", size=9)
            current_row += 1

    col_widths = {1: 8, 2: 14, 3: 25, 4: 20, 5: 18, 6: 16, 7: 25}
    for c_idx, width in col_widths.items():
        ws.column_dimensions[openpyxl.utils.get_column_letter(c_idx)].width = width

    excel_io = io.BytesIO()
    wb.save(excel_io)
    excel_io.seek(0)
    return excel_io

def generate_annexure_k_excel(data, student_df=None):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Assessment Sheet"

    ws.merge_cells("A1:M1")
    t_cell = ws["A1"]
    t_cell.value = "Assesment Sheet for Campus Recruitment Drive in Punjab TEVTA's Institutes"
    t_cell.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    t_cell.fill = PatternFill("solid", fgColor="1F4E79")
    t_cell.alignment = Alignment(horizontal="center", vertical="center")

    ws["A2"] = f"Interviewer: {data.get('company_name', '')}"
    ws["F2"] = "Venue:"
    ws["H2"] = data.get('venue', 'GCT Samanabad FSD')
    ws["K2"] = "Date of Interview:"
    ws["L2"] = str(data.get('drive_date', ''))

    ws["A3"] = f"Time: {data.get('drive_time', '10:00 AM')}"
    ws["F3"] = "Required Trades:"
    ws["H3"] = ", ".join(data.get('target_techs', []))
    ws["K3"] = "Expected Salary:"
    ws["L3"] = str(data.get('salary_range', ''))

    ws.merge_cells("A4:E4")
    ws["A4"] = "Interview evaluation Sheet are to be completed by the District Placement Officer to rank the candidate's. Under each heading, DPO should give the candidate a numerical rating."
    ws["A4"].font = Font(size=8, italic=True)
    ws.merge_cells("F4:M4")
    ws["F4"] = "Scale:    1- Un Satisfactory,   2- Satisfactory,   3- Average,   4- Good,   5- Excellent"
    ws["F4"].font = Font(size=8, bold=True)

    for r in range(2, 5):
        for c in range(1, 14):
            cell = ws.cell(row=r, column=c)
            if r in [2, 3] and cell.value and any(k in str(cell.value) for k in ["Venue", "Date", "Trades", "Salary", "Time", "Interviewer"]):
                cell.font = Font(bold=True)

    headers = [
        "Sr. No", "Candidate Name", "Trade", "Session", "Institute",
        "Quality of CV", "Verbal Communication", "Candidate Interest",
        "Interpersonal Skills", "Prior work Experience", "Technical Knowledge", "Overall Impression", "Remarks"
    ]
    for col_idx, h in enumerate(headers, start=1):
        c = ws.cell(row=5, column=col_idx)
        c.value = h
        c.font = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="2F5597")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    current_row = 6
    if student_df is not None and not student_df.empty:
        for idx, (_, row) in enumerate(student_df.iterrows(), start=1):
            ws.cell(row=current_row, column=1, value=idx).alignment = Alignment(horizontal="center")
            ws.cell(row=current_row, column=2, value=str(row.get('Name', '')))
            ws.cell(row=current_row, column=3, value=str(row.get('Technology', '')))
            ws.cell(row=current_row, column=4, value=str(data.get('session', '')))
            ws.cell(row=current_row, column=5, value="GCT Samanabad")
            
            for c_idx in range(6, 14):
                ws.cell(row=current_row, column=c_idx, value="")

            for c_idx in range(1, 14):
                ws.cell(row=current_row, column=c_idx).border = thin_border
                ws.cell(row=current_row, column=c_idx).font = Font(name="Calibri", size=9)
            current_row += 1
    else:
        for r_offset in range(15):
            r_num = current_row + r_offset
            ws.cell(row=r_num, column=1, value=r_offset + 1).alignment = Alignment(horizontal="center")
            ws.cell(row=r_num, column=5, value="GCT Samanabad")
            for c_idx in range(1, 14):
                ws.cell(row=r_num, column=c_idx).border = thin_border
                ws.cell(row=r_num, column=c_idx).font = Font(name="Calibri", size=9)
        current_row += 15

    sig_row = current_row + 2
    ws.cell(row=sig_row, column=2, value=f"Name of DPO: {data.get('dpo_name', 'Mr. Muhammad Asghar Ch (DPO)')}").font = Font(bold=True)
    ws.cell(row=sig_row, column=10, value="Signature of DPO with date: ______________________").font = Font(bold=True)

    col_widths = {1: 7, 2: 22, 3: 16, 4: 12, 5: 16, 6: 12, 7: 15, 8: 14, 9: 14, 10: 16, 11: 15, 12: 15, 13: 18}
    for c_idx, width in col_widths.items():
        ws.column_dimensions[openpyxl.utils.get_column_letter(c_idx)].width = width

    excel_io = io.BytesIO()
    wb.save(excel_io)
    excel_io.seek(0)
    return excel_io

def generate_annexure_l(data):
    doc = Document()
    
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    t_run = title.add_run('Annexure “L”\nEmployer’s Feedback Form for Recruitment Drive\n')
    t_run.bold = True
    t_run.font.size = Pt(13)

    p_intro = doc.add_paragraph()
    p_intro.paragraph_format.line_spacing = 1.15
    intro_txt = (
        f"Thank you for scheduling On-Campus Recruiting at {data.get('venue', 'GCT Samanabad, Faisalabad')} on Date {data.get('drive_date', '')}.\n"
        f"We hope your efforts are fruitful and your experience conducting a recruitment drive has been on the positive side.\n"
        f"Please take a few minutes of your precious time to answer the questions below. This information will provide valuable feedback for our students and staff in order to conduct better recruitment drives in future.\n\n"
        f"Please review the options describing your experience (E = Excellent, G = Good, UN = Unsatisfactory):"
    )
    p_intro.add_run(intro_txt)

    table = doc.add_table(rows=1, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers = ["Feedback Parameter", "Excellent (E)", "Good (G)", "Unsatisfactory (UN)"]
    for i, h in enumerate(headers):
        c = table.cell(0, i)
        c.text = h
        c.paragraphs[0].runs[0].bold = True
        c.paragraphs[0].runs[0].font.size = Pt(9)

    params_list = [
        ("Feedback regarding Venue:", "", "", ""),
        ("Coordination of recruiting schedule by Placement Officer", data.get('l_coordination', 'E')),
        ("Staff of institute / venue was friendly & helpful", data.get('l_staff', 'E')),
        ("Interview venue was comfortable & equipped", data.get('l_venue_comfort', 'E')),
        ("Feedback regarding Students:", "", "", ""),
        ("Students Technical skills met your expectations", data.get('l_tech_skills', 'E')),
        ("Students resumes were professional", data.get('l_resumes', 'G')),
        ("Students communication skills were commendable", data.get('l_comm_skills', 'G')),
        ("Students were formally dressed", data.get('l_dress', 'E')),
        ("Overall impression of recruitment drive", data.get('l_overall', 'E'))
    ]

    for item in params_list:
        row = table.add_row().cells
        if item[1] == "":
            row[0].text = item[0]
            row[0].paragraphs[0].runs[0].bold = True
            row[0].paragraphs[0].runs[0].font.size = Pt(9)
        else:
            row[0].text = item[0]
            row[0].paragraphs[0].runs[0].font.size = Pt(9)
            val = item[1]
            row[1].text = "☑ E" if val == "E" else "☐ E"
            row[2].text = "☑ G" if val == "G" else "☐ G"
            row[3].text = "☑ UN" if val == "UN" else "☐ UN"
            for col_i in [1, 2, 3]:
                row[col_i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                row[col_i].paragraphs[0].runs[0].font.size = Pt(9)

    doc.add_paragraph()
    p_res = doc.add_paragraph()
    p_res.paragraph_format.line_spacing = 1.2
    p_res.add_run(f"No. of students appeared:  {data.get('students_appeared', '')}\n")
    p_res.add_run(f"No. of students shortlisted:  {data.get('students_shortlisted', '')}\n")
    p_res.add_run(f"Suggestions (if any):  {data.get('l_suggestions', 'Satisfied with overall academic caliber and discipline.')}\n\n")

    p_sig = doc.add_paragraph()
    p_sig.paragraph_format.line_spacing = 1.2
    p_sig.add_run(f"Industry / Organization:  {data.get('company_name', '')}\n")
    p_sig.add_run(f"HR Representative Name:  {data.get('emp_focal', '')}          Designation:  {data.get('emp_desig', '')}\n")
    p_sig.add_run(f"Mobile / Phone:  {data.get('emp_contact', '')}          Email:  {data.get('emp_email', 'hr@industry.com')}\n\n\n")
    p_sig.add_run("_________________________________________\nSignature & Stamp with Date")

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# ----------------- Helper Functions: Direct HTML Print & Annual Report -----------------

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

# --- FORCE SUPABASE CONNECTION ---
import psycopg2

SUPABASE_URL = "postgresql://postgres.mcmesfmsshwfxvtkhnoh:IPO%402025!db@aws-0-ap-southeast-1.pooler.supabase.com:6543/postgres"

try:
    if "postgres" in st.secrets:
        conn = psycopg2.connect(st.secrets["postgres"]["url"])
    else:
        conn = psycopg2.connect(SUPABASE_URL)
    cursor = conn.cursor()
except Exception as e:
    st.error(f"Supabase connection failed: {e}")

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
                if uploaded_file.name.lower().endswith('.csv'):
                    uploaded_df = pd.read_csv(uploaded_file, dtype=str)
                else:
                    uploaded_df = pd.read_excel(uploaded_file, dtype=str)

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
                if comp_file.name.lower().endswith('.csv'):
                    df_up = pd.read_csv(comp_file, dtype=str)
                else:
                    df_up = pd.read_excel(comp_file, dtype=str)

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
# --- 6. Recruitment Drive ---
# =========================================================================
elif choice == "Recruitment Drive":
    st.subheader("On-Campus Recruitment Drive Management & Documentation")
    
    # Ensure new columns exist in table safely
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
                    cursor.execute("SELECT COUNT(*) FROM students WHERE session = ? AND technology = ?", (target_session, t))
                    avail = cursor.fetchone()[0]
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
                        q_t = f"SELECT roll_no AS 'Roll No', name AS 'Name', technology AS 'Technology', cnic AS 'CNIC', phone AS 'Mobile Number', email AS 'Email' FROM students WHERE session = ? AND technology = ? ORDER BY roll_no LIMIT {cnt}"
                        df_sub = pd.read_sql(q_t, conn, params=[target_session, t])
                        if not df_sub.empty:
                            all_eligible_list.append(df_sub)
        else:
            for t in target_techs:
                cursor.execute("SELECT COUNT(*) FROM students WHERE session = ? AND technology = ?", (target_session, t))
                avail = cursor.fetchone()[0]
                custom_counts[t] = avail
                q_t = "SELECT roll_no AS 'Roll No', name AS 'Name', technology AS 'Technology', cnic AS 'CNIC', phone AS 'Mobile Number', email AS 'Email' FROM students WHERE session = ? AND technology = ? ORDER BY roll_no"
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
                    cursor.execute("""
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
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
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

    with rd_tab2:
        df_drives = pd.read_sql("SELECT id, company_name, drive_date, session, target_techs, venue, drive_time, status FROM recruitment_drives ORDER BY id DESC", conn)
        
        if not df_drives.empty:
            drive_map = {f"ID {r['id']} - {r['company_name']} ({r['drive_date']}) [{r['status']}]": r['id'] for _, r in df_drives.iterrows()}
            selected_drive_label = st.selectbox("Select Recruitment Drive to Finalize:", list(drive_map.keys()))
            selected_drive_id = drive_map[selected_drive_label]

            cursor.execute("SELECT * FROM recruitment_drives WHERE id = ?", (selected_drive_id,))
            dr_row = cursor.fetchone()
            dr_cols = [d[0] for d in cursor.description]
            dr_data = dict(zip(dr_cols, dr_row))
            dr_data['target_techs'] = dr_data['target_techs'].split(",") if dr_data['target_techs'] else []

            st.write(f"**Company:** {dr_data['company_name']} | **Date:** {dr_data['drive_date']} | **Trades:** {', '.join(dr_data['target_techs'])}")

            st.markdown("---")
            st.markdown("#### 📝 Annexure “K” Assessment Sheet Setup")
            st.caption("Annexure 'K' unhi students ke liye generate karein jinho nay test ya interview mein hissa liya:")

            saved_part_cnics = []
            if dr_data.get('participated_cnics_json'):
                try:
                    saved_part_cnics = json.loads(dr_data['participated_cnics_json'])
                except:
                    saved_part_cnics = []

            k_gen_mode = st.radio(
                "Assessment Sheet Candidate Selection:",
                ["🎯 Filter by Participated Students (Upload / Paste CNICs)", "📋 Use Default Eligible Roster (From Plan)"],
                horizontal=True
            )

            cand_df = pd.DataFrame(columns=['Roll No', 'Name', 'Technology', 'Session'])
            unmatched_k = []
            k_raw_cnics = []

            if k_gen_mode == "🎯 Filter by Participated Students (Upload / Paste CNICs)":
                k_t1, k_t2 = st.tabs(["📁 Upload File (.xlsx / .csv)", "✍️ Paste CNICs Direct"])

                with k_t1:
                    k_file = st.file_uploader("Upload Interview/Test Participated Students File", type=["xlsx", "csv"], key="k_cand_file")
                    if k_file is not None:
                        try:
                            if k_file.name.lower().endswith(".csv"):
                                df_k_up = pd.read_csv(k_file, dtype=str)
                            else:
                                df_k_up = pd.read_excel(k_file, dtype=str)
                            
                            c_col = None
                            for c in df_k_up.columns:
                                if "cnic" in str(c).lower():
                                    c_col = c
                                    break
                            if not c_col:
                                c_col = st.selectbox("Select CNIC column from sheet:", df_k_up.columns, key="k_sheet_col")
                            
                            if c_col:
                                k_raw_cnics = df_k_up[c_col].dropna().tolist()
                        except Exception as e:
                            st.error(f"File error: {e}")

                with k_t2:
                    default_k_text = "\n".join(saved_part_cnics) if saved_part_cnics else ""
                    k_text = st.text_area("Interview mein shareek students k CNICs paste karein (Har line par aik CNIC):", value=default_k_text, placeholder="33100-1234567-1\n33100-7654321-2", key="k_text_area")
                    if k_text.strip():
                        k_raw_cnics = re.split(r"[\n,\s;]+", k_text.strip())

                if not k_raw_cnics and saved_part_cnics:
                    k_raw_cnics = saved_part_cnics

                if k_raw_cnics:
                    target_k_dashed, lookup_k = parse_cnic_input(k_raw_cnics)
                    if lookup_k:
                        placeholders_k = ",".join("?" * len(lookup_k))
                        sql_k_query = f"""
                            SELECT roll_no AS 'Roll No', name AS 'Name', technology AS 'Technology', session AS 'Session', cnic AS 'CNIC'
                            FROM students 
                            WHERE cnic IN ({placeholders_k})
                            ORDER BY technology, roll_no
                        """
                        cand_df = pd.read_sql(sql_k_query, conn, params=lookup_k)
                        found_k = set(cand_df['CNIC'].tolist())
                        unmatched_k = [c for c in target_k_dashed if c not in found_k]

                        st.success(f"✅ Total **{len(cand_df)}** Participated Candidates active hain.")
                        
                        col_save_k, col_clr_k = st.columns([1, 1])
                        with col_save_k:
                            if st.button("💾 Save Participated Candidates to Database", key="btn_save_part_cands"):
                                cursor.execute("""
                                    UPDATE recruitment_drives 
                                    SET participated_cnics_json = ?, students_appeared = ?
                                    WHERE id = ?
                                """, (json.dumps(list(found_k)), len(cand_df), selected_drive_id))
                                conn.commit()
                                dr_data['participated_cnics_json'] = json.dumps(list(found_k))
                                dr_data['students_appeared'] = len(cand_df)
                                st.success(f"✅ {len(cand_df)} Participated Candidates successfully save ho gaye!")
                                st.rerun()

                        with col_clr_k:
                            if saved_part_cnics:
                                if st.button("🔄 Clear / Reset Participated List", key="btn_clr_part_cands"):
                                    cursor.execute("""
                                        UPDATE recruitment_drives 
                                        SET participated_cnics_json = '[]', students_appeared = 0
                                        WHERE id = ?
                                    """, (selected_drive_id,))
                                    conn.commit()
                                    st.warning("Participated list clear ho gayi hai. Ab aap fresh list daal sakte hain.")
                                    st.rerun()

                        # --- EDIT / REMOVE / ADD SINGLE CANDIDATE SECTION ---
                        if not cand_df.empty:
                            with st.expander("🛠️ Edit / Galti Durust Karein (Participated Candidates Management)"):
                                cand_remove_options = {
                                    f"{r['Roll No']} - {r['Name']} ({r['Technology']}) [{r['CNIC']}]": r['CNIC']
                                    for _, r in cand_df.iterrows()
                                }
                                to_remove_label = st.selectbox("Galat candidate select karein jise hatana hai:", list(cand_remove_options.keys()), key="sel_cand_remove")
                                if st.button("❌ Remove Selected Candidate From List", key="btn_remove_single_cand"):
                                    cnic_to_del = cand_remove_options[to_remove_label]
                                    updated_cnics = [c for c in cand_df['CNIC'].tolist() if c != cnic_to_del]
                                    cursor.execute("""
                                        UPDATE recruitment_drives 
                                        SET participated_cnics_json = ?, students_appeared = ?
                                        WHERE id = ?
                                    """, (json.dumps(updated_cnics), len(updated_cnics), selected_drive_id))
                                    conn.commit()
                                    st.success(f"Candidate {to_remove_label} list se hata diya gaya!")
                                    st.rerun()

                                st.markdown("---")
                                st.markdown("**➕ Ya koi missing candidate shamil karna chahein:**")
                                new_add_cnic = st.text_input("New Student CNIC to Add:", placeholder="33100-1234567-1", key="add_new_part_cnic_input")
                                if st.button("➕ Add This CNIC to Participated List", key="btn_add_part_single_cnic"):
                                    if new_add_cnic.strip():
                                        _, lookup_new = parse_cnic_input([new_add_cnic.strip()])
                                        if lookup_new:
                                            cursor.execute("SELECT cnic FROM students WHERE cnic = ?", (lookup_new[0],))
                                            row_verify = cursor.fetchone()
                                            if row_verify:
                                                curr_cnics = cand_df['CNIC'].tolist()
                                                if row_verify[0] not in curr_cnics:
                                                    curr_cnics.append(row_verify[0])
                                                    cursor.execute("UPDATE recruitment_drives SET participated_cnics_json = ?, students_appeared = ? WHERE id = ?", (json.dumps(curr_cnics), len(curr_cnics), selected_drive_id))
                                                    conn.commit()
                                                    st.success("Candidate kamyabi se shamil ho gaya!")
                                                    st.rerun()
                                                else:
                                                    st.info("Yeh candidate pehle se list mein mojood hai.")
                                            else:
                                                st.error("Diya gaya CNIC student directory mein nahi mila.")

                        if unmatched_k:
                            st.warning(f"⚠️ **{len(unmatched_k)}** CNICs aapke Student Directory mein mojood nahi hain!")
                            with st.expander("Dekhein missing CNICs"):
                                st.write(unmatched_k)
                            
                            unmatched_df = pd.DataFrame({"Missing_CNIC": unmatched_k})
                            unm_buf = io.BytesIO()
                            with pd.ExcelWriter(unm_buf, engine='openpyxl') as wr:
                                unmatched_df.to_excel(wr, index=False, sheet_name="Missing_In_Directory")
                            st.download_button(
                                label="📥 Download Missing CNICs List (.xlsx)",
                                data=unm_buf.getvalue(),
                                file_name=f"Missing_Students_Drive_{dr_data['company_name'].replace(' ', '_')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                else:
                    st.info("Barah-e-karam participate karnay walay students ke CNICs upload ya paste karein.")

            else:
                saved_limits = {}
                if dr_data.get('candidate_limits_json'):
                    try:
                        saved_limits = json.loads(dr_data['candidate_limits_json'])
                    except:
                        pass

                all_cand = []
                for t in dr_data['target_techs']:
                    cnt_lim = saved_limits.get(t)
                    q_k = "SELECT roll_no AS 'Roll No', name AS 'Name', technology AS 'Technology', session AS 'Session', cnic AS 'CNIC' FROM students WHERE session = ? AND technology = ? ORDER BY roll_no"
                    if cnt_lim and cnt_lim > 0:
                        q_k += f" LIMIT {cnt_lim}"
                    df_k_sub = pd.read_sql(q_k, conn, params=[dr_data['session'], t])
                    if not df_k_sub.empty:
                        all_cand.append(df_k_sub)

                cand_df = pd.concat(all_cand, ignore_index=True) if all_cand else pd.DataFrame(columns=['Roll No', 'Name', 'Technology', 'Session', 'CNIC'])

                st.success(f"📋 Roster ke mutabiq total **{len(cand_df)}** Eligible Candidates mojood hain.")

                # OPTION TO SAVE DEFAULT ROSTER AS PARTICIPATED CANDIDATES
                if not cand_df.empty:
                    c_save_roster, _ = st.columns([1, 1])
                    with c_save_roster:
                        if st.button("💾 Save Default Roster Candidates as Annexure K", key="btn_save_roster_as_part"):
                            roster_cnics = cand_df['CNIC'].dropna().tolist()
                            cursor.execute("""
                                UPDATE recruitment_drives 
                                SET participated_cnics_json = ?, students_appeared = ?
                                WHERE id = ?
                            """, (json.dumps(roster_cnics), len(roster_cnics), selected_drive_id))
                            conn.commit()
                            dr_data['participated_cnics_json'] = json.dumps(roster_cnics)
                            dr_data['students_appeared'] = len(roster_cnics)
                            st.success(f"✅ Roster ke {len(roster_cnics)} Candidates Annexure K ke tor par database mein save ho gaye!")
                            st.rerun()

            annex_k_excel = generate_annexure_k_excel(dr_data, cand_df)
            st.download_button(
                f"📊 Download Official Annexure “K” Assessment Sheet ({len(cand_df)} Candidates) (.xlsx)",
                annex_k_excel,
                f"Annexure_K_{dr_data['company_name'].replace(' ', '_')}.xlsx",
                key="btn_dl_annex_k_active"
            )

            st.markdown("---")
            st.markdown("#### 🎯 Post-Interview / Test Finalized Candidates (Selected Students)")
            st.caption("Test aur interview ke baad company ki taraf se **finalize/select kiye gaye** students ke CNICs yahan provide karein:")

            saved_sel_cnics = []
            if dr_data.get('selected_cnics_json'):
                try:
                    saved_sel_cnics = json.loads(dr_data['selected_cnics_json'])
                except:
                    saved_sel_cnics = []

            fn_t1, fn_t2 = st.tabs(["📁 Upload Finalized Students File", "✍️ Paste Finalized CNICs"])
            finalized_raw_cnics = []

            with fn_t1:
                fin_file = st.file_uploader("Upload Company Final Selected / Hired List", type=["xlsx", "csv"], key="final_hired_file")
                if fin_file is not None:
                    try:
                        if fin_file.name.lower().endswith(".csv"):
                            df_fin = pd.read_csv(fin_file, dtype=str)
                        else:
                            df_fin = pd.read_excel(fin_file, dtype=str)
                        
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
                    fin_placeholders = ",".join("?" * len(fin_lookup))
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
                            cursor.execute("""
                                UPDATE recruitment_drives 
                                SET selected_cnics_json = ?, students_shortlisted = ?
                                WHERE id = ?
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

                    # --- EDIT / REMOVE SINGLE SELECTED CANDIDATE SECTION ---
                    with st.expander("🛠️ Edit Selected / Shortlisted List (Remove Incorrect Candidate)"):
                        sel_remove_options = {
                            f"{r['Roll No']} - {r['Name']} ({r['Technology']}) [{r['CNIC']}]": r['CNIC']
                            for _, r in finalized_students_df.iterrows()
                        }
                        to_remove_sel_label = st.selectbox("Galat selected student jise hatana hai:", list(sel_remove_options.keys()), key="sel_to_remove_box")
                        if st.button("❌ Remove From Selected List", key="btn_remove_single_sel"):
                            cnic_sel_del = sel_remove_options[to_remove_sel_label]
                            updated_sel_cnics = [c for c in finalized_students_df['CNIC'].tolist() if c != cnic_sel_del]
                            cursor.execute("""
                                UPDATE recruitment_drives 
                                SET selected_cnics_json = ?, students_shortlisted = ?
                                WHERE id = ?
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

                    st.markdown("---")
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

                    cursor.execute("""
                    UPDATE recruitment_drives SET
                        students_appeared = ?, students_shortlisted = ?,
                        l_coordination = ?, l_staff = ?, l_venue_comfort = ?,
                        l_tech_skills = ?, l_resumes = ?, l_comm_skills = ?, l_dress = ?,
                        l_overall = ?, l_suggestions = ?, budget_utilized = ?, 
                        participated_cnics_json = ?, selected_cnics_json = ?,
                        status = 'Completed'
                    WHERE id = ?
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
            annex_l_docx = generate_annexure_l(dr_data)
            st.download_button("📄 Download Annexure “L” (.docx)", annex_l_docx, f"Annexure_L_{dr_data['company_name'].replace(' ', '_')}.docx")
        else:
            st.info("Pehle Tab 1 se Recruitment Drive plan karein.")

    with rd_tab3:
        st.markdown("##### 📂 Permanent Archived Recruitment Drives, Edit & Delete")
        df_all_drives = pd.read_sql("SELECT id, company_name, drive_date, job_title, target_techs, session, students_appeared, students_shortlisted, budget_utilized, status FROM recruitment_drives ORDER BY id DESC", conn)
        if not df_all_drives.empty:
            st.dataframe(df_all_drives, use_container_width=True)
            
            st.markdown("---")
            st.markdown("###### ⚙️ Manage Archived Drive (Re-download / Edit / Delete)")
            re_did = st.selectbox("Archived Drive Select Karein:", df_all_drives['id'].tolist(), format_func=lambda x: f"ID {x} - {df_all_drives.loc[df_all_drives['id']==x, 'company_name'].values[0]} ({df_all_drives.loc[df_all_drives['id']==x, 'drive_date'].values[0]})")
            
            cursor.execute("SELECT * FROM recruitment_drives WHERE id = ?", (re_did,))
            archive_dr = dict(zip([d[0] for d in cursor.description], cursor.fetchone()))
            archive_dr['target_techs'] = archive_dr['target_techs'].split(",") if archive_dr['target_techs'] else []

            # 1. Roster Candidates
            re_saved_limits = {}
            if archive_dr.get('candidate_limits_json'):
                try:
                    re_saved_limits = json.loads(archive_dr['candidate_limits_json'])
                except:
                    pass

            re_cands = []
            for t in archive_dr['target_techs']:
                cnt_lim = re_saved_limits.get(t)
                q_k = "SELECT roll_no AS 'Roll No', name AS 'Name', technology AS 'Technology', session AS 'Session', cnic AS 'CNIC', phone AS 'Mobile Number', email AS 'Email' FROM students WHERE session = ? AND technology = ? ORDER BY roll_no"
                if cnt_lim and cnt_lim > 0:
                    q_k += f" LIMIT {cnt_lim}"
                df_re_sub = pd.read_sql(q_k, conn, params=[archive_dr['session'], t])
                if not df_re_sub.empty:
                    re_cands.append(df_re_sub)

            re_cand_df = pd.concat(re_cands, ignore_index=True) if re_cands else pd.DataFrame(columns=['Roll No', 'Name', 'Technology', 'Session', 'CNIC', 'Mobile Number', 'Email'])

            # 2. Actual Participated Candidates (Loads saved candidates if present)
            re_part_df = re_cand_df.copy()
            if archive_dr.get('participated_cnics_json'):
                try:
                    p_cnics = json.loads(archive_dr['participated_cnics_json'])
                    if p_cnics:
                        p_holders = ",".join("?" * len(p_cnics))
                        q_part = f"SELECT roll_no AS 'Roll No', name AS 'Name', technology AS 'Technology', session AS 'Session', cnic AS 'CNIC' FROM students WHERE cnic IN ({p_holders}) ORDER BY technology, roll_no"
                        df_part_loaded = pd.read_sql(q_part, conn, params=p_cnics)
                        if not df_part_loaded.empty:
                            re_part_df = df_part_loaded
                except:
                    pass

            # 3. Final Selected Candidates
            re_sel_df = pd.DataFrame()
            if archive_dr.get('selected_cnics_json'):
                try:
                    s_cnics = json.loads(archive_dr['selected_cnics_json'])
                    if s_cnics:
                        s_holders = ",".join("?" * len(s_cnics))
                        q_sel = f"SELECT roll_no AS 'Roll No', name AS 'Name', technology AS 'Technology', session AS 'Session', phone AS 'Mobile Number', cnic AS 'CNIC', email AS 'Email' FROM students WHERE cnic IN ({s_holders}) ORDER BY technology, name"
                        re_sel_df = pd.read_sql(q_sel, conn, params=s_cnics)
                except:
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
                        cursor.execute("""
                            UPDATE recruitment_drives 
                            SET company_name = ?, job_title = ?, students_appeared = ?, students_shortlisted = ?, budget_utilized = ?, status = ?
                            WHERE id = ?
                        """, (e_comp, e_jtitle, e_app, e_short, e_bud_d, e_stat_d, re_did))
                        conn.commit()
                        st.success("Drive data update ho gaya! Refresh karein.")
                        st.rerun()

            with st.expander(f"🗑️ Delete Record (ID: {re_did})"):
                st.warning("⚠️ Kya aap waqai is Recruitment Drive ko permanently delete karna chahte hain?")
                confirm_del_d = st.checkbox("Haan, main is recruitment drive ko delete karna chahta hoon", key=f"del_chk_d_{re_did}")
                if st.button("🔴 Delete This Recruitment Drive", key=f"del_btn_d_{re_did}"):
                    if confirm_del_d:
                        cursor.execute("DELETE FROM recruitment_drives WHERE id = ?", (re_did,))
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
