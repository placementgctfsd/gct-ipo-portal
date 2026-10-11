import sqlite3
import psycopg2
from psycopg2.extras import execute_values

SUPABASE_URI = "postgresql://postgres.mcmesfmsshwfxvtkhnoh:IPO%402025%21db@aws-0-ap-southeast-1.pooler.supabase.com:5432/postgres"

print("Connecting to local and cloud databases...")
conn_local = sqlite3.connect("ipo_management.db")
c_local = conn_local.cursor()

conn_cloud = psycopg2.connect(SUPABASE_URI)
c_cloud = conn_cloud.cursor()

# 1. System Users
print("1/9 Migrating system_users...")
c_local.execute("SELECT username, name, password_hash, role, email, otp FROM system_users")
for row in c_local.fetchall():
    c_cloud.execute("""
        INSERT INTO system_users (username, name, password_hash, role, email, otp)
        VALUES (%s, %s, %s, %s, %s, %s)
        ON CONFLICT (username) DO UPDATE 
        SET name = EXCLUDED.name,
            password_hash = EXCLUDED.password_hash,
            role = EXCLUDED.role,
            email = EXCLUDED.email,
            otp = EXCLUDED.otp;
    """, row)
conn_cloud.commit()

# 2. Students (High-Speed Bulk Insert)
print("2/9 Reading all students from local database...")
c_local.execute("""
    SELECT cnic, roll_no, name, technology, session, phone, status, email, 
           placed_company, placed_salary, placed_date, cv_verified, cv_blob, cv_filename 
    FROM students
""")
students = c_local.fetchall()
total_students = len(students)
print(f"Total students to migrate: {total_students}")

batch_size = 2000
query_students = """
    INSERT INTO students (
        cnic, roll_no, name, technology, session, phone, status, email,
        placed_company, placed_salary, placed_date, cv_verified, cv_blob, cv_filename
    ) VALUES %s
    ON CONFLICT (cnic) DO NOTHING;
"""

for i in range(0, total_students, batch_size):
    batch = students[i:i + batch_size]
    execute_values(c_cloud, query_students, batch, page_size=batch_size)
    conn_cloud.commit()
    print(f"Uploaded {min(i + batch_size, total_students)} / {total_students} students...")

# 3. Activities
print("3/9 Migrating activities...")
c_local.execute("SELECT id, activity_type, title, company_or_speaker, activity_date, venue, details FROM activities")
rows = c_local.fetchall()
if rows:
    execute_values(c_cloud, "INSERT INTO activities (id, activity_type, title, company_or_speaker, activity_date, venue, details) VALUES %s ON CONFLICT (id) DO NOTHING;", rows)
    conn_cloud.commit()

# 4. Activity Participants
print("4/9 Migrating activity_participants...")
c_local.execute("SELECT id, activity_id, student_cnic, status, remarks FROM activity_participants")
rows = c_local.fetchall()
if rows:
    for i in range(0, len(rows), batch_size):
        execute_values(c_cloud, "INSERT INTO activity_participants (id, activity_id, student_cnic, status, remarks) VALUES %s ON CONFLICT (id) DO NOTHING;", rows[i:i + batch_size])
        conn_cloud.commit()

# 5. Guest Lectures
print("5/9 Migrating guest_lectures...")
c_local.execute("SELECT id, speaker_name, speaker_type, qualification, designation, company_org, contact_no, topic, target_tech, target_session, lecture_date, venue, total_attendees, refreshment_heads, backdrop_details, bouquet_details, speaker_rating, student_rating, remarks, budget_utilized, status FROM guest_lectures")
rows = c_local.fetchall()
if rows:
    execute_values(c_cloud, "INSERT INTO guest_lectures (id, speaker_name, speaker_type, qualification, designation, company_org, contact_no, topic, target_tech, target_session, lecture_date, venue, total_attendees, refreshment_heads, backdrop_details, bouquet_details, speaker_rating, student_rating, remarks, budget_utilized, status) VALUES %s ON CONFLICT (id) DO NOTHING;", rows)
    conn_cloud.commit()

# 6. Industry Visits
print("6/9 Migrating industry_visits...")
c_local.execute("SELECT id, industry_name, focal_person, focal_contact, technology, session, tour_date, dep_time, travel_mode, other_mode, students_count, driver_name, conductor_name, aipo_name, assistant_name, approval_copy, rep_name, rep_designation, h_punctuality, h_discipline, h_interest, h_tech_know, h_comm, h_safety, h_recommendation, h_comments, q1_orientation, q2_machinery, r_discipline, r_tech, r_prof, r_satisfaction, budget_utilized, status, q_relevant, q_objectives, candidate_limits_json FROM industry_visits")
rows = c_local.fetchall()
if rows:
    execute_values(c_cloud, "INSERT INTO industry_visits (id, industry_name, focal_person, focal_contact, technology, session, tour_date, dep_time, travel_mode, other_mode, students_count, driver_name, conductor_name, aipo_name, assistant_name, approval_copy, rep_name, rep_designation, h_punctuality, h_discipline, h_interest, h_tech_know, h_comm, h_safety, h_recommendation, h_comments, q1_orientation, q2_machinery, r_discipline, r_tech, r_prof, r_satisfaction, budget_utilized, status, q_relevant, q_objectives, candidate_limits_json) VALUES %s ON CONFLICT (id) DO NOTHING;", rows)
    conn_cloud.commit()

# 7. Recruitment Drives
print("7/9 Migrating recruitment_drives...")
c_local.execute("SELECT id, company_name, company_address, job_location, job_title, experience, vacancies_male, vacancies_female, marks_div, salary_range, job_desc, recruitment_mode, eobi, social_sec, bonuses, accommodation, venue, drive_date, drive_time, written_test, target_techs, session, refreshment_heads, backdrop_details, bouquet_details, dpo_name, ipo_name, tevta_contact, emp_focal, emp_desig, emp_contact, emp_email, students_appeared, students_shortlisted, l_coordination, l_staff, l_venue_comfort, l_tech_skills, l_resumes, l_comm_skills, l_dress, l_overall, l_suggestions, candidate_limits_json, budget_utilized, status FROM recruitment_drives")
rows = c_local.fetchall()
if rows:
    execute_values(c_cloud, "INSERT INTO recruitment_drives (id, company_name, company_address, job_location, job_title, experience, vacancies_male, vacancies_female, marks_div, salary_range, job_desc, recruitment_mode, eobi, social_sec, bonuses, accommodation, venue, drive_date, drive_time, written_test, target_techs, session, refreshment_heads, backdrop_details, bouquet_details, dpo_name, ipo_name, tevta_contact, emp_focal, emp_desig, emp_contact, emp_email, students_appeared, students_shortlisted, l_coordination, l_staff, l_venue_comfort, l_tech_skills, l_resumes, l_comm_skills, l_dress, l_overall, l_suggestions, candidate_limits_json, budget_utilized, status) VALUES %s ON CONFLICT (id) DO NOTHING;", rows)
    conn_cloud.commit()

# 8. Monthly Activity Plans
print("8/9 Migrating monthly_activity_plans...")
c_local.execute("SELECT id, month_year, lec_electrical, lec_mechanical, lec_civil, lec_food, lec_instrument, lec_autofarm, lec_textile, lec_bsc, guest_lectures_expected, industrial_visits_expected, budget_breakdown, total_budget, created_date, lec_shortcourses, recruitment_drives_expected FROM monthly_activity_plans")
rows = c_local.fetchall()
if rows:
    execute_values(c_cloud, "INSERT INTO monthly_activity_plans (id, month_year, lec_electrical, lec_mechanical, lec_civil, lec_food, lec_instrument, lec_autofarm, lec_textile, lec_bsc, guest_lectures_expected, industrial_visits_expected, budget_breakdown, total_budget, created_date, lec_shortcourses, recruitment_drives_expected) VALUES %s ON CONFLICT (id) DO NOTHING;", rows)
    conn_cloud.commit()

# 9. Traceability Forms
print("9/9 Migrating traceability_forms...")
c_local.execute("SELECT id, cnic, student_name, original_filename, uploaded_at, extracted_data_blob, form_blob FROM traceability_forms")
forms = c_local.fetchall()
total_forms = len(forms)
for idx, row in enumerate(forms, 1):
    c_cloud.execute("""
        INSERT INTO traceability_forms (id, cnic, student_name, original_filename, uploaded_at, extracted_data_blob, form_blob)
        VALUES (%s, %s, %s, %s, %s, %s)
        ON CONFLICT (id) DO NOTHING;
    """, row)
    if idx % 10 == 0 or idx == total_forms:
        conn_cloud.commit()
        print(f"Uploaded {idx}/{total_forms} traceability forms...")

conn_cloud.commit()
c_local.close()
conn_local.close()
c_cloud.close()
conn_cloud.close()

print("\nALHAMDULILLAH! Tamam 92,000+ records bulk mode mein kamyabi se upload ho gaye!")
