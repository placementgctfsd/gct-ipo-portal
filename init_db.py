import sqlite3

def init_db():
    conn = sqlite3.connect("ipo_management.db")
    cursor = conn.cursor()

    # 1. Students Master Table (CNIC as primary unique identifier)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS students (
        cnic TEXT PRIMARY KEY,
        roll_no TEXT,
        name TEXT,
        technology TEXT,
        session TEXT,
        phone TEXT,
        status TEXT DEFAULT 'Enrolled'
    )
    """)

    # 2. Activities Table (Visits, Guest Lectures, Drives)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS activities (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        activity_type TEXT,
        title TEXT,
        company_or_speaker TEXT,
        activity_date TEXT,
        venue TEXT,
        details TEXT
    )
    """)

    # 3. Student Activity Participation & Traceability
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS activity_participants (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        activity_id INTEGER,
        student_cnic TEXT,
        status TEXT, -- e.g., 'Attended', 'Shortlisted', 'Hired'
        remarks TEXT,
        FOREIGN KEY (activity_id) REFERENCES activities (id),
        FOREIGN KEY (student_cnic) REFERENCES students (cnic)
    )
    """)

    conn.commit()
    conn.close()
    print("Database & Tables Initialized Successfully!")

if __name__ == "__main__":
    init_db()
