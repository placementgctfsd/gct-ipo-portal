import hashlib
import sqlite3

def make_hash(password):
    return hashlib.sha256(str.encode(password)).hexdigest()

# Database se connect
conn = sqlite3.connect("ipo_management.db")
c = conn.cursor()

# Users Table Create
c.execute('''
CREATE TABLE IF NOT EXISTS system_users (
    username TEXT PRIMARY KEY,
    name TEXT,
    password_hash TEXT,
    role TEXT
)
''')

# Admin account: Username = wasif, Password = Admin@123
admin_pass = make_hash("Admin@123")
c.execute('''
INSERT OR IGNORE INTO system_users (username, name, password_hash, role)
VALUES (?, ?, ?, ?)
''', ('wasif', 'Muhammad Wasif Ali', admin_pass, 'Administrator'))

conn.commit()
conn.close()
print("Success! Users table ban gaya aur Wasif Admin account create ho gaya.")