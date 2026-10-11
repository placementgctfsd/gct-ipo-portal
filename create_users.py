import sqlite3

conn = sqlite3.connect("ipo_management.db")
c = conn.cursor()

# Naye columns add karna
try:
    c.execute("ALTER TABLE system_users ADD COLUMN email TEXT")
except:
    pass

try:
    c.execute("ALTER TABLE system_users ADD COLUMN otp TEXT")
except:
    pass

# Admin (Wasif) ke account ke sath apni recovery email link karein
# Neeche apni actual Gmail address likhein
admin_email = "placement.gctfsd@gmail.com"  # <-- Yahan apna Gmail address likhein

c.execute("UPDATE system_users SET email = ? WHERE username = 'wasif'", (admin_email.strip().lower(),))

conn.commit()
conn.close()
print("Success! Database columns update ho gaye aur Admin email save ho gayi.")
