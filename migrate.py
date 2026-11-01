import sqlite3
import pandas as pd
from sqlalchemy import create_engine

# 1. Local SQLite connection
local_conn = sqlite3.connect("ipo_management.db")

# 2. Supabase PostgreSQL connection (Pooler URL on port 6543)
SUPABASE_URL = "postgresql+psycopg2://postgres.mcmesfmsshwfxvtkhnoh:IPO%402025!db@aws-0-ap-southeast-1.pooler.supabase.com:6543/postgres"

print("Connecting to Supabase...")
engine = create_engine(SUPABASE_URL)

# 3. Aapke local database ke exact tables ki list
tables = [
    "students", 
    "activities",
    "activity_participants",
    "recruitment_drives", 
    "guest_lectures", 
    "industry_visits",
    "monthly_activity_plans",
    "traceability_forms",
    "system_users"
]

# Check existing tables in local DB
cursor = local_conn.cursor()
cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
existing_local_tables = [row[0] for row in cursor.fetchall()]

for table in tables:
    if table in existing_local_tables:
        try:
            df = pd.read_sql(f"SELECT * FROM {table}", local_conn)
            if not df.empty:
                df.to_sql(table, engine, if_exists="replace", index=False)
                print(f"✅ Table '{table}' successfully migrated ({len(df)} records).")
            else:
                print(f"⚠️ Table '{table}' khali thi, skip kar di.")
        except Exception as e:
            print(f"❌ Error migrating '{table}': {e}")
    else:
        print(f"ℹ️ Table '{table}' local DB mein nahi mili.")

print("\n🎉 Migration process complete!")
