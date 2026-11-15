"""One-time, non-destructive migration from ipo_management.db to PostgreSQL.
Run only AFTER launching the corrected portal once against PostgreSQL and logging in,
so the portal's tables/migrations have been created. Existing cloud rows are preserved.
"""
import os, sqlite3, sys
import psycopg2

SQLITE_PATH = os.environ.get('IPO_SQLITE_PATH', 'ipo_management.db')
DATABASE_URL = os.environ.get('DATABASE_URL')
if not DATABASE_URL:
    print('ERROR: Set DATABASE_URL to your Supabase/PostgreSQL connection string first.')
    sys.exit(1)
if not os.path.exists(SQLITE_PATH):
    print(f'ERROR: Local SQLite database not found: {SQLITE_PATH}')
    sys.exit(1)

TABLES = [
    'system_users', 'students', 'traceability_forms', 'activities',
    'guest_lectures', 'industry_visits', 'recruitment_drives',
    'monthly_activity_plans'
]

src = sqlite3.connect(SQLITE_PATH)
src.row_factory = sqlite3.Row
dst = psycopg2.connect(DATABASE_URL, sslmode='require', connect_timeout=20)
try:
    with dst:
        with dst.cursor() as cur:
            for table in TABLES:
                exists = src.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (table,)).fetchone()
                if not exists:
                    print(f'SKIP {table}: not present in local database')
                    continue
                local_cols = [r['name'] for r in src.execute(f'PRAGMA table_info("{table}")').fetchall()]
                cur.execute("SELECT column_name FROM information_schema.columns WHERE table_schema=current_schema() AND table_name=%s ORDER BY ordinal_position", (table,))
                pg_cols = [r[0] for r in cur.fetchall()]
                cols = [c for c in local_cols if c in pg_cols]
                if not cols:
                    print(f'SKIP {table}: no common columns (initialize portal schema first)')
                    continue
                quoted = ', '.join('"' + c.replace('"','""') + '"' for c in cols)
                placeholders = ', '.join(['%s'] * len(cols))
                insert = f'INSERT INTO "{table}" ({quoted}) VALUES ({placeholders}) ON CONFLICT DO NOTHING'
                rows = src.execute(f'SELECT {quoted} FROM "{table}"').fetchall()
                before = cur.rowcount
                for row in rows:
                    cur.execute(insert, tuple(row[c] for c in cols))
                inserted = cur.rowcount
                print(f'{table}: examined {len(rows):,} local rows; cloud duplicates were skipped')
                # Keep SERIAL/IDENTITY counters ahead of imported IDs.
                if 'id' in cols:
                    cur.execute("SELECT pg_get_serial_sequence(%s, 'id')", (table,))
                    seq = cur.fetchone()[0]
                    if seq:
                        cur.execute(f'SELECT MAX(id) FROM "{table}"')
                        max_id = cur.fetchone()[0]
                        if max_id is not None:
                            cur.execute('SELECT setval(%s, %s, true)', (seq, max_id))
    print('\nMigration finished. Your local SQLite file was NOT changed or deleted.')
except Exception:
    dst.rollback()
    raise
finally:
    src.close(); dst.close()
