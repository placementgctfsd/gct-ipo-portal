r"""One-time FAST migration from local SQLite to Supabase PostgreSQL using execute_values."""
import os
import sys
import sqlite3
import psycopg2
from psycopg2 import sql
from psycopg2.extras import execute_values


def quote_ident(name):
    return '"' + name.replace('"', '""') + '"'


def main():
    if len(sys.argv) != 2:
        print(r"Usage: python migrate_sqlite_to_supabase.py path\to\ipo_management.db")
        return 2
    source_path = sys.argv[1]
    pg_url = os.environ.get("SUPABASE_DB_URL")
    if not pg_url:
        print("ERROR: Set SUPABASE_DB_URL environment variable first.")
        return 2
    if not os.path.isfile(source_path):
        print(f"ERROR: SQLite database file not found: {source_path}")
        return 2

    src = sqlite3.connect(source_path)
    src.row_factory = sqlite3.Row
    dst = psycopg2.connect(pg_url)
    migrated = 0
    skipped = []

    try:
        s_cur = src.cursor()
        d_cur = dst.cursor()
        s_cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name")
        tables = [r[0] for r in s_cur.fetchall()]

        for table in tables:
            s_cur.execute(f"PRAGMA table_info({quote_ident(table)})")
            source_cols = [r[1] for r in s_cur.fetchall()]

            d_cur.execute("""
                SELECT column_name FROM information_schema.columns
                WHERE table_schema = current_schema() AND table_name = %s
                ORDER BY ordinal_position
            """, (table,))
            target_cols = [r[0] for r in d_cur.fetchall()]
            if not target_cols:
                skipped.append(f"{table}: target table does not exist in Supabase")
                continue

            cols = [c for c in source_cols if c in target_cols]
            if not cols:
                skipped.append(f"{table}: no common columns")
                continue

            s_cur.execute(f"SELECT {', '.join(quote_ident(c) for c in cols)} FROM {quote_ident(table)}")
            rows = s_cur.fetchall()
            if not rows:
                print(f"{table}: 0 rows")
                continue

            all_records = [tuple(row[c] for c in cols) for row in rows]
            total_records = len(all_records)
            print(f"\nMigrating {table} ({total_records} rows)...")

            # Fast multi-row INSERT query template
            insert_query = f"""
                INSERT INTO {quote_ident(table)} ({', '.join(quote_ident(c) for c in cols)})
                VALUES %s
                ON CONFLICT DO NOTHING
            """

            # 2,000 records per batch via execute_values (extremely fast)
            batch_size = 2000
            for i in range(0, total_records, batch_size):
                batch = all_records[i:i + batch_size]
                execute_values(d_cur, insert_query, batch, page_size=batch_size)
                dst.commit()
                processed = min(i + batch_size, total_records)
                print(f"  --> Progress: {processed} / {total_records} rows ({int(processed/total_records*100)}%)")

            migrated += total_records
            print(f"{table}: Completed {total_records} rows.")

            # Update Serial Sequences
            if "id" in cols and table != "system_users":
                d_cur.execute("SELECT pg_get_serial_sequence(%s, 'id')", (table,))
                seq_res = d_cur.fetchone()
                seq = seq_res[0] if seq_res else None
                if seq:
                    d_cur.execute(sql.SQL("SELECT COALESCE(MAX({}), 0) FROM {}").format(
                        sql.Identifier("id"), sql.Identifier(table)
                    ))
                    max_id = d_cur.fetchone()[0] or 0
                    if max_id > 0:
                        d_cur.execute("SELECT setval(%s, %s, true)", (seq, max_id))
                        dst.commit()

        print(f"\n✅ Migration Successfully Finished!")
        return 0

    except Exception:
        try:
            if not dst.closed:
                dst.rollback()
        except Exception:
            pass
        raise
    finally:
        src.close()
        try:
            if not dst.closed:
                dst.close()
        except Exception:
            pass


if __name__ == "__main__":
    raise SystemExit(main())
