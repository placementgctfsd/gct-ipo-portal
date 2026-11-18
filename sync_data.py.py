"""Safe one-way IPO data sync: Supabase PostgreSQL -> existing local SQLite.

Reads the database connection URL from:
    .streamlit/secrets.toml  -> [postgres] url = "..."
or the SUPABASE_DB_URL environment variable.

Usage:
    python sync_data.py --test-connection   # no SQLite changes
    python sync_data.py                     # backup, then sync 4 tables

This script never creates or migrates SQLite tables, deletes local-only rows,
or sends local data to Supabase. Rows with the same primary key are updated
from the cloud. The four-table operation is one SQLite transaction.
"""

from __future__ import annotations

import argparse
from contextlib import closing
from datetime import date, datetime, time, timedelta
from decimal import Decimal
import json
import os
from pathlib import Path
import socket
import sqlite3
import sys
from urllib.parse import urlparse
import uuid

import psycopg2
from psycopg2 import sql

BASE_DIR = Path(__file__).resolve().parent
LOCAL_DB = BASE_DIR / "ipo_management.db"
TABLES = ("students", "guest_lectures", "industry_visits", "recruitment_drives")


def load_database_url() -> str:
    env_url = os.environ.get("SUPABASE_DB_URL", "").strip()
    if env_url:
        return env_url

    secrets_path = BASE_DIR / ".streamlit" / "secrets.toml"
    if not secrets_path.is_file():
        raise RuntimeError(
            "Missing .streamlit/secrets.toml. Add [postgres] url, "
            "or set SUPABASE_DB_URL."
        )

    try:
        import tomllib  # Python 3.11+
    except ModuleNotFoundError:
        try:
            import tomli as tomllib  # Python 3.10: pip install tomli
        except ModuleNotFoundError as exc:
            raise RuntimeError(
                "Python 3.10 requires 'pip install tomli' to read secrets.toml."
            ) from exc

    with secrets_path.open("rb") as f:
        config = tomllib.load(f)
    url = str(config.get("postgres", {}).get("url", "")).strip()
    if not url:
        raise RuntimeError("Missing [postgres] url in .streamlit/secrets.toml.")
    return url


def validate_url(url: str):
    info = urlparse(url)
    if info.scheme not in ("postgres", "postgresql") or not info.hostname:
        raise RuntimeError("Invalid PostgreSQL URL; copy it from Supabase > Connect.")
    if any(marker in url for marker in ("YOUR_PASSWORD", "[YOUR-PASSWORD]", "POOLER_HOST")):
        raise RuntimeError("Replace all placeholders in your PostgreSQL URL.")
    try:
        port = info.port or 5432
    except ValueError as exc:
        raise RuntimeError("Invalid port in PostgreSQL URL.") from exc
    return info.hostname, port


def sqlite_value(value):
    if isinstance(value, memoryview):
        return value.tobytes()
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, (datetime, date, time)):
        return value.isoformat()
    if isinstance(value, timedelta):
        return str(value)
    if isinstance(value, Decimal):
        return str(value)
    if isinstance(value, uuid.UUID):
        return str(value)
    if isinstance(value, (dict, list, tuple)):
        return json.dumps(value, ensure_ascii=False, default=str)
    return value


def quote_identifier(name: str) -> str:
    return '"' + name.replace('"', '""') + '"'


def backup_sqlite(source: sqlite3.Connection) -> Path:
    """Use SQLite's live backup API; abort synchronization if no backup works."""
    now = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    candidates = []
    if os.environ.get("LOCALAPPDATA"):
        candidates.append(Path(os.environ["LOCALAPPDATA"]) / "GCT_IPO" / "backups")
    candidates.append(BASE_DIR / "ipo_backups")
    candidates.append(Path.home() / "GCT_IPO_backups")

    failures = []
    for folder in candidates:
        target = folder / f"ipo_management_before_sync_{now}.db"
        try:
            folder.mkdir(parents=True, exist_ok=True)
            with closing(sqlite3.connect(str(target), timeout=15)) as dest:
                source.backup(dest)
            # Prove the backup is a valid SQLite database.
            with closing(sqlite3.connect(str(target))) as check:
                result = check.execute("PRAGMA quick_check").fetchone()
            if not result or result[0] != "ok":
                raise RuntimeError("backup integrity check did not pass")
            return target
        except (OSError, sqlite3.Error, RuntimeError) as exc:
            failures.append(f"{folder}: {type(exc).__name__}: {exc}")
            try:
                target.unlink(missing_ok=True)
            except OSError:
                pass

    raise RuntimeError(
        "Cannot create a verified SQLite backup; no records have been modified. "
        + " | ".join(failures)
    )


def sync_one_table(pg_conn, sqlite_conn: sqlite3.Connection, table: str) -> int:
    info = sqlite_conn.execute(
        f"PRAGMA table_info({quote_identifier(table)})"
    ).fetchall()
    if not info:
        raise RuntimeError(f"Local table '{table}' does not exist. No schema was changed.")

    local_columns = {row[1] for row in info}
    primary_keys = [row[1] for row in sorted(info, key=lambda r: r[5]) if row[5] > 0]
    if not primary_keys:
        raise RuntimeError(f"Local table '{table}' has no primary key; refusing unsafe merge.")

    with pg_conn.cursor() as cur:
        cur.execute(sql.SQL("SELECT * FROM {}.{}").format(
            sql.Identifier("public"), sql.Identifier(table)
        ))
        online_columns = [desc[0] for desc in cur.description]
        missing_local = [col for col in online_columns if col not in local_columns]
        if missing_local:
            raise RuntimeError(
                f"'{table}' has online columns not in SQLite: "
                + ", ".join(missing_local)
                + ". No automatic schema changes were made."
            )
        missing_keys = [key for key in primary_keys if key not in online_columns]
        if missing_keys:
            raise RuntimeError(
                f"'{table}' online source lacks primary key(s): " + ", ".join(missing_keys)
            )

        col_sql = ", ".join(quote_identifier(col) for col in online_columns)
        values_sql = ", ".join("?" for _ in online_columns)
        pk_sql = ", ".join(quote_identifier(pk) for pk in primary_keys)
        update_cols = [col for col in online_columns if col not in primary_keys]
        if update_cols:
            update_sql = ", ".join(
                f"{quote_identifier(col)} = excluded.{quote_identifier(col)}"
                for col in update_cols
            )
            conflict_sql = f"ON CONFLICT ({pk_sql}) DO UPDATE SET {update_sql}"
        else:
            conflict_sql = f"ON CONFLICT ({pk_sql}) DO NOTHING"

        statement = (
            f"INSERT INTO {quote_identifier(table)} ({col_sql}) VALUES ({values_sql}) "
            + conflict_sql
        )
        count = 0
        while True:
            batch = cur.fetchmany(500)
            if not batch:
                break
            sqlite_conn.executemany(
                statement,
                [tuple(sqlite_value(value) for value in row) for row in batch],
            )
            count += len(batch)
        return count


def main() -> int:
    parser = argparse.ArgumentParser(description="Safely sync cloud IPO tables to local SQLite")
    parser.add_argument("--test-connection", action="store_true", help="Test Supabase only; do not modify SQLite")
    args = parser.parse_args()

    try:
        url = load_database_url()
        hostname, port = validate_url(url)
    except (OSError, ValueError, RuntimeError) as exc:
        print(f"❌ Configuration problem: {exc}")
        return 1

    print(f"⏳ Checking Supabase DNS: {hostname}:{port}")
    try:
        socket.getaddrinfo(hostname, port, type=socket.SOCK_STREAM)
    except OSError as exc:
        print(f"❌ Cannot resolve this hostname: {exc}")
        print("   Verify the exact Pooler host in Supabase > Connect; test with nslookup.")
        return 1

    print("⏳ Connecting to Supabase Cloud...")
    try:
        pg_conn = psycopg2.connect(url, connect_timeout=15, sslmode="require")
    except psycopg2.Error as exc:
        # Scrub the full URL and password if present in a driver error.
        safe_error = str(exc).replace(url, "[REDACTED CONNECTION URL]")
        if urlparse(url).password:
            safe_error = safe_error.replace(urlparse(url).password, "[REDACTED]")
        print(f"❌ Supabase connection failed: {safe_error}")
        return 1

    with closing(pg_conn):
        if args.test_connection:
            with pg_conn.cursor() as cur:
                cur.execute("SELECT current_database(), current_user")
                database, username = cur.fetchone()
            print(f"✅ Supabase connection verified (database: {database}, user: {username}).")
            print("ℹ️ Test only: local database was not opened or changed.")
            return 0

        if not LOCAL_DB.is_file():
            print(f"❌ Missing local database: {LOCAL_DB}")
            print("   Copy your existing ipo_management.db alongside this script. Nothing was created.")
            return 1

        try:
            with closing(sqlite3.connect(str(LOCAL_DB), timeout=15)) as local:
                local.execute("PRAGMA busy_timeout = 15000")
                print("⏳ Creating a verified backup before changing any SQLite records...")
                backup_file = backup_sqlite(local)
                print(f"✅ SQLite backup: {backup_file}")
                pg_conn.set_session(readonly=True, isolation_level="REPEATABLE READ")

                try:
                    local.execute("BEGIN IMMEDIATE")
                    for table in TABLES:
                        print(f"🔄 Syncing {table}...")
                        count = sync_one_table(pg_conn, local, table)
                        print(f"   ✅ {count} cloud rows merged")
                    local.commit()
                except (psycopg2.Error, sqlite3.Error, RuntimeError, ValueError, TypeError) as exc:
                    local.rollback()
                    print(f"❌ Sync cancelled; local changes rolled back: {exc}")
                    print(f"✅ Original database backup is still available: {backup_file}")
                    return 1

                print("✨ Sync complete: 4 tables updated in local SQLite.")
                print("ℹ️ Local-only rows preserved; matching keys updated from cloud.")
                print("ℹ️ No cloud data was modified.")
                return 0
        except (sqlite3.Error, OSError, RuntimeError) as exc:
            print(f"❌ Local SQLite problem: {exc}")
            print("   Sync aborted without committing any records.")
            return 1


if __name__ == "__main__":
    sys.exit(main())
