"""
One-time migration: SQLite tracker.db -> PostgreSQL configured by DATABASE_URL.

Run from the project folder:
  Windows PowerShell:
    $env:DATABASE_URL="PASTE_RENDER_INTERNAL_DATABASE_URL_HERE"
    python migrate_sqlite_to_postgres.py

  macOS/Linux:
    export DATABASE_URL="PASTE_RENDER_INTERNAL_DATABASE_URL_HERE"
    python migrate_sqlite_to_postgres.py

Keep a backup of tracker.db. This script refuses to overwrite an already-populated
PostgreSQL database. It permits only the default seeded admin account created by
app startup, then replaces that seed with the source database's actual records.
"""
import os
import sys
import sqlite3
from datetime import date, datetime
from sqlalchemy import text
from pathlib import Path

if not os.environ.get("DATABASE_URL"):
    raise SystemExit("ERROR: Set DATABASE_URL to your PostgreSQL connection URL first.")

SOURCE_DB = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).with_name("tracker.db")
if not SOURCE_DB.is_file():
    raise SystemExit(f"ERROR: SQLite database not found: {SOURCE_DB}")

# Importing app configures SQLAlchemy from DATABASE_URL and creates the schema.
from app import app, db, User, DailyEntry, EntryHistory, Roster, Target

TABLES = ("user", "daily_entry", "entry_history", "roster", "target")

def rows(conn, table):
    return [dict(r) for r in conn.execute(f'SELECT * FROM "{table}"').fetchall()]

def parse_date(value):
    return date.fromisoformat(value) if isinstance(value, str) else value

def parse_datetime(value):
    if isinstance(value, str):
        return datetime.fromisoformat(value)
    return value

with sqlite3.connect(f"file:{SOURCE_DB.resolve().as_posix()}?mode=ro", uri=True) as source:
    source.row_factory = sqlite3.Row
    source_rows = {table: rows(source, table) for table in TABLES}

source_counts = {table: len(source_rows[table]) for table in TABLES}
if not source_rows["user"]:
    raise SystemExit("ERROR: Source SQLite database contains no users; nothing migrated.")

with app.app_context():
    # Create missing tables if needed. This does not alter existing data.
    db.create_all()

    destination_counts = {
        "user": User.query.count(),
        "daily_entry": DailyEntry.query.count(),
        "entry_history": EntryHistory.query.count(),
        "roster": Roster.query.count(),
        "target": Target.query.count(),
    }

    # A fresh app import creates a single default admin. Allow only that seed.
    allowed_seed = (
        destination_counts == {
            "user": 1, "daily_entry": 0, "entry_history": 0, "roster": 0, "target": 0
        }
        and User.query.filter_by(username="admin", role="admin").count() == 1
    )
    completely_empty = all(value == 0 for value in destination_counts.values())

    if not (allowed_seed or completely_empty):
        raise SystemExit(
            "STOPPED: PostgreSQL is not empty (or contains data beyond the default "
            f"seed admin). Current counts: {destination_counts}. No records were changed. "
            "Do not rerun or delete data; inspect the destination first."
        )

    try:
        if allowed_seed:
            seed_admin = User.query.filter_by(username="admin", role="admin").one()
            db.session.delete(seed_admin)
            db.session.flush()

        # Insert in foreign-key dependency order and preserve original primary keys.
        db.session.add_all([User(**r) for r in source_rows["user"]])
        db.session.flush()

        db.session.add_all([
            Target(**r) for r in source_rows["target"]
        ])
        db.session.add_all([
            DailyEntry(
                id=r["id"], agent_id=r["agent_id"],
                entry_date=parse_date(r["entry_date"]),
                total_calls=r["total_calls"], approvals=r["approvals"], errors=r["errors"]
            ) for r in source_rows["daily_entry"]
        ])
        db.session.add_all([
            Roster(
                id=r["id"], agent_id=r["agent_id"], off_date=parse_date(r["off_date"])
            ) for r in source_rows["roster"]
        ])
        db.session.flush()

        db.session.add_all([
            EntryHistory(
                id=r["id"], entry_id=r["entry_id"], changed_by=r["changed_by"],
                changed_at=parse_datetime(r["changed_at"]),
                old_values=r["old_values"], new_values=r["new_values"]
            ) for r in source_rows["entry_history"]
        ])
        db.session.commit()

        # Align PostgreSQL auto-increment sequences with the preserved IDs.
        for table in TABLES:
            quoted = f'"{table}"'
            db.session.execute(text(
                f"SELECT setval(pg_get_serial_sequence('{table}', 'id'), "
                f"COALESCE((SELECT MAX(id) FROM {quoted}), 1), "
                f"EXISTS(SELECT 1 FROM {quoted}))"
            ))
        db.session.commit()

    except Exception:
        db.session.rollback()
        raise

    final_counts = {
        "user": User.query.count(),
        "daily_entry": DailyEntry.query.count(),
        "entry_history": EntryHistory.query.count(),
        "roster": Roster.query.count(),
        "target": Target.query.count(),
    }

    if final_counts != source_counts:
        raise SystemExit(
            f"WARNING: Migration ran but count verification differs. "
            f"Source={source_counts}; PostgreSQL={final_counts}. "
            "Do not delete the SQLite backup."
        )

print("Migration completed and row counts verified.")
for table in TABLES:
    print(f"  {table}: {source_counts[table]} rows")
print("Your SQLite file was read-only and has not been changed.")
