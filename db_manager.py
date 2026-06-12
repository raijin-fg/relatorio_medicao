import sqlite3
from datetime import datetime
from pathlib import Path

DB_PATH = Path(__file__).parent / "data" / "prd_data.db"


def get_snapshot_date() -> str:
    """Return current date as YYYY-MM-01 format."""
    now = datetime.now()
    return now.strftime("%Y-%m-01")


def init_db() -> None:
    """Create database tables if they don't exist."""
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)

    with sqlite3.connect(DB_PATH) as conn:
        cursor = conn.cursor()

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS netapp_volumes (
                snapshot_date TEXT NOT NULL,
                volume_name TEXT NOT NULL,
                total_gb REAL,
                used_gb REAL,
                free_gb REAL,
                PRIMARY KEY (snapshot_date, volume_name)
            )
        """)

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS netbackup_jobs (
                snapshot_date TEXT NOT NULL,
                job_id TEXT NOT NULL,
                job_policy TEXT,
                job_schedule TEXT,
                start_time TEXT,
                end_time TEXT,
                elapsed_time INTEGER,
                kilobytes INTEGER,
                kb_sec REAL,
                PRIMARY KEY (snapshot_date, job_id)
            )
        """)

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS netbackup_policies (
                snapshot_date TEXT NOT NULL,
                policy_name TEXT NOT NULL,
                backup_selections TEXT,
                PRIMARY KEY (snapshot_date, policy_name)
            )
        """)

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS netbackup_tapes (
                snapshot_date TEXT NOT NULL,
                media_id TEXT NOT NULL,
                media_type TEXT,
                robot TEXT,
                pool TEXT,
                kilobytes INTEGER,
                PRIMARY KEY (snapshot_date, media_id)
            )
        """)

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS zabbix_availability (
                snapshot_date TEXT NOT NULL,
                group_name TEXT NOT NULL,
                host TEXT NOT NULL,
                item TEXT NOT NULL,
                availability_pct REAL,
                PRIMARY KEY (snapshot_date, group_name, host, item)
            )
        """)

        conn.commit()


def insert_netapp_volumes(rows: list) -> None:
    """Insert NetApp volumes data. Rows: [(volume_name, total_gb, used_gb, free_gb), ...]"""
    if not rows:
        return

    snapshot_date = get_snapshot_date()

    with sqlite3.connect(DB_PATH) as conn:
        cursor = conn.cursor()
        for row in rows:
            cursor.execute("""
                INSERT OR REPLACE INTO netapp_volumes
                (snapshot_date, volume_name, total_gb, used_gb, free_gb)
                VALUES (?, ?, ?, ?, ?)
            """, (snapshot_date, row[0], row[1], row[2], row[3]))
        conn.commit()


def insert_netbackup_jobs(rows: list) -> None:
    """Insert NetBackup jobs. Each row is a dict with keys: job_id, job_policy, job_schedule,
    start_time, end_time, elapsed_time, kilobytes, kb_sec"""
    if not rows:
        return

    snapshot_date = get_snapshot_date()

    with sqlite3.connect(DB_PATH) as conn:
        cursor = conn.cursor()
        for row in rows:
            cursor.execute("""
                INSERT OR REPLACE INTO netbackup_jobs
                (snapshot_date, job_id, job_policy, job_schedule, start_time, end_time,
                 elapsed_time, kilobytes, kb_sec)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (snapshot_date, row["job_id"], row["job_policy"], row["job_schedule"],
                  row["start_time"], row["end_time"], row["elapsed_time"],
                  row["kilobytes"], row["kb_sec"]))
        conn.commit()


def insert_netbackup_policies(rows: list) -> None:
    """Insert NetBackup policies. Each row is a dict with keys: policy_name, backup_selections"""
    if not rows:
        return

    snapshot_date = get_snapshot_date()

    with sqlite3.connect(DB_PATH) as conn:
        cursor = conn.cursor()
        for row in rows:
            cursor.execute("""
                INSERT OR REPLACE INTO netbackup_policies
                (snapshot_date, policy_name, backup_selections)
                VALUES (?, ?, ?)
            """, (snapshot_date, row["policy_name"], row["backup_selections"]))
        conn.commit()


def insert_netbackup_tapes(rows: list) -> None:
    """Insert NetBackup tapes. Rows: [(media_id, media_type, robot, pool, kilobytes), ...]"""
    if not rows:
        return

    snapshot_date = get_snapshot_date()

    with sqlite3.connect(DB_PATH) as conn:
        cursor = conn.cursor()
        for row in rows:
            cursor.execute("""
                INSERT OR REPLACE INTO netbackup_tapes
                (snapshot_date, media_id, media_type, robot, pool, kilobytes)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (snapshot_date, row[0], row[1], row[2], row[3], row[4]))
        conn.commit()


def insert_zabbix_availability(rows: list) -> None:
    """Insert Zabbix availability data. Each row is a dict with keys:
    group_name, host, item, availability_pct"""
    if not rows:
        return

    snapshot_date = get_snapshot_date()

    with sqlite3.connect(DB_PATH) as conn:
        cursor = conn.cursor()
        for row in rows:
            cursor.execute("""
                INSERT OR REPLACE INTO zabbix_availability
                (snapshot_date, group_name, host, item, availability_pct)
                VALUES (?, ?, ?, ?, ?)
            """, (snapshot_date, row["group_name"], row["host"], row["item"],
                  row["availability_pct"]))
        conn.commit()
