"""
Log every HTTP call from persistent_session.py to SQLite, to check a long
run live. `python monitoring.py` prints a refreshing summary.
"""

import os
import sqlite3
import time
from contextlib import contextmanager

DB_PATH = os.environ.get("SPX_MONITOR_DB") or os.path.join(
    os.environ.get("SPX_OUTPUT_DIR") or "./discovery_output", "monitoring.sqlite3"
)

_SCHEMA = """
CREATE TABLE IF NOT EXISTS http_calls (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    timestamp REAL NOT NULL,
    method TEXT NOT NULL,
    url TEXT NOT NULL,
    status_code INTEGER,
    duration_seconds REAL,
    response_bytes INTEGER,
    error TEXT
);
CREATE INDEX IF NOT EXISTS idx_http_calls_timestamp ON http_calls(timestamp);
"""


@contextmanager
def _connect():
    os.makedirs(os.path.dirname(DB_PATH) or ".", exist_ok=True)
    conn = sqlite3.connect(DB_PATH, timeout=10)
    try:
        yield conn
    finally:
        conn.close()


def init_db():
    """Create the table, in WAL mode."""
    with _connect() as conn:
        conn.execute("PRAGMA journal_mode=WAL;")
        conn.executescript(_SCHEMA)
        conn.commit()


_db_ready = False


def record_call(
    method,
    url,
    status_code=None,
    duration_seconds=None,
    response_bytes=None,
    error=None,
):
    """Record one HTTP call. Never raises."""
    global _db_ready
    try:
        if not _db_ready:
            init_db()
            _db_ready = True
        with _connect() as conn:
            conn.execute(
                "INSERT INTO http_calls "
                "(timestamp, method, url, status_code, duration_seconds, response_bytes, error) "
                "VALUES (?, ?, ?, ?, ?, ?, ?)",
                (
                    time.time(),
                    method,
                    url,
                    status_code,
                    duration_seconds,
                    response_bytes,
                    error,
                ),
            )
            conn.commit()
    except Exception:
        pass


def query_recent(n=10):
    """n most recent calls, newest first."""
    with _connect() as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            "SELECT * FROM http_calls ORDER BY timestamp DESC LIMIT ?", (n,)
        ).fetchall()
        return [dict(r) for r in rows]


def query_summary(last_seconds=None):
    """Calls, errors, bytes, average duration and time since last call,
    optionally over the last `last_seconds` only."""
    with _connect() as conn:
        conn.row_factory = sqlite3.Row
        where = ""
        params = ()
        if last_seconds:
            where = "WHERE timestamp >= ?"
            params = (time.time() - last_seconds,)

        row = conn.execute(
            f"""
            SELECT
                COUNT(*) AS total_calls,
                SUM(CASE WHEN error IS NOT NULL OR status_code >= 400 THEN 1 ELSE 0 END) AS error_count,
                SUM(COALESCE(response_bytes, 0)) AS total_bytes,
                AVG(duration_seconds) AS avg_duration_seconds,
                MAX(timestamp) AS last_call_timestamp
            FROM http_calls {where}
            """,
            params,
        ).fetchone()
        summary = dict(row)
        if summary["last_call_timestamp"]:
            summary["seconds_since_last_call"] = (
                time.time() - summary["last_call_timestamp"]
            )
        else:
            summary["seconds_since_last_call"] = None
        return summary


def _format_bytes(n):
    if not n:
        return "0 B"
    for unit in ["B", "KB", "MB", "GB"]:
        if n < 1024:
            return f"{n:.1f} {unit}"
        n /= 1024
    return f"{n:.1f} TB"


def _cli():
    """Print a summary every 5 s."""
    init_db()
    print(f"Monitoring database: {DB_PATH}")
    print("Ctrl+C to stop.\n")
    try:
        while True:
            summary = query_summary(last_seconds=60)
            since_last = summary["seconds_since_last_call"]
            status = (
                "ALIVE"
                if since_last is not None and since_last < 30
                else "NO RECENT ACTIVITY"
            )

            print(
                f"[{status}] last minute: {summary['total_calls'] or 0} calls, "
                f"{summary['error_count'] or 0} errors, "
                f"{_format_bytes(summary['total_bytes'])} transferred, "
                f"avg {summary['avg_duration_seconds'] or 0:.2f}s/call, "
                f"last call {since_last:.0f}s ago"
                if since_last is not None
                else f"[{status}] no calls recorded yet"
            )
            time.sleep(5)
    except KeyboardInterrupt:
        print("\nStopped.")


if __name__ == "__main__":
    _cli()
