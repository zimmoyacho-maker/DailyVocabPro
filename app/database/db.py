import sqlite3
from app.core.paths import db_path

def connect():
    return sqlite3.connect(db_path())

def initialize_database():
    conn = connect()
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS words (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        word TEXT NOT NULL UNIQUE,
        meaning TEXT,
        example TEXT,
        example_ko TEXT,
        memo TEXT,
        level TEXT,
        tags TEXT
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS settings (
        key TEXT PRIMARY KEY,
        value TEXT NOT NULL
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS progress (
        word_id INTEGER PRIMARY KEY,
        score INTEGER NOT NULL DEFAULT 0,
        interval_days INTEGER NOT NULL DEFAULT 1,
        review_count INTEGER NOT NULL DEFAULT 0,
        next_review TEXT,
        last_review TEXT,
        last_result TEXT,
        FOREIGN KEY(word_id) REFERENCES words(id)
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS study_events (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        day TEXT NOT NULL,
        word_id INTEGER NOT NULL,
        mode TEXT NOT NULL,
        result TEXT NOT NULL,
        created_at TEXT NOT NULL,
        FOREIGN KEY(word_id) REFERENCES words(id)
    )
    """)

    defaults = {
        "daily_new": "10",
        "daily_review": "5",
        "daily_time": "21:00",
        "theme": "light",
    }
    for key, value in defaults.items():
        cur.execute(
            "INSERT OR IGNORE INTO settings(key, value) VALUES(?, ?)",
            (key, value),
        )

    conn.commit()
    conn.close()

def get_setting(key: str, default: str = "") -> str:
    conn = connect()
    cur = conn.cursor()
    cur.execute("SELECT value FROM settings WHERE key=?", (key,))
    row = cur.fetchone()
    conn.close()
    return row[0] if row else default

def set_setting(key: str, value: str):
    conn = connect()
    cur = conn.cursor()
    cur.execute("INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)", (key, value))
    conn.commit()
    conn.close()
