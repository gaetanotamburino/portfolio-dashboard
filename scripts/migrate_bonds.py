from pathlib import Path
import sqlite3

DB_PATH = Path(__file__).parent.parent / "data" / "portfolio.db"


def migrate(conn):
    conn.execute("""
        CREATE TABLE IF NOT EXISTS bonds (
            isin     TEXT PRIMARY KEY,
            coupon   REAL,
            freq     INTEGER,
            maturity TEXT
        )
    """)
    conn.execute(
        "INSERT OR IGNORE INTO bonds VALUES (?, ?, ?, ?)",
        ("IT0004286966", 5.0, 2, "2039-08-01"),
    )
    conn.commit()


if __name__ == "__main__":
    with sqlite3.connect(DB_PATH) as conn:
        migrate(conn)
    print("bonds table ready.")
