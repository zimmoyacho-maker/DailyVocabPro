from datetime import date
from app.database.db import connect

def today_stats():
    today = date.today().isoformat()
    conn = connect()
    cur = conn.cursor()

    cur.execute("SELECT COUNT(*) FROM study_events WHERE day=?", (today,))
    total = cur.fetchone()[0]

    cur.execute("SELECT COUNT(*) FROM study_events WHERE day=? AND result='know'", (today,))
    correct = cur.fetchone()[0]

    accuracy = round((correct / total) * 100) if total else 0
    stars = max(1, min(5, round(accuracy / 20))) if total else 0

    conn.close()
    return {
        "day": today,
        "total": total,
        "correct": correct,
        "accuracy": accuracy,
        "stars": "★" * stars + "☆" * (5 - stars) if total else "☆☆☆☆☆",
    }
