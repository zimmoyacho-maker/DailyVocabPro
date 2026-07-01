from datetime import date, timedelta, datetime
from app.database.db import connect

def apply_rating(word_id: int, result: str, mode: str = "study"):
    today = date.today()
    now = datetime.now().isoformat(timespec="seconds")

    conn = connect()
    cur = conn.cursor()

    cur.execute("""
    SELECT score, interval_days, review_count
    FROM progress
    WHERE word_id=?
    """, (word_id,))
    row = cur.fetchone()

    if row:
        score, interval_days, review_count = row
    else:
        score, interval_days, review_count = 0, 1, 0

    if result == "know":
        score += 3
        if review_count == 0:
            interval_days = 3
        else:
            interval_days = min(max(interval_days * 2, 3), 120)
    elif result == "maybe":
        score += 1
        interval_days = max(1, min(interval_days, 3))
    else:
        score -= 2
        interval_days = 1

    review_count += 1
    next_review = (today + timedelta(days=interval_days)).isoformat()

    cur.execute("""
    INSERT INTO progress(
        word_id, score, interval_days, review_count,
        next_review, last_review, last_result
    )
    VALUES(?,?,?,?,?,?,?)
    ON CONFLICT(word_id) DO UPDATE SET
        score=excluded.score,
        interval_days=excluded.interval_days,
        review_count=excluded.review_count,
        next_review=excluded.next_review,
        last_review=excluded.last_review,
        last_result=excluded.last_result
    """, (
        word_id,
        score,
        interval_days,
        review_count,
        next_review,
        today.isoformat(),
        result,
    ))

    cur.execute("""
    INSERT INTO study_events(day, word_id, mode, result, created_at)
    VALUES(?,?,?,?,?)
    """, (today.isoformat(), word_id, mode, result, now))

    conn.commit()
    conn.close()

def due_review_count(today: str) -> int:
    conn = connect()
    cur = conn.cursor()
    cur.execute("""
    SELECT COUNT(*)
    FROM progress
    WHERE next_review IS NOT NULL AND next_review <= ?
    """, (today,))
    n = cur.fetchone()[0]
    conn.close()
    return n
