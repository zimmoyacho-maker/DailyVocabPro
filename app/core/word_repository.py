from app.database.db import connect

def count_words() -> int:
    conn = connect()
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM words")
    n = cur.fetchone()[0]
    conn.close()
    return n

def fetch_new_words(limit: int):
    conn = connect()
    cur = conn.cursor()
    cur.execute("""
    SELECT w.id, w.word, w.meaning, w.example, w.example_ko, w.memo, w.level, w.tags
    FROM words w
    LEFT JOIN progress p ON p.word_id = w.id
    WHERE p.word_id IS NULL
    ORDER BY RANDOM()
    LIMIT ?
    """, (limit,))
    rows = cur.fetchall()
    conn.close()
    return rows

def fetch_review_words(limit: int, today: str):
    conn = connect()
    cur = conn.cursor()
    cur.execute("""
    SELECT w.id, w.word, w.meaning, w.example, w.example_ko, w.memo, w.level, w.tags
    FROM words w
    JOIN progress p ON p.word_id = w.id
    WHERE p.next_review IS NOT NULL AND p.next_review <= ?
    ORDER BY p.next_review ASC, p.score ASC
    LIMIT ?
    """, (today, limit))
    rows = cur.fetchall()
    conn.close()
    return rows

def search_words(query: str, limit: int = 200):
    q = f"%{query.strip()}%"
    conn = connect()
    cur = conn.cursor()
    cur.execute("""
    SELECT id, word, meaning, example, example_ko, memo, level, tags
    FROM words
    WHERE word LIKE ?
       OR meaning LIKE ?
       OR example LIKE ?
       OR example_ko LIKE ?
       OR memo LIKE ?
       OR level LIKE ?
       OR tags LIKE ?
    ORDER BY
      CASE
        WHEN word = ? THEN 0
        WHEN word LIKE ? THEN 1
        ELSE 2
      END,
      word
    LIMIT ?
    """, (q, q, q, q, q, q, q, query, f"{query}%", limit))
    rows = cur.fetchall()
    conn.close()
    return rows

def fetch_words_by_ids(ids):
    if not ids:
        return []
    placeholders = ",".join(["?"] * len(ids))
    conn = connect()
    cur = conn.cursor()
    cur.execute(f"""
    SELECT id, word, meaning, example, example_ko, memo, level, tags
    FROM words
    WHERE id IN ({placeholders})
    """, tuple(ids))
    rows = cur.fetchall()
    conn.close()
    by_id = {r[0]: r for r in rows}
    return [by_id[i] for i in ids if i in by_id]
