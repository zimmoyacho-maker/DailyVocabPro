#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""QA check for DailyVocab SQLite DB."""

from __future__ import annotations

import argparse
import sqlite3
from pathlib import Path


def one(cur, sql: str) -> int:
    cur.execute(sql)
    return cur.fetchone()[0]


def qa_check(db_path: Path) -> None:
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()

    checks = {
        "total_words": "SELECT COUNT(*) FROM words",
        "missing_meaning": "SELECT COUNT(*) FROM words WHERE meaning IS NULL OR TRIM(meaning)=''",
        "missing_level": "SELECT COUNT(*) FROM words WHERE level IS NULL OR TRIM(level)=''",
        "missing_example_en": """
            SELECT COUNT(*) FROM words w
            LEFT JOIN examples e ON e.word_id=w.id AND e.is_primary=1
            WHERE e.example_en IS NULL OR TRIM(e.example_en)=''
        """,
        "missing_example_ko": """
            SELECT COUNT(*) FROM words w
            LEFT JOIN examples e ON e.word_id=w.id AND e.is_primary=1
            WHERE e.example_ko IS NULL OR TRIM(e.example_ko)=''
        """,
        "missing_memo": """
            SELECT COUNT(*) FROM words w
            LEFT JOIN notes n ON n.word_id=w.id
            WHERE n.memo IS NULL OR TRIM(n.memo)=''
        """,
        "missing_tags": """
            SELECT COUNT(*) FROM words w
            LEFT JOIN word_tags t ON t.word_id=w.id
            WHERE t.tag IS NULL
        """,
    }

    print("DailyVocabPro SQLite QA")
    print("-----------------------")
    results = {}
    for name, sql in checks.items():
        value = one(cur, sql)
        results[name] = value
        print(f"{name}: {value}")

    total = results.get("total_words", 0)
    incomplete = (
        results.get("missing_example_ko", 0)
        + results.get("missing_memo", 0)
        + results.get("missing_level", 0)
        + results.get("missing_tags", 0)
    )
    if total:
        score = max(0, int(100 - (incomplete / (total * 4) * 100)))
        print(f"quality_score: {score}%")

    conn.close()


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--db", required=True)
    args = parser.parse_args()
    qa_check(Path(args.db))


if __name__ == "__main__":
    main()
