-- DailyVocabPro SQLite Schema v0.1
PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS words (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    word TEXT NOT NULL UNIQUE,
    meaning TEXT,
    part_of_speech TEXT,
    pronunciation TEXT,
    level TEXT,
    frequency INTEGER,
    source TEXT,
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS examples (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    word_id INTEGER NOT NULL,
    example_en TEXT NOT NULL,
    example_ko TEXT,
    is_primary INTEGER NOT NULL DEFAULT 1,
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY(word_id) REFERENCES words(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS notes (
    word_id INTEGER PRIMARY KEY,
    memo TEXT,
    grammar TEXT,
    collocations TEXT,
    synonyms TEXT,
    antonyms TEXT,
    confusion TEXT,
    memory_tip TEXT,
    ai_reviewed INTEGER NOT NULL DEFAULT 0,
    human_reviewed INTEGER NOT NULL DEFAULT 0,
    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY(word_id) REFERENCES words(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS word_tags (
    word_id INTEGER NOT NULL,
    tag TEXT NOT NULL,
    PRIMARY KEY(word_id, tag),
    FOREIGN KEY(word_id) REFERENCES words(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS qa_flags (
    word_id INTEGER PRIMARY KEY,
    flags TEXT,
    score INTEGER NOT NULL DEFAULT 0,
    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY(word_id) REFERENCES words(id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_words_word ON words(word);
CREATE INDEX IF NOT EXISTS idx_words_level ON words(level);
CREATE INDEX IF NOT EXISTS idx_examples_word_id ON examples(word_id);
CREATE INDEX IF NOT EXISTS idx_word_tags_tag ON word_tags(tag);
