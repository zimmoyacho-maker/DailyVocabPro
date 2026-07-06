# DailyVocabPro Database Schema

This document defines the first SQLite database design for DailyVocabPro.

## Goal

Move from Excel-only editing to a stable SQLite-based data layer.

```text
Excel
  ↓ import
SQLite
  ↓
Dictionary Editor
  ↓
Learning App
```

## Tables

- `words`: main vocabulary entry
- `examples`: English/Korean examples
- `notes`: memo, grammar, collocations, synonyms, antonyms, memory tips
- `word_tags`: multiple tags per word
- `qa_flags`: QA results

## Why SQLite?

- faster search
- better data consistency
- easier statistics
- easier learning app integration
- safer than editing only Excel
