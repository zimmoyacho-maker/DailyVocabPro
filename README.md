# DailyVocab Pro

DailyVocab Pro is a personal English vocabulary learning app built with Python, Tkinter, SQLite, and Excel.

## Current milestone

**V2.8 — Project Refactor**

This version moves the previous prototype features into a maintainable GitHub project structure.

## Features

- Excel vocabulary import
- SQLite local database
- Daily new word study
- SRS-based review scheduling
- Study card UI
- Example → Korean translation → meaning learning flow
- Search across word, meaning, example, translation, memo, level, tags
- Study statistics
- Offline TTS pronunciation with `pyttsx3`
- Offline AI Tutor-style explanation panel
- GitHub-ready project structure

## Folder structure

```text
DailyVocabPro/
├── app/
│   ├── main.py
│   ├── core/
│   ├── database/
│   ├── services/
│   ├── ui/
│   └── utils/
├── data/
├── docs/
├── tests/
├── assets/
├── releases/
├── requirements.txt
└── README.md
```

## Run

```bash
pip install -r requirements.txt
python -m app.main
```

Place your `words.xlsx` file in the `data/` folder.

Supported Excel columns:

```text
word, meaning, example, example_ko, memo, level, tags
```

The first three columns are required. The rest are optional.

## Version

Current: **v2.8.0**
