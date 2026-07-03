# Master Dictionary Editor

`apps/dictionary_editor.py` is a GUI editor for the DailyVocabPro Master Dictionary.

## Features

- Open an Excel dictionary file
- Navigate rows
- Search words and examples
- Edit:
  - `example_ko`
  - `memo`
  - `level`
  - `tags`
- Quick tag buttons
- QA scan
- Auto backup before save
- Translation_Log update

## Install

```bash
pip install openpyxl
```

## Run

From the project root:

```bash
python apps/dictionary_editor.py
```

## Recommended workflow

1. Open the latest Excel file from `data/`.
2. Go to row `1002`.
3. Edit `example_ko`, `memo`, `level`, `tags`.
4. Click `저장 후 다음`.
5. Run QA after every 50~100 rows.
6. Save file.
7. Commit the updated file or export a batch result.

## Keyboard shortcuts

| Shortcut | Action |
|---|---|
| Ctrl+O | Open Excel |
| Ctrl+S | Save |
| Alt+Left | Previous row |
| Alt+Right | Next row |
| F5 | QA scan |

## Notes

The editor creates a backup in a `backups/` folder before saving.
