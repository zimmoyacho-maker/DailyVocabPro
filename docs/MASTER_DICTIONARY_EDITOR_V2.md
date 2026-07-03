# Master Dictionary Editor v2.0

PySide6 기반의 DailyVocabPro Master Dictionary 편집기입니다.

## Features

- Dark modern UI
- Excel open/save
- Word card with progress
- Search result list
- Row navigation
- Editable `example_ko`, `memo`, `level`, `tags`
- Quick tag buttons
- Save state indicator
- QA summary
- Auto backup before save
- Translation_Log update

## Install

```bash
pip install -r apps/requirements_dictionary_editor.txt
```

## Run

```bash
python apps/dictionary_editor_qt.py
```

## Recommended workflow

1. Put latest Excel file in `data/`.
2. Run the editor.
3. Open the Excel file.
4. Go to row `1002`.
5. Edit and press `저장 후 다음`.
6. Run QA every 50~100 rows.
7. Save and commit the updated file.

## Commit message

```text
feat(editor): add PySide6 master dictionary editor
```
