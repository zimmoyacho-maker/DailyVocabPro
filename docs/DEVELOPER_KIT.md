# DailyVocabPro Developer Kit v0.1

This kit prepares the project for SQLite-based development.

## Included

```text
scripts/
  schema.sql
  import_excel_to_sqlite.py
  export_sqlite_to_excel.py
  qa_sqlite.py

docs/
  DATABASE_SCHEMA.md
  DEVELOPER_KIT.md

data/
  excel/
  sqlite/
  backups/
```

## Install

```bash
pip install openpyxl
```

## 1. Put Excel file here

```text
data/excel/words_v3_exampleko_502_1001.xlsx
```

## 2. Import Excel to SQLite

```bash
python scripts/import_excel_to_sqlite.py --input data/excel/words_v3_exampleko_502_1001.xlsx --db data/sqlite/dailyvocab.db
```

## 3. Run QA

```bash
python scripts/qa_sqlite.py --db data/sqlite/dailyvocab.db
```

## 4. Export SQLite back to Excel

```bash
python scripts/export_sqlite_to_excel.py --db data/sqlite/dailyvocab.db --output data/excel/words_export.xlsx
```

## Recommended commit message

```text
feat(data): add SQLite developer kit
```
