# Changelog

## v2.8.0

### Added
- Refactored project structure for GitHub development.
- Added separated modules:
  - `app/core/srs.py`
  - `app/core/statistics.py`
  - `app/database/db.py`
  - `app/database/excel_importer.py`
  - `app/services/speech.py`
  - `app/services/ai_tutor.py`
  - `app/ui/main_window.py`
  - `app/ui/study_window.py`
  - `app/ui/search_panel.py`
- Preserved SRS, search, statistics, TTS, and AI Tutor prototype features.

### Changed
- App entry point changed to:

```bash
python -m app.main
```

### Notes
- This is the first GitHub-ready refactor version.
