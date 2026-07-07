# Editor v2.4 — Knowledge Integrated

## New

- Editor now uses `engine/knowledge_engine.py`
- `Assist All` uses Translation Memory + Phrase Dictionary
- Confidence and reasons are shown in the right panel
- Saved translations are learned into `knowledge/translation_memory.json`
- Phrase matches are reused automatically

## Run

```bash
python apps/dictionary_editor_qt.py
```

## Test Knowledge Engine

```bash
python apps/test_knowledge_engine.py
```

## Commit message

```text
feat(editor): integrate knowledge engine with editor assist
```
