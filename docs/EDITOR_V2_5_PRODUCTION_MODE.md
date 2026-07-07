# Editor v2.5 — Production Mode

## New

- Auto Assist option
- Moving to a new word automatically fills empty Example_KO, Memo, Tags, Level
- `Ctrl+Enter`: Save and Next
- `Ctrl+Space`: Assist All
- `Ctrl+Q`: QA Check
- `Ctrl+B`: Save current translation/memo to Golden Collection
- Golden button in UI

## Speed workflow

1. Open Excel.
2. Set Batch Range.
3. Keep `Auto Assist` checked.
4. Review Example_KO.
5. Press `Ctrl+Enter`.
6. Next word is opened and automatically assisted.

## Run

```bash
python apps/dictionary_editor_qt.py
```

## Commit message

```text
feat(editor): add production mode auto assist workflow
```
