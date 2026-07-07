# Dictionary Editor v2.6 All-in-One

## Included

- Excel open/save
- Batch range
- Incomplete-only filter
- QA check
- Auto Assist
- Save and Next
- Confidence display
- Suggestion reasons
- Golden Collection save
- Translation Memory
- Phrase Dictionary
- Golden Collection
- Noun Phrase Dictionary
- Verb Pattern Dictionary
- Natural Translation Engine v1

## Keyboard shortcuts

- `Ctrl+S`: Save
- `Ctrl+Enter`: Save and Next
- `Ctrl+Space`: Assist All
- `Ctrl+Q`: QA Check
- `Ctrl+B`: Save Golden

## Run

```bash
python apps/dictionary_editor_qt.py
```

## Engine test

```bash
python apps/test_natural_translation_engine.py
```

## Copy into GitHub project

Copy these folders into your existing `DailyVocabPro` repository:

```text
apps/
engine/
knowledge/
docs/
```

## Commit

```bash
git add .
git commit -m "feat(editor): integrate all dictionary editor productivity features"
git push
```
