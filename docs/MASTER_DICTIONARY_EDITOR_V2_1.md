# Master Dictionary Editor v2.1

## New in v2.1

- Left word list
- Batch range filter
- Show only incomplete rows
- Better QA summary scoped to the current batch
- Save-and-next follows the visible batch list

## Run

```bash
pip install -r apps/requirements_dictionary_editor.txt
python apps/dictionary_editor_qt.py
```

## Recommended workflow for Batch003-A

1. Open latest Excel file from `data/`.
2. Set Batch Range: `1002 ~ 1101`.
3. Click `적용`.
4. Turn on `미완료만 보기` if needed.
5. Edit rows using `저장 후 다음`.
6. Run `QA 검사`.
7. Save.

## Commit message

```text
feat(editor): add batch word list and QA filters
```
