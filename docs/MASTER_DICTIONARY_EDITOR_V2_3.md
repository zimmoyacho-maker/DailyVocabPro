# Master Dictionary Editor v2.3 Translation Assist

## New in v2.3

- `Example_KO 초안` button
- `Assist All` now also fills Example_KO when it is empty
- Offline pattern-based Korean draft generation
- No OpenAI API required
- User can freely edit the suggested Korean translation

## Important

The translation is a draft, not final publication quality.

Recommended workflow:

1. Open Excel.
2. Select Batch Range.
3. Click `Example_KO 초안` or `Assist All`.
4. Review and polish Korean translation.
5. Save and Next.

## Run

```bash
python apps/dictionary_editor_qt.py
```

## Commit message

```text
feat(editor): add offline example translation assist
```
