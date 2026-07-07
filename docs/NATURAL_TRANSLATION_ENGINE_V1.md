# Natural Translation Engine v1

## Goal

Reduce weak fallback drafts like:

```text
이 예문에서 'achieve'는 '성취하다'의 의미로 쓰였습니다...
```

## New files

```text
engine/knowledge_engine.py
knowledge/verb_patterns.json
knowledge/noun_phrases.json
apps/test_natural_translation_engine.py
docs/NATURAL_TRANSLATION_ENGINE_V1.md
```

## Test

```bash
python apps/test_natural_translation_engine.py
```

## Commit message

```text
feat(engine): improve natural example translation drafts
```
