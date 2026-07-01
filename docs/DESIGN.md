# Design Notes

## Learning flow

```text
Word
↓
Example
↓
Translation
↓
Meaning
↓
Self rating
↓
SRS scheduling
```

## SRS results

| Result | Effect |
|---|---|
| know | increases score and review interval |
| maybe | small score gain, short review interval |
| dont | score decrease, review tomorrow |

## Database tables

### words

Stores vocabulary data imported from Excel.

### progress

Stores SRS state.

### study_events

Stores daily learning statistics.
