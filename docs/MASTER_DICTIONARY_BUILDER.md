# Master Dictionary Builder Guide

This tool helps build the DailyVocabPro Master Dictionary.

## Purpose

The builder reads an Excel vocabulary file and assists these columns:

- `example_ko`
- `memo`
- `level`
- `tags`
- `part_of_speech`
- `dictionary_def`
- `pronunciation`
- `source_url`
- `qa_flag`

## Recommended workflow

Run the tool in small batches first.

```bash
pip install -r tools/requirements_master_dictionary.txt
```

### Run with translation API

```bash
python tools/master_dictionary_builder.py ^
  --input data/words.xlsx ^
  --output data/words_batch003_partA.xlsx ^
  --start-row 1002 ^
  --end-row 1101 ^
  --batch-name Batch003-PartA ^
  --translate-url http://localhost:5000/translate
```

### Run without translation API

This fills dictionary info, memo, level, tags, and QA flags only.

```bash
python tools/master_dictionary_builder.py ^
  --input data/words.xlsx ^
  --output data/words_batch003_partA_no_translation.xlsx ^
  --start-row 1002 ^
  --end-row 1101 ^
  --batch-name Batch003-PartA ^
  --no-translate
```

## QA flags

The `qa_flag` column helps identify rows that need review.

Common flags:

- `missing_example_ko`
- `review_translation`
- `missing_memo`
- `missing_level`
- `missing_tags`
- `no_translation_api`

## Project Atlas batch plan

| Batch | Rows | Status |
|---|---:|---|
| Batch001 | 2~501 | Complete |
| Batch002 | 502~1001 | Complete |
| Batch003 | 1002~1501 | In Progress |
| Batch004 | 1502~2001 | Pending |
| Batch005 | 2002~2501 | Pending |
| Batch006 | 2502~2999 | Pending |

## Notes

Machine translation is only a draft. Final review is required for publication-quality data.
