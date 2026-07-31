# Conversion commands

Run from the repo root with the project venv active.

## ZIP or folder batch, Head Office

```bash
python3 scripts/kpi_bulk_transform.py \
  --source "/absolute/path/to/KAMUS KPI GROUP 2.zip" \
  --template "input/KPI Upload Template.xlsx" \
  --mapping "configs/production_position_reference.json" \
  --write-discovered-config "output/<run>/discovered.config.json" \
  --output-dir "output/<run>"
```

## Batch spanning affiliates or non-HO companies

`--target-company-id ""` widens the lookup to every company. Head Office runs keep the default so a title shared across companies cannot resolve to the wrong one.

```bash
python3 scripts/kpi_bulk_transform.py \
  --source "/absolute/path/to/KAMUS KPI GROUP 3.zip" \
  --template "input/KPI Upload Template.xlsx" \
  --mapping "configs/production_position_reference.json" \
  --target-company-id "" \
  --output-dir "output/<run>"
```

## Single workbook

Resolve Head Office Kamus KPI paths through `scripts/kamus_source.py` before calling the converter.

```bash
python3 scripts/kpi_bulk_transform.py \
  --source "$(python3 - <<'PY'
from pathlib import Path
import sys
sys.path.insert(0, "scripts")
from kamus_source import resolve_kamus_source_root, resolve_source_workbook, load_inventory_config
ctx = resolve_kamus_source_root()
inv = load_inventory_config(ctx.inventory_config)
print(resolve_source_workbook(ctx.source_root, "Group Pengelolaan SDM/DIREKTORAT SDM & UMUM - Group Pengelolaan SDM.xlsx", inv))
PY
)" \
  --template "input/KPI Upload Template.xlsx" \
  --config "configs/<approved>.json" \
  --mapping "configs/production_position_reference.json" \
  --output "output/<run>/single_conversion.xlsx" \
  --report "output/<run>/single_conversion.report.csv"
```

For a whole approved batch against the canonical HO root, prefer resolving each config row with `resolve_source_workbook(...)` rather than hardcoding `SOURCE_ROOT` in a run script.

```bash
python3 scripts/kpi_bulk_transform.py \
  --source "/absolute/path/to/source.xlsx" \
  --template "input/KPI Upload Template.xlsx" \
  --config "configs/<approved>.json" \
  --mapping "configs/production_position_reference.json" \
  --output "output/<run>/single_conversion.xlsx" \
  --report "output/<run>/single_conversion.report.csv"
```

`--only-sheet "<sheet name>"` narrows any of the above to one worksheet.

## Recap workbook

```bash
python3 scripts/build_conversion_recap.py \
  --output-dir "output/<run>" \
  --config "configs/<approved>.json" \
  --reference "configs/production_position_reference.json" \
  --output "output/<run>/recap.xlsx" \
  --report-scope "Group 2"
```

## Report CSV columns

`severity`, `sheet_name`, `source_row`, `record_type`, `title`, `message`. Filter on `severity = error` before packaging.

## Normalisation the converter already applies

Reading these keeps a post-conversion "fix" from re-doing work the converter did:

- `Triwulan` and `Triwulanan` become `TRIWULANAN`; accepted periods are `BULANAN`, `TRIWULANAN`, `TAHUNAN`, `SEMESTER`, `MONTHLY`, `QUARTERLY`, `WEEKLY`.
- Placeholders such as `(blank)` count as missing.
- Shared KPI Impact fields backfill by title across parsed sheets when one sheet holds a placeholder and another holds the real value.
- OUTPUT and KAI rows drop when their required weight is blank, when `Komentar`, `Comment`, `Status`, or `Alignment` matches `drop_comment_values`, or when the title is numeric-only, which marks a subtotal row rather than a KPI.
