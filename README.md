# KPI Bulk Transformer

Portable Python repo for transforming KPI design workbooks into the official bulk upload template.

## Repo Structure

```text
kpi-bulk-transformer/
  configs/
    all_positions.json
    sample.json
  input/
  output/
  scripts/
    kpi_bulk_transform.py
  .gitignore
  requirements.txt
  README.md
```

## Requirements

- Python 3.11 or newer
- `openpyxl`

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install -r requirements.txt
```

## Inputs

This repo now includes a ready-to-run `input/` folder. It is intended to hold:

- source KPI workbook, for example `Bu Desi - Group Pengelolaan SDM (done Konfirmasi KPI).xlsx`
- official upload template, for example `KPI Upload Template.xlsx`
- folder containing `Master Posisi` exports

Current local layout:

```text
input/
  Bu Desi - Group Pengelolaan SDM (done Konfirmasi KPI).xlsx
  KPI Upload Template.xlsx
  data_master_posisi_31-03-2026_12_56_12/
```

## Run All Configured Positions

```bash
python3 scripts/kpi_bulk_transform.py \
  --source "input/Bu Desi - Group Pengelolaan SDM (done Konfirmasi KPI).xlsx" \
  --template "input/KPI Upload Template.xlsx" \
  --positions-dir "input/data_master_posisi_31-03-2026_12_56_12" \
  --config "configs/all_positions.json" \
  --output "output/KPI Upload Template - All Positions.xlsx" \
  --report "output/KPI Upload Template - All Positions.report.csv"
```

## Run One Sheet Only

```bash
python3 scripts/kpi_bulk_transform.py \
  --source "input/Bu Desi - Group Pengelolaan SDM (done Konfirmasi KPI).xlsx" \
  --template "input/KPI Upload Template.xlsx" \
  --positions-dir "input/data_master_posisi_31-03-2026_12_56_12" \
  --config "configs/all_positions.json" \
  --only-sheet "DH Manajemen Talenta" \
  --output "output/DH Manajemen Talenta.xlsx" \
  --report "output/DH Manajemen Talenta.report.csv"
```

## Config

Each entry in `configs/*.json` maps one worksheet to one upload target position.

Important fields:

- `sheet_name`: source worksheet name
- `position_name`: output position title
- `position_master_id`: expected Position Master ID
- `position_lookup_names`: fallback names when resolving against `Master Posisi`
- `group_name`
- `directorate_name`
- `expected_impact_count`: expected shared Pelindo Impact count, currently `10`
- `drop_comment_values`: comment values that force OUTPUT/KAI drop

## Current Rules Implemented

- Source sheets are parsed as block-style layouts with downward inheritance for merged-looking fields.
- Placeholder values like `(blank)` are treated as missing.
- Positions are exported regardless of `Tipe Posisi` (`Struktural` and `Non-struktural` are both included).
- `Position Master ID` is always required in generated output rows.
- `Position Master Variant ID` is optional and may be blank.
- `System KPI ID` is included in output format and currently left blank.
- `Triwulan` and `Triwulanan` are normalized to `TRIWULANAN`.
- Allowed uploader periods supported by normalization:
  - `BULANAN`
  - `TRIWULANAN`
  - `TAHUNAN`
  - `SEMESTER`
  - `MONTHLY`
  - `QUARTERLY`
  - `WEEKLY`
- Shared KPI Impact fields can be backfilled by title across parsed sheets when one sheet contains placeholders and another contains the valid value.
- OUTPUT/KAI rows are removed when:
  - their required weight is blank
  - their comment is in `drop_comment_values`

## Outputs

The script writes:

- one upload-ready `.xlsx`
- one validation `.csv`

The CSV report contains:

- `severity`
- `sheet_name`
- `source_row`
- `record_type`
- `title`
- `message`

## Current Limitation

- `Position Master Variant ID` and `System KPI ID` are currently not populated by the transformer and remain blank.

## Git Init

To make this a standalone repo:

```bash
git init
git add .
git commit -m "Initial KPI bulk transformer"
```
