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
- For DB export only: Node.js and `mysql2` available from the PMS service checkout

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
- offline position reference JSON, for example `configs/production_position_reference.json`

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
  --mapping "configs/production_position_reference.json" \
  --output "output/KPI Upload Template - All Positions.xlsx" \
  --report "output/KPI Upload Template - All Positions.report.csv"
```

## Offline Position Reference

The converter should not require DB access on every device. Export the production reference once from a machine with read-only DB access, commit or distribute the generated JSON, then run conversion offline with `--mapping`.

```bash
DB_USER="<read-only user>" DB_PASSWORD="<password>" \
node scripts/export_position_reference.mjs \
  --profile production \
  --output configs/production_position_reference.json
```

The generated reference includes:

- `rows`: Position Nomenclature mapping (`cluster_id` / PNID to PMID) enriched with group/company names and active flags.
- `position_master_rows`: Position Master rows enriched with active organization and company context.
- `organization_rows`: active/deleted-filtered organization/group reference.
- `company_rows`: active/deleted-filtered company reference.

## Run Converter Without Codex

From this repo:

```bash
cd /Users/alfredoteja/Documents/portaverse-kpi-bulk-upload-transformer
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install -r requirements.txt
```

Run a ZIP batch:

```bash
python3 scripts/kpi_bulk_transform.py \
  --source "/absolute/path/to/KAMUS KPI GROUP 2.zip" \
  --template "input/KPI Upload Template.xlsx" \
  --mapping "configs/production_position_reference.json" \
  --output-dir "output/group2_conversion_$(date +%Y%m%d_%H%M)"
```

Run one workbook:

```bash
python3 scripts/kpi_bulk_transform.py \
  --source "/absolute/path/to/source.xlsx" \
  --template "input/KPI Upload Template.xlsx" \
  --config "configs/pre_restructure_positions_rw_reviewed_20260609.json" \
  --mapping "configs/production_position_reference.json" \
  --output "output/single_conversion.xlsx" \
  --report "output/single_conversion.report.csv"
```

Build a recap workbook:

```bash
python3 scripts/build_conversion_recap.py \
  --output-dir "output/group2_conversion_YYYYMMDD_HHMM" \
  --config "configs/pre_restructure_positions_rw_reviewed_20260609.json" \
  --reference "configs/production_position_reference.json" \
  --output "output/group2_conversion_recap.xlsx" \
  --report-scope "Group 2"
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
- `position_nomenclature_id`: PNID/cluster id from the offline production reference
- `group_name`
- `directorate_name`
- `expected_impact_count`: expected shared Pelindo Impact count, currently `10`
- `drop_comment_values`: comment values that force OUTPUT/KAI drop

## Current Rules Implemented

- Source sheets are parsed as block-style layouts with downward inheritance for merged-looking fields.
- Placeholder values like `(blank)` are treated as missing.
- Positions are exported regardless of `Tipe Posisi` (`Struktural` and `Non-struktural` are both included).
- `Position Master ID` or `Position Nomenklatur ID` is required in generated output rows.
- If `Position Nomenklatur ID` is present, the output leaves `Position Master ID` blank so the backend importer expands PNID to PMID.
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
