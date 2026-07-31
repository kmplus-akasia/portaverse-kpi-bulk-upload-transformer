# Baseline counts

Last reconciled inventory run, kept as a sanity check for a new download. A fresh package will differ; a wild difference is a signal to re-check the exclusion rules before trusting the numbers.

## Run 2026-07-29

Source folder: `outputs/kamus-ho-config-20260729/source/KAMUS KPI PELINDO GROUP 1 (HO) 5/`  
Inventory JSON: `configs/kamus_kpi_ho_visible_20260729.json`  
Original download (historical only): `/Users/alfredoteja/Downloads/KAMUS KPI PELINDO GROUP 1 (HO) 5`
Output: `configs/kamus_kpi_ho_visible_20260729.json`

| Measure | v2 | v1 pre-restructure |
| --- | --- | --- |
| Workbooks | 39 | 26 |
| Visible worksheets | 500 | 380 |
| Position worksheets | 435 | 337 |
| Visible non-position sheets | 65 | 40 |
| Title missing in sheet | 0 | 3 |
| Hidden worksheets excluded | 128 | 96 |

Excluded from the inventory:

- `Unit Kerja Pelindo per April 2026 - edit.xlsx` — organisation reference, not a kamus.
- Archive folders whose name contains `(Original)`.
