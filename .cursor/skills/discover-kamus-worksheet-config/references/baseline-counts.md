# Baseline counts

Last reconciled inventory run, kept as a sanity check for a new download. A fresh package will differ; a wild difference is a signal to re-check the exclusion rules before trusting the numbers.

## Run 2026-07-27

Source ZIP: `/Users/alfredoteja/Downloads/KAMUS KPI PELINDO GROUP 1 (HO)-20260727T055803Z-1-001.zip`  
Extracted root: `outputs/kamus-ho-config-20260727/source/KAMUS KPI PELINDO GROUP 1 (HO)/`  
Output: `configs/kamus_kpi_ho_visible_20260727.json`

| Measure | v2 | v1 pre-restructure |
| --- | --- | --- |
| Workbooks | 40 | 26 |
| Visible worksheets | 430 | 380 |
| Position worksheets | 398 | 337 |
| Visible non-position sheets | 32 | 40 |
| Title missing in sheet | 0 | 3 |
| Hidden worksheets excluded | 32 | 96 |

Excluded from the inventory:

- `Unit Kerja Pelindo per April 2026 - edit.xlsx` — organisation reference, not a kamus.
- Archive folders whose name contains `(Original)`.

## Prior run 2026-07-20

Output was `configs/temp_visible_kamus_kpi_ho_latest_20260720.json` (removed). That run counted 489 v2 and 411 v1 visible worksheets from folder `(HO) 4`. The 2026-07-27 package is a different download with fewer visible tabs because of tab renames, hides, and cleanup — not because the extractor changed.
