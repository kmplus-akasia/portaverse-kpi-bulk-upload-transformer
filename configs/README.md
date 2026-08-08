# Config files

Active configs used by the KPI upload workflow. Run-scoped conversion artifacts live under `outputs/<run>/` and `output/<run>/`.

| File | Role | Status |
| --- | --- | --- |
| `kamus_kpi_group2_visible_20260807.json` | Worksheet inventory from Kamus KPI Group 2 (Regional, Cabang, Subholding) 2026-08-07 | **Current Group 2** — supersedes `20260805` and interim Subholding-only `20260807` |
| `kamus_kpi_ho_visible_20260806.json` | Worksheet inventory from the 2026-08-06 Kamus KPI HO folder (`KPI MKST.xlsx` installed under Group Sekretariat Perusahaan; old PA folder archived) | **Current HO** — input for `position-mapping-review` |
| `kamus_kpi_mkst_pa_sekretaris_20260806.json` | PA + Sekretaris positions extracted from hidden sheet `Sekretariat` inside `KPI MKST.xlsx` | **Current PA/Sekretaris kamus source** for Sekper mapping/upload |
| `kamus_kpi_ho_visible_20260803.json` | Worksheet inventory from the 2026-08-03 Kamus KPI HO folder | Superseded by `20260806` for HO |
| `kamus_kpi_ho_visible_20260729.json` | Worksheet inventory from the 2026-07-29 Kamus KPI HO folder | Superseded by `20260806` for HO |
| `production_position_reference.json` | **Canonical** read-only production position/org snapshot | Refresh with `scripts/refresh_canonical_production_reference.sh` |
| `production_position_reference.meta.json` | Export timestamp, row counts, review status | Updated on every canonical refresh |
| `pre_restructure_positions.json` | Pre-restructure worksheet inventory (295 rows) | **Historical Q1 only** |
| `pre_restructure_positions_rw_reviewed_20260609.json` | Reviewer-approved historical Q1 identities | **Historical Q1 only** |
| `sample.json` | Minimal converter config example | Fixture |
| `all_positions.json` | SDM group example for README | Fixture |

## Production reference maintenance

One canonical snapshot only:

```sh
DB_READ_WRITE=0 ./scripts/refresh_canonical_production_reference.sh
```

- **JSON:** `configs/production_position_reference.json`
- **Receipt:** `outputs/production-reference/REFERENCE_RECEIPT.md`
- **Metadata:** `configs/production_position_reference.meta.json`

Do not copy production snapshots into run folders. Cite the canonical path and `exported_at` in run receipts instead.

The export contains employee names and NIPP. Treat it as sensitive operational data; do not commit refreshed snapshots without explicit review (~110 MB, exceeds GitHub's 100 MB limit).

Superseded dated snapshots under `outputs/` and `output/` were removed on 2026-07-29 in favour of the canonical path above.
