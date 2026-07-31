# Config files

Active configs used by the KPI upload workflow. Run-scoped conversion artifacts live under `outputs/<run>/` and `output/<run>/`.

| File | Role | Status |
| --- | --- | --- |
| `kamus_kpi_ho_visible_20260729.json` | Worksheet inventory from the 2026-07-29 Kamus KPI HO folder | **Current** — input for `position-mapping-review` |
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
