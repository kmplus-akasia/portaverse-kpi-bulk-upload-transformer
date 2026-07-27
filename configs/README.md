# Config files

Active configs used by the KPI upload workflow. Run-scoped snapshots live under `outputs/<run>/`.

| File | Role | Status |
| --- | --- | --- |
| `kamus_kpi_ho_visible_20260727.json` | Worksheet inventory from the 2026-07-27 Kamus KPI HO download | **Current** — input for `position-mapping-review` |
| `production_position_reference.json` | Read-only production position/org snapshot | **Committed baseline** — last export that fits GitHub's 100 MB limit; refresh locally with `export_position_reference.mjs` |
| `pre_restructure_positions.json` | Pre-restructure worksheet inventory (295 rows) | **Historical Q1 only** — do not use for current v2 mapping |
| `pre_restructure_positions_rw_reviewed_20260609.json` | Reviewer-approved historical Q1 identities | **Historical Q1 only** |
| `sample.json` | Minimal converter config example | Fixture |
| `all_positions.json` | SDM group example for README | Fixture |

Superseded files are removed, not kept in `configs/`. Older production snapshots remain under `outputs/` when a run produced them.

Latest local-only refresh (108 MB, includes structural lookup rows): `outputs/kamus-ho-config-20260727/production_position_reference_20260727.json`.
