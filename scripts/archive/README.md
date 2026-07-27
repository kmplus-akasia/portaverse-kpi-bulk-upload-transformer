# Archived scripts

One-off builders kept for provenance. Each was written for a single dated run and hardcodes that run's source paths, output directory, or business rules, so none of them is a general entry point.

## Restoring one

These scripts compute their repo root as `Path(__file__).resolve().parents[1]`, which resolves to `scripts/` from this folder rather than the repo root. Copy a script back into `scripts/` before running it, and re-point its hardcoded paths at the current run.

## What each one did

### Group 1 Head Office v2, July 2026

A sequence of narrow upload builders produced while the Group 1 HO v2 batch was being reconciled against production. Each targets a specific slice of that reconciliation.

| Script | Slice |
| --- | --- |
| `build_group1_ho_v2_missing_upload.py` | positions absent from the converted batch |
| `build_group1_ho_v2_unconverted_mapped_upload.py` | positions that had a mapping but no conversion |
| `build_group1_ho_v2_two_positions_upload.py` | two named positions |
| `build_group1_ho_v2_reviewed_14_one_upload.py` | the 14 reviewer-approved identities, as one workbook |
| `build_group1_ho_v2_remaining_37_resolution_audit.py` | resolution audit for the remaining 37 |
| `build_group1_ho_v2_bad_production_full_upload.py` | full rebuild for identities with bad production KPI |
| `build_group1_ho_v2_bad_production_kamus_only_upload.py` | the same slice, kamus rows only |
| `build_group1_ho_v2_bad_production_strict_upload.py` | the same slice under strict identity rules |
| `build_group1_ho_v2_delta_remediation.py` | delta against the 2026-07-09 production KPI snapshot |
| `generate_group1_ho_v2_followup_upload.py` | follow-up upload after the first delivery |

The general form of this work now lives in `amend-kpi-upload-form`, which handles add, remove, replace, and identity patching as a delta with a comparison sheet.

### Project positions

| Script | Slice |
| --- | --- |
| `build_project_position_upload_config.py` | config for Pengendalian Proyek positions; hardcodes ten position families and the project tokens BMTH, Kalibaru, NPEA, JICT Koja, and Kijing |
| `split_project_upload_by_project.py` | splits one generated project workbook into per-project workbooks; hardcodes the same project buckets |

Building a config now goes through `discover-kamus-worksheet-config` plus `apply-position-identity-config`.

### Mapping audit

`build_position_mapping_manual_audit.py` built a manual audit workbook from a `position_mapping_report` export sitting in `~/Downloads`. Superseded by `position-mapping-review`.

### Coverage audit

`audit_dashboard_kpi_gaps.py` pulled dashboard data from production and compared it against generated upload workbooks. It answers the same coverage question as the `update-org-kpi-audit-report` skill, which runs from the `dashboard-org-kpi-audit` pipeline and is the maintained path.

### Staging export

`export_staging_nomenclature_mapping.mjs` exported the nomenclature mapping from the staging profile into `configs/staging_nomenclature_mapping.json`. Identity decisions cite production, so the production exporter `export_position_reference.mjs` is the maintained path.
