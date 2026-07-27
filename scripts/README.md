# Scripts

Active entry points, grouped by the workflow stage they serve. `.cursor/skills/` holds the skill that drives each stage; `scripts/archive/` holds one-off builders kept for provenance.

## Converter core

| Script | Role |
| --- | --- |
| `kpi_bulk_transform.py` | workbook-to-template converter; also discovers configs via `--write-discovered-config` |
| `position_mapping.py` | strict position resolver and lookup indexes |
| `build_conversion_recap.py` | recap workbook for a converted batch |

## Mapping and identity

| Script | Role |
| --- | --- |
| `build_mapping_conflict_review.py` | builds the reviewer artifact for unresolved mappings |
| `build_mapping_override_candidates.py` | builds the `Override Candidates` sheet from conflict reviews |
| `apply_position_mapping_review.py` | applies `Reviewer Confirm Mapping` decisions into a config |
| `apply_mapping_override_candidates.py` | applies approved overrides into a config |
| `fix_structural_scope_from_reference.py` | corrects structural scope and PMID against the active reference |

## Historical periods

| Script | Role |
| --- | --- |
| `export_historical_q1_position_reference.mjs` | read-only TEPMS export at a cutoff |
| `historical_q1_reference.mjs` | shared query layer for that export |
| `historical_q1_mapping.py` | historical worksheet resolver |
| `build_historical_q1_position_mapping.py` | writes the historical mapping report and summary |

## Validation and audit

| Script | Role |
| --- | --- |
| `validate_kpi_upload_batch.py` | batch-level fail-closed validation and upload manifest |
| `audit_converted_kpi_identity_scope.py` | retroactive scan of generated workbooks for PMID/PNID scope inversions |

## Inventory

| Script | Role |
| --- | --- |
| `extract_visible_kamus_config.py` | visible worksheet and position-title extraction from a Kamus KPI folder |

## Production reference

| Script | Role |
| --- | --- |
| `export_position_reference.mjs` | read-only production position reference snapshot |
| `export_group1_ho_v2_kpi_production_snapshot.mjs` | read-only production KPI snapshot; `--output`, `--profile`, `--year`, honours `DB_READ_WRITE` |

## Amendment

| Script | Role |
| --- | --- |
| `patch_position_master_ids_in_uploads.py` | replaces exact PMID values inside generated workbooks and writes an audit CSV |

## Dashboard

| Script | Role |
| --- | --- |
| `run_dashboard.py` | launches the KPI planning dashboard; needs the optional `streamlit` dependency |
