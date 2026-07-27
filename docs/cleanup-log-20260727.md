# Cleanup log, 2026-07-27

Repository organisation run. Every deletion below was approved item by item; nothing was mass-deleted, and no run directory, manifest, or validation receipt was touched.

## Deleted

| Path | Size | Why it was safe |
| --- | --- | --- |
| `output/ho_structure_pdf_audit_production_reference_retry_20260702.json` | 109 MB | production reference copy, no document cited it, re-exportable |
| `output/group1_k3_meko_meka_prod_reference_20260706.json` | 109 MB | production reference copy, no document cited it, re-exportable |
| `tmp/group1_ho_v2_20260716_raw/` | 180 MB | extracted copy of a raw Kamus KPI download that lives in `~/Downloads` |
| `tmp/group1_ho_v2_source_20260703.zip` | 48 MB | same, as an archive |
| `outputs/019f85c0-historical-q1-upload-20260722/…188_Identity_20260722.xlsx.inspect.ndjson` | 39 MB | inspection dump; its workbook remains in place |
| `outputs/019f85c0-historical-q1-final-v1-20260723/…TW1_V1_20260723.xlsx.inspect.ndjson` | 21 MB | inspection dump; its workbook remains in place |
| 6 `__pycache__` directories | 6 MB | bytecode cache |

Directory sizes went from `output` 1246 MB, `outputs` 386 MB, `tmp` 267 MB to 1030 MB, 327 MB, and 41 MB.

## Kept deliberately

- `output/production_position_reference_20260716_issue_audit.json` and `output/production_kpi_snapshot_20260716_issue_audit.json` — cited nine and seven times respectively as identity evidence for the issue remediation work.
- Every remaining production reference copy inside a run directory, so each delivered batch keeps the snapshot it was built against.
- Every `UPLOAD_THESE_FILES.md`, `VALIDATION_RECEIPT.md`, and upload-ready workbook.

## Scripts archived

Fifteen one-off builders moved from `scripts/` to `scripts/archive/`, which now carries a README recording what each one did and which skill supersedes it. Four were git-tracked and moved with `git mv`, so history follows them.

`scripts/` keeps eighteen active entry points, indexed in `scripts/README.md`.

## Verification

`python3 -m unittest` over the eight converter, mapping, historical, and validation test modules: 115 tests, OK. `python3 -m py_compile` passed on the converter core, validator, identity-scope audit, and PMID patch scripts.

## .gitignore

Added `tmp/` and `*.inspect.ndjson`. Files already tracked stay tracked; this only stops new artifacts from filling `git status`.
