# Finalize Position Scope and Regeneration

This ExecPlan is a living document. Keep `Progress`, `Surprises & Discoveries`,
`Decision Log`, and `Outcomes & Retrospective` current during execution.

## Purpose / Big Picture

Produce a clean, upload-ready Head Office batch where position identity follows
the production contract: structural positions use PMID and non-structural
positions use PNID. Remove obsolete conversion output after the new batch passes
all checks.

## Context and Orientation

The defect originates in `scripts/fix_structural_scope_from_reference.py`, which
currently converts a non-structural config to structural when its numeric value is
found in `position_master_rows`. The production snapshot is
`configs/production_position_reference.json`; type `5` is structural and PNID is
`rows[].cluster_id`. The reviewed input config is
`configs/pre_restructure_positions_rw_reviewed_20260609.json`. The source ZIP is
`/Users/alfredoteja/Downloads/KAMUS KPI HO PRE-RESTRUCTURE-20260602T070011Z-3-001.zip`.

## Scope and Approach

Use TDD to replace numeric-collision inference with production-type resolution.
Strengthen `scripts/validate_kpi_upload_batch.py` so it checks identity semantics,
not only ID existence. Generate into `output/kpi_upload_final_20260618/`, validate
before cleanup, then delete every sibling under `output/`.

No DB refresh, source workbook edits, template redesign, or unrelated converter
refactor is included.

## Milestones

### Milestone 1: Regression protection

Create focused tests for type `5`, type `6`, type `4`, unique PNID resolution,
ambiguous PNID failure, and validator rejection of a non-structural PMID. Run tests
against current code and observe the expected failure.

### Milestone 2: Type-driven correction

Update the postprocessor and validator with the minimum implementation needed for
the tests. Run focused and full test suites plus Python compile checks.

### Milestone 3: Regeneration and validation

Generate corrected config and audit, run the 26-workbook ZIP conversion, then run
the batch validator. Independently scan each workbook against the reference,
verify ZIP CRC/XML parseability, and assert the two reported examples use PNID 10
and 2901.

### Milestone 4: Output cleanup

After Milestone 3 succeeds, retain only `output/kpi_upload_final_20260618/`.
Read back the final directory, manifest, workbook count, ZIP, and validation
receipt.

## Validation

```bash
.venv/bin/python -m unittest tests.test_fix_structural_scope_from_reference
.venv/bin/python -m unittest tests.test_kpi_bulk_transform
.venv/bin/python -m py_compile scripts/fix_structural_scope_from_reference.py scripts/validate_kpi_upload_batch.py
.venv/bin/python scripts/validate_kpi_upload_batch.py --output-dir output/kpi_upload_final_20260618 --config output/kpi_upload_final_20260618/corrected_positions.json --reference configs/production_position_reference.json --fixed-audit output/kpi_upload_final_20260618/scope_correction_audit.csv --expected-workbooks 26 --upload-ready-dir output/kpi_upload_final_20260618/upload-ready --zip-output output/kpi_upload_final_20260618/KPI_Upload_Final_20260618.zip
```

## Progress

- [x] Root cause and 63-row correction set audited.
- [x] Cleanup scope approved: retain only final batch and audit artifacts.
- [x] Regression tests fail for current behavior.
- [x] PNID-first correction, neglect handling, exact structural lookup, and validator hardening pass tests.
- [x] 26 workbooks regenerate and pass all validations.
- [x] Obsolete `output/` contents are removed.

## Surprises & Discoveries

- Of 63 prior automatic conversions, 54 configured values were already valid PNIDs.
  Only 9 values were invalid as PNIDs and valid structural PMIDs. Numeric namespace
  validity must be checked before master-type inference.
- Thirteen configs had blank identity and scope. They are now explicit `neglect`,
  preventing Master Posisi fallback from emitting non-structural PMIDs.
- Exact structural lookup must run before fuzzy normalization; otherwise numeric
  suffixes such as `Wilayah Timur 1/2` collapse to the wrong PMID.
- The original source ZIP still exists at the runbook path in `~/Downloads`.

## Decision Log

- Decision: Treat a valid explicit PNID as authoritative, then use production type
  only when the configured value is not a valid PNID.
  Rationale: Numeric values legitimately overlap between PMID and PNID namespaces.
  Date/Author: 2026-06-18 / Codex
- Decision: Abort rather than guess when a non-structural PMID maps to zero or
  multiple PNIDs.
  Rationale: Silent guesses caused the current defect.
  Date/Author: 2026-06-18 / Codex
- Decision: Clean `output/` only after fresh batch validation succeeds.
  Rationale: Preserve rollback evidence until a valid replacement exists.
  Date/Author: 2026-06-18 / Codex

## Outcomes & Retrospective

Implemented regression coverage and regenerated 26 workbooks. The final validator
reported 6,118 rows, 3,838 structural rows, 2,280 non-structural rows, and zero
errors. ZIP CRC/XML validation passed for all 26 workbooks (260 XML members), and
two representative workbooks opened successfully through LibreOffice.

Cleanup completed. `output/` now contains only
`output/kpi_upload_final_20260618/`, including the upload-ready workbooks, ZIP,
corrected config, audit, manifest, instructions, reports, and validation receipt.
