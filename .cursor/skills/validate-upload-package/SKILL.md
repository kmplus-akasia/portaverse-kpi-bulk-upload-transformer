---
name: validate-upload-package
description: Decide whether a KPI package may be called upload-ready. Use when a conversion, formulir, or amendment finishes, when a workbook was edited in place, or when an older output directory needs re-checking. Other skills reach for this before claiming completion.
---

# Fail-Closed Upload Gate

`upload-ready` is a claim this skill grants. Until every check passes, the artifact is a draft and gets described as one.

## Branches

- **New run** — validate the directory a producing skill just wrote.
- **Retroactive** — scan existing directories under `output/` and `outputs/` for identity-scope inversions with `scripts/audit_converted_kpi_identity_scope.py`, reading artifacts only and regenerating nothing.

## Checks

Run all five. One failure keeps the package a draft.

1. `scripts/validate_kpi_upload_batch.py` returns zero errors: headers match the template's 24 columns, every row carries exactly one identity, and each PMID or PNID exists in the reference.
2. `unzip -t` passes on every delivered `.xlsx`.
3. `IDKPI` is unique and sequential within each file, every OUTPUT resolves to an IMPACT parent, and every KAI resolves to an OUTPUT parent.
4. No cell holds a formula error, each sheet opens on the template's frozen panes, and the KPI-row conditional-formatting and dropdown-validation ranges reach the final populated KPI row.
5. The manifest lists every delivered workbook and nothing besides workbooks.

## Steps

1. **Run the checks and keep the numbers** each one produced.

   Done when: all five checks have a recorded pass or fail together with the counts that produced it.

2. **Write `VALIDATION_RECEIPT.md`** beside the package: generation date, workbook count, KPI row count, the PMID versus PNID split, error count, checks run, checks skipped, positions deliberately left out, and the resolved Kamus source root from `README_SOURCE.md` when the package came from a Kamus conversion.

   Done when: the receipt exists, every figure in it came from check output rather than from expectation, and `README_SOURCE.md` is present for Kamus-derived packages.

## Scope of the claim

These checks read the file. They confirm the package is internally valid and its identities exist, which is a different question from whether uploading it produces the intended production state. When items were removed, `amend-kpi-upload-form` owns that second question, because the importer appends and a valid file can still leave stale KPIs behind.

## Report back

Respond in Indonesian with the verdict, the counts, the checks skipped and why, and the receipt path. When a check fails, name the failing artifact and call the package a draft.
