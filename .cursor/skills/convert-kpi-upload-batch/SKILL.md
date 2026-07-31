---
name: convert-kpi-upload-batch
description: Convert a whole Kamus KPI batch into upload-ready workbooks. Use when the user asks to convert a group, folder, or ZIP of KPI workbooks, to regenerate a batch after a scope or template fix, or to produce an upload package with a manifest.
---

# Batch Conversion to Upload-Ready

## Steps

1. **Pin the run.** Record the source (workbook, folder, or ZIP), the template version, the config, the reference snapshot with its export timestamp, and a run-scoped output directory. Commands are in `references/conversion-commands.md`.

   For Head Office Kamus KPI, resolve the raw source through `scripts/kamus_source.py` and the inventory at `configs/kamus_kpi_ho_visible_20260729.json`. The canonical repo root is `outputs/kamus-ho-config-20260729/source/KAMUS KPI PELINDO GROUP 1 (HO) 5`. Write `README_SOURCE.md` beside the run before conversion starts.

   Done when: all five are written down, the resolved Kamus source root is inside the repository (not `~/Downloads`), and the output directory is new so a prior run stays readable beside this one.

2. **Settle identity before conversion.** With an approved config, refresh candidates against the reference. Without one, discovery writes a config through `--write-discovered-config`; worksheets it cannot resolve go to `position-mapping-review` and stay out of this batch.

   Done when: every worksheet entering conversion holds exactly one identity shape, and the unresolved ones are listed by name.

3. **Convert.** The converter keeps the official template structure, extends conditional formatting and dropdown validation through each workbook's final populated KPI row, and writes one workbook plus one report CSV per source workbook.

   Done when: the workbook count matches the expected count, every report CSV has been read for rows of severity `error`, and formatting and validation coverage reaches the final KPI row in every workbook regardless of row count.

4. **Package.** Collect the upload-ready workbooks, write `UPLOAD_THESE_FILES.md` listing workbooks only, and produce a ZIP when the batch ships as one file.

   Done when: the manifest names every deliverable workbook and nothing else, so a report, config, or audit file cannot be uploaded by mistake.

5. **Verify** through `validate-upload-package`.

   Done when: that skill reports zero errors and the receipt sits beside the batch.

## Report back

Respond in Indonesian with the run directory, workbook and KPI row counts, the split of rows using PMID versus PNID, the positions deliberately left out, and the validation result.
