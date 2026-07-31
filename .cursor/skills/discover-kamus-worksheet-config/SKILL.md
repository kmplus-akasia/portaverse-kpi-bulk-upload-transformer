---
name: discover-kamus-worksheet-config
description: Inventory the worksheets inside a raw Kamus KPI download into a machine-readable position config. Use when a new Kamus KPI folder or ZIP arrives, when the user asks which worksheets or positions a download contains, or when position-mapping-review needs a candidate catalog.
---

# Inventory Kamus KPI Worksheets

Turn a read-only Kamus KPI download into one inventory that later skills can map. Source workbooks stay exactly as downloaded.

This skill records what each workbook says about itself. Resolving a worksheet to a production identity belongs to `position-mapping-review`.

## Steps

1. **Fix the source.** Record the absolute source root, the download date, and the generation split. Workbooks under `KAMUS KPI HO PRE-RESTRUCTURE` are `v1`; the remaining KPI workbooks are `v2`. Organisation-reference workbooks and archive folders whose name contains `(Original)` stay out of the inventory and go on an explicit exclusion list.

   Done when: workbook counts per generation and the exclusion list are both written down.

2. **Read visibility from OOXML metadata**, not from sheet names. Keep worksheets in state `visible`, and count the `hidden` and `veryHidden` ones left behind so the totals reconcile later.

   Done when: every workbook contributes a visible-worksheet count, and those counts sum to the inventory total.

3. **Extract the position title from sheet content.** Two patterns occur: a value to the right of a `Posisi` or `Nama Posisi` label, and a title sitting directly in `A1`. Support sheets named `Panduan`, `Mapping Organisasi`, `NEW KPI`, or `Jadwal Validator` take precedence over label detection, so they stay out of the position config even when they contain position-like columns. Record the label cell, the value cell, and the extraction method per row.

   Done when: every visible worksheet holds either a title backed by a cell reference, or a `review_status` naming why the title stayed unproven.

4. **Emit the inventory** as JSON under `configs/` plus an editable review workbook under a run-scoped folder in `outputs/`. Use `scripts/extract_visible_kamus_config.py --root <source> --output <json>`. Keep the two generations in separate lists. When the raw folder has been moved into the repository, record the repo-relative moved source root in `metadata.source_root`; later conversion skills must read from that moved root, not from `~/Downloads`.

   Done when: `source_workbook + sheet_name` is unique within each generation list, and `include_in_position_config` is `true` only for rows whose title came from sheet content.

5. **Reconcile and verify.** Compare row counts against the raw visible-worksheet counts, scan the review workbook for formula errors, and run `unzip -t` on it. `references/baseline-counts.md` holds the last reconciled run for comparison.

   Done when: counts match, the archive test passes, and the metadata block records source root, generation timestamp, classification rule, exclusions, and every count.

## Report back

Respond in Indonesian with the source root, workbook and worksheet counts per generation, the number of rows flagged for review, and the paths to both the JSON config and the review workbook.
