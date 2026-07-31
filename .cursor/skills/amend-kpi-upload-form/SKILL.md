---
name: amend-kpi-upload-form
description: Amend a delivered KPI upload form as a delta. Use when the user wants to add KPI items on top of an existing kamus, remove or drop items from a form, replace items after an issue correction, or patch identity values inside a workbook that has already been generated.
---

# Amend a Delivered Formulir

## The importer appends

A KPI already in production stays there when its row is absent from the next upload. A form that simply omits an item therefore leaves production holding the old item *and* the new set, and every structural, identity, and integrity check still passes, because the file itself is valid.

So a removal ships two artifacts: the form stating the intended final state, and an explicit archive list naming the items a human must retire through the replacement workflow.

| Branch | What changes |
| --- | --- |
| Add | new items appended on top of the existing kamus |
| Remove | items absent from the final state, plus the archive list |
| Replace | a removal and an addition bound to the same parent |
| Patch identity | PMID, PMVID, or PNID corrected inside a generated workbook |

## Steps

1. **Establish the before state** from the delivered form, and from a read-only production KPI snapshot when the change answers to what production currently holds. `scripts/export_group1_ho_v2_kpi_production_snapshot.mjs` takes `--output`, `--profile`, and `--year`, and honours `DB_READ_WRITE=0`. When adding KPI rows from Kamus, resolve raw workbooks through `scripts/kamus_source.py` and record the source in `README_SOURCE.md`.

   Done when: the before state names each affected item with its parent and its current identity.

2. **Apply the delta and rebuild the hierarchy.** Renumber `IDKPI` as `1..N`, regenerate every `Parent KPI ID`, restate the weight total per level, and resize the template's conditional-formatting and dropdown-validation ranges through the amended file's final populated KPI row.

   Done when: every OUTPUT has a parent IMPACT and every KAI has a parent OUTPUT inside the amended file, the weight total per level is stated rather than assumed, and every KPI row remains inside the active formatting and validation ranges.

3. **Write the comparison sheet.** One row per touched item: title, parent, before status, after status, and the action a human takes outside the upload.

   Done when: every added, removed, and replaced item appears with both statuses, and each removed item carries its archive instruction.

4. **Verify and version.** Run `validate-upload-package`, then name the file for the revision so the earlier delivery keeps its own name.

   Done when: validation reports zero errors and the previous file is still present under its original name.

## Report back

Respond in Indonesian with the branch, counts of added, removed, and replaced items, the final IMPACT/OUTPUT/KAI totals, the archive actions required outside the upload, and the amended file path.
