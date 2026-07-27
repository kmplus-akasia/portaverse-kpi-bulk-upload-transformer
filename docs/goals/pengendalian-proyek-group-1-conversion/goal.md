# Pengendalian Proyek Group 1 KPI Conversion

## Outcome

Convert the provided Kamus KPI workbook for Pengendalian Proyek, but only the Group 1 KPIs. Produce conversion outputs plus a position-mapping review artifact where each inferred position mapping can be manually reviewed and corrected with PNID/PMID.

## Source

- Input workbook: `/Users/alfredoteja/Downloads/DIREKTORAT TEKNIK - Ibu Ika Oktania - Pengendalian Proyek (Selesai konfirmasi KPI) (1).xlsx`
- Workspace: `/Users/alfredoteja/Documents/portaverse-kpi-bulk-upload-transformer`

## Constraints

- Use Group 1 only. Do not convert Group 2 or Group 3 rows.
- Fetch the newest production reference data before mapping.
- Treat PNID as `position_nomenclature_mapping.cluster_id`, not the raw mapping row id.
- Keep output per workbook unless the converter contract requires a narrower artifact.
- Produce concrete review files, not chat-only mapping notes.

## Completion Proof

- Latest production reference export exists and has usable PNID/PMID reference rows.
- Group 1 conversion output exists.
- Manual-review position mapping artifact exists with workbook, sheet/group/row context, inferred mapping, and editable PNID/PMID fields.
- Verification checks confirm Group 2/3 were excluded from the converted output.

## Run Command

`/goal Follow docs/goals/pengendalian-proyek-group-1-conversion/goal.md.`
