# Generate Production Missing KPI Impact Upload

This ExecPlan is a living document. Keep `Progress`, `Surprises & Discoveries`, `Decision Log`, and `Outcomes & Retrospective` up to date as work proceeds.

## Purpose / Big Picture

Produce one upload-ready KPI workbook for Head Office company ID 1 positions that had no complete KPI dictionary in the latest recoverable production audit, plus narrowly verified post-snapshot corrections reported by the user. Every target receives only the ten shared KPI impact items already used by the validated Head Office upload batch.

## Context and Orientation

The importer contract is the 24-column `KPI Template` schema defined in `scripts/kpi_bulk_transform.py` and present in `input/KPI Upload Template.xlsx`. Structural positions use `Position Master ID (Required)`; non-structural positions use `Position Nomenklatur ID` so the importer expands the nomenclature to matching positions.

The latest recoverable read-only production audit was captured on 2026-06-15 for year 2026 and company ID 1. It contains 79 incomplete units: 33 structural PMIDs and 46 non-structural PNIDs. Production cannot be refreshed in this session because the only local production profile uses `root`, which local access policy forbids for routine verification, and the Notion connector is currently unauthorized.

On 2026-06-19 the user reported that Ficky Alkarim, Data Scientist, was missing from the snapshot output. A read-only staging comparison resolved employee `90003230` to PMID `33711` / PMVID `35658`. The production nomenclature reference independently maps PMID `33711` in Department Monitoring dan Evaluasi Klaster Ekspansi Korporasi to PNID `11435`. A different Data Scientist position, PMID `35541` / PNID `12256`, is intentionally excluded.

The ten shared KPI impact rows will be read from a validated workbook under `output/kpi_upload_final_20260618/upload-ready/`. Position and organization labels come from the 2026-06-15 live audit; the older offline reference is used only for secondary ID-existence checks because at least one PNID label has changed since its 2026-06-07 export.

## Scope and Approach

Create one workbook from a validated 24-column upload workbook because the copy under `input/` still has the older header ordering. Repeat the same ten impact definitions for the 79 audited units plus PNID 11435, changing only position ownership and organization labels. Preserve PMID/PNID exclusivity, leave optional IDs blank unless needed to narrow scope, and exclude every non-IMPACT row. Emit the partial PNID 54 gap as exact PMID 348 plus PMVID 39813 so already-covered variants are not expanded into the batch.

This output is a snapshot-based production draft. It must not be represented as a live 2026-06-19 production gap refresh.

## Milestones

### Milestone 1: Recover target and impact contracts

Reconstruct the 79 audited units from the prior production audit evidence, resolve their organization metadata from the production reference, and inspect the representative workbook for the canonical ten impact rows.

Validation: target counts equal 33 structural PMID units plus 46 PNID units plus one exact non-structural PMID/PMVID unit, target identities are unique, and the source contains exactly ten distinct IMPACT rows.

### Milestone 2: Build the upload workbook

Use the bundled `@oai/artifact-tool` runtime to import the official template, populate 790 rows, preserve readable column sizing and frozen headers, and export the workbook under `outputs/019ede9f-kpi-impact-production/`.

Validation: every target has exactly ten rows; every row has KPI Type `IMPACT`; each row has exactly one of PMID or PNID; headers match the importer contract.

### Milestone 3: Verify content and presentation

Inspect key ranges and scan for formula errors with `artifact-tool`, render the populated sheet, visually inspect it, and run ZIP/XML and row-count checks.

Validation: workbook opens, rendered content is legible, total rows equal 790, and no schema or identity errors are found.

## Validation

- `artifact-tool` table inspection of the header and representative rows.
- `artifact-tool` match scan for formula errors.
- Rendered `KPI Template` visual inspection.
- Deterministic receipt confirming 79 targets, 10 impacts, 790 rows, 33 PMID targets, and 46 PNID targets.

## Progress

- [x] Inspected importer, template locations, current upload batch, and production access policy.
- [x] Recovered the latest production gap summary and all 79 target IDs from prior audit evidence.
- [x] Inspect and lock the canonical ten KPI impact rows.
- [x] Build and export the workbook.
- [x] Complete content and visual validation.

## Surprises & Discoveries

- Notion search returned `RuntimeException: unauthorized`, so no current read-only production credential could be obtained.
- The recoverable audit is from 2026-06-15, not a live 2026-06-19 query.
- PNID `54` was partial: two variants had KPI items and one variant remained without KPI items. PNID ownership expansion may therefore touch already-covered variants during import and requires dry-run review.
- The `input/KPI Upload Template.xlsx` copy uses an older header order. The validated 2026-06-18 upload workbook uses the current backend order.
- The backend matches active IMPACT items globally by case-insensitive title when `System KPI ID` is blank, so the ten existing items can be reused without knowing their numeric IDs.
- The missing PNID 54 detail identifies exact PMID 348 / PMVID 39813, allowing a narrow upload that avoids the two already-covered variants.
- The validated source workbook contained an unused blank `Sheet1`. The final artifact was rebuilt with one `KPI Template` sheet only.
- The missing Ficky row was caused by relying on the 2026-06-15 production snapshot. Later user review rejected PNID 11435 as stale; current area-scope evidence points to PNID 11542 for Ficky Alkarim / Data Scientist.

## Decision Log

- Decision: Do not use the local production `root` credential.
  Rationale: The local connection policy explicitly requires a read-only production user.
  Date/Author: 2026-06-19 / Codex

- Decision: Build from the latest recoverable production audit and label the result snapshot-based.
  Rationale: This provides a concrete artifact without making a false live-production claim.
  Date/Author: 2026-06-19 / Codex

- Decision: Replace PNID `54` expansion with exact PMID `348` and PMVID `39813`.
  Rationale: The importer expands PNID to every mapped position, while the audit shows only this variant is missing. Exact targeting avoids changing already-covered variants.
  Date/Author: 2026-06-19 / Codex

- Decision: Supersede PNID `11435` with PNID `11542` for Ficky Alkarim and exclude PNID `12256`.
  Rationale: User review rejected PNID 11435. Live staging plus current area-scope artifacts identify Ficky's Data Scientist assignment under PNID 11542; PNID 12256 belongs to a separate Data Scientist position.
  Date/Author: 2026-06-23 / Codex

## Outcomes & Retrospective

Generated `outputs/019ede9f-kpi-impact-production/KPI_Upload_Production_HO_Missing_Impact_20260619_SNAPSHOT.xlsx`.
Generated Ficky-only upload form `outputs/019ede9f-kpi-impact-production/KPI_Upload_Ficky_Alkarim_Impact_20260623.xlsx`.

The combined workbook contains 800 IMPACT-only rows for 80 target units: the original 79 audited units plus Ficky Alkarim's PNID 11542 correction. Post-export validation confirmed the current 24-column importer headers, ten distinct impact titles and 100% total impact weight per target, 80 unique target identities, exact PMID 348 / PMVID 39813 handling for the partial PNID 54 gap, no PNID 54 expansion, ten PNID 11542 rows, zero PNID 11435 rows, zero PNID 12256 rows, no formula errors, one worksheet only, valid ZIP members, and readable first/last rendered ranges.

The Ficky-only workbook contains 10 IMPACT-only rows for PNID 11542, one worksheet only, zero stale PNID 11435 rows, zero PNID 12256 rows, no formula errors, valid ZIP members, and a readable rendered preview.

Remaining limitation: the target list is the latest recoverable production audit from 2026-06-15. A current production dry-run is still required before confirmed upload because live read-only production access was unavailable on 2026-06-19.
