# Historical Q1 Position Mapping Design

## Goal

Produce a reviewable mapping from every worksheet in the Head Office pre-restructure KPI dictionaries to the historical PMID or PNID supported by employee-position assignments that ended on 31 March 2026. Conversion into upload forms starts only after the user approves the mapping workbook.

## Scope

- Source KPI dictionaries: 26 workbooks under `KAMUS KPI HO PRE-RESTRUCTURE`.
- Mapping grain: one row per `(source_workbook, worksheet)`.
- Expected worksheet inventory: 295 rows, based on `configs/pre_restructure_positions.json`.
- Historical cutoff: `DATE(tb_employee_position_master_sync.end_date) = '2026-03-31'`.
- Company scope: `company_id = 1`, PT Pelabuhan Indonesia (Persero), matching the converter's Head Office default.
- Conversion output: one upload workbook per source workbook.
- Production access is read-only. No import, update, delete, migration, or other production write is allowed.

## Non-goals

- Do not rebuild Quarter 2 or current-position mappings.
- Do not infer historical positions from current active assignments.
- Do not silently approve fuzzy, ambiguous, or missing mappings.
- Do not convert any worksheet before the user completes the manual review gate.
- Do not change raw KPI dictionary content.

## Source of Truth

Historical worker-position evidence comes from this chain:

`tb_employee_position_master_sync` -> `tb_position_master_variant` -> `tb_position_master_v2`

The historical cohort query must apply:

```sql
WHERE tepms.deletedAt IS NULL
  AND DATE(tepms.end_date) = '2026-03-31'
```

The export must retain the exact `start_date`, `end_date`, `employee_number`, `position_master_variant_id`, `lakhar_id`, and `job_sharing_id` from TEPMS. Employee names may be added from `tb_employee` for review convenience, but employee status must not remove an otherwise valid non-deleted historical TEPMS record.

Historical organization context should be resolved through `tb_position_master_organization_sync`, `tb_group_master`, and `tb_company_in` for records whose effective period contains 31 March 2026. When historical organization context is unavailable, the row remains in the cohort and is marked with missing-organization evidence instead of being dropped.

`position_nomenclature_mapping` is only a bridge from historical PMID to a PNID candidate. It has no effective-date fields and therefore is not historical proof by itself.

The existing files `configs/pre_restructure_positions.json` and `configs/pre_restructure_positions_rw_reviewed_20260609.json` are candidate hints only. They must not override TEPMS evidence.

## Position Identity Rules

### Structural positions

- `tb_position_master_v2.position_master_type_id = 5` means structural.
- The proposed upload identity is the PMID reached from `position_master_variant_id`.
- The PNID field must remain blank.

### Non-structural positions

- A position whose `position_master_type_id` is not `5` is treated as non-structural.
- TEPMS yields a PMID, not a PNID.
- PNID candidates come from `position_nomenclature_mapping.cluster_id` rows associated with the historical PMID.
- Candidate selection uses position title, group, company, and job-class evidence.
- A unique, context-consistent cluster becomes the proposed PNID.
- Multiple plausible clusters, no cluster, or conflicting context produces `NEEDS_CHECK`; no PNID is auto-selected.
- When a PNID is approved, the PMID field must remain blank because the importer expands PNID to its mapped PMIDs.

### Assignment type

- `lakhar_id IS NULL AND job_sharing_id IS NULL` means `PRIMARY`.
- A non-null `lakhar_id` means `LAKHAR`.
- A non-null `job_sharing_id` means `JOB_SHARING`.
- All three assignment types remain visible as evidence.
- Primary evidence has precedence when primary and secondary assignments point to different identities for the same employee and worksheet. The conflict remains visible in the review queue.

## Worksheet Mapping

Each worksheet is matched against only the historical Q1 cohort. Matching evidence is ranked in this order:

1. Exact normalized worksheet position title plus historical group.
2. Exact normalized title with a unique historical position in company `1`.
3. Strong token match for title plus matching group and company.
4. Existing pre-restructure config identity as a comparison hint.

Workbook path and worksheet name form the stable key. Identically named worksheets in different workbooks remain separate review rows.

The mapping result uses these confidence labels:

- `high_confidence`: unique historical identity with exact or equivalent title and consistent group/company evidence.
- `low_confidence`: plausible identity with a weak title or incomplete context.
- `mapping_conflict`: two or more strong identities compete.
- `no_candidate`: no historical identity supports the worksheet.

Confidence is advisory. Every worksheet still requires explicit reviewer approval before conversion.

## Review Workbook

Create one workbook at:

`outputs/019f6a65-722b-72f1-95ee-ec8aa577e55a/Historical_Q1_Position_Mapping_Review_20260716.xlsx`

It contains four sheets:

### Summary

- Cutoff and company scope.
- Counts for workbooks, worksheets, historical assignments, employees, unique PMIDs, and PNID candidates.
- Mapping counts by confidence and reviewer status.
- Clear statement that conversion has not started.

### Historical TEPMS

One row per historical assignment with employee, assignment type, start/end dates, variant, PMID, position title/type, historical organization context, and PNID candidate evidence.

### Position Mapping Report

One row per source workbook and worksheet with:

- Source Workbook
- Worksheet
- Worksheet Position
- Worksheet Group
- Historical Employee Numbers
- Historical Employee Names
- Assignment Types
- Historical End Date
- Inferred Scope
- Candidate PMID
- Candidate PNID
- Candidate Position Title
- Candidate Group
- Candidate Company
- Confidence Label
- Confidence Reason
- Existing Config PMID
- Existing Config PNID
- Reviewer Confirm Mapping
- Reviewer Actual PMID
- Reviewer Actual PNID
- Reviewer Notes

`Reviewer Confirm Mapping` accepts only `YES` or `NEEDS_CHECK`.

### Review Queue

Includes every row that is not `high_confidence`, has multiple historical identities, lacks historical organization evidence, has a PNID ambiguity, or remains unapproved.

## Manual Review Gate

The user reviews and returns the mapping workbook.

Before applying it:

- Every converted row must have `Reviewer Confirm Mapping = YES`.
- Exactly one of Reviewer Actual PMID or Reviewer Actual PNID may be populated.
- A reviewer ID overrides the candidate only after namespace validation.
- A PMID override must resolve to a historical structural identity. An unsupported PMID may be documented in Reviewer Notes but remains blocked from conversion.
- A PNID override must exist in `position_nomenclature_mapping` or include an explanatory reviewer note and remain blocked from conversion.
- `NEEDS_CHECK`, blank, contradictory, or invalid rows remain excluded.

The approved workbook is transformed into a dedicated reviewed config. The original config files remain unchanged.

## Conversion

After the manual review passes:

- Convert approved worksheets only.
- Produce one upload workbook per source workbook.
- Use PMID only for structural rows and PNID only for non-structural rows.
- Preserve the raw KPI content and existing converter normalization rules.
- Emit a validation report and manifest that tie every output workbook back to its approved mapping rows.
- Do not include skipped or unresolved worksheets in upload-ready files.

## Error Handling

- Duplicate TEPMS assignments are retained in raw evidence and deduplicated only for aggregated mapping counts.
- Conflicting PMIDs or PNIDs for one worksheet create `mapping_conflict`.
- Missing employee names do not discard the assignment; employee number remains the identifier.
- Missing historical organization context creates a review warning.
- A production query failure stops the historical export and produces no partial review workbook.
- A worksheet inventory mismatch against 295 rows stops the review artifact from being marked complete.
- An invalid reviewer value, both IDs populated, or wrong namespace blocks config generation.

## Verification

Historical export checks:

- Every exported TEPMS row has `DATE(end_date) = '2026-03-31'`.
- Every exported row has `deletedAt IS NULL`.
- TEPMS row count reconciles with the direct production count query.
- Distinct employee, PMVID, PMID, and assignment-type counts are recorded.

Review workbook checks:

- Exactly 26 source workbooks and 295 mapping rows are present.
- `(Source Workbook, Worksheet)` is unique.
- No row contains both Candidate PMID and Candidate PNID.
- Reviewer input cells are editable and validated.
- All four sheets are visually inspected for clipping, readability, and broken formulas.

Conversion checks:

- No unapproved mapping is converted.
- Every converted row has exactly one upload identity namespace.
- Generated files pass the existing batch validator and ZIP integrity check.
- Manifest counts reconcile to approved worksheet and output-row counts.
- Any skipped check or unresolved row is reported explicitly.

## Acceptance Criteria

1. A read-only historical TEPMS export exists for assignments ending on 31 March 2026.
2. All 295 pre-restructure worksheets appear in the mapping review workbook.
3. Every proposed PMID/PNID includes traceable historical evidence and a confidence reason.
4. Ambiguous or unsupported mappings are blocked and placed in the review queue.
5. The user can approve or correct mappings directly in the workbook.
6. No KPI conversion occurs before the reviewed workbook is returned and validated.
7. After approval, upload forms are generated per source workbook and contain only approved position identities.
