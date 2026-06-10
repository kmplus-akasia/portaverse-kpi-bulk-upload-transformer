# Portaverse KPI Converter Design

## Purpose

This spec defines the ideal behavior for the Kamus KPI Portaverse converter before converting Group 2 and Group 3 workbooks. The converter must produce upload-ready KPI Template workbooks that are structurally correct, enum-safe, and auditable even when the raw Kamus KPI workbook has typos, copied headers, shifted cells, or values from the wrong enum family.

The converter should not silently hide raw-data issues. If a value can be safely normalized, it should normalize it and record the action. If the row looks structurally damaged or semantically ambiguous, it should spotlight the issue in the report so the user can resolve it before upload.

## Current Problem Findings

### PMID and PNID Scope Errors

Some structural positions were manually or automatically mapped into `Position Nomenklatur ID` instead of `Position Master ID`.

Confirmed examples:

- `Manager Rekrutmen-Karir` in Group Pengelolaan SDM was mapped as `position_nomenclature_id = 515`, `position_scope = non_structural`.
- Production reference shows ID `515` is actually `position_master_id` for `Manager Rekrutmen dan Karir`, with `position_master_type_id = 5`, so it is structural.
- `DH Hubungan Lembaga-Investor` was mapped as `position_nomenclature_id = 67`, but production reference shows ID `67` is `position_master_id` for `Department Head Hubungan Lembaga dan Investor`, also structural.

Required behavior:

- A structural position must output PMID only.
- A non-structural/general position must output PNID only.
- The converter must decide PMID vs PNID from position scope and production position identity, not from the numeric ID alone.
- If a configured PNID also exists as a production structural PMID, the converter must still output PMID when the raw position title/sheet resolves to a structural production position.
- If the raw title/scope signals and production reference cannot produce one clear structural or non-structural identity, the converter must block that position and report a mapping conflict instead of guessing.

### Enum and Cross-Column Pollution

Raw Kamus KPI workbooks contain valid enum variants, spelling variants, and values copied from the wrong column.

Observed raw values:

- Period: `Triwulanan`, `triwulanan`, `Triwulan`, `Tahunan`, `tahunan`, `per tahun`, `Semester`, `Semesteran`, `per Semester`, `Triwulanan/Tahunan`.
- Polarity: `Positif`, `Negatif`, plus invalid pollution such as `DIRECT`, `INDIRECT`, `DUPLICATE`, `Diunggah`, `0`, repeated header text, formulas, and long descriptions.
- Cascading: `DIRECT`, `INDIRECT`, `DUPLICATE`, plus pollution such as `SPECIFIC`, `Routine`, `Non Routine`, `Positif`, `Negatif`, header text, and free-text notes.
- Ownership Type: `SPECIFIC`, `SHARED`, `COMMON`, plus `Shared`, `Specific`, `SPESIFIC`, `Routine`, `Non Routine`, `Positif`, `Negatif`, `Pdf`.
- Nature Of Work: `Routine`, `Non Routine`, plus `Non-Routine`, `Non routine`, `non Routine`, `Non-Rotine`, and pollution such as `DIRECT`, `INDIRECT`, `DUPLICATE`, URLs, and `Pdf`.

Required behavior:

- Every enum output column must be generated from a central normalizer, not by passing raw values through.
- Every normalizer must know which values belong to other enum families.
- Cross-family values must be spotlighted in the report.
- Safe defaults may be applied only when the business meaning is stable.

## Conversion Standards

### Upload Enum Allowlist

The converter must only emit these values:

- `Polarity`: `POSITIVE`, `NEGATIVE`, `NEUTRAL`
- `Period`: `BULANAN`, `TRIWULANAN`, `TAHUNAN`, `SEMESTER`, `MONTHLY`, `QUARTERLY`, `WEEKLY`
- `Cascading`: blank, `DIRECT`, `INDIRECT`, `DUPLICATE`
- `Ownership Type`: blank, `SPECIFIC`, `SHARED`, `COMMON`
- `Nature Of Work (KAI Only)`: blank for non-KAI rows, `Routine` or `Non Routine` for KAI rows

If production DB enum values differ from this list, production DB read-only reference wins. The converter must keep the allowlist in one place so it can be updated without changing parsing logic.

### Period Normalization

The converter should normalize:

- `Triwulan`, `Triwulanan`, `triwulanan` -> `TRIWULANAN`
- `Tahunan`, `tahunan`, `per tahun`, `per tahunan`, `tahun` -> `TAHUNAN`
- `Semester`, `Semesteran`, `per semester` -> `SEMESTER`
- `Bulanan`, `bulanan` -> `BULANAN`
- English equivalents where already supported: `Monthly`, `Quarterly`, `Weekly`

For combined values such as `Triwulanan/Tahunan`, the converter must not guess silently. It should prefer the parent KPI period if available and report an ambiguity. If no reliable parent exists, it should block that KPI row as a report error.

### Polarity Normalization

The converter should normalize:

- `Positif`, `positive`, `pos` -> `POSITIVE`
- `Negatif`, `negative`, `neg` -> `NEGATIVE`
- `Netral`, `neutral` -> `NEUTRAL`

If polarity is blank or polluted by another enum family such as `DIRECT`, `INDIRECT`, `DUPLICATE`, `SPECIFIC`, `SHARED`, `COMMON`, `Routine`, or `Non Routine`, the converter may default to `POSITIVE` because current upload behavior treats missing polarity as positive. It must report the defaulting with raw value, normalized value, source workbook, sheet, row, KPI type, and KPI title.

### Cascading Normalization

The converter should normalize:

- `direct` -> `DIRECT`
- `indirect`, `Indirect` -> `INDIRECT`
- `duplicate` -> `DUPLICATE`

If Cascading contains ownership, nature, or polarity values such as `SPECIFIC`, `SHARED`, `COMMON`, `Routine`, `Non Routine`, `Positif`, or `Negatif`, the converter should output `INDIRECT` and report the cross-column correction. This matches the user decision that typo/pollution in the Cascading column should default to indirect unless an explicit valid cascading value exists.

If Cascading contains long free text or a copied header, the converter should output `INDIRECT` only when the row itself is otherwise structurally valid, and report a warning. If other required fields also look shifted, the row should be spotlighted as a possible shifted-row issue.

### Ownership Type Normalization

The converter should normalize:

- `Specific`, `SPESIFIC`, `spesific`, `specific` -> `SPECIFIC`
- `Shared`, `shared` -> `SHARED`
- `Common`, `common` -> `COMMON`

If Ownership Type is blank or contains values from another enum family such as `Routine`, `Non Routine`, `DIRECT`, `INDIRECT`, `DUPLICATE`, `Positif`, or `Negatif`, the converter must not pass the raw value through. It should default output Ownership Type to `SPECIFIC` and report the default or correction.

### Nature Of Work Normalization

The converter should normalize:

- `Routine`, `routine`, `Rutin` -> `Routine`
- `Non Routine`, `Non-Routine`, `Non routine`, `non Routine`, `Non-Rotine`, `non rotine`, `Non Rutin` -> `Non Routine`

If Nature Of Work is blank, infer from period:

- `Non Routine` when Period is `TAHUNAN`
- `Routine` for every other period, including `BULANAN`, `TRIWULANAN`, `SEMESTER`, `MONTHLY`, `QUARTERLY`, and `WEEKLY`

If Nature Of Work contains cascading values such as `DIRECT`, `INDIRECT`, or `DUPLICATE`, ignore the raw value, infer from period, and report the cross-column correction. URLs, `Pdf`, reference labels, and copied headers should be treated the same way.

## Position Mapping Standards

### Source of Truth

Position scope and ID type must be validated against production reference data. Scope resolution has higher priority than numeric ID uniqueness:

- Structural positions use `position_master_rows` and output `Position Master ID`.
- Non-structural positions use nomenclature `rows` and output `Position Nomenklatur ID`.
- `position_master_type_id = 5` means structural.
- Nomenclature/cluster IDs must not be inferred from an arbitrary numeric ID unless the raw position identity resolves to a non-structural production cluster.
- The same number may exist as both PMID and PNID; the converter must choose the ID column from resolved position scope and title identity, not from numeric ID availability.

### Resolution Rules

The converter should resolve position mapping in this order:

1. Normalize raw position identity from sheet name, discovered `Nama Posisi`, group name, and known aliases/abbreviations.
2. Resolve that identity against production `position_master_rows` and nomenclature `rows` scoped to the target company.
3. Decide structural vs non-structural from the resolved production position, with `position_master_type_id = 5` as structural.
4. Apply exact reviewed config only when its scope and ID type agree with the resolved production identity.
5. Use alias-normalized matching when exact title differs but still resolves to one clear production identity.
6. Manual unresolved report.

If reviewed config conflicts with production reference, scope/title identity must win over the numeric ID field:

- Config says PNID `515`, production also has PNID `515`, but raw position identity resolves to structural `Manager Rekrutmen dan Karir` with `position_master_id = 515` and `position_master_type_id = 5`: correct to structural PMID `515`.
- Config says PNID `67`, but raw position identity resolves to structural `Department Head Hubungan Lembaga dan Investor` with `position_master_id = 67` and `position_master_type_id = 5`: correct to structural PMID `67`.
- If title/scope signals point to multiple active candidates, different companies, or both structural and non-structural identities with similar confidence, block and report a mapping conflict.

### Output Invariants

Every generated KPI row for one position must satisfy:

- Structural: `Position Master ID (Required)` is filled and `Position Nomenklatur ID` is blank.
- Non-structural: `Position Nomenklatur ID` is filled and `Position Master ID (Required)` is blank.
- Never output both IDs.
- Never output neither ID unless the sheet is intentionally skipped and reported.

## Raw Workbook Quality Spotlight

The report must distinguish converter defaults from source-data problems.

Report categories:

- `normalized_enum`: raw value was a known spelling/case variant and safely normalized.
- `cross_column_enum`: raw value belonged to another enum family and was corrected.
- `defaulted_enum`: blank or invalid raw value was replaced by a safe default.
- `ambiguous_enum`: raw value had multiple possible meanings and needs review.
- `shifted_row_suspected`: copied headers, formulas, long descriptions, URLs, or repeated section labels appear in enum columns.
- `mapping_corrected`: PMID/PNID scope was corrected from production reference.
- `mapping_conflict`: mapping could not be corrected safely and must be manually reviewed.

Each report row must include:

- source workbook path/name
- sheet name
- source row
- KPI type
- KPI title
- affected field
- raw value
- normalized output value
- severity
- recommended action

Severity rules:

- `info`: safe normalization, no manual action needed.
- `warning`: upload-safe default/correction was applied but source data should be reviewed.
- `error`: converter cannot safely produce correct output for that row or position.

## Converter Architecture

### Enum Normalizer Module

Create a small, testable enum normalization layer with one function per upload enum family:

- `normalize_period`
- `normalize_polarity`
- `normalize_cascading`
- `normalize_ownership_type`
- `normalize_kai_nature`

Each function returns structured output:

- `value`: upload-safe value or blank
- `status`: `ok`, `normalized`, `defaulted`, `cross_column`, `ambiguous`, or `invalid`
- `raw_value`
- `message`

The row builder should consume these normalized objects and append report issues from their statuses.

### Position Scope Validator

Add a mapping validation layer before row generation:

- Validate configured PMID/PNID against production reference.
- Correct PMID/PNID inversions using resolved position scope and production title identity.
- Block ambiguous mappings.
- Record all corrections and conflicts in the conversion report and recap.

### Final Workbook Gate

Before saving or before returning success, scan generated rows and fail the conversion if:

- Any enum column contains a value outside the allowlist.
- Any row has both PMID and PNID.
- Any row has neither PMID nor PNID.
- Any structural row outputs PNID.
- Any non-structural row outputs PMID.
- Any required upload field is missing.

The converter can still write a workbook for inspection if useful, but the process exit status must be nonzero when upload safety is not guaranteed.

## Recap Workbook Updates

The recap workbook should add:

- Enum normalization counts by field and severity.
- Mapping correction counts.
- Mapping conflict rows.
- Raw workbook quality issue rows.
- Ready-to-upload flag based on final workbook gate, not just row count.

The recap should separate upload blockers from review-only warnings so the user can focus first on issues that prevent upload.

## Test Scenarios

### Unit Tests

Add tests for:

- `per tahun` -> `TAHUNAN`
- `Semesteran` and `per Semester` -> `SEMESTER`
- `Triwulanan/Tahunan` produces ambiguous status unless a parent fallback is used.
- Polarity pollution `INDIRECT`, `DUPLICATE`, `SPECIFIC`, `Routine` defaults to `POSITIVE` and reports correction.
- Cascading pollution `SPECIFIC`, `SHARED`, `COMMON`, `Routine`, `Non Routine`, `Positif` outputs `INDIRECT` and reports correction.
- Ownership variants `Shared`, `Specific`, `SPESIFIC` normalize correctly.
- Ownership pollution `Non Routine` defaults to `SPECIFIC` and is not passed through.
- Nature variants `Non-Routine`, `Non-Rotine`, `non Routine` normalize to `Non Routine`.
- Nature pollution `INDIRECT`, URL, and `Pdf` infer from period and report correction.
- Config PNID `515` for Manager Rekrutmen-Karir corrects to structural PMID `515`.
- Config PNID `67` for DH Hubungan Lembaga-Investor corrects to structural PMID `67`.
- A true conflicting ID blocks conversion.

### Fixture Tests

Use small synthetic workbook fixtures that represent:

- Normal source row.
- Repeated header row inside data.
- Shifted row where enum columns contain descriptions/formulas.
- Cross-column enum pollution.
- Structural and non-structural position mappings.

### Batch Verification

Run against:

- Existing Group 1 HO pre-restructure source zip.
- Existing Group 3 source zip.
- Group 2 source zip after it is downloaded.

Acceptance criteria:

- Zero invalid enum values in generated upload workbooks.
- Zero PMID/PNID scope inversions.
- Every auto-correction appears in report.
- Every ambiguous/unsafe row is spotlighted before upload.
- The converter exits nonzero only for true upload blockers.

## Operational Boundaries

- Production DB may be used read-only to verify enum truth and position scope.
- No production write, import, delete, or migration is allowed from this converter workflow.
- Credentials must not be committed or written into repo files.
- Raw source workbooks remain the audit source for KPI content.
- Production reference JSON is the source for position ID type and company scope when DB access is unavailable.

## Open Implementation Notes

- Group 2 source files are not yet available locally. The converter must handle Group 2 with the same parser and validation rules once the zip is downloaded.
- The enum allowlist should be easy to update if production DB reveals additional backend-accepted values.
- Existing manual corrections in Group 1 should be treated as regression cases where possible.
