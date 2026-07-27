# Historical Q1 Position Mapping Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Produce an editable review workbook that maps all 295 Head Office pre-restructure KPI worksheets to evidence-backed historical PMID/PNID candidates from TEPMS assignments ending on 31 March 2026.

**Architecture:** A small Node exporter reads production using the existing read-only profile and writes a traceable historical reference JSON. A Python resolver maps the 295 config rows to the reference, retaining exact worker/assignment evidence and treating PNID as a current nomenclature bridge only. A conversation-scoped JavaScript builder uses `@oai/artifact-tool` to create and visually verify the editable review workbook. No upload conversion occurs until the returned workbook is approved.

**Tech Stack:** Node.js + `mysql2/promise` from the PMS service runtime; Python 3.11 standard library; existing `scripts/position_mapping.py` normalization rules; `@oai/artifact-tool` for `.xlsx` authoring; `unittest` and Node's built-in test runner.

## Global Constraints

- Historical cutoff is exactly `DATE(tepms.end_date) = '2026-03-31'`.
- Only `tepms.deletedAt IS NULL` assignments are historical evidence.
- Head Office mapping candidates require `company_id = 1`; assignments without confirmable historical organization stay in raw evidence but cannot auto-map a worksheet.
- Structural positions use PMID only when `position_master_type_id = 5`; non-structural positions use PNID only.
- PNID is derived from `position_nomenclature_mapping`; it is never treated as a historical TEPMS field.
- Every one of the 295 `(source_workbook, worksheet)` keys appears in the review workbook.
- The workbook exposes `Reviewer Confirm Mapping`, `Reviewer Actual PMID`, `Reviewer Actual PNID`, and `Reviewer Notes`; only `YES` can later be converted.
- Production access remains read-only. Do not alter existing raw workbooks, checked-in configs, or unrelated dirty files.
- The final review workbook is written to `outputs/019f6a65-722b-72f1-95ee-ec8aa577e55a/`.

---

### Task 1: Build the read-only historical-reference exporter

**Files:**

- Create: `scripts/historical_q1_reference.mjs`
- Create: `scripts/export_historical_q1_position_reference.mjs`
- Create: `tests/historical_q1_reference.test.mjs`

**Interfaces:**

- `buildHistoricalAssignmentQuery()` returns SQL that selects TEPMS assignment metadata, PMID, position type/title, and historical organization fields with two placeholders: cutoff date and company ID.
- `classifyAssignment(row)` returns `PRIMARY`, `LAKHAR`, or `JOB_SHARING`.
- `shapeHistoricalPayload({ profile, cutoffDate, companyId, assignmentRows, nomenclatureRows })` returns `{ source, historical_assignment_rows, nomenclature_rows }`.
- CLI: `node scripts/export_historical_q1_position_reference.mjs --profile production --cutoff-date 2026-03-31 --company-id 1 --output <path>`.

- [ ] **Step 1: Write failing Node tests for query and assignment classification**

```js
import test from 'node:test';
import assert from 'node:assert/strict';
import { buildHistoricalAssignmentQuery, classifyAssignment } from '../scripts/historical_q1_reference.mjs';

test('historical query is parameterized and anchored to the requested cutoff', () => {
  const sql = buildHistoricalAssignmentQuery();
  assert.match(sql, /DATE\(tepms\.end_date\) = \?/);
  assert.match(sql, /tepms\.deletedAt IS NULL/);
  assert.match(sql, /tpmos\.organization_master_id/);
});

test('classifies primary and secondary historical assignments', () => {
  assert.equal(classifyAssignment({ lakhar_id: null, job_sharing_id: null }), 'PRIMARY');
  assert.equal(classifyAssignment({ lakhar_id: 2, job_sharing_id: null }), 'LAKHAR');
  assert.equal(classifyAssignment({ lakhar_id: null, job_sharing_id: 9 }), 'JOB_SHARING');
});
```

- [ ] **Step 2: Run the Node test and verify it fails because the module is absent**

Run: `node --test tests/historical_q1_reference.test.mjs`

Expected: `ERR_MODULE_NOT_FOUND` for `scripts/historical_q1_reference.mjs`.

- [ ] **Step 3: Implement the pure query/payload module**

```js
export function classifyAssignment(row) {
  if (row.lakhar_id != null) return 'LAKHAR';
  if (row.job_sharing_id != null) return 'JOB_SHARING';
  return 'PRIMARY';
}

export function buildHistoricalAssignmentQuery() {
  return `SELECT ... WHERE tepms.deletedAt IS NULL AND DATE(tepms.end_date) = ?`;
}
```

The SQL joins TEPMS to variant and position master, uses a left historical organization join with the cutoff date, and retains organization-missing rows. The CLI loads the existing `~/.codex/pms-connections/<profile>.env` contract without printing credentials, runs parameterized queries, writes JSON, and closes the connection in `finally`.

- [ ] **Step 4: Run the Node test and syntax check**

Run: `node --test tests/historical_q1_reference.test.mjs && node --check scripts/export_historical_q1_position_reference.mjs`

Expected: all Node tests pass and syntax check exits 0.

- [ ] **Step 5: Commit**

```bash
git add scripts/historical_q1_reference.mjs scripts/export_historical_q1_position_reference.mjs tests/historical_q1_reference.test.mjs
git commit -m "feat: export historical Q1 position reference"
```

### Task 2: Resolve worksheet rows against historical evidence

**Files:**

- Create: `scripts/historical_q1_mapping.py`
- Create: `scripts/build_historical_q1_position_mapping.py`
- Create: `tests/test_historical_q1_mapping.py`

**Interfaces:**

- `build_mapping_rows(positions, historical_payload, existing_config, company_id='1') -> list[dict]` returns exactly one mapping report row per source-workbook/worksheet key.
- `historical_assignment_type(row) -> str` returns `PRIMARY`, `LAKHAR`, or `JOB_SHARING`.
- `validate_mapping_row(row) -> list[str]` rejects both candidate namespaces, a structural PNID, or a non-structural PMID.
- CLI: `python3 scripts/build_historical_q1_position_mapping.py --historical-reference <json> --config configs/pre_restructure_positions.json --existing-config configs/pre_restructure_positions_rw_reviewed_20260609.json --output-dir <dir>` writes `mapping_report.json`, `mapping_report.csv`, and `summary.json`.

- [ ] **Step 1: Write failing Python tests for historical mapping behavior**

```python
def test_structural_historical_assignment_proposes_pmid_only(self):
    rows = build_mapping_rows([worksheet('Group Head')], historical_reference(structural_row()), {}, '1')
    self.assertEqual(rows[0]['Candidate PMID'], '501')
    self.assertEqual(rows[0]['Candidate PNID'], '')

def test_non_structural_unique_cluster_proposes_pnid_only(self):
    rows = build_mapping_rows([worksheet('Officer Keuangan')], historical_reference(non_structural_row()), {}, '1')
    self.assertEqual(rows[0]['Candidate PMID'], '')
    self.assertEqual(rows[0]['Candidate PNID'], '76')

def test_ambiguous_pnid_stays_needs_check(self):
    rows = build_mapping_rows([worksheet('Officer Keuangan')], historical_reference(ambiguous_non_structural_row()), {}, '1')
    self.assertEqual(rows[0]['Confidence Label'], 'mapping_conflict')
```

- [ ] **Step 2: Run the focused Python test and verify it fails because the module is absent**

Run: `python3 -m unittest tests.test_historical_q1_mapping -v`

Expected: import failure for `historical_q1_mapping`.

- [ ] **Step 3: Implement the resolver and CLI**

The resolver must reuse `position_mapping.normalize_position_lookup`, aggregate worker numbers/names and assignment types, use only company `1` rows as mapping candidates, record missing historical organization evidence, preserve existing config IDs as comparison fields, and rank exact title/group evidence before weaker title matching. It must output `high_confidence`, `low_confidence`, `mapping_conflict`, or `no_candidate` and leave every row unapproved.

- [ ] **Step 4: Run focused and related regression tests**

Run: `python3 -m unittest tests.test_historical_q1_mapping tests.test_position_mapping tests.test_apply_position_mapping_review -v`

Expected: all tests pass.

- [ ] **Step 5: Commit**

```bash
git add scripts/historical_q1_mapping.py scripts/build_historical_q1_position_mapping.py tests/test_historical_q1_mapping.py
git commit -m "feat: map pre-restructure worksheets to historical Q1 positions"
```

### Task 3: Build and validate the editable review workbook

**Files:**

- Create: `outputs/019f6a65-722b-72f1-95ee-ec8aa577e55a/build_historical_q1_review_workbook.mjs`
- Create: `outputs/019f6a65-722b-72f1-95ee-ec8aa577e55a/historical_q1_mapping_inputs.json`
- Create: `outputs/019f6a65-722b-72f1-95ee-ec8aa577e55a/Historical_Q1_Position_Mapping_Review_20260717.xlsx`

**Interfaces:**

- Builder input is the Task 2 mapping report JSON plus the Task 1 historical reference JSON.
- Builder output has `Summary`, `Historical TEPMS`, `Position Mapping Report`, and `Review Queue` sheets.
- `Position Mapping Report` keeps the four reviewer fields editable. `Reviewer Confirm Mapping` has list validation for `YES` and `NEEDS_CHECK`.

- [ ] **Step 1: Write a failing builder-input assertion before workbook authoring**

```js
if (mappingRows.length !== 295) {
  throw new Error(`Expected 295 mapping rows, received ${mappingRows.length}`);
}
if (new Set(mappingRows.map((row) => `${row['Source Workbook']}\u0000${row.Worksheet}`)).size !== 295) {
  throw new Error('Mapping row keys must be unique.');
}
```

- [ ] **Step 2: Run the builder before inputs exist and verify it fails at the expected input gate**

Run: `node outputs/019f6a65-722b-72f1-95ee-ec8aa577e55a/build_historical_q1_review_workbook.mjs`

Expected: failure that names the missing mapping input file or the 295-row input assertion.

- [ ] **Step 3: Implement the artifact-tool workbook builder**

Use `Workbook.create()` and `SpreadsheetFile.exportXlsx`. Apply title/header hierarchy, frozen headers, bounded column widths, date format `yyyy-mm-dd`, autofilters, and meaningful conditional formatting for confidence/review status. Do not use `openpyxl` or another alternative library to author the workbook.

- [ ] **Step 4: Run artifact checks and visual verification**

Run the builder, then inspect all sheet names/key ranges, scan for formula errors, render each sheet, and visually inspect the rendered previews. Verify the workbook ZIP structure with `unzip -t`.

Expected: all four sheets are readable, all 295 mapping rows exist, reviewer fields are visible, no formula error scan results, and ZIP integrity passes.

- [ ] **Step 5: Commit only reusable source code, not generated output**

```bash
git add scripts/historical_q1_reference.mjs scripts/export_historical_q1_position_reference.mjs scripts/historical_q1_mapping.py scripts/build_historical_q1_position_mapping.py tests
git commit -m "feat: prepare historical Q1 mapping review workflow"
```

### Task 4: Execute the production read-only mapping run and stop at review

**Files:**

- Create: `outputs/019f6a65-722b-72f1-95ee-ec8aa577e55a/historical_q1_position_reference.json`
- Create: `outputs/019f6a65-722b-72f1-95ee-ec8aa577e55a/mapping_report.json`
- Create: `outputs/019f6a65-722b-72f1-95ee-ec8aa577e55a/mapping_report.csv`
- Create: `outputs/019f6a65-722b-72f1-95ee-ec8aa577e55a/summary.json`
- Create: `outputs/019f6a65-722b-72f1-95ee-ec8aa577e55a/Historical_Q1_Position_Mapping_Review_20260717.xlsx`

- [ ] **Step 1: Export historical production evidence**

Run: `node scripts/export_historical_q1_position_reference.mjs --profile production --cutoff-date 2026-03-31 --company-id 1 --output outputs/019f6a65-722b-72f1-95ee-ec8aa577e55a/historical_q1_position_reference.json`

Expected: a read-only JSON artifact with source metadata, historical assignment rows, and nomenclature rows; every assignment end date is 31 March 2026.

- [ ] **Step 2: Build mapping artifacts**

Run: `python3 scripts/build_historical_q1_position_mapping.py --historical-reference outputs/019f6a65-722b-72f1-95ee-ec8aa577e55a/historical_q1_position_reference.json --config configs/pre_restructure_positions.json --existing-config configs/pre_restructure_positions_rw_reviewed_20260609.json --output-dir outputs/019f6a65-722b-72f1-95ee-ec8aa577e55a`

Expected: a 295-row mapping report, CSV, and summary that reconcile counts.

- [ ] **Step 3: Build review workbook and verify it**

Run the Task 3 builder against the generated inputs. Inspect and render all sheets before exporting the final workbook.

- [ ] **Step 4: Update the execution plan and stop at the manual review gate**

Record mapping counts, verification output, the review workbook path, and any production-data anomalies in this plan's Progress, Surprises & Discoveries, Decision Log, and Outcomes & Retrospective. Do not generate upload forms.

## Validation

- `node --test tests/historical_q1_reference.test.mjs`
- `node --check scripts/export_historical_q1_position_reference.mjs`
- `python3 -m unittest tests.test_historical_q1_mapping tests.test_position_mapping tests.test_apply_position_mapping_review -v`
- `unzip -t outputs/019f6a65-722b-72f1-95ee-ec8aa577e55a/Historical_Q1_Position_Mapping_Review_20260717.xlsx`
- Artifact-tool inspection for key ranges and formula errors.
- Artifact-tool renders for `Summary`, `Historical TEPMS`, `Position Mapping Report`, and `Review Queue`.
- Direct JSON checks: 295 mapping rows, 295 unique mapping keys, no row with both candidate IDs, and every historical assignment end date on 2026-03-31.

## Progress

- [x] 2026-07-16: Design approved; cutoff set to 31 March 2026 and manual review gate accepted.
- [x] 2026-07-16: Design spec committed as `e2e3bdc`.
- [x] 2026-07-17: Focused pre-existing mapping tests passed: 17 tests, 0 failures.
- [x] Task 1: Build read-only historical-reference exporter. Completed in `fd18abb` and hardened in `2063deb`; re-review clean.
- [x] Task 2: Resolve worksheet rows against historical evidence. Completed in `4be78a6`, hardened through `bbb1a12`; final re-review clean.
- [x] Task 3: Build and validate editable review workbook.
- [x] Task 4: Run production export and prepare review artifact.
- [ ] Manual review: user approves/corrects mappings.
- [ ] Conversion: generate per-workbook upload forms only after manual approval.

- [x] 2026-07-17: Production read-only export wrote 321 TEPMS assignments and 1,686 nomenclature rows for company `1` using the exact `DATE(tepms.end_date) = '2026-03-31'` filter.
- [x] 2026-07-17: Mapping output validated: 295 rows, 295 unique worksheet keys, 99 high-confidence candidates, 14 low-confidence candidates, 182 no-candidate rows, no mixed PMID/PNID candidates, and no prefilled reviewer decision.
- [x] 2026-07-17: Review workbook generated at `outputs/019f6a65-722b-72f1-95ee-ec8aa577e55a/Historical_Q1_Position_Mapping_Review_20260717.xlsx`; four sheets, reviewer validation, artifact re-import, visual previews, and `unzip -t` all passed.

## Surprises & Discoveries

- `tb_employee_position_master_sync` contains a PMVID, not a PNID. PNID requires the non-temporal `position_nomenclature_mapping` bridge and must be reviewed when context is ambiguous.
- The existing production reference exporter is current-active only, so it cannot be reused as the historical Q1 source.
- The current checkout contains substantial user-owned dirty and untracked work. The implementation uses newly named historical files and the conversation output directory only.
- JSON serialization can render date-time values in UTC (including `2026-03-30` for an Indonesia-local 31 March timestamp). The source SQL remains the authority and filters with `DATE(tepms.end_date) = '2026-03-31'`; report evidence is displayed as the intended local date.

## Decision Log

- Decision: Use `DATE(tepms.end_date) = '2026-03-31'` rather than current-active or overlapping-date logic.
  Rationale: The user defined the pre-restructure Q1 cohort by assignments that ended on 31 March 2026.
  Date/Author: 2026-07-16 / Alfredo Teja and Codex.

- Decision: Keep non-confirmable historical organization assignments in raw evidence but exclude them from automatic Head Office candidates.
  Rationale: This preserves auditability without allowing records outside company `1` to map silently.
  Date/Author: 2026-07-17 / Codex.

- Decision: Pause before KPI conversion.
  Rationale: The user explicitly requires manual position-mapping review before conversion.
  Date/Author: 2026-07-16 / Alfredo Teja and Codex.

- Decision: Run the production export and JSON mapping before authoring the workbook.
  Rationale: The artifact builder requires the real, validated 295-row mapping input; this only reorders dependent milestones and does not advance conversion.
  Date/Author: 2026-07-17 / Codex.

## Outcomes & Retrospective

- Current state: the source-backed, editable workbook covers all 295 pre-restructure worksheet mappings and preserves 321 historical TEPMS assignment records as read-only evidence.
- Verification: mapping invariants, reviewer-field blankness, XLSX ZIP integrity, artifact re-import, and visual previews all passed.
- Stopping point reached: manual review is now required. No upload form, KPI conversion, or checked-in configuration was generated or modified.
