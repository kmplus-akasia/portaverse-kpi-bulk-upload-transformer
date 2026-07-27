# Historical Q1 Position Mapping Rerun Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to execute this run task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Produce a fresh, editable Excel review workbook mapping every configured pre-restructure KPI worksheet to evidence-backed historical PMID/PNID candidates for Q1 2026.

**Architecture:** Reuse the already-tested read-only exporter and JSON resolver without changing source code. Write all run-specific evidence and the review workbook only to `outputs/historical-q1-pre-restructure-review-20260721/`; stop before KPI conversion.

**Tech Stack:** Existing Node.js production exporter, Python resolver, and `@oai/artifact-tool` workbook builder.

## Global Constraints

- Historical source rule is exactly `DATE(tb_employee_position_master_sync.end_date) = '2026-03-31'` and `deletedAt IS NULL`.
- Production access uses the existing `production` profile and read-only transaction only.
- Candidate scope is company ID `1`; PNID remains a nomenclature bridge rather than a historical TEPMS field.
- Structural candidates use PMID; non-structural candidates use PNID.
- The config source is `configs/pre_restructure_positions.json`; every configured worksheet must appear once.
- Reviewer columns remain blank. Only a later explicit reviewer `YES` may unlock conversion.
- Do not change raw workbooks, configs, existing review artifacts, or unrelated dirty worktree files.

---

### Task 1: Export and resolve the Q1 historical evidence

**Files:**
- Create: `outputs/historical-q1-pre-restructure-review-20260721/historical_q1_position_reference.json`
- Create: `outputs/historical-q1-pre-restructure-review-20260721/mapping_report.json`
- Create: `outputs/historical-q1-pre-restructure-review-20260721/mapping_report.csv`
- Create: `outputs/historical-q1-pre-restructure-review-20260721/summary.json`

- [x] **Step 1: Run the production read-only exporter**

Run:
```bash
node scripts/export_historical_q1_position_reference.mjs --profile production --cutoff-date 2026-03-31 --company-id 1 --output outputs/historical-q1-pre-restructure-review-20260721/historical_q1_position_reference.json
```

Expected: source metadata identifies production, company `1`, cutoff `2026-03-31`, and `read_only: true`.

- [x] **Step 2: Run the worksheet resolver**

Run:
```bash
python3 scripts/build_historical_q1_position_mapping.py --historical-reference outputs/historical-q1-pre-restructure-review-20260721/historical_q1_position_reference.json --config configs/pre_restructure_positions.json --existing-config configs/pre_restructure_positions_rw_reviewed_20260609.json --output-dir outputs/historical-q1-pre-restructure-review-20260721
```

Expected: one JSON row and one CSV row per configured `(Source Workbook, Worksheet)` key.

- [x] **Step 3: Validate mapping invariants**

Run a JSON check that requires 295 rows, 295 distinct keys, zero rows with both candidate IDs populated, and zero prefilled reviewer fields.

### Task 2: Build and verify the editable review workbook

**Files:**
- Create: `outputs/historical-q1-pre-restructure-review-20260721/Historical_Q1_Position_Mapping_Review_20260721.xlsx`

- [x] **Step 1: Author the workbook with artifact-tool**

Create the four sheets `Summary`, `Historical TEPMS`, `Position Mapping Report`, and `Review Queue`. Keep the four reviewer fields editable; validate `Reviewer Confirm Mapping` against `YES` and `NEEDS_CHECK`.

- [x] **Step 2: Verify the output**

Use artifact-tool to inspect key ranges and formula errors, render all four sheets for visual review, and run:
```bash
unzip -t outputs/historical-q1-pre-restructure-review-20260721/Historical_Q1_Position_Mapping_Review_20260721.xlsx
```

- [x] **Step 3: Stop at manual review**

Report the workbook path and mapping counts. Do not generate KPI upload forms.

## Execution Record

- 2026-07-21: Production read-only exporter returned 322 historical TEPMS assignments and 1,686 nomenclature rows for company ID `1` with cutoff `2026-03-31`.
- 2026-07-21: Resolver returned 295 mapping rows and 295 unique worksheet keys: 99 high confidence, 14 low confidence, and 182 no candidate. Candidate namespaces were mutually exclusive and every reviewer field was blank.
- 2026-07-21: Final workbook `outputs/historical-q1-pre-restructure-review-20260721/Historical_Q1_Position_Mapping_Review_20260721.xlsx` passed artifact inspection, formula-error scan, four-sheet visual review, independent task review, and `unzip -t`.
- 2026-07-21: The system `python3` interpreter stalled while importing standard library modules under overlapping resolver launches. A single rerun through the bundled Python runtime succeeded without source or config changes.
- Manual-review gate is active. No KPI upload form or conversion output was generated.
