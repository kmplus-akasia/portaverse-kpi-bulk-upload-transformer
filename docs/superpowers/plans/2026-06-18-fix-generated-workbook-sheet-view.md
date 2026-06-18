# Fix Generated Workbook Sheet View Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Prevent generated KPI upload workbooks from triggering Excel's unreadable-content repair flow.

**Architecture:** Preserve the official template's valid frozen-pane metadata instead of partially clearing it. Protect the behavior with an XML-level regression test, then regenerate only the affected upload-ready outputs and validate their workbook structure.

**Tech Stack:** Python, openpyxl, unittest, OOXML ZIP/XML inspection

---

### Task 1: Add the failing sheet-view regression test

**Files:**
- Modify: `tests/test_kpi_bulk_transform.py`
- Test: `tests/test_kpi_bulk_transform.py`

- [ ] **Step 1: Add ZIP/XML imports and a regression test**

Add `zipfile` and `xml.etree.ElementTree` imports. Create a template with
`freeze_panes = "B2"`, call `write_output_workbook`, inspect
`xl/worksheets/sheet1.xml`, and assert that `<pane>` exists whenever a
`<selection>` has a `pane` attribute.

- [ ] **Step 2: Run the focused test and verify RED**

Run:

```bash
python3 -m unittest tests.test_kpi_bulk_transform.KpiBulkTransformTest.test_write_output_workbook_preserves_valid_frozen_pane_view -v
```

Expected: FAIL because the generated worksheet has pane-specific selections but
no `<pane>` element.

### Task 2: Preserve the template pane and verify GREEN

**Files:**
- Modify: `scripts/kpi_bulk_transform.py:1622`
- Test: `tests/test_kpi_bulk_transform.py`

- [ ] **Step 1: Remove the pane-clearing statement**

Delete only:

```python
worksheet.freeze_panes = None
```

Keep all data, style, autofilter, and workbook-save behavior unchanged.

- [ ] **Step 2: Run the focused test and verify GREEN**

Run the focused unittest command from Task 1. Expected: PASS.

- [ ] **Step 3: Run the full transformer suite**

Run:

```bash
python3 -m unittest tests/test_kpi_bulk_transform.py
```

Expected: all tests pass with zero failures and zero errors.

### Task 3: Regenerate and validate affected upload-ready workbooks

**Files:**
- Replace generated `.xlsx` files under `output/group1_ho_regenerated_20260615_final_v2/upload-ready/`
- Replace generated `.xlsx` files under `output/project_positions_upload_20260615/upload-ready/by-project/`

- [ ] **Step 1: Re-run the existing Head Office and project conversion commands**

Use the existing source workbooks, reviewed configs, production mapping snapshot,
and official template already recorded in the output manifests. Write results back
to the two affected upload-ready directories without changing filenames.

- [ ] **Step 2: Run the batch validator**

Run `scripts/validate_kpi_upload_batch.py` against both regenerated directories.
Expected: no new validation errors compared with the existing accepted outputs.

- [ ] **Step 3: Validate OOXML structure for every regenerated workbook**

For every non-lock `.xlsx`, verify ZIP CRC, parse all `.xml` and `.rels` members,
and assert that no worksheet has pane-specific selections without a `<pane>`.
Expected: 26/26 Head Office and 6/6 project workbooks pass.

- [ ] **Step 4: Open representative workbooks headlessly**

Use the installed LibreOffice `soffice --headless` conversion path on one Head
Office and one project workbook. Expected: both open and convert without repair or
format errors.

- [ ] **Step 5: Review the final diff and generated-file scope**

Confirm only the generator, regression test, plan/spec documentation, and affected
generated workbooks changed. Do not stage or modify unrelated user files.

