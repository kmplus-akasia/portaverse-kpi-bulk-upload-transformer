# Fix Generated Workbook Sheet View

## Problem

Generated KPI upload workbooks trigger Excel's unreadable-content repair flow. The
official template contains a valid frozen pane and pane-specific selections. The
generator removes the frozen pane with `worksheet.freeze_panes = None`, but the
saved worksheet XML retains the pane-specific selections. Excel rejects that
orphaned `sheetView` state.

## Design

Preserve the official template's frozen pane by removing the generator statement
that clears it. This is the smallest change and keeps the template's view metadata
internally consistent.

Do not change KPI rows, formulas, styles, importer fields, or workbook naming.

## Regression Test

Generate a workbook through `write_output_workbook`, inspect its worksheet XML,
and verify:

- a `<pane>` element remains present when pane-specific selections exist;
- the generated workbook can still be loaded by the workbook validation path;
- generated KPI values remain unchanged.

The test must fail against the current implementation before the production code
is changed.

## Regeneration

Regenerate the affected Head Office and project-position upload-ready workbooks
from their existing configs and source files. Do not modify unrelated outputs.

## Verification

- Run the focused regression test and the existing transformer test suite.
- Run the existing batch validator on regenerated upload-ready workbooks.
- Inspect every regenerated workbook for ZIP/XML integrity.
- Confirm no worksheet contains pane-specific selections without a `<pane>`.
- Open a representative workbook with a real spreadsheet application when
  practical.

