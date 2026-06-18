# Finalize Position Scope and Regeneration Design

## Problem

The 15 June scope postprocessor treated any configured PNID that also existed as a
`position_master_id` as structural. That inference is invalid because every
non-structural position also has an internal PMID. The rule converted 63 configs;
54 values were already valid PNIDs and only 9 were invalid as PNIDs and valid as
structural PMIDs. Thirteen additional configs had no reviewed identity and needed
an explicit `neglect` scope.

The 18 June workbook-view regeneration reused that incorrect config, so it fixed
Excel metadata but preserved wrong PMID/PNID ownership.

## Source of Truth

`configs/production_position_reference.json` is authoritative for this offline
run. `position_master_type_id == 5` means structural. Every other active type is
non-structural. PNID is `rows[].cluster_id`, never a raw mapping-row ID or an
internal PMID.

## Chosen Design

Replace collision-based scope inference in
`scripts/fix_structural_scope_from_reference.py` with type-driven resolution:

1. A valid configured PNID wins even when the same number exists as a PMID.
2. Resolve an invalid PNID against `position_master_rows`.
3. If its production type is `5`, output PMID only.
4. Otherwise, resolve all active nomenclature rows for that PMID. Require exactly
   one unique `cluster_id`, then output that PNID only.
5. Fail visibly when no master, mixed master types, no PNID, or multiple PNIDs are
   found. Never guess from title or numeric collision.
6. Convert blank, unresolved identities to explicit `neglect` and skip them before
   metadata fallback.
7. Record before/after scope and identity in the audit CSV.

The batch validator will independently verify config identities against production
type and PMID-to-PNID relationships. This prevents a future postprocessor defect
from passing validation merely because an ID exists in both namespaces.

## Regeneration and Cleanup

Regenerate all 26 Head Office workbooks from the original ZIP using the corrected
reviewed config and production snapshot. Build a new final directory under
`output/`, validate reports, config identities, workbook identities, ZIP/XML
integrity, and the two reported examples. Only after validation succeeds, delete
all other contents of `output/`.

The retained final directory contains upload-ready workbooks, one upload ZIP,
the corrected config, correction audit, upload manifest, upload instructions, and
validation receipt. Intermediate per-workbook outputs may remain inside this one
final directory only when required to reproduce the manifest.

## Acceptance Criteria

- Regression tests fail on the old collision rule and pass on the type-driven rule.
- `Officer Transaksi dan Proses` remains PNID `44` only.
- `Officer QA` remains PNID `11517` only.
- All active non-structural output rows use PNID; all structural rows use PMID.
- 26/26 workbooks pass batch, schema, report, ZIP CRC, and OOXML validation.
- `output/` contains only the final regenerated batch directory.
