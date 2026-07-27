# Historical period branch

A pre-restructure worksheet describes a position as it existed before the reorganisation. Its identity comes from TEPMS assignment rows retained at a cutoff, so this branch swaps the reference source and the resolver while keeping the rest of the gate identical.

For the established Q1 workflow the cutoff is `2026-03-31` and the company is Head Office, `1`.

## Export the historical reference

Read-only, and it needs credentials:

```bash
node scripts/export_historical_q1_position_reference.mjs \
  --profile production \
  --cutoff-date 2026-03-31 \
  --company-id 1 \
  --output <run-scoped>/historical_q1_position_reference.json
```

The payload's `source` block must show profile `production`, company `1`, the cutoff, and `read_only: true`. Confirm those four before resolving anything.

## Resolve

```bash
python3 scripts/build_historical_q1_position_mapping.py \
  --historical-reference <run-scoped>/historical_q1_position_reference.json \
  --config configs/pre_restructure_positions.json \
  --existing-config configs/pre_restructure_positions_rw_reviewed_20260609.json \
  --output-dir <run-scoped>
```

Outputs are `mapping_report.json`, `mapping_report.csv`, and `summary.json`.

The resolver keeps only company `1` rows as candidates, aggregates worker numbers, names, and assignment types (including `LAKHAR`), records missing historical organisation evidence rather than skipping it, and carries the existing config IDs as comparison fields rather than as answers.

## Baseline

Run 2026-07-21 over `configs/pre_restructure_positions.json`: 321 historical assignments and 1,686 nomenclature rows exported; 295 worksheet keys resolved into 99 high confidence, 14 low confidence, and 182 no candidate. Candidate namespaces were mutually exclusive and every reviewer field was blank.

The large no-candidate share is expected here: many pre-restructure worksheets describe positions that no longer have a retained assignment at the cutoff. Those rows belong in the artifact with their evidence gap stated, which is how the reviewer learns which identities need a manual decision.
