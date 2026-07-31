---
name: apply-position-identity-config
description: Write approved PMID/PNID identity into a KPI position config. Use when reviewer decisions in a mapping workbook are ready to apply, when override candidates have been approved, or when an existing config must be re-checked against a fresher production reference.
---

# Apply Identity to Config

Three sources of decision, one effect: the config's identity fields. Every run writes a new config and an audit of what changed.

| Source of decision | Script | Key input |
| --- | --- | --- |
| Reviewer, mapping workbook | `scripts/apply_position_mapping_review.py` | `Reviewer Confirm Mapping` plus `Reviewer Actual PMID`/`PNID` |
| Reviewer, override sheet | `scripts/apply_mapping_override_candidates.py` | `Approved` on sheet `Override Candidates` |
| Production reference drift | `scripts/fix_structural_scope_from_reference.py` | active reference rows |

Column meanings live in the mapping skill's `references/review-artifact-schema.md`.

## Steps

1. **Match keys first.** The join is `source_workbook` plus `sheet_name`. Run the dry run and read the unmatched list before writing anything.

   Done when: every decision row matches a config key, or the mismatches are reported as blockers and the run stops there.

2. **Apply one identity per position.** A structural decision sets PMID and clears PNID; a non-structural decision sets PNID and clears PMID. A `NEEDS_CHECK` or `HOLD` row sets `mapping_review_status` and keeps the position out of upload scope.

   Done when: no position in the output config holds a PMID and a PNID together, and every held position carries a status naming why it is held.

3. **Write to a new path.** The input config stays as it was, and the output is named for the run. Include a top-level `metadata.source_root` and `metadata.inventory_config` copied from `scripts/kamus_source.py` when the config will feed a Kamus conversion.

   Done when: the output config exists at a new path and the input file is byte-identical to before the run.

4. **Audit every change.** The reference-drift branch runs without a reviewer, so it carries the heavier burden: each changed value needs an audit row, and a change that no active reference row supports is a blocker rather than a silent correction.

   Done when: the audit file lists every changed position with its old and new identity, and the audit row count reconciles against the decisions applied.

## Report back

Respond in Indonesian with the branch used, counts of approved, held, and unmatched rows, the output config path, the audit path, and the identities that stay out of upload scope.
