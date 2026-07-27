---
name: position-mapping-review
description: Gate worksheet-to-position identity behind an editable review artifact. Use when worksheets need PMID/PNID candidates, when a mapping conflict or low-confidence match appears, when a pre-restructure or historical period must be resolved, or when a named subset of identities needs a mapping workbook.
---

# Position Mapping Gate

Produce an editable review artifact and stop there. Conversion resumes only after a reviewer fills the decision columns and `apply-position-identity-config` writes them into a config.

## Steps

1. **Name the reference snapshot.** Record its path, its export timestamp, and whether it is an *active* production export or a *historical* export taken at a cutoff. An active export describes today's organisation, so a request about a past period needs the historical branch.

   Done when: snapshot kind, path, and export timestamp sit in the artifact metadata, and the kind matches the period the request is about.

2. **Choose the branch.**
   - *Current identities* — resolve against the active reference with the strict resolver.
   - *Historical period* — follow `references/historical-q1-branch.md`.
   - *Named subset* — carry the subset's source list, whether that is an audit sheet, a worksheet inventory, or a list of PNIDs, and resolve only those keys.

   Done when: the candidate row count equals the count of worksheet keys the branch was handed.

3. **Settle scope before lookup.** Infer structural or non-structural per worksheet, then search only the matching lookup index built through `build_lookup_indexes(...)`. Title resemblance sets a candidate's score, while the evidence that justifies it comes from the active reference, the worksheet itself, or historical TEPMS rows.

   Done when: no row carries both a candidate PMID and a candidate PNID, and every row holds one confidence label with a stated reason.

4. **Emit the review artifact** as an Excel workbook in a run-scoped folder under `outputs/`, reviewer columns present and empty. Column set and confidence labels are in `references/review-artifact-schema.md`.

   Done when: every worksheet key appears exactly once, every reviewer field is blank, and the confidence distribution is counted.

5. **Stop at the gate.** The run's deliverable is the artifact plus its distribution summary, and the reviewer decides what happens next.

   Done when: the reviewer holds the artifact path, the distribution, and the list of columns to fill.

## Report back

Respond in Indonesian with the reference snapshot kind and export timestamp, the row count, the confidence distribution, the artifact path, and the decision columns the reviewer must complete.
