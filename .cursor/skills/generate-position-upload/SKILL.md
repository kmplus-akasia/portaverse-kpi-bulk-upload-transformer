---
name: generate-position-upload
description: Build one formulir upload KPI for named positions. Use when the user asks for a formulir or upload form for a specific position, a kamus KPI form for one role, one consolidated form covering several identities, or a recap of affected identities.
---

# Formulir Upload for Named Positions

One formulir, built from the official template, carrying only identities that are already resolved.

## Steps

1. **Resolve each requested position to an identity** from an approved config, a reviewed mapping artifact, or the active production reference. The three valid shapes are in `references/identity-shapes.md`. A position whose identity rests only on a title resembling a production position goes to `position-mapping-review` instead.

   Done when: each requested position holds exactly one identity shape, and any position without one is on the blocker list before a single KPI row is read.

2. **Locate the source worksheet** for each identity, recording workbook, sheet, and how that pair was established.

   Done when: every convertible identity names its source worksheet, and identities whose worksheet is missing join the blocker list.

3. **Parse and transform** through the repo's rules in `scripts/kpi_bulk_transform.py`, so enum normalisation, drop rules, and shared-impact backfill match what batch conversion produces.

   Done when: parsed rows preserve the IMPACT to OUTPUT to KAI hierarchy within each identity, and weight totals per level are known.

4. **Assemble against the official template.** For a consolidated formulir covering several identities, renumber `IDKPI` as one global `1..N` sequence and regenerate every `Parent KPI ID` after the merge, since per-identity numbering collides once merged.

   Done when: `IDKPI` is unique and sequential across the whole `KPI Template` sheet, and every OUTPUT and KAI resolves to a parent inside the same file.

5. **Verify** through `validate-upload-package`.

   Done when: that skill reports zero errors and the receipt exists beside the formulir.

6. **Recap the identities.** Count converted identities by shape and list every blocked one with its reason.

   Done when: converted plus blocked equals the total requested, with no identity unaccounted for.

## Report back

Respond in Indonesian with the formulir path, the identity recap by shape, the blocked identities and their reasons, and the validation result.
