---
name: refresh-production-reference
description: Export a fresh read-only production position reference snapshot to the canonical project path.
disable-model-invocation: true
---

# Refresh the Production Reference

Production access here is read-only. Set `DB_READ_WRITE=0` where the exporter supports it, and leave every `.env` file as it is.

## Canonical path

There is **one** production snapshot for the whole project:

| Artifact | Path |
| --- | --- |
| Snapshot JSON | `configs/production_position_reference.json` |
| Metadata | `configs/production_position_reference.meta.json` |
| Receipt | `outputs/production-reference/REFERENCE_RECEIPT.md` |

Do not export duplicate snapshots into run folders. Runs cite `exported_at` from the metadata file.

## Steps

1. **Confirm credentials resolve.** Missing credentials are a blocker; stop and report rather than reusing an older snapshot.

   Done when: the production profile connects, or the blocker is named.

2. **Refresh the canonical snapshot.**

```bash
DB_READ_WRITE=0 ./scripts/refresh_canonical_production_reference.sh
```

   Done when: `configs/production_position_reference.json` and `.meta.json` exist with a new `exported_at`.

3. **Record provenance.** The refresh script updates metadata and the receipt with export timestamp, profile, database, and row counts.

   Done when: `configs/production_position_reference.meta.json` reflects the new export and review status remains `current_snapshot_unreviewed` until a data owner confirms.

## Sensitivity

The export carries employee names and NIPP. Treat it as sensitive operational data and ask before committing it to the repository.

## Report back

Respond in Indonesian with the canonical path, export timestamp, row counts per section, and review status.
