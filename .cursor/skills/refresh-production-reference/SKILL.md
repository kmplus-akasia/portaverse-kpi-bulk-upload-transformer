---
name: refresh-production-reference
description: Export a fresh read-only production position reference snapshot.
disable-model-invocation: true
---

# Refresh the Production Reference

Production access here is read-only. Set `DB_READ_WRITE=0` where the exporter supports it, and leave every `.env` file as it is.

## Steps

1. **Confirm the profile and credentials resolve.** Missing credentials are a blocker to report, and the run stops there rather than continuing against an older snapshot.

   Done when: the profile is named and its connection succeeds, or the blocker is reported with the missing variable.

2. **Export to a run-scoped path**, leaving the previous snapshot in place.

```bash
DB_READ_WRITE=0 node scripts/export_position_reference.mjs \
  --profile production \
  --output <run-scoped>/production_position_reference_<date>.json
```

   Done when: the new file exists alongside the old one and both remain readable.

3. **Record provenance and review status.** Note the export timestamp, the profile, the database, and the row counts per section: `rows`, `position_master_rows`, `organization_rows`, `company_rows`.

   Done when: those five facts are written down, and the snapshot is marked unreviewed until someone confirms it, so identity decisions cite the timestamp they relied on.

## Sensitivity

The export carries employee names and NIPP. Treat it as sensitive operational data and ask before committing it to the repository.

## Report back

Respond in Indonesian with the snapshot path, profile, export timestamp, row counts per section, and its review status.
