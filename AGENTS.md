# Portaverse KPI Bulk Upload Transformer

## Commands

Install runtime:

```sh
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install -r requirements.txt
```

Run focused converter checks:

```sh
python3 -m unittest \
  tests.test_position_mapping \
  tests.test_kpi_bulk_transform \
  tests.test_mapping_override_candidates \
  tests.test_apply_mapping_override_candidates \
  tests.test_position_scope_pipeline -v

python3 -m py_compile \
  scripts/position_mapping.py \
  scripts/kpi_bulk_transform.py \
  scripts/validate_kpi_upload_batch.py
```

Run historical-Q1 checks:

```sh
python3 -m unittest \
  tests.test_historical_q1_mapping \
  tests.test_position_mapping \
  tests.test_apply_position_mapping_review -v
```

For a generated workbook, also verify its archive:

```sh
unzip -t "<output-workbook>.xlsx"
```

Do not treat the full `unittest discover` suite as the default verification path: dashboard tests can require optional `streamlit`.

## Repo map

- `scripts/kpi_bulk_transform.py` — main workbook-to-upload-template converter.
- `scripts/position_mapping.py` — strict position resolver and lookup indexes.
- `scripts/validate_kpi_upload_batch.py` — batch-level fail-closed validation.
- `scripts/export_position_reference.mjs` — read-only production-reference export.
- `scripts/historical_q1_mapping.py` and `scripts/export_historical_q1_position_reference.mjs` — historical Q1 mapping workflow.
- `tests/` — regression tests; add focused coverage with behavioral changes.
- `configs/` — mapping/config snapshots. Treat production-reference files as sensitive.
- `input/` — source workbooks and templates.
- `output/` and `outputs/` — generated artifacts and evidence; never mass-delete or overwrite prior runs.

## Source of truth and identity rules

- Start every conversion by identifying the raw workbook, exact worksheet scope, template version, config, and production-reference snapshot.
- Treat `configs/production_position_reference.json` as an offline snapshot, not timeless truth. Its `current_snapshot_unreviewed` status requires review before identity decisions.
- Use the strict resolver through `build_lookup_indexes(...)`; do not reintroduce obsolete lookup paths.
- Determine worksheet scope before lookup:
  - structural position -> `Position Master ID` / PMID only;
  - non-structural position -> `Position Nomenklatur ID` / PNID (`cluster_id`) only;
  - uncertain, conflicting, low-confidence, or missing mapping -> block conversion and create a review artifact.
- Never populate PMID and PNID together for the same upload identity.
- Do not infer position identity from title similarity alone when active production evidence, worksheet evidence, or historical TEPMS evidence is required.

## Historical Q1 workflow

For pre-restructure Q1 work, preserve this mandatory gate:

```text
TEPMS historical identity -> worksheet mapping -> editable review artifact
-> explicit approval -> conversion -> validation
```

- Use the specified historical cutoff; for the established Q1 workflow it is `2026-03-31`.
- Current active reference exports cannot recreate historical assignments.
- Keep `HOLD`, missing, ambiguous, and unapproved identities out of upload-ready output.
- Preserve raw workbook/worksheet provenance and reviewer decisions in the delivered artifact.

## Production and data boundaries

Always:

- Use production access read-only; require `DB_READ_WRITE=0` where the exporter supports it.
- Preserve timestamps, source paths, identity evidence, validation receipts, and skipped records.
- Report unavailable credentials or dependencies as a blocker; never imply that production data was verified when it was not.
- Treat employee names, NIPP, and production-reference snapshots as sensitive operational data.

Ask first:

- Refreshing production data when it may require credentials or a new environment.
- Bulk conversion or any action that produces an upload-ready package from unresolved mappings.
- Changes to input templates, importer contract, config schemas, or production DB queries.
- Deleting generated artifacts, replacing historical outputs, or committing sensitive reference data.

Never:

- Edit secrets or `.env` files.
- Write to production databases.
- Bypass mapping confidence gates, convert `HOLD` identities, or silently fill uncertain PMID/PNID values.
- Claim an upload package is ready until structural, formula, identity, and XLSX-integrity checks pass.

## Working conventions

- This worktree is already dirty. Preserve unrelated user changes; do not clean, reset, reformat, stage, or overwrite them.
- For conversion/remediation tasks, classify scope first: upload, delete-only, sync, triage, verification, or mapping review.
- Prefer an editable Excel/CSV review artifact over chat-only mapping decisions.
- Make generated outputs run-scoped and retain validation evidence beside them.
- Before reporting completion, review every changed file and state checks run, skipped checks, blocked identities, and remaining approval gates.

## Agent skills

### KPI upload workflow

Repo-local skills live in `.cursor/skills/`, one folder per skill, and are symlinked into `~/.codex/skills/`. Classify the request's scope, then open the matching skill; `kpi-upload-router` holds the routing table.

| Scope | Skill |
| --- | --- |
| Worksheet inventory from a raw download | `discover-kamus-worksheet-config` |
| Worksheet-to-identity mapping, including historical periods | `position-mapping-review` |
| Writing approved PMID/PNID into a config | `apply-position-identity-config` |
| One formulir for named positions | `generate-position-upload` |
| Whole group, folder, or ZIP | `convert-kpi-upload-batch` |
| Changing a formulir already delivered | `amend-kpi-upload-form` |
| Granting the `upload-ready` claim | `validate-upload-package` |
| Fresh read-only production snapshot | `refresh-production-reference` |

### Issue tracker

Issues and PRDs live in this repository's GitHub Issues. See `docs/agents/issue-tracker.md`.

### Triage labels

Uses the five default canonical triage labels. See `docs/agents/triage-labels.md`.

### Domain docs

Single-context layout. See `docs/agents/domain.md`.
