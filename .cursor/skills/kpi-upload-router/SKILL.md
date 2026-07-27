---
name: kpi-upload-router
description: Index of the Portaverse KPI upload skills and when to reach for each.
disable-model-invocation: true
---

# KPI Upload Router

Classify the request's scope first, then open the matching skill.

| Scope | The request sounds like | Skill |
| --- | --- | --- |
| Inventory | a fresh Kamus KPI download arrived; which worksheets does it hold | `discover-kamus-worksheet-config` |
| Mapping | worksheets need PMID/PNID candidates; a conflict or low-confidence match; a historical period | `position-mapping-review` |
| Apply | reviewer decisions are ready; overrides approved; the reference drifted | `apply-position-identity-config` |
| Formulir | one named position, or one consolidated form covering several identities | `generate-position-upload` |
| Batch | a whole group, folder, or ZIP converted at once | `convert-kpi-upload-batch` |
| Amend | add, remove, or replace items on a form already delivered; patch identity | `amend-kpi-upload-form` |
| Verify | is this package upload-ready; re-check an older run | `validate-upload-package` |
| Reference | the production snapshot is stale | `refresh-production-reference` |
| Audit | which production positions still lack a KPI dictionary | `update-org-kpi-audit-report` |

The usual sequence is inventory, then mapping, then apply, then batch or formulir, then verify. Amend enters after a formulir has already been delivered.

Two scopes end without an upload artifact by design: mapping stops at the reviewer, and audit reports coverage. Reaching those means the answer is a review artifact or a report, and the run is complete there.
