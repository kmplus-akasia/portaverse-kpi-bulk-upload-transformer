# Review artifact schema

## Reviewer decision columns

These four stay blank when the artifact is produced. `apply-position-identity-config` reads them back.

| Column | Accepted values |
| --- | --- |
| `Reviewer Confirm Mapping` | `YES`, `NEEDS_CHECK` |
| `Reviewer Actual PMID` | a Position Master ID, when the identity is structural |
| `Reviewer Actual PNID` | a Position Nomenclature ID (`cluster_id`), when the identity is non-structural |
| `Reviewer Notes` | free text |

A row fills either `Reviewer Actual PMID` or `Reviewer Actual PNID`. Filling both is rejected at apply time as an identity conflict.

Leaving both blank on a `YES` row means the reviewer accepts the candidate the resolver proposed, and apply takes the candidate matching the inferred scope.

## Confidence labels

| Label | Meaning | Reaches upload |
| --- | --- | --- |
| `high_confidence` | evidence-backed single candidate | after reviewer `YES` |
| `low_confidence` | a candidate exists, evidence is thin | after reviewer `YES` |
| `mapping_conflict` | more than one defensible candidate | after the reviewer names one |
| `scope_uncertain` | structural versus non-structural undecided | after the reviewer names the scope |
| `no_candidate` | no eligible row in the lookup | stays out until an identity exists |

## Candidate columns

The conflict-review builder writes: `Candidate Rank`, `Candidate Score`, `Candidate Scope`, `Candidate PMID`, `Candidate PNID`, `Candidate Title`, `Candidate Group`, `Candidate Company`, `Candidate Company Code`, `Candidate Source`, `Match Reason`, plus the raw side (`Raw Group`, `Raw Position`, `Direktorat`) and `Confidence Reason`.

Scoring weights title at 0.70, group at 0.20, and company at 0.10.

## Override candidates sheet

`scripts/build_mapping_override_candidates.py` produces a second artifact shape: sheet `Override Candidates`, keyed on `Source Workbook` + `Worksheet`, with an `Approved` column accepting `YES` or `NO`. `scripts/apply_mapping_override_candidates.py` treats `1`, `true`, `yes`, `y`, `approved`, `approve`, and `ok` as approval.

## Historical report columns

The historical branch writes its own column set, defined as `REPORT_COLUMNS` in `scripts/historical_q1_mapping.py`. It adds the evidence fields absent from the current-reference artifact: `Historical Employee Numbers`, `Historical Employee Names`, `Assignment Types`, `Historical End Date`, `Missing Historical Organization Evidence`, `Existing Config PMID`, and `Existing Config PNID`.
