# Output Folder Naming Standard

Use this format for new conversion outputs:

```text
<group>_<scope>_<artifact>_<YYYYMMDD>_<qualifier>
```

Required parts:
- `group`: `group1`, `group2`, `group3`, or another explicit business group.
- `scope`: short business scope, for example `ho_v2`, `k3`, `pengendalian_proyek`.
- `artifact`: what the folder contains, for example `conversion`, `upload_ready`, `readiness_audit`, `delta_remediation`, `followup_upload`.
- `YYYYMMDD`: production/reference date or generation date.
- `qualifier`: source or status, for example `prod`, `staging`, `reviewed`, `latest_prod`.

Examples:
- `group1_ho_v2_conversion_20260709_latest_prod`
- `group1_ho_v2_upload_ready_20260709_prod`
- `group1_ho_v2_readiness_audit_20260710_prod`
- `group1_ho_v2_delta_remediation_20260709_prod`

Rules:
- Keep names lowercase snake_case.
- Put the group and scope first.
- Put the date near the end.
- Avoid generic names like `latest`, `final`, or `v2` unless paired with a date and scope.
- Do not mix multiple groups in one output folder.
- Keep large production snapshots inside the analysis folder that used them.

Current note:
- Group 2 and Group 3 conversion artifacts were moved out of `output/` during cleanup on 2026-07-10.
- Existing active Group 1 artifacts were left in place to avoid breaking current audit references.
