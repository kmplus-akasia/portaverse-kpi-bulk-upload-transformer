# Group 1 HO Issue Remediation Upload Design

## Objective

Produce one consolidated, complete KPI upload workbook for issues 008, 009,
018, 019, 020, 021, 025, and 031. Issue 011 is excluded because its weight
allocation will be handled manually by users in the KPI allocation workflow.

## Sources of Truth

- Raw source: latest third download under
  `/Users/alfredoteja/Downloads/KAMUS KPI PELINDO GROUP 1 (HO) 3`.
- Identity and organization reference: the fresh production reference exported
  read-only on 2026-07-16.
- Current KPI comparison: the fresh 2026 production KPI snapshot exported
  read-only on 2026-07-16.
- Upload schema: `input/KPI Upload Template.xlsx` and the active importer in
  `pms-service/src/modules/performance-hq/services/kpi-template-import.service.ts`.

## Packaging

- Deliver exactly one consolidated `.xlsx` workbook.
- Preserve the official upload-sheet structure and required master-data sheets.
- Multiple structural and non-structural identities may coexist in the same
  upload sheet, provided every row has exactly one valid ownership identity.
- Issue 011 rows must not be included.

## Remediation Rules

### Issue 008

Include `Standarisasi Fasilitas Pendukung Terminal Penumpang dan Kendaraan
Roro` with its intended parent Impact, weight, and KAI. Use the complete raw KPI
definition while resolving the blank-weight/transfer ambiguity into an
upload-ready row set.

### Issue 009

Include both required Outputs:

- `Implementasi Pengelolaan Alur Pelayaran di Pelindo Group`
- `Delta Pengurangan Kecelakaan Kerja Eksternal`

Each Output must have one valid child KAI. Use the raw Delta KAI where valid and
create the missing Alur Pelayaran child from the available issue/source context,
keeping its title and definition explicit and auditable.

### Issue 018

Include the raw Group Head Keberlanjutan Korporasi `Net Income` Output and its
KAI under the correct `Net Income` Impact.

### Issue 019

Set `Manpower productivity (Revenue per total manpower)` to 5% and retain the
corporate Impact set so total Impact weight equals 100%. Keep valid non-drop
Output/KAI rows from the source.

### Issue 020

Remove every source item marked `Drop`. For the remaining non-drop Output rows,
normalize the existing weights proportionally so Output totals exactly 100%.
Each KAI inherits the normalized weight of its parent Output, so KAI also totals
exactly 100%.

### Issue 021

Keep only non-drop items. Preserve weights already populated. Calculate the
remaining weight to 100% and distribute it equally across non-drop items whose
weight is blank. Each KAI inherits its parent Output weight. Use a deterministic
residual adjustment on the final allocated item if decimal rounding is needed.

### Issue 025

Retain the seven valid Output rows from the Manager Pengelolaan Aset source.
Create a distinct KAI row for every parent Output. Identical KAI titles may be
repeated when they belong to different parent Output IDs; they must not be
deduplicated across parents.

### Issue 031

For `Digitalisasi Pengelolaan Keuangan`, include only the KAI
`Percentage progres pengembangan auto generate nota (100%)`. Exclude the bond
and loan monitoring and locking-system indicators from this parent.

## Identity and Parent Rules

- Structural positions use PMID; non-structural positions use PNID.
- Do not populate both identity namespaces on one logical ownership row.
- OUTPUT rows must reference an IMPACT within the same identity.
- KAI rows must reference an OUTPUT within the same identity.
- Parent IDs and titles must be internally consistent and importer-readable.

## Workbook Construction

Use a hybrid source-backed method:

1. Extract complete candidate rows from the raw worksheets.
2. Resolve identities against the fresh production reference.
3. Apply only the issue-specific transformations defined above.
4. Populate the official template without redesigning the upload sheet.
5. Preserve required workbook formatting and master-data tabs.

An audit/receipt sheet may be added only if the importer tolerates extra sheets.
The receipt must identify the issue, identity, source worksheet, row counts,
weight totals, and applied remediation without changing upload semantics.

## Validation and Acceptance Criteria

- One final consolidated workbook exists and opens successfully.
- Issue 011 is absent.
- All eight requested issues are represented.
- Every target identity has the complete 10-row corporate Impact set.
- Impact, Output, and KAI totals are each exactly 100% per target identity.
- No included row is marked `Drop`.
- Every Output and KAI has a valid same-identity parent.
- Issue 025 has one KAI ownership row per Output parent, including repeated KAI
  titles where required.
- Issue 031 has exactly one allowed KAI under Digitalisasi Pengelolaan Keuangan.
- Identity values agree with the fresh production reference.
- The workbook passes the repo batch validator or the closest importer-contract
  validation available.
- Formula-error scan and visual inspection of all workbook sheets pass.

## Safety

- Production access is read-only; no production mutation is performed.
- Raw source workbooks and the official template are not overwritten.
- The final workbook is written under a new conversation-specific `outputs/`
  directory.
