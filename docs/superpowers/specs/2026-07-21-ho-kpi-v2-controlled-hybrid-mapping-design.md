# Controlled Hybrid Mapping for 34 Head Office KPI Identities

## Objective

Produce an auditable mapping review for the 34 unresolved Head Office position identities before any KPI dictionary conversion or upload-form generation.

The scope is:

- 33 identities classified as `Kamus KPI Tidak Tersedia` in the latest organization/KPI audit.
- 1 identity classified as `KPI Parsial - Perlu Review` (PNID 12474).
- One output row per unique `(Jenis Identity, ID Identity)`.

This phase stops after the mapping workbook is produced and verified. It must not convert KPI rows, generate upload forms, or change production data.

## Authoritative Inputs

1. Target identities and current status:
   `/Users/alfredoteja/Downloads/Laporan Audit Organisasi dan KPI - Kantor Pusat - Terbaru.xlsx`, sheet `Audit Posisi`.
2. Latest V2 visible-worksheet inventory:
   `configs/temp_visible_kamus_kpi_ho_latest_20260720.json`.
   Its metadata identifies source root `/Users/alfredoteja/Downloads/KAMUS KPI PELINDO GROUP 1 (HO) 4`, generated at `2026-07-20T17:04:10+07:00`.
3. Production identity reference:
   use the newest validated production reference available in the workspace and record its exact path and export timestamp in the output workbook.
4. Historical evidence:
   previously reviewer-approved mappings and mapping audit workbooks may be used only as supporting evidence. They never override an incompatible current identity or a missing V2 worksheet.

## Controlled Hybrid Strategy

### Stage 1 — Build the V2 candidate catalog

Adapt `kamus_kpi_v2[]` into the mapping resolver's `positions[]` shape without changing the source inventory. Include only rows where:

- `visibility = visible`;
- `include_in_position_config = true`;
- `review_status = ready`;
- workbook and worksheet names are present.

Retain source workbook, worksheet, extracted position name, group name, sheet order, extraction method, and source-config metadata.

### Stage 2 — Resolve strict current candidates

For each of the 34 target identities, rank V2 worksheets using current position and unit context. Reuse normalization and scoring concepts from `scripts/position_mapping.py`:

- title similarity;
- group/unit similarity;
- structural versus non-structural namespace compatibility;
- current production identity activity;
- runner-up gap and conflicts.

Do not use `--best-effort-mapping`. A candidate must not be automatically accepted when scope is uncertain, the current identity is inactive, or runner-up competition is material.

### Stage 3 — Reconcile historical approved evidence

Search prior approved mapping configs and review artifacts by exact source workbook/worksheet and by target PMID/PNID. Historical evidence may upgrade confidence only when:

- the same V2 workbook and worksheet still exist;
- the current target identity is active and matches the correct namespace;
- the historical mapping was explicitly reviewer-approved or has an equivalent trusted provenance;
- no competing current V2 candidate has materially similar evidence.

If any condition fails, keep the row in manual review status and show the historical mapping as evidence only.

### Stage 4 — Assign mapping status

Every target receives exactly one of these statuses:

- `Siap Dipetakan`: one current V2 candidate is supported by strict high confidence or a current-compatible trusted historical approval.
- `Perlu Konfirmasi`: a plausible candidate exists but confidence, namespace, or runner-up evidence is insufficient.
- `Tidak Ada Kandidat`: no eligible V2 worksheet is defensible.

No status in this phase authorizes conversion.

## Mapping Precedence

1. Exact current identity-compatible V2 mapping.
2. Strict high-confidence title and unit/group match.
3. Trusted historical approval validated against current V2 inventory and current production identity.
4. Lower-confidence candidate presented for manual confirmation.
5. No candidate; fail closed.

An incompatible historical mapping must never outrank current namespace or active-identity validation.

## Output Workbook

Create one simple worksheet named `Review Pemetaan 34 Posisi` with a summary band and a filterable mapping table.

Required columns:

### Target identity

- No.
- Jenis Identity
- PMID
- PNID
- Nama Posisi Production
- Nomenclature Production
- Unit Organisasi Production
- Jumlah Pekerja Aktif
- NIPP Unik Terdampak
- Status Audit Awal

### Recommended V2 mapping

- Workbook Kamus V2
- Worksheet Kamus V2
- Posisi pada Worksheet
- Group/Unit pada Kamus
- Metode Pemetaan
- Skor Kandidat
- Kandidat Runner-up
- Skor Runner-up
- Bukti Historis
- Status Pemetaan
- Tingkat Keyakinan
- Alasan/Rekonsiliasi

### Reviewer input

- Konfirmasi Reviewer (`Setuju`, `Perlu Koreksi`, `Tolak`)
- Workbook Hasil Koreksi
- Worksheet Hasil Koreksi
- Catatan Reviewer
- Reviewer
- Tanggal Review

Reviewer-input columns use a distinct editable color and data validation for the confirmation status.

## Data-Quality Rules

- Target count must equal 34.
- `(Jenis Identity, ID Identity)` must be unique and nonblank.
- Report worker impact using distinct NIPP within each identity; retain the source assignment count separately if it differs.
- Each accepted mapping must reference an existing eligible V2 workbook/worksheet pair.
- Structural/non-structural namespace must be consistent with the active production identity.
- Candidate and runner-up must never be the same workbook/worksheet pair.
- `Siap Dipetakan` requires nonblank workbook, worksheet, method, confidence, and evidence.
- `Tidak Ada Kandidat` must include a reason.
- No conversion or upload artifact may be created by the mapping command.

## Verification

Before delivery:

1. Reconcile 34 output rows to the 34 target identities.
2. Confirm all referenced V2 workbook/worksheet pairs exist in the 20 July config.
3. Confirm target identity uniqueness and distinct-NIPP counts.
4. Review every `Siap Dipetakan` row for namespace and runner-up conflicts.
5. Scan for blank required fields and formula errors.
6. Render the full worksheet and verify readable headers, wrapped evidence, visible confirmation fields, and no clipped critical values.

## Delivery and Phase Gate

Deliver only the verified mapping-review workbook and a concise status summary: counts of `Siap Dipetakan`, `Perlu Konfirmasi`, and `Tidak Ada Kandidat`.

Conversion may begin only after the user reviews and approves the mapping decisions. A later conversion phase must consume the approved reviewer fields, not silently reuse unapproved recommendations.
