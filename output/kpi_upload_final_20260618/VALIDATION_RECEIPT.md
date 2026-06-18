# KPI Upload Final Validation Receipt

Generated: 2026-06-18

## Result

- Upload workbooks: 26
- Ready workbooks: 26
- Total KPI rows: 6,118
- Structural rows using PMID: 3,838
- Non-structural rows using PNID: 2,280
- Batch validation errors: 0
- Corrected invalid PNID to structural PMID: 9 positions
- Explicitly neglected unresolved positions: 13 positions

## Reported Examples

- `Officer Transaksi dan Proses`: PMID blank, PNID `44`
- `Officer QA`: PMID blank, PNID `11517`

## Checks

- Transformer and scope regression suite: 42 tests passed
- Python compile checks: passed
- Workbook ZIP CRC and XML parsing: 26/26 passed, 260 XML members parsed
- Worksheet pane/selection consistency: 26/26 passed
- Representative LibreOffice open/PDF conversion: 2/2 passed
- Upload ZIP test: passed

Upload only files listed in `UPLOAD_THESE_FILES.md` or use
`KPI_Upload_Final_20260618.zip`.
