# Identity shapes

Three shapes are valid in the upload template. Every row carries exactly one of them.

| Shape | `Position Master ID` | `Position Master Variant ID` | `Position Nomenklatur ID` |
| --- | --- | --- | --- |
| Structural | the PMID | optional | blank |
| Non-structural | blank | blank | the PNID (`cluster_id`) |
| Assistant (PA) | `77` | the PMVID for that direksi location | blank |

## Structural

The worksheet maps to an active position master. The importer takes the PMID directly.

## Non-structural

The worksheet maps to a nomenclature cluster. Leaving `Position Master ID` blank is deliberate: the importer expands the PNID into its position masters. Filling both fields makes the identity ambiguous and fails validation.

## Assistant

Head Office Personal Assistant assignments all share PMID `77`, so the PMID alone cannot tell one PA apart from another. The PMVID is what distinguishes the direksi location, which makes this the one shape where PMID and PMVID appear together and both are required.

This shape came out of the July 2026 PA and Staff work: the KPI material sat on hidden worksheets, and eleven active PA PMVIDs were covered only after the location-specific workbooks were combined with the historical Direktur Utama and Wakil Direktur Utama packages. Two PA identities under Direktur Utama share one KPI branch across two PMVIDs, which is expected rather than a duplicate.
