#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT"

export DB_READ_WRITE=0

echo "Exporting canonical production reference..."
node scripts/export_position_reference.mjs \
  --profile production \
  --output configs/production_position_reference.json

python3 << 'PY'
import json
from pathlib import Path

canonical = Path("configs/production_position_reference.json")
ref = json.loads(canonical.read_text())
src = ref.get("source", {})
meta = {
    "canonical_path": str(canonical),
    "exported_at": src.get("exported_at"),
    "profile": src.get("profile"),
    "database": src.get("database"),
    "review_status": src.get("review_status", "current_snapshot_unreviewed"),
    "read_only": src.get("read_only"),
    "row_counts": {
        "rows": len(ref.get("rows", [])),
        "position_master_rows": len(ref.get("position_master_rows", [])),
        "structural_lookup_rows": len(ref.get("structural_lookup_rows", [])),
        "non_structural_lookup_rows": len(ref.get("non_structural_lookup_rows", [])),
        "active_assignment_rows": len(ref.get("active_assignment_rows", [])),
        "organization_rows": len(ref.get("organization_rows", [])),
        "company_rows": len(ref.get("company_rows", [])),
    },
    "refreshed_by": "scripts/refresh_canonical_production_reference.sh",
    "notes": "Single canonical production snapshot for mapping and conversion. Refresh before identity decisions when production org data may have drifted.",
}
Path("configs/production_position_reference.meta.json").write_text(
    json.dumps(meta, ensure_ascii=False, indent=2) + "\n",
    encoding="utf-8",
)

receipt_dir = Path("outputs/production-reference")
receipt_dir.mkdir(parents=True, exist_ok=True)
counts = meta["row_counts"]
receipt = f"""# Production Reference Receipt

## Canonical snapshot
- **Path:** `configs/production_position_reference.json`
- **Metadata:** `configs/production_position_reference.meta.json`
- **Exported at:** {src.get('exported_at')}
- **Profile:** {src.get('profile')}
- **Database:** {src.get('database')}
- **Review status:** {src.get('review_status', 'current_snapshot_unreviewed')}

## Row counts
| Section | Count |
| --- | ---: |
| nomenclature rows | {counts['rows']} |
| position_master_rows | {counts['position_master_rows']} |
| structural_lookup_rows | {counts['structural_lookup_rows']} |
| non_structural_lookup_rows | {counts['non_structural_lookup_rows']} |
| active_assignment_rows | {counts['active_assignment_rows']} |
| organization_rows | {counts['organization_rows']} |
| company_rows | {counts['company_rows']} |

## Refresh
```sh
DB_READ_WRITE=0 ./scripts/refresh_canonical_production_reference.sh
```
"""
(receipt_dir / "REFERENCE_RECEIPT.md").write_text(receipt, encoding="utf-8")
print(f"exported_at={meta['exported_at']}")
print(f"structural_lookup_rows={counts['structural_lookup_rows']}")
PY

echo "Updated configs/production_position_reference.meta.json"
echo "See outputs/production-reference/REFERENCE_RECEIPT.md"
