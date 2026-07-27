#!/usr/bin/env python3
"""Split a consolidated upload-rows JSON into N batch workbooks with renumbered IDKPI."""

from __future__ import annotations

import argparse
import json
import math
import sys
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from kpi_bulk_transform import (  # noqa: E402
    UPLOAD_HEADERS,
    normalize_bsc_perspective,
    write_output_workbook,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--rows-json", required=True, type=Path, help="upload_rows.json payload")
    parser.add_argument("--template", required=True, type=Path, help="Official upload template xlsx")
    parser.add_argument("--output-dir", required=True, type=Path, help="Run-scoped output directory")
    parser.add_argument("--batches", type=int, default=3, help="Number of batch workbooks")
    parser.add_argument(
        "--name-prefix",
        default="Formulir_Upload_KPI_Historikal_Q1",
        help="Filename prefix for each batch workbook",
    )
    parser.add_argument("--date-stamp", default=datetime.now(timezone.utc).strftime("%Y%m%d"))
    return parser.parse_args()


def identity_key(row: dict[str, Any]) -> tuple[str, str]:
    pmid = str(row.get("Position Master ID (Required)") or "").strip()
    pmvid = str(row.get("Position Master Variant ID (Optional)") or "").strip()
    pnid = str(row.get("Position Nomenklatur ID") or "").strip()
    if pmvid and pmid == "77":
        return ("PMVID", pmvid)
    if pmid and not pnid:
        return ("PMID", pmid)
    if pnid and not pmid:
        return ("PNID", pnid)
    raise ValueError(f"Invalid upload identity scope: pmid={pmid!r} pnid={pnid!r} pmvid={pmvid!r}")


def group_rows_by_identity(rows: list[dict[str, Any]]) -> list[tuple[tuple[str, str], list[dict[str, Any]]]]:
    groups: list[tuple[tuple[str, str], list[dict[str, Any]]]] = []
    current_key: tuple[str, str] | None = None
    for row in rows:
        key = identity_key(row)
        if key != current_key:
            groups.append((key, []))
            current_key = key
        groups[-1][1].append(row)
    return groups


def apply_bsc_normalization(rows: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    normalized_rows: list[dict[str, Any]] = []
    audit: list[dict[str, Any]] = []
    for index, row in enumerate(rows, start=1):
        updated = dict(row)
        raw_bsc = row.get("BSC Perspective")
        result = normalize_bsc_perspective(raw_bsc)
        if result.value and result.value != raw_bsc:
            updated["BSC Perspective"] = result.value
            audit.append(
                {
                    "row_index": index,
                    "kpi_type": row.get("KPI Type"),
                    "title": row.get("Title"),
                    "from": raw_bsc,
                    "to": result.value,
                    "status": result.status.value,
                }
            )
        normalized_rows.append(updated)
    return normalized_rows, audit


def split_identity_groups(
    groups: list[tuple[tuple[str, str], list[dict[str, Any]]]],
    batch_count: int,
) -> list[list[tuple[tuple[str, str], list[dict[str, Any]]]]]:
    if batch_count < 1:
        raise ValueError("batch_count must be at least 1")
    if not groups:
        return [[] for _ in range(batch_count)]

    per_batch = math.ceil(len(groups) / batch_count)
    batches: list[list[tuple[tuple[str, str], list[dict[str, Any]]]]] = []
    for start in range(0, len(groups), per_batch):
        batches.append(groups[start : start + per_batch])
    while len(batches) < batch_count:
        batches.append([])
    return batches[:batch_count]


def renumber_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    old_to_new: dict[int, int] = {}
    renumbered: list[dict[str, Any]] = []
    next_id = 1
    for row in rows:
        old_id = int(float(row["IDKPI"]))
        old_to_new[old_id] = next_id
        updated = dict(row)
        updated["IDKPI"] = next_id
        renumbered.append(updated)
        next_id += 1

    for row in renumbered:
        parent_raw = row.get("Parent KPI ID")
        if parent_raw in (None, ""):
            continue
        parent_id = int(float(parent_raw))
        row["Parent KPI ID"] = old_to_new[parent_id]

    return renumbered


def rows_to_matrix(rows: list[dict[str, Any]]) -> list[list[Any]]:
    return [[row.get(header) for header in UPLOAD_HEADERS] for row in rows]


def validate_batch_rows(rows: list[dict[str, Any]]) -> list[str]:
    errors: list[str] = []
    idkpis = [int(row["IDKPI"]) for row in rows]
    expected = list(range(1, len(idkpis) + 1))
    if idkpis != expected:
        errors.append("IDKPI is not sequential 1..N")
    id_set = set(idkpis)
    for row in rows:
        ktype = str(row.get("KPI Type") or "").strip().upper()
        if ktype == "OUTPUT":
            bsc = normalize_bsc_perspective(row.get("BSC Perspective"))
            if not bsc.value:
                errors.append(f"OUTPUT row {row['IDKPI']} has invalid BSC Perspective: {row.get('BSC Perspective')!r}")
        parent = row.get("Parent KPI ID")
        if ktype in {"OUTPUT", "KAI"} and parent not in (None, ""):
            if int(parent) not in id_set:
                errors.append(f"Row {row['IDKPI']} parent {parent} missing from batch")
    return errors


def write_manifest(
    output_dir: Path,
    batches: list[dict[str, Any]],
    bsc_audit: list[dict[str, Any]],
    full_rows: list[dict[str, Any]],
) -> None:
    support = output_dir / "support"
    support.mkdir(parents=True, exist_ok=True)
    (support / "split_manifest.json").write_text(json.dumps({"batches": batches}, indent=2) + "\n", encoding="utf-8")
    (support / "bsc_normalization_audit.json").write_text(
        json.dumps({"changes": bsc_audit, "total_rows": len(full_rows)}, indent=2) + "\n",
        encoding="utf-8",
    )

    manifest_lines = ["# Upload these files", ""]
    for batch in batches:
        manifest_lines.append(f"- `{batch['filename']}` — {batch['identity_count']} identities, {batch['row_count']} rows")
    manifest_lines.append("")
    (output_dir / "UPLOAD_THESE_FILES.md").write_text("\n".join(manifest_lines) + "\n", encoding="utf-8")


def main() -> int:
    args = parse_args()
    payload = json.loads(args.rows_json.read_text(encoding="utf-8"))
    headers = payload["headers"]
    if headers != UPLOAD_HEADERS:
        raise SystemExit("upload_rows.json headers do not match converter contract")

    rows, bsc_audit = apply_bsc_normalization(payload["rows"])
    groups = group_rows_by_identity(rows)
    batch_groups = split_identity_groups(groups, args.batches)

    args.output_dir.mkdir(parents=True, exist_ok=True)
    batch_summaries: list[dict[str, Any]] = []
    type_totals: Counter[str] = Counter()

    for batch_index, identity_batch in enumerate(batch_groups, start=1):
        batch_rows = [row for _, group_rows in identity_batch for row in group_rows]
        if not batch_rows:
            continue
        batch_rows = renumber_rows(batch_rows)
        errors = validate_batch_rows(batch_rows)
        if errors:
            raise SystemExit(f"Batch {batch_index} validation failed: {errors[:5]}")

        filename = f"{args.name_prefix}_Batch_{batch_index}_of_{args.batches}_{args.date_stamp}.xlsx"
        output_path = args.output_dir / filename
        write_output_workbook(args.template, output_path, rows_to_matrix(batch_rows))

        identities = [f"{namespace}:{identity}" for (namespace, identity), _ in identity_batch]
        batch_type_counts = Counter(str(row.get("KPI Type") or "") for row in batch_rows)
        type_totals.update(batch_type_counts)
        batch_summaries.append(
            {
                "batch": batch_index,
                "filename": filename,
                "path": str(output_path),
                "identity_count": len(identity_batch),
                "row_count": len(batch_rows),
                "type_counts": dict(batch_type_counts),
                "identities": identities,
            }
        )

    write_manifest(args.output_dir, batch_summaries, bsc_audit, rows)
    print(json.dumps({"output_dir": str(args.output_dir), "batches": batch_summaries, "bsc_changes": len(bsc_audit)}, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
