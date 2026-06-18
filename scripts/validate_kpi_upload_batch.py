#!/usr/bin/env python3
"""Validate generated KPI upload workbooks and create an upload manifest."""

from __future__ import annotations

import argparse
import csv
import json
import re
import shutil
import sys
import zipfile
from pathlib import Path

import openpyxl


EXPECTED_HEADERS = [
    "IDKPI",
    "Group",
    "Direktorat",
    "Posisi",
    "Position Master ID (Required)",
    "Position Master Variant ID (Optional)",
    "BSC Perspective",
    "KPI Type",
    "Parent KPI ID",
    "Parent KPI Title",
    "Title",
    "Description",
    "Unit",
    "Polarity",
    "Period",
    "Formula",
    "Weight (%)",
    "Cascading",
    "Nature Of Work (KAI Only)",
    "External ID (PKPI)",
    "System KPI ID",
    "Ownership Type",
    "Position Nomenklatur ID",
    "RKM Code ID",
]


def norm(value: object) -> str:
    text = str(value or "").lower()
    text = re.sub(r"\bdh\b", "department head", text)
    text = text.replace("&", " dan ")
    text = re.sub(r"[-_/(),.]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-dir", required=True, type=Path)
    parser.add_argument("--config", required=True, type=Path)
    parser.add_argument("--reference", required=True, type=Path)
    parser.add_argument("--fixed-audit", type=Path)
    parser.add_argument("--expected-workbooks", type=int, default=26)
    parser.add_argument("--upload-ready-dir", type=Path)
    parser.add_argument("--zip-output", type=Path)
    return parser.parse_args()


def load_reference_ids(
    reference_path: Path,
) -> tuple[
    set[str],
    set[str],
    dict[str, set[str]],
    dict[str, set[str]],
    dict[str, set[str]],
]:
    with reference_path.open() as handle:
        reference = json.load(handle)
    master_ids = {str(row["position_master_id"]) for row in reference["position_master_rows"]}
    nomenclature_ids = {
        str(row["cluster_id"])
        for row in reference["rows"]
        if row.get("cluster_id") not in (None, "", 0, "0")
    }
    cluster_labels_by_id: dict[str, set[str]] = {}
    master_types_by_id: dict[str, set[str]] = {}
    position_types_by_pnid: dict[str, set[str]] = {}
    for row in reference["position_master_rows"]:
        master_types_by_id.setdefault(str(row["position_master_id"]), set()).add(
            str(row.get("position_master_type_id") or "")
        )
    for row in reference["rows"]:
        cluster_id = row.get("cluster_id")
        if cluster_id in (None, "", 0, "0"):
            continue
        cluster_labels_by_id.setdefault(str(cluster_id), set()).add(norm(row.get("cluster_label")))
        position_types_by_pnid.setdefault(str(cluster_id), set()).add(
            str(row.get("position_master_type_id") or "")
        )
    return (
        master_ids,
        nomenclature_ids,
        cluster_labels_by_id,
        master_types_by_id,
        position_types_by_pnid,
    )


def check_config_scope(
    config_path: Path,
    master_ids: set[str],
    nomenclature_ids: set[str],
    cluster_labels_by_id: dict[str, set[str]],
    master_types_by_id: dict[str, set[str]],
    position_types_by_pnid: dict[str, set[str]],
) -> list[str]:
    errors: list[str] = []
    with config_path.open() as handle:
        config = json.load(handle)
    for pos in config["positions"]:
        scope = (pos.get("position_scope") or "").strip()
        pmid = str(pos.get("position_master_id") or "").strip()
        pnid = str(pos.get("position_nomenclature_id") or "").strip()
        label = f"{pos.get('source_workbook')} :: {pos.get('sheet_name')}"
        if scope not in {"structural", "non_structural", "neglect"}:
            errors.append(f"unsupported position scope {scope or '<blank>'}: {label}")
        if scope == "structural":
            if not pmid:
                errors.append(f"structural config missing PMID: {label}")
            if pnid:
                errors.append(f"structural config still has PNID: {label} -> {pnid}")
            production_types = master_types_by_id.get(pmid, set())
            if pmid and production_types != {"5"}:
                errors.append(
                    f"structural config PMID {pmid} production type "
                    f"{','.join(sorted(production_types)) or 'unknown'} is non-structural: {label}"
                )
        if scope == "non_structural":
            if pmid:
                errors.append(f"non_structural config has PMID populated: {label} -> {pmid}")
            if pnid and pnid not in nomenclature_ids and pnid in master_ids:
                errors.append(
                    f"non_structural config points to structural PMID instead of valid PNID: "
                    f"{label} -> pnid={pnid}"
                )
            production_types = position_types_by_pnid.get(pnid, set())
            if pnid and (not production_types or "5" in production_types):
                errors.append(
                    f"non_structural config PNID {pnid} has invalid production types "
                    f"{sorted(production_types)}: {label}"
                )
    return errors


def load_fixed_pmids(fixed_audit_path: Path | None) -> set[str]:
    if not fixed_audit_path or not fixed_audit_path.exists():
        return set()
    with fixed_audit_path.open(newline="") as handle:
        rows = list(csv.DictReader(handle))
    return {
        str(row["pmid"]).strip()
        for row in rows
        if str(row.get("pmid") or "").strip()
        and (not row.get("resolved_scope") or row.get("resolved_scope") == "structural")
    }


def iter_upload_workbooks(output_dir: Path) -> list[Path]:
    paths = []
    for path in output_dir.rglob("*.xlsx"):
        if path.name.startswith("~$"):
            continue
        if path.parent.name == "upload-ready":
            continue
        if "Conversion Report" in path.name:
            continue
        paths.append(path)
    return sorted(paths)


def check_report_csvs(output_dir: Path) -> list[str]:
    errors: list[str] = []
    for report_path in sorted(output_dir.rglob("*.report.csv")):
        with report_path.open(newline="") as handle:
            for row in csv.DictReader(handle):
                if (row.get("severity") or "").lower() == "error":
                    errors.append(
                        f"{report_path}: row {row.get('source_row')} "
                        f"{row.get('sheet_name')}: {row.get('message')}"
                    )
    return errors


def validate_workbook(
    path: Path,
    fixed_pmids: set[str],
    master_ids: set[str],
    nomenclature_ids: set[str],
    master_types_by_id: dict[str, set[str]],
    position_types_by_pnid: dict[str, set[str]],
) -> tuple[dict[str, object], list[str], set[str]]:
    errors: list[str] = []
    found_fixed: set[str] = set()

    with zipfile.ZipFile(path) as archive:
        bad_file = archive.testzip()
        if bad_file:
            errors.append(f"{path}: bad xlsx zip member {bad_file}")

    workbook = openpyxl.load_workbook(path, read_only=True, data_only=False)
    if "KPI Template" not in workbook.sheetnames:
        errors.append(f"{path}: missing KPI Template sheet")
        return {"path": str(path), "rows": 0, "status": "FAIL"}, errors, found_fixed

    sheet = workbook["KPI Template"]
    row_iter = sheet.iter_rows(
        min_row=1,
        max_col=len(EXPECTED_HEADERS),
        values_only=True,
    )
    headers = list(next(row_iter, ()))
    if headers != EXPECTED_HEADERS:
        errors.append(f"{path}: upload header schema changed")

    rows = 0
    structural_rows = 0
    non_structural_rows = 0
    blank_identity_rows = 0
    double_identity_rows = 0

    for row_idx, row in enumerate(row_iter, 2):
        title = row[10]
        if title in (None, ""):
            continue
        rows += 1
        pmid = str(row[4] or "").strip()
        pnid = str(row[22] or "").strip()
        if pmid and pnid:
            double_identity_rows += 1
            errors.append(f"{path}: row {row_idx} has both PMID={pmid} and PNID={pnid}")
        elif not pmid and not pnid:
            blank_identity_rows += 1
            errors.append(f"{path}: row {row_idx} has neither PMID nor PNID")
        elif pmid:
            structural_rows += 1
            if pmid not in master_ids:
                errors.append(f"{path}: row {row_idx} has invalid PMID={pmid}")
            production_types = master_types_by_id.get(pmid, set())
            if production_types != {"5"}:
                errors.append(
                    f"{path}: row {row_idx} PMID={pmid} has non-structural "
                    f"production types {sorted(production_types)}"
                )
            if pmid in fixed_pmids:
                found_fixed.add(pmid)
        else:
            non_structural_rows += 1
            if pnid not in nomenclature_ids:
                errors.append(f"{path}: row {row_idx} has invalid PNID={pnid}")
            production_types = position_types_by_pnid.get(pnid, set())
            if not production_types or "5" in production_types:
                errors.append(
                    f"{path}: row {row_idx} PNID={pnid} has invalid production "
                    f"types {sorted(production_types)}"
                )
            if pnid in fixed_pmids:
                errors.append(f"{path}: row {row_idx} fixed structural PMID appears in PNID={pnid}")

    return (
        {
            "path": str(path),
            "rows": rows,
            "structural_rows": structural_rows,
            "non_structural_rows": non_structural_rows,
            "blank_identity_rows": blank_identity_rows,
            "double_identity_rows": double_identity_rows,
            "status": "FAIL" if errors else "READY",
        },
        errors,
        found_fixed,
    )


def write_manifest(output_dir: Path, records: list[dict[str, object]], upload_ready_dir: Path | None) -> None:
    csv_path = output_dir / "upload_manifest.csv"
    fields = [
        "status",
        "rows",
        "structural_rows",
        "non_structural_rows",
        "blank_identity_rows",
        "double_identity_rows",
        "path",
        "upload_ready_path",
    ]
    with csv_path.open("w", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        for record in records:
            writer.writerow({field: record.get(field, "") for field in fields})

    md_path = output_dir / "UPLOAD_THESE_FILES.md"
    lines = [
        "# Upload These Files",
        "",
        "Upload the workbooks listed below. Do not upload the conversion report, config, audit, or CSV files.",
        "",
    ]
    for index, record in enumerate(records, 1):
        upload_path = record.get("upload_ready_path") or record["path"]
        lines.append(f"{index}. `{upload_path}`")
    lines.append("")
    md_path.write_text("\n".join(lines))


def prepare_upload_ready(records: list[dict[str, object]], upload_ready_dir: Path) -> None:
    if upload_ready_dir.exists():
        shutil.rmtree(upload_ready_dir)
    upload_ready_dir.mkdir(parents=True, exist_ok=True)
    for record in records:
        source = Path(str(record["path"]))
        target = upload_ready_dir / source.name
        shutil.copy2(source, target)
        record["upload_ready_path"] = str(target)


def write_zip(records: list[dict[str, object]], zip_output: Path) -> None:
    if zip_output.exists():
        zip_output.unlink()
    with zipfile.ZipFile(zip_output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for record in records:
            source = Path(str(record.get("upload_ready_path") or record["path"]))
            archive.write(source, arcname=source.name)
    with zipfile.ZipFile(zip_output) as archive:
        bad_file = archive.testzip()
    if bad_file:
        raise RuntimeError(f"bad upload zip member: {bad_file}")


def main() -> int:
    args = parse_args()
    (
        master_ids,
        nomenclature_ids,
        cluster_labels_by_id,
        master_types_by_id,
        position_types_by_pnid,
    ) = load_reference_ids(args.reference)
    fixed_pmids = load_fixed_pmids(args.fixed_audit)

    errors = check_config_scope(
        args.config,
        master_ids,
        nomenclature_ids,
        cluster_labels_by_id,
        master_types_by_id,
        position_types_by_pnid,
    )
    errors.extend(check_report_csvs(args.output_dir))

    workbook_paths = iter_upload_workbooks(args.output_dir)
    if len(workbook_paths) != args.expected_workbooks:
        errors.append(f"expected {args.expected_workbooks} upload workbooks, found {len(workbook_paths)}")

    records: list[dict[str, object]] = []
    found_fixed_all: set[str] = set()
    for path in workbook_paths:
        record, workbook_errors, found_fixed = validate_workbook(
            path,
            fixed_pmids,
            master_ids,
            nomenclature_ids,
            master_types_by_id,
            position_types_by_pnid,
        )
        records.append(record)
        errors.extend(workbook_errors)
        found_fixed_all.update(found_fixed)

    missing_fixed = fixed_pmids - found_fixed_all
    for pmid in sorted(missing_fixed, key=int):
        errors.append(f"fixed structural PMID not found in generated upload rows: {pmid}")

    ready_records = [record for record in records if record["status"] == "READY"]
    if args.upload_ready_dir:
        prepare_upload_ready(ready_records, args.upload_ready_dir)
    if args.zip_output:
        write_zip(ready_records, args.zip_output)

    write_manifest(args.output_dir, records, args.upload_ready_dir)

    total_rows = sum(int(record["rows"]) for record in records)
    structural_rows = sum(int(record["structural_rows"]) for record in records)
    non_structural_rows = sum(int(record["non_structural_rows"]) for record in records)

    print(f"upload_workbooks={len(workbook_paths)}")
    print(f"ready_workbooks={len(ready_records)}")
    print(f"total_rows={total_rows}")
    print(f"structural_rows={structural_rows}")
    print(f"non_structural_rows={non_structural_rows}")
    print(f"fixed_structural_pmids_checked={len(fixed_pmids)}")
    print(f"errors={len(errors)}")

    if errors:
        for error in errors[:100]:
            print(f"ERROR: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
