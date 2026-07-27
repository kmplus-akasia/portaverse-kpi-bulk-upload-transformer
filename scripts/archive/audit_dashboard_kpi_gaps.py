from __future__ import annotations

import argparse
import csv
import json
import os
import re
import subprocess
import sys
import zipfile
from collections import Counter
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from dashboard.kpi_planning_dashboard import fetch_dashboard_data  # noqa: E402


YEAR = 2026
COMPANY_ID = 1
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
DEFAULT_UPLOAD_DIRS = [
    ROOT / "output/group1_ho_regenerated_20260615_final_v2/upload-ready",
    ROOT / "output/project_positions_upload_20260615/upload-ready/by-project",
]


def load_env_from_pid(pid: str) -> None:
    output = subprocess.check_output(["ps", "eww", "-p", pid], text=True)
    for key in [
        "DB_HOST",
        "DB_PORT",
        "DB_NAME",
        "DB_USER",
        "DB_PASSWORD",
        "DB_SSL",
        "KPI_DASHBOARD_DB_HOST",
        "KPI_DASHBOARD_DB_PORT",
        "KPI_DASHBOARD_DB_NAME",
        "KPI_DASHBOARD_DB_USER",
        "KPI_DASHBOARD_DB_PASSWORD",
        "KPI_DASHBOARD_DB_SSL",
    ]:
        match = re.search(rf"(?:^|\s){re.escape(key)}=([^\s]+)", output)
        if match:
            os.environ[key] = match.group(1)


def shared_strings(zf: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    return [
        "".join(text.text or "" for text in item.iter(NS + "t"))
        for item in root.findall(NS + "si")
    ]


def cell_text(cell: ET.Element, strings: list[str]) -> str:
    value = cell.find(NS + "v")
    if cell.get("t") == "inlineStr":
        inline = cell.find(NS + "is")
        return "" if inline is None else "".join(t.text or "" for t in inline.iter(NS + "t"))
    if value is None:
        return ""
    raw = value.text or ""
    if cell.get("t") == "s" and raw:
        return strings[int(raw)]
    return raw


def iter_xlsx_rows(path: Path) -> list[dict[str, str]]:
    with zipfile.ZipFile(path) as zf:
        strings = shared_strings(zf)
        root = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))
        rows = []
        header: list[str] | None = None
        for row in root.findall(".//" + NS + "row"):
            values = [cell_text(cell, strings) for cell in row.findall(NS + "c")]
            if header is None:
                header = values
                continue
            if not any(str(v).strip() for v in values):
                continue
            rows.append({header[i]: values[i] if i < len(values) else "" for i in range(len(header))})
        return rows


def scan_upload_workbooks(paths: list[Path]) -> dict[str, Any]:
    owners: dict[tuple[str, str], set[str]] = {}
    workbook_count = 0
    row_count = 0
    sample_headers: list[str] = []
    for folder in paths:
        if not folder.exists():
            continue
        for workbook in sorted(folder.glob("*.xlsx")):
            workbook_count += 1
            rows = iter_xlsx_rows(workbook)
            if rows and not sample_headers:
                sample_headers = list(rows[0])
            for row in rows:
                row_count += 1
                pmid = (
                    row.get("Position Master ID (Required)")
                    or row.get("position_master_id")
                    or ""
                ).strip()
                pnid = (
                    row.get("Position Nomenklatur ID")
                    or row.get("position_nomenklatur_id")
                    or ""
                ).strip()
                if pmid:
                    owners.setdefault(("PMID", pmid), set()).add(str(workbook.relative_to(ROOT)))
                if pnid:
                    owners.setdefault(("PNID", pnid), set()).add(str(workbook.relative_to(ROOT)))
    return {
        "owners": owners,
        "workbook_count": workbook_count,
        "row_count": row_count,
        "sample_headers": sample_headers,
    }


def is_ja(value: Any) -> bool:
    return str(value or "").strip().upper().startswith("JA_")


def classify(row: dict[str, Any], upload_owners: dict[tuple[str, str], set[str]]) -> tuple[str, str]:
    category = row["category"]
    if category == "Struktural":
        key = ("PMID", str(row.get("pmid") or ""))
        if key in upload_owners:
            return "already_in_generated_upload", "PMID sudah ada di workbook regenerate terakhir; kemungkinan workbook belum di-upload ulang atau import belum sukses setelah regenerate."
        name = str(row.get("position_name") or "")
        group = str(row.get("group_name") or "")
        if "Proyek" in name or "Proyek" in group or "Investasi" in name or "Terminal Kalibaru" in name:
            return "project_position_not_in_latest_upload", "Posisi struktural proyek/investasi; perlu upload formulir KPI posisi proyek atau mapping KPI proyek yang sesuai."
        if name.startswith(("Direktur", "Wakil Direktur")):
            return "director_level_no_dictionary_source", "Level Direksi aktif tetapi tidak muncul pada workbook upload terakhir; perlu kamus KPI khusus Direksi atau keputusan exclude."
        return "not_found_in_generated_upload", "PMID aktif belum ditemukan di workbook upload regenerate terakhir; perlu sumber kamus KPI atau mapping title ke kamus yang sudah ada."

    pnid = str(row.get("pnid") or "")
    key = ("PNID", pnid)
    if key in upload_owners:
        return "already_in_generated_upload", "PNID sudah ada di workbook regenerate terakhir; kemungkinan workbook belum di-upload ulang atau import belum sukses setelah regenerate."
    return "not_found_in_generated_upload", "PNID aktif belum ditemukan di workbook upload regenerate terakhir; perlu kamus KPI untuk nomenklatur ini atau mapping ke kamus setara."


def dataframe_records(df) -> list[dict[str, Any]]:
    return df.where(df.notna(), None).to_dict(orient="records")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--env-from-pid")
    parser.add_argument("--print-env-meta", action="store_true")
    parser.add_argument("--year", type=int, default=YEAR)
    parser.add_argument("--company-id", type=int, default=COMPANY_ID)
    parser.add_argument("--out", type=Path, default=ROOT / "output/production_kpi_gap_audit_20260615")
    args = parser.parse_args()

    if args.env_from_pid:
        load_env_from_pid(args.env_from_pid)
    if args.print_env_meta:
        for key in [
            "DB_HOST",
            "DB_PORT",
            "DB_NAME",
            "DB_USER",
            "DB_PASSWORD",
            "DB_SSL",
            "KPI_DASHBOARD_DB_HOST",
            "KPI_DASHBOARD_DB_PORT",
            "KPI_DASHBOARD_DB_NAME",
            "KPI_DASHBOARD_DB_USER",
            "KPI_DASHBOARD_DB_PASSWORD",
            "KPI_DASHBOARD_DB_SSL",
        ]:
            value = os.getenv(key)
            if value is not None:
                printable = "<redacted>" if "PASSWORD" in key else value
                print(f"{key}: present len={len(value)} value={printable}")
        return

    args.out.mkdir(parents=True, exist_ok=True)
    data = fetch_dashboard_data(args.year, args.company_id)
    upload = scan_upload_workbooks(DEFAULT_UPLOAD_DIRS)
    upload_owners = upload["owners"]

    structural = dataframe_records(data["structural_not_complete"])
    non_structural = dataframe_records(data["non_structural_not_complete"])
    gap_detail = dataframe_records(data["category_gap_detail"])

    rows: list[dict[str, Any]] = []
    for row in structural:
        if is_ja(row.get("position_name")):
            continue
        enriched = {
            "category": "Struktural",
            "unit": "PMID",
            "pmid": row.get("pmid"),
            "pnid": "",
            "label": row.get("position_name"),
            "group_name": row.get("group_names"),
            "availability_status": row.get("availability_status"),
            "active_variants": row.get("active_variants"),
            "with_kpi_variants": row.get("with_kpi_variants"),
            "without_kpi_variants": row.get("without_kpi_variants"),
            "impact_count": row.get("impact_count"),
            "output_count": row.get("output_count"),
            "kai_count": row.get("kai_count"),
            "active_pmid_list": row.get("pmid"),
            "position_names": row.get("position_name"),
        }
        reason_code, reason = classify(enriched, upload_owners)
        enriched["reason_code"] = reason_code
        enriched["reason"] = reason
        enriched["generated_workbooks"] = " | ".join(sorted(upload_owners.get(("PMID", str(row.get("pmid"))), [])))
        rows.append(enriched)

    for row in non_structural:
        if is_ja(row.get("pnid_label")):
            continue
        enriched = {
            "category": "Non-struktural",
            "unit": "PNID",
            "pmid": "",
            "pnid": row.get("pnid"),
            "label": row.get("pnid_label"),
            "group_name": row.get("group_names"),
            "availability_status": row.get("availability_status"),
            "active_variants": row.get("active_variants"),
            "with_kpi_variants": row.get("with_kpi_variants"),
            "without_kpi_variants": row.get("without_kpi_variants"),
            "impact_count": row.get("impact_count"),
            "output_count": row.get("output_count"),
            "kai_count": row.get("kai_count"),
            "active_pmid_list": row.get("active_pmid_list"),
            "position_names": row.get("position_names"),
        }
        reason_code, reason = classify(enriched, upload_owners)
        enriched["reason_code"] = reason_code
        enriched["reason"] = reason
        enriched["generated_workbooks"] = " | ".join(sorted(upload_owners.get(("PNID", str(row.get("pnid"))), [])))
        rows.append(enriched)

    detail_rows = []
    for row in gap_detail:
        if is_ja(row.get("pnid_label")) or is_ja(row.get("position_name")):
            continue
        detail_rows.append(row)

    fieldnames = [
        "category",
        "unit",
        "pnid",
        "pmid",
        "label",
        "group_name",
        "availability_status",
        "active_variants",
        "with_kpi_variants",
        "without_kpi_variants",
        "impact_count",
        "output_count",
        "kai_count",
        "active_pmid_list",
        "position_names",
        "reason_code",
        "reason",
        "generated_workbooks",
    ]
    with (args.out / "kpi_gap_units_filtered.csv").open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    detail_fields = sorted({key for row in detail_rows for key in row})
    with (args.out / "kpi_gap_position_variants_filtered.csv").open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=detail_fields)
        writer.writeheader()
        writer.writerows(detail_rows)

    reason_summary = Counter(row["reason_code"] for row in rows)
    category_summary = Counter((row["category"], row["availability_status"]) for row in rows)
    summary = {
        "year": args.year,
        "company_id": args.company_id,
        "ignored_rule": "labels starting with JA_",
        "unit_rows_after_filter": len(rows),
        "position_variant_rows_after_filter": len(detail_rows),
        "reason_summary": dict(reason_summary),
        "category_status_summary": {f"{k[0]}::{k[1]}": v for k, v in category_summary.items()},
        "upload_workbooks_scanned": upload["workbook_count"],
        "upload_rows_scanned": upload["row_count"],
        "upload_owner_count": len(upload_owners),
        "sample_upload_headers": upload["sample_headers"],
        "db_info": dataframe_records(data["db_info"])[0],
        "category_anomalies": dataframe_records(data["category_anomalies"]),
    }
    with (args.out / "summary.json").open("w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2, default=str)

    print(json.dumps(summary, ensure_ascii=False, indent=2, default=str))


if __name__ == "__main__":
    main()
