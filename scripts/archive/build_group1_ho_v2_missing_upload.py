"""Build one supplemental upload workbook for missing Group 1 HO V2 sheets."""

from __future__ import annotations

import argparse
import csv
import json
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from collections import Counter, defaultdict
from dataclasses import asdict
from pathlib import Path
from typing import Any

import kpi_bulk_transform as transform


MAIN_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
OFFICE_REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

SKIP_SHEET_TOKENS = (
    "panduan",
    "mapping",
    "sheet",
    "jadwal",
    "coverage",
    "existing",
    "master data",
)


def column_number(cell_ref: str) -> int:
    match = re.match(r"([A-Z]+)", cell_ref or "")
    if not match:
        return 0
    number = 0
    for char in match.group(1):
        number = number * 26 + ord(char) - 64
    return number


def row_number(cell_ref: str) -> int:
    match = re.search(r"(\d+)", cell_ref or "")
    return int(match.group(1)) if match else 0


def read_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    return [
        "".join(text.text or "" for text in item.iter(f"{MAIN_NS}t"))
        for item in root.findall(f"{MAIN_NS}si")
    ]


def cell_text(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return "".join(text.text or "" for text in cell.iter(f"{MAIN_NS}t")).strip()
    value = cell.find(f"{MAIN_NS}v")
    if value is None or value.text is None:
        return ""
    if cell_type == "s":
        try:
            return shared_strings[int(value.text)].strip()
        except (IndexError, ValueError):
            return value.text.strip()
    return value.text.strip()


def workbook_sheets(archive: zipfile.ZipFile) -> list[tuple[str, str, str]]:
    workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
    rel_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    rels = {
        rel.attrib.get("Id"): rel.attrib.get("Target", "").lstrip("/")
        for rel in rel_root.findall(f"{REL_NS}Relationship")
    }
    sheets: list[tuple[str, str, str]] = []
    for sheet in workbook_root.findall(f"{MAIN_NS}sheets/{MAIN_NS}sheet"):
        target = rels.get(sheet.attrib.get(f"{OFFICE_REL_NS}id"), "")
        if target and not target.startswith("xl/"):
            target = f"xl/{target}"
        sheets.append((sheet.attrib.get("name", ""), sheet.attrib.get("state", "visible"), target))
    return sheets


def inspect_sheet(
    archive: zipfile.ZipFile,
    target: str,
    shared_strings: list[str],
) -> dict[str, Any]:
    root = ET.fromstring(archive.read(target))
    tab_color = root.find(f"{MAIN_NS}sheetPr/{MAIN_NS}tabColor")
    color = ""
    if tab_color is not None:
        color = (
            tab_color.attrib.get("rgb")
            or tab_color.attrib.get("indexed")
            or tab_color.attrib.get("theme")
            or ""
        ).upper()

    by_row: dict[int, dict[int, str]] = {}
    max_row = 0
    max_col = 0
    for cell in root.findall(f".//{MAIN_NS}c"):
        ref = cell.attrib.get("r", "")
        row = row_number(ref)
        col = column_number(ref)
        max_row = max(max_row, row)
        max_col = max(max_col, col)
        if row <= 30 and col <= 45:
            text = cell_text(cell, shared_strings)
            if text:
                by_row.setdefault(row, {})[col] = text

    group_name = ""
    position_name = ""
    has_header = False
    output_titles = 0
    output_weights = 0
    for row_index, row in by_row.items():
        row_text = " | ".join(row.get(col, "") for col in range(1, 46)).lower()
        if "kpi impact" in row_text and "kpi output" in row_text:
            has_header = True
        for col, text in row.items():
            normalized = text.strip().lower()
            next_value = row.get(col + 1, "").strip()
            if normalized == "group name":
                group_name = next_value
            elif normalized == "posisi":
                position_name = next_value
        if 9 <= row_index <= 29:
            output_title = (row.get(8, "") or "").strip().lower()
            output_weight = (row.get(9, "") or "").strip().lower()
            if output_title and output_title != "(blank)":
                output_titles += 1
            if output_weight and output_weight not in {"(blank)", "0"}:
                output_weights += 1

    return {
        "tab_color": color,
        "group_name": group_name,
        "position_name": position_name,
        "has_header": has_header,
        "max_row": max_row,
        "max_col": max_col,
        "output_titles": output_titles,
        "output_weights": output_weights,
    }


def discover_missing_candidates(
    source_root: Path,
    configured_keys: set[tuple[str | None, str | None]],
) -> list[dict[str, Any]]:
    candidates: list[dict[str, Any]] = []
    for workbook_path in sorted(source_root.rglob("*.xlsx")):
        if workbook_path.name.startswith("~$"):
            continue
        source_workbook = str(workbook_path.relative_to(source_root))
        if transform.should_skip_source_workbook(source_workbook):
            continue
        with zipfile.ZipFile(workbook_path) as archive:
            shared_strings = read_shared_strings(archive)
            for sheet_name, state, target in workbook_sheets(archive):
                if state != "visible" or not target:
                    continue
                if any(token in sheet_name.lower() for token in SKIP_SHEET_TOKENS):
                    continue
                metadata = inspect_sheet(archive, target, shared_strings)
                if not metadata["has_header"]:
                    continue
                if not (metadata["position_name"] or metadata["group_name"] or metadata["output_titles"]):
                    continue
                if (source_workbook, sheet_name) in configured_keys:
                    continue
                candidates.append(
                    {
                        "source_workbook": source_workbook,
                        "workbook_path": str(workbook_path),
                        "sheet_name": sheet_name,
                        **metadata,
                    }
                )
    return candidates


def config_from_candidate(candidate: dict[str, Any]) -> transform.PositionConfig:
    position_name = candidate["position_name"] or candidate["sheet_name"]
    source_workbook = candidate["source_workbook"]
    return transform.PositionConfig(
        source_workbook=source_workbook,
        sheet_name=candidate["sheet_name"],
        position_name=position_name,
        group_name=candidate["group_name"] or "",
        directorate_name=Path(source_workbook).stem.split(" - ")[0],
        position_lookup_names=[position_name, candidate["sheet_name"]],
    )


def csv_row(config: transform.PositionConfig, candidate: dict[str, Any]) -> dict[str, Any]:
    return {
        "source_workbook": config.source_workbook,
        "sheet_name": config.sheet_name,
        "position_name": config.position_name,
        "tab_color": candidate["tab_color"],
        "output_titles_sample": candidate["output_titles"],
        "output_weights_sample": candidate["output_weights"],
        "mapping_confidence_label": config.mapping_confidence_label,
        "position_scope": config.position_scope,
        "position_master_id": config.position_master_id,
        "position_nomenclature_id": config.position_nomenclature_id,
        "candidate_position_master_id": config.candidate_position_master_id,
        "candidate_position_nomenclature_id": config.candidate_position_nomenclature_id,
        "candidate_score": config.candidate_score,
        "portaverse_position_title": config.portaverse_position_title,
        "portaverse_group_name": config.portaverse_group_name,
        "mapping_confidence_reason": config.mapping_confidence_reason,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source-root", required=True, type=Path)
    parser.add_argument("--reviewed-config", required=True, type=Path)
    parser.add_argument("--template", required=True, type=Path)
    parser.add_argument("--mapping", required=True, type=Path)
    parser.add_argument("--output-dir", required=True, type=Path)
    args = parser.parse_args()

    reviewed = json.loads(args.reviewed_config.read_text(encoding="utf-8"))
    configured_keys = {
        (item.get("source_workbook"), item.get("sheet_name"))
        for item in reviewed.get("positions", [])
    }
    candidates = discover_missing_candidates(args.source_root, configured_keys)

    nomenclature_mapping, strict_indexes = transform.load_position_reference(
        args.mapping,
        transform.TARGET_COMPANY_ID_DEFAULT,
        None,
    )
    configs: list[tuple[transform.PositionConfig, dict[str, Any]]] = []
    for candidate in candidates:
        config = config_from_candidate(candidate)
        transform.refresh_config_from_mapping(config, nomenclature_mapping, None, strict_indexes)
        configs.append((config, candidate))

    safe_configs = [
        config
        for config, _ in configs
        if config.mapping_confidence_label == "high_confidence"
        and bool(config.position_master_id or config.position_nomenclature_id)
    ]

    args.output_dir.mkdir(parents=True, exist_ok=True)
    all_config_path = args.output_dir / "group1_ho_v2_missing_all_candidates.config.json"
    safe_config_path = args.output_dir / "group1_ho_v2_missing_high_confidence.config.json"
    audit_csv_path = args.output_dir / "group1_ho_v2_missing_mapping_audit.csv"
    output_xlsx_path = args.output_dir / "group1_ho_v2_missing_high_confidence_upload.xlsx"
    report_csv_path = args.output_dir / "group1_ho_v2_missing_high_confidence_upload.report.csv"

    all_config_path.write_text(
        json.dumps({"positions": [asdict(config) for config, _ in configs]}, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    safe_config_path.write_text(
        json.dumps({"positions": [asdict(config) for config in safe_configs]}, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )

    fieldnames = list(csv_row(configs[0][0], configs[0][1]).keys()) if configs else []
    with audit_csv_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for config, candidate in configs:
            writer.writerow(csv_row(config, candidate))

    configs_by_workbook: dict[str, list[transform.PositionConfig]] = defaultdict(list)
    workbook_paths: dict[str, Path] = {}
    for config, candidate in configs:
        if config in safe_configs and config.source_workbook:
            configs_by_workbook[config.source_workbook].append(config)
            workbook_paths[config.source_workbook] = Path(candidate["workbook_path"])

    issues: list[transform.ValidationIssue] = []
    parsed_sheets: list[transform.ParsedSheet] = []
    for source_workbook in sorted(configs_by_workbook):
        parsed_sheets.extend(
            transform.collect_parsed_sheets(
                workbook_paths[source_workbook],
                None,
                configs_by_workbook[source_workbook],
                issues,
            )
        )

    transform.backfill_shared_impact_fields(parsed_sheets)
    rows, errors, warnings, infos = transform.write_transformed_workbook(
        args.template,
        output_xlsx_path,
        report_csv_path,
        parsed_sheets,
        issues,
    )

    status_counts = Counter(config.mapping_confidence_label for config, _ in configs)
    print(f"missing_candidates={len(configs)}")
    print(f"status_counts={dict(status_counts)}")
    print(f"safe_high_confidence={len(safe_configs)}")
    print(f"generated_rows={rows}")
    print(f"errors={errors}")
    print(f"warnings={warnings}")
    print(f"infos={infos}")
    print(f"upload={output_xlsx_path}")
    print(f"report={report_csv_path}")
    print(f"audit={audit_csv_path}")
    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
