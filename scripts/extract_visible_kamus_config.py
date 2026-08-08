#!/usr/bin/env python3
"""Extract visible KPI worksheets and in-sheet position titles from OOXML workbooks."""

from __future__ import annotations

import argparse
import json
import re
import zipfile
from datetime import datetime
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree as ET

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CELL_REF_RE = re.compile(r"([A-Z]+)(\d+)")

POSITION_LABELS = {
    "nama posisi",
    "posisi",
    "nama jabatan",
    "jabatan",
    "position name",
    "position",
}
GROUP_LABELS = {"group name", "nama group", "nama grup", "group", "grup", "unit kerja"}
SUPPORT_SHEET_RE = re.compile(
    r"^(?:\d+[a-z]?\.\s*)?(?:panduan|jadwal validator|mapping organisasi|kpi coverage|kpi existing|new kpi|master data|sheet\d*)$",
    flags=re.I,
)
GENERIC_TOP_LEFT = {
    "group name",
    "nama group",
    "nama grup",
    "group",
    "grup",
    "unit kerja",
    "nama posisi",
    "posisi",
    "nama jabatan",
    "jabatan",
    "bsc perspective",
    "jenis posisi",
    "kpi impact",
    "kpi output",
    "financial",
    "customer",
    "internal process",
    "learning & growth",
    "learning and growth",
}
GENERIC_KAMUS_TAB_TITLES = {
    "kamus kpi bagian",
    "kamus kpi officer",
    "kamus kpi staff",
    "kamus kpi sub bagian",
    "kamus kpi kawasan",
    "kamus kpi kelompok kerja",
    "kamus kpi unit",
    "kamus kpi pelabuhan kawasan",
    "kamus kpi sub regional head",
}


def normalize_label(value: str) -> str:
    return re.sub(r"\s+", " ", value.strip().rstrip(":").strip()).casefold()


def shared_strings(archive: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    return ["".join(node.text or "" for node in si.iter(f"{{{MAIN_NS}}}t")) for si in root]


def workbook_sheets(archive: zipfile.ZipFile) -> list[dict[str, str | int]]:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rels.findall(f"{{{PKG_REL_NS}}}Relationship")
    }
    result = []
    for order, sheet in enumerate(workbook.findall(f".//{{{MAIN_NS}}}sheet"), start=1):
        rel_id = sheet.attrib[f"{{{REL_NS}}}id"]
        target = targets[rel_id]
        xml_path = target.lstrip("/") if target.startswith("/") else str(PurePosixPath("xl") / target)
        result.append(
            {
                "sheet_order": order,
                "sheet_name": sheet.attrib["name"],
                "visibility": sheet.attrib.get("state", "visible"),
                "xml_path": xml_path,
            }
        )
    return result


def cell_value(cell: ET.Element, strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.iter(f"{{{MAIN_NS}}}t"))
    value = cell.find(f"{{{MAIN_NS}}}v")
    raw = "" if value is None else (value.text or "")
    if cell_type == "s" and raw.isdigit():
        index = int(raw)
        return strings[index] if index < len(strings) else raw
    return raw


def column_number(ref: str) -> int:
    match = CELL_REF_RE.fullmatch(ref)
    letters = match.group(1) if match else "A"
    number = 0
    for char in letters:
        number = number * 26 + ord(char) - 64
    return number


def sheet_cells(
    archive: zipfile.ZipFile,
    xml_path: str,
    strings: list[str],
    max_row: int = 40,
    max_col: int = 24,
) -> dict[tuple[int, int], tuple[str, str]]:
    root = ET.fromstring(archive.read(xml_path))
    cells: dict[tuple[int, int], tuple[str, str]] = {}
    for cell in root.findall(f".//{{{MAIN_NS}}}c"):
        ref = cell.attrib.get("r", "")
        match = CELL_REF_RE.fullmatch(ref)
        if not match:
            continue
        row = int(match.group(2))
        col = column_number(ref)
        if row > max_row or col > max_col:
            continue
        value = cell_value(cell, strings).strip()
        if value:
            cells[(row, col)] = (ref, value)
    return cells


def extract_labeled_value(
    cells: dict[tuple[int, int], tuple[str, str]], labels: set[str]
) -> tuple[str, str, str]:
    first_label_ref = ""
    for (row, col), (label_ref, raw_label) in sorted(cells.items()):
        normalized = normalize_label(raw_label)
        inline_match = re.match(r"^([^:]{1,80}):\s*(.+)$", raw_label.strip(), flags=re.S)
        if inline_match and normalize_label(inline_match.group(1)) in labels:
            return inline_match.group(2).strip(), label_ref, label_ref
        if normalized not in labels:
            continue
        first_label_ref = first_label_ref or label_ref
        for candidate_col in range(col + 1, min(col + 7, 25)):
            candidate = cells.get((row, candidate_col))
            if not candidate:
                continue
            value_ref, value = candidate
            if normalize_label(value) not in POSITION_LABELS | GROUP_LABELS:
                return value, label_ref, value_ref
    return "", first_label_ref, ""


def is_generic_top_left(value: str) -> bool:
    normalized = normalize_label(value)
    return (
        not normalized
        or normalized in GENERIC_TOP_LEFT
        or normalized in POSITION_LABELS | GROUP_LABELS
        or normalized in GENERIC_KAMUS_TAB_TITLES
    )


def extract_position(
    cells: dict[tuple[int, int], tuple[str, str]], sheet_name: str
) -> tuple[str, str, str, str]:
    if SUPPORT_SHEET_RE.match(sheet_name.strip()):
        return "", "", "", "support_sheet_excluded"
    position, label_ref, value_ref = extract_labeled_value(cells, POSITION_LABELS)
    if position:
        return position, label_ref, value_ref, "labeled_value"
    if label_ref:
        return "", label_ref, "", "labeled_value_missing"
    for row in (1, 2):
        top_left = cells.get((row, 1))
        if not top_left:
            continue
        ref, value = top_left
        if is_generic_top_left(value):
            continue
        return value, "", ref, "top_left_title"
    # Group 2 Subholding/Cabang templates often start the KPI grid at A1
    # (BSC Perspective / Jenis Posisi) and keep the position title only on the tab.
    a1 = cells.get((1, 1), ("", ""))[1]
    a2 = cells.get((2, 1), ("", ""))[1]
    tab = sheet_name.strip()
    if (
        tab
        and not SUPPORT_SHEET_RE.match(tab)
        and normalize_label(tab) not in GENERIC_KAMUS_TAB_TITLES
        and not is_generic_top_left(tab)
        and (not a1 or is_generic_top_left(a1))
        and (not a2 or is_generic_top_left(a2))
    ):
        return tab, "", "", "sheet_tab_title"
    return "", "", "", "not_found"


def generation_for(relative_path: Path) -> str:
    return (
        "v1_pre_restructure"
        if relative_path.parts and relative_path.parts[0] == "KAMUS KPI HO PRE-RESTRUCTURE"
        else "v2"
    )


def is_source_workbook(path: Path, root: Path) -> bool:
    if path.name.startswith("~$"):
        return False
    relative = path.relative_to(root)
    if any("(Original)" in part for part in relative.parts[:-1]):
        return False
    if len(relative.parts) == 1 and relative.name == "Unit Kerja Pelindo per April 2026 - edit.xlsx":
        return False
    return True


def extract(root: Path) -> dict:
    result = {
        "metadata": {
            "source_root": str(root),
            "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
            "scope": "visible worksheets only; source workbooks unchanged",
            "classification_rule": {
                "v1_pre_restructure": "Workbook under KAMUS KPI HO PRE-RESTRUCTURE",
                "v2": "Other KPI workbooks in the download package",
            },
            "excluded_files": [
                "Unit Kerja Pelindo per April 2026 - edit.xlsx",
                "Archive folders whose name contains (Original)",
            ],
        },
        "kamus_kpi_v2": [],
        "kamus_kpi_v1_pre_restructure": [],
    }
    workbooks = sorted(
        (
            path
            for path in root.rglob("*")
            if path.is_file()
            and path.suffix.casefold() in {".xlsx", ".xlsm"}
            and is_source_workbook(path, root)
        ),
        key=lambda path: str(path.relative_to(root)).casefold(),
    )
    workbook_counts = {"v2": 0, "v1_pre_restructure": 0}
    hidden_counts = {"v2": 0, "v1_pre_restructure": 0}
    for workbook_path in workbooks:
        relative = workbook_path.relative_to(root)
        generation = generation_for(relative)
        workbook_counts[generation] += 1
        try:
            with zipfile.ZipFile(workbook_path) as archive:
                strings = shared_strings(archive)
                for sheet in workbook_sheets(archive):
                    if sheet["visibility"] != "visible":
                        hidden_counts[generation] += 1
                        continue
                    cells = sheet_cells(archive, str(sheet["xml_path"]), strings)
                    position, position_label_cell, position_value_cell, extraction_method = extract_position(
                        cells, str(sheet["sheet_name"])
                    )
                    is_support_sheet = bool(SUPPORT_SHEET_RE.match(str(sheet["sheet_name"]).strip()))
                    if is_support_sheet:
                        group_name, group_label_cell, group_value_cell = "", "", ""
                    else:
                        group_name, group_label_cell, group_value_cell = extract_labeled_value(
                            cells, GROUP_LABELS
                        )
                    entry = {
                        "generation": generation,
                        "source_folder": relative.parts[0] if len(relative.parts) > 1 else "",
                        "source_workbook": relative.as_posix(),
                        "sheet_order": sheet["sheet_order"],
                        "sheet_name": sheet["sheet_name"],
                        "visibility": sheet["visibility"],
                        "position_name": position,
                        "position_label_cell": position_label_cell,
                        "position_value_cell": position_value_cell,
                        "position_extraction_method": extraction_method,
                        "group_name": group_name,
                        "group_label_cell": group_label_cell,
                        "group_value_cell": group_value_cell,
                        "include_in_position_config": bool(position),
                        "review_status": (
                            "ready"
                            if position
                            else (
                                "visible_non_position_sheet"
                                if is_support_sheet
                                else "position_title_missing_in_sheet"
                            )
                        ),
                    }
                    target = (
                        result["kamus_kpi_v2"]
                        if generation == "v2"
                        else result["kamus_kpi_v1_pre_restructure"]
                    )
                    target.append(entry)
        except (zipfile.BadZipFile, KeyError, ET.ParseError) as error:
            target = (
                result["kamus_kpi_v2"]
                if generation == "v2"
                else result["kamus_kpi_v1_pre_restructure"]
            )
            target.append(
                {
                    "generation": generation,
                    "source_folder": relative.parts[0] if len(relative.parts) > 1 else "",
                    "source_workbook": relative.as_posix(),
                    "sheet_order": None,
                    "sheet_name": "",
                    "visibility": "unreadable",
                    "position_name": "",
                    "position_label_cell": "",
                    "position_value_cell": "",
                    "position_extraction_method": "workbook_read_error",
                    "group_name": "",
                    "group_label_cell": "",
                    "group_value_cell": "",
                    "include_in_position_config": False,
                    "review_status": f"workbook_read_error: {type(error).__name__}: {error}",
                }
            )
    result["metadata"]["counts"] = {
        "workbooks_v2": workbook_counts["v2"],
        "workbooks_v1_pre_restructure": workbook_counts["v1_pre_restructure"],
        "visible_worksheets_v2": len(result["kamus_kpi_v2"]),
        "visible_worksheets_v1_pre_restructure": len(result["kamus_kpi_v1_pre_restructure"]),
        "position_worksheets_v2": sum(
            1 for row in result["kamus_kpi_v2"] if row["include_in_position_config"]
        ),
        "position_worksheets_v1_pre_restructure": sum(
            1
            for row in result["kamus_kpi_v1_pre_restructure"]
            if row["include_in_position_config"]
        ),
        "visible_non_position_or_unresolved_v2": sum(
            1 for row in result["kamus_kpi_v2"] if not row["include_in_position_config"]
        ),
        "visible_non_position_or_unresolved_v1_pre_restructure": sum(
            1
            for row in result["kamus_kpi_v1_pre_restructure"]
            if not row["include_in_position_config"]
        ),
        "visible_non_position_sheets_v2": sum(
            1
            for row in result["kamus_kpi_v2"]
            if row["review_status"] == "visible_non_position_sheet"
        ),
        "visible_non_position_sheets_v1_pre_restructure": sum(
            1
            for row in result["kamus_kpi_v1_pre_restructure"]
            if row["review_status"] == "visible_non_position_sheet"
        ),
        "position_title_missing_v2": sum(
            1
            for row in result["kamus_kpi_v2"]
            if row["review_status"] == "position_title_missing_in_sheet"
        ),
        "position_title_missing_v1_pre_restructure": sum(
            1
            for row in result["kamus_kpi_v1_pre_restructure"]
            if row["review_status"] == "position_title_missing_in_sheet"
        ),
        "excluded_hidden_worksheets_v2": hidden_counts["v2"],
        "excluded_hidden_worksheets_v1_pre_restructure": hidden_counts["v1_pre_restructure"],
    }
    return result


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()
    result = extract(args.root.resolve())
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(result["metadata"]["counts"], ensure_ascii=False, indent=2))
    for key in ("kamus_kpi_v2", "kamus_kpi_v1_pre_restructure"):
        unresolved = [
            {
                "source_workbook": row["source_workbook"],
                "sheet_name": row["sheet_name"],
                "review_status": row["review_status"],
            }
            for row in result[key]
            if not row["include_in_position_config"]
        ]
        print(json.dumps({key + "_unresolved": unresolved}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
