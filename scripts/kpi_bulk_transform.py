#!/usr/bin/env python3
"""Transform KPI design workbooks into the official bulk upload template."""

from __future__ import annotations

import argparse
import csv
import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

try:
    from openpyxl import load_workbook
except ImportError as exc:  # pragma: no cover - runtime guidance
    raise SystemExit(
        "Missing dependency: openpyxl. Install it with `python3 -m pip install openpyxl`."
    ) from exc


XLSX_NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

UPLOAD_HEADERS = [
    "IDKPI",
    "Group",
    "Direktorat",
    "Posisi",
    "Position Master ID (Required)",
    "Position Master Variant ID (Required only for Non-Struktural)",
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
    "Ownership Type",
    "Nature Of Work (KAI Only)",
    "External ID (PKPI)",
]


def norm_text(value: Any) -> str | None:
    if value is None:
        return None
    if isinstance(value, str):
        value = value.replace("\r\n", "\n").replace("\r", "\n").strip()
        return value or None
    return str(value)


def normalize_title(value: str | None) -> str:
    value = (norm_text(value) or "").lower()
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def is_placeholder(value: Any) -> bool:
    text = norm_text(value)
    if text is None:
        return True
    return text.strip().lower() in {"(blank)", "blank"}


def to_upper_enum(value: str | None, mapping: dict[str, str]) -> str | None:
    value = norm_text(value)
    if not value or is_placeholder(value):
        return None
    return mapping.get(value, value.upper())


def uploader_period(value: str | None) -> str | None:
    return to_upper_enum(
        value,
        {
            "Triwulan": "TRIWULANAN",
            "Triwulanan": "TRIWULANAN",
            "Tahunan": "TAHUNAN",
            "Semester": "SEMESTER",
            "Bulanan": "BULANAN",
            "Monthly": "MONTHLY",
            "Quarterly": "QUARTERLY",
            "Weekly": "WEEKLY",
        },
    )


def uploader_polarity(value: str | None) -> str | None:
    return to_upper_enum(
        value,
        {
            "Positif": "POSITIVE",
            "Negatif": "NEGATIVE",
            "Netral": "NEUTRAL",
            "Positive": "POSITIVE",
            "Negative": "NEGATIVE",
            "Neutral": "NEUTRAL",
        },
    )


def col_to_num(col_ref: str) -> int:
    value = 0
    for ch in col_ref:
        if ch.isalpha():
            value = (value * 26) + (ord(ch.upper()) - 64)
    return value


def read_xlsx_sheet(path: Path, sheet_name: str) -> list[list[Any]]:
    with zipfile.ZipFile(path) as workbook_zip:
        workbook = ET.fromstring(workbook_zip.read("xl/workbook.xml"))
        rels = ET.fromstring(workbook_zip.read("xl/_rels/workbook.xml.rels"))
        relmap = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}

        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in workbook_zip.namelist():
            root = ET.fromstring(workbook_zip.read("xl/sharedStrings.xml"))
            for si in root:
                shared_strings.append(
                    "".join(
                        t.text or ""
                        for t in si.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")
                    )
                )

        target = None
        for sheet in workbook.find("a:sheets", XLSX_NS):
            if sheet.attrib["name"] == sheet_name:
                rel_target = relmap[
                    sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
                ]
                target = f"xl/{rel_target}" if not rel_target.startswith("xl/") else rel_target
                break
        if not target:
            raise KeyError(f"Sheet '{sheet_name}' not found in {path}")

        root = ET.fromstring(workbook_zip.read(target))
        rows = root.find("a:sheetData", XLSX_NS)
        parsed_rows: list[list[Any]] = []
        for row in rows:
            values: dict[int, Any] = {}
            for cell in row:
                match = re.match(r"([A-Z]+)(\d+)", cell.attrib.get("r", ""))
                if not match:
                    continue
                col_num = col_to_num(match.group(1))
                cell_type = cell.attrib.get("t")
                raw_value = cell.find("a:v", XLSX_NS)
                inline = cell.find("a:is", XLSX_NS)
                value = None
                if cell_type == "s" and raw_value is not None:
                    value = shared_strings[int(raw_value.text)]
                elif cell_type == "inlineStr" and inline is not None:
                    value = "".join(
                        t.text or ""
                        for t in inline.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")
                    )
                elif raw_value is not None:
                    value = raw_value.text
                values[col_num] = norm_text(value)
            max_col = max(values) if values else 0
            parsed_rows.append([values.get(i) for i in range(1, max_col + 1)])
        return parsed_rows


@dataclass
class PositionMetadata:
    position_master_id: str
    position_name: str
    organization_name: str | None
    company_name: str | None
    position_type: str | None


@dataclass
class ValidationIssue:
    severity: str
    sheet_name: str
    source_row: int | None
    record_type: str
    title: str | None
    message: str


@dataclass
class PositionConfig:
    sheet_name: str
    position_name: str
    group_name: str
    directorate_name: str
    position_master_id: str | None = None
    position_lookup_names: list[str] = field(default_factory=list)
    drop_comment_values: list[str] = field(default_factory=lambda: ["Drop"])
    expected_impact_count: int = 10


@dataclass
class ImpactRecord:
    bsc: str | None
    title: str
    unit: str | None
    period: str | None
    formula: str | None
    polarity: str | None
    weight: str | None
    outputs: list[dict[str, Any]] = field(default_factory=list)


@dataclass
class ParsedSheet:
    config: PositionConfig
    metadata: PositionMetadata | None
    impacts: list[ImpactRecord]


def backfill_shared_impact_fields(parsed_sheets: list[ParsedSheet]) -> None:
    canonical: dict[str, dict[str, str | None]] = {}
    for parsed in parsed_sheets:
        for impact in parsed.impacts:
            key = normalize_title(impact.title)
            fields = canonical.setdefault(
                key,
                {
                    "bsc": None,
                    "unit": None,
                    "period": None,
                    "formula": None,
                    "polarity": None,
                    "weight": None,
                },
            )
            for field_name in fields:
                current_value = getattr(impact, field_name)
                if not fields[field_name] and norm_text(current_value) and not is_placeholder(current_value):
                    fields[field_name] = current_value

    for parsed in parsed_sheets:
        for impact in parsed.impacts:
            fields = canonical.get(normalize_title(impact.title), {})
            for field_name, fallback in fields.items():
                current_value = getattr(impact, field_name)
                if (not norm_text(current_value) or is_placeholder(current_value)) and fallback:
                    setattr(impact, field_name, fallback)


class PositionMasterIndex:
    def __init__(self) -> None:
        self.by_id: dict[str, PositionMetadata] = {}
        self.by_title: dict[str, PositionMetadata] = {}

    @classmethod
    def load(cls, root_dir: Path) -> "PositionMasterIndex":
        index = cls()
        files = sorted(root_dir.glob("*.xlsx"))
        for workbook_path in files:
            rows = read_xlsx_sheet(workbook_path, "Master Posisi")
            if not rows:
                continue
            for row in rows[1:]:
                row = row + [None] * max(0, 8 - len(row))
                position_id = norm_text(row[0])
                position_name = norm_text(row[1])
                if not position_id or not position_name:
                    continue
                metadata = PositionMetadata(
                    position_master_id=position_id,
                    position_name=position_name,
                    organization_name=norm_text(row[2]) if len(row) > 2 else None,
                    company_name=norm_text(row[3]) if len(row) > 3 else None,
                    position_type=norm_text(row[7]) if len(row) > 7 else None,
                )
                index.by_id[position_id] = metadata
                index.by_title[normalize_title(position_name)] = metadata
        return index

    def resolve(self, config: PositionConfig) -> PositionMetadata | None:
        if config.position_master_id:
            metadata = self.by_id.get(str(config.position_master_id))
            if metadata:
                return metadata
        lookup_names = [config.position_name, *config.position_lookup_names]
        for name in lookup_names:
            metadata = self.by_title.get(normalize_title(name))
            if metadata:
                return metadata
        return None


def find_header_row(rows: list[list[Any]]) -> tuple[int, dict[str, int]]:
    for index, row in enumerate(rows, start=1):
        headers = {norm_text(value): idx for idx, value in enumerate(row, start=1) if norm_text(value)}
        if "KPI Impact" in headers and "KPI Output" in headers and "Key Activity Indicator (KAI)" in headers:
            return index, headers
    raise ValueError("Could not find KPI header row")


def row_value(row: list[Any], header_map: dict[str, int], header: str) -> str | None:
    col = header_map.get(header)
    if not col or col - 1 >= len(row):
        return None
    return norm_text(row[col - 1])


def parse_block_sheet(
    rows: list[list[Any]],
    config: PositionConfig,
    issues: list[ValidationIssue],
) -> list[ImpactRecord]:
    header_row, header_map = find_header_row(rows)
    comment_header = "Komentar" if "Komentar" in header_map else None

    impacts: list[ImpactRecord] = []
    current_impact: ImpactRecord | None = None
    current_bsc: str | None = None
    current_impact_defaults: dict[str, str | None] = {}
    current_output_defaults: dict[str, str | None] = {}

    for source_row in range(header_row + 1, len(rows) + 1):
        row = rows[source_row - 1]
        first_cell = norm_text(row[0]) if row else None
        impact_title = row_value(row, header_map, "KPI Impact")
        output_title = row_value(row, header_map, "KPI Output")
        kai_title = row_value(row, header_map, "Key Activity Indicator (KAI)")

        if first_cell == "TOTAL":
            break
        if not any(norm_text(value) for value in row):
            continue

        bsc = row_value(row, header_map, "BSC Perspective")
        if bsc:
            current_bsc = bsc

        if impact_title and not is_placeholder(impact_title):
            impact_unit = row_value(row, header_map, "KPI Impact Unit") or current_impact_defaults.get("unit")
            impact_period = row_value(row, header_map, "KPI Impact Frequency") or current_impact_defaults.get("period")
            impact_formula = row_value(row, header_map, "KPI Impact Formula") or current_impact_defaults.get("formula")
            impact_polarity = row_value(row, header_map, "KPI Impact Polarity") or current_impact_defaults.get("polarity")
            impact_weight = row_value(row, header_map, "%Weight (Impact)") or current_impact_defaults.get("weight")
            current_impact = ImpactRecord(
                bsc=current_bsc,
                title=impact_title,
                unit=impact_unit,
                period=impact_period,
                formula=impact_formula,
                polarity=impact_polarity,
                weight=impact_weight,
            )
            current_impact_defaults = {
                "unit": impact_unit,
                "period": impact_period,
                "formula": impact_formula,
                "polarity": impact_polarity,
                "weight": impact_weight,
            }
            impacts.append(current_impact)
        elif current_impact is None:
            issues.append(
                ValidationIssue(
                    severity="error",
                    sheet_name=config.sheet_name,
                    source_row=source_row,
                    record_type="row",
                    title=None,
                    message="Row appears before any KPI Impact block and cannot inherit a parent impact.",
                )
            )
            continue

        comment = row_value(row, header_map, comment_header) if comment_header else None
        drop_comment = norm_text(comment) in set(config.drop_comment_values)

        output_weight = row_value(row, header_map, "%Weight (Output)")
        kai_weight = row_value(row, header_map, "%Weight (Activity)")

        output_period = row_value(row, header_map, "KPI Output Frequency") or current_output_defaults.get("period")
        output_polarity = row_value(row, header_map, "KPI Output Polarity") or current_output_defaults.get("polarity")
        output_unit = row_value(row, header_map, "KPI Output Unit") or current_output_defaults.get("unit")
        output_formula = row_value(row, header_map, "KPI Output Formula") or current_output_defaults.get("formula")
        output_definition = row_value(row, header_map, "KPI Output Definition") or current_output_defaults.get("description")
        cascading_output = row_value(row, header_map, "Cascading Tagging (KPI Output)") or current_output_defaults.get("cascading")
        coverage_output = row_value(row, header_map, "Coverage KPI Output") or current_output_defaults.get("ownership_type")
        nature_of_work = row_value(row, header_map, "Nature of Work (KAI)") or current_output_defaults.get("nature_of_work")

        keep_output = (
            not is_placeholder(output_title)
            and not drop_comment
            and not is_placeholder(output_weight)
        )
        keep_kai = (
            keep_output
            and not is_placeholder(kai_title)
            and not drop_comment
            and not is_placeholder(kai_weight)
        )

        if output_title and not keep_output:
            reasons = []
            if drop_comment:
                reasons.append("comment is Drop")
            if is_placeholder(output_weight):
                reasons.append("output weight is blank")
            issues.append(
                ValidationIssue(
                    severity="info",
                    sheet_name=config.sheet_name,
                    source_row=source_row,
                    record_type="output",
                    title=output_title,
                    message=f"Dropped OUTPUT row because {' and '.join(reasons)}.",
                )
            )

        if keep_output:
            current_output_defaults = {
                "period": output_period,
                "polarity": output_polarity,
                "unit": output_unit,
                "formula": output_formula,
                "description": output_definition,
                "cascading": cascading_output,
                "ownership_type": coverage_output,
                "nature_of_work": nature_of_work,
            }
            output_record = {
                "source_row": source_row,
                "title": output_title,
                "description": output_definition,
                "unit": output_unit,
                "period": output_period,
                "formula": output_formula,
                "polarity": output_polarity,
                "weight": output_weight,
                "cascading": cascading_output,
                "ownership_type": coverage_output,
                "kai": None,
            }
            current_impact.outputs.append(output_record)

            if kai_title and not keep_kai:
                reasons = []
                if drop_comment:
                    reasons.append("comment is Drop")
                if is_placeholder(kai_weight):
                    reasons.append("KAI weight is blank")
                issues.append(
                    ValidationIssue(
                        severity="info",
                        sheet_name=config.sheet_name,
                        source_row=source_row,
                        record_type="kai",
                        title=kai_title,
                        message=f"Dropped KAI row because {' and '.join(reasons)}.",
                    )
                )

            if keep_kai:
                output_record["kai"] = {
                    "source_row": source_row,
                    "title": kai_title,
                    "description": row_value(row, header_map, "KPI KAI Definition"),
                    "formula": row_value(row, header_map, "KPI KAI Formula"),
                    "weight": kai_weight,
                    "nature_of_work": nature_of_work,
                    "period": output_period,
                    "polarity": output_polarity,
                    "cascading": cascading_output,
                    "ownership_type": coverage_output,
                }
                if output_period:
                    issues.append(
                        ValidationIssue(
                            severity="info",
                            sheet_name=config.sheet_name,
                            source_row=source_row,
                            record_type="kai",
                            title=kai_title,
                            message="KAI period inferred from KPI Output Frequency because the source sheet does not provide a separate KAI period column.",
                        )
                    )
        elif kai_title and not is_placeholder(kai_title):
            issues.append(
                ValidationIssue(
                    severity="warning",
                    sheet_name=config.sheet_name,
                    source_row=source_row,
                    record_type="kai",
                    title=kai_title,
                    message="Skipped KAI because its OUTPUT row was dropped or missing.",
                )
            )

    return impacts


def build_upload_rows(
    config: PositionConfig,
    metadata: PositionMetadata | None,
    impacts: list[ImpactRecord],
    start_id: int,
) -> tuple[list[list[Any]], int]:
    position_master_id = metadata.position_master_id if metadata else str(config.position_master_id)
    position_name = config.position_name

    rows: list[list[Any]] = []
    next_id = start_id
    impact_ids: dict[str, str] = {}

    for impact in impacts:
        impact_id = str(next_id)
        next_id += 1
        impact_ids[impact.title] = impact_id
        rows.append(
            [
                impact_id,
                config.group_name,
                config.directorate_name,
                position_name,
                position_master_id,
                None,
                impact.bsc,
                "IMPACT",
                None,
                "#N/A",
                impact.title,
                None,
                impact.unit,
                uploader_polarity(impact.polarity),
                uploader_period(impact.period),
                impact.formula,
                impact.weight,
                None,
                None,
                None,
                None,
            ]
        )

    for impact in impacts:
        for output in impact.outputs:
            output_id = str(next_id)
            next_id += 1
            output["_generated_id"] = output_id
            rows.append(
                [
                    output_id,
                    config.group_name,
                    config.directorate_name,
                    position_name,
                    position_master_id,
                    None,
                    impact.bsc,
                    "OUTPUT",
                    impact_ids[impact.title],
                    impact.title,
                    output["title"],
                    output["description"],
                    output["unit"],
                    uploader_polarity(output["polarity"]),
                    uploader_period(output["period"]),
                    output["formula"],
                    output["weight"],
                    output["cascading"],
                    output["ownership_type"],
                    None,
                    None,
                ]
            )

    for impact in impacts:
        for output in impact.outputs:
            kai = output.get("kai")
            if not kai:
                continue
            rows.append(
                [
                    str(next_id),
                    config.group_name,
                    config.directorate_name,
                    position_name,
                    position_master_id,
                    None,
                    impact.bsc,
                    "KAI",
                    output["_generated_id"],
                    output["title"],
                    kai["title"],
                    kai["description"],
                    None,
                    uploader_polarity(kai["polarity"]),
                    uploader_period(kai["period"]),
                    kai["formula"],
                    kai["weight"],
                    kai["cascading"],
                    kai["ownership_type"],
                    kai["nature_of_work"],
                    None,
                ]
            )
            next_id += 1

    return rows, next_id


def validate_output_rows(
    config: PositionConfig,
    rows: list[list[Any]],
    issues: list[ValidationIssue],
) -> None:
    all_ids = {str(row[0]) for row in rows}
    for row in rows:
        row_map = dict(zip(UPLOAD_HEADERS, row))
        title = row_map["Title"]
        record_type = row_map["KPI Type"]
        for key in [
            "IDKPI",
            "KPI Type",
            "Title",
            "Position Master ID (Required)",
            "Polarity",
            "Weight (%)",
        ]:
            if not norm_text(row_map.get(key)):
                issues.append(
                    ValidationIssue(
                        severity="error",
                        sheet_name=config.sheet_name,
                        source_row=None,
                        record_type=record_type or "row",
                        title=title,
                        message=f"Missing required upload field: {key}",
                    )
                )
        if record_type == "OUTPUT":
            if not norm_text(row_map["BSC Perspective"]):
                issues.append(
                    ValidationIssue(
                        severity="error",
                        sheet_name=config.sheet_name,
                        source_row=None,
                        record_type="OUTPUT",
                        title=title,
                        message="OUTPUT row is missing BSC Perspective.",
                    )
                )
            if not norm_text(row_map["Parent KPI ID"]):
                issues.append(
                    ValidationIssue(
                        severity="error",
                        sheet_name=config.sheet_name,
                        source_row=None,
                        record_type="OUTPUT",
                        title=title,
                        message="OUTPUT row is missing Parent KPI ID.",
                    )
                )
        if record_type == "KAI":
            if not norm_text(row_map["Nature Of Work (KAI Only)"]):
                issues.append(
                    ValidationIssue(
                        severity="error",
                        sheet_name=config.sheet_name,
                        source_row=None,
                        record_type="KAI",
                        title=title,
                        message="KAI row is missing Nature Of Work.",
                    )
                )
            if not norm_text(row_map["Parent KPI ID"]):
                issues.append(
                    ValidationIssue(
                        severity="error",
                        sheet_name=config.sheet_name,
                        source_row=None,
                        record_type="KAI",
                        title=title,
                        message="KAI row is missing Parent KPI ID.",
                    )
                )
        parent_id = norm_text(row_map["Parent KPI ID"])
        if record_type in {"OUTPUT", "KAI"} and parent_id and parent_id not in all_ids:
            issues.append(
                ValidationIssue(
                    severity="error",
                    sheet_name=config.sheet_name,
                    source_row=None,
                    record_type=record_type,
                    title=title,
                    message=f"Parent KPI ID {parent_id} does not exist in generated rows.",
                )
            )


def write_output_workbook(template_path: Path, output_path: Path, rows: list[list[Any]]) -> None:
    workbook = load_workbook(template_path)
    worksheet = workbook["KPI Template"]
    if worksheet.max_row > 1:
        worksheet.delete_rows(2, worksheet.max_row - 1)
    for row_index, row_values in enumerate(rows, start=2):
        for col_index, value in enumerate(row_values, start=1):
            worksheet.cell(row=row_index, column=col_index, value=value)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)


def write_report(report_path: Path, issues: list[ValidationIssue]) -> None:
    report_path.parent.mkdir(parents=True, exist_ok=True)
    with report_path.open("w", newline="", encoding="utf-8") as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(["severity", "sheet_name", "source_row", "record_type", "title", "message"])
        for issue in issues:
            writer.writerow(
                [
                    issue.severity,
                    issue.sheet_name,
                    issue.source_row,
                    issue.record_type,
                    issue.title,
                    issue.message,
                ]
            )


def load_config(config_path: Path) -> list[PositionConfig]:
    data = json.loads(config_path.read_text(encoding="utf-8"))
    configs = []
    for item in data["positions"]:
        configs.append(
            PositionConfig(
                sheet_name=item["sheet_name"],
                position_name=item["position_name"],
                group_name=item["group_name"],
                directorate_name=item["directorate_name"],
                position_master_id=str(item["position_master_id"]) if item.get("position_master_id") is not None else None,
                position_lookup_names=item.get("position_lookup_names", []),
                drop_comment_values=item.get("drop_comment_values", ["Drop"]),
                expected_impact_count=int(item.get("expected_impact_count", 10)),
            )
        )
    return configs


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", required=True, type=Path, help="Source KPI workbook (.xlsx)")
    parser.add_argument("--template", required=True, type=Path, help="Official upload template (.xlsx)")
    parser.add_argument("--positions-dir", required=True, type=Path, help="Directory with data_master_posisi xlsx exports")
    parser.add_argument("--config", required=True, type=Path, help="JSON config describing sheets to export")
    parser.add_argument("--output", required=True, type=Path, help="Output upload workbook (.xlsx)")
    parser.add_argument("--report", required=True, type=Path, help="Validation report (.csv)")
    parser.add_argument(
        "--only-sheet",
        action="append",
        default=[],
        help="Limit export to a specific sheet name. May be passed multiple times.",
    )
    args = parser.parse_args()

    configs = load_config(args.config)
    if args.only_sheet:
        selected = set(args.only_sheet)
        configs = [cfg for cfg in configs if cfg.sheet_name in selected]

    master_index = PositionMasterIndex.load(args.positions_dir)
    issues: list[ValidationIssue] = []
    parsed_sheets: list[ParsedSheet] = []
    output_rows: list[list[Any]] = []
    next_global_id = 1

    for config in configs:
        metadata = master_index.resolve(config)
        if metadata is None:
            issues.append(
                ValidationIssue(
                    severity="error",
                    sheet_name=config.sheet_name,
                    source_row=None,
                    record_type="sheet",
                    title=config.position_name,
                    message="Could not resolve position metadata from the master position exports.",
                )
            )
            continue
        if config.position_master_id and metadata.position_master_id != str(config.position_master_id):
            issues.append(
                ValidationIssue(
                    severity="warning",
                    sheet_name=config.sheet_name,
                    source_row=None,
                    record_type="sheet",
                    title=config.position_name,
                    message=(
                        f"Config position_master_id={config.position_master_id} differs from master lookup "
                        f"{metadata.position_master_id}; using the config value in output."
                    ),
                )
            )
        if metadata.position_type and metadata.position_type != "Struktural":
            issues.append(
                ValidationIssue(
                    severity="warning",
                    sheet_name=config.sheet_name,
                    source_row=None,
                    record_type="sheet",
                    title=config.position_name,
                    message=(
                        f"Master position type is {metadata.position_type}; uploader may require "
                        "Position Master Variant ID for non-struktural roles."
                    ),
                )
            )

        sheet_rows = read_xlsx_sheet(args.source, config.sheet_name)
        impacts = parse_block_sheet(sheet_rows, config, issues)
        if len(impacts) != config.expected_impact_count:
            issues.append(
                ValidationIssue(
                    severity="warning",
                    sheet_name=config.sheet_name,
                    source_row=None,
                    record_type="sheet",
                    title=config.position_name,
                    message=(
                        f"Parsed {len(impacts)} KPI Impact rows; expected "
                        f"{config.expected_impact_count} shared Pelindo impacts."
                    ),
                )
            )
        parsed_sheets.append(ParsedSheet(config=config, metadata=metadata, impacts=impacts))

    backfill_shared_impact_fields(parsed_sheets)

    for parsed in parsed_sheets:
        rows, next_global_id = build_upload_rows(
            parsed.config,
            parsed.metadata,
            parsed.impacts,
            next_global_id,
        )
        validate_output_rows(parsed.config, rows, issues)
        output_rows.extend(rows)

    write_output_workbook(args.template, args.output, output_rows)
    write_report(args.report, issues)

    errors = sum(1 for issue in issues if issue.severity == "error")
    warnings = sum(1 for issue in issues if issue.severity == "warning")
    infos = sum(1 for issue in issues if issue.severity == "info")
    print(f"Wrote workbook: {args.output}")
    print(f"Wrote report: {args.report}")
    print(f"Generated rows: {len(output_rows)}")
    print(f"Issues: errors={errors} warnings={warnings} info={infos}")
    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
