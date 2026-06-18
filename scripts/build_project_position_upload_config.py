#!/usr/bin/env python3
"""Build a KPI upload config for project-specific positions."""

from __future__ import annotations

import argparse
import csv
import json
import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


SOURCE_WORKBOOK = (
    "KAMUS KPI HO PRE-RESTRUCTURE/"
    "DIREKTORAT TEKNIK - Ibu Ika Oktania - Pengendalian Proyek (Selesai konfirmasi KPI).xlsx"
)

BASE_SHEET_BY_FAMILY = {
    "pimpinan_proyek": "Pimpinan Proyek",
    "deputy_administrasi": "Deputy Administrasi",
    "deputy_konstruksi": "Deputy Konstruksi",
    "deputy_perencanaan": "Deputy Perencanaan Proyek",
    "manager_administrasi": "Manager Administrasi",
    "manager_konstruksi": "Manager Konstruksi",
    "manager_perencanaan": "Manager Perencanaan Proyek",
    "officer_administrasi": "Officer Administrasi",
    "officer_konstruksi": "Officer Konstruksi",
    "officer_perencanaan": "Officer Perencanaan Proyek",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--reference", required=True, type=Path)
    parser.add_argument("--output-config", required=True, type=Path)
    parser.add_argument("--audit-output", required=True, type=Path)
    return parser.parse_args()


def norm(value: object) -> str:
    text = str(value or "").lower()
    text = text.replace("&", " dan ")
    text = re.sub(r"[-_/(),.]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def is_active_master_row(row: dict[str, Any]) -> bool:
    return row.get("is_position_active") == 1 and row.get("is_position_organization_active") == 1


def is_project_position(row: dict[str, Any]) -> bool:
    if norm(row.get("company_name")) != "pt pelabuhan indonesia persero":
        return False
    text = " ".join(
        norm(row.get(key))
        for key in ["position_name", "group_name", "company_name"]
    )
    if "proyek" not in text:
        return False
    project_tokens = [
        "bali maritime tourism hub",
        "bmth",
        "terminal kalibaru",
        "npea",
        "jict koja",
        "kijing",
    ]
    return any(token in text for token in project_tokens)


def classify_family(position_name: str) -> str | None:
    text = norm(position_name)
    if "pimpinan proyek" in text or "pimpro" in text and not text.startswith("deput"):
        return "pimpinan_proyek"
    if "deput" in text:
        if "administrasi" in text:
            return "deputy_administrasi"
        if "konstruksi" in text:
            return "deputy_konstruksi"
        if "perencanaan" in text:
            return "deputy_perencanaan"
    if "manager" in text or "manajer" in text:
        if "administrasi" in text:
            return "manager_administrasi"
        if "konstruksi" in text:
            return "manager_konstruksi"
        if "perencanaan" in text:
            return "manager_perencanaan"
    if "officer" in text or "administrator" in text:
        if "administrasi" in text:
            return "officer_administrasi"
        if "konstruksi" in text:
            return "officer_konstruksi"
        if "perencanaan" in text:
            return "officer_perencanaan"
    return None


def project_name_from_group(group_name: str) -> str:
    text = group_name
    replacements = [
        ("Bidang Perencanaan Proyek ", ""),
        ("Bidang Administrasi Proyek ", ""),
        ("Bidang Konstruksi 1 Proyek ", ""),
        ("Bidang Konstruksi 2 Proyek ", ""),
        ("Bidang Konstruksi 3 Proyek ", ""),
        ("Bidang Konstruksi Proyek ", ""),
        ("Bidang Konstruksi/Teknik Sipil/ME/IT Proyek ", ""),
        ("Subbidang Perencanaan Proyek ", ""),
        ("Subbidang Administrasi Proyek ", ""),
        ("Subbidang Konstruksi Proyek ", ""),
        ("Proyek ", ""),
    ]
    for old, new in replacements:
        text = text.replace(old, new)
    return text.strip()


def build_config_row(row: dict[str, Any], family: str) -> dict[str, Any]:
    sheet_name = BASE_SHEET_BY_FAMILY[family]
    return {
        "source_workbook": SOURCE_WORKBOOK,
        "sheet_name": sheet_name,
        "position_name": row["position_name"],
        "position_master_id": str(row["position_master_id"]),
        "position_nomenclature_id": None,
        "position_scope": "structural",
        "portaverse_position_title": row["position_name"],
        "portaverse_group_name": row.get("group_name"),
        "portaverse_company_name": row.get("company_name"),
        "cluster_label": None,
        "position_lookup_names": [row["position_name"], sheet_name],
        "group_name": row.get("group_name") or project_name_from_group(row.get("group_name") or ""),
        "directorate_name": "DIREKTORAT TEKNIK",
        "expected_impact_count": 10,
        "drop_comment_values": ["Drop"],
    }


def main() -> int:
    args = parse_args()
    with args.reference.open() as handle:
        reference = json.load(handle)

    selected: list[tuple[dict[str, Any], str]] = []
    seen: set[str] = set()
    for row in reference["position_master_rows"]:
        if not is_active_master_row(row) or not is_project_position(row):
            continue
        family = classify_family(row.get("position_name") or "")
        if not family:
            continue
        pmid = str(row["position_master_id"])
        if pmid in seen:
            continue
        seen.add(pmid)
        selected.append((row, family))

    selected.sort(key=lambda item: (item[0].get("group_name") or "", item[0].get("position_name") or ""))
    positions = [build_config_row(row, family) for row, family in selected]
    payload = {
        "reference_source": {
            "config_generated_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
            "file": str(args.reference),
            "source_workbook": SOURCE_WORKBOOK,
            "selection_rule": (
                "Active PT Pelabuhan Indonesia project positions for BMTH, Terminal Kalibaru, "
                "NPEA, JICT KOJA, and Kijing mapped to Pengendalian Proyek base KPI sheets."
            ),
            "selected_positions": len(positions),
        },
        "positions": positions,
    }

    args.output_config.parent.mkdir(parents=True, exist_ok=True)
    args.output_config.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n")

    args.audit_output.parent.mkdir(parents=True, exist_ok=True)
    with args.audit_output.open("w", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=[
                "position_master_id",
                "position_name",
                "group_name",
                "company_name",
                "family",
                "base_sheet",
            ],
        )
        writer.writeheader()
        for row, family in selected:
            writer.writerow(
                {
                    "position_master_id": row["position_master_id"],
                    "position_name": row["position_name"],
                    "group_name": row.get("group_name"),
                    "company_name": row.get("company_name"),
                    "family": family,
                    "base_sheet": BASE_SHEET_BY_FAMILY[family],
                }
            )

    print(f"selected_positions={len(positions)}")
    print(f"output_config={args.output_config}")
    print(f"audit_output={args.audit_output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
