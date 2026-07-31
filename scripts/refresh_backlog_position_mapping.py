#!/usr/bin/env python3
"""Refresh backlog position mapping against an active production reference snapshot."""

from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

import kpi_bulk_transform as transform
import position_mapping

DEFAULT_INPUT = ROOT / "outputs/backlog-kamus-kpi-mapping-20260727/support/backlog_target_positions.config.json"
DEFAULT_REFERENCE = ROOT / "configs/production_position_reference.json"
DEFAULT_OUTPUT_DIR = ROOT / "outputs/backlog-kamus-kpi-mapping-20260727"
PENGENDALIAN_INV = DEFAULT_OUTPUT_DIR / "support/pengendalian_proyek_visible_20260727.json"
KEBERLANJUTAN_CONFIG = ROOT / "output/group_keberlanjutan_5_positions_20260727/support/group_keberlanjutan_5_positions.config.json"

PENGENDALIAN_PIMPINAN_PMID: list[tuple[tuple[str, ...], str, str]] = [
    (("batang",), "23088", "Pimpinan Proyek Investasi Pelabuhan Batang"),
    (("kalibaru", "npea"), "35775", "Pimpinan Proyek Investasi Terminal Kalibaru dan NPEA"),
    (("kijing",), "37585", "Pimpinan Proyek Investasi Kijing"),
    (("jict", "koja"), "35828", "Pimpinan Proyek Investasi JICT Koja"),
    (("bmth",), "37579", "Pimpinan Proyek Invetasi Bali Maritime Tourism Hub"),
]

PRODUCTION_VALIDATED_OVERRIDES: dict[tuple[str, str], dict[str, str | None]] = {
    (
        "Group Layanan Strategis SDM/DIREKTORAT SDM & UMUM - Group Layanan Strategis SDM.xlsx",
        "Officer Pelaporan dan Analitik ",
    ): {
        "scope": "non_structural",
        "position_nomenclature_id": "12566",
        "reason": "Production cluster Officer Manajemen Pelaporan dan Analitik SDM (PNID 12566) matches worksheet Officer Pelaporan dan Analitik SDM.",
    },
    (
        "Group Keberlanjutan Korporasi/DIREKTORAT HUBUNGAN KELEMBAGAAN - Group Keberlanjutan Korporasi.xlsx",
        "Officer Implementasi dan Pelapo",
    ): {
        "scope": "non_structural",
        "position_nomenclature_id": "12533",
        "reason": "Production cluster Officer Implementasi dan Pelaporan Keberlanjutan Korporasi (PNID 12533).",
    },
    (
        "Group Pengadaan/DIREKTORAT WAKIL DIREKTUR UTAMA - Group Pengadaan.xlsx",
        "Officer Pengadaan Teknik",
    ): {
        "scope": "non_structural",
        "position_nomenclature_id": "112",
        "reason": "Exact production cluster match for Officer Pengadaan Teknik (PNID 112).",
    },
    (
        "Group Pengadaan/DIREKTORAT WAKIL DIREKTUR UTAMA - Group Pengadaan.xlsx",
        "Manager Pengadaan Teknik",
    ): {
        "scope": "structural",
        "position_master_id": "753",
        "reason": "Exact production structural match for Manager Pengadaan Teknik (PMID 753).",
    },
    (
        "Group Pengadaan/DIREKTORAT WAKIL DIREKTUR UTAMA - Group Pengadaan.xlsx",
        "Officer Strategi dan Perencanaa",
    ): {
        "scope": "non_structural",
        "position_nomenclature_id": "111",
        "reason": "Exact production cluster match for Officer Strategi Dan Perencanaan Pengadaan (PNID 111).",
    },
    (
        "Group Strategi Korporasi dan Pengembangan Bisnis/DIREKTORAT PENGEMBANGAN USAHA - Group Strategi Korporasi dan Pengembangan Bisnis.xlsx",
        "Officer Perencanaan Strategis ",
    ): {
        "scope": "non_structural",
        "position_nomenclature_id": "12474",
        "reason": "Exact production cluster match for Officer Perencanaan Strategis (PNID 12474).",
    },
    (
        "Group Strategi Korporasi dan Pengembangan Bisnis/DIREKTORAT PENGEMBANGAN USAHA - Group Strategi Korporasi dan Pengembangan Bisnis.xlsx",
        "Group Head",
    ): {
        "scope": "structural",
        "position_master_id": "35743",
        "reason": "Production structural match for Group Head Strategi Korporasi dan Pengembangan Bisnis (PMID 35743).",
    },
    (
        "Group Sekretariat Perusahaan/DIREKTORAT UTAMA - Group Sekretariat Perusahaan.xlsx",
        "Junior Officer Komunikasi Korpo",
    ): {
        "scope": "non_structural",
        "position_nomenclature_id": "12",
        "reason": "Production uses shared cluster Officer Komunikasi Korporasi (PNID 12) for junior tier worksheets.",
    },
    (
        "Group Sekretariat Perusahaan/DIREKTORAT UTAMA - Group Sekretariat Perusahaan.xlsx",
        "Senior Officer Komunikasi Korpo",
    ): {
        "scope": "non_structural",
        "position_nomenclature_id": "12",
        "reason": "Production uses shared cluster Officer Komunikasi Korporasi (PNID 12) for senior tier worksheets.",
    },
}


def norm(value: object) -> str:
    text = str(value or "").lower().strip()
    text = re.sub(r"\bdh\b", "department head", text)
    text = text.replace("&", " dan ")
    text = re.sub(r"[-_/(),.]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def config_key(source_workbook: str | None, sheet_name: str | None) -> tuple[str, str]:
    return (norm(source_workbook), norm(sheet_name))


def load_pengendalian_position_names() -> dict[tuple[str, str], str]:
    if not PENGENDALIAN_INV.exists():
        return {}
    payload = json.loads(PENGENDALIAN_INV.read_text(encoding="utf-8"))
    names: dict[tuple[str, str], str] = {}
    for row in payload.get("kamus_kpi_v2", []):
        workbook = f"Group Pengendalian Proyek/{row.get('source_workbook')}"
        key = config_key(workbook, str(row.get("sheet_name")))
        position_name = str(row.get("position_name") or row.get("sheet_name") or "")
        if position_name:
            names[key] = position_name
    return names


def load_keberlanjutan_approved() -> dict[tuple[str, str], dict[str, Any]]:
    if not KEBERLANJUTAN_CONFIG.exists():
        return {}
    payload = json.loads(KEBERLANJUTAN_CONFIG.read_text(encoding="utf-8"))
    approved: dict[tuple[str, str], dict[str, Any]] = {}
    for row in payload.get("positions", []):
        key = config_key(row.get("source_workbook"), row.get("sheet_name"))
        approved[key] = row
    return approved


def match_pengendalian_pmid(position_name: str, sheet_name: str) -> tuple[str, str] | None:
    combined = norm(f"{position_name} {sheet_name}")
    if combined.strip() == "pimpinan proyek":
        return None
    for tokens, pmid, title in PENGENDALIAN_PIMPINAN_PMID:
        if all(token in combined for token in tokens):
            return pmid, title
    return None


def apply_validated_identity(
    config: transform.PositionConfig,
    scope: str,
    pmid: str | None,
    pnid: str | None,
    reason: str,
    indexes: position_mapping.LookupIndexes,
    trust_source: str = "production_reference",
) -> bool:
    validation = position_mapping.validate_manual_override(
        inferred_scope=scope,
        position_master_id=pmid,
        position_nomenclature_id=pnid,
        indexes=indexes,
    )
    if not validation.allowed:
        return False
    transform.apply_validated_override_candidate(config, validation)
    config.mapping_confidence_label = position_mapping.HIGH_CONFIDENCE
    config.mapping_confidence_reason = reason
    config.mapping_review_status = "approved"
    config.mapping_override_approved = True
    config.mapping_override_trust_source = trust_source
    return True


def apply_keberlanjutan_approved(
    config: transform.PositionConfig,
    row: dict[str, Any],
    indexes: position_mapping.LookupIndexes,
) -> None:
    scope = transform.normalize_position_scope(row.get("position_scope"))
    pmid = str(row.get("position_master_id")) if row.get("position_master_id") not in (None, "", 0) else None
    pnid = (
        str(row.get("position_nomenclature_id"))
        if row.get("position_nomenclature_id") not in (None, "", 0)
        else None
    )
    if scope == "structural" and pmid:
        validation = position_mapping.validate_manual_override("structural", pmid, None, indexes)
        if validation.allowed:
            transform.apply_validated_override_candidate(config, validation)
        else:
            config.position_scope = "structural"
            config.position_master_id = pmid
            config.position_nomenclature_id = None
            config.portaverse_position_title = row.get("portaverse_position_title")
            config.portaverse_group_name = row.get("portaverse_group_name")
            config.portaverse_company_name = row.get("portaverse_company_name")
            config.portaverse_company_code = row.get("portaverse_company_code")
            transform.enforce_position_scope_ids(config)
    elif scope == "non_structural" and pnid:
        if apply_validated_identity(
            config,
            "non_structural",
            None,
            pnid,
            str(row.get("mapping_confidence_reason") or "Approved non-structural identity from Keberlanjutan run."),
            indexes,
            trust_source=str(row.get("mapping_override_trust_source") or "reviewer_manual"),
        ):
            if row.get("position_name"):
                config.position_name = str(row["position_name"])
            return
        config.position_scope = "non_structural"
        config.position_master_id = None
        config.position_nomenclature_id = pnid
        config.portaverse_position_title = row.get("portaverse_position_title")
        config.portaverse_group_name = row.get("portaverse_group_name")
        config.portaverse_company_name = row.get("portaverse_company_name")
        config.portaverse_company_code = row.get("portaverse_company_code")
        config.cluster_label = row.get("cluster_label")
        transform.enforce_position_scope_ids(config)
    config.mapping_confidence_label = row.get("mapping_confidence_label") or position_mapping.HIGH_CONFIDENCE
    config.mapping_confidence_reason = row.get("mapping_confidence_reason")
    config.mapping_review_status = row.get("mapping_review_status") or "approved"
    config.mapping_override_approved = True
    config.mapping_override_trust_source = row.get("mapping_override_trust_source") or "reviewer_manual"
    if row.get("position_name"):
        config.position_name = str(row["position_name"])


def refresh_positions(
    configs: list[transform.PositionConfig],
    indexes: position_mapping.LookupIndexes,
    pengendalian_names: dict[tuple[str, str], str],
    keberlanjutan_approved: dict[tuple[str, str], dict[str, Any]],
) -> list[dict[str, Any]]:
    audit: list[dict[str, Any]] = []
    for config in configs:
        key = config_key(config.source_workbook, config.sheet_name)
        before_pmid = config.position_master_id
        before_pnid = config.position_nomenclature_id
        before_label = config.mapping_confidence_label

        enriched_name = pengendalian_names.get(key)
        if enriched_name:
            config.position_name = enriched_name

        if key in keberlanjutan_approved:
            apply_keberlanjutan_approved(config, keberlanjutan_approved[key], indexes)
        else:
            resolution = position_mapping.resolve_mapping(
                worksheet=config.sheet_name,
                worksheet_title=config.position_name or config.sheet_name,
                group_name=config.group_name,
                source_workbook=config.source_workbook,
                indexes=indexes,
            )
            transform.apply_strict_resolution(config, resolution)

            override = PRODUCTION_VALIDATED_OVERRIDES.get((config.source_workbook or "", config.sheet_name))
            if override:
                apply_validated_identity(
                    config,
                    str(override["scope"]),
                    str(override["position_master_id"]) if override.get("position_master_id") else None,
                    str(override["position_nomenclature_id"]) if override.get("position_nomenclature_id") else None,
                    str(override["reason"]),
                    indexes,
                )
            elif config.group_name == "Group Pengendalian Proyek" and norm(config.sheet_name).startswith("pimpinan proyek"):
                match = match_pengendalian_pmid(config.position_name, config.sheet_name)
                if match:
                    pmid, title = match
                    applied = apply_validated_identity(
                        config,
                        "structural",
                        pmid,
                        None,
                        f"Production-validated project PMID {pmid} ({title}) from worksheet project token match.",
                        indexes,
                    )
                    if not applied:
                        config.mapping_confidence_label = position_mapping.MAPPING_CONFLICT
                        config.mapping_confidence_reason = (
                            f"Project PMID {pmid} ({title}) exists in reference but is not in active structural lookup."
                        )
                        config.mapping_review_status = "NEEDS_CHECK"
                elif config.mapping_confidence_label in {
                    position_mapping.NO_CANDIDATE,
                    position_mapping.MAPPING_CONFLICT,
                    position_mapping.LOW_CONFIDENCE,
                }:
                    config.mapping_review_status = "HOLD"
                    config.position_master_id = None
                    config.position_nomenclature_id = None
                    transform.enforce_position_scope_ids(config)
                    suffix = (
                        " No active Pimpinan Proyek PMID found in production reference for this project name."
                        if config.mapping_confidence_label == position_mapping.NO_CANDIDATE
                        else " Generic or ambiguous Pimpinan Proyek worksheet requires reviewer PMID selection."
                    )
                    config.mapping_confidence_reason = ((config.mapping_confidence_reason or "") + suffix).strip()

        if (
            config.mapping_confidence_label != before_label
            or config.position_master_id != before_pmid
            or config.position_nomenclature_id != before_pnid
        ):
            audit.append(
                {
                    "source_workbook": config.source_workbook,
                    "sheet_name": config.sheet_name,
                    "position_name": config.position_name,
                    "before_label": before_label,
                    "after_label": config.mapping_confidence_label,
                    "before_pmid": before_pmid,
                    "after_pmid": config.position_master_id,
                    "before_pnid": before_pnid,
                    "after_pnid": config.position_nomenclature_id,
                    "reason": config.mapping_confidence_reason,
                    "review_status": config.mapping_review_status,
                }
            )
    return audit


def write_crosscheck_csv(
    configs: list[transform.PositionConfig],
    output_path: Path,
    reference_path: Path,
) -> None:
    rows: list[dict[str, str]] = []
    for config in configs:
        rows.append(
            {
                "Backlog Group": config.group_name or "",
                "Backlog Direktorat": config.directorate_name or "",
                "Backlog Posisi Diminta": str(getattr(config, "backlog_posisi_diminta", "") or ""),
                "Source Workbook": config.source_workbook or "",
                "Worksheet": config.sheet_name,
                "Raw Position": config.position_name,
                "Confidence Label": config.mapping_confidence_label or "",
                "Candidate Scope": config.position_scope or "",
                "Candidate PMID": config.position_master_id or config.candidate_position_master_id or "",
                "Candidate PNID": config.position_nomenclature_id or config.candidate_position_nomenclature_id or "",
                "Candidate Title": config.portaverse_position_title or "",
                "Candidate Group": config.portaverse_group_name or "",
                "Active Employee Name": config.active_employee_name or "",
                "Active Employee NIPP": config.active_employee_nipp or "",
                "Match Reason": config.mapping_confidence_reason or "",
                "Recommended Action": (
                    "No action required; auto-mapped."
                    if config.mapping_confidence_label == position_mapping.HIGH_CONFIDENCE
                    else "Review candidate before allowing upload rows."
                    if config.mapping_confidence_label == position_mapping.LOW_CONFIDENCE
                    else "Check active reference or source worksheet title."
                    if config.mapping_confidence_label == position_mapping.NO_CANDIDATE
                    else "Choose one candidate or create manual override."
                ),
                "Reviewer Confirm Mapping": "",
                "Reviewer Actual PMID": "",
                "Reviewer Actual PNID": "",
                "Reviewer Notes": "",
            }
        )
    if not rows:
        return
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input-config", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--reference", type=Path, default=DEFAULT_REFERENCE)
    parser.add_argument("--output-config", type=Path, default=DEFAULT_OUTPUT_DIR / "support/backlog_target_positions_mapped.config.json")
    parser.add_argument("--audit-output", type=Path, default=DEFAULT_OUTPUT_DIR / "mapping_refresh_audit.csv")
    parser.add_argument("--crosscheck-output", type=Path, default=DEFAULT_OUTPUT_DIR / "crosscheck_posisi_identity_pekerja.csv")
    parser.add_argument("--summary-output", type=Path, default=DEFAULT_OUTPUT_DIR / "mapping_summary.json")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    payload = json.loads(args.reference.read_text(encoding="utf-8"))
    indexes = position_mapping.build_lookup_indexes(payload, None)
    raw_payload = json.loads(args.input_config.read_text(encoding="utf-8"))
    configs = transform.load_config(args.input_config)
    backlog_fields = {
        config_key(row.get("source_workbook"), row.get("sheet_name")): row
        for row in raw_payload.get("positions", [])
    }
    for config in configs:
        extra = backlog_fields.get(config_key(config.source_workbook, config.sheet_name), {})
        for field in ("backlog_nama_file", "backlog_posisi_diminta"):
            if extra.get(field):
                setattr(config, field, extra[field])

    audit = refresh_positions(
        configs,
        indexes,
        load_pengendalian_position_names(),
        load_keberlanjutan_approved(),
    )

    output_positions = [transform.config_to_dict(config) for config in configs]
    for row, config in zip(output_positions, configs):
        for field in ("backlog_nama_file", "backlog_posisi_diminta"):
            if hasattr(config, field):
                row[field] = getattr(config, field)

    args.output_config.parent.mkdir(parents=True, exist_ok=True)
    args.output_config.write_text(
        json.dumps({"positions": output_positions}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    if audit:
        args.audit_output.parent.mkdir(parents=True, exist_ok=True)
        with args.audit_output.open("w", encoding="utf-8", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=list(audit[0].keys()))
            writer.writeheader()
            writer.writerows(audit)

    write_crosscheck_csv(configs, args.crosscheck_output, args.reference)

    distribution = Counter(config.mapping_confidence_label or "unknown" for config in configs)
    hold_count = sum(1 for config in configs if config.mapping_review_status in {"HOLD", "NEEDS_CHECK"})
    summary = {
        "reference": str(args.reference),
        "reference_exported_at": payload.get("exported_at") or payload.get("source", {}).get("exported_at"),
        "reference_kind": "active_production",
        "refreshed_at": datetime.now(timezone.utc).isoformat(),
        "input_config": str(args.input_config),
        "output_config": str(args.output_config),
        "worksheet_rows": len(configs),
        "confidence_distribution": dict(sorted(distribution.items())),
        "held_or_needs_check": hold_count,
        "audit_rows": len(audit),
        "crosscheck_csv": str(args.crosscheck_output),
    }
    args.summary_output.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
