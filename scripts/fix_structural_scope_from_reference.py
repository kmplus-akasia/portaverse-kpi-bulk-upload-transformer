#!/usr/bin/env python3
"""Create a derived KPI position config with structural scope corrected."""

from __future__ import annotations

import argparse
import csv
import json
import re
from datetime import datetime, timezone
from pathlib import Path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-config", required=True, type=Path)
    parser.add_argument("--reference", required=True, type=Path)
    parser.add_argument("--output-config", required=True, type=Path)
    parser.add_argument("--audit-output", required=True, type=Path)
    return parser.parse_args()


def norm(value: object) -> str:
    text = str(value or "").lower()
    text = re.sub(r"\bdh\b", "department head", text)
    text = text.replace("&", " dan ")
    text = re.sub(r"[-_/(),.]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def is_truthy_flag(value: object) -> bool:
    return value in (True, 1, "1", "true", "TRUE", "Y", "y")


def is_active_reference_row(row: dict[str, object]) -> bool:
    for field in ("is_company_active", "is_group_active", "is_position_active"):
        if row.get(field) not in (None, "") and not is_truthy_flag(row.get(field)):
            return False
    if row.get("is_position_organization_active") not in (None, ""):
        return is_truthy_flag(row.get("is_position_organization_active"))
    return True


def main() -> int:
    args = parse_args()

    with args.reference.open() as handle:
        reference = json.load(handle)
    masters_by_id: dict[str, list[dict[str, object]]] = {}
    for row in reference["position_master_rows"]:
        masters_by_id.setdefault(str(row["position_master_id"]), []).append(row)
    nomenclature_ids = {
        str(row["cluster_id"])
        for row in reference["rows"]
        if row.get("cluster_id") not in (None, "", 0, "0")
    }
    cluster_labels_by_id: dict[str, set[str]] = {}
    pnids_by_pmid: dict[str, set[str]] = {}
    for row in reference["rows"]:
        if not is_active_reference_row(row):
            continue
        cluster_id = row.get("cluster_id")
        if cluster_id in (None, "", 0, "0"):
            continue
        cluster_labels_by_id.setdefault(str(cluster_id), set()).add(norm(row.get("cluster_label")))
        position_master_id = row.get("position_master_id")
        if position_master_id not in (None, "", 0, "0"):
            pnids_by_pmid.setdefault(str(position_master_id), set()).add(str(cluster_id))

    with args.input_config.open() as handle:
        config = json.load(handle)

    fixes: list[dict[str, str]] = []
    for pos in config["positions"]:
        scope = (pos.get("position_scope") or "").strip()
        pnid = str(pos.get("position_nomenclature_id") or "").strip()
        pmid = str(pos.get("position_master_id") or "").strip()
        if not scope and not pmid and not pnid:
            before = {
                "position_master_id": pos.get("position_master_id"),
                "position_nomenclature_id": pos.get("position_nomenclature_id"),
                "position_scope": pos.get("position_scope"),
                "cluster_label": pos.get("cluster_label"),
                "position_name": pos.get("position_name"),
                "portaverse_position_title": pos.get("portaverse_position_title"),
                "portaverse_group_name": pos.get("portaverse_group_name"),
                "portaverse_company_name": pos.get("portaverse_company_name"),
            }
            pos["position_scope"] = "neglect"
            after = dict(before)
            after["position_scope"] = "neglect"
            fixes.append(
                {
                    "source_workbook": pos.get("source_workbook", ""),
                    "sheet_name": pos.get("sheet_name", ""),
                    "pmid": "",
                    "production_type_id": "",
                    "resolved_scope": "neglect",
                    "before": json.dumps(before, ensure_ascii=False, sort_keys=True),
                    "after": json.dumps(after, ensure_ascii=False, sort_keys=True),
                }
            )
            continue
        if scope != "non_structural" or not pnid:
            continue
        # The field is explicitly PNID. A valid value in the PNID namespace wins
        # even when the same number also exists as an internal PMID.
        if pnid in nomenclature_ids:
            continue
        if pnid not in masters_by_id:
            continue

        before = {
            "position_master_id": pos.get("position_master_id"),
            "position_nomenclature_id": pos.get("position_nomenclature_id"),
            "position_scope": pos.get("position_scope"),
            "cluster_label": pos.get("cluster_label"),
            "position_name": pos.get("position_name"),
            "portaverse_position_title": pos.get("portaverse_position_title"),
            "portaverse_group_name": pos.get("portaverse_group_name"),
            "portaverse_company_name": pos.get("portaverse_company_name"),
        }
        masters = masters_by_id[pnid]
        master_types = {str(row.get("position_master_type_id") or "") for row in masters}
        if len(master_types) != 1:
            raise ValueError(
                f"mixed production position types for PMID {pnid}: {sorted(master_types)}"
            )
        master_type = next(iter(master_types))
        master = masters[0]

        if master_type == "5":
            resolved_scope = "structural"
            resolved_pmid = pnid
            resolved_pnid = None
            resolved_cluster_label = None
        else:
            candidate_pnids = pnids_by_pmid.get(pnid, set())
            if not candidate_pnids:
                raise ValueError(f"no PNID found for non-structural PMID {pnid}")
            if len(candidate_pnids) != 1:
                raise ValueError(
                    f"multiple PNIDs found for non-structural PMID {pnid}: "
                    f"{sorted(candidate_pnids)}"
                )
            resolved_scope = "non_structural"
            resolved_pmid = None
            resolved_pnid = next(iter(candidate_pnids))
            labels = cluster_labels_by_id.get(resolved_pnid, set())
            resolved_cluster_label = next(iter(labels)) if len(labels) == 1 else None

        pos["position_scope"] = resolved_scope
        pos["position_master_id"] = resolved_pmid
        pos["position_nomenclature_id"] = resolved_pnid
        pos["cluster_label"] = resolved_cluster_label
        pos["position_name"] = master.get("position_name") or pos.get("position_name")
        pos["portaverse_position_title"] = master.get("position_name")
        pos["portaverse_group_name"] = master.get("group_name")
        pos["portaverse_company_name"] = master.get("company_name")
        lookup_names = pos.get("position_lookup_names") or []
        if master.get("position_name") and master["position_name"] not in lookup_names:
            lookup_names.insert(0, master["position_name"])
        pos["position_lookup_names"] = lookup_names

        after = {
            "position_master_id": pos.get("position_master_id"),
            "position_nomenclature_id": pos.get("position_nomenclature_id"),
            "position_scope": pos.get("position_scope"),
            "cluster_label": pos.get("cluster_label"),
            "position_name": pos.get("position_name"),
            "portaverse_position_title": pos.get("portaverse_position_title"),
            "portaverse_group_name": pos.get("portaverse_group_name"),
            "portaverse_company_name": pos.get("portaverse_company_name"),
        }
        fixes.append(
            {
                "source_workbook": pos.get("source_workbook", ""),
                "sheet_name": pos.get("sheet_name", ""),
                "pmid": pnid,
                "production_type_id": master_type,
                "resolved_scope": resolved_scope,
                "before": json.dumps(before, ensure_ascii=False, sort_keys=True),
                "after": json.dumps(after, ensure_ascii=False, sort_keys=True),
            }
        )

    config["reference_source"]["structural_scope_autofix_generated_at"] = datetime.now(
        timezone.utc
    ).isoformat()
    config["reference_source"]["structural_scope_autofix_count"] = len(fixes)
    config["reference_source"]["structural_scope_autofix_rule"] = (
        "identity collisions are resolved by production position_master_type_id; "
        "type 5 uses PMID and all other types require one unique PNID"
    )

    args.output_config.parent.mkdir(parents=True, exist_ok=True)
    with args.output_config.open("w") as handle:
        json.dump(config, handle, ensure_ascii=False, indent=2)
        handle.write("\n")

    args.audit_output.parent.mkdir(parents=True, exist_ok=True)
    with args.audit_output.open("w", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=[
                "source_workbook",
                "sheet_name",
                "pmid",
                "production_type_id",
                "resolved_scope",
                "before",
                "after",
            ],
        )
        writer.writeheader()
        writer.writerows(fixes)

    print(f"fixes={len(fixes)}")
    print(f"output_config={args.output_config}")
    print(f"audit_output={args.audit_output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
