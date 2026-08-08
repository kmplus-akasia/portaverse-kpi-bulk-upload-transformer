#!/usr/bin/env python3
"""Build approved upload config + source ZIP for Subholding identities marked Siap konversi.

Reads the slim Position First identity conversion report, resolves each siap identity
to an inventory workbook+sheet, writes:
  - approved positions config JSON
  - ZIP of only the needed source workbooks (keys = inventory source_workbook paths)
  - README_SOURCE.md / allowlist receipt
"""

from __future__ import annotations

import argparse
import json
import re
import zipfile
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

import position_mapping as pm

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def norm(value: Any) -> str:
    return "" if value is None else str(value).strip()


def norm_sheet(value: Any) -> str:
    """Preserve intentional trailing spaces in Excel sheet names (max 31 chars)."""
    if value is None:
        return ""
    return str(value).strip("\n\r\t")


def is_na(value: Any) -> bool:
    return norm(value).upper() in {"", "#N/A", "N/A", "NA", "-", "NULL"}


def load_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    out: list[str] = []
    for item in root.findall("m:si", NS):
        texts = [
            node.text or ""
            for node in item.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")
        ]
        out.append("".join(texts))
    return out


def cell_value(cell: ET.Element, shared: list[str]) -> Any:
    cell_type = cell.attrib.get("t")
    inline = cell.find("m:is", NS)
    if inline is not None:
        return "".join(
            node.text or ""
            for node in inline.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")
        )
    value_node = cell.find("m:v", NS)
    if value_node is None:
        return None
    raw = value_node.text
    return shared[int(raw)] if cell_type == "s" else raw


def col_index(cell_ref: str) -> int:
    match = re.match(r"([A-Z]+)", cell_ref or "A1")
    assert match is not None
    total = 0
    for char in match.group(1):
        total = total * 26 + (ord(char) - 64)
    return total - 1


def sheet_table(path: Path, sheet_name: str) -> list[dict[str, str]]:
    with zipfile.ZipFile(path) as archive:
        shared = load_shared_strings(archive)
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        rid = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
        target = None
        for sheet in workbook.findall("m:sheets/m:sheet", NS):
            if sheet.attrib["name"] == sheet_name:
                target = rid[sheet.attrib[f"{REL}id"]].lstrip("/")
                if not target.startswith("xl/"):
                    target = f"xl/{target}"
                break
        if target is None:
            raise KeyError(sheet_name)
        root = ET.fromstring(archive.read(target))
        rows: list[list[Any]] = []
        for row in root.findall("m:sheetData/m:row", NS):
            cells: dict[int, Any] = {}
            for cell in row.findall("m:c", NS):
                cells[col_index(cell.attrib.get("r", "A1"))] = cell_value(cell, shared)
            if not cells:
                continue
            width = max(cells) + 1
            rows.append([cells.get(index) for index in range(width)])
    header_idx = None
    header: list[str] = []
    for index, row in enumerate(rows[:10]):
        values = [norm(cell) for cell in row]
        if "Status Kesiapan" in values:
            header_idx = index
            header = values
            break
    if header_idx is None:
        raise RuntimeError(f"Header not found in {sheet_name}")
    out: list[dict[str, str]] = []
    for row in rows[header_idx + 1 :]:
        if not row or all(cell in (None, "") for cell in row):
            continue
        out.append(
            {
                name: (norm(row[idx]) if idx < len(row) else "")
                for idx, name in enumerate(header)
                if name
            }
        )
    return out


def token_score(left: str, right: str) -> float:
    a = set(pm.normalize_position_lookup(left).split())
    b = set(pm.normalize_position_lookup(right).split())
    if not a or not b:
        return 0.0
    return len(a & b) / max(len(a), len(b))


def pick_workbook_hint(row: dict[str, str]) -> str:
    alasan = row.get("Alasan Status", "")
    rw = norm(row.get("Nama File R&W (raw)"))
    auto = norm(row.get("File Kamus (usulan otomatis)"))
    folder = norm(row.get("Folder R&W (raw)"))
    if "Automated" in alasan or is_na(rw):
        if auto:
            return Path(auto).name
    if rw and not is_na(rw):
        return Path(rw).name
    if auto:
        return Path(auto).name
    if folder.lower().endswith(".xlsx"):
        return Path(folder).name
    return ""


def resolve_inventory_entry(
    row: dict[str, str],
    by_base: dict[str, list[dict[str, Any]]],
) -> dict[str, Any]:
    sheet_hint = (
        norm(row.get("Sheet Inventory (resolved)"))
        or norm(row.get("Sheet Kamus (usulan otomatis)"))
        or norm(row.get("Worksheet R&W (raw)"))
    )
    workbook_hint = pick_workbook_hint(row)
    hint_path = Path(workbook_hint)
    base = hint_path.name.casefold()
    stem = hint_path.stem.casefold()
    candidates = by_base.get(base) or by_base.get(stem) or []
    if not candidates:
        raise KeyError(f"workbook not in inventory: {workbook_hint}")
    # Deduplicate entries appended under both basename and stem keys.
    deduped_candidates = {
        f"{entry['source_workbook']}::{entry['sheet_name']}": entry for entry in candidates
    }
    candidates = list(deduped_candidates.values())
    exact: list[dict[str, Any]] = []
    soft: list[tuple[float, dict[str, Any]]] = []
    for entry in candidates:
        sheet_name = norm(entry.get("sheet_name"))
        position_name = norm(entry.get("position_name"))
        if sheet_hint.casefold() in {sheet_name.casefold(), position_name.casefold()}:
            exact.append(entry)
            continue
        if len(sheet_name) == 31 and sheet_hint.casefold().startswith(sheet_name.casefold().rstrip()):
            exact.append(entry)
            continue
        if sheet_hint[:31].casefold() == sheet_name.casefold():
            exact.append(entry)
            continue
        score = max(
            token_score(sheet_hint, sheet_name),
            token_score(sheet_hint, position_name),
        )
        if score >= 0.6:
            soft.append((score, entry))
    if exact:
        unique = {f"{e['source_workbook']}|{e['sheet_name']}": e for e in exact}
        if len(unique) == 1:
            return next(iter(unique.values()))
        # prefer sheet_name exact over position_name
        for entry in exact:
            if norm(entry.get("sheet_name")).casefold() == sheet_hint.casefold():
                return entry
        return exact[0]
    if not soft:
        raise KeyError(f"sheet not in workbook {workbook_hint}: {sheet_hint}")
    soft.sort(key=lambda item: (-item[0], item[1]["sheet_name"]))
    return soft[0][1]


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--identity-report",
        type=Path,
        default=Path(
            "outputs/kamus-group2-subholding-rw-reconcile-20260807/"
            "Position_First_Identity_Conversion_Subholding_LATEST.xlsx"
        ),
    )
    parser.add_argument(
        "--inventory",
        type=Path,
        default=Path("configs/kamus_kpi_group2_visible_20260807.json"),
    )
    parser.add_argument(
        "--reference",
        type=Path,
        default=Path("configs/production_position_reference.json"),
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=None,
    )
    args = parser.parse_args()

    stamp = datetime.now().astimezone().strftime("%Y%m%d_%H%M%S")
    generated_at = datetime.now().astimezone().isoformat(timespec="seconds")
    output_dir = args.output_dir or Path(
        f"outputs/kamus-group2-subholding-upload-ready-siap-{stamp}"
    )
    output_dir.mkdir(parents=True, exist_ok=True)

    inventory = json.loads(args.inventory.read_text(encoding="utf-8"))
    source_root = Path(inventory["metadata"]["source_root"])
    if not source_root.is_absolute():
        source_root = Path.cwd() / source_root
    by_base: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in inventory.get("kamus_kpi_v2", []):
        if not row.get("include_in_position_config"):
            continue
        workbook_path = Path(row["source_workbook"])
        by_base[workbook_path.name.casefold()].append(row)
        # R&W often writes .xlsx while source is .xlsm (or vice versa).
        by_base[workbook_path.stem.casefold()].append(row)

    reference = json.loads(args.reference.read_text(encoding="utf-8"))
    ref_meta = {}
    meta_path = Path("configs/production_position_reference.meta.json")
    if meta_path.exists():
        ref_meta = json.loads(meta_path.read_text(encoding="utf-8"))

    indexes = pm.build_lookup_indexes(reference)
    by_pmid = {
        str(c.position_master_id): c
        for c in indexes.structural
        if c.position_master_id
    }
    by_pnid = {
        str(c.position_nomenclature_id): c
        for c in indexes.non_structural
        if c.position_nomenclature_id
    }

    # Prefer Pemetaan Kamus KPI (tracker workbook); fall back to legacy Siap Konversi sheet.
    try:
        mapping_rows = sheet_table(args.identity_report, "Pemetaan Kamus KPI")
        siap_rows = [
            row for row in mapping_rows if norm(row.get("Status Kesiapan")) == "Siap konversi"
        ]
    except KeyError:
        siap_rows = sheet_table(args.identity_report, "Siap Konversi")
    positions: list[dict[str, Any]] = []
    allowlist_rows: list[dict[str, Any]] = []
    held_unresolved: list[dict[str, Any]] = []
    alasan_counts = Counter()

    for row in siap_rows:
        pmid = norm(row.get("PMID"))
        pnid = norm(row.get("PNID"))
        alasan = norm(row.get("Alasan Status"))
        alasan_counts[alasan] += 1
        try:
            entry = resolve_inventory_entry(row, by_base)
        except KeyError as exc:
            held_unresolved.append(
                {
                    "pmid": pmid,
                    "pnid": pnid,
                    "judul_posisi": norm(row.get("Judul Posisi")),
                    "alasan_status": alasan,
                    "batch_status": "held_inventory_unresolved",
                    "reason": str(exc),
                    "workbook_hint": pick_workbook_hint(row),
                    "sheet_hint": (
                        norm(row.get("Sheet Inventory (resolved)"))
                        or norm(row.get("Sheet Kamus (usulan otomatis)"))
                        or norm(row.get("Worksheet R&W (raw)"))
                    ),
                    "nipp_count": row.get("Jumlah NIPP"),
                }
            )
            continue

        source_workbook = norm(entry["source_workbook"])
        sheet_name = norm_sheet(entry["sheet_name"])
        position_name = norm(entry.get("position_name")) or sheet_name.rstrip()

        candidate = by_pmid.get(pmid) if pmid else None
        if candidate is None and pnid:
            candidate = by_pnid.get(pnid)

        group_name = (
            norm(row.get("Unit / Group"))
            or (candidate.group_name if candidate else "")
            or Path(source_workbook).stem
        )
        directorate_name = (
            (candidate.company_name if candidate else "")
            or norm(row.get("Perusahaan"))
            or "Subholding"
        )

        if pmid and not pnid:
            scope = "structural"
            position_master_id = pmid
            position_nomenclature_id = None
        elif pnid and not pmid:
            scope = "non_structural"
            position_master_id = None
            position_nomenclature_id = pnid
        elif pmid and pnid:
            # Prefer structural when both present
            scope = "structural"
            position_master_id = pmid
            position_nomenclature_id = None
        else:
            held_unresolved.append(
                {
                    "pmid": pmid,
                    "pnid": pnid,
                    "judul_posisi": norm(row.get("Judul Posisi")),
                    "alasan_status": alasan,
                    "batch_status": "held_missing_identity",
                    "reason": "missing PMID/PNID",
                    "nipp_count": row.get("Jumlah NIPP"),
                }
            )
            continue

        positions.append(
            {
                "sheet_name": sheet_name,
                "position_name": position_name,
                "group_name": group_name,
                "directorate_name": directorate_name,
                "source_workbook": source_workbook,
                "position_master_id": position_master_id,
                "position_nomenclature_id": position_nomenclature_id,
                "position_scope": scope,
                "portaverse_position_title": norm(row.get("Judul Posisi"))
                or (candidate.title if candidate else position_name),
                "portaverse_group_name": group_name,
                "portaverse_company_name": directorate_name,
                "portaverse_company_code": "",
                "company_in_id": norm(row.get("company_in_id"))
                or (str(candidate.company_id) if candidate and candidate.company_id else ""),
                "mapping_confidence_label": "approved_siap_konversi",
                "mapping_confidence_reason": alasan,
                "mapping_review_status": "approved",
                "mapping_override_approved": True,
                "mapping_override_trust_source": alasan,
                "cluster_label": "Subholding",
                "expected_impact_count": 10,
                "drop_comment_values": ["Drop"],
                "active_employee_count": int(float(row.get("Jumlah Pegawai") or row.get("Jumlah NIPP") or 0))
                if norm(row.get("Jumlah Pegawai") or row.get("Jumlah NIPP"))
                else 0,
                "_alasan": alasan,
                "_judul": norm(row.get("Judul Posisi")),
                "_pmid": pmid,
                "_pnid": pnid,
                "_nipp_count": row.get("Jumlah NIPP"),
            }
        )

    def alasan_rank(alasan: str) -> int:
        if alasan.startswith("R&W resolve"):
            return 0
        if "folder mismatch" in alasan:
            return 1
        if alasan.startswith("Automated high_confidence (path R&W kosong)"):
            return 2
        if alasan.startswith("Automated high_confidence (R&W"):
            return 3
        if "alias/urutan kata" in alasan:
            return 4
        return 9

    def title_score(judul: str, sheet_name: str, position_name: str) -> float:
        return max(token_score(judul, sheet_name), token_score(judul, position_name))

    # Satu sheet kamus boleh dipakai berulang untuk banyak identity (PMID/PNID berbeda).
    # Hanya collapse jika identity yang sama muncul lebih dari sekali.
    selected: list[dict[str, Any]] = []
    allowlist_rows: list[dict[str, Any]] = []
    workbooks_needed: set[str] = set()
    reused_sheet_identities = 0
    by_sheet_counts: Counter[str] = Counter()
    seen_identity: set[str] = set()
    for item in sorted(
        positions,
        key=lambda entry: (
            alasan_rank(entry["_alasan"]),
            -title_score(entry["_judul"], entry["sheet_name"], entry["position_name"]),
            entry["_pmid"],
            entry["_pnid"],
        ),
    ):
        identity_key = f"pmid:{item['_pmid']}" if item["_pmid"] else f"pnid:{item['_pnid']}"
        if identity_key in seen_identity:
            continue
        seen_identity.add(identity_key)
        sheet_key = f"{item['source_workbook']}::{item['sheet_name']}"
        by_sheet_counts[sheet_key] += 1
        if by_sheet_counts[sheet_key] > 1:
            reused_sheet_identities += 1
        workbooks_needed.add(item["source_workbook"])
        clean = {k: v for k, v in item.items() if not k.startswith("_")}
        selected.append(clean)
        allowlist_rows.append(
            {
                "pmid": item["_pmid"],
                "pnid": item["_pnid"],
                "judul_posisi": item["_judul"],
                "alasan_status": item["_alasan"],
                "source_workbook": item["source_workbook"],
                "sheet_name": item["sheet_name"],
                "nipp_count": item["_nipp_count"],
                "batch_status": "included",
                "sheet_reuse_index": by_sheet_counts[sheet_key],
            }
        )

    sheets_reused = sum(1 for count in by_sheet_counts.values() if count > 1)

    config_path = output_dir / f"subholding_siap_upload_config_{stamp}.json"
    config_payload = {
        "metadata": {
            "title": "Subholding siap-konversi upload config",
            "generated_at": generated_at,
            "source_root": str(source_root),
            "inventory_config": str(args.inventory),
            "identity_report": str(args.identity_report),
            "production_reference": str(args.reference),
            "production_reference_exported_at": ref_meta.get("exported_at")
            or reference.get("exported_at"),
            "siap_identity_count": len(siap_rows),
            "position_config_count": len(selected),
            "sheet_reuse_allowed": True,
            "sheets_reused_count": sheets_reused,
            "extra_identities_via_sheet_reuse": reused_sheet_identities,
            "held_inventory_unresolved": len(held_unresolved),
            "workbook_count": len(workbooks_needed),
            "alasan_status_counts": dict(alasan_counts),
            "left_out": (
                "Belum konversi identities from the same report; "
                "held_inventory_unresolved"
            ),
        },
        "positions": selected,
    }
    config_path.write_text(json.dumps(config_payload, ensure_ascii=False, indent=2), encoding="utf-8")

    zip_path = output_dir / f"subholding_siap_source_workbooks_{stamp}.zip"
    missing_files: list[str] = []
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for relative in sorted(workbooks_needed):
            absolute = source_root / relative
            if not absolute.exists():
                missing_files.append(relative)
                continue
            archive.write(absolute, arcname=relative)
    if missing_files:
        raise SystemExit(f"Missing source workbooks ({len(missing_files)}): {missing_files[:5]}")

    allowlist_path = output_dir / f"siap_allowlist_{stamp}.json"
    allowlist_path.write_text(
        json.dumps(
            {
                "generated_at": generated_at,
                "included_count": len(allowlist_rows),
                "sheets_reused_count": sheets_reused,
                "extra_identities_via_sheet_reuse": reused_sheet_identities,
                "held_inventory_unresolved_count": len(held_unresolved),
                "rows": allowlist_rows,
                "held_inventory_unresolved": held_unresolved,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )

    readme = output_dir / "README_SOURCE.md"
    readme.write_text(
        f"""# Subholding Siap-Konversi Upload Batch

Generated: `{generated_at}`

## Pins
- Identity report: `{args.identity_report}`
- Inventory: `{args.inventory}` (`generated_at` {inventory['metadata'].get('generated_at')})
- Source root: `{source_root}`
- Production reference: `{args.reference}` (exported_at: {ref_meta.get('exported_at') or reference.get('exported_at')})
- Template: `input/KPI Upload Template.xlsx`

## Scope
- Siap identities in report: **{len(siap_rows)}**
- Position configs included: **{len(selected)}** (sheet reuse allowed)
- Sheets reused by >1 identity: **{sheets_reused}** (+{reused_sheet_identities} extra identities)
- Held inventory unresolved (path tidak resolve ke inventory terbaru): **{len(held_unresolved)}**
- Source workbooks in ZIP: **{len(workbooks_needed)}**
- Left out of conversion: Belum konversi + held_inventory_unresolved

## Alasan Status
{json.dumps(dict(alasan_counts), ensure_ascii=False, indent=2)}

## Artifacts
- Config: `{config_path.name}`
- Source ZIP: `{zip_path.name}`
- Allowlist: `{allowlist_path.name}`
""",
        encoding="utf-8",
    )

    print(
        json.dumps(
            {
                "output_dir": str(output_dir),
                "config": str(config_path),
                "source_zip": str(zip_path),
                "allowlist": str(allowlist_path),
                "siap_identities": len(siap_rows),
                "config_positions": len(selected),
                "sheets_reused_count": sheets_reused,
                "extra_identities_via_sheet_reuse": reused_sheet_identities,
                "held_inventory_unresolved": len(held_unresolved),
                "workbooks": len(workbooks_needed),
                "alasan_status_counts": dict(alasan_counts),
            },
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
