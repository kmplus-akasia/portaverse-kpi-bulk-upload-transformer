#!/usr/bin/env python3
"""Align position config source_workbook paths with the canonical Kamus inventory."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from kamus_source import (
    attach_config_metadata,
    canonicalize_source_workbook,
    load_inventory_config,
    resolve_kamus_source_root,
)


def normalize_config(config_path: Path, *, write: bool) -> dict[str, object]:
    source_context = resolve_kamus_source_root()
    inventory = load_inventory_config(source_context.inventory_config)
    payload = json.loads(config_path.read_text(encoding="utf-8"))
    changes: list[dict[str, str]] = []

    for item in payload.get("positions", []):
        original = str(item.get("source_workbook") or "").strip()
        if not original:
            continue
        try:
            canonical = canonicalize_source_workbook(original, inventory)
        except (FileNotFoundError, ValueError):
            continue
        if canonical != original:
            changes.append(
                {
                    "position_name": str(item.get("position_name") or ""),
                    "sheet_name": str(item.get("sheet_name") or ""),
                    "before": original,
                    "after": canonical,
                }
            )
            item["source_workbook"] = canonical

    updated = attach_config_metadata(payload, source_context)
    if write:
        config_path.write_text(json.dumps(updated, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return {"config": str(config_path), "changes": changes, "metadata": updated.get("metadata")}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("configs", nargs="+", type=Path, help="Position config JSON files to normalize")
    parser.add_argument("--write", action="store_true", help="Write normalized configs in place")
    args = parser.parse_args()

    total_changes = 0
    for config_path in args.configs:
        result = normalize_config(config_path.resolve(), write=args.write)
        changes = result["changes"]
        total_changes += len(changes)
        print(f"{config_path}: {len(changes)} source_workbook updates")
        for change in changes:
            print(
                f"  - {change['position_name']} ({change['sheet_name']}): "
                f"{change['before']} -> {change['after']}"
            )
    if args.write:
        print(f"Wrote {len(args.configs)} config(s); {total_changes} workbook path updates.")
    else:
        print("Dry run only. Re-run with --write to persist changes.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
