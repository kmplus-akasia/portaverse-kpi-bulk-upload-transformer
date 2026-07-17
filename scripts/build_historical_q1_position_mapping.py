#!/usr/bin/env python3
"""Build JSON and CSV historical Q1 worksheet mapping reports from local JSON inputs."""

from __future__ import annotations

import argparse
import csv
import json
from pathlib import Path

from historical_q1_mapping import REPORT_COLUMNS, build_mapping_rows, mapping_summary


def _load_json(path: Path) -> dict:
    with path.open(encoding="utf-8") as handle:
        return json.load(handle)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--historical-reference", required=True, type=Path)
    parser.add_argument("--config", required=True, type=Path)
    parser.add_argument("--existing-config", required=True, type=Path)
    parser.add_argument("--output-dir", required=True, type=Path)
    parser.add_argument("--company-id", default="1")
    args = parser.parse_args()

    historical_reference = _load_json(args.historical_reference)
    config = _load_json(args.config)
    existing_config = _load_json(args.existing_config)
    positions = config.get("positions", [])
    if not isinstance(positions, list):
        raise ValueError("Config positions must be a list.")

    rows = build_mapping_rows(positions, historical_reference, existing_config, args.company_id)
    if len(rows) != len(positions):
        raise ValueError("Mapping report did not preserve every config position row.")

    args.output_dir.mkdir(parents=True, exist_ok=True)
    json_path = args.output_dir / "mapping_report.json"
    csv_path = args.output_dir / "mapping_report.csv"
    summary_path = args.output_dir / "summary.json"
    json_path.write_text(json.dumps(rows, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    with csv_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=REPORT_COLUMNS)
        writer.writeheader()
        writer.writerows(rows)
    summary_path.write_text(
        json.dumps(mapping_summary(rows, historical_reference), indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )
    print(f"Wrote {len(rows)} mapping rows to {args.output_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
