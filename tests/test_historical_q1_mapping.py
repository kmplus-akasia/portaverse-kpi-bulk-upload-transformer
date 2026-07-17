import csv
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from historical_q1_mapping import (  # noqa: E402
    build_mapping_rows,
    historical_assignment_type,
    validate_mapping_row,
)


def worksheet(title, *, group="Group Keuangan", source="source.xlsx", **overrides):
    row = {
        "source_workbook": source,
        "sheet_name": title,
        "position_name": title,
        "position_lookup_names": [title],
        "group_name": group,
        "position_scope": None,
        "position_master_id": None,
        "position_nomenclature_id": None,
    }
    row.update(overrides)
    return row


def historical_row(
    *,
    pmid="501",
    title="Group Head Keuangan",
    type_id="5",
    group="Group Keuangan",
    company="1",
    employee="100",
    employee_name="Ayu",
    missing_org=0,
    lakhar_id=None,
    job_sharing_id=None,
):
    return {
        "position_master_id": pmid,
        "position_title": title,
        "position_master_type_id": type_id,
        "group_name": group,
        "company_id": company,
        "company_name": "PT Pelabuhan Indonesia (Persero)",
        "company_code": "PLD",
        "employee_number": employee,
        "employee_name": employee_name,
        "assignment_end_date": "2026-03-31",
        "missing_historical_organization": missing_org,
        "lakhar_id": lakhar_id,
        "job_sharing_id": job_sharing_id,
    }


def historical_reference(*rows, nomenclature_rows=None):
    return {
        "source": {"cutoff_date": "2026-03-31", "company_id": "1", "read_only": True},
        "historical_assignment_rows": list(rows),
        "nomenclature_rows": list(nomenclature_rows or []),
    }


class HistoricalQ1MappingTest(unittest.TestCase):
    def test_structural_historical_assignment_proposes_pmid_only(self):
        rows = build_mapping_rows(
            [worksheet("Group Head Keuangan")],
            historical_reference(historical_row()),
            {},
            "1",
        )

        self.assertEqual(rows[0]["Candidate PMID"], "501")
        self.assertEqual(rows[0]["Candidate PNID"], "")
        self.assertEqual(rows[0]["Confidence Label"], "high_confidence")
        self.assertEqual(rows[0]["Reviewer Confirm Mapping"], "")

    def test_non_structural_unique_cluster_proposes_pnid_only(self):
        rows = build_mapping_rows(
            [worksheet("Officer Keuangan")],
            historical_reference(
                historical_row(pmid="701", title="Officer Keuangan", type_id="6"),
                nomenclature_rows=[
                    {
                        "position_master_id": "701",
                        "cluster_id": "76",
                        "cluster_label": "Officer Keuangan",
                        "group_name": "Group Keuangan",
                        "company_id": "1",
                    }
                ],
            ),
            {},
            "1",
        )

        self.assertEqual(rows[0]["Candidate PMID"], "")
        self.assertEqual(rows[0]["Candidate PNID"], "76")
        self.assertEqual(rows[0]["Inferred Scope"], "non_structural")

    def test_ambiguous_pnid_stays_needs_check(self):
        rows = build_mapping_rows(
            [worksheet("Officer Keuangan")],
            historical_reference(
                historical_row(pmid="701", title="Officer Keuangan", type_id="6"),
                nomenclature_rows=[
                    {
                        "position_master_id": "701",
                        "cluster_id": "76",
                        "cluster_label": "Officer Keuangan",
                        "group_name": "Group Keuangan",
                        "company_id": "1",
                    },
                    {
                        "position_master_id": "701",
                        "cluster_id": "77",
                        "cluster_label": "Officer Keuangan",
                        "group_name": "Group Keuangan",
                        "company_id": "1",
                    },
                ],
            ),
            {},
            "1",
        )

        self.assertEqual(rows[0]["Confidence Label"], "mapping_conflict")
        self.assertEqual(rows[0]["Candidate PMID"], "")
        self.assertEqual(rows[0]["Candidate PNID"], "")

    def test_records_missing_organization_as_raw_evidence_only(self):
        rows = build_mapping_rows(
            [worksheet("Group Head Keuangan")],
            historical_reference(historical_row(missing_org=1)),
            {},
            "1",
        )

        self.assertEqual(rows[0]["Confidence Label"], "no_candidate")
        self.assertEqual(rows[0]["Candidate PMID"], "")
        self.assertEqual(rows[0]["Historical Employee Numbers"], "100")
        self.assertEqual(rows[0]["Missing Historical Organization Evidence"], "YES")

    def test_excludes_other_company_from_automatic_candidates(self):
        rows = build_mapping_rows(
            [worksheet("Group Head Keuangan")],
            historical_reference(historical_row(company="2")),
            {},
            "1",
        )

        self.assertEqual(rows[0]["Confidence Label"], "no_candidate")
        self.assertEqual(rows[0]["Candidate PMID"], "")

    def test_aggregates_primary_and_secondary_assignment_evidence(self):
        rows = build_mapping_rows(
            [worksheet("Group Head Keuangan")],
            historical_reference(
                historical_row(employee="100", employee_name="Ayu"),
                historical_row(
                    employee="101",
                    employee_name="Bima",
                    lakhar_id="22",
                    job_sharing_id=None,
                ),
                historical_row(
                    employee="102",
                    employee_name="Citra",
                    lakhar_id=None,
                    job_sharing_id="33",
                ),
            ),
            {},
            "1",
        )

        self.assertEqual(rows[0]["Historical Employee Numbers"], "100; 101; 102")
        self.assertEqual(rows[0]["Historical Employee Names"], "Ayu; Bima; Citra")
        self.assertEqual(rows[0]["Assignment Types"], "PRIMARY; LAKHAR; JOB_SHARING")

    def test_primary_assignment_precedence_is_recorded_when_secondary_identity_conflicts(self):
        rows = build_mapping_rows(
            [worksheet("Group Head Keuangan")],
            historical_reference(
                historical_row(pmid="501", employee="100"),
                historical_row(pmid="502", employee="100", lakhar_id="22"),
            ),
            {},
            "1",
        )

        self.assertEqual(rows[0]["Confidence Label"], "mapping_conflict")
        self.assertIn("PRIMARY assignment evidence takes precedence", rows[0]["Confidence Reason"])

    def test_primary_secondary_identity_conflict_blocks_lower_ranked_secondary_identity(self):
        rows = build_mapping_rows(
            [worksheet("Group Head Keuangan")],
            historical_reference(
                historical_row(pmid="501", employee="100", group="Group Keuangan"),
                historical_row(
                    pmid="502",
                    employee="100",
                    lakhar_id="22",
                    group="Group Lain",
                ),
            ),
            {},
            "1",
        )

        self.assertEqual(rows[0]["Confidence Label"], "mapping_conflict")
        self.assertEqual(rows[0]["Candidate PMID"], "")
        self.assertEqual(rows[0]["Candidate PNID"], "")

    def test_two_secondary_identities_for_employee_always_conflict(self):
        rows = build_mapping_rows(
            [worksheet("Group Head Keuangan")],
            historical_reference(
                historical_row(pmid="501", employee="100", lakhar_id="22"),
                historical_row(
                    pmid="502",
                    employee="100",
                    job_sharing_id="33",
                    group="Group Lain",
                ),
            ),
            {},
            "1",
        )

        self.assertEqual(rows[0]["Confidence Label"], "mapping_conflict")
        self.assertEqual(rows[0]["Candidate PMID"], "")
        self.assertEqual(rows[0]["Candidate PNID"], "")

    def test_two_primary_identities_for_employee_always_conflict(self):
        rows = build_mapping_rows(
            [worksheet("Group Head Keuangan")],
            historical_reference(
                historical_row(pmid="501", employee="100"),
                historical_row(pmid="502", employee="100", group="Group Lain"),
            ),
            {},
            "1",
        )

        self.assertEqual(rows[0]["Confidence Label"], "mapping_conflict")
        self.assertEqual(rows[0]["Candidate PMID"], "")
        self.assertEqual(rows[0]["Candidate PNID"], "")

    def test_duplicate_worksheet_key_is_rejected_before_mapping(self):
        with self.assertRaisesRegex(ValueError, "Duplicate source-workbook/worksheet config key"):
            build_mapping_rows(
                [worksheet("Group Head Keuangan"), worksheet("Group Head Keuangan")],
                historical_reference(historical_row()),
                {},
                "1",
            )

    def test_api_rejects_non_head_office_company_id(self):
        with self.assertRaisesRegex(ValueError, "company ID '1'"):
            build_mapping_rows(
                [worksheet("Group Head Keuangan")],
                historical_reference(historical_row(company="2")),
                {},
                "2",
            )

    def test_api_rejects_payload_from_non_head_office_company(self):
        payload = historical_reference(historical_row())
        payload["source"]["company_id"] = "2"

        with self.assertRaisesRegex(ValueError, "source company_id must be '1'"):
            build_mapping_rows([worksheet("Group Head Keuangan")], payload, {}, "1")

    def test_existing_config_ids_are_comparison_fields_only(self):
        position = worksheet("Group Head Keuangan")
        existing = {
            "positions": [
                {
                    **position,
                    "position_master_id": "old-pmid",
                    "position_nomenclature_id": "old-pnid",
                }
            ]
        }
        rows = build_mapping_rows([position], historical_reference(historical_row()), existing, "1")

        self.assertEqual(rows[0]["Candidate PMID"], "501")
        self.assertEqual(rows[0]["Existing Config PMID"], "old-pmid")
        self.assertEqual(rows[0]["Existing Config PNID"], "old-pnid")

    def test_assignment_type_prefers_lakhar_then_job_sharing(self):
        self.assertEqual(historical_assignment_type({}), "PRIMARY")
        self.assertEqual(historical_assignment_type({"lakhar_id": "1"}), "LAKHAR")
        self.assertEqual(historical_assignment_type({"job_sharing_id": "1"}), "JOB_SHARING")
        self.assertEqual(
            historical_assignment_type({"lakhar_id": "1", "job_sharing_id": "2"}),
            "LAKHAR",
        )

    def test_validation_rejects_candidate_namespace_mismatches(self):
        both = validate_mapping_row(
            {"Inferred Scope": "structural", "Candidate PMID": "501", "Candidate PNID": "76"}
        )
        structural_pnid = validate_mapping_row(
            {"Inferred Scope": "structural", "Candidate PMID": "", "Candidate PNID": "76"}
        )
        non_structural_pmid = validate_mapping_row(
            {"Inferred Scope": "non_structural", "Candidate PMID": "501", "Candidate PNID": ""}
        )

        self.assertTrue(any("both" in error.lower() for error in both))
        self.assertTrue(any("structural" in error.lower() for error in structural_pnid))
        self.assertTrue(any("non-structural" in error.lower() for error in non_structural_pmid))

    def test_cli_writes_json_csv_and_summary(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            reference_path = temp / "reference.json"
            config_path = temp / "config.json"
            existing_path = temp / "existing.json"
            output_dir = temp / "output"
            reference_path.write_text(json.dumps(historical_reference(historical_row())), encoding="utf-8")
            config_path.write_text(json.dumps({"positions": [worksheet("Group Head Keuangan")]}), encoding="utf-8")
            existing_path.write_text(json.dumps({"positions": []}), encoding="utf-8")

            completed = subprocess.run(
                [
                    sys.executable,
                    str(ROOT / "scripts" / "build_historical_q1_position_mapping.py"),
                    "--historical-reference",
                    str(reference_path),
                    "--config",
                    str(config_path),
                    "--existing-config",
                    str(existing_path),
                    "--output-dir",
                    str(output_dir),
                ],
                check=True,
                capture_output=True,
                text=True,
            )

            rows = json.loads((output_dir / "mapping_report.json").read_text(encoding="utf-8"))
            summary = json.loads((output_dir / "summary.json").read_text(encoding="utf-8"))
            with (output_dir / "mapping_report.csv").open(newline="", encoding="utf-8") as handle:
                csv_rows = list(csv.DictReader(handle))

        self.assertIn("Wrote 1 mapping rows", completed.stdout)
        self.assertEqual(rows[0]["Candidate PMID"], "501")
        self.assertEqual(len(csv_rows), 1)
        self.assertEqual(summary["mapping_rows"], 1)
        self.assertEqual(summary["reviewer_approved_rows"], 0)

    def test_cli_rejects_non_head_office_company_id(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            reference_path = temp / "reference.json"
            config_path = temp / "config.json"
            existing_path = temp / "existing.json"
            output_dir = temp / "output"
            reference_path.write_text(json.dumps(historical_reference(historical_row())), encoding="utf-8")
            config_path.write_text(json.dumps({"positions": [worksheet("Group Head Keuangan")]}), encoding="utf-8")
            existing_path.write_text(json.dumps({"positions": []}), encoding="utf-8")

            completed = subprocess.run(
                [
                    sys.executable,
                    str(ROOT / "scripts" / "build_historical_q1_position_mapping.py"),
                    "--historical-reference",
                    str(reference_path),
                    "--config",
                    str(config_path),
                    "--existing-config",
                    str(existing_path),
                    "--output-dir",
                    str(output_dir),
                    "--company-id",
                    "2",
                ],
                check=False,
                capture_output=True,
                text=True,
            )

        self.assertNotEqual(completed.returncode, 0)
        self.assertIn("company ID '1'", completed.stderr)

    def test_cli_rejects_duplicate_worksheet_key(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            reference_path = temp / "reference.json"
            config_path = temp / "config.json"
            existing_path = temp / "existing.json"
            output_dir = temp / "output"
            reference_path.write_text(json.dumps(historical_reference(historical_row())), encoding="utf-8")
            config_path.write_text(
                json.dumps(
                    {
                        "positions": [
                            worksheet("Group Head Keuangan"),
                            worksheet("Group Head Keuangan"),
                        ]
                    }
                ),
                encoding="utf-8",
            )
            existing_path.write_text(json.dumps({"positions": []}), encoding="utf-8")

            completed = subprocess.run(
                [
                    sys.executable,
                    str(ROOT / "scripts" / "build_historical_q1_position_mapping.py"),
                    "--historical-reference",
                    str(reference_path),
                    "--config",
                    str(config_path),
                    "--existing-config",
                    str(existing_path),
                    "--output-dir",
                    str(output_dir),
                ],
                check=False,
                capture_output=True,
                text=True,
            )

        self.assertNotEqual(completed.returncode, 0)
        self.assertIn("Duplicate source-workbook/worksheet config key", completed.stderr)


if __name__ == "__main__":
    unittest.main()
