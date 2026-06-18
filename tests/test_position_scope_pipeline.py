import csv
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from scripts import fix_structural_scope_from_reference as scope_fix
from scripts import validate_kpi_upload_batch as validator


class PositionScopePipelineTest(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)

    def tearDown(self):
        self.temp_dir.cleanup()

    def write_json(self, name, payload):
        path = self.root / name
        path.write_text(json.dumps(payload), encoding="utf-8")
        return path

    def reference(self, master_type=6, cluster_ids=(10,)):
        master = {
            "position_master_id": 44,
            "position_name": "Officer II Monitoring dan Evaluasi Pengawasan Intern",
            "position_master_type_id": master_type,
            "group_name": "Department Monitoring dan Evaluasi Pengawasan Intern",
            "company_name": "PT Pelabuhan Indonesia (Persero)",
        }
        return {
            "position_master_rows": [master],
            "rows": [
                {
                    **master,
                    "cluster_id": cluster_id,
                    "cluster_label": "Officer Monitoring Dan Evaluasi Pengawasan Intern",
                }
                for cluster_id in cluster_ids
            ],
        }

    def collision_reference(self):
        reference = self.reference(master_type=6, cluster_ids=(10,))
        reference["position_master_rows"].append(
            {
                "position_master_id": 100,
                "position_name": "Officer I Transaksi dan Proses Data",
                "position_master_type_id": 6,
            }
        )
        reference["rows"].append(
            {
                "position_master_id": 100,
                "position_name": "Officer I Transaksi dan Proses Data",
                "position_master_type_id": 6,
                "cluster_id": 44,
                "cluster_label": "Officer Transaksi Dan Proses Data - Company 1",
            }
        )
        return reference

    def config(self):
        return {
            "reference_source": {},
            "positions": [
                {
                    "source_workbook": "source.xlsx",
                    "sheet_name": "Officer Transaksi dan Proses",
                    "position_name": "Officer Transaksi dan Proses",
                    "position_master_id": None,
                    "position_nomenclature_id": "44",
                    "position_scope": "non_structural",
                    "cluster_label": "Officer Transaksi dan Proses",
                    "position_lookup_names": ["Officer Transaksi dan Proses"],
                }
            ],
        }

    def run_scope_fix(self, reference, config=None):
        input_path = self.write_json("input.json", config or self.config())
        reference_path = self.write_json("reference.json", reference)
        output_path = self.root / "output.json"
        audit_path = self.root / "audit.csv"
        argv = [
            "fix_structural_scope_from_reference.py",
            "--input-config",
            str(input_path),
            "--reference",
            str(reference_path),
            "--output-config",
            str(output_path),
            "--audit-output",
            str(audit_path),
        ]
        with patch.object(sys, "argv", argv):
            scope_fix.main()
        with audit_path.open(newline="") as handle:
            audit = list(csv.DictReader(handle))
        return json.loads(output_path.read_text(encoding="utf-8")), audit

    def test_type_6_position_is_corrected_to_unique_pnid(self):
        output, audit = self.run_scope_fix(self.reference(master_type=6, cluster_ids=(10,)))

        position = output["positions"][0]
        self.assertEqual(position["position_scope"], "non_structural")
        self.assertIsNone(position["position_master_id"])
        self.assertEqual(position["position_nomenclature_id"], "10")
        self.assertEqual(len(audit), 1)

    def test_valid_pnid_wins_when_same_number_is_also_a_pmid(self):
        output, audit = self.run_scope_fix(self.collision_reference())

        position = output["positions"][0]
        self.assertEqual(position["position_scope"], "non_structural")
        self.assertIsNone(position["position_master_id"])
        self.assertEqual(position["position_nomenclature_id"], "44")
        self.assertEqual(position["position_name"], "Officer Transaksi dan Proses")
        self.assertEqual(audit, [])

    def test_type_4_position_is_also_non_structural(self):
        output, _ = self.run_scope_fix(self.reference(master_type=4, cluster_ids=(10,)))

        position = output["positions"][0]
        self.assertEqual(position["position_scope"], "non_structural")
        self.assertEqual(position["position_nomenclature_id"], "10")

    def test_type_5_position_is_corrected_to_pmid(self):
        output, _ = self.run_scope_fix(self.reference(master_type=5, cluster_ids=()))

        position = output["positions"][0]
        self.assertEqual(position["position_scope"], "structural")
        self.assertEqual(position["position_master_id"], "44")
        self.assertIsNone(position["position_nomenclature_id"])

    def test_type_5_position_remains_resolvable_when_old_organization_is_inactive(self):
        reference = self.reference(master_type=5, cluster_ids=())
        reference["position_master_rows"][0].update(
            {
                "is_position_active": 1,
                "is_position_organization_active": 0,
                "is_group_active": 0,
                "is_company_active": 1,
            }
        )

        output, _ = self.run_scope_fix(reference)

        position = output["positions"][0]
        self.assertEqual(position["position_scope"], "structural")
        self.assertEqual(position["position_master_id"], "44")

    def test_non_structural_position_with_multiple_pnids_fails_visibly(self):
        with self.assertRaisesRegex(ValueError, "multiple PNIDs"):
            self.run_scope_fix(self.reference(master_type=6, cluster_ids=(10, 11)))

    def test_blank_identity_is_made_explicit_neglect(self):
        config = self.config()
        config["positions"][0].update(
            {
                "position_scope": None,
                "position_master_id": None,
                "position_nomenclature_id": None,
            }
        )

        output, audit = self.run_scope_fix(self.reference(), config)

        self.assertEqual(output["positions"][0]["position_scope"], "neglect")
        self.assertEqual(audit[0]["resolved_scope"], "neglect")

    def test_validator_rejects_unresolved_scope(self):
        config = self.config()
        config["positions"][0].update(
            {
                "position_scope": None,
                "position_master_id": None,
                "position_nomenclature_id": None,
            }
        )
        config_path = self.write_json("unresolved-config.json", config)
        reference_path = self.write_json("unresolved-reference.json", self.reference())

        errors = validator.check_config_scope(
            config_path, *validator.load_reference_ids(reference_path)
        )

        self.assertTrue(any("unsupported position scope" in error for error in errors))

    def test_validator_rejects_non_structural_pmid_in_structural_config(self):
        config = self.config()
        config["positions"][0].update(
            {
                "position_scope": "structural",
                "position_master_id": "44",
                "position_nomenclature_id": None,
            }
        )
        config_path = self.write_json("invalid-config.json", config)
        reference_path = self.write_json("validator-reference.json", self.reference())
        reference = validator.load_reference_ids(reference_path)

        errors = validator.check_config_scope(config_path, *reference)

        self.assertTrue(
            any("production type 6 is non-structural" in error for error in errors),
            errors,
        )

    def test_validator_accepts_valid_non_structural_pnid_pmid_collision(self):
        config_path = self.write_json("collision-config.json", self.config())
        reference_path = self.write_json(
            "collision-reference.json", self.collision_reference()
        )
        reference = validator.load_reference_ids(reference_path)

        errors = validator.check_config_scope(config_path, *reference)

        self.assertEqual(errors, [])


if __name__ == "__main__":
    unittest.main()
