import json
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook


class ApplyPositionMappingReviewTest(unittest.TestCase):
    def test_applies_manual_yes_rows_and_holds_needs_check_rows(self):
        from scripts.apply_position_mapping_review import apply_review_to_config

        base_config = {
            "positions": [
                {
                    "source_workbook": "Group A/book.xlsx",
                    "sheet_name": "Manager A",
                    "position_name": "Manager A",
                    "position_scope": "mapping_conflict",
                    "position_master_id": "111",
                    "position_nomenclature_id": None,
                    "mapping_confidence_label": "mapping_conflict",
                },
                {
                    "source_workbook": "Group A/book.xlsx",
                    "sheet_name": "Officer A",
                    "position_name": "Officer A",
                    "position_scope": "non_structural",
                    "position_master_id": None,
                    "position_nomenclature_id": "222",
                    "mapping_confidence_label": "low_confidence",
                },
                {
                    "source_workbook": "Group A/book.xlsx",
                    "sheet_name": "Officer Pending",
                    "position_name": "Officer Pending",
                    "position_scope": "non_structural",
                    "position_master_id": None,
                    "position_nomenclature_id": "333",
                    "mapping_confidence_label": "low_confidence",
                },
            ]
        }

        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            config_path = temp / "config.json"
            review_path = temp / "review.xlsx"
            output_path = temp / "reviewed.json"
            config_path.write_text(json.dumps(base_config), encoding="utf-8")
            self._write_review_workbook(review_path)

            stats = apply_review_to_config(config_path, review_path, output_path)

            reviewed = json.loads(output_path.read_text(encoding="utf-8"))
            positions = {
                (row["source_workbook"], row["sheet_name"]): row
                for row in reviewed["positions"]
            }

        manager = positions[("Group A/book.xlsx", "Manager A")]
        self.assertEqual(manager["position_scope"], "structural")
        self.assertEqual(manager["position_master_id"], "999")
        self.assertIsNone(manager["position_nomenclature_id"])
        self.assertTrue(manager["mapping_override_approved"])
        self.assertEqual(manager["mapping_review_status"], "approved")
        self.assertEqual(manager["mapping_override_trust_source"], "reviewer_manual")

        officer = positions[("Group A/book.xlsx", "Officer A")]
        self.assertEqual(officer["position_scope"], "non_structural")
        self.assertIsNone(officer["position_master_id"])
        self.assertEqual(officer["position_nomenclature_id"], "888")
        self.assertEqual(officer["mapping_override_trust_source"], "reviewer_manual")

        pending = positions[("Group A/book.xlsx", "Officer Pending")]
        self.assertEqual(pending["mapping_review_status"], "needs_check")
        self.assertFalse(pending["mapping_override_approved"])
        self.assertIsNone(pending.get("mapping_override_trust_source"))
        self.assertEqual(pending["position_nomenclature_id"], "333")

        self.assertEqual(stats["review_yes_rows"], 2)
        self.assertEqual(stats["review_needs_check_rows"], 1)
        self.assertEqual(stats["manual_override_rows"], 2)

    def test_allows_identical_duplicate_review_rows(self):
        from scripts.apply_position_mapping_review import apply_review_to_config

        base_config = {
            "positions": [
                {
                    "source_workbook": "Group A/book.xlsx",
                    "sheet_name": "Officer A",
                    "position_name": "Officer A",
                    "position_scope": "non_structural",
                    "position_master_id": None,
                    "position_nomenclature_id": "222",
                    "mapping_confidence_label": "high_confidence",
                },
            ]
        }

        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            config_path = temp / "config.json"
            review_path = temp / "review.xlsx"
            output_path = temp / "reviewed.json"
            config_path.write_text(json.dumps(base_config), encoding="utf-8")
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Position Mapping Report"
            sheet.append(
                [
                    "Source Workbook",
                    "Worksheet",
                    "Reviewer Confirm Mapping",
                    "Reviewer Actual PMID",
                    "Reviewer Actual PNID",
                    "Candidate PNID",
                ]
            )
            duplicate = ["Group A/book.xlsx", "Officer A", "YES", None, None, "222"]
            sheet.append(duplicate)
            sheet.append(duplicate)
            workbook.save(review_path)

            stats = apply_review_to_config(config_path, review_path, output_path)

        self.assertEqual(stats["review_yes_rows"], 1)

    def _write_review_workbook(self, path: Path) -> None:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Position Mapping Report"
        sheet.append(
            [
                "Source Workbook",
                "Worksheet",
                "Reviewer Confirm Mapping",
                "Reviewer Actual PMID",
                "Reviewer Actual PNID",
            ]
        )
        sheet.append(["Group A/book.xlsx", "Manager A", "YES", 999, None])
        sheet.append(["Group A/book.xlsx", "Officer A", "YES", None, 888])
        sheet.append(["Group A/book.xlsx", "Officer Pending", "NEEDS_CHECK", None, 777])
        workbook.save(path)


if __name__ == "__main__":
    unittest.main()
