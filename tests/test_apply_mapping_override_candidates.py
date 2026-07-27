import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "apply_mapping_override_candidates.py"


class ApplyMappingOverrideCandidatesTest(unittest.TestCase):
    def test_apply_only_explicitly_approved_rows(self):
        config = {
            "reference_source": {},
            "positions": [
                {
                    "source_workbook": "source-a.xlsx",
                    "sheet_name": "Role A",
                    "position_name": "Role A",
                    "group_name": "Group A",
                    "directorate_name": "Directorate A",
                    "position_scope": "mapping_conflict",
                    "position_master_id": None,
                    "position_nomenclature_id": None,
                },
                {
                    "source_workbook": "source-b.xlsx",
                    "sheet_name": "Role B",
                    "position_name": "Role B",
                    "group_name": "Group B",
                    "directorate_name": "Directorate B",
                    "position_scope": "mapping_conflict",
                    "position_master_id": None,
                    "position_nomenclature_id": None,
                },
            ],
        }

        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            config_path = tmp_path / "config.json"
            output_path = tmp_path / "output.json"
            overrides_path = tmp_path / "overrides.xlsx"
            config_path.write_text(json.dumps(config), encoding="utf-8")
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = "Override Candidates"
            worksheet.append(
                [
                    "Approved",
                    "Source Workbook",
                    "Sheet",
                    "Suggested Position Scope",
                    "Suggested Position Master ID",
                    "Suggested Position Nomenklatur ID",
                    "Suggested Position Title",
                    "Suggested Group",
                    "Suggested Company",
                ]
            )
            worksheet.append(["YES", "source-a.xlsx", "Role A", "structural", "100", "", "Role A Portaverse", "Group A PV", "PT A"])
            worksheet.append(["", "source-b.xlsx", "Role B", "non_structural", "", "200", "Role B Cluster", "Group B PV", "PT B"])
            workbook.save(overrides_path)

            result = subprocess.run(
                [
                    sys.executable,
                    str(SCRIPT),
                    "--config",
                    str(config_path),
                    "--overrides",
                    str(overrides_path),
                    "--output",
                    str(output_path),
                ],
                check=True,
                text=True,
                capture_output=True,
            )

            self.assertIn("applied_rows=1", result.stdout)
            output = json.loads(output_path.read_text(encoding="utf-8"))
            first, second = output["positions"]
            self.assertEqual(first["position_scope"], "structural")
            self.assertEqual(first["position_master_id"], "100")
            self.assertIsNone(first["position_nomenclature_id"])
            self.assertEqual(first["portaverse_position_title"], "Role A Portaverse")
            self.assertEqual(second["position_scope"], "mapping_conflict")
            self.assertIsNone(second["position_master_id"])
            self.assertIsNone(second["position_nomenclature_id"])


if __name__ == "__main__":
    unittest.main()
