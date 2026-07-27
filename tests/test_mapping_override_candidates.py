import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from build_mapping_override_candidates import build_workbook, build_override_rows  # noqa: E402


class MappingOverrideCandidatesTest(unittest.TestCase):
    def test_build_override_rows_classifies_high_confidence_and_ambiguous(self):
        base = {
            "Source Workbook": "source.xlsx",
            "Sheet": "Officer A",
            "Raw Group": "Group A",
            "Raw Position": "Officer A",
            "Direktorat": "Direktorat A",
            "Candidate Scope": "structural",
            "Candidate PMID": "100",
            "Candidate PNID": "",
            "Candidate Title": "Officer A",
            "Candidate Group": "Group A",
            "Candidate Company": "PT A",
            "Candidate Company Code": "A",
            "Candidate Score": 0.92,
            "Match Reason": "title=1.00; group=0.80",
        }
        ambiguous = dict(base)
        ambiguous.update(
            {
                "Sheet": "Officer B",
                "Raw Position": "Officer B",
                "Candidate PMID": "200",
                "Candidate Score": 0.70,
            }
        )
        runner_up = dict(ambiguous)
        runner_up.update({"Candidate PMID": "201", "Candidate Score": 0.66})

        rows = build_override_rows([base, ambiguous, runner_up])

        by_sheet = {row["Sheet"]: row for row in rows}
        self.assertEqual(by_sheet["Officer A"]["Review Status"], "review_recommended_high_confidence")
        self.assertEqual(by_sheet["Officer A"]["Suggested Position Master ID"], "100")
        self.assertEqual(by_sheet["Officer B"]["Review Status"], "ambiguous_candidates")
        self.assertEqual(by_sheet["Officer B"]["Candidate Count"], 2)

    def test_build_workbook_adds_review_tabs_and_approval_validation(self):
        rows = [
            {
                "Review Status": "review_recommended_high_confidence",
                "Approved": "",
                "Source Workbook": "source.xlsx",
                "Sheet": "Officer A",
                "Raw Group": "Group A",
                "Raw Position": "Officer A",
                "Direktorat": "Direktorat A",
                "Suggested Position Scope": "structural",
                "Suggested Position Master ID": "100",
                "Suggested Position Nomenklatur ID": "",
                "Suggested Position Title": "Officer A",
                "Suggested Group": "Group A",
                "Suggested Company": "PT A",
                "Suggested Company Code": "A",
                "Candidate Score": 0.92,
                "Runner-up Score": "",
                "Candidate Count": 1,
                "Match Reason": "title=1.00",
                "Reviewer Notes": "",
            }
        ]

        workbook = build_workbook(rows, Path("review.xlsx"))

        self.assertIn("High Confidence", workbook.sheetnames)
        self.assertIn("All Recommended", workbook.sheetnames)
        self.assertIn("Override Candidates", workbook.sheetnames)
        validations = list(workbook["Override Candidates"].data_validations.dataValidation)
        self.assertEqual(len(validations), 1)
        self.assertEqual(validations[0].formula1, '"YES,NO"')


if __name__ == "__main__":
    unittest.main()
