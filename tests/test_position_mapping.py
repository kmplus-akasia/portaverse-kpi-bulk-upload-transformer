import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from position_mapping import (  # noqa: E402
    HIGH_CONFIDENCE,
    LOW_CONFIDENCE,
    MAPPING_CONFLICT,
    NON_STRUCTURAL,
    NO_CANDIDATE,
    SCOPE_UNCERTAIN,
    STRUCTURAL,
    build_lookup_indexes,
    infer_worksheet_scope,
    mapping_report_row,
    normalize_position_lookup,
    resolve_mapping,
    validate_manual_override,
)


def reference_payload() -> dict:
    return {
        "source": {
            "profile": "production",
            "database": "portaverse",
            "review_status": "current_snapshot_unreviewed",
        },
        "structural_lookup_rows": [
            {
                "position_master_id": "509",
                "position_name": "Group Head Pengelolaan SDM",
                "position_master_type_id": "5",
                "group_name": "Group Pengelolaan SDM",
                "company_id": "1",
                "company_name": "PT Pelabuhan Indonesia (Persero)",
                "company_code": "PLD",
                "active_variant_count": 1,
                "active_employee_count": 2,
                "definitive_employee_count": 2,
                "secondary_employee_count": 0,
                "active_employee_names": ["Budi Santoso", "Citra Dewi"],
                "active_employee_nipps": ["K001", "K002"],
            },
            {
                "position_master_id": "510",
                "position_name": "Group Head Pengembangan SDM",
                "position_master_type_id": "5",
                "group_name": "Group Pengembangan SDM",
                "company_id": "1",
                "company_name": "PT Pelabuhan Indonesia (Persero)",
                "company_code": "PLD",
                "active_variant_count": 1,
                "active_employee_count": 1,
                "definitive_employee_count": 1,
                "secondary_employee_count": 0,
                "active_employee_names": ["Dewi Utami"],
                "active_employee_nipps": ["K003"],
            },
            {
                "position_master_id": "76",
                "position_name": "Manager Keuangan",
                "position_master_type_id": "5",
                "group_name": "Group Keuangan",
                "company_id": "1",
                "company_name": "PT Pelabuhan Indonesia (Persero)",
                "company_code": "PLD",
                "active_variant_count": 1,
                "active_employee_count": 1,
                "definitive_employee_count": 1,
                "secondary_employee_count": 0,
                "active_employee_names": ["Eka Pratama"],
                "active_employee_nipps": ["K004"],
            },
            {
                "position_master_id": "999",
                "position_name": "Manager Vacant",
                "position_master_type_id": "5",
                "group_name": "Group Keuangan",
                "company_id": "1",
                "company_name": "PT Pelabuhan Indonesia (Persero)",
                "company_code": "PLD",
                "active_variant_count": 1,
                "active_employee_count": 0,
                "definitive_employee_count": 0,
                "secondary_employee_count": 0,
            },
        ],
        "non_structural_lookup_rows": [
            {
                "cluster_id": "76",
                "cluster_label": "Officer Keuangan",
                "position_master_id": "701",
                "position_name": "Officer Keuangan",
                "position_master_type_id": "6",
                "group_name": "Group Keuangan",
                "company_id": "1",
                "company_name": "PT Pelabuhan Indonesia (Persero)",
                "company_code": "PLD",
                "active_variant_count": 2,
                "active_employee_count": 2,
                "definitive_employee_count": 0,
                "secondary_employee_count": 2,
                "active_employee_names": ["Fajar Nugroho", "Gita Lestari"],
                "active_employee_nipps": ["K005", "K006"],
            },
            {
                "cluster_id": "80",
                "cluster_label": "Officer Kepatuhan",
                "position_master_id": "702",
                "position_name": "Officer Kepatuhan",
                "position_master_type_id": "6",
                "group_name": "Group Kepatuhan",
                "company_id": "1",
                "company_name": "PT Pelabuhan Indonesia (Persero)",
                "company_code": "PLD",
                "active_variant_count": 1,
                "active_employee_count": 1,
                "definitive_employee_count": 1,
                "secondary_employee_count": 0,
                "active_employee_names": ["Hana Putri"],
                "active_employee_nipps": ["K007"],
            },
        ],
        "organization_rows": [
            {
                "group_master_id": "77",
                "group_name": "Group Pengelolaan SDM",
                "parent_id": "69",
                "company_id": "1",
                "is_group_active": 1,
                "is_company_active": 1,
            },
            {
                "group_master_id": "4364",
                "group_name": "Unit Pendukung Strategi dan Pengelolaan Pembelajaran",
                "parent_id": "77",
                "company_id": "1",
                "is_group_active": 1,
                "is_company_active": 1,
            },
            {
                "group_master_id": "4394",
                "group_name": "Unit Pendukung Talenta dan Rekrutmen",
                "parent_id": "77",
                "company_id": "1",
                "is_group_active": 1,
                "is_company_active": 1,
            },
        ],
    }


class PositionMappingTest(unittest.TestCase):
    def test_normalizes_ampersand_as_dan_for_position_titles(self):
        self.assertEqual(
            normalize_position_lookup("Manager Talenta & Rekrutmen"),
            normalize_position_lookup("Manager Talenta dan Rekrutmen"),
        )

    def test_normalizer_preserves_area_wilayah_and_project_numbers(self):
        self.assertNotEqual(
            normalize_position_lookup("Administrator Operasi Wilayah I Group B Tanjung Priok"),
            normalize_position_lookup("Administrator Operasi Wilayah II Group A Tanjung Priok"),
        )
        self.assertIn("area 2", normalize_position_lookup("Administrator Pelayanan Terminal Area2"))
        self.assertIn("proyek 2", normalize_position_lookup("Officer Pengendalian Kinerja P2"))

    def test_tl_and_project_roles_infer_structural_scope(self):
        self.assertEqual(infer_worksheet_scope("TL Source to Contract 2").scope, STRUCTURAL)
        self.assertEqual(infer_worksheet_scope("Pimpro Satker Single ERP").scope, STRUCTURAL)
        self.assertEqual(infer_worksheet_scope("Principle Expert Auditor").scope, NON_STRUCTURAL)

    def test_infers_scope_from_worksheet_title_before_lookup(self):
        self.assertEqual(infer_worksheet_scope("Group Head Pengelolaan SDM").scope, STRUCTURAL)
        self.assertEqual(infer_worksheet_scope("Officer Keuangan").scope, NON_STRUCTURAL)
        self.assertEqual(infer_worksheet_scope("Officer Manager Keuangan").scope, SCOPE_UNCERTAIN)
        self.assertEqual(infer_worksheet_scope("Kamus KPI Bagian").scope, SCOPE_UNCERTAIN)

    def test_high_confidence_structural_mapping_outputs_pmid_only(self):
        indexes = build_lookup_indexes(reference_payload(), target_company_id="1")

        result = resolve_mapping(
            worksheet="Group Head Pengelolaan SDM",
            worksheet_title="Group Head Pengelolaan SDM",
            group_name="Group Pengelolaan SDM",
            source_workbook="Kamus KPI Group SDM.xlsx",
            indexes=indexes,
        )

        self.assertEqual(result.confidence_label, HIGH_CONFIDENCE)
        self.assertTrue(result.upload_allowed)
        self.assertEqual(result.inferred_scope, STRUCTURAL)
        self.assertEqual(result.position_master_id, "509")
        self.assertIsNone(result.position_nomenclature_id)

    def test_non_structural_mapping_uses_pnid_namespace_despite_numeric_collision(self):
        indexes = build_lookup_indexes(reference_payload(), target_company_id="1")

        result = resolve_mapping(
            worksheet="Officer Keuangan",
            worksheet_title="Officer Keuangan",
            group_name="Group Keuangan",
            source_workbook="Kamus KPI Keuangan.xlsx",
            indexes=indexes,
        )

        self.assertEqual(result.confidence_label, HIGH_CONFIDENCE)
        self.assertEqual(result.inferred_scope, NON_STRUCTURAL)
        self.assertIsNone(result.position_master_id)
        self.assertEqual(result.position_nomenclature_id, "76")

    def test_low_confidence_candidate_is_blocked_until_review(self):
        indexes = build_lookup_indexes(reference_payload(), target_company_id="1")

        result = resolve_mapping(
            worksheet="Officer Keu",
            worksheet_title="Officer Keu",
            group_name="Group Keuangan",
            source_workbook="Kamus KPI Keuangan.xlsx",
            indexes=indexes,
        )

        self.assertEqual(result.confidence_label, LOW_CONFIDENCE)
        self.assertFalse(result.upload_allowed)
        self.assertEqual(result.position_nomenclature_id, "76")

    def test_duplicate_strong_candidates_are_mapping_conflict(self):
        payload = reference_payload()
        payload["non_structural_lookup_rows"].append(
            {
                "cluster_id": "77",
                "cluster_label": "Officer Keuangan",
                "position_master_id": "703",
                "position_name": "Officer Keuangan",
                "position_master_type_id": "6",
                "group_name": "Group Keuangan",
                "company_id": "1",
                "company_name": "PT Pelabuhan Indonesia (Persero)",
                "company_code": "PLD",
                "active_variant_count": 1,
                "active_employee_count": 1,
                "definitive_employee_count": 1,
                "secondary_employee_count": 0,
                "active_employee_names": ["Indra Wijaya"],
                "active_employee_nipps": ["K008"],
            }
        )
        indexes = build_lookup_indexes(payload, target_company_id="1")

        result = resolve_mapping(
            worksheet="Officer Keuangan",
            worksheet_title="Officer Keuangan",
            group_name="Group Keuangan",
            source_workbook="Kamus KPI Keuangan.xlsx",
            indexes=indexes,
        )

        self.assertEqual(result.confidence_label, MAPPING_CONFLICT)
        self.assertFalse(result.upload_allowed)

    def test_scope_uncertain_never_forces_lookup(self):
        indexes = build_lookup_indexes(reference_payload(), target_company_id="1")

        result = resolve_mapping(
            worksheet="Officer Manager Keuangan",
            worksheet_title="Officer Manager Keuangan",
            group_name="Group Keuangan",
            source_workbook="Kamus KPI Keuangan.xlsx",
            indexes=indexes,
        )

        self.assertEqual(result.confidence_label, SCOPE_UNCERTAIN)
        self.assertFalse(result.upload_allowed)
        self.assertIsNone(result.position_master_id)
        self.assertIsNone(result.position_nomenclature_id)

    def test_mapping_report_row_aggregates_active_employee_names_and_nipps(self):
        indexes = build_lookup_indexes(reference_payload(), target_company_id="1")
        result = resolve_mapping(
            worksheet="Officer Keuangan",
            worksheet_title="Officer Keuangan",
            group_name="Group Keuangan",
            source_workbook="Kamus KPI Keuangan.xlsx",
            indexes=indexes,
        )

        row = mapping_report_row(result)

        self.assertEqual(row["Active Employee Name"], "Fajar Nugroho; Gita Lestari")
        self.assertEqual(row["Active Employee NIPP"], "K005; K006")
        self.assertEqual(row["Recommended Action"], "No action required; auto-mapped.")

    def test_manual_override_validation_requires_active_id_in_matching_scope(self):
        indexes = build_lookup_indexes(reference_payload(), target_company_id="1")

        valid = validate_manual_override(
            inferred_scope=NON_STRUCTURAL,
            position_master_id=None,
            position_nomenclature_id="76",
            indexes=indexes,
        )
        wrong_namespace = validate_manual_override(
            inferred_scope=NON_STRUCTURAL,
            position_master_id="76",
            position_nomenclature_id=None,
            indexes=indexes,
        )
        inactive = validate_manual_override(
            inferred_scope=STRUCTURAL,
            position_master_id="999",
            position_nomenclature_id=None,
            indexes=indexes,
        )

        self.assertTrue(valid.allowed)
        self.assertFalse(wrong_namespace.allowed)
        self.assertEqual(wrong_namespace.reason, "Manual non-structural override must provide PNID only.")
        self.assertFalse(inactive.allowed)
        self.assertEqual(inactive.reason, "Manual structural PMID is not present in active structural lookup.")

    def test_no_candidate_when_scope_known_but_active_lookup_has_no_match(self):
        indexes = build_lookup_indexes(reference_payload(), target_company_id="1")

        result = resolve_mapping(
            worksheet="Auditor Teknologi Informasi",
            worksheet_title="Auditor Teknologi Informasi",
            group_name="Group SPI",
            source_workbook="Kamus KPI SPI.xlsx",
            indexes=indexes,
        )

        self.assertEqual(result.confidence_label, NO_CANDIDATE)
        self.assertFalse(result.upload_allowed)

    def test_structural_candidate_matches_group_ancestor_context(self):
        payload = reference_payload()
        payload["structural_lookup_rows"].append(
            {
                "position_master_id": "37526",
                "position_name": "Manager Talenta & Rekrutmen",
                "position_master_type_id": "5",
                "group_master_id": "4394",
                "group_name": "Unit Pendukung Talenta dan Rekrutmen",
                "company_id": "1",
                "company_name": "PT Pelabuhan Indonesia (Persero)",
                "company_code": "PLD",
                "active_variant_count": 1,
                "active_employee_count": 1,
                "definitive_employee_count": 1,
                "secondary_employee_count": 0,
                "active_employee_names": ["Fajar Apriadi Darwita"],
                "active_employee_nipps": ["106316"],
            }
        )
        indexes = build_lookup_indexes(payload, target_company_id="1")

        result = resolve_mapping(
            worksheet="Manager Talenta-Rekrutmen",
            worksheet_title="Manager Talenta dan Rekrutmen",
            group_name="Group Pengelolaan SDM",
            source_workbook="Kamus KPI Group Pengelolaan SDM.xlsx",
            indexes=indexes,
        )

        self.assertEqual(result.confidence_label, HIGH_CONFIDENCE)
        self.assertEqual(result.position_master_id, "37526")
        self.assertEqual(result.candidate_group, "Unit Pendukung Talenta dan Rekrutmen")

    def test_non_structural_candidate_matches_group_ancestor_context(self):
        payload = reference_payload()
        payload["non_structural_lookup_rows"].append(
            {
                "cluster_id": "900",
                "cluster_label": "Officer Strategi dan Pengelolaan Pembelajaran",
                "position_master_id": "36018",
                "position_name": "Officer II Strategi dan Pengelolaan Pembelajaran",
                "position_master_type_id": "6",
                "group_master_id": "4364",
                "group_name": "Unit Pendukung Strategi dan Pengelolaan Pembelajaran",
                "company_id": "1",
                "company_name": "PT Pelabuhan Indonesia (Persero)",
                "company_code": "PLD",
                "active_variant_count": 1,
                "active_employee_count": 1,
                "definitive_employee_count": 1,
                "secondary_employee_count": 0,
                "active_employee_names": ["Wahid Rahmat Setiawan"],
                "active_employee_nipps": ["107105"],
            }
        )
        indexes = build_lookup_indexes(payload, target_company_id="1")

        result = resolve_mapping(
            worksheet="Officer Strategi-Pengelolaan",
            worksheet_title="Officer Strategi & Pengelolaan Pembelajaran",
            group_name="Group Pengelolaan SDM",
            source_workbook="Kamus KPI Group Pengelolaan SDM.xlsx",
            indexes=indexes,
        )

        self.assertEqual(result.confidence_label, HIGH_CONFIDENCE)
        self.assertEqual(result.position_nomenclature_id, "900")

    def test_duplicate_rows_for_same_pnid_do_not_block_high_confidence(self):
        payload = reference_payload()
        payload["non_structural_lookup_rows"].extend(
            [
                {
                    "cluster_id": "901",
                    "cluster_label": "Officer Evaluasi Pembelajaran dan Manajemen Pengetahuan",
                    "position_master_id": "35928",
                    "position_name": "Senior Officer III Evaluasi Pembelajaran dan Manajemen Pengetahuan",
                    "position_master_type_id": "6",
                    "group_master_id": "4364",
                    "group_name": "Unit Pendukung Strategi dan Pengelolaan Pembelajaran",
                    "company_id": "1",
                    "company_name": "PT Pelabuhan Indonesia (Persero)",
                    "company_code": "PLD",
                    "active_variant_count": 1,
                    "active_employee_count": 1,
                    "active_employee_names": ["Wita Fitriani"],
                    "active_employee_nipps": ["106006"],
                },
                {
                    "cluster_id": "901",
                    "cluster_label": "Officer Evaluasi Pembelajaran dan Manajemen Pengetahuan",
                    "position_master_id": "35974",
                    "position_name": "Senior Officer II Evaluasi Pembelajaran dan Manajemen Pengetahuan",
                    "position_master_type_id": "6",
                    "group_master_id": "4364",
                    "group_name": "Unit Pendukung Strategi dan Pengelolaan Pembelajaran",
                    "company_id": "1",
                    "company_name": "PT Pelabuhan Indonesia (Persero)",
                    "company_code": "PLD",
                    "active_variant_count": 1,
                    "active_employee_count": 1,
                    "active_employee_names": ["Meutya Dewi Aprilia"],
                    "active_employee_nipps": ["106566"],
                },
            ]
        )
        indexes = build_lookup_indexes(payload, target_company_id="1")

        result = resolve_mapping(
            worksheet="Officer Evaluasi Pembelajaran d",
            worksheet_title="Officer Evaluasi Pembelajaran dan Manajemen Pengetahuan",
            group_name="Group Pengelolaan SDM",
            source_workbook="Kamus KPI Group Pengelolaan SDM.xlsx",
            indexes=indexes,
        )

        self.assertEqual(result.confidence_label, HIGH_CONFIDENCE)
        self.assertEqual(result.position_nomenclature_id, "901")
        self.assertEqual(result.active_employee_count, 2)
        self.assertEqual(result.active_employee_name, "Wita Fitriani; Meutya Dewi Aprilia")
        self.assertEqual(result.active_employee_nipp, "106006; 106566")


if __name__ == "__main__":
    unittest.main()
