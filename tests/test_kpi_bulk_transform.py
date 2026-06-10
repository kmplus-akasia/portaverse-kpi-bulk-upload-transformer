import json
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from kpi_bulk_transform import (  # noqa: E402
    ImpactRecord,
    NormalizedEnum,
    NormalizationStatus,
    PositionConfig,
    UPLOAD_HEADERS,
    append_enum_issue,
    build_upload_rows,
    is_active_valid_sheet,
    load_config,
    load_nomenclature_mapping,
    parse_block_sheet,
    normalize_cascading,
    normalize_kai_nature,
    normalize_ownership_type,
    normalize_period,
    normalize_polarity,
    refresh_configs_from_mapping,
    uploader_kai_nature,
    uploader_polarity,
    validate_output_rows,
    write_output_workbook,
)

from openpyxl import Workbook, load_workbook  # noqa: E402


class KpiBulkTransformTest(unittest.TestCase):
    def test_latest_upload_headers_include_optional_pnid_columns(self):
        self.assertEqual(len(UPLOAD_HEADERS), 24)
        self.assertEqual(
            UPLOAD_HEADERS[-3:],
            ["Ownership Type", "Position Nomenklatur ID", "RKM Code ID"],
        )
        self.assertEqual(UPLOAD_HEADERS[18], "Nature Of Work (KAI Only)")
        self.assertEqual(UPLOAD_HEADERS[21], "Ownership Type")

    def test_position_nomenclature_id_overrides_position_master_id_in_output_rows(self):
        config = PositionConfig(
            sheet_name="Officer Kinerja Individu",
            position_name="Officer I Kinerja Individu",
            group_name="Group Pengelolaan SDM",
            directorate_name="Direktorat Sumber Daya Manusia dan Umum",
            position_master_id="528",
            position_nomenclature_id="76",
            position_scope="non_structural",
        )
        impacts = [
            ImpactRecord(
                bsc="Financial",
                title="Net Income",
                unit="Rupiah",
                period="Triwulan",
                formula="Revenue - Cost",
                polarity="Positif",
                weight="15",
            )
        ]

        rows, next_id = build_upload_rows(config, "528", impacts, 1)

        self.assertEqual(next_id, 2)
        self.assertEqual(len(rows[0]), 24)
        row_map = dict(zip(UPLOAD_HEADERS, rows[0]))
        self.assertIsNone(row_map["Position Master ID (Required)"])
        self.assertEqual(row_map["Position Nomenklatur ID"], "76")
        self.assertIsNone(row_map["RKM Code ID"])

    def test_structural_scope_uses_only_position_master_id_even_when_pnid_exists(self):
        config = PositionConfig(
            sheet_name="Group Head",
            position_name="Group Head Pengelolaan SDM",
            group_name="Group Pengelolaan SDM",
            directorate_name="Direktorat Sumber Daya Manusia dan Umum",
            position_master_id="509",
            position_nomenclature_id="74",
            position_scope="structural",
        )
        impact = ImpactRecord(
            bsc="Financial",
            title="Net Income",
            unit="Rupiah",
            period="Triwulan",
            formula="Revenue - Cost",
            polarity="Positif",
            weight="15",
        )

        rows, _ = build_upload_rows(config, "509", [impact], 1)

        row_map = dict(zip(UPLOAD_HEADERS, rows[0]))
        self.assertEqual(row_map["Position Master ID (Required)"], "509")
        self.assertIsNone(row_map["Position Nomenklatur ID"])

    def test_validate_output_rows_rejects_invalid_upload_enums_and_dual_position_ids(self):
        config = PositionConfig(
            sheet_name="Manager Rekrutmen-Karir",
            position_name="Manager Rekrutmen dan Karir",
            group_name="Group Pengelolaan SDM",
            directorate_name="Direktorat SDM & Umum",
        )
        impact = dict.fromkeys(UPLOAD_HEADERS)
        impact.update(
            {
                "IDKPI": "1",
                "Group": "Group Pengelolaan SDM",
                "Direktorat": "Direktorat SDM & Umum",
                "Posisi": "Manager Rekrutmen dan Karir",
                "Position Master ID (Required)": "515",
                "BSC Perspective": "Learning & Growth",
                "KPI Type": "IMPACT",
                "Title": "Pemenuhan formasi",
                "Unit": "%",
                "Polarity": "POSITIVE",
                "Period": "PER TAHUN",
                "Formula": "realisasi/target",
                "Weight (%)": "10",
                "Cascading": "SPECIFIC",
                "Ownership Type": "Non Routine",
                "Position Nomenklatur ID": "515",
            }
        )
        kai = dict(impact)
        kai.update(
            {
                "IDKPI": "2",
                "KPI Type": "KAI",
                "Parent KPI ID": "1",
                "Parent KPI Title": "Pemenuhan formasi",
                "Title": "Follow up rekrutmen",
                "Period": "TAHUNAN",
                "Cascading": "DIRECT",
                "Nature Of Work (KAI Only)": "Kadang",
                "Ownership Type": "SPECIFIC",
                "Position Nomenklatur ID": None,
            }
        )
        issues = []

        validate_output_rows(config, [[row[header] for header in UPLOAD_HEADERS] for row in [impact, kai]], issues)

        messages = [issue.message for issue in issues]
        self.assertIn("Invalid Period enum: PER TAHUN", messages)
        self.assertIn("Invalid Cascading enum: SPECIFIC", messages)
        self.assertIn("Invalid Ownership Type enum: Non Routine", messages)
        self.assertIn("Invalid upload scope: row has both PMID and PNID.", messages)
        self.assertIn("Invalid KAI Nature enum: Kadang", messages)

    def test_load_config_treats_zero_position_master_id_as_missing_and_loads_pnid(self):
        payload = {
            "positions": [
                {
                    "sheet_name": "Officer Kinerja Individu",
                    "position_name": "Officer I Kinerja Individu",
                    "group_name": "Group Pengelolaan SDM",
                    "directorate_name": "Direktorat Sumber Daya Manusia dan Umum",
                    "position_master_id": 0,
                    "position_nomenclature_id": 76,
                    "position_scope": "non_structural",
                }
            ]
        }
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "config.json"
            path.write_text(json.dumps(payload), encoding="utf-8")

            configs = load_config(path)

        self.assertIsNone(configs[0].position_master_id)
        self.assertEqual(configs[0].position_nomenclature_id, "76")
        self.assertEqual(configs[0].position_scope, "non_structural")

    def test_load_config_structural_scope_drops_pnid_and_non_structural_scope_drops_pmid(self):
        payload = {
            "positions": [
                {
                    "sheet_name": "Group Head",
                    "position_name": "Group Head",
                    "group_name": "Group",
                    "directorate_name": "Direktorat",
                    "position_master_id": 946,
                    "position_nomenclature_id": 99,
                    "position_scope": "structural",
                },
                {
                    "sheet_name": "Officer",
                    "position_name": "Officer",
                    "group_name": "Group",
                    "directorate_name": "Direktorat",
                    "position_master_id": 528,
                    "position_nomenclature_id": 82,
                    "position_scope": "non_structural",
                },
            ]
        }
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "config.json"
            path.write_text(json.dumps(payload), encoding="utf-8")

            configs = load_config(path)

        self.assertEqual(configs[0].position_master_id, "946")
        self.assertIsNone(configs[0].position_nomenclature_id)
        self.assertIsNone(configs[1].position_master_id)
        self.assertEqual(configs[1].position_nomenclature_id, "82")

    def test_output_and_kai_inherit_impact_polarity_and_default_kai_nature(self):
        config = PositionConfig(
            sheet_name="Officer Kinerja Individu",
            position_name="Officer I Kinerja Individu",
            group_name="Group Pengelolaan SDM",
            directorate_name="Direktorat Sumber Daya Manusia dan Umum",
        )
        rows = [
            [
                "Jenis Posisi",
                "BSC Perspective",
                "KPI Impact",
                "KPI Impact Unit",
                "KPI Impact Frequency",
                "KPI Impact Formula",
                "%Weight (Impact)",
                "KPI Impact Polarity",
                "KPI Output",
                "%Weight (Output)",
                "KPI Output Definition",
                "KPI Output Unit",
                "KPI Output Frequency",
                "KPI Output Formula",
                "KPI Output Polarity",
                "Coverage KPI Output",
                "Cascading Tagging (KPI Output)",
                "Key Activity Indicator (KAI)",
                "%Weight (Activity)",
            ],
            [
                "Struktural",
                "Financial",
                "Net Income",
                "Rupiah",
                "Triwulan",
                "Revenue - Cost",
                "15",
                "Positif",
                "Penyempurnaan PMS",
                "10",
                "Enhancement PMS",
                "%",
                "Triwulan",
                "realisasi/target",
                None,
                "SPECIFIC",
                "DIRECT",
                "update SOP",
                "5",
            ],
        ]
        issues = []

        impacts = parse_block_sheet(rows, config, issues)

        self.assertEqual(impacts[0].source_row, 2)
        self.assertEqual(impacts[0].outputs[0]["polarity"], "Positif")
        self.assertEqual(impacts[0].outputs[0]["kai"]["polarity"], "Positif")
        self.assertIsNone(impacts[0].outputs[0]["kai"]["nature_of_work"])

    def test_missing_polarity_defaults_to_positive_and_missing_kai_nature_by_period(self):
        config = PositionConfig(
            sheet_name="Officer",
            position_name="Officer",
            group_name="Group Raw",
            directorate_name="Direktorat Raw",
            position_nomenclature_id="333",
            position_scope="non_structural",
        )
        impact = ImpactRecord(
            bsc="Financial",
            title="Revenue",
            unit="Rupiah",
            period="Triwulan",
            formula="Revenue",
            polarity=None,
            weight="100",
            outputs=[
                {
                    "source_row": 2,
                    "title": "Output",
                    "description": "desc",
                    "unit": "%",
                    "period": "Triwulan",
                    "formula": "realisasi/target",
                    "polarity": None,
                    "weight": "100",
                    "cascading": "DIRECT",
                    "ownership_type": "SPECIFIC",
                    "kai": {
                        "source_row": 2,
                        "title": "KAI",
                        "description": "desc",
                        "formula": "realisasi/target",
                        "weight": "100",
                        "nature_of_work": None,
                        "period": "Triwulan",
                        "polarity": None,
                        "cascading": "DIRECT",
                        "ownership_type": "SPECIFIC",
                    },
                }
            ],
        )

        rows, _ = build_upload_rows(config, None, [impact], 1)
        row_maps = [dict(zip(UPLOAD_HEADERS, row)) for row in rows]

        self.assertEqual([row["Polarity"] for row in row_maps], ["POSITIVE", "POSITIVE", "POSITIVE"])
        self.assertEqual(row_maps[2]["Nature Of Work (KAI Only)"], "Routine")

        impact.outputs[0]["kai"]["period"] = "Tahunan"
        rows, _ = build_upload_rows(config, None, [impact], 1)
        yearly_kai = dict(zip(UPLOAD_HEADERS, rows[2]))
        self.assertEqual(yearly_kai["Nature Of Work (KAI Only)"], "Non Routine")

    def test_append_enum_issue_uses_warning_for_defaulted_and_error_for_invalid_without_value(self):
        config = PositionConfig(
            sheet_name="Officer",
            position_name="Officer",
            group_name="Group",
            directorate_name="Direktorat",
        )
        issues = []

        append_enum_issue(
            issues,
            config,
            6,
            "IMPACT",
            "Impact Title",
            "Polarity",
            NormalizedEnum(
                value="POSITIVE",
                status=NormalizationStatus.NORMALIZED,
                raw_value="Positif",
                message="Polarity normalized to POSITIVE.",
            ),
        )
        append_enum_issue(
            issues,
            config,
            7,
            "OUTPUT",
            "Output Title",
            "Period",
            NormalizedEnum(
                value="TRIWULANAN",
                status=NormalizationStatus.DEFAULTED,
                raw_value=None,
                message="Period missing; defaulted to fallback period TRIWULANAN.",
            ),
        )
        append_enum_issue(
            issues,
            config,
            8,
            "OUTPUT",
            "Output Title",
            "Period",
            NormalizedEnum(
                value=None,
                status=NormalizationStatus.INVALID,
                raw_value="bogus",
                message="Invalid period.",
            ),
        )

        self.assertEqual(issues[0].severity, "info")
        self.assertIn(
            "enum_issue category=normalized; field=Polarity; raw=Positif; normalized=POSITIVE; Polarity normalized to POSITIVE.",
            issues[0].message,
        )
        self.assertEqual(issues[1].severity, "warning")
        self.assertIn(
            "enum_issue category=defaulted; field=Period; raw=None; normalized=TRIWULANAN; Period missing; defaulted to fallback period TRIWULANAN.",
            issues[1].message,
        )
        self.assertEqual(issues[2].severity, "error")
        self.assertIn(
            "enum_issue category=invalid; field=Period; raw=bogus; normalized=None; Invalid period.",
            issues[2].message,
        )

    def test_build_upload_rows_uses_impact_source_row_for_normalized_polarity_issue(self):
        config = PositionConfig(
            sheet_name="Officer",
            position_name="Officer",
            group_name="Group Raw",
            directorate_name="Direktorat Raw",
            position_nomenclature_id="333",
            position_scope="non_structural",
        )
        impact = ImpactRecord(
            source_row=12,
            bsc="Financial",
            title="Revenue",
            unit="Rupiah",
            period="TRIWULANAN",
            formula="Revenue",
            polarity="Positif",
            weight="100",
        )
        issues = []

        rows, _ = build_upload_rows(config, None, [impact], 1, issues)

        self.assertEqual(dict(zip(UPLOAD_HEADERS, rows[0]))["Polarity"], "POSITIVE")
        self.assertEqual(len(issues), 1)
        self.assertEqual(issues[0].severity, "info")
        self.assertEqual(issues[0].source_row, 12)
        self.assertEqual(issues[0].record_type, "IMPACT")
        self.assertIn(
            "enum_issue category=normalized; field=Polarity; raw=Positif; normalized=POSITIVE; Polarity normalized to POSITIVE.",
            issues[0].message,
        )

    def test_parsed_kai_inherited_fields_do_not_duplicate_output_enum_issues(self):
        config = PositionConfig(
            sheet_name="Officer Kinerja Individu",
            position_name="Officer I Kinerja Individu",
            group_name="Group Pengelolaan SDM",
            directorate_name="Direktorat Sumber Daya Manusia dan Umum",
        )
        rows = [
            [
                "Jenis Posisi",
                "BSC Perspective",
                "KPI Impact",
                "KPI Impact Unit",
                "KPI Impact Frequency",
                "KPI Impact Formula",
                "%Weight (Impact)",
                "KPI Impact Polarity",
                "KPI Output",
                "%Weight (Output)",
                "KPI Output Definition",
                "KPI Output Unit",
                "KPI Output Frequency",
                "KPI Output Formula",
                "KPI Output Polarity",
                "Coverage KPI Output",
                "Cascading Tagging (KPI Output)",
                "Nature of Work (KAI)",
                "Key Activity Indicator (KAI)",
                "%Weight (Activity)",
            ],
            [
                "Struktural",
                "Financial",
                "Net Income",
                "Rupiah",
                "TRIWULANAN",
                "Revenue - Cost",
                "100",
                "POSITIVE",
                "Penyempurnaan PMS",
                "100",
                "Enhancement PMS",
                "%",
                "Triwulan",
                "realisasi/target",
                "Positif",
                "Non Routine",
                "SPECIFIC",
                "Routine",
                "update SOP",
                "100",
            ],
        ]
        parse_issues = []
        issues = []

        impacts = parse_block_sheet(rows, config, parse_issues)
        self.assertEqual(impacts[0].source_row, 2)

        upload_rows, _ = build_upload_rows(config, "528", impacts, 1, issues)
        row_maps = [dict(zip(UPLOAD_HEADERS, row)) for row in upload_rows]

        self.assertEqual([row["KPI Type"] for row in row_maps], ["IMPACT", "OUTPUT", "KAI"])
        messages = [issue.message for issue in issues]
        self.assertTrue(any("enum_issue category=normalized; field=Period" in message for message in messages))
        self.assertTrue(any("enum_issue category=normalized; field=Polarity" in message for message in messages))
        self.assertTrue(any("enum_issue category=cross_column; field=Cascading" in message for message in messages))
        self.assertTrue(any("enum_issue category=cross_column; field=Ownership Type" in message for message in messages))
        self.assertFalse(
            any(
                issue.record_type == "KAI" and "field=Cascading" in issue.message
                for issue in issues
            )
        )
        self.assertFalse(
            any(
                issue.record_type == "KAI" and "field=Ownership Type" in issue.message
                for issue in issues
            )
        )

    def test_build_upload_rows_reports_ambiguous_output_period(self):
        config = PositionConfig(
            sheet_name="Officer Kinerja Individu",
            position_name="Officer I Kinerja Individu",
            group_name="Group Pengelolaan SDM",
            directorate_name="Direktorat Sumber Daya Manusia dan Umum",
        )
        impact = ImpactRecord(
            bsc="Financial",
            title="Net Income",
            unit="Rupiah",
            period="Triwulan",
            formula="Revenue - Cost",
            polarity="Positif",
            weight="15",
            outputs=[
                {
                    "source_row": 2,
                    "title": "Penyempurnaan PMS",
                    "description": "Enhancement PMS",
                    "unit": "%",
                    "period": "Triwulanan/Tahunan",
                    "formula": "realisasi/target",
                    "polarity": "Positif",
                    "weight": "10",
                    "cascading": "DIRECT",
                    "ownership_type": "SPECIFIC",
                }
            ],
        )
        issues = []

        rows, _ = build_upload_rows(config, "528", [impact], 1, issues)
        output_row = next(
            row for row in (dict(zip(UPLOAD_HEADERS, row)) for row in rows) if row["KPI Type"] == "OUTPUT"
        )

        self.assertEqual(output_row["Period"], "TRIWULANAN")
        matching = [issue for issue in issues if "enum_issue category=ambiguous; field=Period" in issue.message]
        self.assertEqual(len(matching), 1)
        self.assertEqual(matching[0].severity, "warning")

    def test_build_upload_rows_reports_cross_column_output_cascading(self):
        config = PositionConfig(
            sheet_name="Officer Kinerja Individu",
            position_name="Officer I Kinerja Individu",
            group_name="Group Pengelolaan SDM",
            directorate_name="Direktorat Sumber Daya Manusia dan Umum",
        )
        impact = ImpactRecord(
            bsc="Financial",
            title="Net Income",
            unit="Rupiah",
            period="Triwulan",
            formula="Revenue - Cost",
            polarity="Positif",
            weight="15",
            outputs=[
                {
                    "source_row": 2,
                    "title": "Penyempurnaan PMS",
                    "description": "Enhancement PMS",
                    "unit": "%",
                    "period": "Triwulan",
                    "formula": "realisasi/target",
                    "polarity": "Positif",
                    "weight": "10",
                    "cascading": "SPECIFIC",
                    "ownership_type": "SPECIFIC",
                }
            ],
        )
        issues = []

        rows, _ = build_upload_rows(config, "528", [impact], 1, issues)
        output_row = next(
            row for row in (dict(zip(UPLOAD_HEADERS, row)) for row in rows) if row["KPI Type"] == "OUTPUT"
        )

        self.assertEqual(output_row["Cascading"], "INDIRECT")
        matching = [issue for issue in issues if "enum_issue category=cross_column; field=Cascading" in issue.message]
        self.assertEqual(len(matching), 1)
        self.assertEqual(matching[0].severity, "warning")

    def test_build_upload_rows_reports_cross_column_kai_nature(self):
        config = PositionConfig(
            sheet_name="Officer Kinerja Individu",
            position_name="Officer I Kinerja Individu",
            group_name="Group Pengelolaan SDM",
            directorate_name="Direktorat Sumber Daya Manusia dan Umum",
        )
        impact = ImpactRecord(
            bsc="Financial",
            title="Net Income",
            unit="Rupiah",
            period="Triwulan",
            formula="Revenue - Cost",
            polarity="Positif",
            weight="15",
            outputs=[
                {
                    "source_row": 2,
                    "title": "Penyempurnaan PMS",
                    "description": "Enhancement PMS",
                    "unit": "%",
                    "period": "Triwulan",
                    "formula": "realisasi/target",
                    "polarity": "Positif",
                    "weight": "10",
                    "cascading": "DIRECT",
                    "ownership_type": "SPECIFIC",
                    "kai": {
                        "source_row": 2,
                        "title": "update SOP",
                        "description": "desc kai",
                        "formula": "progress x 100%",
                        "weight": "5",
                        "nature_of_work": "Pdf",
                        "period": "Tahunan",
                        "polarity": "Positif",
                    },
                }
            ],
        )
        issues = []

        rows, _ = build_upload_rows(config, "528", [impact], 1, issues)
        kai_row = next(row for row in (dict(zip(UPLOAD_HEADERS, row)) for row in rows) if row["KPI Type"] == "KAI")

        self.assertEqual(kai_row["Period"], "TAHUNAN")
        self.assertEqual(kai_row["Nature Of Work (KAI Only)"], "Non Routine")
        matching = [
            issue for issue in issues if "enum_issue category=cross_column; field=Nature Of Work" in issue.message
        ]
        self.assertEqual(len(matching), 1)
        self.assertEqual(matching[0].severity, "warning")

    def test_unknown_polarity_values_default_to_positive_enum(self):
        self.assertEqual(uploader_polarity("INDIRECT"), "POSITIVE")
        self.assertEqual(uploader_polarity("DUPLICATE"), "POSITIVE")
        self.assertEqual(uploader_polarity("SPECIFIC"), "POSITIVE")
        self.assertEqual(uploader_polarity("Negatif"), "NEGATIVE")

    def test_enum_normalizers_handle_variants_pollution_and_ambiguity(self):
        self.assertEqual(normalize_period("per tahun").value, "TAHUNAN")
        self.assertEqual(normalize_period("Semesteran").value, "SEMESTER")

        combo = normalize_period("Triwulanan/Tahunan")
        self.assertEqual(combo.status, NormalizationStatus.AMBIGUOUS)
        self.assertIsNone(combo.value)

        combo_with_fallback = normalize_period("Triwulanan/Tahunan", fallback="TRIWULANAN")
        self.assertEqual(combo_with_fallback.status, NormalizationStatus.AMBIGUOUS)
        self.assertEqual(combo_with_fallback.value, "TRIWULANAN")

        polarity = normalize_polarity("INDIRECT")
        self.assertEqual(polarity.value, "POSITIVE")
        self.assertEqual(polarity.status, NormalizationStatus.CROSS_COLUMN)
        self.assertEqual(normalize_polarity("Negatif").value, "NEGATIVE")

        cascading = normalize_cascading("SPECIFIC")
        self.assertEqual(cascading.value, "INDIRECT")
        self.assertEqual(cascading.status, NormalizationStatus.CROSS_COLUMN)
        self.assertEqual(normalize_cascading("Indirect").value, "INDIRECT")

        ownership_blank = normalize_ownership_type(None)
        self.assertEqual(ownership_blank.value, "SPECIFIC")
        self.assertEqual(ownership_blank.status, NormalizationStatus.DEFAULTED)
        ownership = normalize_ownership_type("Non Routine")
        self.assertEqual(ownership.value, "SPECIFIC")
        self.assertEqual(ownership.status, NormalizationStatus.CROSS_COLUMN)
        self.assertEqual(normalize_ownership_type("SPESIFIC").value, "SPECIFIC")

        self.assertEqual(normalize_kai_nature(None, "TAHUNAN").value, "Non Routine")
        self.assertEqual(normalize_kai_nature(None, "TRIWULANAN").value, "Routine")
        polluted_nature = normalize_kai_nature("INDIRECT", "BULANAN")
        self.assertEqual(polluted_nature.value, "Routine")
        self.assertEqual(polluted_nature.status, NormalizationStatus.CROSS_COLUMN)
        pdf_nature = normalize_kai_nature("Pdf", "TAHUNAN")
        self.assertEqual(pdf_nature.value, "Non Routine")
        self.assertEqual(pdf_nature.status, NormalizationStatus.CROSS_COLUMN)
        uploaded_nature = normalize_kai_nature("Diunggah", "BULANAN")
        self.assertEqual(uploaded_nature.value, "Routine")
        self.assertEqual(uploaded_nature.status, NormalizationStatus.CROSS_COLUMN)
        url_nature = normalize_kai_nature("https://example.com/report.pdf", "BULANAN")
        self.assertEqual(url_nature.status, NormalizationStatus.CROSS_COLUMN)
        self.assertEqual(normalize_kai_nature("Non-Rotine", "BULANAN").value, "Non Routine")

    def test_zero_weight_rows_are_dropped_and_period_inherits_from_parent(self):
        config = PositionConfig(
            sheet_name="Group Head",
            position_name="Group Head",
            group_name="Group Raw",
            directorate_name="Direktorat Raw",
            position_master_id="58",
            position_scope="structural",
        )
        rows = [
            [
                "Jenis Posisi",
                "BSC Perspective",
                "KPI Impact",
                "KPI Impact Unit",
                "KPI Impact Frequency",
                "KPI Impact Formula",
                "%Weight (Impact)",
                "KPI Impact Polarity",
                "KPI Output",
                "%Weight (Output)",
                "KPI Output Definition",
                "KPI Output Unit",
                "KPI Output Frequency",
                "KPI Output Formula",
                "KPI Output Polarity",
                "Coverage KPI Output",
                "Cascading Tagging (KPI Output)",
                "Key Activity Indicator (KAI)",
                "%Weight (Activity)",
            ],
            [
                "Struktural",
                "Financial",
                "Net Income",
                "Rupiah",
                "Triwulan",
                "Revenue - Cost",
                "100",
                "Positif",
                "Zero Output",
                "0",
                "Should be dropped",
                "%",
                None,
                "realisasi/target",
                "INDIRECT",
                "SPECIFIC",
                "DIRECT",
                "Zero KAI",
                "0",
            ],
            [
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                None,
                "Valid Output",
                "100",
                "Kept",
                "%",
                None,
                "realisasi/target",
                "DUPLICATE",
                "SPECIFIC",
                "INDIRECT",
                "Valid KAI",
                "100",
            ],
        ]
        issues = []

        impacts = parse_block_sheet(rows, config, issues)
        upload_rows, _ = build_upload_rows(config, "58", impacts, 1)
        row_maps = [dict(zip(UPLOAD_HEADERS, row)) for row in upload_rows]

        self.assertEqual([row["Title"] for row in row_maps], ["Net Income", "Valid Output", "Valid KAI"])
        self.assertEqual([row["Period"] for row in row_maps], ["TRIWULANAN", "TRIWULANAN", "TRIWULANAN"])
        self.assertEqual([row["Polarity"] for row in row_maps], ["POSITIVE", "POSITIVE", "POSITIVE"])

    def test_kai_nature_normalizes_to_backend_allowed_values(self):
        self.assertEqual(uploader_kai_nature("Routine", "Tahunan"), "Routine")
        self.assertEqual(uploader_kai_nature("routine", "Tahunan"), "Routine")
        self.assertEqual(uploader_kai_nature("Non Routine", "Semester"), "Non Routine")
        self.assertEqual(uploader_kai_nature("Non-Routine", "Semester"), "Non Routine")
        self.assertEqual(uploader_kai_nature("Non-Rotine", "Semester"), "Non Routine")
        self.assertEqual(uploader_kai_nature("(blank)", "Semester"), "Routine")
        self.assertEqual(uploader_kai_nature("Diunggah", "Tahunan"), "Non Routine")

    def test_kai_rows_force_specific_indirect_percent_unit(self):
        config = PositionConfig(
            sheet_name="Officer Kinerja Individu",
            position_name="Officer Kinerja Individu",
            group_name="Group Pengelolaan SDM",
            directorate_name="Direktorat SDM",
            position_nomenclature_id="76",
        )
        impact = ImpactRecord(
            bsc="Learning & Growth",
            title="Manpower productivity",
            unit="IDR",
            period="Triwulan",
            formula="Revenue / Employee",
            polarity="Positif",
            weight="5",
            outputs=[
                {
                    "source_row": 20,
                    "title": "Penyempurnaan PMS",
                    "description": "desc",
                    "unit": "%",
                    "period": "Triwulan",
                    "formula": "realisasi/target",
                    "polarity": "Positif",
                    "weight": "5",
                    "cascading": "DIRECT",
                    "ownership_type": "SPECIFIC",
                    "kai": {
                        "source_row": 20,
                        "title": "update SOP",
                        "description": "desc kai",
                        "formula": "progress x 100%",
                        "weight": "5",
                        "nature_of_work": "Routine",
                        "period": "Triwulan",
                        "polarity": "Positif",
                        "cascading": "SPECIFIC",
                        "ownership_type": "Non Routine",
                    },
                }
            ],
        )
        issues = []
        rows, _ = build_upload_rows(config, "528", [impact], 1, issues)

        kai_row = dict(zip(UPLOAD_HEADERS, rows[2]))
        self.assertEqual(kai_row["KPI Type"], "KAI")
        self.assertEqual(kai_row["Unit"], "%")
        self.assertEqual(kai_row["Cascading"], "INDIRECT")
        self.assertEqual(kai_row["Ownership Type"], "SPECIFIC")
        messages = [issue.message for issue in issues]
        self.assertEqual(
            sum("enum_issue category=cross_column; field=Cascading" in message for message in messages),
            1,
        )
        self.assertEqual(
            sum("enum_issue category=cross_column; field=Ownership Type" in message for message in messages),
            1,
        )

    def test_duplicate_outputs_merge_weight_and_reparent_kai_children(self):
        config = PositionConfig(
            sheet_name="Officer Kinerja Individu",
            position_name="Officer Kinerja Individu",
            group_name="Group Pengelolaan SDM",
            directorate_name="Direktorat SDM",
            position_nomenclature_id="76",
        )
        duplicated_title = "Penyempurnaan Pengelolaan Talent Management berbasis Kompetensi (manajemen kinerja individu)"
        impact = ImpactRecord(
            bsc="Learning & Growth",
            title="Manpower productivity (Revenue per total manpower)",
            unit="IDR",
            period="Triwulan",
            formula="Revenue / Employee",
            polarity="Positif",
            weight="5",
            outputs=[],
        )
        for index, weight in enumerate(["5", "5", "5", "15", "5", "5", "5"], start=1):
            impact.outputs.append(
                {
                    "source_row": 20 + index,
                    "title": duplicated_title,
                    "description": "desc",
                    "unit": "%",
                    "period": "Triwulan",
                    "formula": "realisasi/target",
                    "polarity": "Positif",
                    "weight": weight,
                    "cascading": "DIRECT",
                    "ownership_type": "SPECIFIC",
                    "kai": {
                        "source_row": 20 + index,
                        "title": f"KAI {index}",
                        "description": None,
                        "formula": "progress x 100%",
                        "weight": weight,
                        "nature_of_work": "Routine",
                        "period": "Triwulan",
                        "polarity": "Positif",
                        "cascading": "DIRECT",
                        "ownership_type": "SPECIFIC",
                    },
                }
            )

        rows, _ = build_upload_rows(config, "528", [impact], 1)

        output_rows = [dict(zip(UPLOAD_HEADERS, row)) for row in rows if row[7] == "OUTPUT"]
        kai_rows = [dict(zip(UPLOAD_HEADERS, row)) for row in rows if row[7] == "KAI"]
        self.assertEqual(len(output_rows), 1)
        self.assertEqual(output_rows[0]["Title"], duplicated_title)
        self.assertEqual(output_rows[0]["Weight (%)"], "45")
        self.assertEqual(len(kai_rows), 7)
        self.assertEqual({row["Parent KPI ID"] for row in kai_rows}, {output_rows[0]["IDKPI"]})

    def test_active_valid_sheet_requires_visible_yellow_tab(self):
        class SheetLike:
            sheet_state = "visible"

            class sheet_properties:
                class tabColor:
                    rgb = "FFFFC000"
                    indexed = None
                    theme = None

        self.assertTrue(is_active_valid_sheet(SheetLike()))

        SheetLike.sheet_state = "hidden"
        self.assertFalse(is_active_valid_sheet(SheetLike()))

        SheetLike.sheet_state = "visible"
        SheetLike.sheet_properties.tabColor.rgb = "FF70AD47"
        self.assertFalse(is_active_valid_sheet(SheetLike()))

    def test_mapping_restricts_lookup_to_target_company_and_uses_cluster_label(self):
        payload = {
            "position_master_rows": [
                {
                    "position_master_id": 111,
                    "position_name": "Officer Perencanaan",
                    "position_master_type_id": 5,
                    "company_id": 136,
                    "company_name": "PT Company Lain",
                    "group_name": "Group Lain",
                    "is_company_active": 1,
                    "is_group_active": 1,
                    "is_position_active": 1,
                    "is_position_organization_active": 1,
                },
                {
                    "position_master_id": 222,
                    "position_name": "Manager Perencanaan",
                    "position_master_type_id": 5,
                    "company_id": 1,
                    "company_name": "PT Pelabuhan Indonesia (Persero)",
                    "group_name": "Group HO",
                    "is_company_active": 1,
                    "is_group_active": 1,
                    "is_position_active": 1,
                    "is_position_organization_active": 1,
                },
            ],
            "rows": [
                {
                    "cluster_id": 333,
                    "cluster_label": "Officer Perencanaan",
                    "position_master_id": 3330,
                    "position_name": "Officer I Perencanaan",
                    "position_master_type_id": 6,
                    "type_name": "General",
                    "company_id": 1,
                    "company_name": "PT Pelabuhan Indonesia (Persero)",
                    "active_company_name": "PT Pelabuhan Indonesia (Persero)",
                    "group_name": "Group HO",
                    "active_group_name": "Group HO",
                    "is_company_active": 1,
                    "is_group_active": 1,
                }
            ],
        }
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "reference.json"
            path.write_text(json.dumps(payload), encoding="utf-8")

            mapping = load_nomenclature_mapping(path, "1")

        self.assertNotIn("officer perencanaan", {k for k, v in mapping.items() if v.get("portaverse_company_name") == "PT Company Lain"})
        self.assertEqual(mapping["officer perencanaan"]["position_nomenclature_id"], "333")
        self.assertEqual(mapping["officer perencanaan"]["position_scope"], "non_structural")
        self.assertEqual(mapping["manager perencanaan"]["position_master_id"], "222")
        self.assertIsNone(mapping["manager perencanaan"]["position_nomenclature_id"])

    def test_refresh_config_from_mapping_clears_old_other_company_id(self):
        mapping = {
            "officer perencanaan": {
                "position_master_id": "3330",
                "position_nomenclature_id": "333",
                "position_scope": "non_structural",
                "portaverse_position_title": "Officer I Perencanaan",
                "portaverse_group_name": "Group HO",
                "portaverse_company_name": "PT Pelabuhan Indonesia (Persero)",
                "cluster_label": "Officer Perencanaan",
            }
        }
        config = PositionConfig(
            sheet_name="Officer Perencanaan",
            position_name="Officer Perencanaan",
            group_name="Group Raw",
            directorate_name="Direktorat Raw",
            position_master_id="111",
            position_scope="structural",
            portaverse_company_name="PT Company Lain",
        )

        refresh_configs_from_mapping([config], mapping)

        self.assertIsNone(config.position_master_id)
        self.assertEqual(config.position_nomenclature_id, "333")
        self.assertEqual(config.position_scope, "non_structural")
        self.assertEqual(config.group_name, "Group Raw")

    def test_refresh_config_from_mapping_uses_structural_scope_when_number_exists_as_pmid_and_pnid(self):
        mapping = {
            "manager rekrutmen karir": {
                "position_master_id": "515",
                "position_nomenclature_id": "515",
                "position_scope": "non_structural",
                "position_master_type_id": "5",
                "portaverse_position_title": "Manager Rekrutmen dan Karir",
                "portaverse_group_name": "Unit Pendukung Rekrutmen dan Karir",
                "portaverse_company_name": "PT Pelabuhan Indonesia (Persero)",
                "cluster_label": "Manager Rekrutmen-Karir",
            }
        }
        config = PositionConfig(
            sheet_name="Manager Rekrutmen-Karir",
            position_name="Manager Rekrutmen-Karir",
            group_name="Group Pengelolaan SDM",
            directorate_name="Direktorat SDM & Umum",
            position_nomenclature_id="515",
            position_scope="non_structural",
        )

        refresh_configs_from_mapping([config], mapping)

        self.assertEqual(config.position_master_id, "515")
        self.assertIsNone(config.position_nomenclature_id)
        self.assertEqual(config.position_scope, "structural")

    def test_load_nomenclature_mapping_structural_title_wins_over_same_number_pnid(self):
        payload = {
            "position_master_rows": [
                {
                    "position_master_id": 515,
                    "position_name": "Manager Rekrutmen dan Karir",
                    "position_master_type_id": 5,
                    "company_id": 1,
                    "company_name": "PT Pelabuhan Indonesia (Persero)",
                    "group_name": "Unit Pendukung Rekrutmen dan Karir",
                    "is_company_active": 1,
                    "is_group_active": 1,
                    "is_position_active": 1,
                    "is_position_organization_active": 1,
                }
            ],
            "rows": [
                {
                    "cluster_id": 515,
                    "cluster_label": "Manager Rekrutmen dan Karir",
                    "position_master_id": 9999,
                    "position_name": "Officer I Rekrutmen dan Karir",
                    "position_master_type_id": 6,
                    "type_name": "General",
                    "company_id": 1,
                    "company_name": "PT Pelabuhan Indonesia (Persero)",
                    "active_company_name": "PT Pelabuhan Indonesia (Persero)",
                    "group_name": "Unit Pendukung Rekrutmen dan Karir",
                    "active_group_name": "Unit Pendukung Rekrutmen dan Karir",
                    "is_company_active": 1,
                    "is_group_active": 1,
                }
            ],
        }
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "reference.json"
            path.write_text(json.dumps(payload), encoding="utf-8")

            mapping = load_nomenclature_mapping(path, "1")

        manager = mapping["manager rekrutmen dan karir"]
        self.assertEqual(manager["position_master_id"], "515")
        self.assertIsNone(manager["position_nomenclature_id"])
        self.assertEqual(manager["position_scope"], "structural")

    def test_refresh_config_from_mapping_preserves_reviewed_manual_pmid(self):
        mapping = {
            "different manager title": {
                "position_master_id": "999",
                "position_nomenclature_id": None,
                "position_scope": "structural",
                "portaverse_position_title": "Different Manager Title",
                "portaverse_group_name": "Group HO",
                "portaverse_company_name": "PT Pelabuhan Indonesia (Persero)",
                "cluster_label": None,
            }
        }
        config = PositionConfig(
            sheet_name="Manager Performa Keuangan",
            position_name="Manager Performa Keuangan",
            group_name="Group Raw",
            directorate_name="Direktorat Raw",
            position_master_id="316",
            position_scope="structural",
        )

        refresh_configs_from_mapping([config], mapping)

        self.assertEqual(config.position_master_id, "316")
        self.assertIsNone(config.position_nomenclature_id)
        self.assertEqual(config.position_scope, "structural")

    def test_refresh_config_from_mapping_preserves_reviewed_manual_pnid(self):
        mapping = {
            "different officer title": {
                "position_master_id": None,
                "position_nomenclature_id": "999",
                "position_scope": "non_structural",
                "portaverse_position_title": "Different Officer Title",
                "portaverse_group_name": "Group HO",
                "portaverse_company_name": "PT Pelabuhan Indonesia (Persero)",
                "cluster_label": "Different Officer Title",
            }
        }
        config = PositionConfig(
            sheet_name="Officer QA",
            position_name="Officer QA",
            group_name="Group Raw",
            directorate_name="Direktorat Raw",
            position_nomenclature_id="11517",
            position_scope="non_structural",
        )

        refresh_configs_from_mapping([config], mapping)

        self.assertIsNone(config.position_master_id)
        self.assertEqual(config.position_nomenclature_id, "11517")
        self.assertEqual(config.position_scope, "non_structural")

    def test_write_output_workbook_saves_weight_as_numeric_cell(self):
        with tempfile.TemporaryDirectory() as tmp:
            template = Path(tmp) / "template.xlsx"
            output = Path(tmp) / "output.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "KPI Template"
            ws.append(UPLOAD_HEADERS)
            ws.append([None] * len(UPLOAD_HEADERS))
            wb.save(template)
            row = [
                "1",
                "Group Raw",
                "Direktorat Raw",
                "Officer",
                None,
                None,
                "Financial",
                "IMPACT",
                None,
                "#N/A",
                "Title",
                None,
                "%",
                "POSITIVE",
                "TRIWULANAN",
                "x/y",
                "12.5",
                None,
                None,
                None,
                None,
                None,
                "333",
                None,
            ]

            write_output_workbook(template, output, [row])
            saved = load_workbook(output, data_only=True)
            weight_cell = saved["KPI Template"].cell(row=2, column=17)

        self.assertIsInstance(weight_cell.value, float)
        self.assertEqual(weight_cell.value, 12.5)
        self.assertEqual(saved["KPI Template"].column_dimensions["P"].width, 72)


if __name__ == "__main__":
    unittest.main()
