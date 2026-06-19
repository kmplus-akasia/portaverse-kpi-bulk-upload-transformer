from __future__ import annotations

import unittest

import pandas as pd
from streamlit.testing.v1 import AppTest

from dashboard import kpi_planning_dashboard as dashboard


def dashboard_fixture_data():
    from dashboard import kpi_planning_dashboard as app

    worker_progress = app.enrich_worker_progress(
        pd.DataFrame(
            [
                {
                    "company_id": 1,
                    "company_name": "PT Pelabuhan Indonesia (Persero)",
                    "employee_number": "100",
                    "employee_name": "Ayu",
                    "corporate_email": "ayu@example.com",
                    "position_master_variant_id": 10,
                    "position_master_id": 20,
                    "position_name": "Manager Operasi",
                    "position_type_name": "Struktural",
                    "position_master_type_id": 5,
                    "group_name": "Operasi",
                    "assignment_type": "DEFINITIF",
                    "pmid": 20,
                    "pnid": None,
                    "pnid_label": None,
                    "pnid_mapping_count": 0,
                    "system_kpi_count": 1,
                    "manual_kpi_count": 0,
                    "unknown_origin_kpi_count": 0,
                    "total_kpi_count": 1,
                    "impact_count": 1,
                    "output_count": 0,
                    "kai_count": 0,
                    "total_weight": 100,
                    "draft_status_count": 1,
                    "subordinate_review_count": 0,
                    "manager_decision_count": 0,
                    "approved_count": 0,
                    "last_kpi_update": "2026-06-19 10:00:00",
                }
            ]
        )
    )
    return {
        "db_info": pd.DataFrame([{"db_name": "fixture", "host_name": "test", "server_read_only": 1, "db_now": "2026-06-19"}]),
        "schedule": pd.DataFrame([{"name": "Planning 2026", "start_date": "2026-01-01", "end_date": "2026-12-31"}]),
        "company_options": pd.DataFrame(
            [
                {"company_id": 1, "company_name": "PT Pelabuhan Indonesia (Persero)"},
                {"company_id": 22, "company_name": "Cabang Pelabuhan Bengkulu"},
            ]
        ),
        "worker_progress": worker_progress,
        "coverage": pd.DataFrame([{
            "active_position_variants": 1,
            "active_position_masters": 1,
            "variants_with_any_kpi": 1,
            "variants_without_kpi": 0,
            "total_kpi_ownership_rows": 1,
        }]),
        "category_summary": pd.DataFrame([
            {"category": "Struktural", "unit": "PMID", "active_units": 1, "complete_units": 1, "partial_units": 0, "missing_units": 0, "not_complete_units": 0, "coverage_pct": 100.0},
            {"category": "Non-struktural", "unit": "PNID", "active_units": 0, "complete_units": 0, "partial_units": 0, "missing_units": 0, "not_complete_units": 0, "coverage_pct": 0.0},
        ]),
        "category_anomalies": pd.DataFrame([{"anomaly": "active_non_structural_without_pnid", "active_pmids": 0, "active_variants": 0}]),
        "structural_detail": app.empty_frame(["with_kpi_variants"]),
        "non_structural_detail": app.empty_frame(["with_kpi_variants"]),
        "structural_not_complete": app.empty_frame(["pmid"]),
        "non_structural_not_complete": app.empty_frame(["pnid"]),
        "category_gap_detail": app.empty_frame(["category"]),
        "group_coverage": app.empty_frame(["company_id", "company_name", "group_name", "active_variants", "active_primary_variants", "with_kpi", "without_kpi", "coverage_pct", "impact_count", "output_count", "kai_count"]),
        "status": app.empty_frame(["type", "item_approval_status", "allocation_status", "weight_approval_status", "ownership_rows", "kpi_ids", "active_variants"]),
        "import_agg": app.empty_frame(["dry_run", "status", "logs", "total_rows", "created_count", "updated_count", "affected_rows", "invalid_count", "total_positions_sum", "first_log_at", "last_log_at"]),
        "success_by_file": app.empty_frame(["file_name"]),
        "failed_imports": app.empty_frame(["import_log_id"]),
        "anomalies": pd.DataFrame([{"anomaly": "fixture", "cnt": 0}]),
        "without_kpi": app.empty_frame(["company_id", "company_name", "group_name"]),
    }


def fixture_streamlit_app():
    from dashboard import kpi_planning_dashboard as app
    from tests.test_kpi_planning_dashboard import dashboard_fixture_data

    data = dashboard_fixture_data()

    app.run_app(
        data_loader=lambda year, company_id: data,
        company_loader=lambda: data["company_options"],
    )


def unavailable_streamlit_app():
    from dashboard import kpi_planning_dashboard as app

    def unavailable():
        raise RuntimeError("Missing database environment variables: DB_HOST")

    app.run_app(company_loader=unavailable)


class KpiPlanningDashboardTest(unittest.TestCase):
    def test_streamlit_app_renders_all_company_progress_and_followup(self):
        app = AppTest.from_function(fixture_streamlit_app, default_timeout=10).run()

        self.assertEqual(len(app.exception), 0)
        self.assertIn("KPI Planning Dashboard Production", [item.value for item in app.title])
        self.assertIn("Company ID - Name", [item.label for item in app.selectbox])
        self.assertIn(
            "Progress perencanaan KPI per pekerja",
            [item.value for item in app.subheader],
        )
        self.assertIn("Daftar pekerja untuk follow-up", [item.value for item in app.subheader])
        self.assertTrue(any(item.label == "Unduh CSV follow-up" for item in app.get("download_button")))

    def test_streamlit_app_stops_without_cached_metrics_when_database_is_unavailable(self):
        app = AppTest.from_function(unavailable_streamlit_app, default_timeout=10).run()

        self.assertEqual(len(app.exception), 0)
        self.assertTrue(any("Unable to load live production data" in item.value for item in app.error))
        self.assertTrue(any("Dashboard unavailable" in item.value for item in app.caption))
        self.assertEqual(len(app.metric), 0)

    def test_check_payload_includes_all_company_worker_progress(self):
        data = dashboard_fixture_data()

        payload = dashboard.build_check_payload(data, 2026, None)

        self.assertEqual(payload["company_scope"], "ALL_PELINDO")
        self.assertEqual(payload["worker_progress"]["active_workers"], 1)
        self.assertEqual(payload["portfolio_origin_summary"], {"KAMUS_KPI": 1})
        self.assertEqual(payload["worker_readiness_summary"], {"DRAFT_PERENCANAAN": 1})

    def test_parse_company_id_defaults_to_all(self):
        self.assertIsNone(dashboard.parse_company_id(None))
        self.assertIsNone(dashboard.parse_company_id(""))
        self.assertIsNone(dashboard.parse_company_id("all"))
        self.assertEqual(dashboard.parse_company_id(" 22 "), 22)
        with self.assertRaisesRegex(ValueError, "Company ID"):
            dashboard.parse_company_id("not-a-number")

    def test_sql_scope_uses_parameter_only_for_selected_company(self):
        all_scope = dashboard.build_active_cte(None)
        selected_scope = dashboard.build_active_cte(22)

        self.assertNotIn("tgm.company_id = %(company_id)s", all_scope)
        self.assertIn("tgm.company_id = %(company_id)s", selected_scope)
        self.assertIn("te.archived_at IS NULL", all_scope)
        self.assertIn("tgm.company_id AS company_id", all_scope)
        self.assertIn("k.created_by_pov = 'SYSTEM'", all_scope)

    def test_worker_progress_sql_aggregates_origin_and_current_worker_portfolio(self):
        sql = dashboard.build_worker_progress_sql(None)

        self.assertIn("active_worker_positions AS", sql)
        self.assertIn("position_origin AS", sql)
        self.assertIn("employee_portfolio AS", sql)
        self.assertIn("k.created_by_pov = 'SYSTEM'", sql)
        self.assertIn("k.created_by_pov IN ('WORKER', 'SUPERIOR')", sql)
        self.assertIn("ep.employee_number = awp.employee_number", sql)
        position_origin_sql = sql.split("position_origin AS", 1)[1].split(
            "employee_portfolio AS", 1
        )[0]
        self.assertIn("GROUP BY ko.position_master_id", position_origin_sql)
        self.assertNotIn(
            "GROUP BY ko.position_master_variant_id", position_origin_sql
        )
        self.assertIn("GROUP_CONCAT(DISTINCT CASE", sql)
        self.assertIn("GROUP BY", sql.split("position_origin AS", 1)[0])
        self.assertNotIn("tgm.company_id = %(company_id)s", sql)

        scoped = dashboard.build_worker_progress_sql(22)
        self.assertIn("tgm.company_id = %(company_id)s", scoped)

    def test_origin_prefers_dictionary_and_flags_unknown_values(self):
        self.assertEqual(dashboard.classify_portfolio_origin(2, 3, 0), "KAMUS_KPI")
        self.assertEqual(
            dashboard.classify_portfolio_origin(0, 3, 0),
            "MANUAL_TANPA_KAMUS",
        )
        self.assertEqual(
            dashboard.classify_portfolio_origin(0, 0, 0),
            "BELUM_ADA_PORTFOLIO",
        )
        self.assertEqual(
            dashboard.classify_portfolio_origin(0, 0, 1),
            "ORIGIN_TIDAK_DIKENAL",
        )

    def test_readiness_uses_official_priority(self):
        self.assertEqual(
            dashboard.classify_portfolio_readiness(0, 0, 0, 0, 0),
            "BELUM_ADA_DRAFT",
        )
        self.assertEqual(
            dashboard.classify_portfolio_readiness(4, 1, 1, 1, 1),
            "DRAFT_PERENCANAAN",
        )
        self.assertEqual(
            dashboard.classify_portfolio_readiness(3, 0, 1, 1, 1),
            "MENUNGGU_REVIEW_BAWAHAN",
        )
        self.assertEqual(
            dashboard.classify_portfolio_readiness(2, 0, 0, 1, 1),
            "MENUNGGU_KEPUTUSAN_ANDA",
        )
        self.assertEqual(
            dashboard.classify_portfolio_readiness(2, 0, 0, 0, 2),
            "DISETUJUI",
        )
        self.assertEqual(
            dashboard.classify_portfolio_readiness(2, 0, 0, 0, 1),
            "DRAFT_PERENCANAAN",
        )

    def test_enrich_progress_and_roll_up_least_advanced_assignment(self):
        raw = pd.DataFrame(
            [
                self.row("100", 1, system=2, total=2, approved=2),
                self.row("100", 2, manual=1, total=1, draft=1),
                self.row("200", 3),
            ]
        )

        detail = dashboard.enrich_worker_progress(raw)
        workers = dashboard.build_worker_level_progress(detail).set_index(
            "employee_number"
        )

        self.assertEqual(workers.loc["100", "readiness_status"], "DRAFT_PERENCANAAN")
        self.assertEqual(workers.loc["100", "active_assignment_count"], 2)
        self.assertEqual(workers.loc["200", "readiness_status"], "BELUM_ADA_DRAFT")
        self.assertEqual(detail.loc[0, "portfolio_origin"], "KAMUS_KPI")
        self.assertEqual(detail.loc[1, "portfolio_origin"], "MANUAL_TANPA_KAMUS")

    def test_followup_filter_defaults_to_non_approved_and_searches_worker_or_group(self):
        detail = dashboard.enrich_worker_progress(
            pd.DataFrame(
                [
                    self.row("100", 1, name="Ayu", group="Operasi", total=1, approved=1),
                    self.row("200", 2, name="Bima", group="Keuangan"),
                ]
            )
        )

        followup = dashboard.filter_worker_progress(detail)
        approved = dashboard.filter_worker_progress(detail, include_approved=True)
        searched = dashboard.filter_worker_progress(
            detail, include_approved=True, search="operasi"
        )

        self.assertEqual(followup["employee_number"].tolist(), ["200"])
        self.assertEqual(len(approved), 2)
        self.assertEqual(searched["employee_number"].tolist(), ["100"])

    def test_summary_counts_unique_workers_assignments_and_followup(self):
        detail = dashboard.enrich_worker_progress(
            pd.DataFrame(
                [
                    self.row("100", 1, system=1, total=1, approved=1),
                    self.row("100", 2, manual=1, total=1, draft=1),
                    self.row("200", 3),
                ]
            )
        )

        summary = dashboard.build_progress_summary(detail)

        self.assertEqual(summary["active_workers"], 2)
        self.assertEqual(summary["active_assignments"], 3)
        self.assertEqual(summary["followup_workers"], 2)
        self.assertEqual(summary["dictionary_assignments"], 1)
        self.assertEqual(summary["manual_assignments"], 1)

    def test_csv_uses_utf8_bom_and_only_selected_columns(self):
        detail = dashboard.enrich_worker_progress(pd.DataFrame([self.row("100", 1)]))

        payload = dashboard.worker_progress_csv(detail)

        self.assertTrue(payload.startswith(b"\xef\xbb\xbf"))
        text = payload.decode("utf-8-sig")
        self.assertIn("employee_number", text)
        self.assertIn("portfolio_origin", text)
        self.assertNotIn("draft_status_count", text)

    @staticmethod
    def row(
        employee_number: str,
        variant_id: int,
        *,
        name: str = "Pekerja",
        group: str = "Group",
        system: int = 0,
        manual: int = 0,
        unknown: int = 0,
        total: int = 0,
        draft: int = 0,
        review: int = 0,
        decision: int = 0,
        approved: int = 0,
    ) -> dict[str, object]:
        return {
            "company_id": 1,
            "company_name": "PT Pelabuhan Indonesia (Persero)",
            "employee_number": employee_number,
            "employee_name": name,
            "corporate_email": f"{employee_number}@example.com",
            "position_master_variant_id": variant_id,
            "position_master_id": variant_id + 100,
            "position_name": f"Position {variant_id}",
            "position_type_name": "Struktural",
            "position_master_type_id": 5,
            "group_name": group,
            "assignment_type": "DEFINITIF",
            "pmid": variant_id + 100,
            "pnid": None,
            "pnid_label": None,
            "pnid_mapping_count": 0,
            "system_kpi_count": system,
            "manual_kpi_count": manual,
            "unknown_origin_kpi_count": unknown,
            "total_kpi_count": total,
            "impact_count": 0,
            "output_count": 0,
            "kai_count": 0,
            "total_weight": 0,
            "draft_status_count": draft,
            "subordinate_review_count": review,
            "manager_decision_count": decision,
            "approved_count": approved,
            "last_kpi_update": None,
        }


if __name__ == "__main__":
    unittest.main()
