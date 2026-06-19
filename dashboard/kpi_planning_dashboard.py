from __future__ import annotations

import argparse
import json
import os
from dataclasses import dataclass
from typing import Any

import pandas as pd
import pymysql
from pymysql.cursors import DictCursor
from pymysql.err import OperationalError


DEFAULT_YEAR = int(os.getenv("KPI_DASHBOARD_YEAR", "2026"))

PORTFOLIO_ORIGINS = [
    "KAMUS_KPI",
    "MANUAL_TANPA_KAMUS",
    "BELUM_ADA_PORTFOLIO",
    "ORIGIN_TIDAK_DIKENAL",
]
READINESS_ORDER = [
    "BELUM_ADA_DRAFT",
    "DRAFT_PERENCANAAN",
    "MENUNGGU_REVIEW_BAWAHAN",
    "MENUNGGU_KEPUTUSAN_ANDA",
    "DISETUJUI",
]
READINESS_LABELS = {
    "BELUM_ADA_DRAFT": "Belum Ada Draft",
    "DRAFT_PERENCANAAN": "Draft Perencanaan",
    "MENUNGGU_REVIEW_BAWAHAN": "Menunggu Review Bawahan",
    "MENUNGGU_KEPUTUSAN_ANDA": "Menunggu Keputusan Anda",
    "DISETUJUI": "Disetujui",
}
ORIGIN_LABELS = {
    "KAMUS_KPI": "Kamus KPI / Performance Tree",
    "MANUAL_TANPA_KAMUS": "Manual tanpa kamus",
    "BELUM_ADA_PORTFOLIO": "Belum ada portfolio",
    "ORIGIN_TIDAK_DIKENAL": "Origin tidak dikenal",
}


def parse_company_id(value: Any) -> int | None:
    if value is None or str(value).strip().lower() in {"", "all"}:
        return None
    try:
        company_id = int(str(value).strip())
    except (TypeError, ValueError) as exc:
        raise ValueError("Company ID must be a positive integer or 'all'.") from exc
    if company_id <= 0:
        raise ValueError("Company ID must be a positive integer or 'all'.")
    return company_id


DEFAULT_COMPANY_ID = parse_company_id(os.getenv("KPI_DASHBOARD_COMPANY_ID"))


def classify_portfolio_origin(
    system_kpi_count: Any,
    manual_kpi_count: Any,
    unknown_origin_kpi_count: Any,
) -> str:
    if int(system_kpi_count or 0) > 0:
        return "KAMUS_KPI"
    if int(manual_kpi_count or 0) > 0:
        return "MANUAL_TANPA_KAMUS"
    if int(unknown_origin_kpi_count or 0) > 0:
        return "ORIGIN_TIDAK_DIKENAL"
    return "BELUM_ADA_PORTFOLIO"


def classify_portfolio_readiness(
    total_kpi_count: Any,
    draft_status_count: Any,
    subordinate_review_count: Any,
    manager_decision_count: Any,
    approved_count: Any,
) -> str:
    total = int(total_kpi_count or 0)
    if total == 0:
        return "BELUM_ADA_DRAFT"
    if int(draft_status_count or 0) > 0:
        return "DRAFT_PERENCANAAN"
    if int(subordinate_review_count or 0) > 0:
        return "MENUNGGU_REVIEW_BAWAHAN"
    if int(manager_decision_count or 0) > 0:
        return "MENUNGGU_KEPUTUSAN_ANDA"
    if int(approved_count or 0) == total:
        return "DISETUJUI"
    return "DRAFT_PERENCANAAN"


def enrich_worker_progress(df: pd.DataFrame) -> pd.DataFrame:
    enriched = df.copy()
    if enriched.empty:
        for column in [
            "portfolio_origin",
            "portfolio_origin_label",
            "readiness_status",
            "readiness_label",
            "mapping_anomaly",
        ]:
            enriched[column] = pd.Series(dtype="object")
        return enriched

    numeric_columns = [
        "system_kpi_count",
        "manual_kpi_count",
        "unknown_origin_kpi_count",
        "total_kpi_count",
        "draft_status_count",
        "subordinate_review_count",
        "manager_decision_count",
        "approved_count",
    ]
    for column in numeric_columns:
        enriched[column] = pd.to_numeric(enriched[column], errors="coerce").fillna(0)

    enriched["portfolio_origin"] = enriched.apply(
        lambda row: classify_portfolio_origin(
            row["system_kpi_count"],
            row["manual_kpi_count"],
            row["unknown_origin_kpi_count"],
        ),
        axis=1,
    )
    enriched["readiness_status"] = enriched.apply(
        lambda row: classify_portfolio_readiness(
            row["total_kpi_count"],
            row["draft_status_count"],
            row["subordinate_review_count"],
            row["manager_decision_count"],
            row["approved_count"],
        ),
        axis=1,
    )
    enriched["portfolio_origin_label"] = enriched["portfolio_origin"].map(
        ORIGIN_LABELS
    )
    enriched["readiness_label"] = enriched["readiness_status"].map(
        READINESS_LABELS
    )
    mapping_count = pd.to_numeric(
        enriched.get("pnid_mapping_count", 0), errors="coerce"
    ).fillna(0)
    structural = pd.to_numeric(
        enriched.get("position_master_type_id", 0), errors="coerce"
    ).fillna(0) == 5
    enriched["mapping_anomaly"] = ""
    enriched.loc[~structural & (mapping_count == 0), "mapping_anomaly"] = "PNID_MISSING"
    enriched.loc[~structural & (mapping_count > 1), "mapping_anomaly"] = "PNID_MULTIPLE"
    return enriched


def build_worker_level_progress(detail: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "employee_number",
        "employee_name",
        "active_assignment_count",
        "readiness_status",
        "readiness_label",
    ]
    if detail.empty:
        return empty_frame(columns)

    ranked = detail.copy()
    ranked["_readiness_rank"] = ranked["readiness_status"].map(
        {status: rank for rank, status in enumerate(READINESS_ORDER)}
    ).fillna(1)
    worker = (
        ranked.sort_values(["employee_number", "_readiness_rank"])
        .groupby("employee_number", as_index=False)
        .agg(
            employee_name=("employee_name", "first"),
            active_assignment_count=("position_master_variant_id", "size"),
            _readiness_rank=("_readiness_rank", "min"),
        )
    )
    worker["readiness_status"] = worker["_readiness_rank"].map(
        {rank: status for rank, status in enumerate(READINESS_ORDER)}
    )
    worker["readiness_label"] = worker["readiness_status"].map(READINESS_LABELS)
    return worker[columns]


def filter_worker_progress(
    detail: pd.DataFrame,
    *,
    search: str = "",
    origins: list[str] | None = None,
    statuses: list[str] | None = None,
    include_approved: bool = False,
) -> pd.DataFrame:
    filtered = detail.copy()
    if not include_approved and "readiness_status" in filtered:
        filtered = filtered[filtered["readiness_status"] != "DISETUJUI"]
    if origins:
        filtered = filtered[filtered["portfolio_origin"].isin(origins)]
    if statuses:
        filtered = filtered[filtered["readiness_status"].isin(statuses)]
    token = search.strip()
    if token and not filtered.empty:
        searchable = [
            column
            for column in [
                "employee_number",
                "employee_name",
                "corporate_email",
                "group_name",
                "position_name",
            ]
            if column in filtered
        ]
        mask = pd.Series(False, index=filtered.index)
        for column in searchable:
            mask |= filtered[column].astype(str).str.contains(
                token, case=False, na=False, regex=False
            )
        filtered = filtered[mask]
    return filtered.reset_index(drop=True)


def build_progress_summary(detail: pd.DataFrame) -> dict[str, int]:
    workers = build_worker_level_progress(detail)
    return {
        "active_workers": int(detail["employee_number"].nunique()) if not detail.empty else 0,
        "active_assignments": int(len(detail)),
        "followup_workers": int(
            (workers["readiness_status"] != "DISETUJUI").sum()
        ) if not workers.empty else 0,
        "dictionary_assignments": int(
            (detail["portfolio_origin"] == "KAMUS_KPI").sum()
        ) if not detail.empty else 0,
        "manual_assignments": int(
            (detail["portfolio_origin"] == "MANUAL_TANPA_KAMUS").sum()
        ) if not detail.empty else 0,
        "approved_workers": int(
            (workers["readiness_status"] == "DISETUJUI").sum()
        ) if not workers.empty else 0,
    }


WORKER_PROGRESS_EXPORT_COLUMNS = [
    "company_id",
    "company_name",
    "employee_number",
    "employee_name",
    "corporate_email",
    "assignment_type",
    "group_name",
    "position_name",
    "pmid",
    "pnid",
    "pnid_label",
    "mapping_anomaly",
    "portfolio_origin",
    "portfolio_origin_label",
    "readiness_status",
    "readiness_label",
    "total_kpi_count",
    "impact_count",
    "output_count",
    "kai_count",
    "total_weight",
    "last_kpi_update",
]


def worker_progress_csv(detail: pd.DataFrame) -> bytes:
    columns = [column for column in WORKER_PROGRESS_EXPORT_COLUMNS if column in detail]
    return detail[columns].to_csv(index=False).encode("utf-8-sig")


@dataclass(frozen=True)
class DbConfig:
    host: str
    port: int
    database: str
    user: str
    password: str
    ssl_enabled: bool = False


def _env(name: str, fallback: str | None = None) -> str | None:
    return os.getenv(f"KPI_DASHBOARD_{name}") or os.getenv(name) or fallback


def load_db_config() -> DbConfig:
    missing = [
        name
        for name in ["DB_HOST", "DB_PORT", "DB_NAME", "DB_USER", "DB_PASSWORD"]
        if not _env(name)
    ]
    if missing:
        raise RuntimeError(
            "Missing database environment variables: " + ", ".join(missing)
        )
    return DbConfig(
        host=str(_env("DB_HOST")),
        port=int(str(_env("DB_PORT"))),
        database=str(_env("DB_NAME")),
        user=str(_env("DB_USER")),
        password=str(_env("DB_PASSWORD")),
        ssl_enabled=str(_env("DB_SSL", "0")) == "1",
    )


def connect(config: DbConfig):
    return pymysql.connect(
        host=config.host,
        port=config.port,
        user=config.user,
        password=config.password,
        database=config.database,
        cursorclass=DictCursor,
        autocommit=True,
        connect_timeout=10,
        read_timeout=30,
        write_timeout=30,
        ssl={} if config.ssl_enabled else None,
    )


def read_sql(conn, sql: str, params: dict[str, Any] | None = None) -> pd.DataFrame:
    normalized = sql.strip().lower()
    if not (normalized.startswith("select") or normalized.startswith("with")):
        raise ValueError("Dashboard query must be read-only SELECT/WITH SQL.")
    with conn.cursor() as cursor:
        cursor.execute(sql, params or {})
        return pd.DataFrame(cursor.fetchall())


ACTIVE_CTE_TEMPLATE = """
active_positions AS (
  SELECT
    tpmv.position_master_variant_id,
    tpm.position_master_id,
    tpm.name AS position_name,
    tpm.position_master_type_id,
    tpmt.name AS position_type_name,
    tpm.cohort_id,
    tpmos.organization_master_id AS group_master_id,
    tgm.name AS group_name,
    tgm.company_id AS company_id,
    tci.name AS company_name,
    COUNT(DISTINCT tepms.employee_number) AS active_employee_count,
    SUM(CASE WHEN tepms.lakhar_id IS NULL AND tepms.job_sharing_id IS NULL THEN 1 ELSE 0 END) AS primary_assignment_count,
    SUM(CASE WHEN tepms.lakhar_id IS NOT NULL OR tepms.job_sharing_id IS NOT NULL THEN 1 ELSE 0 END) AS secondary_assignment_count
  FROM tb_employee_position_master_sync tepms
  JOIN tb_employee te
    ON te.employee_number = tepms.employee_number
   AND te.deletedAt IS NULL
   AND te.archived_at IS NULL
  JOIN tb_position_master_variant tpmv
    ON tpmv.position_master_variant_id = tepms.position_master_variant_id
   AND tpmv.deletedAt IS NULL
  JOIN tb_position_master_v2 tpm
    ON tpm.position_master_id = tpmv.position_master_id
   AND tpm.deletedAt IS NULL
   AND CURRENT_TIMESTAMP() BETWEEN COALESCE(tpm.start_date, '1000-01-01') AND COALESCE(tpm.end_date, '9999-12-31')
  LEFT JOIN tb_position_master_type tpmt
    ON tpmt.position_master_type_id = tpm.position_master_type_id
  JOIN tb_position_master_organization_sync tpmos
    ON tpmos.position_master_id = tpm.position_master_id
   AND tpmos.deletedAt IS NULL
   AND CURRENT_TIMESTAMP() BETWEEN COALESCE(tpmos.start_date, '1000-01-01') AND COALESCE(tpmos.end_date, '9999-12-31')
  JOIN tb_group_master tgm
    ON tgm.group_master_id = tpmos.organization_master_id
   AND tgm.deletedAt IS NULL
   {company_filter}
   AND CURRENT_TIMESTAMP() BETWEEN COALESCE(tgm.start_date, '1000-01-01') AND COALESCE(tgm.end_date, '9999-12-31')
  JOIN tb_company_in tci
    ON tci.company_in_id = tgm.company_id
   AND tci.deletedAt IS NULL
   AND CURRENT_TIMESTAMP() BETWEEN COALESCE(tci.start_date, '1000-01-01') AND COALESCE(tci.end_date, '9999-12-31')
  WHERE tepms.deletedAt IS NULL
    AND CURRENT_TIMESTAMP() BETWEEN tepms.start_date AND COALESCE(tepms.end_date, '9999-12-31')
    AND tpm.name NOT REGEXP '^(JA_|JS_)'
  GROUP BY
    tpmv.position_master_variant_id,
    tpm.position_master_id,
    tpm.name,
    tpm.position_master_type_id,
    tpmt.name,
    tpm.cohort_id,
    tpmos.organization_master_id,
    tgm.name,
    tgm.company_id,
    tci.name
),
kpi_counts AS (
  SELECT
    ko.position_master_variant_id,
    ko.position_master_id,
    COUNT(DISTINCT ko.kpi_id) AS kpi_count,
    COUNT(DISTINCT CASE WHEN k.type = 'IMPACT' THEN ko.kpi_id END) AS impact_count,
    COUNT(DISTINCT CASE WHEN k.type = 'OUTPUT' THEN ko.kpi_id END) AS output_count,
    COUNT(DISTINCT CASE WHEN k.type = 'KAI' THEN ko.kpi_id END) AS kai_count,
    COUNT(DISTINCT CASE WHEN k.type = 'SUB_IMPACT' THEN ko.kpi_id END) AS sub_impact_count,
    COUNT(DISTINCT CASE WHEN k.item_approval_status IN ('APPROVED', 'APPROVED_ADJUSTED') THEN ko.kpi_id END) AS approved_item_count,
    COUNT(DISTINCT CASE WHEN ko.weight_approval_status = 'APPROVED' THEN ko.kpi_id END) AS approved_weight_count
  FROM kpi_ownership_v3 ko
  JOIN kpi_v3 k ON k.kpi_id = ko.kpi_id
  WHERE ko.deleted_at IS NULL
    AND k.deleted_at IS NULL
    AND k.is_active = 1
    AND COALESCE(k.status, 'ACTIVE') = 'ACTIVE'
    AND k.created_by_pov = 'SYSTEM'
    AND ko.year = %(year)s
    AND k.year = %(year)s
  GROUP BY ko.position_master_variant_id, ko.position_master_id
)
"""


def _company_filter(company_id: int | None) -> str:
    return "AND tgm.company_id = %(company_id)s" if company_id is not None else ""


def build_active_cte(company_id: int | None) -> str:
    return ACTIVE_CTE_TEMPLATE.format(company_filter=_company_filter(company_id))


def build_worker_progress_sql(company_id: int | None) -> str:
    return f"""
    WITH
    pnid_mapping AS (
      SELECT
        pnm.company_id,
        pnm.group_master_id,
        pnm.position_master_id,
        COUNT(DISTINCT pnm.cluster_id) AS pnid_mapping_count,
        GROUP_CONCAT(DISTINCT pnm.cluster_id ORDER BY pnm.cluster_id SEPARATOR ', ') AS pnid,
        GROUP_CONCAT(DISTINCT pnm.cluster_label ORDER BY pnm.cluster_label SEPARATOR ' | ') AS pnid_label
      FROM position_nomenclature_mapping pnm
      WHERE pnm.cluster_id IS NOT NULL
      GROUP BY pnm.company_id, pnm.group_master_id, pnm.position_master_id
    ),
    active_worker_positions AS (
      SELECT
        tgm.company_id,
        tci.name AS company_name,
        tepms.employee_number,
        TRIM(CONCAT_WS(' ', te.firstname, te.middlename, te.lastname)) AS employee_name,
        te.corporate_email,
        tpmv.position_master_variant_id,
        tpm.position_master_id,
        tpm.name AS position_name,
        tpm.position_master_type_id,
        COALESCE(tpmt.name, CASE WHEN tpm.position_master_type_id = 5 THEN 'Struktural' ELSE 'Non-struktural' END) AS position_type_name,
        tpmos.organization_master_id AS group_master_id,
        tgm.name AS group_name,
        GROUP_CONCAT(DISTINCT CASE
          WHEN tepms.lakhar_id IS NOT NULL THEN 'LAKHAR'
          WHEN tepms.job_sharing_id IS NOT NULL THEN 'JOB_SHARING'
          ELSE 'DEFINITIF'
        END ORDER BY CASE
          WHEN tepms.lakhar_id IS NULL AND tepms.job_sharing_id IS NULL THEN 0
          WHEN tepms.lakhar_id IS NOT NULL THEN 1
          ELSE 2
        END SEPARATOR ' | ') AS assignment_type,
        CASE WHEN tpm.position_master_type_id = 5 THEN tpm.position_master_id END AS pmid,
        CASE WHEN tpm.position_master_type_id <> 5 THEN pm.pnid END AS pnid,
        CASE WHEN tpm.position_master_type_id <> 5 THEN pm.pnid_label END AS pnid_label,
        CASE WHEN tpm.position_master_type_id <> 5 THEN COALESCE(pm.pnid_mapping_count, 0) ELSE 0 END AS pnid_mapping_count
      FROM tb_employee_position_master_sync tepms
      JOIN tb_employee te
        ON te.employee_number = tepms.employee_number
       AND te.deletedAt IS NULL
       AND te.archived_at IS NULL
      JOIN tb_position_master_variant tpmv
        ON tpmv.position_master_variant_id = tepms.position_master_variant_id
       AND tpmv.deletedAt IS NULL
      JOIN tb_position_master_v2 tpm
        ON tpm.position_master_id = tpmv.position_master_id
       AND tpm.deletedAt IS NULL
       AND CURRENT_TIMESTAMP() BETWEEN COALESCE(tpm.start_date, '1000-01-01') AND COALESCE(tpm.end_date, '9999-12-31')
      LEFT JOIN tb_position_master_type tpmt
        ON tpmt.position_master_type_id = tpm.position_master_type_id
      JOIN tb_position_master_organization_sync tpmos
        ON tpmos.position_master_id = tpm.position_master_id
       AND tpmos.deletedAt IS NULL
       AND CURRENT_TIMESTAMP() BETWEEN COALESCE(tpmos.start_date, '1000-01-01') AND COALESCE(tpmos.end_date, '9999-12-31')
      JOIN tb_group_master tgm
        ON tgm.group_master_id = tpmos.organization_master_id
       AND tgm.deletedAt IS NULL
       {_company_filter(company_id)}
       AND CURRENT_TIMESTAMP() BETWEEN COALESCE(tgm.start_date, '1000-01-01') AND COALESCE(tgm.end_date, '9999-12-31')
      JOIN tb_company_in tci
        ON tci.company_in_id = tgm.company_id
       AND tci.deletedAt IS NULL
       AND CURRENT_TIMESTAMP() BETWEEN COALESCE(tci.start_date, '1000-01-01') AND COALESCE(tci.end_date, '9999-12-31')
      LEFT JOIN pnid_mapping pm
        ON pm.company_id = tgm.company_id
       AND (pm.group_master_id <=> tpmos.organization_master_id)
       AND pm.position_master_id = tpm.position_master_id
      WHERE tepms.deletedAt IS NULL
        AND CURRENT_TIMESTAMP() BETWEEN tepms.start_date AND COALESCE(tepms.end_date, '9999-12-31')
        AND tpm.name NOT REGEXP '^(JA_|JS_)'
      GROUP BY
        tgm.company_id,
        tci.name,
        tepms.employee_number,
        te.firstname,
        te.middlename,
        te.lastname,
        te.corporate_email,
        tpmv.position_master_variant_id,
        tpm.position_master_id,
        tpm.name,
        tpm.position_master_type_id,
        tpmt.name,
        tpmos.organization_master_id,
        tgm.name,
        pm.pnid,
        pm.pnid_label,
        pm.pnid_mapping_count
    ),
    position_origin AS (
      SELECT
        ko.position_master_id,
        COUNT(DISTINCT CASE WHEN k.created_by_pov = 'SYSTEM' THEN ko.kpi_id END) AS system_kpi_count
      FROM kpi_ownership_v3 ko
      JOIN kpi_v3 k ON k.kpi_id = ko.kpi_id
      WHERE ko.deleted_at IS NULL
        AND k.deleted_at IS NULL
        AND k.is_active = 1
        AND COALESCE(k.status, 'ACTIVE') = 'ACTIVE'
        AND ko.year = %(year)s
        AND k.year = %(year)s
      GROUP BY ko.position_master_id
    ),
    employee_portfolio AS (
      SELECT
        ko.employee_number,
        ko.position_master_variant_id,
        ko.position_master_id,
        COUNT(DISTINCT ko.kpi_id) AS total_kpi_count,
        COUNT(DISTINCT CASE WHEN k.created_by_pov IN ('WORKER', 'SUPERIOR') THEN ko.kpi_id END) AS manual_kpi_count,
        COUNT(DISTINCT CASE WHEN k.created_by_pov IS NULL OR k.created_by_pov NOT IN ('SYSTEM', 'WORKER', 'SUPERIOR') THEN ko.kpi_id END) AS unknown_origin_kpi_count,
        COUNT(DISTINCT CASE WHEN k.type = 'IMPACT' THEN ko.kpi_id END) AS impact_count,
        COUNT(DISTINCT CASE WHEN k.type = 'OUTPUT' THEN ko.kpi_id END) AS output_count,
        COUNT(DISTINCT CASE WHEN k.type = 'KAI' THEN ko.kpi_id END) AS kai_count,
        SUM(COALESCE(ko.weight, 0)) AS total_weight,
        COUNT(DISTINCT CASE WHEN k.item_approval_status IN ('UNALLOCATED', 'ALLOCATED', 'DRAFT', 'DRAFT_FROM_SUPERIOR', 'DRAFT_FROM_SUBORDINATE', 'REJECTED') THEN ko.kpi_id END) AS draft_status_count,
        COUNT(DISTINCT CASE WHEN k.item_approval_status IN ('WAITING_REVIEW', 'WAITING_FOR_SUBORDINATE_REVIEW') THEN ko.kpi_id END) AS subordinate_review_count,
        COUNT(DISTINCT CASE WHEN k.item_approval_status IN ('WAITING_FOR_APPROVAL', 'PENDING_CLARIFICATION') THEN ko.kpi_id END) AS manager_decision_count,
        COUNT(DISTINCT CASE WHEN k.item_approval_status IN ('APPROVED', 'APPROVED_ADJUSTED') THEN ko.kpi_id END) AS approved_count,
        MAX(GREATEST(k.updated_at, ko.updated_at)) AS last_kpi_update
      FROM kpi_ownership_v3 ko
      JOIN kpi_v3 k ON k.kpi_id = ko.kpi_id
      WHERE ko.deleted_at IS NULL
        AND k.deleted_at IS NULL
        AND k.is_active = 1
        AND COALESCE(k.status, 'ACTIVE') = 'ACTIVE'
        AND ko.employee_number IS NOT NULL
        AND ko.year = %(year)s
        AND k.year = %(year)s
      GROUP BY ko.employee_number, ko.position_master_variant_id, ko.position_master_id
    )
    SELECT
      awp.*,
      COALESCE(po.system_kpi_count, 0) AS system_kpi_count,
      COALESCE(ep.manual_kpi_count, 0) AS manual_kpi_count,
      COALESCE(ep.unknown_origin_kpi_count, 0) AS unknown_origin_kpi_count,
      COALESCE(ep.total_kpi_count, 0) AS total_kpi_count,
      COALESCE(ep.impact_count, 0) AS impact_count,
      COALESCE(ep.output_count, 0) AS output_count,
      COALESCE(ep.kai_count, 0) AS kai_count,
      COALESCE(ep.total_weight, 0) AS total_weight,
      COALESCE(ep.draft_status_count, 0) AS draft_status_count,
      COALESCE(ep.subordinate_review_count, 0) AS subordinate_review_count,
      COALESCE(ep.manager_decision_count, 0) AS manager_decision_count,
      COALESCE(ep.approved_count, 0) AS approved_count,
      ep.last_kpi_update
    FROM active_worker_positions awp
    LEFT JOIN position_origin po
      ON po.position_master_id = awp.position_master_id
    LEFT JOIN employee_portfolio ep
      ON ep.employee_number = awp.employee_number
     AND ep.position_master_id = awp.position_master_id
     AND (ep.position_master_variant_id <=> awp.position_master_variant_id)
    ORDER BY awp.company_id, awp.group_name, awp.employee_name, awp.position_name
    """


def build_category_summary(
    structural_detail: pd.DataFrame, non_structural_detail: pd.DataFrame
) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    for category, df, unit_label in [
        ("Struktural", structural_detail, "PMID"),
        ("Non-struktural", non_structural_detail, "PNID"),
    ]:
        active_units = len(df)
        complete = int((df["availability_status"] == "Complete").sum()) if active_units else 0
        partial = int((df["availability_status"] == "Partial").sum()) if active_units else 0
        missing = int((df["availability_status"] == "Missing").sum()) if active_units else 0
        rows.append(
            {
                "category": category,
                "unit": unit_label,
                "active_units": active_units,
                "complete_units": complete,
                "partial_units": partial,
                "missing_units": missing,
                "not_complete_units": partial + missing,
                "coverage_pct": round(100 * complete / max(active_units, 1), 1),
            }
        )
    return pd.DataFrame(rows)


def empty_frame(columns: list[str]) -> pd.DataFrame:
    return pd.DataFrame(columns=columns)


def format_db_error(exc: Exception) -> str:
    if isinstance(exc, OperationalError):
        code = exc.args[0] if exc.args else None
        if code == 2013:
            return (
                "Production MySQL connection dropped while the dashboard was loading. "
                "Live data cannot be fetched right now."
            )
        if code == 1045:
            return "Production MySQL authentication failed. Check the dashboard DB credentials."
    return "Unable to load live production data."


def apply_kpi_scope_filter(df: pd.DataFrame, scope: str) -> pd.DataFrame:
    if df.empty or scope == "All":
        return df
    if "with_kpi_variants" not in df.columns:
        return df
    with_kpi = df["with_kpi_variants"].fillna(0).astype(float) > 0
    if scope == "With KPI":
        return df[with_kpi].copy()
    if scope == "Without KPI":
        return df[~with_kpi].copy()
    return df


def fetch_dashboard_data(
    year: int, company_id: int | None
) -> dict[str, pd.DataFrame]:
    config = load_db_config()
    params = {"year": year, "company_id": company_id}
    active_cte = build_active_cte(company_id)
    schedule_company_filter = (
        "AND (company_id = %(company_id)s OR company_id IS NULL)"
        if company_id is not None
        else ""
    )
    with connect(config) as conn:
        data = {
            "db_info": read_sql(
                conn,
                """
                SELECT DATABASE() AS db_name,
                       @@hostname AS host_name,
                       @@read_only AS server_read_only,
                       CURRENT_TIMESTAMP() AS db_now
                """,
            ),
            "schedule": read_sql(
                conn,
                f"""
                SELECT kpi_schedule_id, name, type, year, quarter, company_id, is_active,
                       start_date, end_date, created_at, updated_at
                FROM kpi_schedule_v3
                WHERE deleted_at IS NULL
                  {schedule_company_filter}
                  AND year = %(year)s
                ORDER BY is_active DESC, start_date DESC, kpi_schedule_id DESC
                LIMIT 20
                """,
                params,
            ),
            "company_options": read_sql(
                conn,
                f"""
                WITH {build_active_cte(None)}
                SELECT DISTINCT ap.company_id, ap.company_name
                FROM active_positions ap
                ORDER BY ap.company_id
                """,
                params,
            ),
            "worker_progress": read_sql(
                conn,
                build_worker_progress_sql(company_id),
                params,
            ),
            "coverage": read_sql(
                conn,
                f"""
                WITH {active_cte}
                SELECT
                  COUNT(*) AS active_position_variants,
                  COUNT(DISTINCT ap.position_master_id) AS active_position_masters,
                  SUM(ap.active_employee_count) AS active_employee_assignments,
                  SUM(CASE WHEN ap.primary_assignment_count > 0 THEN 1 ELSE 0 END) AS variants_with_primary_assignment,
                  SUM(CASE WHEN ap.secondary_assignment_count > 0 THEN 1 ELSE 0 END) AS variants_with_secondary_assignment,
                  SUM(CASE WHEN COALESCE(kc.kpi_count, 0) > 0 THEN 1 ELSE 0 END) AS variants_with_any_kpi,
                  SUM(CASE WHEN COALESCE(kc.kpi_count, 0) = 0 THEN 1 ELSE 0 END) AS variants_without_kpi,
                  SUM(CASE WHEN COALESCE(kc.impact_count, 0) > 0
                            AND COALESCE(kc.output_count, 0) > 0
                            AND COALESCE(kc.kai_count, 0) > 0
                           THEN 1 ELSE 0 END) AS variants_with_impact_output_kai,
                  SUM(COALESCE(kc.kpi_count, 0)) AS total_kpi_ownership_rows,
                  SUM(COALESCE(kc.impact_count, 0)) AS total_impact,
                  SUM(COALESCE(kc.output_count, 0)) AS total_output,
                  SUM(COALESCE(kc.kai_count, 0)) AS total_kai,
                  SUM(COALESCE(kc.approved_item_count, 0)) AS approved_items,
                  SUM(COALESCE(kc.approved_weight_count, 0)) AS approved_weights
                FROM active_positions ap
                LEFT JOIN kpi_counts kc
                  ON kc.position_master_id = ap.position_master_id
                 AND (kc.position_master_variant_id <=> ap.position_master_variant_id)
                """,
                params,
            ),
            "group_coverage": read_sql(
                conn,
                f"""
                WITH {active_cte}
                SELECT
                  ap.company_id,
                  ap.company_name,
                  ap.group_name,
                  COUNT(*) AS active_variants,
                  SUM(CASE WHEN ap.primary_assignment_count > 0 THEN 1 ELSE 0 END) AS active_primary_variants,
                  SUM(CASE WHEN COALESCE(kc.kpi_count, 0) > 0 THEN 1 ELSE 0 END) AS with_kpi,
                  SUM(CASE WHEN COALESCE(kc.kpi_count, 0) = 0 THEN 1 ELSE 0 END) AS without_kpi,
                  ROUND(100 * SUM(CASE WHEN COALESCE(kc.kpi_count, 0) > 0 THEN 1 ELSE 0 END) / COUNT(*), 1) AS coverage_pct,
                  SUM(COALESCE(kc.impact_count, 0)) AS impact_count,
                  SUM(COALESCE(kc.output_count, 0)) AS output_count,
                  SUM(COALESCE(kc.kai_count, 0)) AS kai_count
                FROM active_positions ap
                LEFT JOIN kpi_counts kc
                  ON kc.position_master_id = ap.position_master_id
                 AND (kc.position_master_variant_id <=> ap.position_master_variant_id)
                GROUP BY ap.company_id, ap.company_name, ap.group_name
                ORDER BY without_kpi DESC, active_variants DESC, ap.company_id, ap.group_name ASC
                """,
                params,
            ),
            "status": read_sql(
                conn,
                f"""
                WITH {active_cte}
                SELECT
                  k.type,
                  k.item_approval_status,
                  ko.allocation_status,
                  ko.weight_approval_status,
                  COUNT(DISTINCT ko.kpi_ownership_id) AS ownership_rows,
                  COUNT(DISTINCT ko.kpi_id) AS kpi_ids,
                  COUNT(DISTINCT ap.position_master_variant_id) AS active_variants
                FROM active_positions ap
                JOIN kpi_ownership_v3 ko
                  ON ko.position_master_id = ap.position_master_id
                 AND (ko.position_master_variant_id <=> ap.position_master_variant_id)
                JOIN kpi_v3 k ON k.kpi_id = ko.kpi_id
                WHERE ko.deleted_at IS NULL
                  AND k.deleted_at IS NULL
                  AND k.is_active = 1
                  AND COALESCE(k.status, 'ACTIVE') = 'ACTIVE'
                  AND ko.year = %(year)s
                  AND k.year = %(year)s
                GROUP BY k.type, k.item_approval_status, ko.allocation_status, ko.weight_approval_status
                ORDER BY k.type, ownership_rows DESC
                """,
                params,
            ),
            "import_agg": read_sql(
                conn,
                """
                SELECT dry_run, status, COUNT(*) AS logs, SUM(total_rows) AS total_rows,
                       SUM(created_count) AS created_count, SUM(updated_count) AS updated_count,
                       SUM(affected_rows) AS affected_rows, SUM(invalid_count) AS invalid_count,
                       SUM(COALESCE(total_positions, 0)) AS total_positions_sum,
                       MIN(created_at) AS first_log_at, MAX(created_at) AS last_log_at
                FROM kpi_template_import_log
                WHERE planning_year = %(year)s
                GROUP BY dry_run, status
                ORDER BY dry_run, status
                """,
                params,
            ),
            "success_by_file": read_sql(
                conn,
                """
                SELECT file_name, COUNT(*) AS success_logs, SUM(total_rows) AS total_rows,
                       SUM(created_count) AS created_count, SUM(updated_count) AS updated_count,
                       SUM(affected_rows) AS affected_rows, SUM(total_positions) AS total_positions,
                       MIN(created_at) AS first_success_at, MAX(created_at) AS last_success_at
                FROM kpi_template_import_log
                WHERE planning_year = %(year)s AND dry_run = 0 AND status = 'SUCCESS'
                GROUP BY file_name
                ORDER BY last_success_at DESC
                """,
                params,
            ),
            "failed_imports": read_sql(
                conn,
                """
                SELECT import_log_id, created_at, file_name, dry_run, status,
                       total_rows, invalid_count, total_positions,
                       LEFT(COALESCE(error_summary, ''), 500) AS error_summary
                FROM kpi_template_import_log
                WHERE planning_year = %(year)s
                  AND status <> 'SUCCESS'
                ORDER BY created_at DESC, import_log_id DESC
                LIMIT 100
                """,
                params,
            ),
            "without_kpi": read_sql(
                conn,
                f"""
                WITH {active_cte}
                SELECT
                  ap.company_id,
                  ap.company_name,
                  ap.group_name,
                  ap.position_master_id,
                  ap.position_master_variant_id,
                  ap.position_name,
                  ap.active_employee_count,
                  ap.primary_assignment_count,
                  ap.secondary_assignment_count
                FROM active_positions ap
                LEFT JOIN kpi_counts kc
                  ON kc.position_master_id = ap.position_master_id
                 AND (kc.position_master_variant_id <=> ap.position_master_variant_id)
                WHERE COALESCE(kc.kpi_count, 0) = 0
                ORDER BY ap.group_name, ap.position_name
                """,
                params,
            ),
            "structural_detail": read_sql(
                conn,
                f"""
                WITH {active_cte}
                SELECT
                  x.company_id,
                  x.company_name,
                  x.pmid,
                  x.position_name,
                  x.position_type_name,
                  x.group_names,
                  x.active_variants,
                  x.with_kpi_variants,
                  x.without_kpi_variants,
                  x.impact_count,
                  x.output_count,
                  x.kai_count,
                  CASE
                    WHEN x.with_kpi_variants = 0 THEN 'Missing'
                    WHEN x.with_kpi_variants < x.active_variants THEN 'Partial'
                    ELSE 'Complete'
                  END AS availability_status
                FROM (
                  SELECT
                    ap.company_id,
                    ap.company_name,
                    ap.position_master_id AS pmid,
                    MIN(ap.position_name) AS position_name,
                    MAX(COALESCE(ap.position_type_name, 'Struktural')) AS position_type_name,
                    GROUP_CONCAT(DISTINCT ap.group_name ORDER BY ap.group_name SEPARATOR ' | ') AS group_names,
                    COUNT(DISTINCT ap.position_master_variant_id) AS active_variants,
                    COUNT(DISTINCT CASE WHEN COALESCE(kc.kpi_count, 0) > 0 THEN ap.position_master_variant_id END) AS with_kpi_variants,
                    COUNT(DISTINCT CASE WHEN COALESCE(kc.kpi_count, 0) = 0 THEN ap.position_master_variant_id END) AS without_kpi_variants,
                    SUM(COALESCE(kc.impact_count, 0)) AS impact_count,
                    SUM(COALESCE(kc.output_count, 0)) AS output_count,
                    SUM(COALESCE(kc.kai_count, 0)) AS kai_count
                  FROM active_positions ap
                  LEFT JOIN kpi_counts kc
                    ON kc.position_master_id = ap.position_master_id
                   AND (kc.position_master_variant_id <=> ap.position_master_variant_id)
                  WHERE ap.position_master_type_id = 5
                  GROUP BY ap.company_id, ap.company_name, ap.position_master_id
                ) x
                ORDER BY
                  CASE
                    WHEN x.with_kpi_variants = 0 THEN 0
                    WHEN x.with_kpi_variants < x.active_variants THEN 1
                    ELSE 2
                  END,
                  x.position_name
                """,
                params,
            ),
            "non_structural_detail": read_sql(
                conn,
                f"""
                WITH {active_cte}
                SELECT
                  x.company_id,
                  x.company_name,
                  x.pnid,
                  x.pnid_label,
                  x.position_type_names,
                  x.group_names,
                  x.active_pmids,
                  x.active_variants,
                  x.with_kpi_variants,
                  x.without_kpi_variants,
                  x.impact_count,
                  x.output_count,
                  x.kai_count,
                  x.active_pmid_list,
                  x.position_names,
                  CASE
                    WHEN x.with_kpi_variants = 0 THEN 'Missing'
                    WHEN x.with_kpi_variants < x.active_variants THEN 'Partial'
                    ELSE 'Complete'
                  END AS availability_status
                FROM (
                  SELECT
                    ap.company_id,
                    ap.company_name,
                    pnm.cluster_id AS pnid,
                    MAX(COALESCE(pnm.cluster_label, CONCAT('PNID ', pnm.cluster_id))) AS pnid_label,
                    GROUP_CONCAT(DISTINCT COALESCE(ap.position_type_name, 'Non-struktural') ORDER BY ap.position_type_name SEPARATOR ' | ') AS position_type_names,
                    GROUP_CONCAT(DISTINCT ap.group_name ORDER BY ap.group_name SEPARATOR ' | ') AS group_names,
                    COUNT(DISTINCT ap.position_master_id) AS active_pmids,
                    COUNT(DISTINCT ap.position_master_variant_id) AS active_variants,
                    COUNT(DISTINCT CASE WHEN COALESCE(kc.kpi_count, 0) > 0 THEN ap.position_master_variant_id END) AS with_kpi_variants,
                    COUNT(DISTINCT CASE WHEN COALESCE(kc.kpi_count, 0) = 0 THEN ap.position_master_variant_id END) AS without_kpi_variants,
                    SUM(COALESCE(kc.impact_count, 0)) AS impact_count,
                    SUM(COALESCE(kc.output_count, 0)) AS output_count,
                    SUM(COALESCE(kc.kai_count, 0)) AS kai_count,
                    GROUP_CONCAT(DISTINCT ap.position_master_id ORDER BY ap.position_master_id SEPARATOR ', ') AS active_pmid_list,
                    GROUP_CONCAT(DISTINCT ap.position_name ORDER BY ap.position_name SEPARATOR ' | ') AS position_names
                  FROM active_positions ap
                  JOIN position_nomenclature_mapping pnm
                    ON pnm.position_master_id = ap.position_master_id
                   AND pnm.company_id = ap.company_id
                   AND (pnm.group_master_id <=> ap.group_master_id)
                   AND pnm.cluster_id IS NOT NULL
                  LEFT JOIN kpi_counts kc
                    ON kc.position_master_id = ap.position_master_id
                   AND (kc.position_master_variant_id <=> ap.position_master_variant_id)
                  WHERE ap.position_master_type_id <> 5
                  GROUP BY ap.company_id, ap.company_name, pnm.cluster_id
                  HAVING active_pmids >= 1 AND active_variants >= 1
                ) x
                ORDER BY
                  CASE
                    WHEN x.with_kpi_variants = 0 THEN 0
                    WHEN x.with_kpi_variants < x.active_variants THEN 1
                    ELSE 2
                  END,
                  x.pnid_label
                """,
                params,
            ),
            "category_gap_detail": read_sql(
                conn,
                f"""
                WITH {active_cte}
                SELECT
                  ap.company_id,
                  ap.company_name,
                  CASE
                    WHEN ap.position_master_type_id = 5 THEN 'Struktural'
                    ELSE 'Non-struktural'
                  END AS category,
                  pnm.cluster_id AS pnid,
                  COALESCE(pnm.cluster_label, CONCAT('PNID ', pnm.cluster_id)) AS pnid_label,
                  ap.position_master_id AS pmid,
                  ap.position_master_variant_id,
                  ap.position_name,
                  ap.position_type_name,
                  ap.group_name,
                  ap.active_employee_count,
                  ap.primary_assignment_count,
                  ap.secondary_assignment_count
                FROM active_positions ap
                LEFT JOIN position_nomenclature_mapping pnm
                  ON pnm.position_master_id = ap.position_master_id
                 AND pnm.company_id = ap.company_id
                 AND (pnm.group_master_id <=> ap.group_master_id)
                 AND pnm.cluster_id IS NOT NULL
                LEFT JOIN kpi_counts kc
                  ON kc.position_master_id = ap.position_master_id
                 AND (kc.position_master_variant_id <=> ap.position_master_variant_id)
                WHERE COALESCE(kc.kpi_count, 0) = 0
                  AND (
                    ap.position_master_type_id = 5
                    OR (ap.position_master_type_id <> 5 AND pnm.cluster_id IS NOT NULL)
                  )
                ORDER BY category, pnm.cluster_id, ap.position_master_id, ap.position_name
                """,
                params,
            ),
            "category_anomalies": read_sql(
                conn,
                f"""
                WITH {active_cte}
                SELECT
                  'active_non_structural_without_pnid' AS anomaly,
                  COUNT(DISTINCT ap.position_master_id) AS active_pmids,
                  COUNT(DISTINCT ap.position_master_variant_id) AS active_variants
                FROM active_positions ap
                LEFT JOIN position_nomenclature_mapping pnm
                  ON pnm.position_master_id = ap.position_master_id
                 AND pnm.company_id = ap.company_id
                 AND (pnm.group_master_id <=> ap.group_master_id)
                 AND pnm.cluster_id IS NOT NULL
                WHERE ap.position_master_type_id <> 5
                  AND pnm.cluster_id IS NULL
                """,
                params,
            ),
            "anomalies": read_sql(
                conn,
                """
                SELECT anomaly, cnt FROM (
                  SELECT 'active_child_null_parent' AS anomaly, COUNT(*) AS cnt
                  FROM kpi_v3 k
                  WHERE k.deleted_at IS NULL AND k.is_active = 1 AND k.year = %(year)s
                    AND k.type IN ('OUTPUT', 'KAI', 'SUB_IMPACT') AND k.parent_kpi_id IS NULL
                  UNION ALL
                  SELECT 'active_child_parent_missing', COUNT(*)
                  FROM kpi_v3 k
                  LEFT JOIN kpi_v3 p ON p.kpi_id = k.parent_kpi_id AND p.deleted_at IS NULL
                  WHERE k.deleted_at IS NULL AND k.is_active = 1 AND k.year = %(year)s
                    AND k.type IN ('OUTPUT', 'KAI', 'SUB_IMPACT')
                    AND k.parent_kpi_id IS NOT NULL
                    AND p.kpi_id IS NULL
                  UNION ALL
                  SELECT 'active_kpi_missing_title', COUNT(*)
                  FROM kpi_v3 k
                  WHERE k.deleted_at IS NULL AND k.is_active = 1 AND k.year = %(year)s
                    AND (k.title IS NULL OR TRIM(k.title) = '')
                  UNION ALL
                  SELECT 'kpi_target_rows_2026', COUNT(*)
                  FROM kpi_target_v3 kt
                  JOIN kpi_v3 k ON k.kpi_id = kt.kpi_id
                  WHERE kt.deleted_at IS NULL AND k.year = %(year)s
                ) x
                ORDER BY anomaly
                """,
                params,
            ),
        }
        data["worker_progress"] = enrich_worker_progress(data["worker_progress"])
        data["category_summary"] = build_category_summary(
            data["structural_detail"], data["non_structural_detail"]
        )
        data["structural_not_complete"] = data["structural_detail"][
            data["structural_detail"]["availability_status"] != "Complete"
        ].copy()
        data["non_structural_not_complete"] = data["non_structural_detail"][
            data["non_structural_detail"]["availability_status"] != "Complete"
        ].copy()
        return data


def _df_records(df: pd.DataFrame) -> list[dict[str, Any]]:
    records = df.where(pd.notna(df), None).to_dict(orient="records")
    return records


def fetch_company_options() -> pd.DataFrame:
    config = load_db_config()
    with connect(config) as conn:
        return read_sql(
            conn,
            """
            SELECT DISTINCT tci.company_in_id AS company_id, tci.name AS company_name
            FROM tb_employee_position_master_sync tepms
            JOIN tb_employee te
              ON te.employee_number = tepms.employee_number
             AND te.deletedAt IS NULL
             AND te.archived_at IS NULL
            JOIN tb_position_master_variant tpmv
              ON tpmv.position_master_variant_id = tepms.position_master_variant_id
             AND tpmv.deletedAt IS NULL
            JOIN tb_position_master_v2 tpm
              ON tpm.position_master_id = tpmv.position_master_id
             AND tpm.deletedAt IS NULL
             AND CURRENT_TIMESTAMP() BETWEEN COALESCE(tpm.start_date, '1000-01-01') AND COALESCE(tpm.end_date, '9999-12-31')
            JOIN tb_position_master_organization_sync tpmos
              ON tpmos.position_master_id = tpm.position_master_id
             AND tpmos.deletedAt IS NULL
             AND CURRENT_TIMESTAMP() BETWEEN COALESCE(tpmos.start_date, '1000-01-01') AND COALESCE(tpmos.end_date, '9999-12-31')
            JOIN tb_group_master tgm
              ON tgm.group_master_id = tpmos.organization_master_id
             AND tgm.deletedAt IS NULL
             AND CURRENT_TIMESTAMP() BETWEEN COALESCE(tgm.start_date, '1000-01-01') AND COALESCE(tgm.end_date, '9999-12-31')
            JOIN tb_company_in tci
              ON tci.company_in_id = tgm.company_id
             AND tci.deletedAt IS NULL
             AND CURRENT_TIMESTAMP() BETWEEN COALESCE(tci.start_date, '1000-01-01') AND COALESCE(tci.end_date, '9999-12-31')
            WHERE tepms.deletedAt IS NULL
              AND CURRENT_TIMESTAMP() BETWEEN tepms.start_date AND COALESCE(tepms.end_date, '9999-12-31')
            ORDER BY tci.company_in_id
            """,
        )


def build_check_payload(
    data: dict[str, pd.DataFrame], year: int, company_id: int | None
) -> dict[str, Any]:
    coverage = _df_records(data["coverage"])[0]
    worker_detail = data["worker_progress"]
    worker_summary = build_worker_level_progress(worker_detail)
    progress_summary = build_progress_summary(worker_detail)
    origin_summary = (
        worker_detail.groupby("portfolio_origin").size().to_dict()
        if not worker_detail.empty
        else {}
    )
    readiness_summary = (
        worker_summary.groupby("readiness_status").size().to_dict()
        if not worker_summary.empty
        else {}
    )
    db_info = _df_records(data["db_info"])[0]
    return {
        "db": {
            "name": db_info.get("db_name"),
            "host": db_info.get("host_name"),
            "server_read_only": db_info.get("server_read_only"),
            "db_now": str(db_info.get("db_now")),
        },
        "year": year,
        "company_id": company_id,
        "company_scope": "ALL_PELINDO" if company_id is None else "SELECTED_COMPANY",
        "schedule": _df_records(data["schedule"][:20]),
        "coverage": coverage,
        "worker_progress": progress_summary,
        "portfolio_origin_summary": origin_summary,
        "worker_readiness_summary": readiness_summary,
        "company_count": int(data["company_options"]["company_id"].nunique()),
        "import_agg": _df_records(data["import_agg"]),
        "anomalies": _df_records(data["anomalies"]),
        "category_summary": _df_records(data["category_summary"]),
        "structural_rows": int(len(data["structural_detail"])),
        "structural_not_complete_rows": int(len(data["structural_not_complete"])),
        "non_structural_rows": int(len(data["non_structural_detail"])),
        "non_structural_not_complete_rows": int(len(data["non_structural_not_complete"])),
        "category_gap_rows": int(len(data["category_gap_detail"])),
        "category_anomalies": _df_records(data["category_anomalies"]),
        "without_kpi_rows": int(len(data["without_kpi"])),
        "group_rows": int(len(data["group_coverage"])),
    }


def run_check(year: int, company_id: int | None) -> None:
    data = fetch_dashboard_data(year, company_id)
    print(
        json.dumps(
            build_check_payload(data, year, company_id),
            default=str,
            indent=2,
        )
    )


def format_int(value: Any) -> str:
    try:
        return f"{int(value):,}".replace(",", ".")
    except (TypeError, ValueError):
        return "0"


def run_app(
    *,
    data_loader=fetch_dashboard_data,
    company_loader=fetch_company_options,
) -> None:
    import plotly.express as px
    import streamlit as st

    st.set_page_config(
        page_title="KPI Planning Dashboard Production",
        page_icon="KPI",
        layout="wide",
    )

    st.title("KPI Planning Dashboard Production")

    try:
        company_options = company_loader()
    except Exception as exc:
        st.error(format_db_error(exc))
        st.caption("Dashboard unavailable. Live production data could not be loaded.")
        with st.expander("Technical details"):
            st.code(str(exc))
        st.stop()

    company_names = {
        int(row.company_id): str(row.company_name)
        for row in company_options.itertuples(index=False)
    }
    company_values: list[int | None] = [None, *company_names.keys()]
    default_company_index = (
        company_values.index(DEFAULT_COMPANY_ID)
        if DEFAULT_COMPANY_ID in company_values
        else 0
    )

    with st.sidebar:
        st.header("Filters")
        year = st.number_input("Planning year", min_value=2024, max_value=2030, value=DEFAULT_YEAR, step=1)
        company_id = st.selectbox(
            "Company ID - Name",
            company_values,
            index=default_company_index,
            format_func=lambda value: (
                "Seluruh Pelindo"
                if value is None
                else f"{value} - {company_names.get(value, 'Unknown company')}"
            ),
        )
        worker_search = st.text_input("Cari pekerja, NIPP, posisi, atau group", "")
        selected_origins = st.multiselect(
            "Origin portfolio",
            PORTFOLIO_ORIGINS,
            default=PORTFOLIO_ORIGINS,
            format_func=lambda value: ORIGIN_LABELS[value],
        )
        selected_statuses = st.multiselect(
            "Status readiness",
            READINESS_ORDER,
            default=READINESS_ORDER,
            format_func=lambda value: READINESS_LABELS[value],
        )
        include_approved = st.checkbox("Tampilkan Disetujui", value=False)
        only_gaps = st.checkbox("Show only groups with gaps", value=True)
        max_rows = st.slider("Detail rows", min_value=20, max_value=500, value=120, step=20)
        refresh = st.button("Refresh data")

    try:
        data = data_loader(int(year), company_id)
    except Exception as exc:
        st.error(format_db_error(exc))
        st.caption("Dashboard unavailable. Live production data could not be loaded.")
        with st.expander("Technical details"):
            st.code(str(exc))
        st.stop()

    coverage = data["coverage"].iloc[0].to_dict()
    completion = (
        float(coverage.get("variants_with_any_kpi") or 0)
        / max(float(coverage.get("active_position_variants") or 0), 1)
    )

    db_info = data["db_info"].iloc[0].to_dict()
    schedule = data["schedule"]
    schedule_text = "No active schedule found"
    if not schedule.empty:
        active = schedule.iloc[0]
        schedule_text = (
            f"{active.get('name')} | {active.get('start_date')} - {active.get('end_date')}"
        )

    st.caption(
        f"Source: production `{db_info.get('db_name')}` on `{db_info.get('host_name')}`. "
        f"DB time: {db_info.get('db_now')}. Schedule: {schedule_text}."
    )
    if int(db_info.get("server_read_only") or 0) == 0:
        st.warning("Production server is not read-only. This dashboard only runs internal SELECT queries.")

    kpi_cols = st.columns(6)
    kpi_cols[0].metric("Active variants", format_int(coverage.get("active_position_variants")))
    kpi_cols[1].metric("Active masters", format_int(coverage.get("active_position_masters")))
    kpi_cols[2].metric("With KPI", format_int(coverage.get("variants_with_any_kpi")))
    kpi_cols[3].metric("Without KPI", format_int(coverage.get("variants_without_kpi")))
    kpi_cols[4].metric("Completion", f"{completion:.1%}")
    kpi_cols[5].metric("KPI ownership rows", format_int(coverage.get("total_kpi_ownership_rows")))

    st.divider()

    st.subheader("Progress perencanaan KPI per pekerja")
    worker_detail = data["worker_progress"].copy()
    progress_summary = build_progress_summary(worker_detail)
    progress_cols = st.columns(6)
    progress_cols[0].metric("Pekerja aktif", format_int(progress_summary["active_workers"]))
    progress_cols[1].metric("Assignment aktif", format_int(progress_summary["active_assignments"]))
    progress_cols[2].metric("Assignment dengan kamus", format_int(progress_summary["dictionary_assignments"]))
    progress_cols[3].metric("Manual tanpa kamus", format_int(progress_summary["manual_assignments"]))
    progress_cols[4].metric("Perlu follow-up", format_int(progress_summary["followup_workers"]))
    progress_cols[5].metric("Pekerja disetujui", format_int(progress_summary["approved_workers"]))

    filtered_workers = filter_worker_progress(
        worker_detail,
        search=worker_search,
        origins=selected_origins,
        statuses=selected_statuses,
        include_approved=include_approved,
    )

    progress_left, progress_right = st.columns([1.2, 1])
    with progress_left:
        progress_plot = (
            worker_detail.groupby(["portfolio_origin_label", "readiness_label"], as_index=False)
            .size()
            .rename(columns={"size": "assignment_count"})
        )
        if not progress_plot.empty:
            progress_fig = px.bar(
                progress_plot,
                x="portfolio_origin_label",
                y="assignment_count",
                color="readiness_label",
                text="assignment_count",
                labels={
                    "portfolio_origin_label": "Origin portfolio",
                    "assignment_count": "Assignment",
                    "readiness_label": "Readiness",
                },
            )
            progress_fig.update_layout(barmode="stack", height=420)
            st.plotly_chart(progress_fig, width="stretch")
        else:
            st.info("Tidak ada assignment pekerja aktif pada scope ini.")

    with progress_right:
        worker_level = build_worker_level_progress(worker_detail)
        readiness_table = (
            worker_level.groupby(["readiness_status", "readiness_label"], as_index=False)
            .size()
            .rename(columns={"size": "worker_count"})
        )
        if not readiness_table.empty:
            readiness_table["_order"] = readiness_table["readiness_status"].map(
                {status: index for index, status in enumerate(READINESS_ORDER)}
            )
            readiness_table = readiness_table.sort_values("_order").drop(columns="_order")
        st.markdown("**Readiness pekerja unik**")
        st.dataframe(readiness_table, width="stretch", hide_index=True)

    if company_id is None and not worker_detail.empty:
        company_summary = (
            worker_detail.groupby(["company_id", "company_name"], as_index=False)
            .agg(
                active_workers=("employee_number", "nunique"),
                active_assignments=("position_master_variant_id", "size"),
                followup_assignments=("readiness_status", lambda values: (values != "DISETUJUI").sum()),
            )
            .sort_values(["followup_assignments", "active_workers"], ascending=False)
        )
        with st.expander("Ringkasan progress per company", expanded=True):
            st.dataframe(company_summary, width="stretch", hide_index=True, height=360)

    st.subheader("Daftar pekerja untuk follow-up")
    visible_columns = [
        column for column in WORKER_PROGRESS_EXPORT_COLUMNS if column in filtered_workers
    ]
    st.dataframe(
        filtered_workers[visible_columns].head(max_rows),
        width="stretch",
        hide_index=True,
        height=520,
    )
    st.download_button(
        "Unduh CSV follow-up",
        data=worker_progress_csv(filtered_workers),
        file_name=f"kpi-planning-followup-{year}-{'all' if company_id is None else company_id}.csv",
        mime="text/csv",
    )

    anomaly_rows = worker_detail[
        (worker_detail["mapping_anomaly"] != "")
        | (worker_detail["portfolio_origin"] == "ORIGIN_TIDAK_DIKENAL")
    ]
    if not anomaly_rows.empty:
        st.warning("Ada assignment dengan mapping PNID atau origin KPI yang perlu diperiksa.")
        with st.expander("Anomali progress pekerja"):
            st.dataframe(anomaly_rows[visible_columns], width="stretch", hide_index=True)

    st.divider()

    st.subheader("Ketersediaan kamus KPI by kategori posisi")
    category_summary = data["category_summary"].copy()
    structural_row = category_summary[category_summary["category"] == "Struktural"].iloc[0].to_dict()
    non_structural_row = category_summary[category_summary["category"] == "Non-struktural"].iloc[0].to_dict()

    category_cols = st.columns(6)
    category_cols[0].metric("Active structural PMIDs", format_int(structural_row.get("active_units")))
    category_cols[1].metric("Structural complete", format_int(structural_row.get("complete_units")))
    category_cols[2].metric("Structural not complete", format_int(structural_row.get("not_complete_units")))
    category_cols[3].metric("Active non-structural PNIDs", format_int(non_structural_row.get("active_units")))
    category_cols[4].metric("Non-structural complete", format_int(non_structural_row.get("complete_units")))
    category_cols[5].metric("Non-structural not complete", format_int(non_structural_row.get("not_complete_units")))

    category_plot = category_summary.melt(
        id_vars=["category", "unit"],
        value_vars=["complete_units", "partial_units", "missing_units"],
        var_name="availability_status",
        value_name="count",
    )
    category_plot["availability_status"] = category_plot["availability_status"].map(
        {
            "complete_units": "Complete",
            "partial_units": "Partial",
            "missing_units": "Missing",
        }
    )
    category_fig = px.bar(
        category_plot,
        x="category",
        y="count",
        color="availability_status",
        text="count",
        labels={"category": "Kategori", "count": "Jumlah", "availability_status": "Status"},
    )
    category_fig.update_layout(barmode="stack", height=360)
    st.plotly_chart(category_fig, width="stretch")

    category_anomalies = data["category_anomalies"].copy()
    if not category_anomalies.empty:
        unmapped = int(category_anomalies.iloc[0].get("active_pmids") or 0)
        if unmapped > 0:
            st.warning(
                "Ada active non-structural PMID yang belum punya PNID mapping. "
                "Baris ini tidak masuk denominator PNID sampai mapping tersedia."
            )
            st.dataframe(category_anomalies, width="stretch", hide_index=True)

    structural_tab, non_structural_tab, missing_tab = st.tabs(
        ["Structural PMIDs", "Non-structural PNIDs", "Belum lengkap"]
    )
    with structural_tab:
        structural_scope = st.selectbox(
            "KPI status filter",
            ["All", "With KPI", "Without KPI"],
            key="structural_kpi_status_filter",
        )
        st.dataframe(
            apply_kpi_scope_filter(data["structural_detail"], structural_scope),
            width="stretch",
            hide_index=True,
            height=420,
        )
    with non_structural_tab:
        non_structural_scope = st.selectbox(
            "KPI status filter",
            ["All", "With KPI", "Without KPI"],
            key="non_structural_kpi_status_filter",
        )
        st.dataframe(
            apply_kpi_scope_filter(data["non_structural_detail"], non_structural_scope),
            width="stretch",
            hide_index=True,
            height=420,
        )
    with missing_tab:
        st.markdown("Structural PMID belum lengkap")
        st.dataframe(data["structural_not_complete"], width="stretch", hide_index=True, height=320)
        st.markdown("Non-structural PNID belum lengkap")
        st.dataframe(data["non_structural_not_complete"], width="stretch", hide_index=True, height=320)
        st.markdown("Active PNID/PMID/Position Variant tanpa KPI")
        st.dataframe(data["category_gap_detail"], width="stretch", hide_index=True, height=520)

    st.divider()

    group_df = data["group_coverage"].copy()
    if worker_search:
        group_df = group_df[
            group_df["group_name"].str.contains(worker_search, case=False, na=False)
        ]
    if only_gaps:
        group_df = group_df[group_df["without_kpi"].astype(float) > 0]

    left, right = st.columns([1.2, 1])
    with left:
        st.subheader("Coverage gaps by group")
        plot_df = group_df.head(30).sort_values("without_kpi", ascending=True)
        fig = px.bar(
            plot_df,
            x="without_kpi",
            y="group_name",
            orientation="h",
            hover_data=["active_variants", "with_kpi", "coverage_pct"],
            labels={"without_kpi": "Without KPI", "group_name": "Group"},
        )
        fig.update_layout(height=max(420, min(900, 28 * len(plot_df) + 120)))
        st.plotly_chart(fig, width="stretch")

    with right:
        st.subheader("Planning status by KPI type")
        status_df = data["status"].copy()
        fig = px.bar(
            status_df,
            x="type",
            y="ownership_rows",
            color="weight_approval_status",
            hover_data=["item_approval_status", "allocation_status", "kpi_ids", "active_variants"],
            labels={"ownership_rows": "Ownership rows", "type": "KPI type"},
        )
        st.plotly_chart(fig, width="stretch")
        st.dataframe(status_df, width="stretch", hide_index=True)

    st.subheader("Upload audit")
    st.caption("Global untuk seluruh Pelindo; log import tidak memiliki Company ID yang andal untuk filter.")
    import_cols = st.columns([1, 1])
    with import_cols[0]:
        agg = data["import_agg"]
        fig = px.bar(
            agg,
            x="status",
            y="logs",
            color=agg["dry_run"].map({0: "Actual", 1: "Dry run"}),
            labels={"color": "Mode", "logs": "Log count"},
        )
        st.plotly_chart(fig, width="stretch")
        st.dataframe(agg, width="stretch", hide_index=True)
    with import_cols[1]:
        st.markdown("Latest successful uploads")
        st.dataframe(
            data["success_by_file"].head(12),
            width="stretch",
            hide_index=True,
        )

    st.subheader("Anomaly checks")
    anomaly_df = data["anomalies"].copy()
    st.dataframe(anomaly_df, width="stretch", hide_index=True)
    if not data["failed_imports"].empty:
        st.markdown("Failed or non-success import logs")
        st.dataframe(data["failed_imports"].head(max_rows), width="stretch", hide_index=True)

    st.subheader("Active position variants without KPI")
    st.dataframe(data["without_kpi"].head(max_rows), width="stretch", hide_index=True)

    with st.expander("Group coverage detail"):
        st.dataframe(group_df.head(max_rows), width="stretch", hide_index=True)

    if refresh:
        st.rerun()


def main() -> None:
    parser = argparse.ArgumentParser(description="KPI planning dashboard")
    parser.add_argument("--check", action="store_true", help="Run a production read smoke test and print JSON")
    parser.add_argument("--year", type=int, default=DEFAULT_YEAR)
    parser.add_argument(
        "--company-id",
        type=parse_company_id,
        default=DEFAULT_COMPANY_ID,
        help="Company ID; omit or use 'all' for all Pelindo companies",
    )
    args = parser.parse_args()
    if args.check:
        run_check(args.year, args.company_id)
    else:
        run_app()


if __name__ == "__main__":
    main()
