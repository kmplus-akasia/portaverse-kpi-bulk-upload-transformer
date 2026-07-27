#!/usr/bin/env node
import { createRequire } from "node:module";
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";

const args = process.argv.slice(2);
const outputIndex = args.indexOf("--output");
const profileIndex = args.indexOf("--profile");

const outputPath =
  outputIndex >= 0 && args[outputIndex + 1]
    ? args[outputIndex + 1]
    : "configs/production_position_reference.json";
const profile =
  profileIndex >= 0 && args[profileIndex + 1] ? args[profileIndex + 1] : "production";

const envPath = path.join(process.env.HOME, ".codex/pms-connections", `${profile}.env`);
const env = {};
if (existsSync(envPath)) {
  for (const line of readFileSync(envPath, "utf8").split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;
    const idx = trimmed.indexOf("=");
    if (idx < 0) continue;
    env[trimmed.slice(0, idx).trim()] = trimmed
      .slice(idx + 1)
      .trim()
      .replace(/^['"]|['"]$/g, "");
  }
}
for (const key of [
  "DB_ENGINE",
  "DB_HOST",
  "DB_PORT",
  "DB_NAME",
  "DB_USER",
  "DB_PASSWORD",
  "DB_SSL",
]) {
  if (process.env[key]) env[key] = process.env[key];
}

const required = ["DB_ENGINE", "DB_HOST", "DB_PORT", "DB_NAME", "DB_USER", "DB_PASSWORD"];
const missing = required.filter((key) => !env[key]);
if (missing.length) {
  console.error(`Profile ${profile} is incomplete. Missing: ${missing.join(", ")}`);
  process.exit(1);
}
if (!["mysql", "mariadb"].includes(String(env.DB_ENGINE).toLowerCase())) {
  console.error(`Unsupported DB_ENGINE: ${env.DB_ENGINE}`);
  process.exit(1);
}

const servicePackage = "/Users/alfredoteja/Documents/pms-codebase/pms-service/package.json";
const serviceRequire = createRequire(servicePackage);
const mysql = serviceRequire("mysql2/promise");

const connection = await mysql.createConnection({
  host: env.DB_HOST,
  port: Number(env.DB_PORT),
  user: env.DB_USER,
  password: env.DB_PASSWORD,
  database: env.DB_NAME,
  ssl: env.DB_SSL === "1" ? {} : undefined,
});

try {
  const [nomenclatureRows] = await connection.query(
    `SELECT pnm.id,
            pnm.cluster_id,
            pnm.cluster_label,
            pnm.position_master_id,
            pnm.job_class_level,
            pnm.position_name,
            pnm.position_master_type_id,
            pnm.type_name,
            pnm.group_master_id,
            pnm.group_name,
            pnm.company_id,
            pnm.company_name,
            tgm.name AS active_group_name,
            tgm.org_level AS active_group_org_level,
            tgm.org_type AS active_group_org_type,
            tgm.costcenter AS active_group_costcenter,
            CASE
              WHEN tgm.group_master_id IS NOT NULL
               AND tgm.deletedAt IS NULL
               AND (CURRENT_TIMESTAMP() BETWEEN COALESCE(tgm.start_date, '1000-01-01')
                    AND COALESCE(tgm.end_date, '9999-12-31'))
              THEN 1 ELSE 0
            END AS is_group_active,
            tci.name AS active_company_name,
            tci.code AS active_company_code,
            tci.type_org AS active_company_type_org,
            CASE
              WHEN tci.company_in_id IS NOT NULL
               AND tci.deletedAt IS NULL
               AND (CURRENT_TIMESTAMP() BETWEEN COALESCE(tci.start_date, '1000-01-01')
                    AND COALESCE(tci.end_date, '9999-12-31'))
              THEN 1 ELSE 0
            END AS is_company_active
       FROM position_nomenclature_mapping pnm
       LEFT JOIN tb_group_master tgm
         ON tgm.group_master_id = pnm.group_master_id
       LEFT JOIN tb_company_in tci
         ON tci.company_in_id = pnm.company_id
      WHERE pnm.position_master_id IS NOT NULL
      ORDER BY pnm.cluster_id ASC, pnm.position_master_id ASC, pnm.group_master_id ASC`,
  );

  const [positionMasterRows] = await connection.query(
    `SELECT tpm.position_master_id,
            tpm.name AS position_name,
            tpm.job_class_level,
            tpm.job_score,
            tpm.total_position_max,
            tpm.work_unit,
            tpm.position_master_type_id,
            tpm.position_master_urgency_id,
            tpm.is_job_assignment,
            tpm.is_career_path,
            tpm.cohort_id,
            tpm.start_date AS position_start_date,
            tpm.end_date AS position_end_date,
            CASE
              WHEN tpm.deletedAt IS NULL
               AND (CURRENT_TIMESTAMP() BETWEEN COALESCE(tpm.start_date, '1000-01-01')
                    AND COALESCE(tpm.end_date, '9999-12-31'))
              THEN 1 ELSE 0
            END AS is_position_active,
            tpmos.position_master_organization_sync_id,
            tpmos.organization_master_id AS group_master_id,
            tpmos.start_date AS organization_start_date,
            tpmos.end_date AS organization_end_date,
            CASE
              WHEN tpmos.position_master_organization_sync_id IS NOT NULL
               AND tpmos.deletedAt IS NULL
               AND (CURRENT_TIMESTAMP() BETWEEN COALESCE(tpmos.start_date, '1000-01-01')
                    AND COALESCE(tpmos.end_date, '9999-12-31'))
              THEN 1 ELSE 0
            END AS is_position_organization_active,
            tgm.name AS group_name,
            tgm.org_level AS group_org_level,
            tgm.org_type AS group_org_type,
            tgm.costcenter AS group_costcenter,
            CASE
              WHEN tgm.group_master_id IS NOT NULL
               AND tgm.deletedAt IS NULL
               AND (CURRENT_TIMESTAMP() BETWEEN COALESCE(tgm.start_date, '1000-01-01')
                    AND COALESCE(tgm.end_date, '9999-12-31'))
              THEN 1 ELSE 0
            END AS is_group_active,
            tgm.company_id,
            tci.name AS company_name,
            tci.code AS company_code,
            tci.type_org AS company_type_org,
            CASE
              WHEN tci.company_in_id IS NOT NULL
               AND tci.deletedAt IS NULL
               AND (CURRENT_TIMESTAMP() BETWEEN COALESCE(tci.start_date, '1000-01-01')
                    AND COALESCE(tci.end_date, '9999-12-31'))
              THEN 1 ELSE 0
            END AS is_company_active
       FROM tb_position_master_v2 tpm
       LEFT JOIN tb_position_master_organization_sync tpmos
         ON tpmos.position_master_id = tpm.position_master_id
        AND tpmos.deletedAt IS NULL
       LEFT JOIN tb_group_master tgm
         ON tgm.group_master_id = tpmos.organization_master_id
       LEFT JOIN tb_company_in tci
         ON tci.company_in_id = tgm.company_id
      WHERE tpm.deletedAt IS NULL
      ORDER BY tpm.position_master_id ASC, tpmos.organization_master_id ASC`,
  );

  const [organizationRows] = await connection.query(
    `SELECT tgm.group_master_id,
            tgm.name AS group_name,
            tgm.parent_id,
            tgm.company_id,
            tci.name AS company_name,
            tci.code AS company_code,
            tgm.org_level,
            tgm.org_type,
            tgm.costcenter,
            tgm.start_date,
            tgm.end_date,
            CASE
              WHEN tgm.deletedAt IS NULL
               AND (CURRENT_TIMESTAMP() BETWEEN COALESCE(tgm.start_date, '1000-01-01')
                    AND COALESCE(tgm.end_date, '9999-12-31'))
              THEN 1 ELSE 0
            END AS is_group_active,
            CASE
              WHEN tci.company_in_id IS NOT NULL
               AND tci.deletedAt IS NULL
               AND (CURRENT_TIMESTAMP() BETWEEN COALESCE(tci.start_date, '1000-01-01')
                    AND COALESCE(tci.end_date, '9999-12-31'))
              THEN 1 ELSE 0
            END AS is_company_active
       FROM tb_group_master tgm
       LEFT JOIN tb_company_in tci
         ON tci.company_in_id = tgm.company_id
      WHERE tgm.deletedAt IS NULL
      ORDER BY tgm.group_master_id ASC`,
  );

  const [companyRows] = await connection.query(
    `SELECT company_in_id,
            parent_id,
            objid,
            name AS company_name,
            code AS company_code,
            type_org,
            company_type_id,
            tier,
            is_internal,
            start_date,
            end_date,
            CASE
              WHEN deletedAt IS NULL
               AND (CURRENT_TIMESTAMP() BETWEEN COALESCE(start_date, '1000-01-01')
                    AND COALESCE(end_date, '9999-12-31'))
              THEN 1 ELSE 0
            END AS is_company_active
       FROM tb_company_in
      WHERE deletedAt IS NULL
      ORDER BY company_in_id ASC`,
  );

  const [activeAssignmentRows] = await connection.query(
    `SELECT tpm.position_master_id,
            tpmos.organization_master_id AS group_master_id,
            tgm.company_id,
            COUNT(DISTINCT tpmv.position_master_variant_id) AS active_variant_count,
            COUNT(DISTINCT tepms.employee_number) AS active_employee_count,
            COUNT(DISTINCT CASE
              WHEN tepms.lakhar_id IS NULL AND tepms.job_sharing_id IS NULL THEN tepms.employee_number
            END) AS definitive_employee_count,
            COUNT(DISTINCT CASE
              WHEN tepms.lakhar_id IS NOT NULL OR tepms.job_sharing_id IS NOT NULL THEN tepms.employee_number
            END) AS secondary_employee_count,
            GROUP_CONCAT(DISTINCT TRIM(CONCAT_WS(' ', te.firstname, te.middlename, te.lastname))
              ORDER BY CASE
                WHEN tepms.lakhar_id IS NULL AND tepms.job_sharing_id IS NULL THEN 0
                WHEN tepms.lakhar_id IS NOT NULL THEN 1
                ELSE 2
              END, te.employee_number SEPARATOR '; ') AS active_employee_names,
            GROUP_CONCAT(DISTINCT tepms.employee_number
              ORDER BY CASE
                WHEN tepms.lakhar_id IS NULL AND tepms.job_sharing_id IS NULL THEN 0
                WHEN tepms.lakhar_id IS NOT NULL THEN 1
                ELSE 2
              END, te.employee_number SEPARATOR '; ') AS active_employee_nipps,
            GROUP_CONCAT(DISTINCT CASE
              WHEN tepms.lakhar_id IS NULL AND tepms.job_sharing_id IS NULL
              THEN TRIM(CONCAT_WS(' ', te.firstname, te.middlename, te.lastname))
            END ORDER BY te.employee_number SEPARATOR '; ') AS definitive_employee_names,
            GROUP_CONCAT(DISTINCT CASE
              WHEN tepms.lakhar_id IS NULL AND tepms.job_sharing_id IS NULL
              THEN tepms.employee_number
            END ORDER BY te.employee_number SEPARATOR '; ') AS definitive_employee_nipps,
            GROUP_CONCAT(DISTINCT CASE
              WHEN tepms.lakhar_id IS NOT NULL OR tepms.job_sharing_id IS NOT NULL
              THEN TRIM(CONCAT_WS(' ', te.firstname, te.middlename, te.lastname))
            END ORDER BY te.employee_number SEPARATOR '; ') AS secondary_employee_names,
            GROUP_CONCAT(DISTINCT CASE
              WHEN tepms.lakhar_id IS NOT NULL OR tepms.job_sharing_id IS NOT NULL
              THEN tepms.employee_number
            END ORDER BY te.employee_number SEPARATOR '; ') AS secondary_employee_nipps
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
        AND tpm.name NOT REGEXP '^(JA_|JS_)'
      GROUP BY tpm.position_master_id, tpmos.organization_master_id, tgm.company_id`,
  );

  const activeByPositionOrg = new Map();
  for (const row of activeAssignmentRows) {
    activeByPositionOrg.set(
      `${row.position_master_id}|${row.group_master_id}|${row.company_id}`,
      row,
    );
  }
  const withActiveAssignment = (row) => {
    const active = activeByPositionOrg.get(
      `${row.position_master_id}|${row.group_master_id}|${row.company_id}`,
    );
    if (!active) return null;
    return {
      ...row,
      active_variant_count: Number(active.active_variant_count || 0),
      active_employee_count: Number(active.active_employee_count || 0),
      definitive_employee_count: Number(active.definitive_employee_count || 0),
      secondary_employee_count: Number(active.secondary_employee_count || 0),
      active_employee_names: active.active_employee_names || "",
      active_employee_nipps: active.active_employee_nipps || "",
      definitive_employee_names: active.definitive_employee_names || "",
      definitive_employee_nipps: active.definitive_employee_nipps || "",
      secondary_employee_names: active.secondary_employee_names || "",
      secondary_employee_nipps: active.secondary_employee_nipps || "",
    };
  };
  const structuralLookupRows = positionMasterRows
    .filter(
      (row) =>
        String(row.position_master_type_id || "") === "5" &&
        Number(row.is_position_active || 0) === 1 &&
        Number(row.is_position_organization_active || 0) === 1 &&
        Number(row.is_group_active || 0) === 1 &&
        Number(row.is_company_active || 0) === 1,
    )
    .map(withActiveAssignment)
    .filter(Boolean);
  const nonStructuralLookupRows = nomenclatureRows
    .filter(
      (row) =>
        String(row.position_master_type_id || "") !== "5" &&
        row.cluster_id &&
        Number(row.is_group_active || 0) === 1 &&
        Number(row.is_company_active || 0) === 1,
    )
    .map(withActiveAssignment)
    .filter(Boolean);

  const payload = {
    source: {
      profile,
      database: env.DB_NAME,
      exported_at: new Date().toISOString(),
      read_only: true,
      review_status: "current_snapshot_unreviewed",
      review_note:
        "Current production snapshot is exported for implementation/testing and still requires data-owner review before final business use.",
      tables: [
        "position_nomenclature_mapping",
        "tb_position_master_v2",
        "tb_position_master_organization_sync",
        "tb_position_master_variant",
        "tb_employee_position_master_sync",
        "tb_employee",
        "tb_group_master",
        "tb_company_in",
      ],
      notes:
        "Offline reference dataset for KPI converter. Compatible rows and position_master_rows keys are retained for scripts/kpi_bulk_transform.py.",
    },
    structural_lookup_rows: structuralLookupRows,
    non_structural_lookup_rows: nonStructuralLookupRows,
    rows: nomenclatureRows,
    position_master_rows: positionMasterRows,
    active_assignment_rows: activeAssignmentRows,
    organization_rows: organizationRows,
    company_rows: companyRows,
  };
  mkdirSync(path.dirname(outputPath), { recursive: true });
  writeFileSync(outputPath, `${JSON.stringify(payload, null, 2)}\n`, "utf8");
  console.log(`Wrote nomenclature rows: ${nomenclatureRows.length}`);
  console.log(`Wrote position master/org rows: ${positionMasterRows.length}`);
  console.log(`Wrote active structural lookup rows: ${structuralLookupRows.length}`);
  console.log(`Wrote active non-structural lookup rows: ${nonStructuralLookupRows.length}`);
  console.log(`Wrote active assignment rows: ${activeAssignmentRows.length}`);
  console.log(`Wrote organization rows: ${organizationRows.length}`);
  console.log(`Wrote company rows: ${companyRows.length}`);
  console.log(`Output: ${outputPath}`);
} finally {
  await connection.end();
}
