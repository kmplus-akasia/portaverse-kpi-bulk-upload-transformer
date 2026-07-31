#!/usr/bin/env node
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { requireFromPmsService, userHomeDir } from "./pms_service_require.mjs";

const args = process.argv.slice(2);
const outputIndex = args.indexOf("--output");
const profileIndex = args.indexOf("--profile");
const yearIndex = args.indexOf("--year");

const outputPath =
  outputIndex >= 0 && args[outputIndex + 1]
    ? args[outputIndex + 1]
    : "output/group1_ho_v2_delta_remediation_20260709/production_kpi_snapshot_20260709.json";
const profile =
  profileIndex >= 0 && args[profileIndex + 1] ? args[profileIndex + 1] : "production";
const year = yearIndex >= 0 && args[yearIndex + 1] ? Number(args[yearIndex + 1]) : 2026;

const envPath = path.join(userHomeDir(), ".codex/pms-connections", `${profile}.env`);
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
  "DB_READ_WRITE",
]) {
  if (process.env[key]) env[key] = process.env[key];
}

const required = ["DB_ENGINE", "DB_HOST", "DB_PORT", "DB_NAME", "DB_USER", "DB_PASSWORD"];
const missing = required.filter((key) => !env[key]);
if (missing.length) {
  console.error(`Profile ${profile} is incomplete. Missing: ${missing.join(", ")}`);
  process.exit(1);
}
if (String(env.DB_READ_WRITE ?? "0") !== "0") {
  console.error("Refusing to query production snapshot unless DB_READ_WRITE=0.");
  process.exit(1);
}
if (!["mysql", "mariadb"].includes(String(env.DB_ENGINE).toLowerCase())) {
  console.error(`Unsupported DB_ENGINE: ${env.DB_ENGINE}`);
  process.exit(1);
}

const mysql = requireFromPmsService("mysql2/promise");

const connection = await mysql.createConnection({
  host: env.DB_HOST,
  port: Number(env.DB_PORT),
  user: env.DB_USER,
  password: env.DB_PASSWORD,
  database: env.DB_NAME,
  ssl: env.DB_SSL === "1" ? {} : undefined,
});

try {
  const [rows] = await connection.query(
    `SELECT
       ko.kpi_ownership_id,
       ko.kpi_id,
       CAST(ko.position_master_id AS CHAR) AS pmid,
       CAST(ko.position_master_variant_id AS CHAR) AS pmvid,
       CAST(ko.position_nomenclature_id AS CHAR) AS pnid,
       ko.year AS ownership_year,
       ko.allocation_status,
       ko.ownership_type,
       ko.weight_approval_status,
       ko.deleted_at AS ownership_deleted_at,
       k.year AS kpi_year,
       k.type AS kpi_type,
       k.title,
       k.item_approval_status,
       k.is_active,
       k.deleted_at AS kpi_deleted_at,
       k.external_id,
       k.parent_kpi_id,
       k.perspective,
       k.polarity,
       k.monitoring_period,
       k.formula,
       k.target_unit,
       k.cascading_method,
       k.nature_of_work,
       k.kpi_ownership_type,
       k.description,
       k.rkm_code_id,
       ko.weight AS ownership_weight
     FROM kpi_ownership_v3 ko
     INNER JOIN kpi_v3 k ON k.kpi_id = ko.kpi_id
     WHERE ko.year = ?
       AND k.year = ?
       AND ko.deleted_at IS NULL
       AND k.deleted_at IS NULL
       AND k.is_active = 1
     ORDER BY ko.position_master_id ASC, ko.position_nomenclature_id ASC, k.type ASC, k.kpi_id ASC`,
    [year, year],
  );

  const snapshot = {
    generated_at: new Date().toISOString(),
    profile,
    year,
    connection: {
      host: env.DB_HOST,
      database: env.DB_NAME,
      user: env.DB_USER,
      read_only: env.DB_READ_WRITE === "0",
    },
    rows,
  };
  mkdirSync(path.dirname(outputPath), { recursive: true });
  writeFileSync(outputPath, JSON.stringify(snapshot, null, 2), "utf8");
  console.log(
    JSON.stringify(
      {
        output: outputPath,
        year,
        rows: rows.length,
        read_only: env.DB_READ_WRITE === "0",
      },
      null,
      2,
    ),
  );
} finally {
  await connection.end();
}
