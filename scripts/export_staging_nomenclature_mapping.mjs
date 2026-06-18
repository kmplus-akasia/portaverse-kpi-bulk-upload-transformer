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
    : "configs/staging_nomenclature_mapping.json";
const profile =
  profileIndex >= 0 && args[profileIndex + 1] ? args[profileIndex + 1] : "staging";

const envPath = path.join(
  process.env.HOME,
  ".codex/pms-connections",
  `${profile}.env`,
);

if (!existsSync(envPath)) {
  console.error(`Missing profile: ${envPath}`);
  process.exit(1);
}

const env = {};
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
  const [rows] = await connection.query(
    `SELECT cluster_id,
            cluster_label,
            position_master_id,
            position_name,
            group_name,
            type_name
       FROM position_nomenclature_mapping
      WHERE position_master_id IS NOT NULL
      ORDER BY cluster_id ASC, position_master_id ASC`,
  );
  const [positionMasterRows] = await connection.query(
    `SELECT position_master_id,
            name AS position_name,
            position_master_type_id
       FROM tb_position_master_v2
      WHERE deletedAt IS NULL
        AND position_master_id IS NOT NULL
        AND name IS NOT NULL
      ORDER BY position_master_id ASC`,
  );
  const payload = {
    source: {
      profile,
      tables: ["position_nomenclature_mapping", "tb_position_master_v2"],
      query: "read-only SELECT mapping rows plus active position master names",
    },
    exported_at: new Date().toISOString(),
    rows,
    position_master_rows: positionMasterRows,
  };
  mkdirSync(path.dirname(outputPath), { recursive: true });
  writeFileSync(outputPath, `${JSON.stringify(payload, null, 2)}\n`, "utf8");
  console.log(`Wrote mapping rows: ${rows.length}`);
  console.log(`Wrote position master rows: ${positionMasterRows.length}`);
  console.log(`Output: ${outputPath}`);
} finally {
  await connection.end();
}
