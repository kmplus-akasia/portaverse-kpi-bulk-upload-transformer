#!/usr/bin/env node
import { createRequire } from "node:module";
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import {
  buildHistoricalAssignmentQuery,
  buildHistoricalNomenclatureQuery,
  shapeHistoricalPayload,
} from "./historical_q1_reference.mjs";

const Q1_CUTOFF_DATE = "2026-03-31";
const HEAD_OFFICE_COMPANY_ID = "1";

function optionValue(args, name, fallback) {
  const index = args.indexOf(name);
  return index >= 0 && args[index + 1] ? args[index + 1] : fallback;
}

function loadProfile(profile) {
  const envPath = path.join(process.env.HOME || "", ".codex/pms-connections", `${profile}.env`);
  const env = {};

  if (existsSync(envPath)) {
    for (const line of readFileSync(envPath, "utf8").split(/\r?\n/)) {
      const trimmed = line.trim();
      if (!trimmed || trimmed.startsWith("#")) continue;
      const separator = trimmed.indexOf("=");
      if (separator < 0) continue;
      env[trimmed.slice(0, separator).trim()] = trimmed
        .slice(separator + 1)
        .trim()
        .replace(/^['"]|['"]$/g, "");
    }
  }

  for (const key of ["DB_ENGINE", "DB_HOST", "DB_PORT", "DB_NAME", "DB_USER", "DB_PASSWORD", "DB_SSL"]) {
    if (process.env[key]) env[key] = process.env[key];
  }

  return env;
}

function fail(message) {
  console.error(message);
  process.exit(1);
}

const args = process.argv.slice(2);
const outputPath = optionValue(args, "--output", "");
const profile = optionValue(args, "--profile", "production");
const cutoffDate = optionValue(args, "--cutoff-date", Q1_CUTOFF_DATE);
const companyId = optionValue(args, "--company-id", HEAD_OFFICE_COMPANY_ID);

if (!outputPath) fail("Missing required --output path.");
if (cutoffDate !== Q1_CUTOFF_DATE) {
  fail(`Historical Q1 exporter only supports cutoff date ${Q1_CUTOFF_DATE}.`);
}
if (companyId !== HEAD_OFFICE_COMPANY_ID) {
  fail(`Historical Q1 exporter only supports company ID ${HEAD_OFFICE_COMPANY_ID}.`);
}

const env = loadProfile(profile);
const required = ["DB_ENGINE", "DB_HOST", "DB_PORT", "DB_NAME", "DB_USER", "DB_PASSWORD"];
const missing = required.filter((key) => !env[key]);
if (missing.length) fail(`Profile ${profile} is incomplete. Missing: ${missing.join(", ")}`);
if (!['mysql', 'mariadb'].includes(String(env.DB_ENGINE).toLowerCase())) {
  fail(`Unsupported DB_ENGINE: ${env.DB_ENGINE}`);
}

const servicePackage = "/Users/alfredoteja/Documents/pms-codebase/pms-service/package.json";
const serviceRequire = createRequire(servicePackage);
const mysql = serviceRequire("mysql2/promise");

let connection;
try {
  connection = await mysql.createConnection({
    host: env.DB_HOST,
    port: Number(env.DB_PORT),
    user: env.DB_USER,
    password: env.DB_PASSWORD,
    database: env.DB_NAME,
    ssl: env.DB_SSL === "1" ? {} : undefined,
  });

  const [assignmentRows] = await connection.query(buildHistoricalAssignmentQuery(), [
    cutoffDate,
    companyId,
  ]);
  const [nomenclatureRows] = await connection.query(buildHistoricalNomenclatureQuery(), [companyId]);
  const payload = shapeHistoricalPayload({
    profile,
    cutoffDate,
    companyId,
    assignmentRows,
    nomenclatureRows,
  });

  mkdirSync(path.dirname(outputPath), { recursive: true });
  writeFileSync(outputPath, `${JSON.stringify(payload, null, 2)}\n`, "utf8");
  console.log(`Wrote historical assignment rows: ${assignmentRows.length}`);
  console.log(`Wrote nomenclature rows: ${nomenclatureRows.length}`);
  console.log(`Output: ${outputPath}`);
} finally {
  if (connection) await connection.end();
}
