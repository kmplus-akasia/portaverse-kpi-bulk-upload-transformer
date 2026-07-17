#!/usr/bin/env node
import { createRequire } from "node:module";
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { pathToFileURL } from "node:url";
import {
  buildHistoricalAssignmentQuery,
  buildHistoricalNomenclatureQuery,
  shapeHistoricalPayload,
} from "./historical_q1_reference.mjs";

const Q1_CUTOFF_DATE = "2026-03-31";
const HEAD_OFFICE_COMPANY_ID = "1";
const PRODUCTION_PROFILE = "production";

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

export function assertProductionProfile(profile) {
  if (profile !== PRODUCTION_PROFILE) {
    throw new Error("Historical Q1 exporter requires --profile production.");
  }
  return profile;
}

export async function runHistoricalReferenceExport({
  createConnection,
  connectionOptions,
  profile,
  cutoffDate,
  companyId,
}) {
  assertProductionProfile(profile);

  let connection;
  try {
    connection = await createConnection(connectionOptions);
    await connection.query("SET SESSION TRANSACTION READ ONLY");
    await connection.query("START TRANSACTION READ ONLY");

    const [assignmentRows] = await connection.query(buildHistoricalAssignmentQuery(), [
      cutoffDate,
      companyId,
    ]);
    const [nomenclatureRows] = await connection.query(buildHistoricalNomenclatureQuery(), [companyId]);
    return shapeHistoricalPayload({
      profile,
      cutoffDate,
      companyId,
      assignmentRows,
      nomenclatureRows,
    });
  } finally {
    if (connection) await connection.end();
  }
}

export async function main(args = process.argv.slice(2)) {
  const outputPath = optionValue(args, "--output", "");
  const profile = optionValue(args, "--profile", PRODUCTION_PROFILE);
  const cutoffDate = optionValue(args, "--cutoff-date", Q1_CUTOFF_DATE);
  const companyId = optionValue(args, "--company-id", HEAD_OFFICE_COMPANY_ID);

  if (!outputPath) throw new Error("Missing required --output path.");
  assertProductionProfile(profile);
  if (cutoffDate !== Q1_CUTOFF_DATE) {
    throw new Error(`Historical Q1 exporter only supports cutoff date ${Q1_CUTOFF_DATE}.`);
  }
  if (companyId !== HEAD_OFFICE_COMPANY_ID) {
    throw new Error(`Historical Q1 exporter only supports company ID ${HEAD_OFFICE_COMPANY_ID}.`);
  }

  const env = loadProfile(profile);
  const required = ["DB_ENGINE", "DB_HOST", "DB_PORT", "DB_NAME", "DB_USER", "DB_PASSWORD"];
  const missing = required.filter((key) => !env[key]);
  if (missing.length) {
    throw new Error(`Profile ${profile} is incomplete. Missing: ${missing.join(", ")}`);
  }
  if (!["mysql", "mariadb"].includes(String(env.DB_ENGINE).toLowerCase())) {
    throw new Error(`Unsupported DB_ENGINE: ${env.DB_ENGINE}`);
  }

  const servicePackage = "/Users/alfredoteja/Documents/pms-codebase/pms-service/package.json";
  const serviceRequire = createRequire(servicePackage);
  const mysql = serviceRequire("mysql2/promise");
  const payload = await runHistoricalReferenceExport({
    createConnection: mysql.createConnection.bind(mysql),
    connectionOptions: {
      host: env.DB_HOST,
      port: Number(env.DB_PORT),
      user: env.DB_USER,
      password: env.DB_PASSWORD,
      database: env.DB_NAME,
      ssl: env.DB_SSL === "1" ? {} : undefined,
    },
    profile,
    cutoffDate,
    companyId,
  });

  mkdirSync(path.dirname(outputPath), { recursive: true });
  writeFileSync(outputPath, `${JSON.stringify(payload, null, 2)}\n`, "utf8");
  console.log(`Wrote historical assignment rows: ${payload.historical_assignment_rows.length}`);
  console.log(`Wrote nomenclature rows: ${payload.nomenclature_rows.length}`);
  console.log(`Output: ${outputPath}`);
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  main().catch((error) => {
    console.error(error.message);
    process.exitCode = 1;
  });
}
