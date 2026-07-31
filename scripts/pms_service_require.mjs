import { createRequire } from "node:module";
import { existsSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const scriptsDir = path.dirname(fileURLToPath(import.meta.url));
const kpiRepoRoot = path.resolve(scriptsDir, "..");

export function userHomeDir() {
  return process.env.HOME || process.env.USERPROFILE || "";
}

export function resolvePmsServicePackage() {
  const candidates = [];

  if (process.env.PMS_SERVICE_ROOT) {
    candidates.push(path.join(process.env.PMS_SERVICE_ROOT, "package.json"));
  }

  candidates.push(
    path.join(kpiRepoRoot, "..", "pms-codebase", "pms-service", "package.json"),
    path.join(kpiRepoRoot, "..", "pms-service", "package.json"),
  );

  for (const candidate of candidates) {
    if (existsSync(candidate)) {
      return candidate;
    }
  }

  throw new Error(
    "pms-service package.json not found. Clone pms-service under pms-codebase or set PMS_SERVICE_ROOT.",
  );
}

export function requireFromPmsService(specifier) {
  return createRequire(resolvePmsServicePackage())(specifier);
}
