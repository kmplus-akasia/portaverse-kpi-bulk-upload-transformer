import test from "node:test";
import assert from "node:assert/strict";
import {
  buildHistoricalAssignmentQuery,
  buildHistoricalNomenclatureQuery,
  classifyAssignment,
  shapeHistoricalPayload,
} from "../scripts/historical_q1_reference.mjs";
import {
  assertProductionProfile,
  runHistoricalReferenceExport,
} from "../scripts/export_historical_q1_position_reference.mjs";

test("historical query is parameterized and anchored to the requested cutoff", () => {
  const sql = buildHistoricalAssignmentQuery();

  assert.match(sql, /DATE\(tepms\.end_date\) = \?/);
  assert.match(sql, /tepms\.deletedAt IS NULL/);
  assert.match(sql, /tpmos\.organization_master_id/);
  assert.equal((sql.match(/\?/g) || []).length, 2);
});

test("classifies primary and secondary historical assignments", () => {
  assert.equal(classifyAssignment({ lakhar_id: null, job_sharing_id: null }), "PRIMARY");
  assert.equal(classifyAssignment({ lakhar_id: 2, job_sharing_id: null }), "LAKHAR");
  assert.equal(classifyAssignment({ lakhar_id: null, job_sharing_id: 9 }), "JOB_SHARING");
});

test("nomenclature query scopes directly to the requested company", () => {
  const sql = buildHistoricalNomenclatureQuery();

  assert.match(sql, /pnm\.company_id = \?/);
  assert.doesNotMatch(sql, /tgm\.company_id = \? OR tgm\.group_master_id IS NULL/);
});

test("production-only runner enables a read-only session before selecting", async () => {
  assert.equal(assertProductionProfile("production"), "production");
  assert.throws(
    () => assertProductionProfile("staging"),
    /requires --profile production/,
  );

  const calls = [];
  let ended = false;
  const connection = {
    async query(sql, params) {
      calls.push({ sql, params });
      if (sql === "SET SESSION TRANSACTION READ ONLY") return [[]];
      if (sql.includes("tb_employee_position_master_sync")) return [[{ employee_number: "100" }]];
      return [[{ position_master_id: 501 }]];
    },
    async end() {
      ended = true;
    },
  };

  const payload = await runHistoricalReferenceExport({
    createConnection: async () => connection,
    connectionOptions: { host: "database.example" },
    profile: "production",
    cutoffDate: "2026-03-31",
    companyId: "1",
  });

  const firstSelectIndex = calls.findIndex(({ sql }) => sql.includes("SELECT"));
  assert.equal(calls[0].sql, "SET SESSION TRANSACTION READ ONLY");
  assert.ok(firstSelectIndex > 0);
  assert.deepEqual(calls[firstSelectIndex].params, ["2026-03-31", "1"]);
  assert.deepEqual(calls[firstSelectIndex + 1].params, ["1"]);
  assert.equal(payload.historical_assignment_rows[0].assignment_type, "PRIMARY");
  assert.equal(ended, true);
});

test("shapes historical payload with assignment evidence and source scope", () => {
  const payload = shapeHistoricalPayload({
    profile: "production",
    cutoffDate: "2026-03-31",
    companyId: "1",
    assignmentRows: [{ employee_number: "100", lakhar_id: null, job_sharing_id: null }],
    nomenclatureRows: [{ cluster_id: 76, position_master_id: 501 }],
  });

  assert.deepEqual(payload.source, {
    profile: "production",
    cutoff_date: "2026-03-31",
    company_id: "1",
    read_only: true,
  });
  assert.deepEqual(payload.historical_assignment_rows, [
    { employee_number: "100", lakhar_id: null, job_sharing_id: null, assignment_type: "PRIMARY" },
  ]);
  assert.deepEqual(payload.nomenclature_rows, [{ cluster_id: 76, position_master_id: 501 }]);
});
