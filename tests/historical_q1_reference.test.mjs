import test from "node:test";
import assert from "node:assert/strict";
import {
  buildHistoricalAssignmentQuery,
  classifyAssignment,
  shapeHistoricalPayload,
} from "../scripts/historical_q1_reference.mjs";

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
