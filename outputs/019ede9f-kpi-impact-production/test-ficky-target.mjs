import assert from "node:assert/strict";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const workbook = await SpreadsheetFile.importXlsx(
  await FileBlob.load(
    "KPI_Upload_Production_HO_Missing_Impact_20260619_SNAPSHOT.xlsx",
  ),
);
const sheet = workbook.worksheets.getItem("KPI Template");
const used = sheet.getUsedRange(true).values;
const rows = used.slice(1);
const fickyRows = rows.filter((row) => Number(row[22]) === 11435);
const otherDataScientistRows = rows.filter((row) => Number(row[22]) === 12256);

assert.equal(
  fickyRows.length,
  10,
  "PNID 11435 for Ficky Alkarim / Data Scientist must have ten KPI IMPACT rows",
);
assert.equal(
  new Set(fickyRows.map((row) => row[10])).size,
  10,
  "PNID 11435 must contain ten distinct KPI IMPACT titles",
);
assert.ok(
  fickyRows.every(
    (row) =>
      row[7] === "IMPACT" &&
      !row[4] &&
      !row[5] &&
      Number(row[22]) === 11435,
  ),
  "Ficky rows must use PNID 11435 only",
);
assert.equal(
  otherDataScientistRows.length,
  0,
  "PNID 12256 is a different Data Scientist position and must not be added",
);

console.log("PASS: Ficky Alkarim target is present and narrowly scoped.");
