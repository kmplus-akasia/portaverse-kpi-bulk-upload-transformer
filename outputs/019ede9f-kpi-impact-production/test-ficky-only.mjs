import assert from "node:assert/strict";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const workbook = await SpreadsheetFile.importXlsx(
  await FileBlob.load("KPI_Upload_Ficky_Alkarim_Impact_20260623.xlsx"),
);

const sheets = workbook.worksheets.items;
assert.equal(sheets.length, 1, "Ficky-only workbook must contain exactly one sheet");

const sheet = workbook.worksheets.getItem("KPI Template");
const used = sheet.getUsedRange(true).values;
const rows = used.slice(1);

assert.equal(rows.length, 10, "Ficky-only workbook must contain ten KPI IMPACT rows");
assert.equal(
  new Set(rows.map((row) => row[10])).size,
  10,
  "Ficky-only workbook must contain ten distinct KPI IMPACT titles",
);
assert.ok(
  rows.every(
    (row) =>
      row[7] === "IMPACT" &&
      !row[4] &&
      !row[5] &&
      Number(row[22]) === 11542,
  ),
  "Every Ficky-only row must use PNID 11542 and KPI Type IMPACT",
);
assert.equal(
  rows.filter((row) => Number(row[22]) === 11435).length,
  0,
  "Ficky-only workbook must not contain stale PNID 11435",
);
assert.equal(
  rows.filter((row) => Number(row[22]) === 12256).length,
  0,
  "Ficky-only workbook must not contain other Data Scientist PNID 12256",
);

console.log("PASS: Ficky-only upload workbook is narrowly scoped.");
