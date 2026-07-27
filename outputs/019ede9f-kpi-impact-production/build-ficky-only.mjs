import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "/Users/alfredoteja/Documents/portaverse-kpi-bulk-upload-transformer/outputs/019ede9f-kpi-impact-production";
const sourcePath = path.join(
  outputDir,
  "KPI_Upload_Production_HO_Missing_Impact_20260619_SNAPSHOT.xlsx",
);
const outputPath = path.join(outputDir, "KPI_Upload_Ficky_Alkarim_Impact_20260623.xlsx");

const sourceWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(sourcePath));
const sourceSheet = sourceWorkbook.worksheets.getItem("KPI Template");
const used = sourceSheet.getUsedRange(true).values;
const headers = used[0];
const fickyRows = used
  .slice(1)
  .filter((row) => Number(row[22]) === 11542)
  .map((row, index) => {
    const next = [...row];
    next[0] = index + 1;
    return next;
  });

if (fickyRows.length !== 10) {
  throw new Error(`Expected ten Ficky rows for PNID 11542, found ${fickyRows.length}.`);
}
if (
  fickyRows.some(
    (row) => row[7] !== "IMPACT" || row[4] || row[5] || Number(row[22]) !== 11542,
  )
) {
  throw new Error("Ficky rows must be PNID 11542 IMPACT-only rows.");
}
if (new Set(fickyRows.map((row) => String(row[10]))).size !== 10) {
  throw new Error("Ficky rows must contain ten distinct KPI IMPACT titles.");
}

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("KPI Template");
sheet.getRange("A1:X1").values = [headers];
sheet.getRange("A2:X11").values = fickyRows;
sheet.freezePanes.freezeRows(1);

sheet.getRange("A1:X1").format = {
  fill: "#1F4E78",
  font: { bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
  rowHeight: 38,
};
sheet.getRange("A2:X11").format = {
  fill: "#FFFFFF",
  font: { color: "#1F2937", size: 10 },
  verticalAlignment: "top",
  wrapText: true,
  rowHeight: 34,
};

const widths = {
  A: 10, B: 34, C: 30, D: 38, E: 18, F: 18, G: 22, H: 14,
  I: 14, J: 24, K: 48, L: 34, M: 14, N: 14, O: 16, P: 58,
  Q: 14, R: 16, S: 24, T: 18, U: 18, V: 18, W: 20, X: 16,
};
for (const [column, width] of Object.entries(widths)) {
  sheet.getRange(`${column}1:${column}11`).format.columnWidth = width;
}
sheet.getRange("E2:F11").format.numberFormat = "0";
sheet.getRange("Q2:Q11").format.numberFormat = "0";
sheet.getRange("W2:W11").format.numberFormat = "0";

const errorCheck = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?",
  options: { useRegex: true, maxResults: 300 },
  summary: "Ficky-only formula error scan",
});
const preview = await workbook.render({
  sheetName: "KPI Template",
  range: "A1:X11",
  scale: 1,
  format: "png",
});
await fs.writeFile(
  path.join(outputDir, "preview-ficky-only.png"),
  new Uint8Array(await preview.arrayBuffer()),
);

const exported = await SpreadsheetFile.exportXlsx(workbook);
await exported.save(outputPath);

console.log(
  JSON.stringify({
    outputPath,
    uploadRows: fickyRows.length,
    pnid: 11542,
    formulaErrors: errorCheck.ndjson,
  }),
);
