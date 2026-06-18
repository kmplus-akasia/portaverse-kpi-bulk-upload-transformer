import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputWorkbook = "/Users/alfredoteja/Downloads/KPI_Upload_Template_Portaverse_2026.xlsx";
const referencePath = "configs/production_position_reference.json";
const outputDir = "outputs/kpi-template-master-data";
const outputWorkbook = `${outputDir}/KPI_Upload_Template_Portaverse_2026_with_master_data.xlsx`;

const colLetter = (index) => {
  let n = index;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
};

const asValue = (value) => (value === undefined ? null : value);

const writeRows = (sheet, startRow, rows) => {
  if (!rows.length) return;
  const cols = rows[0].length;
  const endRow = startRow + rows.length - 1;
  const endCol = colLetter(cols);
  sheet.getRange(`A${startRow}:${endCol}${endRow}`).values = rows;
};

const writeRowsChunked = (sheet, rows, chunkSize = 5000) => {
  for (let offset = 0; offset < rows.length; offset += chunkSize) {
    writeRows(sheet, offset + 1, rows.slice(offset, offset + chunkSize));
  }
};

const ref = JSON.parse(await fs.readFile(referencePath, "utf8"));
const input = await FileBlob.load(inputWorkbook);
const workbook = await SpreadsheetFile.importXlsx(input);

const nomenclatureSheet = workbook.worksheets.getItem("Master Data Nomenclature");
const positionMasterSheet = workbook.worksheets.getItem("Master Data Position Master");
const formSheet = workbook.worksheets.getItem("Formulir Upload KPI");
const guideSheet = workbook.worksheets.getItem("Aturan dan Panduan Pengisian");

const nomenclatureHeaders = [
  "cluster_id (PNID)",
  "mapping_row_id",
  "cluster_label",
  "position_master_id (PMID)",
  "position_name",
  "group_name",
  "type_name",
  "company_name",
  "group_master_id",
  "company_id",
  "job_class_level",
  "is_group_active",
  "is_company_active",
  "company_code",
];

const nomenclatureRows = ref.rows.map((r) => [
  asValue(r.cluster_id),
  asValue(r.id),
  asValue(r.cluster_label),
  asValue(r.position_master_id),
  asValue(r.position_name),
  asValue(r.group_name),
  asValue(r.type_name),
  asValue(r.company_name ?? r.active_company_name),
  asValue(r.group_master_id),
  asValue(r.company_id),
  asValue(r.job_class_level),
  asValue(r.is_group_active),
  asValue(r.is_company_active),
  asValue(r.active_company_code),
]);

const positionMasterHeaders = [
  "position_master_id (PMID)",
  "position_name",
  "job_class_level",
  "company_name",
  "group_name",
  "group_master_id",
  "company_id",
  "company_code",
  "position_master_type_id",
  "is_position_active",
  "is_position_organization_active",
  "is_group_active",
  "is_company_active",
  "position_start_date",
  "position_end_date",
];

const positionMasterRows = [...ref.position_master_rows]
  .sort((a, b) => {
    const id = Number(a.position_master_id) - Number(b.position_master_id);
    if (id) return id;
    for (const key of [
      "is_position_active",
      "is_position_organization_active",
      "is_group_active",
      "is_company_active",
    ]) {
      const diff = Number(b[key] || 0) - Number(a[key] || 0);
      if (diff) return diff;
    }
    return Number(a.group_master_id || 0) - Number(b.group_master_id || 0);
  })
  .map((r) => [
    asValue(r.position_master_id),
    asValue(r.position_name),
    asValue(r.job_class_level),
    asValue(r.company_name),
    asValue(r.group_name),
    asValue(r.group_master_id),
    asValue(r.company_id),
    asValue(r.company_code),
    asValue(r.position_master_type_id),
    asValue(r.is_position_active),
    asValue(r.is_position_organization_active),
    asValue(r.is_group_active),
    asValue(r.is_company_active),
    asValue(r.position_start_date),
    asValue(r.position_end_date),
  ]);

writeRowsChunked(nomenclatureSheet, [nomenclatureHeaders, ...nomenclatureRows]);
writeRowsChunked(positionMasterSheet, [positionMasterHeaders, ...positionMasterRows]);

const formulaRows = [];
for (let row = 5; row <= 504; row += 1) {
  formulaRows.push([
    `=IF(B${row}<>"","",C${row})`,
    `=IF(B${row}<>"",IFERROR(INDEX('Master Data Nomenclature'!$C:$C,MATCH(B${row},'Master Data Nomenclature'!$A:$A,0)),"Tidak ditemukan"),IF(C${row}<>"",IFERROR(INDEX('Master Data Position Master'!$B:$B,MATCH(C${row},'Master Data Position Master'!$A:$A,0)),"Tidak ditemukan"),""))`,
    `=IF(B${row}<>"",IFERROR(INDEX('Master Data Nomenclature'!$F:$F,MATCH(B${row},'Master Data Nomenclature'!$A:$A,0)),"Tidak ditemukan"),IF(C${row}<>"",IFERROR(INDEX('Master Data Position Master'!$E:$E,MATCH(C${row},'Master Data Position Master'!$A:$A,0)),"Tidak ditemukan"),""))`,
    `=IF(B${row}<>"",IFERROR(INDEX('Master Data Nomenclature'!$H:$H,MATCH(B${row},'Master Data Nomenclature'!$A:$A,0)),"Tidak ditemukan"),IF(C${row}<>"",IFERROR(INDEX('Master Data Position Master'!$D:$D,MATCH(C${row},'Master Data Position Master'!$A:$A,0)),"Tidak ditemukan"),""))`,
  ]);
}
formSheet.getRange("D5:G504").formulas = formulaRows;

guideSheet.getRange("B16").values = [
  [
    "- Catat nilai pada kolom cluster_id (= PNID) di sheet Nomenclature, atau position_master_id (= PMID) di sheet Position Master.",
  ],
];
guideSheet.getRange("B18").values = [
  [
    "- Relasi: cluster_id pada Nomenclature adalah PNID; position_master_id menautkan Nomenclature ke Position Master. Untuk PNID, Company/Group diambil dari baris Nomenclature yang sama.",
  ],
];
guideSheet.getRange("E35").values = [
  ["Terisi otomatis dari Master Data via PNID; jika PNID kosong, fallback memakai PMID."],
];

await fs.mkdir(outputDir, { recursive: true });
const exported = await SpreadsheetFile.exportXlsx(workbook);
await exported.save(outputWorkbook);

console.log(
  JSON.stringify(
    {
      outputWorkbook,
      source: ref.source,
      nomenclatureRows: ref.rows.length,
      positionMasterRows: ref.position_master_rows.length,
      formFormulaRowsUpdated: 500,
    },
    null,
    2,
  ),
);
