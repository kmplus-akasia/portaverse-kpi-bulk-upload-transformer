import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const root = "/Users/alfredoteja/Documents/portaverse-kpi-bulk-upload-transformer";
const outputDir = path.join(root, "outputs/019ede9f-kpi-impact-production");
const uploadDir = path.join(root, "output/kpi_upload_final_20260618/upload-ready");
const outputPath = path.join(
  outputDir,
  "KPI_Upload_Production_HO_Missing_Impact_20260619_SNAPSHOT.xlsx",
);
const EXPECTED_HEADERS = [
  "IDKPI", "Group", "Direktorat", "Posisi",
  "Position Master ID (Required)",
  "Position Master Variant ID (Optional)",
  "BSC Perspective", "KPI Type", "Parent KPI ID", "Parent KPI Title",
  "Title", "Description", "Unit", "Polarity", "Period", "Formula",
  "Weight (%)", "Cascading", "Nature Of Work (KAI Only)",
  "External ID (PKPI)", "System KPI ID", "Ownership Type",
  "Position Nomenklatur ID", "RKM Code ID",
];

const TARGETS_TSV = String.raw`
PMID	67		Department Head Hubungan Lembaga dan Investor	Department Hubungan Lembaga dan Investor	Missing	already_in_generated_upload
PMID	571		Department Head Manajemen Kearsipan, HRIS dan Administrasi SDM	Department Manajemen Kearsipan dan HRIS	Missing	already_in_generated_upload
PMID	458		Department Head Perencanaan SDM dan Kebijakan Remunerasi	Department Perencanaan SDM dan Kebijakan Remunerasi	Missing	already_in_generated_upload
PMID	797		Department Head Riset dan Inovasi	Department Riset dan Inovasi	Missing	already_in_generated_upload
PMID	599		Department Head SSC SDM dan Pengadaan	Department SSC	Missing	already_in_generated_upload
PMID	94		Department Head Tanggung Jawab Sosial dan Lingkungan	Department Tanggung Jawab Sosial dan Lingkungan	Missing	already_in_generated_upload
PMID	23047		Deputi Pimpro Bidang Administrasi Proyek Investasi BMTH	Bidang Administrasi Proyek Bali Maritime Tourism Hub (BMTH)	Missing	already_in_generated_upload
PMID	23076		Deputi Pimpro Bidang Administrasi Proyek Investasi JICT KOJA	Bidang Administrasi Proyek JICT Koja	Missing	already_in_generated_upload
PMID	23050		Deputi Pimpro Bidang Konstruksi 1 Proyek Investasi BMTH	Bidang Konstruksi 1 Proyek Bali Maritime Tourism Hub (BMTH)	Missing	already_in_generated_upload
PMID	23079		Deputi Pimpro Bidang Konstruksi 1 Proyek Investasi JICT KOJA	Bidang Konstruksi 1 Proyek JICT Koja	Missing	already_in_generated_upload
PMID	23053		Deputi Pimpro Bidang Konstruksi 2 Proyek Investasi BMTH	Bidang Konstruksi 2 Proyek Bali Maritime Tourism Hub (BMTH)	Missing	already_in_generated_upload
PMID	23082		Deputi Pimpro Bidang Konstruksi 2 Proyek Investasi JICT KOJA	Bidang Konstruksi 2 Proyek JICT Koja	Missing	already_in_generated_upload
PMID	23056		Deputi Pimpro Bidang Konstruksi 3 Proyek Investasi BMTH	Bidang Konstruksi 3 Proyek Bali Maritime Tourism Hub (BMTH)	Missing	already_in_generated_upload
PMID	23085		Deputi Pimpro Bidang Konstruksi 3 Proyek Investasi JICT KOJA	Bidang Konstruksi 3 Proyek JICT Koja	Missing	already_in_generated_upload
PMID	23073		Deputi Pimpro Bidang Perencanaan Proyek Investasi JICT KOJA	Bidang Perencanaan Proyek JICT Koja	Missing	already_in_generated_upload
PMID	48		Manager ADM dan Hubungan Antar Lembaga Bidang Pengawasan	Unit Pendukung ADM dan Hubungan Antar Lembaga Bidang Pengawasan	Missing	already_in_generated_upload
PMID	918		Manager Persetujuan, Pemantauan, dan Pengelolaan Lingkungan	Unit Pendukung Persetujuan, Pemantauan, dan Pengelolaan Lingkungan	Missing	already_in_generated_upload
PMID	23043		Pimpinan Proyek Investasi Bali Maritime Tourism Hub (BMTH)	Proyek Bali Maritime Tourism Hub (BMTH)	Missing	already_in_generated_upload
PMID	23072		Pimpinan Proyek Investasi JICT KOJA	Proyek JICT Koja	Missing	already_in_generated_upload
PNID	89		Officer Compensation Dan Benefit	Kelompok Kerja Paytoll dan Comben 2	Missing	already_in_generated_upload
PNID	92		Officer Data Management	Kelompok Kerja Employee Service 2	Missing	already_in_generated_upload
PNID	88		Officer Payroll	Kelompok Kerja Paytoll dan Comben 2	Missing	already_in_generated_upload
PNID	86		Officer Payroll	Kelompok Kerja Paytoll dan Comben 1	Missing	already_in_generated_upload
PNID	45		Officer Perencanaan Keuangan	Unit Pendukung Anggaran	Missing	already_in_generated_upload
PNID	47		Officer Strategi Perencanaan Keuangan	Unit Pendukung Anggaran	Missing	already_in_generated_upload
PNID	84		Officer Talent Services	Kelompok Kerja Talent Services	Missing	already_in_generated_upload
PNID	66		Staff Utama Direktur Sdm Dan Umum	Direktorat SDM dan Umum	Missing	already_in_generated_upload
PMID	23092		Deputi Pimpro Bidang Administrasi Proyek Investasi Pelabuhan Batang	Bidang Administrasi Proyek Batang	Missing	project_position_not_in_latest_upload
PMID	23098		Deputi Pimpro Bidang Konstruksi 2 Proyek Investasi Pelabuhan Batang	Bidang Konstruksi 2 Proyek Batang	Missing	project_position_not_in_latest_upload
PMID	23089		Deputi Pimpro Bidang Perencanaan Proyek Investasi Pelabuhan Batang	Bidang Perencanaan Proyek Batang	Missing	project_position_not_in_latest_upload
PMID	23088		Pimpinan Proyek Investasi Pelabuhan Batang	Proyek Batang	Missing	project_position_not_in_latest_upload
PMID	21070		Direktur Keuangan	Direktorat Keuangan dan Manajemen Risiko	Missing	not_found_in_generated_upload
PMID	35435		Direktur Komersial	Direktorat Komersial	Missing	not_found_in_generated_upload
PMID	35433		Direktur Manajemen Risiko	Direktorat Manajemen Risiko	Missing	not_found_in_generated_upload
PMID	35432		Direktur Operasi	Direktorat Operasi	Missing	not_found_in_generated_upload
PMID	35434		Direktur Pengembangan Usaha	Direktorat Pengembangan Usaha	Missing	not_found_in_generated_upload
PMID	21071		Direktur SDM dan Umum	Direktorat SDM dan Umum	Missing	not_found_in_generated_upload
PMID	35436		Direktur Teknik	Direktorat Teknik	Missing	not_found_in_generated_upload
PMID	21068		Direktur Utama	Direktorat Utama	Missing	not_found_in_generated_upload
PMID	35367		Group Head Strategi Korporasi dan Pengembangan Bisnis	Group Strategi Korporasi dan Pengembangan Bisnis	Missing	not_found_in_generated_upload
PMID	21069		Wakil Direktur Utama	Direktorat Wakil Direktur Utama	Missing	not_found_in_generated_upload
PNID	7911		Officer Bidang Administrasi Proyek Investasi Jict Koja	Bidang Administrasi Proyek JICT Koja	Missing	not_found_in_generated_upload
PNID	7916		Officer Bidang Administrasi Proyek Investasi Pelabuhan Batang	Bidang Administrasi Proyek Batang	Missing	not_found_in_generated_upload
PNID	7903		Officer Bidang Konstruksi Proyek Investasi Bmth	Bidang Konstruksi 2 Proyek Bali Maritime Tourism Hub (BMTH)	Missing	not_found_in_generated_upload
PNID	7902		Officer Bidang Konstruksi Proyek Investasi Bmth	Bidang Konstruksi 1 Proyek Bali Maritime Tourism Hub (BMTH)	Missing	not_found_in_generated_upload
PNID	7914		Officer Bidang Konstruksi Proyek Investasi Jict Koja	Bidang Konstruksi 3 Proyek JICT Koja	Missing	not_found_in_generated_upload
PNID	7913		Officer Bidang Konstruksi Proyek Investasi Jict Koja	Bidang Konstruksi 2 Proyek JICT Koja	Missing	not_found_in_generated_upload
PNID	7900		Officer Bidang Perencanaan Proyek Investasi Bmth	Bidang Perencanaan Proyek Bali Maritime Tourism Hub (BMTH)	Missing	not_found_in_generated_upload
PNID	7915		Officer Bidang Perencanaan Proyek Investasi Pelabuhan Batang	Bidang Perencanaan Proyek Batang	Missing	not_found_in_generated_upload
PNID	90		Officer Data Management	Kelompok Kerja Employee Service 1	Missing	not_found_in_generated_upload
PNID	94		Officer Data Management	Kelompok Kerja Employee Service 3	Missing	not_found_in_generated_upload
PNID	9730		Officer Implementasi Dan Pelaporan Corporate Sustainability	Unit Pendukung Implementasi dan Pelaporan	Missing	not_found_in_generated_upload
PNID	8807		Officer Integrasi Proyek Implementasi Single Erp	Unit Pendukung Integrasi	Missing	not_found_in_generated_upload
PNID	155		Officer Keselamatan Dan Kesehatan Kerja	Department Keselamatan dan Kesehatan Kerja	Missing	not_found_in_generated_upload
PNID	9725		Officer Pengelolaan Pelanggan	Unit Pendukung Pengelolaan Pelanggan	Missing	not_found_in_generated_upload
PNID	8804		Officer Pengendalian Kinerja Proyek Investasi	Unit Pendukung Pengendalian Kinerja Proyek Investasi 2	Missing	not_found_in_generated_upload
PNID	8802		Officer Pengendalian Proyek Kinerja Investasi	Unit Pendukung Pengendalian Kinerja Proyek Investasi 1	Missing	not_found_in_generated_upload
PNID	122		Officer Pengolahan Dan Pelaporan Data	Unit Pendukung Pengolahaan dan Pelaporan Data	Missing	not_found_in_generated_upload
PNID	99		Officer Source-To-Contract	Kelompok Kerja Source-to-Contract 1	Missing	not_found_in_generated_upload
PNID	100		Officer Source-To-Contract	Kelompok Kerja Source-to-Contract 2	Missing	not_found_in_generated_upload
PNID	104		Officer Strategi Dan Perencanaan Pengadaan	Department Strategi dan Perencanaan Pengadaaan	Missing	not_found_in_generated_upload
PNID	119		Officer Strategi Dan Tata Kelola Ti	Department Strategi dan Tata Kelola TI	Missing	not_found_in_generated_upload
PNID	95		Officer Travel Management	Kelompok Kerja Employee Service 3	Missing	not_found_in_generated_upload
PNID	93		Officer Travel Management	Kelompok Kerja Employee Service 2	Missing	not_found_in_generated_upload
PNID	91		Officer Travel Management	Kelompok Kerja Employee Service 1	Missing	not_found_in_generated_upload
PNID	9489		Penugasan Sebagai Direktur Pada Grup Bisnis Di Luar Pt Pelabuhan Indonesia (Persero)	PT Pelabuhan Indonesia (Persero)	Missing	not_found_in_generated_upload
PNID	13		Personal Assistant Direksi	Department Manajemen Kesekretariatan	Missing	not_found_in_generated_upload
PNID	25		Staff Group Hukum	Group Hukum	Missing	not_found_in_generated_upload
PNID	9727		Staff Group K3 dan Sistem Manajemen	Group K3 dan Sistem Manajemen	Missing	not_found_in_generated_upload
PNID	9698		Staff Group Layanan SDM	Group Layanan SDM	Missing	not_found_in_generated_upload
PNID	8835		Staff Group Manajemen Risiko, Tata Kelola, dan Kepatuhan	Group Manajemen Risiko, Tata Kelola, dan Kepatuhan	Missing	not_found_in_generated_upload
PNID	9696		Staff Group Pengelolaan SDM	Group Pengelolaan SDM	Missing	not_found_in_generated_upload
PNID	9715		Staff Group Pengendalian Proyek	Group Pengendalian Proyek	Missing	not_found_in_generated_upload
PNID	9713		Staff Group Peralatan Pelabuhan	Group Peralatan Pelabuhan	Missing	not_found_in_generated_upload
PNID	9689		Staff Group Perencanaan dan Performa Keuangan	Group Perencanaan dan Performa Keuangan	Missing	not_found_in_generated_upload
PNID	9694		Staff Group Strategi SDM	Group Strategi SDM	Missing	not_found_in_generated_upload
PNID	139		Staff Utama Direktur Pengelola	Direktorat Pengelola	Missing	not_found_in_generated_upload
PNID	11541		Sustainability Analyst - Unit Pendukung Implementasi Dan Pelaporan	Unit Pendukung Implementasi dan Pelaporan	Missing	not_found_in_generated_upload
PNID	11542		Data Scientist	Department  Monitoring & Evaluasi Klaster Ekspansi Korporasi	Missing	user_reported_missing_after_snapshot
PMID_VARIANT	348	39813	Officer Pertambahan Nilai	Unit Pendukung Pertambahan Nilai	Partial	not_found_in_generated_upload
`;

const targets = TARGETS_TSV.trim()
  .split("\n")
  .map((line) => {
    const [kind, id, pmvid, label, group, availabilityStatus, reasonCode] =
      line.split("\t");
    return { kind, id, pmvid, label, group, availabilityStatus, reasonCode };
  });

async function firstUploadWorkbook() {
  const names = (await fs.readdir(uploadDir))
    .filter((name) => name.endsWith(".xlsx"))
    .sort();
  if (!names.length) throw new Error("No validated upload workbook found.");
  return path.join(uploadDir, names[0]);
}

async function importWorkbook(filePath) {
  return SpreadsheetFile.importXlsx(await FileBlob.load(filePath));
}

function targetIdentity(target) {
  return target.kind === "PMID_VARIANT"
    ? `PMID:${target.id}:PMVID:${target.pmvid}`
    : `${target.kind}:${target.id}`;
}

function buildDirectorateMap(config) {
  const values = new Map();
  for (const item of config.positions ?? []) {
    const keys = [];
    if (item.position_master_id) keys.push(`PMID:${item.position_master_id}`);
    if (item.position_nomenclature_id) {
      keys.push(`PNID:${item.position_nomenclature_id}`);
    }
    for (const key of keys) {
      if (!values.has(key)) values.set(key, new Set());
      if (item.directorate_name) values.get(key).add(item.directorate_name);
    }
  }
  return new Map(
    [...values].map(([key, names]) => [
      key,
      names.size === 1 ? [...names][0] : "",
    ]),
  );
}

function resolveDirectorate(target, directorates) {
  if (/^Direktorat\b/i.test(target.group)) return target.group;
  const lookupKey =
    target.kind === "PMID_VARIANT" ? "PNID:54" : `${target.kind}:${target.id}`;
  return directorates.get(lookupKey) ?? "";
}

function validateTargets() {
  if (targets.length !== 80) {
    throw new Error(`Expected 80 targets, found ${targets.length}.`);
  }
  const identities = targets.map(targetIdentity);
  if (new Set(identities).size !== identities.length) {
    throw new Error("Duplicate target identity found.");
  }
  const structural = targets.filter((t) => t.kind === "PMID").length;
  const pnid = targets.filter((t) => t.kind === "PNID").length;
  const exactVariant = targets.filter((t) => t.kind === "PMID_VARIANT").length;
  if (structural !== 33 || pnid !== 46 || exactVariant !== 1) {
    throw new Error(
      `Unexpected target mix PMID=${structural} PNID=${pnid} exact=${exactVariant}.`,
    );
  }
}

function validateImpactRows(rows) {
  if (rows.length !== 10 || rows.some((row) => row[7] !== "IMPACT")) {
    throw new Error("Source does not contain exactly ten IMPACT rows.");
  }
  const titles = rows.map((row) => String(row[10] ?? "").trim());
  if (titles.some((title) => !title) || new Set(titles).size !== 10) {
    throw new Error("Impact titles are blank or duplicated.");
  }
  const weightTotal = rows.reduce((sum, row) => sum + Number(row[16] ?? 0), 0);
  if (weightTotal !== 100) {
    throw new Error(`Impact weight total must be 100, found ${weightTotal}.`);
  }
  return { titles, weightTotal };
}

function buildRows(impactRows, directorates) {
  const rows = [];
  let idKpi = 1;
  for (const target of targets) {
    for (const impact of impactRows) {
      const row = [...impact];
      row[0] = idKpi++;
      row[1] = target.group;
      row[2] = resolveDirectorate(target, directorates);
      row[3] = target.label;
      row[4] = target.kind === "PNID" ? null : Number(target.id);
      row[5] = target.pmvid ? Number(target.pmvid) : null;
      row[7] = "IMPACT";
      row[8] = null;
      row[9] = "#N/A";
      row[17] = null;
      row[18] = null;
      row[19] = null;
      row[20] = null;
      row[21] = null;
      row[22] = target.kind === "PNID" ? Number(target.id) : null;
      row[23] = null;
      rows.push(row);
    }
  }
  return rows;
}

function applyLayout(sheet, lastRow) {
  const header = sheet.getRange("A1:X1");
  header.format = {
    fill: "#1F4E78",
    font: { bold: true, color: "#FFFFFF" },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    rowHeight: 38,
  };
  const body = sheet.getRange(`A2:X${lastRow}`);
  body.conditionalFormats.deleteAll();
  body.format = {
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
    sheet.getRange(`${column}1:${column}${lastRow}`).format.columnWidth = width;
  }
  sheet.getRange(`E2:F${lastRow}`).format.numberFormat = "0";
  sheet.getRange(`Q2:Q${lastRow}`).format.numberFormat = "0";
  sheet.getRange(`W2:W${lastRow}`).format.numberFormat = "0";
}

async function build() {
  validateTargets();
  const sourcePath = await firstUploadWorkbook();
  const sourceWorkbook = await importWorkbook(sourcePath);
  const sourceSheet = sourceWorkbook.worksheets.getItem("KPI Template");
  const impactRows = sourceSheet.getRange("A2:X11").values.map((row) => [...row]);
  const impactContract = validateImpactRows(impactRows);
  const config = JSON.parse(
    await fs.readFile(
      path.join(root, "output/kpi_upload_final_20260618/corrected_positions.json"),
      "utf8",
    ),
  );
  const rows = buildRows(impactRows, buildDirectorateMap(config));
  if (rows.length !== 800) {
    throw new Error(`Expected 800 upload rows, found ${rows.length}.`);
  }

  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add("KPI Template");
  const lastRow = rows.length + 1;
  sheet.getRange("A1:X1").values = [EXPECTED_HEADERS];
  sheet.getRange(`A2:X${lastRow}`).values = rows;
  sheet.freezePanes.freezeRows(1);
  applyLayout(sheet, lastRow);

  const keyCheck = await workbook.inspect({
    kind: "table",
    range: "KPI Template!A1:X15",
    include: "values,formulas",
    tableMaxRows: 15,
    tableMaxCols: 24,
  });
  const tailCheck = await workbook.inspect({
    kind: "table",
    range: `KPI Template!A${lastRow - 10}:X${lastRow}`,
    include: "values,formulas",
    tableMaxRows: 11,
    tableMaxCols: 24,
  });
  const errorCheck = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?",
    options: { useRegex: true, maxResults: 300 },
    summary: "final formula error scan",
  });

  const firstPreview = await workbook.render({
    sheetName: "KPI Template",
    range: "A1:X25",
    scale: 1,
    format: "png",
  });
  await fs.writeFile(
    path.join(outputDir, "preview-first.png"),
    new Uint8Array(await firstPreview.arrayBuffer()),
  );
  const lastPreview = await workbook.render({
    sheetName: "KPI Template",
    range: `A${lastRow - 20}:X${lastRow}`,
    scale: 1,
    format: "png",
  });
  await fs.writeFile(
    path.join(outputDir, "preview-last.png"),
    new Uint8Array(await lastPreview.arrayBuffer()),
  );

  const exported = await SpreadsheetFile.exportXlsx(workbook);
  await exported.save(outputPath);

  const reopened = await importWorkbook(outputPath);
  const reopenedSheet = reopened.worksheets.getItem("KPI Template");
  const exportedHeaders = reopenedSheet.getRange("A1:X1").values[0];
  const exportedRows = reopenedSheet.getRange(`A2:X${lastRow}`).values;
  if (JSON.stringify(exportedHeaders) !== JSON.stringify(EXPECTED_HEADERS)) {
    throw new Error("Post-export header contract mismatch.");
  }
  const exportedIdentityCounts = new Map();
  for (const row of exportedRows) {
    const pmid = row[4];
    const pmvid = row[5];
    const pnid = row[22];
    if (row[7] !== "IMPACT" || Boolean(pmid) === Boolean(pnid)) {
      throw new Error("Post-export KPI type or owner identity validation failed.");
    }
    const identity = pnid ? `PNID:${pnid}` : `PMID:${pmid}:PMVID:${pmvid ?? ""}`;
    if (!exportedIdentityCounts.has(identity)) {
      exportedIdentityCounts.set(identity, { rows: 0, titles: new Set(), weight: 0 });
    }
    const count = exportedIdentityCounts.get(identity);
    count.rows += 1;
    count.titles.add(String(row[10]));
    count.weight += Number(row[16]);
  }
  if (
    exportedIdentityCounts.size !== 80 ||
    [...exportedIdentityCounts.values()].some(
      (value) => value.rows !== 10 || value.titles.size !== 10 || value.weight !== 100,
    )
  ) {
    throw new Error("Post-export per-target KPI contract validation failed.");
  }
  if (
    !exportedIdentityCounts.has("PMID:348:PMVID:39813") ||
    exportedIdentityCounts.has("PNID:54")
  ) {
    throw new Error("Partial PNID 54 was not narrowed to exact PMID/PMVID.");
  }

  const receipt = {
    artifact: outputPath,
    generatedAt: new Date().toISOString(),
    sourceAudit: {
      environment: "production",
      companyId: 1,
      year: 2026,
      capturedAt: "2026-06-15T16:13:10+07:00",
      liveRefreshOn20260619: false,
    },
    sourceWorkbook: sourcePath,
    sourceImpactTitles: impactContract.titles,
    impactWeightTotal: impactContract.weightTotal,
    targetUnits: targets.length,
    structuralPmidUnits: 33,
    pnidUnits: 46,
    exactPartialVariantUnits: 1,
    exactPartialVariant: {
      originalPnid: 54,
      positionMasterId: 348,
      positionMasterVariantId: 39813,
    },
    additionsAfterSnapshot: [
      {
        source: "User report 2026-06-19; identity cross-checked against staging assignment and current area-scope nomenclature evidence",
        employeeName: "Ficky Alkarim",
        employeeNumber: "90003230",
        positionName: "Data Scientist",
        positionMasterId: 33711,
        stagingPositionMasterVariantId: 35658,
        positionNomenclatureId: 11542,
        staleProductionSnapshotPositionNomenclatureId: 11435,
        excludedDifferentPositionNomenclatureId: 12256,
      },
    ],
    uploadRows: rows.length,
    kpiTypeCounts: { IMPACT: rows.length, OUTPUT: 0, KAI: 0 },
    identityCounts: {
      pmidRows: rows.filter((row) => row[4] !== null).length,
      pnidRows: rows.filter((row) => row[22] !== null).length,
      exactPmvidRows: rows.filter((row) => row[5] !== null).length,
    },
    postExportValidation: {
      headersMatchImporter: true,
      targetIdentities: exportedIdentityCounts.size,
      rowsPerTarget: 10,
      uniqueImpactTitlesPerTarget: 10,
      impactWeightPerTarget: 100,
      exactPartialVariantPresent: true,
      pnid54ExpansionAbsent: true,
    },
    caution:
      "Snapshot-based draft. Run production dry-run against a current gap export before confirmed upload.",
    inspect: {
      first: keyCheck.ndjson,
      tail: tailCheck.ndjson,
      formulaErrors: errorCheck.ndjson,
    },
  };
  await fs.writeFile(
    path.join(outputDir, "validation_receipt.json"),
    JSON.stringify(receipt, null, 2),
  );
  console.log(
    JSON.stringify({
      outputPath,
      targetUnits: targets.length,
      uploadRows: rows.length,
      impactTitles: impactContract.titles,
      formulaErrors: errorCheck.ndjson,
    }),
  );
}

await fs.mkdir(outputDir, { recursive: true });
await build();
