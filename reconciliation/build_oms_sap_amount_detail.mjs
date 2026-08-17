import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

const scriptRoot = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const root = process.env.SALES_TOC_ROOT || scriptRoot;
const inputFile = path.join(root, "reconciliation/results/oms_sap_amount_match_detail.json");
const outputDir = process.env.SALES_TOC_OUTPUT_DIR || path.join(root, "outputs/sales_toc_workpaper_final_20260101_20260630");
const outputFile = process.env.OMS_SAP_DETAIL_OUTPUT_FILE || path.join(outputDir, "SAP开票-OMS月结金额匹配明细_20260101-20260630.xlsx");
const data = JSON.parse(await fs.readFile(inputFile, "utf8"));

const wb = Workbook.create();
const summary = wb.worksheets.add("匹配汇总");
const detail = wb.worksheets.add("匹配明细");
const C = { navy: "#17365D", blue: "#2F75B5", pale: "#DDEBF7", white: "#FFFFFF", text: "#203040", line: "#B4C6E7", green: "#E2F0D9", amber: "#FFF2CC", red: "#FCE4D6" };
const write = (sheet, row, column, rows) => { if (rows.length) sheet.getRangeByIndexes(row, column, rows.length, rows[0].length).values = rows; };
const header = range => { range.format = { fill: C.blue, font: { bold: true, color: C.white }, horizontalAlignment: "center", verticalAlignment: "center", wrapText: true, borders: { preset: "all", style: "thin", color: C.line } }; };
const body = range => { range.format = { font: { color: C.text, size: 10 }, verticalAlignment: "center", borders: { insideHorizontal: { style: "thin", color: "#E7E6E6" } } }; };
function title(sheet, text, subtitle, last) {
  sheet.showGridLines = false;
  sheet.getRange(`A1:${last}1`).merge(); sheet.getRange("A1").values = [[text]];
  sheet.getRange(`A1:${last}1`).format = { fill: C.navy, font: { bold: true, color: C.white, size: 16 }, verticalAlignment: "center" };
  sheet.getRange(`A1:${last}1`).format.rowHeight = 30;
  sheet.getRange(`A2:${last}2`).merge(); sheet.getRange("A2").values = [[subtitle]];
  sheet.getRange(`A2:${last}2`).format = { fill: C.pale, font: { italic: true, color: C.text }, wrapText: true };
  sheet.getRange(`A2:${last}2`).format.rowHeight = 34;
}

const fields = [
  "match_category", "sap_invoice_no", "sap_invoice_date", "sap_posting_date", "sap_customer_code", "sap_customer_name",
  "sap_invoice_amount", "sap_original_oms_sales_no", "oms_monthly_document_no", "oms_monthly_amount", "material_code", "sales_unit", "match_fields",
  "amount_difference", "sap_quantity", "oms_quantity", "source_mapping_result",
];
const headers = [
  "匹配结果分类", "SAP发票编号", "SAP开票日期", "SAP记账日期", "SAP客户/售达方编码", "SAP客户/售达方名称",
  "SAP开票金额（含税）", "SAP原录入OMS销售单号", "OMS月结单号（钩稽口径）", "OMS月结金额（share_amount）", "物料编码", "销售单位", "匹配字段",
  "差异金额（SAP-OMS）", "SAP开票数量", "OMS月结数量", "原始映射结果",
];
const detailRows = data.rows.map(row => fields.map(field => row[field] ?? ""));
const detailFirst = 5;
const detailLast = detailFirst + detailRows.length - 1;

title(detail, "SAP开票金额—OMS月结金额匹配明细", `期间：${data.period}｜SAP标准发票（2C） vs OMS Y001月结`, "Q");
write(detail, 3, 0, [headers]); header(detail.getRange("A4:Q4")); detail.getRange("A4:Q4").format.rowHeight = 38;
for (let i = 0; i < detailRows.length; i += 2500) write(detail, 4 + i, 0, detailRows.slice(i, i + 2500));
if (detailRows.length) {
  body(detail.getRange(`A5:Q${detailLast}`));
  for (let row = detailFirst; row <= detailLast; row += 2500) {
    const end = Math.min(detailLast, row + 2499);
    detail.getRange(`N${row}:N${end}`).formulas = Array.from({ length: end - row + 1 }, (_, index) => [`=G${row + index}-J${row + index}`]);
  }
  detail.getRange(`G5:G${detailLast}`).setNumberFormat("#,##0.00;[Red](#,##0.00);-");
  detail.getRange(`J5:J${detailLast}`).setNumberFormat("#,##0.00;[Red](#,##0.00);-");
  detail.getRange(`N5:N${detailLast}`).setNumberFormat("#,##0.00;[Red](#,##0.00);-");
  detail.getRange(`O5:P${detailLast}`).setNumberFormat("#,##0.00");
  for (const col of ["B", "E", "H", "I", "K"]) detail.getRange(`${col}5:${col}${detailLast}`).setNumberFormat("@");
  const resultRange = detail.getRange(`A5:A${detailLast}`);
  resultRange.conditionalFormats.add("containsText", { text: "完全匹配", format: { fill: C.green } });
  resultRange.conditionalFormats.add("containsText", { text: "差异", format: { fill: C.red } });
  resultRange.conditionalFormats.add("containsText", { text: "不存在", format: { fill: C.amber } });
}
[22, 24, 15, 15, 20, 32, 20, 28, 28, 22, 18, 14, 32, 20, 18, 18, 22].forEach((w, i) => detail.getRange(`${String.fromCharCode(65 + i)}:${String.fromCharCode(65 + i)}`).format.columnWidth = w);
detail.freezePanes.freezeRows(4); detail.freezePanes.freezeColumns(2);

title(summary, "SAP开票金额—OMS月结金额匹配汇总", "汇总金额由“匹配明细”按匹配结果分类透视计算，可直接与主底稿总览勾稽。", "G");
summary.getRange("A4:G4").merge(); summary.getRange("A4").values = [["金额口径及匹配结果"]]; summary.getRange("A4:G4").format = { fill: C.pale, font: { bold: true, color: C.navy } };
write(summary, 4, 0, [["金额层级", "SAP金额", "OMS金额", "差异金额（SAP-OMS）", "金额匹配率", "SAP覆盖率", "OMS覆盖率"]]); header(summary.getRange("A5:G5"));
write(summary, 5, 0, [
  ["执行总金额", data.summary.sap_execution_total, data.summary.oms_execution_total, null, "N/A", "N/A", "N/A"],
  ["参与匹配金额（双方共同键）", null, null, null, null, "N/A", "N/A"],
  ["实际匹配金额（完全匹配）", null, null, null, null, null, null],
]);
summary.getRange("B7").formulas = [[`=SUMIFS('匹配明细'!$G$${detailFirst}:$G$${detailLast},'匹配明细'!$A$${detailFirst}:$A$${detailLast},"完全匹配")+SUMIFS('匹配明细'!$G$${detailFirst}:$G$${detailLast},'匹配明细'!$A$${detailFirst}:$A$${detailLast},"金额差异")`]];
summary.getRange("C7").formulas = [[`=SUMIFS('匹配明细'!$J$${detailFirst}:$J$${detailLast},'匹配明细'!$A$${detailFirst}:$A$${detailLast},"完全匹配")+SUMIFS('匹配明细'!$J$${detailFirst}:$J$${detailLast},'匹配明细'!$A$${detailFirst}:$A$${detailLast},"金额差异")`]];
summary.getRange("B8").formulas = [[`=SUMIFS('匹配明细'!$G$${detailFirst}:$G$${detailLast},'匹配明细'!$A$${detailFirst}:$A$${detailLast},"完全匹配")`]];
summary.getRange("C8").formulas = [[`=SUMIFS('匹配明细'!$J$${detailFirst}:$J$${detailLast},'匹配明细'!$A$${detailFirst}:$A$${detailLast},"完全匹配")`]];
for (const row of [6, 7, 8]) summary.getRange(`D${row}`).formulas = [[`=B${row}-C${row}`]];
for (const row of [7, 8]) summary.getRange(`E${row}`).formulas = [[`=IFERROR(MIN(ABS(B${row}),ABS(C${row}))/MAX(ABS(B${row}),ABS(C${row})),0)`]];
summary.getRange("F8").formulas = [["=IFERROR(ABS(B8)/ABS(B6),0)"]]; summary.getRange("G8").formulas = [["=IFERROR(ABS(C8)/ABS(C6),0)"]];
body(summary.getRange("A6:G8")); summary.getRange("A8:G8").format.fill = C.green; summary.getRange("B6:D8").setNumberFormat("#,##0.00;[Red](#,##0.00);-"); summary.getRange("E6:G8").setNumberFormat("0.00%");

summary.getRange("A10:G10").merge(); summary.getRange("A10").values = [["匹配结果分类汇总"]]; summary.getRange("A10:G10").format = { fill: C.pale, font: { bold: true, color: C.navy } };
write(summary, 10, 0, [["匹配结果分类", "明细键数", "SAP金额", "OMS金额", "差异金额（SAP-OMS）", "SAP金额占执行总额", "OMS金额占执行总额"]]); header(summary.getRange("A11:G11"));
const categories = ["完全匹配", "金额差异", "SAP存在但OMS不存在", "OMS存在但SAP不存在"];
write(summary, 11, 0, categories.map(name => [name, null, null, null, null, null, null]));
for (let i = 0; i < categories.length; i++) {
  const row = 12 + i;
  summary.getRange(`B${row}`).formulas = [[`=COUNTIF('匹配明细'!$A$${detailFirst}:$A$${detailLast},A${row})`]];
  summary.getRange(`C${row}`).formulas = [[`=SUMIF('匹配明细'!$A$${detailFirst}:$A$${detailLast},A${row},'匹配明细'!$G$${detailFirst}:$G$${detailLast})`]];
  summary.getRange(`D${row}`).formulas = [[`=SUMIF('匹配明细'!$A$${detailFirst}:$A$${detailLast},A${row},'匹配明细'!$J$${detailFirst}:$J$${detailLast})`]];
  summary.getRange(`E${row}`).formulas = [[`=C${row}-D${row}`]];
  summary.getRange(`F${row}`).formulas = [[`=IFERROR(ABS(C${row})/ABS($B$6),0)`]];
  summary.getRange(`G${row}`).formulas = [[`=IFERROR(ABS(D${row})/ABS($C$6),0)`]];
}
write(summary, 15, 0, [["合计", null, null, null, null, null, null]]);
summary.getRange("B16").formulas = [["=SUM(B12:B15)"]]; summary.getRange("C16").formulas = [["=SUM(C12:C15)"]]; summary.getRange("D16").formulas = [["=SUM(D12:D15)"]]; summary.getRange("E16").formulas = [["=C16-D16"]]; summary.getRange("F16").formulas = [["=SUM(F12:F15)"]]; summary.getRange("G16").formulas = [["=SUM(G12:G15)"]];
body(summary.getRange("A12:G16")); summary.getRange("A16:G16").format = { fill: C.green, font: { bold: true, color: C.text }, borders: { top: { style: "medium", color: C.blue } } }; summary.getRange("B12:B16").setNumberFormat("#,##0"); summary.getRange("C12:E16").setNumberFormat("#,##0.00;[Red](#,##0.00);-"); summary.getRange("F12:G16").setNumberFormat("0.00%");

summary.getRange("A18:G18").merge(); summary.getRange("A18").values = [["Notes:"]]; summary.getRange("A18:G18").format = { fill: C.pale, font: { bold: true, color: C.navy } };
const notes = [
  "1. SAP开票金额取SAP开票清单“含税金额”，过滤发票类型=标准发票（2C）；SAP开票日期取“发票创建日期”，客户/售达方取“客户编码、客户名称”。",
  "2. OMS月结金额取OMS系统日结月结查询记录“share_amount”，过滤业务类型=Y001且出库日期在2026-01-01至2026-06-30。",
  "3. 参与匹配金额：OMS销售单号+物料编码+销售单位在双方均存在的共同键金额；实际匹配金额：共同键且金额差异不超过0.01元的金额。",
  "4. 客户确认的SAP错录销售单号仅在钩稽层映射更正；“SAP原录入OMS销售单号”保留原值，“OMS月结单号（钩稽口径）”列示更正后匹配值，原始SAP清单不修改。",
  "5. 金额匹配率=双方实际匹配金额较小值/较大值；覆盖率=各方实际匹配金额/各方执行总金额。分类汇总及上方指标均由匹配明细公式计算。",
];
notes.forEach((note, index) => { const row = 19 + index; summary.getRange(`A${row}:G${row}`).merge(); summary.getRange(`A${row}`).values = [[note]]; summary.getRange(`A${row}:G${row}`).format.wrapText = true; summary.getRange(`A${row}:G${row}`).format.rowHeight = 31; });
[34, 21, 21, 22, 18, 18, 18].forEach((w, i) => summary.getRange(`${String.fromCharCode(65 + i)}:${String.fromCharCode(65 + i)}`).format.columnWidth = w);
summary.freezePanes.freezeRows(5);

await fs.mkdir(outputDir, { recursive: true });
const qaDir = path.join(root, "reconciliation/qa_previews"); await fs.mkdir(qaDir, { recursive: true });
console.log("SUMMARY\n" + (await wb.inspect({ kind: "table", range: "匹配汇总!A1:G23", include: "values,formulas", tableMaxRows: 25, tableMaxCols: 8, maxChars: 22000 })).ndjson);
console.log("DETAIL_SAMPLE\n" + (await wb.inspect({ kind: "table", range: "匹配明细!A1:Q14", include: "values,formulas", tableMaxRows: 16, tableMaxCols: 18, maxChars: 20000 })).ndjson);
console.log("ERRORS\n" + (await wb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 100 }, summary: "formula errors", maxChars: 4000 })).ndjson);
const preview = await wb.render({ sheetName: "匹配汇总", range: "A1:G23", scale: 1.3, format: "png" }); await fs.writeFile(path.join(qaDir, "SAP-OMS金额匹配汇总.png"), new Uint8Array(await preview.arrayBuffer()));
const detailPreview = await wb.render({ sheetName: "匹配明细", range: "A1:Q14", scale: 1.05, format: "png" }); await fs.writeFile(path.join(qaDir, "SAP-OMS金额匹配明细.png"), new Uint8Array(await detailPreview.arrayBuffer()));
const out = await SpreadsheetFile.exportXlsx(wb); await out.save(outputFile);

// 同步刷新SAP未参与匹配发票清单，避免更正后的销售单号仍被列作SAP单边。
const unmatchedOutputFile = path.join(outputDir, "SAP未参与匹配发票清单_20260101-20260630.xlsx");
const unmatchedRows = data.rows.filter(row => row.match_category === "SAP存在但OMS不存在");
const unmatchedDetailFirst = 5;
const unmatchedDetailLast = 4 + unmatchedRows.length;
const unmatchedWb = Workbook.create();
const unmatchedSummary = unmatchedWb.worksheets.add("未匹配发票汇总");
const unmatchedDetail = unmatchedWb.worksheets.add("未参与匹配发票明细");
title(unmatchedSummary, "SAP未参与匹配发票清单", "范围：2026-01-01至2026-06-30｜SAP标准发票（2C）｜原始SAP清单不修改。", "G");
unmatchedSummary.getRange("A4:G4").merge(); unmatchedSummary.getRange("A4").values = [["金额及数量控制数"]]; unmatchedSummary.getRange("A4:G4").format = { fill: C.pale, font: { bold: true, color: C.navy } };
write(unmatchedSummary, 4, 0, [["项目", "SAP含税金额", "占SAP全量", "SAP数量", "明细键数", "OMS销售单号状态", "口径说明"]]); header(unmatchedSummary.getRange("A5:G5"));
write(unmatchedSummary, 5, 0, [
  ["SAP标准发票（2C）全量", data.summary.sap_execution_total, 1, null, data.rows.length, "全量", "SAP开票清单含税金额"],
  ["已参与共同键匹配", data.summary.sap_participating_amount, null, null, data.categories["完全匹配"].keys + data.categories["金额差异"].keys, "销售单号存在且共同键命中", "OMS销售单号+物料编码+销售单位"],
  ["未参与共同键匹配", null, null, null, null, "当前均为空", "本文件明细公式汇总"],
]);
unmatchedSummary.getRange("C7").formulas = [["=B7/B6"]];
unmatchedSummary.getRange("B8").formulas = [[`=SUM('未参与匹配发票明细'!K${unmatchedDetailFirst}:K${unmatchedDetailLast})`]];
unmatchedSummary.getRange("C8").formulas = [["=B8/B6"]];
unmatchedSummary.getRange("D8").formulas = [[`=SUM('未参与匹配发票明细'!J${unmatchedDetailFirst}:J${unmatchedDetailLast})`]];
unmatchedSummary.getRange("E8").formulas = [[`=COUNTA('未参与匹配发票明细'!A${unmatchedDetailFirst}:A${unmatchedDetailLast})`]];
body(unmatchedSummary.getRange("A6:G8")); unmatchedSummary.getRange("A8:G8").format.fill = C.amber; unmatchedSummary.getRange("B6:B8").setNumberFormat("#,##0.00;[Red](#,##0.00);-"); unmatchedSummary.getRange("C6:C8").setNumberFormat("0.00%"); unmatchedSummary.getRange("D6:E8").setNumberFormat("#,##0");
unmatchedSummary.getRange("A10:G10").merge(); unmatchedSummary.getRange("A10").values = [["Notes:"]]; unmatchedSummary.getRange("A10:G10").format = { fill: C.pale, font: { bold: true, color: C.navy } };
const unmatchedNotes = [
  "1. 客户确认SAP发票9001010368的OMS销售单号误录后，已在钩稽层更正并重新匹配，因此不再列入本清单。",
  "2. 当前未参与匹配记录均因SAP原始字段“OMS销售单号”为空，缺少OMS销售单号+物料编码+销售单位共同键中的主键。",
  "3. 未参与匹配金额及数量由下方明细公式汇总；原始SAP文件不修改。",
];
unmatchedNotes.forEach((note, index) => { const row = 11 + index; unmatchedSummary.getRange(`A${row}:G${row}`).merge(); unmatchedSummary.getRange(`A${row}`).values = [[note]]; unmatchedSummary.getRange(`A${row}:G${row}`).format.wrapText = true; unmatchedSummary.getRange(`A${row}:G${row}`).format.rowHeight = 30; });
[28, 20, 16, 16, 16, 26, 42].forEach((w, i) => unmatchedSummary.getRange(`${String.fromCharCode(65 + i)}:${String.fromCharCode(65 + i)}`).format.columnWidth = w); unmatchedSummary.freezePanes.freezeRows(5);

const unmatchedHeaders = ["序号", "SAP文件月份", "SAP发票编号", "SAP开票日期", "SAP客户/售达方编码", "SAP客户/售达方名称", "SAP原录入OMS销售单号", "物料编码", "销售单位", "SAP开票数量", "SAP含税金额", "未参与匹配原因", "匹配说明"];
title(unmatchedDetail, "SAP未参与匹配发票明细", `正式范围：2026-01-01至2026-06-30｜共${unmatchedRows.length.toLocaleString("zh-CN")}个核对键`, "M");
write(unmatchedDetail, 3, 0, [unmatchedHeaders]); header(unmatchedDetail.getRange("A4:M4")); unmatchedDetail.getRange("A4:M4").format.rowHeight = 38;
const unmatchedDetailRows = unmatchedRows.map((row, index) => [index + 1, row.sap_invoice_date?.slice(0, 7) || "", row.sap_invoice_no || "", row.sap_invoice_date || "", row.sap_customer_code || "", row.sap_customer_name || "", row.sap_original_oms_sales_no || "", row.material_code || "", row.sales_unit || "", row.sap_quantity || 0, row.sap_invoice_amount || 0, "OMS销售单号为空", "缺少主匹配字段，未进入OMS销售单号+物料编码+销售单位共同键"]);
write(unmatchedDetail, 4, 0, unmatchedDetailRows); body(unmatchedDetail.getRange(`A5:M${unmatchedDetailLast}`)); unmatchedDetail.getRange(`A5:M${unmatchedDetailLast}`).format.wrapText = true; unmatchedDetail.getRange(`A5:M${unmatchedDetailLast}`).format.rowHeight = 30; for (const c of ["C","E","G","H"]) unmatchedDetail.getRange(`${c}5:${c}${unmatchedDetailLast}`).setNumberFormat("@"); unmatchedDetail.getRange(`J5:J${unmatchedDetailLast}`).setNumberFormat("#,##0"); unmatchedDetail.getRange(`K5:K${unmatchedDetailLast}`).setNumberFormat("#,##0.00;[Red](#,##0.00);-"); unmatchedDetail.getRange(`L5:M${unmatchedDetailLast}`).format.fill = C.amber;
[10, 16, 20, 16, 20, 34, 28, 18, 14, 18, 20, 24, 46].forEach((w, i) => unmatchedDetail.getRange(`${String.fromCharCode(65 + i)}:${String.fromCharCode(65 + i)}`).format.columnWidth = w); unmatchedDetail.freezePanes.freezeRows(4); unmatchedDetail.freezePanes.freezeColumns(3);
console.log("UNMATCHED_SUMMARY\n" + (await unmatchedWb.inspect({ kind: "table", range: "未匹配发票汇总!A1:G13", include: "values,formulas", tableMaxRows: 15, tableMaxCols: 8, maxChars: 15000 })).ndjson);
console.log("UNMATCHED_ERRORS\n" + (await unmatchedWb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 100 }, summary: "formula errors", maxChars: 4000 })).ndjson);
const unmatchedPreview = await unmatchedWb.render({ sheetName: "未匹配发票汇总", range: "A1:G13", scale: 1.25, format: "png" }); await fs.writeFile(path.join(qaDir, "SAP未参与匹配发票汇总.png"), new Uint8Array(await unmatchedPreview.arrayBuffer()));
const unmatchedOut = await SpreadsheetFile.exportXlsx(unmatchedWb); await unmatchedOut.save(unmatchedOutputFile);
console.log(JSON.stringify({ outputFile, detailRows: detailRows.length, unmatchedOutputFile, unmatchedRows: unmatchedRows.length }, null, 2));
