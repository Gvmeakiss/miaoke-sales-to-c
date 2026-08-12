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
  "sap_invoice_amount", "oms_monthly_document_no", "oms_monthly_amount", "material_code", "sales_unit", "match_fields",
  "amount_difference", "sap_quantity", "oms_quantity", "source_mapping_result",
];
const headers = [
  "匹配结果分类", "SAP发票编号", "SAP开票日期", "SAP记账日期", "SAP客户/售达方编码", "SAP客户/售达方名称",
  "SAP开票金额（含税）", "OMS月结单号", "OMS月结金额（share_amount）", "物料编码", "销售单位", "匹配字段",
  "差异金额（SAP-OMS）", "SAP开票数量", "OMS月结数量", "原始映射结果",
];
const detailRows = data.rows.map(row => fields.map(field => row[field] ?? ""));
const detailFirst = 5;
const detailLast = detailFirst + detailRows.length - 1;

title(detail, "SAP开票金额—OMS月结金额匹配明细", `期间：${data.period}｜SAP标准发票（2C） vs OMS Y001月结`, "P");
write(detail, 3, 0, [headers]); header(detail.getRange("A4:P4")); detail.getRange("A4:P4").format.rowHeight = 38;
for (let i = 0; i < detailRows.length; i += 2500) write(detail, 4 + i, 0, detailRows.slice(i, i + 2500));
if (detailRows.length) {
  body(detail.getRange(`A5:P${detailLast}`));
  for (let row = detailFirst; row <= detailLast; row += 2500) {
    const end = Math.min(detailLast, row + 2499);
    detail.getRange(`M${row}:M${end}`).formulas = Array.from({ length: end - row + 1 }, (_, index) => [`=G${row + index}-I${row + index}`]);
  }
  detail.getRange(`G5:G${detailLast}`).setNumberFormat("#,##0.00;[Red](#,##0.00);-");
  detail.getRange(`I5:I${detailLast}`).setNumberFormat("#,##0.00;[Red](#,##0.00);-");
  detail.getRange(`M5:M${detailLast}`).setNumberFormat("#,##0.00;[Red](#,##0.00);-");
  detail.getRange(`N5:O${detailLast}`).setNumberFormat("#,##0.00");
  for (const col of ["B", "E", "H", "J"]) detail.getRange(`${col}5:${col}${detailLast}`).setNumberFormat("@");
  const resultRange = detail.getRange(`A5:A${detailLast}`);
  resultRange.conditionalFormats.add("containsText", { text: "完全匹配", format: { fill: C.green } });
  resultRange.conditionalFormats.add("containsText", { text: "差异", format: { fill: C.red } });
  resultRange.conditionalFormats.add("containsText", { text: "不存在", format: { fill: C.amber } });
}
[22, 24, 15, 15, 20, 32, 20, 24, 22, 18, 14, 32, 20, 18, 18, 22].forEach((w, i) => detail.getRange(`${String.fromCharCode(65 + i)}:${String.fromCharCode(65 + i)}`).format.columnWidth = w);
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
summary.getRange("C7").formulas = [[`=SUMIFS('匹配明细'!$I$${detailFirst}:$I$${detailLast},'匹配明细'!$A$${detailFirst}:$A$${detailLast},"完全匹配")+SUMIFS('匹配明细'!$I$${detailFirst}:$I$${detailLast},'匹配明细'!$A$${detailFirst}:$A$${detailLast},"金额差异")`]];
summary.getRange("B8").formulas = [[`=SUMIFS('匹配明细'!$G$${detailFirst}:$G$${detailLast},'匹配明细'!$A$${detailFirst}:$A$${detailLast},"完全匹配")`]];
summary.getRange("C8").formulas = [[`=SUMIFS('匹配明细'!$I$${detailFirst}:$I$${detailLast},'匹配明细'!$A$${detailFirst}:$A$${detailLast},"完全匹配")`]];
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
  summary.getRange(`D${row}`).formulas = [[`=SUMIF('匹配明细'!$A$${detailFirst}:$A$${detailLast},A${row},'匹配明细'!$I$${detailFirst}:$I$${detailLast})`]];
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
  "4. 金额匹配率=双方实际匹配金额较小值/较大值；覆盖率=各方实际匹配金额/各方执行总金额。分类汇总及上方指标均由匹配明细公式计算。",
];
notes.forEach((note, index) => { const row = 19 + index; summary.getRange(`A${row}:G${row}`).merge(); summary.getRange(`A${row}`).values = [[note]]; summary.getRange(`A${row}:G${row}`).format.wrapText = true; summary.getRange(`A${row}:G${row}`).format.rowHeight = 31; });
[34, 21, 21, 22, 18, 18, 18].forEach((w, i) => summary.getRange(`${String.fromCharCode(65 + i)}:${String.fromCharCode(65 + i)}`).format.columnWidth = w);
summary.freezePanes.freezeRows(5);

await fs.mkdir(outputDir, { recursive: true });
const qaDir = path.join(root, "reconciliation/qa_previews"); await fs.mkdir(qaDir, { recursive: true });
console.log("SUMMARY\n" + (await wb.inspect({ kind: "table", range: "匹配汇总!A1:G22", include: "values,formulas", tableMaxRows: 24, tableMaxCols: 8, maxChars: 20000 })).ndjson);
console.log("DETAIL_SAMPLE\n" + (await wb.inspect({ kind: "table", range: "匹配明细!A1:P14", include: "values,formulas", tableMaxRows: 16, tableMaxCols: 17, maxChars: 18000 })).ndjson);
console.log("ERRORS\n" + (await wb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 100 }, summary: "formula errors", maxChars: 4000 })).ndjson);
const preview = await wb.render({ sheetName: "匹配汇总", range: "A1:G22", scale: 1.3, format: "png" }); await fs.writeFile(path.join(qaDir, "SAP-OMS金额匹配汇总.png"), new Uint8Array(await preview.arrayBuffer()));
const detailPreview = await wb.render({ sheetName: "匹配明细", range: "A1:P14", scale: 1.05, format: "png" }); await fs.writeFile(path.join(qaDir, "SAP-OMS金额匹配明细.png"), new Uint8Array(await detailPreview.arrayBuffer()));
const out = await SpreadsheetFile.exportXlsx(wb); await out.save(outputFile);
console.log(JSON.stringify({ outputFile, detailRows: detailRows.length }, null, 2));
