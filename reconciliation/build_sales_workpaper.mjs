import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

const scriptRoot = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const root = process.env.SALES_TOC_ROOT || scriptRoot;
const dataDir = process.env.SALES_TOC_DATA_DIR || path.join(root, "reconciliation/results");
const outputDir = process.env.SALES_TOC_OUTPUT_DIR || path.join(root, "outputs/sales_toc_workpaper_final_20260101_20260630");
const outputFile = process.env.SALES_TOC_OUTPUT_FILE || path.join(outputDir, "销售ToC业务流程核对底稿_20260101-20260630.xlsx");
const read = async name => JSON.parse(await fs.readFile(path.join(dataDir, name), "utf8"));
const formalDetailDir = path.join(outputDir, "2026年度正式范围旺店通订单匹配明细");
const formalDetailManifest = await fs.readFile(path.join(formalDetailDir, "明细拆分索引及校验结果.json"), "utf8")
  .then(JSON.parse)
  .catch(() => ({ categories: [], totals: {}, verification: {} }));
const S = await read("summary.json");
const cutoffSensitivity = await read("wdt_cutoff_sensitivity.json").catch(() => ({ data_check: {}, sensitivity: [], definitions: {} }));
const data = {
  internal: await read("huice_internal_recon_workbook.json"),
  billOms: await read("bill_oms_month_recon_workbook.json"),
  qty: await read("order_bill_oms_qty_recon_workbook.json"),
  omsSap: await read("oms_sap_field_map_workbook.json"),
  shop: await read("huice_shop_map_workbook.json"),
};

const wb = Workbook.create();
const names = [
  "1.全局口径与总览",
  "2.订单-账单汇总",
  "4.惠策内部核对汇总",
  "5.惠策内部核对明细",
  "6.账单-OMS月结汇总",
  "7.账单-OMS月结明细",
  "8.数量核对汇总",
  "9.数量核对明细",
  "10.OMS月结-SAP汇总",
  "旺店通订单匹配明细",
  "13.店铺客户映射",
];
for (const name of names) wb.worksheets.add(name);

const scope = "账单/实际结算期间：2026-01-01至2026-06-30｜旺店通正式匹配期间：2026-01-01至2026-06-30";
const C = {
  navy: "#17365D", blue: "#2F75B5", pale: "#DDEBF7", pale2: "#EAF3F8", white: "#FFFFFF",
  text: "#203040", line: "#B4C6E7", green: "#E2F0D9", greenText: "#375623",
  amber: "#FFF2CC", amberText: "#7F6000", red: "#FCE4D6", redText: "#9C0006",
};
const ws = name => wb.worksheets.getItem(name);
const col = index => { let n = index + 1, s = ""; while (n) { n -= 1; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26); } return s; };
const clean = value => value === undefined || value === null || (typeof value === "number" && !Number.isFinite(value)) ? null : (typeof value === "string" && /^\d{12,}$/.test(value) ? `\u200B${value}` : value);
const write = (sheet, row, column, rows) => { if (rows.length && rows[0].length) sheet.getRangeByIndexes(row, column, rows.length, rows[0].length).values = rows.map(r => r.map(clean)); };
const formula = (sheet, cell, value) => sheet.getRange(cell).formulas = [[value]];
function title(sheet, text, subtitle, last = "H") {
  sheet.showGridLines = false;
  sheet.getRange(`A1:${last}1`).merge(); sheet.getRange("A1").values = [[text]];
  sheet.getRange(`A1:${last}1`).format = { fill: C.navy, font: { bold: true, color: C.white, size: 16 }, verticalAlignment: "center" };
  sheet.getRange(`A1:${last}1`).format.rowHeight = 30;
  sheet.getRange(`A2:${last}2`).merge(); sheet.getRange("A2").values = [[`${scope}｜${subtitle}`]];
  sheet.getRange(`A2:${last}2`).format = { fill: C.pale2, font: { italic: true, color: C.text, size: 10 }, wrapText: true, verticalAlignment: "center" };
  sheet.getRange(`A2:${last}2`).format.rowHeight = 34;
}
function header(sheet, range) { sheet.getRange(range).format = { fill: C.blue, font: { bold: true, color: C.white }, horizontalAlignment: "center", verticalAlignment: "center", wrapText: true, borders: { preset: "all", style: "thin", color: C.line } }; }
function section(sheet, range) { sheet.getRange(range).format = { fill: C.pale, font: { bold: true, color: C.navy, size: 11 }, borders: { bottom: { style: "medium", color: C.blue } } }; }
function body(sheet, range) { sheet.getRange(range).format = { font: { color: C.text, size: 10 }, verticalAlignment: "center", borders: { insideHorizontal: { style: "thin", color: "#E7E6E6" } } }; }
function status(sheet, range) {
  const x = sheet.getRange(range);
  x.conditionalFormats.add("containsText", { text: "差异", format: { fill: C.red, font: { bold: true, color: C.redText } } });
  x.conditionalFormats.add("containsText", { text: "未映射", format: { fill: C.red, font: { bold: true, color: C.redText } } });
  x.conditionalFormats.add("containsText", { text: "仅", format: { fill: C.amber, font: { bold: true, color: C.amberText } } });
  x.conditionalFormats.add("containsText", { text: "附属单", format: { fill: C.amber, font: { bold: true, color: C.amberText } } });
  x.conditionalFormats.add("containsText", { text: "一致", format: { fill: C.green, font: { bold: true, color: C.greenText } } });
}
const headerZh = {
  bill_month: "账单月份", platform: "平台", huice_shop: "惠策店铺", customer_code: "OMS客户编码", customer_name: "OMS客户名称",
  detail_rows: "惠策明细行数", summary_rows: "惠策汇总源行数", detail_cash: "明细实际实收", summary_cash: "汇总实际结算", cash_difference: "实收差异", result: "核对结果",
  bill_cash: "惠策实际结算", oms_docs: "OMS单据数", oms_qty: "OMS数量", oms_amount: "OMS实际结算", material_code: "物料编码", ship_month: "发货月份", wdt_shop: "旺店通店铺", unit: "单位",
  order_bill_qty: "惠策账单对应订单数量", qty_difference: "数量差异", mapping_result: "匹配分类", sap_qty: "SAP数量", sap_amount: "SAP含税金额", wdt_item_amount: "旺店通商品金额",
  billed_orders: "账单证据平台单数", source_rows: "源文件行数", bill_record_count: "账单记录数", mapping_status: "映射状态", mapping_source: "映射来源",
  oms_sales_no: "OMS销售单号", sales_unit: "销售单位", file_month: "SAP文件月份", outbound_month: "OMS出库月份", sap_invoice_nos: "SAP发票号", sap_rows: "SAP行数", oms_rows: "OMS行数", quantity_difference: "数量差异", amount_difference: "金额差异", mapped_qty: "映射数量", mapped_amount: "映射金额", source_result: "来源分类",
};
function detail(name, reportTitle, subtitle, dataset, widths = {}, omit = []) {
  const keep = dataset.headers.map((h, i) => ({ h, i })).filter(x => !omit.includes(x.h));
  const headers = keep.map(x => x.h);
  const rows = keep.length === dataset.headers.length ? dataset.rows : dataset.rows.map(r => keep.map(x => r[x.i]));
  const sheet = ws(name), last = col(headers.length - 1), titleLast = col(Math.min(headers.length, 10) - 1);
  title(sheet, reportTitle, subtitle, titleLast);
  write(sheet, 3, 0, [headers.map(h => headerZh[h] ? `${h}\n${headerZh[h]}` : h)]); header(sheet, `A4:${last}4`); sheet.getRange(`A4:${last}4`).format.rowHeight = 34;
  for (let i = 0; i < rows.length; i += 3000) write(sheet, 4 + i, 0, rows.slice(i, i + 3000));
  const end = 4 + rows.length; if (rows.length) { body(sheet, `A5:${last}${end}`); for (const key of ["result", "mapping_result"]) { const index = headers.indexOf(key); if (index >= 0) status(sheet, `${col(index)}5:${col(index)}${end}`); } }
  const textFields = new Set(["customer_code", "material_code", "platform_order_no", "oms_sales_no", "sap_invoice_nos", "reconcile_ids"]);
  headers.forEach((h, i) => { const width = widths[h] || (/name|shop|orders|invoice|reconcile/.test(h) ? 25 : /amount|difference|cash/.test(h) ? 17 : 14); sheet.getRange(`${col(i)}:${col(i)}`).format.columnWidth = width; if (textFields.has(h)) sheet.getRange(`${col(i)}5:${col(i)}${end}`).setNumberFormat("@"); else if (/amount|difference|cash/.test(h)) sheet.getRange(`${col(i)}5:${col(i)}${end}`).setNumberFormat("#,##0.00;[Red](#,##0.00);-"); else if (/qty|count|rows|lines|groups|keys|docs/.test(h)) sheet.getRange(`${col(i)}5:${col(i)}${end}`).setNumberFormat("#,##0"); });
  sheet.freezePanes.freezeRows(4); sheet.freezePanes.freezeColumns(Math.min(2, headers.length));
}
const by = (arr, key) => Object.fromEntries(arr.map(x => [x[key], x]));
const ctl = S.controls, dm = S.display_metrics || {}, orderResults = S.order_bill_results, internalResults = S.huice_internal_results, billOmsResults = S.bill_oms_results, qtyResults = S.qty_results;
const omsSapResults = S.oms_sap_results, exactMap = by(omsSapResults, "mapping_result")["双向字段一致"] || { keys: 0, sap_qty: 0, oms_qty: 0, sap_amount: 0, oms_amount: 0 };
const findResult = (rows, label) => rows.find(x => x.result === label) || { groups: 0, wdt_amount: 0, bill_cash: 0, detail_cash: 0, summary_cash: 0, oms_amount: 0, oms_qty: 0, order_bill_qty: 0 };
const findMap = label => omsSapResults.find(x => x.mapping_result === label) || { keys: 0, sap_qty: 0, oms_qty: 0, sap_amount: 0, oms_amount: 0 };
const orderMatchCategories = ["单号分摊实收一致", "单号订单实收一致"];
const orderCategories = ["单号分摊实收一致", "单号订单实收一致", "单号一致金额差异", "仅账单", "仅订单"];
const matchedWdtAmount = orderMatchCategories.reduce((sum, label) => { const x = findResult(orderResults, label); return sum + (label === "单号订单实收一致" ? (x.wdt_header_amount || 0) : (x.wdt_amount || 0)); }, 0);
const matchedBillCash = orderMatchCategories.reduce((sum, label) => sum + (findResult(orderResults, label).bill_cash || 0), 0);
const matchedOrderGroups = orderMatchCategories.reduce((sum, label) => sum + (findResult(orderResults, label).groups || 0), 0);
const fullWdtComparisonAmount = orderCategories.reduce((sum, label) => sum + (findResult(orderResults, label).wdt_amount || 0), 0);
const cutoffDaily = S.huice_cutoff_daily || [{ business_date: "2026-06-28", bill_count: 0, bill_cash: 0 }, { business_date: "2026-06-29", bill_count: 0, bill_cash: 0 }];
const billOmsCategories = ["实际结算金额一致", "SAP辅助实际结算金额一致", "实际结算金额差异", "仅账单", "仅OMS月结", "店铺未映射"].filter(label => { const x = findResult(billOmsResults, label); return (x.groups || 0) > 0 || Math.abs(x.bill_cash || 0) > 0.000001 || Math.abs(x.oms_amount || 0) > 0.000001; });

// 场景1：订单—账单汇总。成功匹配金额单独列示；分类表仅以惠策实际实收为唯一金额口径。
const orderRefs = {};
{
  const sheet = ws("2.订单-账单汇总"); title(sheet, "订单—账单实际实收金额核对汇总", "账单期间内，以惠策实际实收为主口径；旺店通金额仅用于成功匹配订单的桥接核对。", "H");
  section(sheet, "A4:H4"); sheet.getRange("A4").values = [["惠策账单—旺店通订单匹配金额汇总"]]; sheet.getRange("A4:H4").merge();
  write(sheet, 4, 0, [["项目", "金额/比率", "来源系统", "原始清单名称", "使用字段", "汇总逻辑", "与总览勾稽", "审计口径"]]); header(sheet, "A5:H5");
  write(sheet, 5, 0, [
    ["惠策账单匹配金额", matchedBillCash, "惠策系统", "惠策账单清单", "收款金额（正实收）；退款金额（负实收）", "账单净实收按平台订单号汇总，仅纳入金额一致的订单", "总览：订单—账单成功匹配", "实际实收"],
    ["对应旺店通订单金额", matchedWdtAmount, "旺店通", "旺店通订单清单", "allocated_total；订单总金额", "优先使用分摊金额；订单总金额一致场景使用订单头金额", "总览：订单—账单成功匹配", "订单桥接金额"],
    ["差异金额（旺店通-惠策）", null, "—", "—", "—", "对应旺店通订单金额-惠策账单匹配金额", "总览：订单—账单成功匹配", "核对差异"],
    ["金额匹配率", null, "—", "—", "—", "较小金额/较大金额", "总览：订单—账单成功匹配", "匹配率"],
  ]); formula(sheet, "B8", "=B7-B6"); formula(sheet, "B9", "=IFERROR(MIN(ABS(B6),ABS(B7))/MAX(ABS(B6),ABS(B7)),0)"); body(sheet, "A6:H9"); sheet.getRange("B6:B8").setNumberFormat("#,##0.00;[Red](#,##0.00);-"); sheet.getRange("B9").setNumberFormat("0.00%"); sheet.getRange("A6:H9").format.wrapText = true; orderRefs.matchedBillRow = 6; orderRefs.matchedWdtRow = 7; orderRefs.matchedDiffRow = 8; orderRefs.matchedRateRow = 9;
  section(sheet, "A11:H11"); sheet.getRange("A11").values = [["旺店通—惠策明细全量金额比对"]]; sheet.getRange("A11:H11").merge();
  write(sheet, 11, 0, [["项目", "金额/比率", "来源系统", "原始清单名称", "使用字段", "汇总逻辑", "与总览勾稽", "审计口径"]]); header(sheet, "A12:H12");
  write(sheet, 12, 0, [
    ["旺店通全量比对金额", fullWdtComparisonAmount, "旺店通", "旺店通订单清单", "allocated_total / wdt_amount", "按平台订单号汇总分摊金额，纳入订单—账单全部分类，零金额附属单不重复计入", "总览：旺店通—惠策明细全量", "订单桥接金额"],
    ["惠策明细全量实际实收", null, "惠策系统", "惠策账单清单", "收款金额（正实收）；退款金额（负实收）", "按平台订单号分类汇总全部净实收", "总览及下一行惠策内部", "实际实收"],
    ["差异金额（旺店通-惠策）", null, "—", "—", "—", "旺店通全量比对金额-惠策明细全量实际实收", "总览：旺店通—惠策明细全量", "核对差异"],
    ["金额匹配率", null, "—", "—", "—", "较小金额/较大金额", "总览：旺店通—惠策明细全量", "匹配率"],
  ]); body(sheet, "A13:H16"); sheet.getRange("A13:H16").format.wrapText = true; orderRefs.fullWdtRow = 13; orderRefs.fullBillRow = 14; orderRefs.fullDiffRow = 15; orderRefs.fullRateRow = 16;
  section(sheet, "A18:H18"); sheet.getRange("A18").values = [["惠策账单实际实收分类金额及占比"]]; sheet.getRange("A18:H18").merge();
  write(sheet, 18, 0, [["核对分类", "订单组数", "惠策实际实收金额", "惠策实收占比", "来源系统/清单", "使用字段", "汇总逻辑", "与总览勾稽"]]); header(sheet, "A19:H19");
  const rows = orderCategories.map(label => { const x = findResult(orderResults, label); return [label, x.groups || 0, x.bill_cash || 0, null, "惠策系统/惠策账单清单", "收款金额（正实收）；退款金额（负实收）", "按平台订单号分类后汇总净实收", "分类合计=总览惠策账单金额"]; });
  write(sheet, 19, 0, rows); const first = 20, last = 19 + orderCategories.length, total = last + 1; orderRefs.categoryRows = Object.fromEntries(orderCategories.map((x, i) => [x, first + i])); orderRefs.totalRow = total;
  for (let r = first; r <= last; r++) formula(sheet, `D${r}`, `=IFERROR(ABS(C${r})/ABS($C$${total}),0)`);
  write(sheet, total - 1, 0, [["合计", null, null, null, "惠策系统/惠策账单清单", "收款金额-退款金额", "全部分类金额合计", "直接勾稽总览"]]); formula(sheet, `B${total}`, `=SUM(B${first}:B${last})`); formula(sheet, `C${total}`, `=SUM(C${first}:C${last})`); formula(sheet, `D${total}`, `=SUM(D${first}:D${last})`);
  formula(sheet, "B14", `=C${total}`); formula(sheet, "B15", "=B13-B14"); formula(sheet, "B16", "=IFERROR(MIN(ABS(B13),ABS(B14))/MAX(ABS(B13),ABS(B14)),0)"); sheet.getRange("B13:B15").setNumberFormat("#,##0.00;[Red](#,##0.00);-"); sheet.getRange("B16").setNumberFormat("0.00%");
  body(sheet, `A20:H${total}`); status(sheet, `A20:A${last}`); sheet.getRange(`B20:B${total}`).setNumberFormat("#,##0"); sheet.getRange(`C20:C${total}`).setNumberFormat("#,##0.00;[Red](#,##0.00);-"); sheet.getRange(`D20:D${total}`).setNumberFormat("0.00%"); sheet.getRange(`A${total}:H${total}`).format = { fill: C.green, font: { bold: true, color: C.greenText }, borders: { top: { style: "medium", color: C.blue } } };
  [30, 18, 20, 18, 28, 32, 36, 28].forEach((w, i) => sheet.getRange(`${col(i)}:${col(i)}`).format.columnWidth = w); sheet.freezePanes.freezeRows(19);
}

// 场景2：惠策明细—店铺汇总，仅保留实际实收。
const internalRefs = {};
{
  const sheet = ws("4.惠策内部核对汇总"); title(sheet, "惠策明细—惠策店铺汇总实际实收核对汇总", "两份惠策清单均按导出结算月份+平台+店铺核对。", "H");
  section(sheet, "A4:H4"); sheet.getRange("A4").values = [["金额字段定义"]]; sheet.getRange("A4:H4").merge();
  write(sheet, 4, 0, [["金额项目", "来源清单", "原始字段", "计算逻辑", "总额", "勾稽用途", "对应总览项目", "备注"]]); header(sheet, "A5:H5");
  write(sheet, 5, 0, [["惠策明细实际实收", "惠策账单清单", "收款金额（正实收）；退款金额（负实收）", "对账流水号净额=收款金额-退款金额，再按结算月+平台+店铺汇总", null, "明细侧", "惠策内部", "本期结算清单"], ["惠策汇总实际结算", "惠策店铺汇总清单", "成功金额；不一致实收金额；单边实收金额", "成功金额+不一致实收金额+单边实收金额，按结算月+平台+店铺汇总", null, "汇总侧", "惠策内部", "店铺汇总分类重构"]]); body(sheet, "A6:H7"); sheet.getRange("A6:H7").format.wrapText = true; sheet.getRange("A6:H7").format.rowHeight = 36;
  formula(sheet, "E6", `=SUM(C11:C${10 + Math.max(1, internalResults.length)})`); formula(sheet, "E7", `=SUM(D11:D${10 + Math.max(1, internalResults.length)})`); sheet.getRange("E6:E7").setNumberFormat("#,##0.00;[Red](#,##0.00);-");
  section(sheet, "A9:H9"); sheet.getRange("A9").values = [["分类金额及匹配率"]]; sheet.getRange("A9:H9").merge(); write(sheet, 9, 0, [["核对结果", "月份/平台/店铺组合数", "惠策明细实际实收", "惠策汇总实际结算", "差异（汇总-明细）", "金额匹配率", "分类处理", "备注"]]); header(sheet, "A10:H10");
  const rows = internalResults.length ? internalResults.map(x => [x.result, x.groups || 0, x.detail_cash || 0, x.summary_cash || 0, null, null, x.result === "实收一致" ? "金额一致分类" : "差异分类", "按结算月份+平台+店铺核对"]) : [["无记录", 0, 0, 0, null, null, "分类列示", ""]]; write(sheet, 10, 0, rows); const first = 11, last = 10 + rows.length, total = last + 1; internalRefs.totalRow = total;
  for (let r = first; r <= last; r++) { formula(sheet, `E${r}`, `=D${r}-C${r}`); formula(sheet, `F${r}`, `=IFERROR(MIN(ABS(C${r}),ABS(D${r}))/MAX(ABS(C${r}),ABS(D${r})),0)`); }
  write(sheet, total - 1, 0, [["合计", null, null, null, null, null, "分类合计与总览勾稽", ""]]); formula(sheet, `B${total}`, `=SUM(B${first}:B${last})`); formula(sheet, `C${total}`, `=SUM(C${first}:C${last})`); formula(sheet, `D${total}`, `=SUM(D${first}:D${last})`); formula(sheet, `E${total}`, `=D${total}-C${total}`); formula(sheet, `F${total}`, `=IFERROR(MIN(ABS(C${total}),ABS(D${total}))/MAX(ABS(C${total}),ABS(D${total})),0)`); body(sheet, `A11:H${total}`); status(sheet, `A11:A${last}`); sheet.getRange(`B11:B${total}`).setNumberFormat("#,##0"); sheet.getRange(`C11:E${total}`).setNumberFormat("#,##0.00;[Red](#,##0.00);-"); sheet.getRange(`F11:F${total}`).setNumberFormat("0.00%"); sheet.getRange(`A${total}:H${total}`).format = { fill: C.green, font: { bold: true, color: C.greenText }, borders: { top: { style: "medium", color: C.blue } } }; [30, 18, 20, 20, 20, 17, 22, 28].forEach((w, i) => sheet.getRange(`${col(i)}:${col(i)}`).format.columnWidth = w); sheet.getRange("A6:H7").format.wrapText = true; sheet.freezePanes.freezeRows(10);
}

// 场景3：惠策店铺汇总—OMS月结，主口径为实际结算金额。
const billOmsRefs = {};
{
  const sheet = ws("6.账单-OMS月结汇总"); title(sheet, "惠策店铺汇总—OMS月结实际结算金额核对汇总", "惠策侧采用店铺汇总实际结算金额，OMS侧采用Y001月结结算金额。", "H");
  section(sheet, "A4:H4"); sheet.getRange("A4").values = [["金额字段定义"]]; sheet.getRange("A4:H4").merge(); write(sheet, 4, 0, [["金额项目", "来源清单", "原始字段", "计算逻辑", "总额", "勾稽用途", "对应总览项目", "备注"]]); header(sheet, "A5:H5");
  write(sheet, 5, 0, [["惠策实际结算金额", "惠策店铺汇总清单", "成功金额；不一致实收金额；单边实收金额", "成功金额+不一致实收金额+单边实收金额，按结算月+平台+店铺汇总", null, "账单侧", "账单—OMS", "实际结算口径"], ["OMS月结金额", "OMS系统日结月结查询记录SQL", "share_amount（结算金额）", "按出库月份+OMS客户编码汇总Y001单据", null, "OMS侧", "账单—OMS", "实际结算口径"]]); body(sheet, "A6:H7"); sheet.getRange("A6:H7").format.wrapText = true; sheet.getRange("A6:H7").format.rowHeight = 36;
  formula(sheet, "E6", `=SUM(C11:C${10 + billOmsCategories.length})`); formula(sheet, "E7", `=SUM(D11:D${10 + billOmsCategories.length})`); sheet.getRange("E6:E7").setNumberFormat("#,##0.00;[Red](#,##0.00);-");
  section(sheet, "A9:H9"); sheet.getRange("A9").values = [["分类金额及匹配率"]]; sheet.getRange("A9:H9").merge(); write(sheet, 9, 0, [["核对结果", "月份/客户组合数", "惠策实际结算金额", "OMS结算金额", "差异（OMS-惠策）", "金额匹配率", "分类处理", "备注"]]); header(sheet, "A10:H10");
  const rows = billOmsCategories.map(label => { const x = findResult(billOmsResults, label); return [label, x.groups || 0, x.bill_cash || 0, x.oms_amount || 0, null, null, label.includes("一致") ? "金额一致分类" : "差异分类", label === "店铺未映射" ? "店铺未取得OMS客户编码" : label === "仅账单" ? "无对应OMS月结" : label === "仅OMS月结" ? "无对应惠策汇总" : "按月份+客户编码核对"]; }); write(sheet, 10, 0, rows); const first = 11, last = 10 + rows.length, total = last + 1; billOmsRefs.categoryRows = Object.fromEntries(billOmsCategories.map((x, i) => [x, first + i])); billOmsRefs.totalRow = total;
  for (let r = first; r <= last; r++) { formula(sheet, `E${r}`, `=D${r}-C${r}`); formula(sheet, `F${r}`, `=IFERROR(MIN(ABS(C${r}),ABS(D${r}))/MAX(ABS(C${r}),ABS(D${r})),0)`); }
  write(sheet, total - 1, 0, [["合计", null, null, null, null, null, "分类合计与总览勾稽", ""]]); formula(sheet, `B${total}`, `=SUM(B${first}:B${last})`); formula(sheet, `C${total}`, `=SUM(C${first}:C${last})`); formula(sheet, `D${total}`, `=SUM(D${first}:D${last})`); formula(sheet, `E${total}`, `=D${total}-C${total}`); formula(sheet, `F${total}`, `=IFERROR(MIN(ABS(C${total}),ABS(D${total}))/MAX(ABS(C${total}),ABS(D${total})),0)`); body(sheet, `A11:H${total}`); status(sheet, `A11:A${last}`); sheet.getRange(`B11:B${total}`).setNumberFormat("#,##0"); sheet.getRange(`C11:E${total}`).setNumberFormat("#,##0.00;[Red](#,##0.00);-"); sheet.getRange(`F11:F${total}`).setNumberFormat("0.00%"); sheet.getRange(`A${total}:H${total}`).format = { fill: C.green, font: { bold: true, color: C.greenText }, borders: { top: { style: "medium", color: C.blue } } }; [30, 18, 20, 20, 20, 17, 22, 28].forEach((w, i) => sheet.getRange(`${col(i)}:${col(i)}`).format.columnWidth = w); sheet.getRange("A6:H7").format.wrapText = true; sheet.freezePanes.freezeRows(10);
}

// 场景4：数量沿业务流分两步核对，不进行四方直接比较。
const qtyRefs = {};
{
  const sheet = ws("8.数量核对汇总"); title(sheet, "销售数量逐层核对汇总", "惠策无原生商品数量；Step 1 使用惠策账单已匹配订单对应的旺店通商品数量作为账单数量证据。", "K");
  section(sheet, "A4:K4"); sheet.getRange("A4").values = [["Step 1｜惠策账单订单证据数量 vs OMS月结数量"]]; sheet.getRange("A4:K4").merge();
  write(sheet, 4, 0, [["核对步骤", "左侧数量项目", "来源系统/原始清单", "使用字段/汇总逻辑", "左侧数量", "右侧数量项目", "来源系统/原始清单", "使用字段/汇总逻辑", "右侧数量", "差异数量（右-左）", "差异率"]]); header(sheet, "A5:K5");
  write(sheet, 5, 0, [["Step 1", "惠策账单已匹配订单对应的旺店通商品数量", "惠策账单清单+旺店通订单清单", "平台订单号匹配后汇总旺店通quantity", ctl.billed_wdt_qty || 0, "OMS月结数量", "OMS系统日结月结查询记录SQL", "item_num；Y001按出库月+客户+物料汇总", ctl.oms_month_qty || 0, null, null]]); formula(sheet, "J6", "=I6-E6"); formula(sheet, "K6", "=IFERROR(ABS(J6)/MAX(ABS(E6),ABS(I6)),0)"); qtyRefs.step1Row = 6;
  section(sheet, "A8:K8"); sheet.getRange("A8").values = [["Step 2｜OMS月结数量 vs SAP标准发票（2C）数量"]]; sheet.getRange("A8:K8").merge();
  write(sheet, 8, 0, [["核对步骤", "左侧数量项目", "来源系统/原始清单", "使用字段/汇总逻辑", "左侧数量", "右侧数量项目", "来源系统/原始清单", "使用字段/汇总逻辑", "右侧数量", "差异数量（右-左）", "差异率"]]); header(sheet, "A9:K9");
  write(sheet, 9, 0, [["Step 2", "OMS共同键数量", "OMS系统日结月结查询记录SQL", "item_num；销售单号+物料+销售单位共同键汇总", exactMap.oms_qty || 0, "SAP共同键数量", "SAP开票清单", "invoice_qty；标准发票（2C）共同键汇总", exactMap.sap_qty || 0, null, null]]); formula(sheet, "J10", "=I10-E10"); formula(sheet, "K10", "=IFERROR(ABS(J10)/MAX(ABS(E10),ABS(I10)),0)"); qtyRefs.step2Row = 10;
  body(sheet, "A6:K6"); body(sheet, "A10:K10"); sheet.getRange("E6:E10").setNumberFormat("#,##0"); sheet.getRange("I6:J10").setNumberFormat("#,##0"); sheet.getRange("K6:K10").setNumberFormat("0.00%"); sheet.getRange("A6:K10").format.wrapText = true; sheet.getRange("A6:K6").format.rowHeight = 48; sheet.getRange("A10:K10").format.rowHeight = 48;
  write(sheet, 12, 0, [["明细支持", "9.数量核对明细 Step 1", "惠策账单订单证据—OMS", "按发货月份+客户/店铺+物料", null, "9.数量核对明细 Step 2", "OMS—SAP", "按销售单号+物料+销售单位", null, null, null]]); body(sheet, "A13:K13");
  [14, 34, 30, 38, 17, 25, 30, 38, 17, 20, 16].forEach((w, i) => sheet.getRange(`${col(i)}:${col(i)}`).format.columnWidth = w); sheet.freezePanes.freezeRows(5);
}

// 场景5：OMS—SAP汇总。同时列示全量与共同键范围，供去年式总览直接勾稽。
const sapRefs = {};
{
  const sheet = ws("10.OMS月结-SAP汇总"); title(sheet, "OMS月结—SAP标准发票（2C）金额及数量核对汇总", "全量汇总用于总量核对；共同键范围用于开票明细核对。", "K");
  section(sheet, "A4:K4"); sheet.getRange("A4").values = [["全量及共同键核对"]]; sheet.getRange("A4:K4").merge();
  write(sheet, 4, 0, [["项目", "OMS来源系统/清单", "OMS字段/汇总逻辑", "OMS金额/数量", "SAP来源系统/清单", "SAP字段/汇总逻辑", "SAP金额/数量", "差异（SAP-OMS）", "匹配率", "匹配键", "与总览勾稽"]]); header(sheet, "A5:K5");
  write(sheet, 5, 0, [
    ["全量金额", "OMS/OMS系统日结月结查询记录SQL", "share_amount；2026年1—6月Y001月结全量汇总", ctl.oms_month_amount || 0, "SAP/SAP开票清单", "tax_amount；2026年1—6月标准发票（2C）全量汇总", ctl.sap_full_amount || 0, null, null, "全量汇总", "总览金额匹配1.总量匹配"],
    ["全量数量", "OMS/OMS系统日结月结查询记录SQL", "item_num；2026年1—6月Y001月结全量汇总", ctl.oms_month_qty || 0, "SAP/SAP开票清单", "invoice_qty；2026年1—6月标准发票（2C）全量汇总", ctl.sap_full_qty || 0, null, null, "全量汇总", "总览数量匹配1.总量匹配"],
    ["共同键金额", "OMS/OMS系统日结月结查询记录SQL", "share_amount；共同键汇总Y001", exactMap.oms_amount || 0, "SAP/SAP开票清单", "tax_amount；标准发票（2C）共同键汇总", exactMap.sap_amount || 0, null, null, "OMS销售单号+物料号+销售单位", "总览金额匹配2.开票明细匹配"],
    ["共同键数量", "OMS/OMS系统日结月结查询记录SQL", "item_num；共同键汇总Y001", exactMap.oms_qty || 0, "SAP/SAP开票清单", "invoice_qty；标准发票（2C）共同键汇总", exactMap.sap_qty || 0, null, null, "OMS销售单号+物料号+销售单位", "总览数量匹配2.开票明细匹配"],
  ]);
  for (let r = 6; r <= 9; r++) { formula(sheet, `H${r}`, `=G${r}-D${r}`); formula(sheet, `I${r}`, `=IFERROR(MIN(ABS(D${r}),ABS(G${r}))/MAX(ABS(D${r}),ABS(G${r})),0)`); }
  sapRefs.totalAmountRow = 6; sapRefs.totalQtyRow = 7; sapRefs.commonAmountRow = 8; sapRefs.commonQtyRow = 9;
  body(sheet, "A6:K9"); sheet.getRange("D6:H6").setNumberFormat("#,##0.00;[Red](#,##0.00);-"); sheet.getRange("D7:H7").setNumberFormat("#,##0"); sheet.getRange("D8:H8").setNumberFormat("#,##0.00;[Red](#,##0.00);-"); sheet.getRange("D9:H9").setNumberFormat("#,##0"); sheet.getRange("I6:I9").setNumberFormat("0.00%"); sheet.getRange("A6:K9").format.wrapText = true; sheet.getRange("A6:K9").format.rowHeight = 46; sheet.getRange("A6:K9").format.fill = C.green;
  [18, 34, 38, 20, 28, 38, 20, 20, 16, 34, 24].forEach((w, i) => sheet.getRange(`${col(i)}:${col(i)}`).format.columnWidth = w); sheet.freezePanes.freezeRows(5);
}

// 全局总览：以SAP财务确认端为起点，向OMS、惠策汇总、惠策明细及旺店通订单逐层穿透。
{
  const sheet = ws("1.全局口径与总览"); title(sheet, "销售ToC三单匹配结果总览", "审计追溯链路：SAP开票 → OMS月结 → 惠策账单汇总 → 惠策账单明细 → 旺店通订单。", "K");

  section(sheet, "A4:K4"); sheet.getRange("A4").values = [["金额匹配"]]; sheet.getRange("A4:K4").merge();
  write(sheet, 4, 0, [["核对步骤", "参与匹配数据", "匹配字段/维度", "SAP开票金额", "OMS月结金额", "惠策汇总金额", "惠策明细金额", "旺店通订单金额", "差异（前项-后项）", "金额匹配率", "匹配覆盖率"]]); header(sheet, "A5:K5");
  write(sheet, 5, 0, [
    ["1.SAP开票—OMS月结（逐条核对）", "SAP标准发票（2C）—OMS Y001月结", "OMS销售单号+物料编码+销售单位", null, null, "N/A", "N/A", "N/A", null, null, "N/A"],
    ["2.OMS月结—惠策账单汇总", "OMS Y001月结—惠策店铺汇总", "结算月份+店铺映射后的OMS客户编码", "N/A", null, null, "N/A", "N/A", null, null, "N/A"],
    ["3.惠策账单汇总—惠策账单明细", "惠策店铺汇总—惠策账单清单", "结算月份+平台+店铺", "N/A", "N/A", null, null, "N/A", null, null, "N/A"],
    ["4.惠策账单明细—旺店通订单（总量比对）", "惠策账单清单—旺店通订单清单", "账单期全量与订单追溯期全量", "N/A", "N/A", "N/A", null, null, null, null, "N/A"],
  ]);
  formula(sheet, "D6", `='10.OMS月结-SAP汇总'!G${sapRefs.commonAmountRow}`); formula(sheet, "E6", `='10.OMS月结-SAP汇总'!D${sapRefs.commonAmountRow}`); formula(sheet, "I6", "=D6-E6"); formula(sheet, "J6", "=IFERROR(MIN(ABS(D6),ABS(E6))/MAX(ABS(D6),ABS(E6)),0)");
  formula(sheet, "E7", `='6.账单-OMS月结汇总'!D${billOmsRefs.totalRow}`); formula(sheet, "F7", `='6.账单-OMS月结汇总'!C${billOmsRefs.totalRow}`); formula(sheet, "I7", "=E7-F7"); formula(sheet, "J7", "=IFERROR(MIN(ABS(E7),ABS(F7))/MAX(ABS(E7),ABS(F7)),0)");
  formula(sheet, "F8", `='4.惠策内部核对汇总'!D${internalRefs.totalRow}`); formula(sheet, "G8", `='4.惠策内部核对汇总'!C${internalRefs.totalRow}`); formula(sheet, "I8", "=F8-G8"); formula(sheet, "J8", "=IFERROR(MIN(ABS(F8),ABS(G8))/MAX(ABS(F8),ABS(G8)),0)");
  formula(sheet, "G9", `='2.订单-账单汇总'!B${orderRefs.fullBillRow}`); formula(sheet, "H9", `='2.订单-账单汇总'!B${orderRefs.fullWdtRow}`); formula(sheet, "I9", "=G9-H9"); formula(sheet, "J9", "=IFERROR(MIN(ABS(G9),ABS(H9))/MAX(ABS(G9),ABS(H9)),0)");
  const sensitivity = cutoffSensitivity.sensitivity || [];
  const sensitivityByRange = Object.fromEntries(sensitivity.map(x => [x.range, x]));
  const sensitivityDisplay = [
    ["2026年度", "5.惠策账单明细—旺店通订单（2026年度正式范围可匹配条目）", "惠策账单明细表 VS 旺店通订单清单（2026.01.01-2026.06.30）"],
    ["2025.12-2026.06", "6.惠策账单明细—旺店通订单（Cut-off截至2025年12月可匹配条目）", "惠策账单明细表 VS 旺店通订单清单（2025.12.01-2026.06.30）"],
    ["2025.11-2026.06", "7.惠策账单明细—旺店通订单（Cut-off截至2025年11月可匹配条目）", "惠策账单明细表 VS 旺店通订单清单（2025.11.01-2026.06.30）"],
  ].map(([key, step, participant]) => {
    const x = sensitivityByRange[key] || {};
    return [step, participant, "惠策平台订单号=旺店通原始单号；金额一致范围", "N/A", "N/A", "N/A", x.amount_exact_cash || 0, x.amount_exact_wdt_amount || 0, null, null, x.amount_match_rate || 0];
  });
  write(sheet, 9, 0, sensitivityDisplay);
  for (let r = 10; r <= 12; r++) { formula(sheet, `I${r}`, `=G${r}-H${r}`); formula(sheet, `J${r}`, `=IFERROR(MIN(ABS(G${r}),ABS(H${r}))/MAX(ABS(G${r}),ABS(H${r})),0)`); }
  body(sheet, "A6:K12"); sheet.getRange("A6:K12").format.borders = { preset: "all", style: "thin", color: C.line }; sheet.getRange("D6:I12").setNumberFormat("#,##0.00;[Red](#,##0.00);-"); sheet.getRange("D6:K12").format.horizontalAlignment = "right"; sheet.getRange("J6:K12").setNumberFormat("0.00%"); sheet.getRange("A6:K12").format.wrapText = true; sheet.getRange("A6:K12").format.rowHeight = 46; sheet.getRange("A10:K10").format.fill = C.green; sheet.getRange("A11:K12").format.fill = C.amber;

  section(sheet, "A15:E15"); sheet.getRange("A15").values = [["惠策系统Cut-off日期核对"]]; sheet.getRange("A15:E15").merge();
  write(sheet, 15, 0, [["日期", "惠策账单数量", "惠策账单金额", "日期字段", "核对结论"]]); header(sheet, "A16:E16");
  const cutoffRows = cutoffDaily.map(x => [new Date(`${x.business_date}T00:00:00`), x.bill_count || 0, x.bill_cash || 0, "business_date（业务日期）", "采用与总体核对一致的全量状态范围及净实收口径；记录已纳入期末核对"]); write(sheet, 16, 0, cutoffRows); body(sheet, "A17:E18"); sheet.getRange("A17:E18").format.borders = { preset: "all", style: "thin", color: C.line }; sheet.getRange("A17:A18").setNumberFormat("yyyy-mm-dd"); sheet.getRange("B17:B18").setNumberFormat("#,##0"); sheet.getRange("C17:C18").setNumberFormat("#,##0.00;[Red](#,##0.00);-"); sheet.getRange("A17:E18").format.wrapText = true; sheet.getRange("A17:E18").format.rowHeight = 46;

  section(sheet, "A20:K20"); sheet.getRange("A20").values = [["数量匹配"]]; sheet.getRange("A20:K20").merge();
  write(sheet, 20, 0, [["核对步骤", "参与匹配数据", "匹配字段/维度", "SAP数量", "OMS数量", "惠策汇总数量", "惠策明细数量", "旺店通数量", "差异（前项-后项）", "数量匹配率", "数量口径"]]); header(sheet, "A21:K21");
  write(sheet, 21, 0, [
    ["1.SAP开票—OMS月结", "SAP标准发票（2C）—OMS Y001月结", "OMS销售单号+物料编码+销售单位", null, null, "N/A", "N/A", "N/A", null, null, "双方原生商品数量；共同键范围"],
    ["2.OMS月结—惠策账单订单证据", "OMS Y001月结—惠策账单已匹配订单", "惠策平台订单号连接旺店通quantity；OMS按月份+客户+物料", "N/A", null, "N/A", "N/A", null, null, null, "惠策无原生商品数量；旺店通quantity仅作为账单订单证据"],
  ]);
  formula(sheet, "D22", `='10.OMS月结-SAP汇总'!G${sapRefs.commonQtyRow}`); formula(sheet, "E22", `='10.OMS月结-SAP汇总'!D${sapRefs.commonQtyRow}`); formula(sheet, "I22", "=D22-E22"); formula(sheet, "J22", "=IFERROR(MIN(ABS(D22),ABS(E22))/MAX(ABS(D22),ABS(E22)),0)");
  formula(sheet, "E23", `='8.数量核对汇总'!I${qtyRefs.step1Row}`); formula(sheet, "H23", `='8.数量核对汇总'!E${qtyRefs.step1Row}`); formula(sheet, "I23", "=E23-H23"); formula(sheet, "J23", "=IFERROR(MIN(ABS(E23),ABS(H23))/MAX(ABS(E23),ABS(H23)),0)");
  body(sheet, "A22:K23"); sheet.getRange("A22:K23").format.borders = { preset: "all", style: "thin", color: C.line }; sheet.getRange("D22:I23").setNumberFormat("#,##0"); sheet.getRange("D22:J23").format.horizontalAlignment = "right"; sheet.getRange("J22:J23").setNumberFormat("0.00%"); sheet.getRange("A22:K23").format.wrapText = true; sheet.getRange("A22:K23").format.rowHeight = 50;

  section(sheet, "A25:K25"); sheet.getRange("A25").values = [["Note｜本年新增穿透测试"]]; sheet.getRange("A25:K25").merge();
  const auditNotes = [
    "今年总览从SAP财务确认结果出发，依次穿透OMS业务结算、惠策店铺汇总、惠策账单明细及旺店通原始订单。",
    "新增OMS—惠策账单汇总实体级核对：以结算月份+店铺映射后的OMS客户编码汇总，不执行订单级强行匹配。",
    "新增惠策账单汇总—明细完整性核对：以结算月份+平台+店铺汇总验证两份惠策清单的实际结算金额。",
    "新增惠策账单明细—旺店通订单穿透：正式范围仅使用2026年度订单；2025年12月及11月订单仅作为Cut-off敏感性参考。匹配覆盖率=惠策金额一致订单实际实收/惠策明细全量实际实收。",
  ];
  auditNotes.forEach((note, i) => { const r = 26 + i; sheet.getRange(`A${r}:K${r}`).merge(); sheet.getRange(`A${r}`).values = [[note]]; }); body(sheet, "A26:K29"); sheet.getRange("A26:K29").format.wrapText = true; sheet.getRange("A26:K29").format.rowHeight = 30;

  section(sheet, "A31:B31"); sheet.getRange("A31").values = [["Lead｜金额"]]; sheet.getRange("A31:B31").merge();
  write(sheet, 31, 0, [["Lead", "金额"]]); header(sheet, "A32:B32");
  write(sheet, 32, 0, [["1.1 SAP开票金额", null], ["1.2 OMS月结金额", null], ["1.3 惠策账单汇总金额", null], ["1.4 惠策账单明细金额", null], ["1.5 旺店通订单金额（正式范围）", null]]);
  formula(sheet, "B33", `='10.OMS月结-SAP汇总'!G${sapRefs.totalAmountRow}`); formula(sheet, "B34", `='10.OMS月结-SAP汇总'!D${sapRefs.totalAmountRow}`); formula(sheet, "B35", `='4.惠策内部核对汇总'!D${internalRefs.totalRow}`); formula(sheet, "B36", `='4.惠策内部核对汇总'!C${internalRefs.totalRow}`); formula(sheet, "B37", `='2.订单-账单汇总'!B${orderRefs.fullWdtRow}`); body(sheet, "A33:B37"); sheet.getRange("A33:B37").format.borders = { preset: "all", style: "thin", color: C.line }; sheet.getRange("B33:B37").setNumberFormat("#,##0.00;[Red](#,##0.00);-");

  section(sheet, "A39:B39"); sheet.getRange("A39").values = [["Lead｜数量"]]; sheet.getRange("A39:B39").merge();
  write(sheet, 39, 0, [["Lead", "数量"]]); header(sheet, "A40:B40");
  write(sheet, 40, 0, [["1.1 SAP开票数量", null], ["1.2 OMS月结数量", null], ["1.3 惠策账单数量", "N/A"], ["1.4 旺店通订单数量（正式范围）", ctl.wdt_qty || 0]]);
  formula(sheet, "B41", `='10.OMS月结-SAP汇总'!G${sapRefs.totalQtyRow}`); formula(sheet, "B42", `='10.OMS月结-SAP汇总'!D${sapRefs.totalQtyRow}`); body(sheet, "A41:B44"); sheet.getRange("A41:B44").format.borders = { preset: "all", style: "thin", color: C.line }; sheet.getRange("B41:B44").setNumberFormat("#,##0");

  section(sheet, "A46:K46"); sheet.getRange("A46").values = [["Notes:"]]; sheet.getRange("A46:K46").merge();
  const sourceNotes = [
    "SAP开票：SAP开票清单tax_amount及invoice_qty；过滤2026年1—6月标准发票（2C）。OMS月结：OMS SQL的share_amount及item_num；过滤2026年1—6月Y001月结单据。",
    "惠策汇总：惠策店铺汇总清单实际结算金额=成功金额+不一致实收金额+单边实收金额。惠策明细：收款金额-退款金额净实收。",
    "旺店通订单：正式匹配仅采用2026-01-01至2026-06-30订单，金额按平台订单号归集allocated_total；可匹配条目以原始单号连接惠策平台订单号。",
    "惠策清单未提供商品编码及原生商品数量，因此惠策数量在Lead中列示N/A；数量核对仅保留SAP—OMS共同键及OMS—账单订单证据两项具有合理计量基础的比较。",
    "截至2026年6月30日，惠策系统未生成当日账单记录，因此选取截止日前最近两个业务日期2026年6月28日及2026年6月29日账单数据执行Cut-off测试，以验证期末附近销售业务记录完整性。",
    "本年度销售ToC三单匹配正式范围为2026年度旺店通订单数据。考虑期末Cut-off影响，额外扩展获取2025年12月及2025年11月旺店通订单数据，并按照相同匹配逻辑执行补充测试，用于向Audit Team展示订单范围变化对匹配覆盖率的影响。扩展数据仅用于Cut-off分析，不纳入正式匹配结果。",
  ];
  sourceNotes.forEach((note, i) => { const r = 47 + i; sheet.getRange(`A${r}:K${r}`).merge(); sheet.getRange(`A${r}`).values = [[note]]; }); body(sheet, "A47:K52"); sheet.getRange("A47:K52").format.wrapText = true; sheet.getRange("A47:K52").format.rowHeight = 32;

  [31, 32, 40, 20, 20, 21, 21, 21, 21, 17, 17].forEach((w, i) => sheet.getRange(`${col(i)}:${col(i)}`).format.columnWidth = w); sheet.freezePanes.freezeRows(5);
}

// 正式范围旺店通订单明细：总览保留控制数、规则及全量分片索引；完整记录存放于同目录CSV分片。
{
  const sheet = ws("旺店通订单匹配明细");
  title(sheet, "2026年度正式范围旺店通订单匹配明细", "惠策账单明细表 ↔ 旺店通订单清单；仅含2026年度正式范围，不含2025年11月及12月Cut-off扩展数据。", "H");
  const categoryMap = Object.fromEntries((formalDetailManifest.categories || []).map(item => [item.category, item]));
  const matched = categoryMap["可匹配条目"] || {};
  const billOnly = categoryMap["仅账单未匹配"] || {};
  const orderOnly = categoryMap["仅订单未匹配"] || {};

  section(sheet, "A4:H4"); sheet.getRange("A4").values = [["汇总统计及总览勾稽"]]; sheet.getRange("A4:H4").merge();
  write(sheet, 4, 0, [["匹配状态", "平台订单数量", "惠策实际实收", "旺店通订单金额"]]); header(sheet, "A5:D5");
  write(sheet, 5, 0, [
    ["总惠策账单明细", ctl.huice_orders || 0, ctl.huice_bill_cash || 0, "N/A"],
    ["总旺店通订单", ctl.wdt_orders || 0, "N/A", ctl.wdt_allocated_amount || 0],
    ["成功匹配（订单号及金额一致）", matched.rows || 0, matched.huice_amount || 0, matched.wdt_amount || 0],
    ["仅账单/账单未成功匹配", billOnly.rows || 0, billOnly.huice_amount || 0, billOnly.wdt_amount || 0],
    ["仅订单未匹配", orderOnly.rows || 0, orderOnly.huice_amount || 0, orderOnly.wdt_amount || 0],
  ]);
  body(sheet, "A6:D10"); sheet.getRange("A6:D10").format.borders = { preset: "all", style: "thin", color: C.line }; sheet.getRange("B6:B10").setNumberFormat("#,##0"); sheet.getRange("C6:D10").setNumberFormat("#,##0.00;[Red](#,##0.00);-"); sheet.getRange("A8:D8").format.fill = C.green; sheet.getRange("A9:D10").format.fill = C.amber;

  section(sheet, "A12:H12"); sheet.getRange("A12").values = [["匹配规则说明"]]; sheet.getRange("A12:H12").merge();
  const ruleNotes = [
    "1. 匹配起点：惠策平台订单号=旺店通原始单号；本明细不执行店铺、商品编码、日期或金额的模糊组合补配。",
    "2. 金额一致优先级：先比较旺店通平台订单分摊金额allocated_total与惠策净实收；未通过时，再比较旺店通订单头金额receivable_amount；差异不超过0.01元视为一致。",
    "3. 可匹配条目：订单号命中且上述任一金额口径一致；其数量及金额与总览“2026年度正式范围可匹配条目”直接勾稽。",
    "4. 仅账单/账单未成功匹配：包括旺店通正式范围无对应订单、订单号为空、订单日期在正式范围外，以及订单号命中但金额不一致。",
    "5. 仅订单未匹配：包括惠策无对应账单、旺店通平台订单号为空及同内部订单零金额附属单。",
    "6. 多对一/一对多：按平台订单号汇总后比较；同一内部订单包含多个平台订单号时分别列示，并保留旺店通内部订单号集合。",
  ];
  ruleNotes.forEach((note, index) => { const row = 13 + index; sheet.getRange(`A${row}:H${row}`).merge(); sheet.getRange(`A${row}`).values = [[note]]; });
  body(sheet, "A13:H18"); sheet.getRange("A13:H18").format.wrapText = true; sheet.getRange("A13:H18").format.rowHeight = 30;

  section(sheet, "A20:H20"); sheet.getRange("A20").values = [["全量明细文件索引"]]; sheet.getRange("A20:H20").merge();
  write(sheet, 20, 0, [["类别", "部分", "起始序号", "结束序号", "明细条数", "惠策实际实收", "旺店通订单金额", "文件名"]]); header(sheet, "A21:H21");
  const fileRows = (formalDetailManifest.categories || []).flatMap(category => (category.files || []).map((file, index) => [
    category.category, index + 1, file.start, file.end, file.rows, file.huice_amount, file.wdt_amount, path.basename(file.file),
  ]));
  write(sheet, 21, 0, fileRows);
  const fileEnd = 21 + Math.max(1, fileRows.length); body(sheet, `A22:H${fileEnd}`); sheet.getRange(`A22:H${fileEnd}`).format.borders = { preset: "all", style: "thin", color: C.line }; sheet.getRange(`B22:E${fileEnd}`).setNumberFormat("#,##0"); sheet.getRange(`F22:G${fileEnd}`).setNumberFormat("#,##0.00;[Red](#,##0.00);-"); sheet.getRange(`H22:H${fileEnd}`).format.wrapText = true;

  const noteRow = fileEnd + 2;
  section(sheet, `A${noteRow}:H${noteRow}`); sheet.getRange(`A${noteRow}`).values = [["Notes:"]]; sheet.getRange(`A${noteRow}:H${noteRow}`).merge();
  const detailNotes = [
    "完整明细共5,896,663条。受Excel单Sheet最大1,048,576行限制，按每个CSV不超过1,000,000条拆分；CSV可由Excel直接打开并筛选。",
    "惠策账单清单未提供商品编码和商品名称；商品编码取自旺店通商品明细，商品名称由OMS物料名称按商品编码辅助映射。惠策单边记录相应字段显示N/A。",
    "全量文件位于主底稿同目录“2026年度正式范围旺店通订单匹配明细/明细数据”；文件级行数、金额及SHA-256校验值详见“明细拆分索引及校验结果.json”。",
  ];
  detailNotes.forEach((note, index) => { const row = noteRow + 1 + index; sheet.getRange(`A${row}:H${row}`).merge(); sheet.getRange(`A${row}`).values = [[note]]; }); body(sheet, `A${noteRow + 1}:H${noteRow + 3}`); sheet.getRange(`A${noteRow + 1}:H${noteRow + 3}`).format.wrapText = true; sheet.getRange(`A${noteRow + 1}:H${noteRow + 3}`).format.rowHeight = 32;

  [30, 12, 16, 16, 16, 20, 20, 48].forEach((width, index) => sheet.getRange(`${col(index)}:${col(index)}`).format.columnWidth = width);
  sheet.freezePanes.freezeRows(5);
}

// 精简后的支持页：只保留与最终口径直接相关的字段。
detail("5.惠策内部核对明细", "惠策明细—惠策店铺汇总实际实收核对明细", "按结算月份+平台+店铺逐项核对实际实收金额。", data.internal, { huice_shop: 34, result: 20 }, ["detail_success_amount", "summary_success_amount", "success_difference", "detail_receivable", "summary_receivable", "receivable_difference", "historical_rows", "historical_receivable", "historical_cash"]);
detail("7.账单-OMS月结明细", "惠策店铺汇总—OMS月结实际结算核对明细", "按月份+OMS客户编码逐项核对惠策实际结算金额与OMS结算金额。", data.billOms, { huice_shop: 34, customer_name: 40, result: 24 }, ["mapping_status", "mapping_source", "bill_record_count", "success_count", "bill_success_amount", "bill_receivable", "success_difference", "receivable_difference", "sap_assisted_qty", "sap_assisted_amount", "sap_success_difference"]);
{
  const sheet = ws("9.数量核对明细"); title(sheet, "销售数量逐层核对明细", "Step 1为惠策账单订单证据—OMS；Step 2为OMS—SAP共同键核对。", "L");
  const select = (dataset, fields) => { const indexes = fields.map(h => dataset.headers.indexOf(h)); return dataset.rows.map(row => indexes.map(i => row[i])); };
  const step1Fields = ["ship_month", "wdt_shop", "customer_code", "customer_name", "material_code", "order_bill_qty", "oms_qty", "qty_difference", "result"];
  const step1Rows = select(data.qty, step1Fields); section(sheet, "A4:I4"); sheet.getRange("A4").values = [["Step 1｜惠策账单已匹配订单对应的旺店通商品数量 vs OMS月结数量"]]; sheet.getRange("A4:I4").merge(); write(sheet, 4, 0, [step1Fields.map(h => headerZh[h] ? `${h}\n${headerZh[h]}` : h)]); header(sheet, "A5:I5"); sheet.getRange("A5:I5").format.rowHeight = 36;
  for (let i = 0; i < step1Rows.length; i += 3000) write(sheet, 5 + i, 0, step1Rows.slice(i, i + 3000)); const step1End = 5 + step1Rows.length; body(sheet, `A6:I${step1End}`); status(sheet, `I6:I${step1End}`); sheet.getRange(`C6:C${step1End}`).setNumberFormat("@"); sheet.getRange(`E6:E${step1End}`).setNumberFormat("@"); sheet.getRange(`F6:H${step1End}`).setNumberFormat("#,##0");
  const step2Section = step1End + 2; const step2Header = step2Section + 1; const step2Start = step2Section + 2;
  const step2Fields = ["outbound_month", "file_month", "oms_sales_no", "material_code", "sales_unit", "customer_code", "customer_name", "oms_qty", "sap_qty", "quantity_difference", "mapping_result", "source_result"];
  const step2Rows = select(data.omsSap, step2Fields); section(sheet, `A${step2Section}:L${step2Section}`); sheet.getRange(`A${step2Section}`).values = [["Step 2｜OMS月结数量 vs SAP标准发票（2C）数量"]]; sheet.getRange(`A${step2Section}:L${step2Section}`).merge(); write(sheet, step2Header - 1, 0, [step2Fields.map(h => headerZh[h] ? `${h}\n${headerZh[h]}` : h)]); header(sheet, `A${step2Header}:L${step2Header}`); sheet.getRange(`A${step2Header}:L${step2Header}`).format.rowHeight = 36;
  for (let i = 0; i < step2Rows.length; i += 3000) write(sheet, step2Start - 1 + i, 0, step2Rows.slice(i, i + 3000)); const step2End = step2Start + step2Rows.length - 1; body(sheet, `A${step2Start}:L${step2End}`); status(sheet, `K${step2Start}:L${step2End}`); for (const c of ["C","D","F"]) sheet.getRange(`${c}${step2Start}:${c}${step2End}`).setNumberFormat("@"); sheet.getRange(`H${step2Start}:J${step2End}`).setNumberFormat("#,##0");
  [14, 28, 18, 36, 17, 18, 38, 18, 18, 18, 24, 20].forEach((w, i) => sheet.getRange(`${col(i)}:${col(i)}`).format.columnWidth = w); sheet.freezePanes.freezeRows(5); sheet.freezePanes.freezeColumns(2);
}
detail("13.店铺客户映射", "惠策店铺—OMS客户映射", "仅展示店铺与OMS客户编码映射结果，金额口径以实际结算为准。", data.shop, { huice_shop: 34, customer_name: 40 }, ["mapping_status", "mapping_source", "bill_record_count", "bill_receivable", "source_rows", "success_count", "bill_success_amount"]);

await fs.mkdir(outputDir, { recursive: true });
const previews = path.join(root, "reconciliation/qa_previews"); await fs.mkdir(previews, { recursive: true });
console.log("OVERVIEW\n" + (await wb.inspect({ kind: "table", range: "1.全局口径与总览!A1:K52", include: "values,formulas", tableMaxRows: 56, tableMaxCols: 12, maxChars: 48000 })).ndjson);
console.log("ORDER_SUMMARY\n" + (await wb.inspect({ kind: "table", range: "2.订单-账单汇总!A1:H28", include: "values,formulas", tableMaxRows: 30, tableMaxCols: 9, maxChars: 18000 })).ndjson);
console.log("HUICE_INTERNAL\n" + (await wb.inspect({ kind: "table", range: "4.惠策内部核对汇总!A1:H16", include: "values,formulas", tableMaxRows: 18, tableMaxCols: 9, maxChars: 12000 })).ndjson);
console.log("BILL_OMS\n" + (await wb.inspect({ kind: "table", range: "6.账单-OMS月结汇总!A1:H22", include: "values,formulas", tableMaxRows: 24, tableMaxCols: 9, maxChars: 14000 })).ndjson);
console.log("FORMAL_ORDER_DETAIL\n" + (await wb.inspect({ kind: "table", range: "旺店通订单匹配明细!A1:H34", include: "values,formulas", tableMaxRows: 36, tableMaxCols: 9, maxChars: 18000 })).ndjson);
console.log("ERRORS\n" + (await wb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 300 }, summary: "formula errors", maxChars: 5000 })).ndjson);
for (const name of names) { const sheet = ws(name), used = sheet.getUsedRange(true), maxCols = Math.min(used.columnCount || 8, name.includes("明细") || name.includes("映射") ? 10 : 12), maxRows = name === "1.全局口径与总览" ? 68 : name.includes("明细") ? 20 : 30; const blob = await wb.render({ sheetName: name, range: `A1:${col(maxCols - 1)}${maxRows}`, scale: 1.15, format: "png" }); await fs.writeFile(path.join(previews, `${name}.png`), new Uint8Array(await blob.arrayBuffer())); }
const qtyStep2Section = 7 + data.qty.rows.length; const qtyStep2Preview = await wb.render({ sheetName: "9.数量核对明细", range: `A${qtyStep2Section}:L${qtyStep2Section + 17}`, scale: 1.15, format: "png" }); await fs.writeFile(path.join(previews, "9.数量核对明细-Step2.png"), new Uint8Array(await qtyStep2Preview.arrayBuffer()));
const out = await SpreadsheetFile.exportXlsx(wb); await out.save(outputFile); console.log(JSON.stringify({ outputFile, sheets: names, previews }, null, 2));
