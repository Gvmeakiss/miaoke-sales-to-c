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
const settlementOms = await read("settlement_oms_workpaper.json").catch(() => ({ definitions: {}, summary_rows: [], available_total: {}, detail_headers: [], detail_rows: [] }));
const data = {
  internal: await read("huice_internal_recon_workbook.json"),
  omsSap: await read("oms_sap_field_map_workbook.json"),
  shop: await read("huice_shop_map_workbook.json"),
};

const wb = Workbook.create();
const names = [
  "1.全局口径与总览",
  "2.OMS月结-SAP汇总",
  "3.对账结果-OMS月结汇总",
  "4.对账结果-OMS月结明细",
  "5.数量核对汇总",
  "6.数量核对明细",
  "7.旺店通订单匹配明细",
  "8.惠策内部核对汇总",
  "9.惠策内部核对明细",
  "10.店铺客户映射",
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
  x.conditionalFormats.add("containsText", { text: "待获取", format: { fill: C.amber, font: { bold: true, color: C.amberText } } });
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
  month: "月份", match_category: "匹配分类", shop_customer_code: "店铺/OMS客户编码", wdt_shop_name: "旺店通店铺名称", oms_customer_name: "OMS客户名称",
  reconciliation_rows: "发货对账行数", reconciliation_amount: "发货对账收款金额", reconciliation_quantity: "发货对账实际数量", oms_quantity: "OMS月结数量", amount_match_rate: "金额匹配率", quantity_match_rate: "数量匹配率",
  source_files: "来源文件", settlement_periods: "原账期范围", business_months: "业务发生月份", cross_period_flag: "跨期补充标识",
  cross_period_adjustment_amount: "跨期结算调整金额", adjusted_reconciliation_amount: "调整后发货对账金额", adjusted_amount_difference: "调整后金额差异", adjusted_amount_match_rate: "调整后金额匹配率", adjustment_basis: "跨期调整依据",
};
function detail(name, reportTitle, subtitle, dataset, widths = {}, omit = []) {
  const keep = dataset.headers.map((h, i) => ({ h, i })).filter(x => !omit.includes(x.h));
  const headers = keep.map(x => x.h);
  const rows = keep.length === dataset.headers.length ? dataset.rows : dataset.rows.map(r => keep.map(x => r[x.i]));
  const sheet = ws(name), last = col(headers.length - 1), titleLast = col(Math.min(headers.length, 10) - 1);
  title(sheet, reportTitle, subtitle, titleLast);
  write(sheet, 3, 0, [headers.map(h => headerZh[h] ? `${h}\n${headerZh[h]}` : h)]); header(sheet, `A4:${last}4`); sheet.getRange(`A4:${last}4`).format.rowHeight = 34;
  for (let i = 0; i < rows.length; i += 3000) write(sheet, 4 + i, 0, rows.slice(i, i + 3000));
  const end = 4 + rows.length; if (rows.length) { body(sheet, `A5:${last}${end}`); for (const key of ["result", "mapping_result"]) { const index = headers.indexOf(key); if (index >= 0) status(sheet, `${col(index)}5:${col(index)}${end}`); } const crossIndex = headers.indexOf("cross_period_flag"); if (crossIndex >= 0) sheet.getRange(`${col(crossIndex)}5:${col(crossIndex)}${end}`).conditionalFormats.add("containsText", { text: "是", format: { fill: C.amber, font: { bold: true, color: C.amberText } } }); }
  const textFields = new Set(["customer_code", "shop_customer_code", "material_code", "platform_order_no", "oms_sales_no", "sap_invoice_nos", "reconcile_ids", "source_files", "settlement_periods", "business_months", "cross_period_flag", "adjustment_basis"]);
  headers.forEach((h, i) => { const width = widths[h] || (/name|shop|orders|invoice|reconcile/.test(h) ? 25 : /amount|difference|cash/.test(h) ? 17 : 14); sheet.getRange(`${col(i)}:${col(i)}`).format.columnWidth = width; if (textFields.has(h)) sheet.getRange(`${col(i)}5:${col(i)}${end}`).setNumberFormat("@"); else if (/rate/.test(h)) sheet.getRange(`${col(i)}5:${col(i)}${end}`).setNumberFormat("0.00%"); else if (/amount|difference|cash/.test(h)) sheet.getRange(`${col(i)}5:${col(i)}${end}`).setNumberFormat("#,##0.00;[Red](#,##0.00);-"); else if (/qty|quantity|count|rows|lines|groups|keys|docs/.test(h)) sheet.getRange(`${col(i)}5:${col(i)}${end}`).setNumberFormat("#,##0.00"); });
  sheet.freezePanes.freezeRows(4); sheet.freezePanes.freezeColumns(Math.min(2, headers.length));
}
const by = (arr, key) => Object.fromEntries(arr.map(x => [x[key], x]));
const ctl = S.controls, orderResults = S.order_bill_results, internalResults = S.huice_internal_results;
const omsSapResults = S.oms_sap_results, exactMap = by(omsSapResults, "mapping_result")["双向字段一致"] || { keys: 0, sap_qty: 0, oms_qty: 0, sap_amount: 0, oms_amount: 0 };
const findResult = (rows, label) => rows.find(x => x.result === label) || { groups: 0, wdt_amount: 0, bill_cash: 0, detail_cash: 0, summary_cash: 0, oms_amount: 0, oms_qty: 0, order_bill_qty: 0 };
const findMap = label => omsSapResults.find(x => x.mapping_result === label) || { keys: 0, sap_qty: 0, oms_qty: 0, sap_amount: 0, oms_amount: 0 };
const orderMatchCategories = ["单号分摊实收一致", "单号订单实收一致"];
const orderCategories = ["单号分摊实收一致", "单号订单实收一致", "单号一致金额差异", "仅账单", "仅订单"];
const matchedWdtAmount = orderMatchCategories.reduce((sum, label) => { const x = findResult(orderResults, label); return sum + (label === "单号订单实收一致" ? (x.wdt_header_amount || 0) : (x.wdt_amount || 0)); }, 0);
const matchedBillCash = orderMatchCategories.reduce((sum, label) => sum + (findResult(orderResults, label).bill_cash || 0), 0);
const matchedOrderGroups = orderMatchCategories.reduce((sum, label) => sum + (findResult(orderResults, label).groups || 0), 0);
const fullWdtComparisonAmount = orderCategories.reduce((sum, label) => sum + (findResult(orderResults, label).wdt_amount || 0), 0);
const orderDifference = findResult(orderResults, "单号一致金额差异");
const participatingBillCash = matchedBillCash + (orderDifference.bill_cash || 0);
const participatingWdtAmount = matchedWdtAmount + (orderDifference.wdt_amount || 0);
const cutoffDecember = (cutoffSensitivity.sensitivity || []).find(x => x.range === "2025.12-2026.06") || {};
const settlementSummary = settlementOms.summary_rows || [];
const settlementTotal = settlementOms.available_total || {};
const settlementDetail = { headers: settlementOms.detail_headers || [], rows: settlementOms.detail_rows || [] };
const monthLabel = month => String(month || "").replace(/^(\d{4})-(\d{2})$/, (_, year, value) => `${year}年${Number(value)}月`);
const availableMonthsLabel = (settlementTotal.available_months || []).map(monthLabel).join("、") || "暂无";
const pendingMonthsLabel = (settlementTotal.pending_months || settlementSummary.filter(x => !String(x.status || "").startsWith("已完成核对")).map(x => x.month)).map(monthLabel).join("、") || "无";
const hasPendingSettlementMonths = pendingMonthsLabel !== "无";
const settlementRangeNote = `已完成月份：${availableMonthsLabel}${hasPendingSettlementMonths ? `；待补月份：${pendingMonthsLabel}` : "；1—6月资料已齐"}`;

// 场景2：惠策明细—店铺汇总，仅保留实际实收。
const internalRefs = {};
{
  const sheet = ws("8.惠策内部核对汇总"); title(sheet, "惠策明细—惠策店铺汇总实际实收核对汇总", "两份惠策清单均按导出结算月份+平台+店铺核对。", "H");
  section(sheet, "A4:H4"); sheet.getRange("A4").values = [["金额字段定义"]]; sheet.getRange("A4:H4").merge();
  write(sheet, 4, 0, [["金额项目", "来源清单", "原始字段", "计算逻辑", "总额", "勾稽用途", "对应总览项目", "备注"]]); header(sheet, "A5:H5");
  write(sheet, 5, 0, [["惠策明细实际实收", "惠策账单清单（含历史遗漏补充）", "收款金额（正实收）；退款金额（负实收）", "按对账流水号去重后计算净额=收款金额-退款金额，再按结算月+平台+店铺汇总", null, "明细侧", "惠策内部", "合并后唯一账单记录"], ["惠策汇总实际结算", "惠策店铺汇总清单+历史遗漏补充", "成功金额；不一致实收金额；单边实收金额；补充明细净实收", "原店铺汇总实际结算金额+补充明细净实收，再按结算月+平台+店铺汇总", null, "汇总侧", "惠策内部", "核对层重构，不修改原文件"]]); body(sheet, "A6:H7"); sheet.getRange("A6:H7").format.wrapText = true; sheet.getRange("A6:H7").format.rowHeight = 36;
  formula(sheet, "E6", `=SUM(C11:C${10 + Math.max(1, internalResults.length)})`); formula(sheet, "E7", `=SUM(D11:D${10 + Math.max(1, internalResults.length)})`); sheet.getRange("E6:E7").setNumberFormat("#,##0.00;[Red](#,##0.00);-");
  section(sheet, "A9:H9"); sheet.getRange("A9").values = [["分类金额及匹配率"]]; sheet.getRange("A9:H9").merge(); write(sheet, 9, 0, [["核对结果", "月份/平台/店铺组合数", "惠策明细实际实收", "惠策汇总实际结算", "差异（汇总-明细）", "金额匹配率", "分类处理", "备注"]]); header(sheet, "A10:H10");
  const rows = internalResults.length ? internalResults.map(x => [x.result, x.groups || 0, x.detail_cash || 0, x.summary_cash || 0, null, null, x.result === "实收一致" ? "金额一致分类" : "差异分类", "按结算月份+平台+店铺核对"]) : [["无记录", 0, 0, 0, null, null, "分类列示", ""]]; write(sheet, 10, 0, rows); const first = 11, last = 10 + rows.length, total = last + 1; internalRefs.totalRow = total; internalRefs.categoryRows = Object.fromEntries(rows.map((x, i) => [x[0], first + i]));
  for (let r = first; r <= last; r++) { formula(sheet, `E${r}`, `=D${r}-C${r}`); formula(sheet, `F${r}`, `=IFERROR(MIN(ABS(C${r}),ABS(D${r}))/MAX(ABS(C${r}),ABS(D${r})),0)`); }
  write(sheet, total - 1, 0, [["合计", null, null, null, null, null, "分类合计与总览勾稽", ""]]); formula(sheet, `B${total}`, `=SUM(B${first}:B${last})`); formula(sheet, `C${total}`, `=SUM(C${first}:C${last})`); formula(sheet, `D${total}`, `=SUM(D${first}:D${last})`); formula(sheet, `E${total}`, `=D${total}-C${total}`); formula(sheet, `F${total}`, `=IFERROR(MIN(ABS(C${total}),ABS(D${total}))/MAX(ABS(C${total}),ABS(D${total})),0)`); body(sheet, `A11:H${total}`); status(sheet, `A11:A${last}`); sheet.getRange(`B11:B${total}`).setNumberFormat("#,##0"); sheet.getRange(`C11:E${total}`).setNumberFormat("#,##0.00;[Red](#,##0.00);-"); sheet.getRange(`F11:F${total}`).setNumberFormat("0.00%"); sheet.getRange(`A${total}:H${total}`).format = { fill: C.green, font: { bold: true, color: C.greenText }, borders: { top: { style: "medium", color: C.blue } } }; [30, 18, 20, 20, 20, 17, 22, 28].forEach((w, i) => sheet.getRange(`${col(i)}:${col(i)}`).format.columnWidth = w); sheet.getRange("A6:H7").format.wrapText = true; sheet.freezePanes.freezeRows(10);
}

// 场景3：发货对账结果—OMS月结。替代原惠策账单—OMS月结主链生态位。
const settlementOmsRefs = {};
{
  const sheet = ws("3.对账结果-OMS月结汇总"); title(sheet, "发货对账结果—OMS月结Y001核对汇总", `按月份+店铺/客户编码+货品/商品编码核对；客户确认跨期结算差异已作桥接调整；${settlementRangeNote}。`, "O");
  section(sheet, "A4:O4"); sheet.getRange("A4").values = [["月份核对结果"]]; sheet.getRange("A4:O4").merge();
  write(sheet, 4, 0, [["月份", "资料状态", "发货对账执行总金额", "OMS月结执行总金额", "发货对账原共同键金额", "跨期结算调整金额", "发货对账调整后共同键金额", "OMS月结共同键金额", "调整后差异（发货对账-OMS）", "调整后金额匹配率", "OMS金额覆盖率", "发货对账单边金额", "OMS单边金额", "来源文件/原账期", "核对口径及说明"]]); header(sheet, "A5:O5");
  const monthRows = settlementSummary.map(x => [
    x.month, x.status, x.reconciliation_total_amount ?? null, x.oms_total_amount ?? null,
    x.reconciliation_common_amount ?? null, x.cross_period_adjustment_amount ?? null,
    x.adjusted_reconciliation_common_amount ?? null, x.oms_common_amount ?? null,
    x.adjusted_common_amount_difference ?? null, x.adjusted_common_amount_match_rate ?? null,
    x.oms_amount_coverage ?? null, x.reconciliation_only_amount ?? null, x.oms_only_amount ?? null,
    [x.source_files, x.settlement_periods].filter(Boolean).join("｜"),
    String(x.status || "").startsWith("待获取") ? `${x.status}；不参与合计及匹配率` : `${x.match_dimension}${x.cross_period_files ? `；含跨期补充：${x.cross_period_files}` : ""}`,
  ]);
  write(sheet, 5, 0, monthRows); const first = 6, last = 5 + monthRows.length, total = last + 1; settlementOmsRefs.totalRow = total;
  write(sheet, total - 1, 0, [[`已完成月份合计（${availableMonthsLabel}）`, "仅汇总资料齐备月份", settlementTotal.reconciliation_total_amount ?? null, settlementTotal.oms_total_amount ?? null, settlementTotal.reconciliation_common_amount ?? null, settlementTotal.cross_period_adjustment_amount ?? null, settlementTotal.adjusted_reconciliation_common_amount ?? null, settlementTotal.oms_common_amount ?? null, settlementTotal.adjusted_common_amount_difference ?? null, settlementTotal.adjusted_common_amount_match_rate ?? null, settlementTotal.oms_amount_coverage ?? null, settlementTotal.reconciliation_only_amount ?? null, settlementTotal.oms_only_amount ?? null, settlementRangeNote, `跨期桥接调整不覆盖原始金额；跨文件重复业务行剔除${settlementTotal.duplicate_rows_removed || 0}条`]]);
  body(sheet, `A6:O${total}`); status(sheet, `B6:B${last}`); monthRows.forEach((row, index) => { const targetRow = first + index; if (String(row[1] || "").includes("含跨期补充")) { sheet.getRange(`B${targetRow}:B${targetRow}`).format = { fill: C.amber, font: { bold: true, color: C.amberText } }; sheet.getRange(`N${targetRow}:O${targetRow}`).format.fill = C.amber; } else if (String(row[1] || "").startsWith("已完成核对")) { sheet.getRange(`B${targetRow}:B${targetRow}`).format = { fill: C.green, font: { bold: true, color: C.greenText } }; } }); sheet.getRange(`A${total}:O${total}`).format = { fill: C.green, font: { bold: true, color: C.greenText }, borders: { top: { style: "medium", color: C.blue } } };
  sheet.getRange(`C6:I${total}`).setNumberFormat("#,##0.00;[Red](#,##0.00);-"); sheet.getRange(`L6:M${total}`).setNumberFormat("#,##0.00;[Red](#,##0.00);-"); sheet.getRange(`J6:K${total}`).setNumberFormat("0.00%"); sheet.getRange(`A6:O${total}`).format.wrapText = true; sheet.getRange(`A6:O${total}`).format.rowHeight = 52; sheet.getRange(`A${total}:O${total}`).format.rowHeight = 68;
  section(sheet, `A${total + 2}:O${total + 2}`); sheet.getRange(`A${total + 2}`).values = [["字段与口径"]]; sheet.getRange(`A${total + 2}:O${total + 2}`).merge();
  const definitionRows = [
    ["发货对账结果", "对账明细（to oms 月结）", "收款金额、实际数量", `${settlementOms.definitions?.reconciliation_filter || "对账状态=对账成功"}；跨月导出仅纳入账期截止不晚于2026-06-30`, "按结算月+店铺编码+货品编码汇总；允许6月末记录在7月初完成对账", "与OMS共同键核对"],
    ["OMS月结", "OMS系统日结月结查询记录SQL", "share_amount、item_num", settlementOms.definitions?.oms_filter || "业务类型=Y001", "按月+客户编码+商品编码汇总", "与发货对账共同键核对"],
    ["跨期结算桥接调整", "客户确认+两侧共同键明细", "业务时间、对账时间、账期结束日期、共同键金额差异", "仅限具备跨期证据的共同键；单边记录不调整", "调整金额=OMS共同键金额-发货对账原共同键金额；保留原始金额及原差异", "形成调整后可比金额"],
  ];
  write(sheet, total + 2, 0, [["金额项目", "来源清单", "原始字段", "过滤条件", "汇总逻辑", "勾稽用途"]]); header(sheet, `A${total + 3}:F${total + 3}`); write(sheet, total + 3, 0, definitionRows); body(sheet, `A${total + 4}:F${total + 6}`); sheet.getRange(`A${total + 4}:F${total + 6}`).format.wrapText = true; sheet.getRange(`A${total + 4}:F${total + 6}`).format.rowHeight = 44;
  [18, 26, 19, 19, 19, 18, 21, 19, 20, 17, 17, 19, 19, 44, 46].forEach((w, i) => sheet.getRange(`${col(i)}:${col(i)}`).format.columnWidth = w); sheet.freezePanes.freezeRows(5);
}

// 场景4：数量沿业务流分两步核对，不进行四方直接比较。
const qtyRefs = {};
{
  const sheet = ws("5.数量核对汇总"); title(sheet, "销售数量逐层核对汇总", "Step 1为发货对账结果—OMS月结（仅已获取月份）；Step 2为OMS月结—SAP标准发票（2C）。", "K");
  section(sheet, "A4:K4"); sheet.getRange("A4").values = [["Step 1｜OMS月结数量 vs 发货对账实际数量"]]; sheet.getRange("A4:K4").merge();
  write(sheet, 4, 0, [["核对步骤", "左侧数量项目", "来源系统/原始清单", "使用字段/汇总逻辑", "左侧数量", "右侧数量项目", "来源系统/原始清单", "使用字段/汇总逻辑", "右侧数量", "差异数量（右-左）", "差异率"]]); header(sheet, "A5:K5");
  write(sheet, 5, 0, [["Step 1", "OMS月结共同键数量", "OMS系统日结月结查询记录SQL", `item_num；Y001按月份+客户编码+商品编码共同键汇总；${settlementRangeNote}`, settlementTotal.oms_common_quantity ?? null, "发货对账共同键实际数量", "对账明细（to oms 月结）", `实际数量；对账成功记录按月份+店铺编码+货品编码共同键汇总；${settlementRangeNote}`, settlementTotal.reconciliation_common_quantity ?? null, null, null]]); formula(sheet, "J6", "=I6-E6"); formula(sheet, "K6", "=IFERROR(ABS(J6)/MAX(ABS(E6),ABS(I6)),0)"); qtyRefs.step1Row = 6;
  section(sheet, "A8:K8"); sheet.getRange("A8").values = [["Step 2｜OMS月结数量 vs SAP标准发票（2C）数量"]]; sheet.getRange("A8:K8").merge();
  write(sheet, 8, 0, [["核对步骤", "左侧数量项目", "来源系统/原始清单", "使用字段/汇总逻辑", "左侧数量", "右侧数量项目", "来源系统/原始清单", "使用字段/汇总逻辑", "右侧数量", "差异数量（右-左）", "差异率"]]); header(sheet, "A9:K9");
  write(sheet, 9, 0, [["Step 2", "OMS共同键数量", "OMS系统日结月结查询记录SQL", "item_num；销售单号+物料+销售单位共同键汇总", exactMap.oms_qty || 0, "SAP共同键数量", "SAP开票清单", "invoice_qty；标准发票（2C）共同键汇总", exactMap.sap_qty || 0, null, null]]); formula(sheet, "J10", "=I10-E10"); formula(sheet, "K10", "=IFERROR(ABS(J10)/MAX(ABS(E10),ABS(I10)),0)"); qtyRefs.step2Row = 10;
  body(sheet, "A6:K6"); body(sheet, "A10:K10"); sheet.getRange("E6:E10").setNumberFormat("#,##0"); sheet.getRange("I6:J10").setNumberFormat("#,##0"); sheet.getRange("K6:K10").setNumberFormat("0.00%"); sheet.getRange("A6:K10").format.wrapText = true; sheet.getRange("A6:K6").format.rowHeight = 48; sheet.getRange("A10:K10").format.rowHeight = 48;
  write(sheet, 12, 0, [["明细支持", "6.数量核对明细 Step 1", "OMS月结—发货对账结果", "按月份+客户/店铺编码+商品/货品编码", null, "6.数量核对明细 Step 2", "OMS—SAP", "按销售单号+物料+销售单位", null, null, null]]); body(sheet, "A13:K13");
  [14, 34, 30, 38, 17, 25, 30, 38, 17, 20, 16].forEach((w, i) => sheet.getRange(`${col(i)}:${col(i)}`).format.columnWidth = w); sheet.freezePanes.freezeRows(5);
}

// 场景5：OMS—SAP汇总。同时列示全量与共同键范围，供去年式总览直接勾稽。
const sapRefs = {};
{
  const sheet = ws("2.OMS月结-SAP汇总"); title(sheet, "OMS月结—SAP标准发票（2C）金额及数量核对汇总", "全量汇总用于总量核对；共同键范围用于开票明细核对。", "K");
  section(sheet, "A4:K4"); sheet.getRange("A4").values = [["全量及共同键核对"]]; sheet.getRange("A4:K4").merge();
  write(sheet, 4, 0, [["项目", "OMS来源系统/清单", "OMS字段/汇总逻辑", "OMS金额/数量", "SAP来源系统/清单", "SAP字段/汇总逻辑", "SAP金额/数量", "差异（SAP-OMS）", "匹配率", "匹配键", "与总览勾稽"]]); header(sheet, "A5:K5");
  write(sheet, 5, 0, [
    ["专项范围金额", "OMS/OMS系统日结月结查询记录SQL", "share_amount；2026年1—6月月结Y001标准结算子集汇总", ctl.oms_month_amount || 0, "SAP/SAP开票清单", "tax_amount；2026年1—6月标准发票（2C）全量汇总", ctl.sap_full_amount || 0, null, null, "专项范围汇总", "总览金额匹配1.总量匹配"],
    ["专项范围数量", "OMS/OMS系统日结月结查询记录SQL", "item_num；2026年1—6月月结Y001标准结算子集汇总", ctl.oms_month_qty || 0, "SAP/SAP开票清单", "invoice_qty；2026年1—6月标准发票（2C）全量汇总", ctl.sap_full_qty || 0, null, null, "专项范围汇总", "总览数量匹配1.总量匹配"],
    ["共同键金额", "OMS/OMS系统日结月结查询记录SQL", "share_amount；共同键汇总Y001", exactMap.oms_amount || 0, "SAP/SAP开票清单", "tax_amount；标准发票（2C）共同键汇总", exactMap.sap_amount || 0, null, null, "OMS销售单号+物料号+销售单位", "总览金额匹配2.开票明细匹配"],
    ["共同键数量", "OMS/OMS系统日结月结查询记录SQL", "item_num；共同键汇总Y001", exactMap.oms_qty || 0, "SAP/SAP开票清单", "invoice_qty；标准发票（2C）共同键汇总", exactMap.sap_qty || 0, null, null, "OMS销售单号+物料号+销售单位", "总览数量匹配2.开票明细匹配"],
  ]);
  for (let r = 6; r <= 9; r++) { formula(sheet, `H${r}`, `=G${r}-D${r}`); formula(sheet, `I${r}`, `=IFERROR(MIN(ABS(D${r}),ABS(G${r}))/MAX(ABS(D${r}),ABS(G${r})),0)`); }
  sapRefs.totalAmountRow = 6; sapRefs.totalQtyRow = 7; sapRefs.commonAmountRow = 8; sapRefs.commonQtyRow = 9;
  body(sheet, "A6:K9"); sheet.getRange("D6:H6").setNumberFormat("#,##0.00;[Red](#,##0.00);-"); sheet.getRange("D7:H7").setNumberFormat("#,##0"); sheet.getRange("D8:H8").setNumberFormat("#,##0.00;[Red](#,##0.00);-"); sheet.getRange("D9:H9").setNumberFormat("#,##0"); sheet.getRange("I6:I9").setNumberFormat("0.00%"); sheet.getRange("A6:K9").format.wrapText = true; sheet.getRange("A6:K9").format.rowHeight = 46; sheet.getRange("A6:K9").format.fill = C.green;
  [18, 34, 38, 20, 28, 38, 20, 20, 16, 34, 24].forEach((w, i) => sheet.getRange(`${col(i)}:${col(i)}`).format.columnWidth = w); sheet.freezePanes.freezeRows(5);
}

// 全局总览：主链展示SAP—OMS月结—发货对账结果；订单—账单稽核紧接数量核对列示。
{
  const sheet = ws("1.全局口径与总览"); title(sheet, "销售ToC三单匹配结果总览", "审计核对方向：SAP开票 → OMS月结 → 发货对账结果；订单—账单及12月追溯在数量核对下方单独稽核。", "P");
  section(sheet, "A4:L4"); sheet.getRange("A4").values = [["金额匹配及逐层勾稽"]]; sheet.getRange("A4:L4").merge();
  write(sheet, 4, 0, [["核对步骤", "参与匹配数据", "匹配字段/维度", "SAP开票总金额", "SAP开票匹配金额", "OMS月结总金额", "OMS月结匹配金额", "发货对账总金额", "发货对账匹配金额", "差异（前项-后项）", "金额匹配率", "匹配覆盖率"]]); header(sheet, "A5:L5");
  write(sheet, 5, 0, [
    ["1.SAP开票—OMS月结", "SAP标准发票（2C）—OMS Y001月结", "OMS销售单号+物料编码+销售单位", null, null, null, null, "N/A", "N/A", null, null, null],
    ["2.OMS月结—发货对账结果（资料齐备期间）", "OMS Y001月结—发货对账成功记录", "月份+OMS客户/店铺编码+商品/货品编码", "N/A", "N/A", null, null, null, null, null, null, null],
  ]);
  formula(sheet, "D6", `='2.OMS月结-SAP汇总'!G${sapRefs.totalAmountRow}`); formula(sheet, "E6", `='2.OMS月结-SAP汇总'!G${sapRefs.commonAmountRow}`); formula(sheet, "F6", `='2.OMS月结-SAP汇总'!D${sapRefs.totalAmountRow}`); formula(sheet, "G6", `='2.OMS月结-SAP汇总'!D${sapRefs.commonAmountRow}`); formula(sheet, "J6", "=ROUND(E6-G6,2)"); formula(sheet, "K6", "=IFERROR(MIN(ABS(E6),ABS(G6))/MAX(ABS(E6),ABS(G6)),0)"); formula(sheet, "L6", "=IFERROR(ABS(E6)/ABS(D6),0)");
  formula(sheet, "F7", `='3.对账结果-OMS月结汇总'!D${settlementOmsRefs.totalRow}`); formula(sheet, "G7", `='3.对账结果-OMS月结汇总'!H${settlementOmsRefs.totalRow}`); formula(sheet, "H7", `='3.对账结果-OMS月结汇总'!C${settlementOmsRefs.totalRow}`); formula(sheet, "I7", `='3.对账结果-OMS月结汇总'!G${settlementOmsRefs.totalRow}`); formula(sheet, "J7", "=ROUND(G7-I7,2)"); formula(sheet, "K7", "=IFERROR(MIN(ABS(G7),ABS(I7))/MAX(ABS(G7),ABS(I7)),0)"); formula(sheet, "L7", "=IFERROR(ABS(G7)/ABS(F7),0)");
  body(sheet, "A6:L7"); sheet.getRange("A6:L7").format.borders = { preset: "all", style: "thin", color: C.line }; sheet.getRange("D6:J7").setNumberFormat("#,##0.00;[Red](#,##0.00);-"); sheet.getRange("K6:L7").setNumberFormat("0.00%"); sheet.getRange("A6:L7").format.wrapText = true; sheet.getRange("A6:L7").format.rowHeight = 54; sheet.getRange("A6:L7").format.fill = C.green;

  section(sheet, "A10:L10"); sheet.getRange("A10").values = [["数量匹配及逐层勾稽"]]; sheet.getRange("A10:L10").merge();
  write(sheet, 10, 0, [["核对步骤", "参与匹配数据", "匹配字段/维度", "SAP开票总数量", "SAP开票匹配数量", "OMS月结总数量", "OMS月结匹配数量", "发货对账总数量", "发货对账匹配数量", "差异（前项-后项）", "数量匹配率", "匹配覆盖率"]]); header(sheet, "A11:L11");
  write(sheet, 11, 0, [
    ["1.SAP开票—OMS月结", "SAP标准发票（2C）—OMS Y001月结", "OMS销售单号+物料编码+销售单位", null, null, null, null, "N/A", "N/A", null, null, null],
    ["2.OMS月结—发货对账结果（资料齐备期间）", "OMS Y001月结—发货对账成功记录", "月份+客户/店铺编码+商品/货品编码", "N/A", "N/A", null, null, null, null, null, null, null],
  ]);
  formula(sheet, "D12", `='2.OMS月结-SAP汇总'!G${sapRefs.totalQtyRow}`); formula(sheet, "E12", `='2.OMS月结-SAP汇总'!G${sapRefs.commonQtyRow}`); formula(sheet, "F12", `='2.OMS月结-SAP汇总'!D${sapRefs.totalQtyRow}`); formula(sheet, "G12", `='2.OMS月结-SAP汇总'!D${sapRefs.commonQtyRow}`); formula(sheet, "J12", "=ROUND(E12-G12,2)"); formula(sheet, "K12", "=IFERROR(MIN(ABS(E12),ABS(G12))/MAX(ABS(E12),ABS(G12)),0)"); formula(sheet, "L12", "=IFERROR(ABS(E12)/ABS(D12),0)");
  write(sheet, 12, 5, [[settlementTotal.oms_total_quantity ?? null, settlementTotal.oms_common_quantity ?? null, settlementTotal.reconciliation_total_quantity ?? null, settlementTotal.reconciliation_common_quantity ?? null]]); formula(sheet, "J13", "=ROUND(G13-I13,2)"); formula(sheet, "K13", "=IFERROR(MIN(ABS(G13),ABS(I13))/MAX(ABS(G13),ABS(I13)),0)"); formula(sheet, "L13", "=IFERROR(ABS(G13)/ABS(F13),0)");
  body(sheet, "A12:L13"); sheet.getRange("A12:L13").format.borders = { preset: "all", style: "thin", color: C.line }; sheet.getRange("D12:J13").setNumberFormat("#,##0.00"); sheet.getRange("K12:L13").setNumberFormat("0.00%"); sheet.getRange("A12:L13").format.wrapText = true; sheet.getRange("A12:L13").format.rowHeight = 54; sheet.getRange("A12:L13").format.fill = C.green;

  section(sheet, "A16:L16"); sheet.getRange("A16").values = [["订单-账单 稽核"]]; sheet.getRange("A16:L16").merge();
  section(sheet, "A18:L18"); sheet.getRange("A18").values = [["惠策账单—旺店通订单金额匹配及勾稽"]]; sheet.getRange("A18:L18").merge();
  write(sheet, 18, 0, [["核对步骤", "参与匹配数据", "匹配字段/维度", "惠策账单总金额", "惠策账单参与匹配金额", "惠策账单实际匹配金额", "旺店通订单总金额", "旺店通订单参与匹配金额", "旺店通订单实际匹配金额", "差异（惠策-旺店通）", "金额匹配率", "惠策匹配覆盖率"]]); header(sheet, "A19:L19");
  write(sheet, 19, 0, [["惠策账单—旺店通订单", "惠策账单清单—旺店通订单清单", "惠策平台订单号=旺店通原始单号；金额差异≤0.01元", null, participatingBillCash, matchedBillCash, fullWdtComparisonAmount, participatingWdtAmount, matchedWdtAmount, null, null, null]]);
  formula(sheet, "D20", "=C30"); formula(sheet, "J20", "=ROUND(F20-I20,2)"); formula(sheet, "K20", "=IFERROR(MIN(ABS(F20),ABS(I20))/MAX(ABS(F20),ABS(I20)),0)"); formula(sheet, "L20", "=IFERROR(ABS(F20)/ABS(D20),0)"); body(sheet, "A20:L20"); sheet.getRange("A20:L20").format.borders = { preset: "all", style: "thin", color: C.line }; sheet.getRange("D20:J20").setNumberFormat("#,##0.00;[Red](#,##0.00);-"); sheet.getRange("K20:L20").setNumberFormat("0.00%"); sheet.getRange("A20:L20").format.wrapText = true; sheet.getRange("A20:L20").format.rowHeight = 58; sheet.getRange("A20:L20").format.fill = C.green;

  section(sheet, "A23:H23"); sheet.getRange("A23").values = [["惠策账单实际实收分类金额及占比"]]; sheet.getRange("A23:H23").merge();
  write(sheet, 23, 0, [["核对分类", "订单组数", "惠策实际实收金额", "惠策实收占比", "来源系统/清单", "使用字段", "汇总逻辑", "勾稽位置"]]); header(sheet, "A24:H24");
  const orderAuditRows = orderCategories.map(label => { const x = findResult(orderResults, label); return [label, x.groups || 0, x.bill_cash || 0, null, "惠策系统/惠策账单清单", "收款金额；退款金额", "按平台订单号分类后汇总净实收", "分类合计=惠策明细全量"]; });
  write(sheet, 24, 0, orderAuditRows); for (let r = 25; r <= 29; r++) formula(sheet, `D${r}`, `=IFERROR(ABS(C${r})/ABS($C$30),0)`); write(sheet, 29, 0, [["合计", null, null, null, "惠策系统/惠策账单清单", "收款金额-退款金额", "全部分类金额合计", "直接勾稽本区域"]]); formula(sheet, "B30", "=SUM(B25:B29)"); formula(sheet, "C30", "=SUM(C25:C29)"); formula(sheet, "D30", "=SUM(D25:D29)");
  body(sheet, "A25:H30"); status(sheet, "A25:A29"); sheet.getRange("B25:B30").setNumberFormat("#,##0"); sheet.getRange("C25:C30").setNumberFormat("#,##0.00;[Red](#,##0.00);-"); sheet.getRange("D25:D30").setNumberFormat("0.00%"); sheet.getRange("A30:H30").format = { fill: C.green, font: { bold: true, color: C.greenText }, borders: { top: { style: "medium", color: C.blue } } };

  section(sheet, "A32:H32"); sheet.getRange("A32").values = [["订单—账单金额层级勾稽"]]; sheet.getRange("A32:H32").merge();
  write(sheet, 32, 0, [["金额层级", "惠策账单金额", "旺店通订单金额", "差异（惠策-旺店通）", "金额匹配率", "惠策覆盖率", "旺店通覆盖率", "口径"]]); header(sheet, "A33:H33");
  write(sheet, 33, 0, [
    ["执行总金额", null, fullWdtComparisonAmount, null, "N/A", "N/A", "N/A", "双方正式范围全量"],
    ["参与匹配金额", participatingBillCash, participatingWdtAmount, null, null, "N/A", "N/A", "平台订单号共同键，含金额差异"],
    ["实际匹配金额", matchedBillCash, matchedWdtAmount, null, null, null, null, "平台订单号共同键且金额差异不超过0.01元"],
  ]); formula(sheet, "B34", "=C30"); for (const r of [34,35,36]) formula(sheet, `D${r}`, `=B${r}-C${r}`); for (const r of [35,36]) formula(sheet, `E${r}`, `=IFERROR(MIN(ABS(B${r}),ABS(C${r}))/MAX(ABS(B${r}),ABS(C${r})),0)`); formula(sheet, "F36", "=IFERROR(ABS(B36)/ABS(B34),0)"); formula(sheet, "G36", "=IFERROR(ABS(C36)/ABS(C34),0)");
  body(sheet, "A34:H36"); sheet.getRange("A36:H36").format.fill = C.green; sheet.getRange("B34:D36").setNumberFormat("#,##0.00;[Red](#,##0.00);-"); sheet.getRange("E34:G36").setNumberFormat("0.00%");

  section(sheet, "A38:H38"); sheet.getRange("A38").values = [["订单—账单范围敏感性统计"]]; sheet.getRange("A38:H38").merge();
  write(sheet, 38, 0, [["旺店通订单范围", "惠策账单执行总金额", "惠策实际匹配金额", "旺店通订单执行总金额", "旺店通实际匹配金额", "差异（惠策-旺店通）", "金额匹配率", "惠策覆盖率"]]); header(sheet, "A39:H39");
  write(sheet, 39, 0, [
    ["2026年度正式范围", null, matchedBillCash, fullWdtComparisonAmount, matchedWdtAmount, null, null, null],
    ["Cut-off参考范围（2025.12-2026.06）", cutoffDecember.huice_cash ?? null, cutoffDecember.amount_exact_cash ?? null, cutoffDecember.order_amount ?? null, cutoffDecember.amount_exact_wdt_amount ?? null, null, cutoffDecember.amount_match_rate ?? null, cutoffDecember.huice_coverage_rate ?? null],
  ]); formula(sheet, "B40", "=C30"); formula(sheet, "F40", "=C40-E40"); formula(sheet, "G40", "=IFERROR(MIN(ABS(C40),ABS(E40))/MAX(ABS(C40),ABS(E40)),0)"); formula(sheet, "H40", "=IFERROR(ABS(C40)/ABS(B40),0)"); formula(sheet, "F41", "=C41-E41");
  body(sheet, "A40:H41"); sheet.getRange("A40:H40").format.fill = C.green; sheet.getRange("A41:H41").format.fill = C.amber; sheet.getRange("B40:F41").setNumberFormat("#,##0.00;[Red](#,##0.00);-"); sheet.getRange("G40:H41").setNumberFormat("0.00%");

  section(sheet, "A44:B44"); sheet.getRange("A44").values = [["Lead｜执行总金额"]]; sheet.getRange("A44:B44").merge(); write(sheet, 44, 0, [["清单+字段", "执行总金额"]]); header(sheet, "A45:B45");
  write(sheet, 45, 0, [["SAP开票清单｜含税金额（标准发票（2C））", null], ["OMS月结SQL｜share_amount（Y001）", null], ["发货对账结果｜收款金额（对账成功）", null]]); formula(sheet, "B46", "=D6"); formula(sheet, "B47", "=F6"); formula(sheet, "B48", "=H7"); body(sheet, "A46:B48"); sheet.getRange("A46:B48").format.borders = { preset: "all", style: "thin", color: C.line }; sheet.getRange("B46:B48").setNumberFormat("#,##0.00;[Red](#,##0.00);-");

  section(sheet, "A51:P51"); sheet.getRange("A51").values = [["Notes:"]]; sheet.getRange("A51:P51").merge();
  const sourceNotes = [
    "1. SAP开票使用SAP开票清单“含税金额、开票数量”，过滤发票类型=标准发票（2C）；OMS月结使用Y001标准结算子集“share_amount、item_num”。",
    "2. 发货对账结果使用“收款金额、实际数量”，仅纳入对账状态=对账成功；与OMS按月份+店铺/客户编码+货品/商品编码核对。客户确认共同键差异由跨期结算导致，原始金额保留，跨期差异通过桥接调整列示。",
    `3. 发货对账—OMS月结当前资料范围：${settlementRangeNote}。所有对账记录统一按原始字段“账期结束日期”归属结算月；6月末D-1/D-2记录即使在7月初完成对账仍纳入，账期结束在2025年12月或2026年7月的记录作为审计期外数据不纳入。`,
    "4. 跨期调整金额=OMS共同键金额-发货对账原共同键金额（仅限具备跨期证据的共同键）；调整后差异=前项匹配金额-后项调整后匹配金额；金额/数量匹配率=双方较小值÷较大值。",
    "5. 惠策账单—旺店通订单及2025年12月Cut-off追溯不纳入SAP—OMS—发货对账主链，统一放在本页“订单-账单 稽核”区域。",
  ];
  sourceNotes.forEach((note, i) => { const r = 52 + i; sheet.getRange(`A${r}:P${r}`).merge(); sheet.getRange(`A${r}`).values = [[note]]; }); body(sheet, "A52:P56"); sheet.getRange("A52:P56").format.wrapText = true; sheet.getRange("A52:P56").format.rowHeight = 31;

  const flowGroups = [["A", "B"], ["C", "E"], ["F", "G"], ["H", "I"], ["J", "K"], ["L", "M"], ["N", "N"], ["O", "P"]];
  const setFlowRow = (row, values) => { flowGroups.forEach(([start, end], index) => { const range = `${start}${row}:${end}${row}`; if (start !== end) sheet.getRange(range).merge(); sheet.getRange(`${start}${row}`).values = [[values[index]]]; }); };
  section(sheet, "A59:P59"); sheet.getRange("A59").values = [["业务数据流及核对结果"]]; sheet.getRange("A59:P59").merge();
  setFlowRow(60, ["业务步骤", "数据来源及字段", "匹配字段/维度", "验证范围", "金额匹配率", "数量匹配率", "覆盖率", "核对结论"]); header(sheet, "A60:P60"); sheet.getRange("A60:P60").format.rowHeight = 34;
  setFlowRow(61, ["1.SAP开票 → OMS月结Y001", "SAP：含税金额/开票数量；OMS：share_amount/item_num", "OMS销售单号+物料编码+销售单位", "2026年1—6月", null, null, null, "逐条共同键核对；详见2.OMS月结-SAP汇总"]);
  setFlowRow(62, ["2.OMS月结Y001 → 发货对账结果", "OMS：share_amount/item_num；发货对账：收款金额/实际数量", "月份+客户/店铺编码+商品/货品编码", `资料齐备范围：${availableMonthsLabel}`, null, null, null, hasPendingSettlementMonths ? `${pendingMonthsLabel}待获取发货对账明细` : "1—6月均已完成核对"]);
  setFlowRow(63, ["补充稽核：惠策账单 → 旺店通订单", "惠策：收款金额-退款金额；旺店通：allocated_total/订单头金额", "平台订单号；金额一致范围", "2026正式范围及2025年12月Cut-off", "见本页上方", "见本页上方", "见本页上方", "不纳入SAP—OMS—发货对账主链"]);
  formula(sheet, "J61", "=K6"); formula(sheet, "L61", "=K12"); formula(sheet, "N61", "=L6"); formula(sheet, "J62", "=K7"); formula(sheet, "L62", "=K13"); formula(sheet, "N62", "=L7");
  body(sheet, "A61:P63"); sheet.getRange("A61:P63").format.borders = { preset: "all", style: "thin", color: C.line }; sheet.getRange("A61:P63").format.wrapText = true; sheet.getRange("A61:P63").format.rowHeight = 62; for (const c of ["J", "L", "N"]) sheet.getRange(`${c}61:${c}62`).setNumberFormat("0.00%"); sheet.getRange("A61:P62").format.fill = C.green; sheet.getRange("A63:P63").format.fill = C.amber;

  section(sheet, "A66:P66"); sheet.getRange("A66").values = [["业务流信息（供财务人员理解）"]]; sheet.getRange("A66:P66").merge();
  const businessNotes = [
    "1. SAP标准发票（2C）作为财务确认端，OMS月结Y001作为结算执行端，发货对账结果作为平台结算中心核销成功明细的穿透支持。",
    "2. 发货对账结果与OMS月结按月、客户/店铺编码及物料编码汇总核对，可同时验证收款金额和实际数量；该链路替代原惠策账单—OMS月结在主底稿中的位置。",
    "3. 惠策账单与旺店通订单仍用于订单来源追溯和Cut-off敏感性分析，已放在本页数量核对下方的“订单-账单 稽核”，但不与主链金额强制串联。",
    `4. ${hasPendingSettlementMonths ? `待取得${pendingMonthsLabel}发货对账明细后` : "后续如取得补充对账明细"}，只需放入约定目录并重新运行一键构建程序，月份行、资料范围、汇总金额及明细将自动刷新。`,
  ];
  businessNotes.forEach((note, index) => { const row = 67 + index; sheet.getRange(`A${row}:P${row}`).merge(); sheet.getRange(`A${row}`).values = [[note]]; }); body(sheet, "A67:P70"); sheet.getRange("A67:P70").format.wrapText = true; sheet.getRange("A67:P70").format.rowHeight = 35;
  [34, 38, 40, 20, 20, 20, 20, 20, 20, 22, 17, 17, 18, 18, 18, 28].forEach((w, i) => sheet.getRange(`${col(i)}:${col(i)}`).format.columnWidth = w); sheet.freezePanes.freezeRows(5);
}

// 正式范围旺店通订单明细：总览保留控制数、规则及全量分片索引；完整记录存放于同目录CSV分片。
{
  const sheet = ws("7.旺店通订单匹配明细");
  title(sheet, "2026年度正式范围旺店通订单匹配明细", "惠策账单明细表 ↔ 旺店通订单清单；仅含2026年度正式范围，2025年12月Cut-off扩展数据未纳入本明细。", "H");
  const categoryMap = Object.fromEntries((formalDetailManifest.categories || []).map(item => [item.category, item]));
  const matched = categoryMap["可匹配条目"] || {};
  const billOnly = categoryMap["仅账单未匹配"] || {};
  const orderOnly = categoryMap["仅订单未匹配"] || {};
  const formalHuiceTotal = (matched.huice_amount || 0) + (billOnly.huice_amount || 0);

  section(sheet, "A4:H4"); sheet.getRange("A4").values = [["汇总统计及总览勾稽"]]; sheet.getRange("A4:H4").merge();
  write(sheet, 4, 0, [["匹配状态", "平台订单数量", "惠策实际实收", "旺店通订单金额"]]); header(sheet, "A5:D5");
  write(sheet, 5, 0, [
    ["总惠策账单明细", ctl.huice_orders || 0, formalHuiceTotal, "N/A"],
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
    `完整明细共${(formalDetailManifest.totals?.union_rows || 0).toLocaleString("zh-CN")}条。受Excel单Sheet最大1,048,576行限制，按每个CSV不超过1,000,000条拆分；CSV可由Excel直接打开并筛选。`,
    "惠策账单清单未提供商品编码和商品名称；商品编码取自旺店通商品明细，商品名称由OMS物料名称按商品编码辅助映射。惠策单边记录相应字段显示N/A。",
    "全量文件位于主底稿同目录“2026年度正式范围旺店通订单匹配明细/明细数据”；文件级行数、金额及SHA-256校验值详见“明细拆分索引及校验结果.json”。",
  ];
  detailNotes.forEach((note, index) => { const row = noteRow + 1 + index; sheet.getRange(`A${row}:H${row}`).merge(); sheet.getRange(`A${row}`).values = [[note]]; }); body(sheet, `A${noteRow + 1}:H${noteRow + 3}`); sheet.getRange(`A${noteRow + 1}:H${noteRow + 3}`).format.wrapText = true; sheet.getRange(`A${noteRow + 1}:H${noteRow + 3}`).format.rowHeight = 32;

  [30, 12, 16, 16, 16, 20, 20, 48].forEach((width, index) => sheet.getRange(`${col(index)}:${col(index)}`).format.columnWidth = width);
  sheet.freezePanes.freezeRows(5);
}

// 精简后的支持页：只保留与最终口径直接相关的字段。
detail("9.惠策内部核对明细", "惠策明细—惠策店铺汇总实际实收核对明细", "按结算月份+平台+店铺逐项核对实际实收金额。", data.internal, { huice_shop: 34, result: 20 }, ["detail_success_amount", "summary_success_amount", "success_difference", "detail_receivable", "summary_receivable", "receivable_difference", "historical_rows", "historical_receivable", "historical_cash"]);
detail("4.对账结果-OMS月结明细", "发货对账结果—OMS月结Y001核对明细", `原始金额、跨期调整及调整后金额并列展示；客户确认跨期结算差异已作桥接调整；${settlementRangeNote}。`, settlementDetail, { wdt_shop_name: 34, oms_customer_name: 40, material_code: 18, result: 24, source_files: 44, settlement_periods: 32, business_months: 22, cross_period_flag: 18, adjustment_basis: 46, cross_period_adjustment_amount: 20, adjusted_reconciliation_amount: 20, adjusted_amount_difference: 20, adjusted_amount_match_rate: 18 }, []);
{
  const sheet = ws("6.数量核对明细"); title(sheet, "销售数量逐层核对明细", "Step 1为OMS月结—发货对账结果；Step 2为OMS—SAP共同键核对。", "L");
  const select = (dataset, fields) => { const indexes = fields.map(h => dataset.headers.indexOf(h)); return dataset.rows.map(row => indexes.map(i => row[i])); };
  const step1Fields = ["month", "shop_customer_code", "wdt_shop_name", "oms_customer_name", "material_code", "reconciliation_quantity", "oms_quantity", "quantity_difference", "quantity_match_rate", "result"];
  const step1Rows = select(settlementDetail, step1Fields); section(sheet, "A4:J4"); sheet.getRange("A4").values = [["Step 1｜OMS月结数量 vs 发货对账实际数量"]]; sheet.getRange("A4:J4").merge(); write(sheet, 4, 0, [step1Fields.map(h => headerZh[h] ? `${h}\n${headerZh[h]}` : h)]); header(sheet, "A5:J5"); sheet.getRange("A5:J5").format.rowHeight = 36;
  for (let i = 0; i < step1Rows.length; i += 3000) write(sheet, 5 + i, 0, step1Rows.slice(i, i + 3000)); const step1End = 5 + step1Rows.length; body(sheet, `A6:J${step1End}`); status(sheet, `J6:J${step1End}`); sheet.getRange(`B6:B${step1End}`).setNumberFormat("@"); sheet.getRange(`E6:E${step1End}`).setNumberFormat("@"); sheet.getRange(`F6:H${step1End}`).setNumberFormat("#,##0.00"); sheet.getRange(`I6:I${step1End}`).setNumberFormat("0.00%");
  const step2Section = step1End + 2; const step2Header = step2Section + 1; const step2Start = step2Section + 2;
  const step2Fields = ["outbound_month", "file_month", "oms_sales_no", "material_code", "sales_unit", "customer_code", "customer_name", "oms_qty", "sap_qty", "quantity_difference", "mapping_result", "source_result"];
  const step2Rows = select(data.omsSap, step2Fields); section(sheet, `A${step2Section}:L${step2Section}`); sheet.getRange(`A${step2Section}`).values = [["Step 2｜OMS月结数量 vs SAP标准发票（2C）数量"]]; sheet.getRange(`A${step2Section}:L${step2Section}`).merge(); write(sheet, step2Header - 1, 0, [step2Fields.map(h => headerZh[h] ? `${h}\n${headerZh[h]}` : h)]); header(sheet, `A${step2Header}:L${step2Header}`); sheet.getRange(`A${step2Header}:L${step2Header}`).format.rowHeight = 36;
  for (let i = 0; i < step2Rows.length; i += 3000) write(sheet, step2Start - 1 + i, 0, step2Rows.slice(i, i + 3000)); const step2End = step2Start + step2Rows.length - 1; body(sheet, `A${step2Start}:L${step2End}`); status(sheet, `K${step2Start}:L${step2End}`); for (const c of ["C","D","F"]) sheet.getRange(`${c}${step2Start}:${c}${step2End}`).setNumberFormat("@"); sheet.getRange(`H${step2Start}:J${step2End}`).setNumberFormat("#,##0");
  [14, 20, 30, 38, 18, 20, 20, 20, 17, 24, 24, 20].forEach((w, i) => sheet.getRange(`${col(i)}:${col(i)}`).format.columnWidth = w); sheet.freezePanes.freezeRows(5); sheet.freezePanes.freezeColumns(2);
}
detail("10.店铺客户映射", "惠策店铺—OMS客户映射", "仅展示店铺与OMS客户编码映射结果，金额口径以实际结算为准。", data.shop, { huice_shop: 34, customer_name: 40 }, ["mapping_status", "mapping_source", "bill_record_count", "bill_receivable", "source_rows", "success_count", "bill_success_amount"]);

await fs.mkdir(outputDir, { recursive: true });
const previews = path.join(root, "reconciliation/qa_previews"); await fs.mkdir(previews, { recursive: true });
console.log("OVERVIEW\n" + (await wb.inspect({ kind: "table", range: "1.全局口径与总览!A1:P82", include: "values,formulas", tableMaxRows: 84, tableMaxCols: 17, maxChars: 82000 })).ndjson);
console.log("HUICE_INTERNAL\n" + (await wb.inspect({ kind: "table", range: "8.惠策内部核对汇总!A1:H16", include: "values,formulas", tableMaxRows: 18, tableMaxCols: 9, maxChars: 12000 })).ndjson);
console.log("SETTLEMENT_OMS\n" + (await wb.inspect({ kind: "table", range: "3.对账结果-OMS月结汇总!A1:O18", include: "values,formulas", tableMaxRows: 20, tableMaxCols: 16, maxChars: 28000 })).ndjson);
console.log("FORMAL_ORDER_DETAIL\n" + (await wb.inspect({ kind: "table", range: "7.旺店通订单匹配明细!A1:H34", include: "values,formulas", tableMaxRows: 36, tableMaxCols: 9, maxChars: 18000 })).ndjson);
console.log("ERRORS\n" + (await wb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 300 }, summary: "formula errors", maxChars: 5000 })).ndjson);
for (const name of names) { const sheet = ws(name), used = sheet.getUsedRange(true), maxCols = Math.min(used.columnCount || 8, name === "1.全局口径与总览" ? 16 : name === "3.对账结果-OMS月结汇总" ? 15 : name.includes("明细") || name.includes("映射") ? 10 : 12), maxRows = name === "1.全局口径与总览" ? 82 : name.includes("明细") ? 20 : 30; const blob = await wb.render({ sheetName: name, range: `A1:${col(maxCols - 1)}${maxRows}`, scale: 1.15, format: "png" }); await fs.writeFile(path.join(previews, `${name}.png`), new Uint8Array(await blob.arrayBuffer())); }
const adjustmentPreview = await wb.render({ sheetName: "4.对账结果-OMS月结明细", range: "I1:Y20", scale: 1.15, format: "png" }); await fs.writeFile(path.join(previews, "4.对账结果-OMS月结明细-跨期调整.png"), new Uint8Array(await adjustmentPreview.arrayBuffer()));
const qtyStep2Section = 7 + settlementDetail.rows.length; const qtyStep2Preview = await wb.render({ sheetName: "6.数量核对明细", range: `A${qtyStep2Section}:L${qtyStep2Section + 17}`, scale: 1.15, format: "png" }); await fs.writeFile(path.join(previews, "6.数量核对明细-Step2.png"), new Uint8Array(await qtyStep2Preview.arrayBuffer()));
const out = await SpreadsheetFile.exportXlsx(wb); await out.save(outputFile); console.log(JSON.stringify({ outputFile, sheets: names, previews }, null, 2));
