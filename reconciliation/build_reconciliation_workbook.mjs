import fs from "node:fs/promises";
import path from "node:path";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

const root = path.resolve(process.argv[2] || "/Users/aatrox/Desktop/miaoke sales to c/reconciliation/output_full");
const outputDir = path.resolve(process.argv[3] || "/Users/aatrox/Desktop/miaoke sales to c/outputs/sales_toc_reconciliation_20260726");
const outputFile = path.join(outputDir, "销售ToC全链路核对底稿_2025年12月至2026年6月.xlsx");

const C = {
  navy: "#17365D", blue: "#2F75B5", mid: "#5B9BD5", pale: "#DDEBF7",
  lighter: "#EAF3F8", gray: "#F2F2F2", line: "#B4C6E7", text: "#203040",
  green: "#E2F0D9", greenText: "#375623", amber: "#FFF2CC", amberText: "#7F6000",
  red: "#FCE4D6", redText: "#9C0006", white: "#FFFFFF",
};

const summary = JSON.parse(await fs.readFile(path.join(root, "summary.json"), "utf8"));
const readDetail = async (name) => JSON.parse(await fs.readFile(path.join(root, `${name}_workbook.json`), "utf8"));
const details = {
  wdt: await readDetail("wdt_huice_detail"),
  huice: await readDetail("huice_summary_detail"),
  sap: await readDetail("oms_sap_detail"),
  omsHuice: await readDetail("oms_huice_detail"),
  mapping: await readDetail("customer_shop_mapping"),
};
const wdtEmbedded = { ...details.wdt, rows: details.wdt.rows.slice(0, 20000) };

const workbook = Workbook.create();
const names = [
  "1.总汇总", "2.参数与说明", "3.旺店通-惠策汇总", "4.旺店通-惠策明细",
  "5.惠策明细-汇总核对", "6.惠策汇总核对明细", "7.OMS-SAP汇总", "8.OMS-SAP明细",
  "9.OMS-惠策汇总", "10.OMS-惠策明细", "11.店铺映射", "12.完整明细索引",
];
for (const name of names) workbook.worksheets.add(name);

function sheet(name) { return workbook.worksheets.getItem(name); }
function colName(index) {
  let n = index + 1, s = "";
  while (n) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26); }
  return s;
}
function clean(v) {
  if (v === undefined || v === null || (typeof v === "number" && !Number.isFinite(v))) return null;
  return v;
}
function styleTitle(ws, title, subtitle, lastCol = "J") {
  ws.showGridLines = false;
  ws.getRange(`A1:${lastCol}1`).merge();
  ws.getRange("A1").values = [[title]];
  ws.getRange(`A1:${lastCol}1`).format = {
    fill: C.navy, font: { bold: true, color: C.white, size: 16 },
    verticalAlignment: "center", horizontalAlignment: "left",
  };
  ws.getRange(`A1:${lastCol}1`).format.rowHeight = 30;
  ws.getRange(`A2:${lastCol}2`).merge();
  ws.getRange("A2").values = [[subtitle]];
  ws.getRange(`A2:${lastCol}2`).format = {
    fill: C.lighter, font: { color: C.text, italic: true, size: 10 },
    wrapText: true, verticalAlignment: "center",
  };
  ws.getRange(`A2:${lastCol}2`).format.rowHeight = 34;
}
function styleHeader(ws, range) {
  ws.getRange(range).format = {
    fill: C.blue, font: { bold: true, color: C.white },
    horizontalAlignment: "center", verticalAlignment: "center", wrapText: true,
    borders: { preset: "all", style: "thin", color: C.line },
  };
}
function styleSection(ws, range) {
  ws.getRange(range).format = {
    fill: C.pale, font: { bold: true, color: C.navy, size: 11 },
    borders: { bottom: { style: "medium", color: C.blue } },
  };
}
function styleBody(ws, range) {
  ws.getRange(range).format = {
    font: { color: C.text, size: 10 }, verticalAlignment: "center",
    borders: { insideHorizontal: { style: "thin", color: "#E7E6E6" } },
  };
}
function statusFormatting(ws, range) {
  const r = ws.getRange(range);
  r.conditionalFormats.add("containsText", { text: "一致", format: { fill: C.green, font: { color: C.greenText, bold: true } } });
  r.conditionalFormats.add("containsText", { text: "差异", format: { fill: C.red, font: { color: C.redText, bold: true } } });
  r.conditionalFormats.add("containsText", { text: "未映射", format: { fill: C.amber, font: { color: C.amberText, bold: true } } });
  r.conditionalFormats.add("containsText", { text: "仅", format: { fill: C.amber, font: { color: C.amberText, bold: true } } });
}
function writeBlock(ws, startRow, startCol, matrix) {
  if (!matrix.length || !matrix[0].length) return;
  ws.getRangeByIndexes(startRow, startCol, matrix.length, matrix[0].length).values = matrix.map(row => row.map(clean));
}
function formatNumericColumns(ws, headers, firstRow, lastRow) {
  headers.forEach((h, i) => {
    const letter = colName(i);
    const label = String(h).toLowerCase();
    const r = ws.getRange(`${letter}${firstRow}:${letter}${lastRow}`);
    if (/date/.test(label)) r.setNumberFormat("yyyy-mm-dd");
    else if (/amount|receivable|cash|tax|score|差异|金额/.test(label)) r.setNumberFormat("#,##0.00");
    else if (/qty|quantity|count|rows|数量|行数/.test(label)) r.setNumberFormat("#,##0");
  });
}
function detailSheet(name, title, note, data, widths = {}) {
  const ws = sheet(name);
  const lastCol = colName(data.headers.length - 1);
  styleTitle(ws, title, note, lastCol);
  writeBlock(ws, 3, 0, [data.headers]);
  styleHeader(ws, `A4:${lastCol}4`);
  const chunkSize = 5000;
  for (let i = 0; i < data.rows.length; i += chunkSize) {
    writeBlock(ws, 4 + i, 0, data.rows.slice(i, i + chunkSize));
  }
  const lastRow = 4 + data.rows.length;
  if (data.rows.length) {
    styleBody(ws, `A5:${lastCol}${lastRow}`);
    formatNumericColumns(ws, data.headers, 5, lastRow);
    const resultCol = data.headers.findIndex(h => h === "result" || h === "mapping_status");
    if (resultCol >= 0) statusFormatting(ws, `${colName(resultCol)}5:${colName(resultCol)}${lastRow}`);
  }
  data.headers.forEach((h, i) => {
    const letter = colName(i);
    let width = widths[h] || (/order|invoice|reconcile|customer_name|shop/.test(h) ? 23 : /amount|receivable|difference/.test(h) ? 16 : 13);
    ws.getRange(`${letter}:${letter}`).format.columnWidth = width;
  });
  ws.freezePanes.freezeRows(4);
  ws.freezePanes.freezeColumns(Math.min(2, data.headers.length));
  return ws;
}

// 1. 总汇总
{
  const ws = sheet("1.总汇总");
  styleTitle(ws, "销售 ToC 全链路核对总汇总", "核对期间：2025年12月至2026年6月；SAP限定“标准发票（2C）”。蓝色结构参考上年汇总底稿，结论以本期可取得资料为准。", "K");
  ws.getRange("A4:K4").merge(); ws.getRange("A4").values = [["金额匹配"]]; styleSection(ws, "A4:K4");
  const headers = ["序号","核对项目","参与数据","匹配/汇总口径","左方合计","右方合计","差异","覆盖率","结论","明细页面","备注"];
  writeBlock(ws, 4, 0, [headers]); styleHeader(ws, "A5:K5");
  const rows = [
    [1,"订单金额核对","旺店通－惠策明细","原始单号=平台订单号；匹配单号金额",null,null,null,null,"","4.旺店通-惠策明细","惠策净应收受平台补贴、退款及结算口径影响"],
    [2,"开票金额核对","OMS－SAP标准发票（2C）","OMS销售单号+物料编码+销售单位",null,null,null,null,"","8.OMS-SAP明细","一致键合计仅有1.38元舍入差"],
    [3,"账单金额核对","OMS未开票池－惠策","月度+客户店铺映射",null,null,null,null,"","10.OMS-惠策明细","店铺映射不足，当前仅用于识别差异方向"],
    [4,"惠策内部核对","惠策明细－惠策店铺汇总","月度+平台+店铺",null,null,null,null,"","6.惠策汇总核对明细","两表粒度不同，不能直接判定源数据错误"],
  ];
  writeBlock(ws, 5, 0, rows);
  ws.getRange("E6").formulas = [["='3.旺店通-惠策汇总'!F6"]];
  ws.getRange("F6").formulas = [["='3.旺店通-惠策汇总'!F7"]];
  ws.getRange("G6").formulas = [["=E6-F6"]]; ws.getRange("H6").formulas = [["='3.旺店通-惠策汇总'!B8"]];
  ws.getRange("I6").formulas = [["='3.旺店通-惠策汇总'!H6"]];
  ws.getRange("E7").formulas = [["='7.OMS-SAP汇总'!B6"]]; ws.getRange("F7").formulas = [["='7.OMS-SAP汇总'!C6"]];
  ws.getRange("G7").formulas = [["=F7-E7"]]; ws.getRange("H7").formulas = [["='7.OMS-SAP汇总'!B14"]]; ws.getRange("I7").formulas = [["='7.OMS-SAP汇总'!H6"]];
  ws.getRange("E8").formulas = [["='9.OMS-惠策汇总'!B5"]]; ws.getRange("F8").formulas = [["='9.OMS-惠策汇总'!B6"]];
  ws.getRange("G8").formulas = [["=E8-F8"]]; ws.getRange("H8").formulas = [["='9.OMS-惠策汇总'!F9"]]; ws.getRange("I8").formulas = [["='9.OMS-惠策汇总'!H5"]];
  ws.getRange("E9").formulas = [["='5.惠策明细-汇总核对'!E6"]]; ws.getRange("F9").formulas = [["='5.惠策明细-汇总核对'!F6"]];
  ws.getRange("G9").formulas = [["=E9-F9"]]; ws.getRange("H9").formulas = [["='5.惠策明细-汇总核对'!F9"]]; ws.getRange("I9").formulas = [["='5.惠策明细-汇总核对'!H6"]];
  styleBody(ws, "A6:K9");

  ws.getRange("A11:K11").merge(); ws.getRange("A11").values = [["数量与覆盖率匹配"]]; styleSection(ws, "A11:K11");
  writeBlock(ws, 11, 0, [headers]); styleHeader(ws, "A12:K12");
  writeBlock(ws, 12, 0, [
    [1,"订单覆盖率","旺店通－惠策明细","平台订单号",null,null,null,null,"","4.旺店通-惠策明细","以旺店通可匹配订单为分母"],
    [2,"发货数量核对","SAP－OMS","OMS销售单号+物料+单位",null,null,null,null,"","8.OMS-SAP明细","SAP有80个组合未在OMS SQL中找到"],
    [3,"店铺映射覆盖率","OMS客户－惠策店铺","客户名称相似度自动映射",null,null,null,null,"","11.店铺映射","高置信映射16/61，须保守解读"],
  ]);
  ws.getRange("E13").formulas = [["='3.旺店通-惠策汇总'!B5"]]; ws.getRange("F13").formulas = [["='3.旺店通-惠策汇总'!B7"]]; ws.getRange("G13").formulas = [["=E13-F13"]]; ws.getRange("H13").formulas = [["=F13/E13"]]; ws.getRange("I13").formulas = [["='3.旺店通-惠策汇总'!H5"]];
  ws.getRange("E14").formulas = [["='7.OMS-SAP汇总'!F5"]]; ws.getRange("F14").formulas = [["='7.OMS-SAP汇总'!G5"]]; ws.getRange("G14").formulas = [["=F14-E14"]]; ws.getRange("H14").formulas = [["=F14/E14"]]; ws.getRange("I14").formulas = [["='7.OMS-SAP汇总'!H5"]];
  ws.getRange("E15").formulas = [["='9.OMS-惠策汇总'!F7"]]; ws.getRange("F15").formulas = [["='9.OMS-惠策汇总'!F8"]]; ws.getRange("G15").formulas = [["=E15-F15"]]; ws.getRange("H15").formulas = [["=F15/E15"]]; ws.getRange("I15").formulas = [["='9.OMS-惠策汇总'!H6"]];
  styleBody(ws, "A13:K15"); statusFormatting(ws, "I6:I9"); statusFormatting(ws, "I13:I15");
  ws.getRange("E6:G9").setNumberFormat("#,##0.00"); ws.getRange("E13:G15").setNumberFormat("#,##0"); ws.getRange("H6:H9").setNumberFormat("0.00%"); ws.getRange("H13:H15").setNumberFormat("0.00%");
  const widths = [7,18,22,30,16,16,16,12,18,22,35]; widths.forEach((w,i)=>ws.getRange(`${colName(i)}:${colName(i)}`).format.columnWidth=w);
  ws.getRange("D6:D15").format.wrapText = true; ws.getRange("K6:K15").format.wrapText = true;
  ws.freezePanes.freezeRows(5);
}

// 2. 参数与说明
{
  const ws = sheet("2.参数与说明");
  styleTitle(ws, "核对参数、数据结构与审阅口径", "本页是审计追溯入口：列明数据流、主键、容差、限制及完整明细位置。", "G");
  writeBlock(ws, 3, 0, [["项目","左侧数据","右侧数据","主钩稽字段","辅助维度/容差","预期关系","本期限制"]]); styleHeader(ws, "A4:G4");
  writeBlock(ws, 4, 0, [
    ["旺店通－惠策","旺店通订单明细","惠策对账明细","原始单号=平台订单号","金额容差0.01元","订单号一对一或多行汇总后一致","惠策净应收包含平台结算调整；旺店通存在空原始单号"],
    ["惠策明细－汇总","惠策对账明细","惠策店铺汇总","月份+平台+店铺","金额容差0.01元","同口径汇总一致","汇总计数明显为更细粒度，现有明细无法还原全部计数"],
    ["OMS－SAP","OMS日结/月结SQL","SAP发票清单","document_no=OMS销售单号","物料编码+销售单位；合计容差2元","数量及含税金额一致","SAP筛选标准发票（2C）；80个SAP组合仅SAP"],
    ["OMS－惠策","OMS未匹配SAP池","惠策对账明细","月度+客户/店铺映射","金额容差0.01元","发货金额与账单净应收具备解释性关系","客户名到店铺名映射覆盖率低，不能作为最终结论"],
  ]); styleBody(ws, "A5:G8");
  ws.getRange("A10:G10").merge(); ws.getRange("A10").values = [["关键筛选与技术参数"]]; styleSection(ws, "A10:G10");
  writeBlock(ws, 10, 0, [["参数","值","说明"],
    ["核对期间",summary.parameters.date_scope,"OMS包含2025-12；SAP与惠策主体为2026-01至2026-06"],
    ["SAP发票筛选",summary.parameters.sap_filter,"排除其他发票类型"],
    ["旺店通-惠策金额容差",summary.parameters.wdt_huice_amount_tolerance,"绝对差额不超过该值视为一致"],
    ["OMS-SAP行金额容差",summary.parameters.oms_sap_line_amount_tolerance,"逐键判断"],
    ["OMS-SAP合计金额容差",summary.parameters.oms_sap_total_amount_tolerance,"考虑税额/舍入尾差"],
    ["生成时间",summary.generated_at,"由可复用核对程序生成"],
  ]); styleHeader(ws, "A11:C11"); styleBody(ws, "A12:C17");
  [22,34,34,32,28,28,40].forEach((w,i)=>ws.getRange(`${colName(i)}:${colName(i)}`).format.columnWidth=w);
  ws.getRange("A4:G17").format.wrapText = true; ws.freezePanes.freezeRows(4);
}

// 3. 旺店通-惠策汇总
{
  const ws = sheet("3.旺店通-惠策汇总");
  styleTitle(ws, "旺店通订单与惠策明细核对汇总", "主键：旺店通.原始单号 = 惠策明细.平台订单号；金额比较为旺店通分摊金额与惠策净应收。", "H");
  writeBlock(ws, 3, 0, [["订单控制","数值","说明","","金额控制","数值","说明","结论"]]); styleHeader(ws, "A4:H4");
  writeBlock(ws, 4, 0, [
    ["旺店通平台订单数",summary.controls.wdt_platform_orders,"含空原始单号记录","","旺店通总金额",summary.controls.wdt_amount,"订单明细口径",""],
    ["惠策平台订单数",summary.controls.huice_platform_orders,"平台订单号口径","","匹配单号旺店通金额",summary.controls.wdt_huice_matched_wdt_amount,"仅匹配主键",""],
    ["匹配订单数",summary.controls.wdt_huice_matched_orders,"两侧同时存在","","匹配单号惠策净应收",summary.controls.wdt_huice_matched_huice_amount,"仅匹配主键",""],
    ["匹配覆盖率",null,"匹配订单数/旺店通平台订单数","","匹配金额差异",null,"旺店通-惠策",""],
  ]);
  ws.getRange("B8").formulas = [["=B7/B5"]]; ws.getRange("F8").formulas = [["=F6-F7"]];
  ws.getRange("H5").values = [["需解释"]]; ws.getRange("H6").values = [["需解释"]]; ws.getRange("H7").values = [["需解释"]]; ws.getRange("H8").values = [["需解释"]];
  styleBody(ws, "A5:H8"); statusFormatting(ws, "H5:H8");
  ws.getRange("A10:H10").merge(); ws.getRange("A10").values = [["逐单号核对结果分布"]]; styleSection(ws, "A10:H10");
  writeBlock(ws, 10, 0, [["结果","订单数","旺店通金额","惠策净应收","差异","订单占比","判断","明细页"]]); styleHeader(ws, "A11:H11");
  const data = summary.wdt_huice_summary.map(x=>[x.result,x.order_count,clean(x.wdt_amount),clean(x.huice_net_receivable),clean(x.amount_difference),null,x.result.includes("一致")?"一致":"异常/待解释","4.旺店通-惠策明细"]);
  writeBlock(ws, 11, 0, data);
  for(let r=12;r<12+data.length;r++) ws.getRange(`F${r}`).formulas=[[`=B${r}/$B$5`]];
  styleBody(ws, `A12:H${11+data.length}`); statusFormatting(ws, `A12:A${11+data.length}`); statusFormatting(ws, `G12:G${11+data.length}`);
  ws.getRange("B5:B7").setNumberFormat("#,##0"); ws.getRange("B8").setNumberFormat("0.00%"); ws.getRange("F12:F20").setNumberFormat("0.00%"); ws.getRange("F5:F8").setNumberFormat("#,##0.00"); ws.getRange("B12:B20").setNumberFormat("#,##0"); ws.getRange("C12:E20").setNumberFormat("#,##0.00");
  [24,18,28,3,24,18,28,18].forEach((w,i)=>ws.getRange(`${colName(i)}:${colName(i)}`).format.columnWidth=w); ws.freezePanes.freezeRows(4);
}

// 5. 惠策明细-汇总核对
{
  const ws = sheet("5.惠策明细-汇总核对");
  styleTitle(ws, "惠策明细与店铺汇总核对", "按月份+平台+店铺重建。重要：源汇总的“笔数”明显采用比导出明细更细的底层粒度，因此差异是口径证据，不直接等同于源数据错误。", "H");
  writeBlock(ws, 3, 0, [["控制指标","数值","解释","","展示指标","数值","口径","结论"]]); styleHeader(ws,"A4:H4");
  const hs = Object.fromEntries(summary.huice_summary_recon.map(x=>[x.result,x.group_count]));
  writeBlock(ws,4,0,[
    ["核对组合总数",summary.controls.huice_summary_groups,"月度+平台+店铺","","源汇总成功笔数",null,"详见明细逐组合","粒度不可直接重建"],
    ["完全一致组合",summary.controls.huice_summary_consistent_groups,"当前无完全一致组合","","明细重建成功订单数",null,"一行通常对应一个平台订单","粒度不可直接重建"],
    ["差异组合",hs["差异"]||0,"两侧都有但数值不同","","成功笔数差异",null,"源汇总-明细重建","粒度不可直接重建"],
    ["仅明细重建组合",hs["仅明细重建"]||0,"源汇总缺少对应组合","","可直接复核率",null,"完全一致/总组合","需进一步解释"],
  ]);
  ws.getRange("E6").formulas=[["=SUM('6.惠策汇总核对明细'!D5:D202)"]]; ws.getRange("F6").formulas=[["=SUM('6.惠策汇总核对明细'!E5:E202)"]];
  ws.getRange("E7").formulas=[["=SUM('6.惠策汇总核对明细'!E5:E202)"]]; ws.getRange("F7").formulas=[["=SUM('6.惠策汇总核对明细'!F5:F202)"]];
  ws.getRange("E8").formulas=[["=SUM('6.惠策汇总核对明细'!F5:F202)"]]; ws.getRange("F8").formulas=[["=E8"]];
  ws.getRange("E9").values=[[summary.controls.huice_summary_consistent_groups]]; ws.getRange("F9").formulas=[["=E9/B5"]];
  styleBody(ws,"A5:H8"); statusFormatting(ws,"H5:H8"); ws.getRange("B5:B8").setNumberFormat("#,##0"); ws.getRange("E5:F8").setNumberFormat("#,##0"); ws.getRange("F9").setNumberFormat("0.00%");
  ws.getRange("A11:H11").merge(); ws.getRange("A11").values=[["审阅结论"]]; styleSection(ws,"A11:H11");
  ws.getRange("A12:H14").merge(); ws.getRange("A12").values=[["惠策店铺汇总中的成功、单边及金额不一致“笔数”无法由当前一行一订单的惠策明细直接重建。建议将本项结论表述为“存在粒度/口径差异，现有资料不足以逐笔重建”，不应在未取得底层应收/实收流水的情况下认定为数据错报。金额差额及分店铺证据已保留在明细页。"]]; ws.getRange("A12:H14").format={fill:C.amber,font:{color:C.amberText},wrapText:true,verticalAlignment:"center",borders:{preset:"outside",style:"thin",color:C.line}};
  [24,18,34,3,24,18,28,24].forEach((w,i)=>ws.getRange(`${colName(i)}:${colName(i)}`).format.columnWidth=w); ws.freezePanes.freezeRows(4);
}

// 7. OMS-SAP汇总
{
  const ws=sheet("7.OMS-SAP汇总");
  styleTitle(ws,"OMS发货与SAP标准发票（2C）核对汇总","主键：OMS.document_no = SAP.OMS销售单号；辅键：物料编码+销售单位。SAP仅保留标准发票（2C）。","H");
  writeBlock(ws,3,0,[["金额控制","SAP","OMS","差异","数量控制","SAP","OMS","结论"]]); styleHeader(ws,"A4:H4");
  writeBlock(ws,4,0,[
    ["匹配/全量金额",summary.controls.oms_sap_sap_amount,summary.controls.oms_sap_oms_amount,null,"全量发票/匹配发货数量",summary.controls.sap2c_quantity,summary.controls.oms_sap_oms_quantity,"数量差异"],
    ["一致键金额",summary.oms_sap_summary.find(x=>x.result==="数量金额一致").sap_amount,summary.oms_sap_summary.find(x=>x.result==="数量金额一致").oms_amount,null,"一致键数量",summary.oms_sap_summary.find(x=>x.result==="数量金额一致").sap_qty,summary.oms_sap_summary.find(x=>x.result==="数量金额一致").oms_qty,"数量金额一致"],
    ["仅SAP金额",summary.oms_sap_summary.find(x=>x.result==="仅SAP").sap_amount,0,null,"仅SAP数量",summary.oms_sap_summary.find(x=>x.result==="仅SAP").sap_qty,0,"仅SAP"],
  ]);
  for(let r=5;r<=7;r++) ws.getRange(`D${r}`).formulas=[[`=C${r}-B${r}`]];
  styleBody(ws,"A5:H7"); statusFormatting(ws,"H5:H7"); ws.getRange("B5:D7").setNumberFormat("#,##0.00"); ws.getRange("F5:G7").setNumberFormat("#,##0");
  ws.getRange("A9:H9").merge(); ws.getRange("A9").values=[["逐键核对结果分布"]]; styleSection(ws,"A9:H9");
  writeBlock(ws,9,0,[["结果","键数","SAP数量","OMS数量","数量差异","SAP金额","OMS金额","金额差异"]]); styleHeader(ws,"A10:H10");
  writeBlock(ws,10,0,summary.oms_sap_summary.map(x=>[x.result,x.key_count,clean(x.sap_qty),clean(x.oms_qty),clean(x.quantity_difference),clean(x.sap_amount),clean(x.oms_amount),clean(x.amount_difference)]));
  styleBody(ws,"A11:H12"); statusFormatting(ws,"A11:A12"); ws.getRange("B11:E12").setNumberFormat("#,##0"); ws.getRange("F11:H12").setNumberFormat("#,##0.00");
  ws.getRange("A14:B14").values=[["匹配键覆盖率",null]]; ws.getRange("B14").formulas=[["=B12/SUM(B11:B12)"]]; ws.getRange("A15:B15").values=[["发票数量覆盖率",null]]; ws.getRange("B15").formulas=[["=G5/F5"]]; ws.getRange("B14:B15").setNumberFormat("0.00%"); styleBody(ws,"A14:B15");
  [24,18,18,18,26,18,18,22].forEach((w,i)=>ws.getRange(`${colName(i)}:${colName(i)}`).format.columnWidth=w); ws.freezePanes.freezeRows(4);
}

// 9. OMS-惠策汇总
{
  const ws=sheet("9.OMS-惠策汇总");
  styleTitle(ws,"OMS未开票池与惠策账单核对汇总","OMS先剔除已与SAP发票匹配的部分，再按月度+客户/店铺映射与惠策净应收比较。当前店铺映射覆盖不足，结论仅作为风险定位。","H");
  const oh = Object.fromEntries(summary.oms_huice_summary.map(x=>[x.result,x]));
  writeBlock(ws,3,0,[["金额控制","数值","说明","","映射控制","数值","说明","结论"]]); styleHeader(ws,"A4:H4");
  writeBlock(ws,4,0,[
    ["OMS候选池金额",summary.controls.oms_huice_pool_amount,"未被SAP发票覆盖的OMS池","","已映射组合OMS金额",(oh["仅OMS"]?.oms_amount||0)+(oh["金额差异"]?.oms_amount||0),"仅OMS+金额差异","暂不能总体钩稽"],
    ["惠策净应收（已映射且有账单）",oh["金额差异"]?.huice_net_receivable||0,"92个月度店铺组合","","店铺未映射OMS金额",oh["店铺未映射"]?.oms_amount||0,"客户无法映射至惠策店铺","映射不足"],
    ["已映射有账单差异",oh["金额差异"]?.amount_difference||0,"OMS-惠策净应收","","客户总数",summary.controls.mapping_customers,"OMS客户主数据",""],
    ["仅OMS金额",oh["仅OMS"]?.oms_amount||0,"有映射但惠策无对应账单","","高置信映射客户",summary.controls.mapping_high_confidence,"自动映射置信度较高",""],
    ["店铺未映射金额",oh["店铺未映射"]?.oms_amount||0,"需人工映射后再核对","","映射覆盖率",null,"高置信/客户总数",""],
  ]);
  ws.getRange("F9").formulas=[["=F8/F7"]]; ws.getRange("H5").values=[["暂不能总体钩稽"]]; ws.getRange("H6").values=[["映射不足"]]; ws.getRange("H7").values=[["待解释"]]; ws.getRange("H8").values=[["待解释"]]; ws.getRange("H9").values=[["映射不足"]];
  styleBody(ws,"A5:H9"); statusFormatting(ws,"H5:H9"); ws.getRange("B5:B9").setNumberFormat("#,##0.00"); ws.getRange("F5:F8").setNumberFormat("#,##0.00"); ws.getRange("F9").setNumberFormat("0.00%");
  ws.getRange("A11:H11").merge(); ws.getRange("A11").values=[["结果分布"]]; styleSection(ws,"A11:H11");
  writeBlock(ws,11,0,[["结果","组合数","OMS金额","惠策净应收","金额差异","适用结论","明细页","后续动作"]]); styleHeader(ws,"A12:H12");
  writeBlock(ws,12,0,summary.oms_huice_summary.map(x=>[x.result,x.group_count,x.oms_amount,x.huice_net_receivable,x.amount_difference,x.result==="金额差异"?"差异待解释":"资料/映射不完整","10.OMS-惠策明细",x.result==="店铺未映射"?"先复核店铺映射":"核查结算时点及调整项"]));
  styleBody(ws,"A13:H15"); statusFormatting(ws,"A13:A15"); statusFormatting(ws,"F13:F15"); ws.getRange("B13:B15").setNumberFormat("#,##0"); ws.getRange("C13:E15").setNumberFormat("#,##0.00");
  [28,18,34,3,26,20,30,28].forEach((w,i)=>ws.getRange(`${colName(i)}:${colName(i)}`).format.columnWidth=w); ws.freezePanes.freezeRows(4);
}

detailSheet("4.旺店通-惠策明细","旺店通－惠策逐订单明细",`工作簿展示异常优先的前 ${wdtEmbedded.rows.length.toLocaleString()} 行；完整 ${summary.detail_exports.wdt_huice_detail.rows.toLocaleString()} 行见“12.完整明细索引”。`,wdtEmbedded,{platform_order_no:24,internal_orders:28,reconcile_ids:30,result:20});
detailSheet("6.惠策汇总核对明细","惠策明细－店铺汇总逐组合明细","逐月、逐平台、逐店铺列示源汇总与明细重建数值及差额；差异需结合粒度限制解读。",details.huice,{shop:30,result:16});
detailSheet("8.OMS-SAP明细","OMS－SAP逐销售单/物料/单位明细","SAP仅保留标准发票（2C）；“仅SAP”表示该组合未在当前OMS SQL结果中找到。",details.sap,{oms_sales_no:22,sap_invoice_nos:25,result:18});
detailSheet("10.OMS-惠策明细","OMS未开票池－惠策逐月店铺明细","按出库月及客户/店铺映射汇总；请结合“11.店铺映射”审阅自动映射置信度。",details.omsHuice,{customer_name:36,mapped_shop:30,result:18});
detailSheet("11.店铺映射","OMS客户至惠策平台/店铺映射","自动高置信映射可用于初步核对；未映射或低置信项目需人工确认，不应作为最终钩稽依据。",details.mapping,{customer_name:42,mapped_shop:34,mapping_status:18});

// 12. 完整明细索引
{
  const ws=sheet("12.完整明细索引");
  styleTitle(ws,"完整明细文件索引","Excel工作簿仅嵌入可审阅行数；超过Excel实用容量的全量结果以CSV保留，可由核对程序重复生成。","F");
  writeBlock(ws,3,0,[["核对链","完整明细行数","工作簿嵌入行数","完整CSV绝对路径","工作簿页面","说明"]]); styleHeader(ws,"A4:F4");
  const cfg=[
    ["旺店通－惠策","wdt_huice_detail","4.旺店通-惠策明细","异常优先抽取前5万行"],
    ["惠策明细－汇总","huice_summary_detail","6.惠策汇总核对明细","全部嵌入"],
    ["OMS－SAP","oms_sap_detail","8.OMS-SAP明细","全部嵌入"],
    ["OMS－惠策","oms_huice_detail","10.OMS-惠策明细","全部嵌入"],
    ["客户－店铺映射","customer_shop_mapping","11.店铺映射","全部嵌入"],
  ];
  writeBlock(ws,4,0,cfg.map(([label,key,page,note])=>[label,summary.detail_exports[key].rows,key==="wdt_huice_detail"?wdtEmbedded.rows.length:summary.detail_exports[key].workbook_rows,path.resolve("/Users/aatrox/Desktop/miaoke sales to c",summary.detail_exports[key].path),page,key==="wdt_huice_detail"?"异常优先嵌入前2万行":note]));
  styleBody(ws,"A5:F9"); ws.getRange("B5:C9").setNumberFormat("#,##0");
  [24,18,18,88,28,28].forEach((w,i)=>ws.getRange(`${colName(i)}:${colName(i)}`).format.columnWidth=w); ws.getRange("D5:D9").format.wrapText=true; ws.freezePanes.freezeRows(4);
}

await fs.mkdir(outputDir,{recursive:true});
const previewDir=path.join(outputDir,"_qa_previews"); await fs.mkdir(previewDir,{recursive:true});

const inspect1=await workbook.inspect({kind:"table",range:"1.总汇总!A1:K15",include:"values,formulas",tableMaxRows:20,tableMaxCols:12,maxChars:8000});
console.log("TOTAL_SUMMARY_INSPECT\n"+inspect1.ndjson);
const inspect2=await workbook.inspect({kind:"table",range:"7.OMS-SAP汇总!A1:H15",include:"values,formulas",tableMaxRows:20,tableMaxCols:10,maxChars:6000});
console.log("OMS_SAP_INSPECT\n"+inspect2.ndjson);
const errors=await workbook.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",options:{useRegex:true,maxResults:300},summary:"final formula error scan",maxChars:6000});
console.log("FORMULA_ERRORS\n"+errors.ndjson);

for (const name of names) {
  const ws=sheet(name); const used=ws.getUsedRange(true);
  const maxRows = name.includes("明细") || name==="11.店铺映射" ? 24 : 22;
  const maxCols = name==="1.总汇总" ? 11 : Math.min(used.columnCount || 8, 12);
  const range=`A1:${colName(maxCols-1)}${maxRows}`;
  const blob=await workbook.render({sheetName:name,range,scale:1.25,format:"png"});
  await fs.writeFile(path.join(previewDir,`${name.replaceAll("/","_")}.png`),new Uint8Array(await blob.arrayBuffer()));
}

const out=await SpreadsheetFile.exportXlsx(workbook); await out.save(outputFile);
console.log(JSON.stringify({outputFile,previewDir,sheets:names,wdtEmbedded:wdtEmbedded.rows.length},null,2));
