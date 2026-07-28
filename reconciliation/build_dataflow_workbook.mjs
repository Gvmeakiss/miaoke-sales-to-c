import fs from "node:fs/promises";
import path from "node:path";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

const baseDir = "/Users/aatrox/Desktop/miaoke sales to c";
const oldDir = path.join(baseDir, "reconciliation/output_full");
const flowDir = path.join(baseDir, "reconciliation/output_flow_exploration");
const outputDir = path.join(baseDir, "outputs/sales_toc_dataflow_reconciliation_20260726");
const outputFile = path.join(outputDir, "销售ToC数据流核对底稿_V2_2025年12月至2026年6月.xlsx");

const readJson = async p => JSON.parse(await fs.readFile(p, "utf8"));
const base = await readJson(path.join(oldDir, "summary.json"));
const flow = await readJson(path.join(flowDir, "exploration_summary.json"));
const data = {
  wdtHuice: await readJson(path.join(oldDir, "wdt_huice_detail_workbook.json")),
  omsSap: await readJson(path.join(oldDir, "oms_sap_detail_workbook.json")),
  huiceInternal: await readJson(path.join(oldDir, "huice_summary_detail_workbook.json")),
  wdtOms: await readJson(path.join(flowDir, "wdt_oms_sales_day_item_recon_workbook.json")),
  typeMap: await readJson(path.join(flowDir, "wdt_oms_order_type_recon_workbook.json")),
  shopMap: await readJson(path.join(flowDir, "wdt_oms_shop_map_workbook.json")),
};
data.wdtHuice.rows = data.wdtHuice.rows.slice(0, 15000);
data.wdtOms.rows = data.wdtOms.rows.slice(0, 20000);

const wb = Workbook.create();
const sheets = [
  "1.数据流总览", "2.匹配瀑布与规则", "3.旺店通-惠策汇总", "4.旺店通-惠策明细",
  "5.旺店通-OMS汇总", "6.旺店通-OMS明细", "7.业务类型映射",
  "8.OMS-SAP汇总", "9.OMS-SAP明细", "10.惠策内部汇总", "11.惠策内部明细",
  "12.客户店铺映射", "13.完整明细索引",
];
for (const name of sheets) wb.worksheets.add(name);

const C = {
  navy: "#17365D", blue: "#2F75B5", pale: "#DDEBF7", pale2: "#EAF3F8",
  white: "#FFFFFF", text: "#203040", line: "#B4C6E7", gray: "#F2F2F2",
  green: "#E2F0D9", greenText: "#375623", amber: "#FFF2CC", amberText: "#7F6000",
  red: "#FCE4D6", redText: "#9C0006",
};
const ws = n => wb.worksheets.getItem(n);
const col = i => { let n=i+1,s=""; while(n){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26);} return s; };
const clean = v => v === undefined || v === null || (typeof v === "number" && !Number.isFinite(v)) ? null : v;
function write(sheet,row,colIndex,rows){ if(rows.length&&rows[0].length) sheet.getRangeByIndexes(row,colIndex,rows.length,rows[0].length).values=rows.map(r=>r.map(clean)); }
function title(sheet,text,sub,last="J"){
  sheet.showGridLines=false; sheet.getRange(`A1:${last}1`).merge(); sheet.getRange("A1").values=[[text]];
  sheet.getRange(`A1:${last}1`).format={fill:C.navy,font:{bold:true,color:C.white,size:16},verticalAlignment:"center"}; sheet.getRange(`A1:${last}1`).format.rowHeight=30;
  sheet.getRange(`A2:${last}2`).merge(); sheet.getRange("A2").values=[[sub]];
  sheet.getRange(`A2:${last}2`).format={fill:C.pale2,font:{italic:true,color:C.text,size:10},wrapText:true,verticalAlignment:"center"}; sheet.getRange(`A2:${last}2`).format.rowHeight=34;
}
function header(sheet,range){ sheet.getRange(range).format={fill:C.blue,font:{bold:true,color:C.white},horizontalAlignment:"center",verticalAlignment:"center",wrapText:true,borders:{preset:"all",style:"thin",color:C.line}}; }
function section(sheet,range){ sheet.getRange(range).format={fill:C.pale,font:{bold:true,color:C.navy,size:11},borders:{bottom:{style:"medium",color:C.blue}}}; }
function body(sheet,range){ sheet.getRange(range).format={font:{color:C.text,size:10},verticalAlignment:"center",borders:{insideHorizontal:{style:"thin",color:"#E7E6E6"}}}; }
function status(sheet,range){
  const r=sheet.getRange(range);
  r.conditionalFormats.add("containsText",{text:"一致",format:{fill:C.green,font:{bold:true,color:C.greenText}}});
  r.conditionalFormats.add("containsText",{text:"差异",format:{fill:C.red,font:{bold:true,color:C.redText}}});
  r.conditionalFormats.add("containsText",{text:"仅",format:{fill:C.amber,font:{bold:true,color:C.amberText}}});
  r.conditionalFormats.add("containsText",{text:"复核",format:{fill:C.amber,font:{bold:true,color:C.amberText}}});
  r.conditionalFormats.add("containsText",{text:"不可",format:{fill:C.red,font:{bold:true,color:C.redText}}});
}
function detail(name,t,sub,d,widths={}){
  const s=ws(name), last=col(d.headers.length-1); title(s,t,sub,last); write(s,3,0,[d.headers]); header(s,`A4:${last}4`);
  for(let i=0;i<d.rows.length;i+=4000) write(s,4+i,0,d.rows.slice(i,i+4000));
  const end=4+d.rows.length;
  if(d.rows.length){body(s,`A5:${last}${end}`); const ri=d.headers.findIndex(h=>h==="result"||h==="mapping_status"); if(ri>=0)status(s,`${col(ri)}5:${col(ri)}${end}`);}
  d.headers.forEach((h,i)=>{let w=widths[h]||(/name|shop|orders|invoice|reconcile/.test(h)?24:/amount|difference|receivable/.test(h)?16:13);s.getRange(`${col(i)}:${col(i)}`).format.columnWidth=w;
    if(/amount|difference|receivable|cash|score/.test(h))s.getRange(`${col(i)}5:${col(i)}${end}`).setNumberFormat("#,##0.00");
    else if(/qty|count|rows|lines|groups/.test(h))s.getRange(`${col(i)}5:${col(i)}${end}`).setNumberFormat("#,##0");
  });
  s.freezePanes.freezeRows(4);s.freezePanes.freezeColumns(Math.min(2,d.headers.length));
}

const salesCtl=flow.sales_controls[0];
const salesResults=Object.fromEntries(flow.sales_day_item_results.map(x=>[x.result,x]));
const exactAmount=(salesResults["数量金额一致"]?.wdt_amount||0)+(salesResults["金额一致数量差异"]?.wdt_amount||0);
const exactAmountGroups=(salesResults["数量金额一致"]?.groups||0)+(salesResults["金额一致数量差异"]?.groups||0);
const exactQty=(salesResults["数量金额一致"]?.wdt_qty||0)+(salesResults["数量一致金额差异"]?.wdt_qty||0);
const exactQtyGroups=(salesResults["数量金额一致"]?.groups||0)+(salesResults["数量一致金额差异"]?.groups||0);

// 1 数据流总览
{
  const s=ws("1.数据流总览"); title(s,"销售 ToC 数据流核对总览","主线按订单、发运、月结、开票顺序组织；惠策作为订单资金分支。OMS日结与月结分别使用Y005和Y001，禁止重复计入。","K");
  s.getRange("A4:K4").merge();s.getRange("A4").values=[["业务数据流与核对结论"]];section(s,"A4:K4");
  write(s,4,0,[["步骤","上游数据","下游数据","主要颗粒度/字段","上游数量/金额","下游数量/金额","匹配或覆盖","差异","结论","明细页","说明"]]);header(s,"A5:K5");
  write(s,5,0,[
    [1,"旺店通订单","惠策明细","原始单号=平台订单号",null,null,null,null,"","4.旺店通-惠策明细","订单与账单分支"],
    [2,"旺店通网店销售","OMS日结Y005","发货日+店铺映射+SAP物料；先按金额剔除",null,null,null,null,"","6.旺店通-OMS明细","多订单归并至日结颗粒度"],
    [3,"OMS月结Y001","SAP标准发票（2C）","OMS销售单号+物料+单位",null,null,null,null,"","9.OMS-SAP明细","已验证的强主键核对"],
    [4,"惠策明细","惠策店铺汇总","月度+平台+店铺+状态分类",null,null,null,null,"","11.惠策内部明细","完整性控制，不连接OMS"],
  ]);
  s.getRange("E6").formulas=[["='3.旺店通-惠策汇总'!B5"]];s.getRange("F6").formulas=[["='3.旺店通-惠策汇总'!B7"]];s.getRange("G6").formulas=[["='3.旺店通-惠策汇总'!B8"]];s.getRange("H6").formulas=[["=E6-F6"]];s.getRange("I6").values=[["部分覆盖，剩余待解释"]];
  s.getRange("E7").formulas=[["='5.旺店通-OMS汇总'!B5"]];s.getRange("F7").formulas=[["='5.旺店通-OMS汇总'!B6"]];s.getRange("G7").formulas=[["='5.旺店通-OMS汇总'!B8"]];s.getRange("H7").formulas=[["=E7-F7"]];s.getRange("I7").values=[["金额核对可作为主规则"]];
  s.getRange("E8").formulas=[["='8.OMS-SAP汇总'!B6"]];s.getRange("F8").formulas=[["='8.OMS-SAP汇总'!C6"]];s.getRange("G8").formulas=[["='8.OMS-SAP汇总'!B10"]];s.getRange("H8").formulas=[["=F8-E8"]];s.getRange("I8").values=[["数量金额一致"]];
  s.getRange("E9").formulas=[["='10.惠策内部汇总'!B7"]];s.getRange("F9").formulas=[["='10.惠策内部汇总'!F5"]];s.getRange("G9").formulas=[["='10.惠策内部汇总'!F7"]];s.getRange("H9").formulas=[["=E9-F9"]];s.getRange("I9").values=[["粒度不同，不可直接重建"]];
  body(s,"A6:K9");status(s,"I6:I9");s.getRange("E6:F9").setNumberFormat("#,##0.00");s.getRange("G6:G9").setNumberFormat("0.00%");s.getRange("H6:H9").setNumberFormat("#,##0.00");
  s.getRange("A11:K11").merge();s.getRange("A11").values=[["总体控制量：用于识别OMS日结/月结重复，不直接作为逐单结论"]];section(s,"A11:K11");
  write(s,11,0,[["数据集","数量","金额","相对旺店通数量","相对旺店通金额","判断"],...flow.totals.map(x=>[x.source,x.quantity,x.amount,null,null,x.source==="WDT"?"基准":x.source==="OMS_Y005"?"日结候选":"月结候选"])]);header(s,"A12:F12");
  for(let r=13;r<=15;r++){s.getRange(`D${r}`).formulas=[[`=B${r}/$B$13`]];s.getRange(`E${r}`).formulas=[[`=C${r}/$C$13`]];}body(s,"A13:F15");s.getRange("B13:C15").setNumberFormat("#,##0.00");s.getRange("D13:E15").setNumberFormat("0.00%");
  [8,23,24,34,18,18,14,18,22,24,34].forEach((w,i)=>s.getRange(`${col(i)}:${col(i)}`).format.columnWidth=w);s.freezePanes.freezeRows(5);
}

// 2 匹配瀑布
{
  const s=ws("2.匹配瀑布与规则"); title(s,"匹配瀑布与剔除规则","先使用强条件匹配并从两侧总体中剔除，再逐层放宽颗粒度；不得把同一记录在多个层级重复计入。","J");
  write(s,3,0,[["层级","适用链路","归并颗粒度","判断条件","命中组数","覆盖数量/订单","覆盖金额","覆盖率","处理","风险说明"]]);header(s,"A4:J4");
  const exactOrders=base.wdt_huice_summary.find(x=>x.result==="单号金额一致");
  write(s,4,0,[
    [1,"旺店通→惠策","平台订单号","单号存在且金额差≤0.01",exactOrders.order_count,exactOrders.order_count,exactOrders.wdt_amount,null,"匹配后剔除","强主键"],
    [2,"旺店通→OMS日结","发货日+店铺映射+SAP物料","金额差≤0.01",exactAmountGroups,null,exactAmount,null,"匹配后剔除","数量口径可不同"],
    [3,"旺店通→OMS日结","同上","数量完全一致",exactQtyGroups,exactQty,null,null,"作为数量辅助结论","金额可能受退补影响"],
    [4,"非网店订单→OMS业务类型","日期+物料+业务类型","数量金额严格一致",null,null,null,null,"按类型分别剔除","详见业务类型映射"],
    [5,"OMS月结→SAP","OMS销售单号+物料+单位","数量、金额均一致",base.controls.oms_sap_exact_keys,base.oms_sap_summary.find(x=>x.result==="数量金额一致").sap_qty,base.oms_sap_summary.find(x=>x.result==="数量金额一致").sap_amount,null,"匹配后剔除","合计尾差1.38元"],
    [6,"剩余总体","保持原颗粒度","不得因总额接近直接视为一致",null,null,null,null,"列入差异明细","避免用总体抵销掩盖错配"],
  ]);
  s.getRange("H5").formulas=[["=F5/'3.旺店通-惠策汇总'!B5"]];s.getRange("H6").formulas=[["=G6/'5.旺店通-OMS汇总'!B5"]];s.getRange("H7").formulas=[["=F7/'5.旺店通-OMS汇总'!F5"]];s.getRange("H9").formulas=[["=F9/'8.OMS-SAP汇总'!F5"]];
  body(s,"A5:J10");s.getRange("E5:G10").setNumberFormat("#,##0.00");s.getRange("H5:H10").setNumberFormat("0.00%");
  s.getRange("A12:J12").merge();s.getRange("A12").values=[["审阅判断：金额一致覆盖率较高时，可支持收入/发运金额链路；数量差异须单独保留，不能因金额一致而自动关闭。"]];s.getRange("A12:J12").format={fill:C.amber,font:{color:C.amberText,bold:true},wrapText:true,borders:{preset:"outside",style:"thin",color:C.line}};
  [8,24,28,28,14,18,18,14,18,34].forEach((w,i)=>s.getRange(`${col(i)}:${col(i)}`).format.columnWidth=w);s.freezePanes.freezeRows(4);
}

// 3 旺店通-惠策汇总
{
  const s=ws("3.旺店通-惠策汇总");title(s,"旺店通订单与惠策账单核对汇总","订单资金分支。主键：旺店通原始单号=惠策平台订单号；先匹配单号金额一致记录，再剔除并分析剩余。","H");
  write(s,3,0,[["控制指标","数值","说明","","金额指标","数值","说明","结论"]]);header(s,"A4:H4");
  write(s,4,0,[["旺店通平台订单数",base.controls.wdt_platform_orders,"旺店通总体","","旺店通总金额",base.controls.wdt_amount,"订单应收","部分覆盖"],["惠策平台订单数",base.controls.huice_platform_orders,"惠策总体","","惠策净应收",base.controls.huice_net_receivable,"账单口径","口径差异"],["匹配订单数",base.controls.wdt_huice_matched_orders,"两侧单号存在","","匹配旺店通金额",base.controls.wdt_huice_matched_wdt_amount,"匹配单号",""],["订单覆盖率",null,"匹配/旺店通","","匹配惠策金额",base.controls.wdt_huice_matched_huice_amount,"匹配单号","" ]]);
  s.getRange("B8").formulas=[["=B7/B5"]];body(s,"A5:H8");status(s,"H5:H8");s.getRange("B5:B7").setNumberFormat("#,##0");s.getRange("B8").setNumberFormat("0.00%");s.getRange("F5:F8").setNumberFormat("#,##0.00");
  s.getRange("A10:H10").merge();s.getRange("A10").values=[["匹配结果分层"]];section(s,"A10:H10");write(s,10,0,[["结果","订单数","旺店通金额","惠策净应收","差异","订单占比","处理","明细页"]]);header(s,"A11:H11");
  write(s,11,0,base.wdt_huice_summary.map(x=>[x.result,x.order_count,x.wdt_amount,x.huice_net_receivable,x.amount_difference,null,x.result==="单号金额一致"?"剔除":"保留复核","4.旺店通-惠策明细"]));
  for(let r=12;r<12+base.wdt_huice_summary.length;r++)s.getRange(`F${r}`).formulas=[[`=B${r}/$B$5`]];body(s,`A12:H${11+base.wdt_huice_summary.length}`);status(s,`A12:A${11+base.wdt_huice_summary.length}`);s.getRange("B12:B20").setNumberFormat("#,##0");s.getRange("C12:E20").setNumberFormat("#,##0.00");s.getRange("F12:F20").setNumberFormat("0.00%");
  [24,17,20,18,24,18,28,22].forEach((w,i)=>s.getRange(`${col(i)}:${col(i)}`).format.columnWidth=w);s.freezePanes.freezeRows(4);
}

// 5 旺店通-OMS汇总
{
  const s=ws("5.旺店通-OMS汇总");title(s,"旺店通网店销售与OMS日结核对汇总","主规则：旺店通网店销售按发货日+店铺映射+SAP物料归并，与OMS Y005日结比较。金额一致作为主结论，数量单独评价。","H");
  write(s,3,0,[["金额控制","数值","说明","","数量控制","数值","说明","结论"]]);header(s,"A4:H4");
  write(s,4,0,[["旺店通网店销售金额",salesCtl.wdt_amount,"已映射日+店铺+物料总体","","旺店通网店销售数量",salesCtl.wdt_qty,"同口径","金额主匹配"],["金额一致覆盖金额",exactAmount,"差额≤0.01元","","数量一致覆盖数量",exactQty,"数量完全相等","数量辅助"],["剩余待解释金额",null,"总体-金额一致","","剩余数量",null,"总体-数量一致","保留差异"],["金额覆盖率",null,"金额一致/总体","","数量覆盖率",null,"数量一致/总体",""],["金额一致组数",exactAmountGroups,"多订单已归并","","数量一致组数",exactQtyGroups,"逐日店铺物料组","" ]]);
  s.getRange("B7").formulas=[["=B5-B6"]];s.getRange("F7").formulas=[["=F5-F6"]];s.getRange("B8").formulas=[["=B6/B5"]];s.getRange("F8").formulas=[["=F6/F5"]];body(s,"A5:H9");status(s,"H5:H9");s.getRange("B5:B7").setNumberFormat("#,##0.00");s.getRange("F5:F7").setNumberFormat("#,##0");s.getRange("B8:F8").setNumberFormat("0.00%");s.getRange("B9:F9").setNumberFormat("#,##0");
  s.getRange("A11:H11").merge();s.getRange("A11").values=[["逐日店铺物料结果分布"]];section(s,"A11:H11");write(s,11,0,[["结果","组数","旺店通数量","OMS数量","旺店通金额","OMS金额","处理","解释"]]);header(s,"A12:H12");
  write(s,12,0,flow.sales_day_item_results.map(x=>[x.result,x.groups,x.wdt_qty,x.oms_qty,x.wdt_amount,x.oms_amount,(x.result==="数量金额一致"||x.result==="金额一致数量差异")?"金额匹配剔除":x.result==="数量一致金额差异"?"数量匹配，金额复核":"保留差异",x.result==="金额一致数量差异"?"金额同源性强，数量口径不同":""]));
  body(s,"A13:H18");status(s,"A13:A18");s.getRange("B13:D18").setNumberFormat("#,##0");s.getRange("E13:F18").setNumberFormat("#,##0.00");
  [25,16,18,18,25,18,25,34].forEach((w,i)=>s.getRange(`${col(i)}:${col(i)}`).format.columnWidth=w);s.freezePanes.freezeRows(4);
}

// 7 业务类型映射
{
  const s=ws("7.业务类型映射");title(s,"旺店通订单类型与OMS业务类型候选映射","辅助规则由日期+物料的数量金额严格一致记录反推。非网店类型覆盖率较高的映射可用于批量剔除；低覆盖率仅作线索。","I");
  write(s,3,0,[data.typeMap.headers]);header(s,"A4:I4");write(s,4,0,data.typeMap.rows);body(s,`A5:I${4+data.typeMap.rows.length}`);
  const h=data.typeMap.headers;h.forEach((x,i)=>{s.getRange(`${col(i)}:${col(i)}`).format.columnWidth=/type/.test(x)?20:16;if(/amount/.test(x))s.getRange(`${col(i)}5:${col(i)}20`).setNumberFormat("#,##0.00");else if(/coverage/.test(x))s.getRange(`${col(i)}5:${col(i)}20`).setNumberFormat("0.00%");else if(/qty|groups/.test(x))s.getRange(`${col(i)}5:${col(i)}20`).setNumberFormat("#,##0");});
  s.getRange("A13:I13").merge();s.getRange("A13").values=[["建议自动规则：分销订单→Y051、样品发货→Z003、退货损失→Z004、赠品→Z001。线下订单→Z006覆盖较低，暂不自动关闭；网店销售仍以店铺映射后的金额匹配为主。"]];s.getRange("A13:I13").format={fill:C.amber,font:{color:C.amberText,bold:true},wrapText:true};s.freezePanes.freezeRows(4);
}

// 8 OMS-SAP
{
  const s=ws("8.OMS-SAP汇总");title(s,"OMS月结与SAP标准发票（2C）核对汇总","OMS仅使用Y001月结；主键：document_no=SAP.OMS销售单号，辅键为物料编码+销售单位。","H");
  const exact=base.oms_sap_summary.find(x=>x.result==="数量金额一致"), only=base.oms_sap_summary.find(x=>x.result==="仅SAP");
  write(s,3,0,[["金额指标","SAP","OMS","差异","数量指标","SAP","OMS","结论"]]);header(s,"A4:H4");write(s,4,0,[["一致键金额",exact.sap_amount,exact.oms_amount,null,"全量/匹配数量",base.controls.sap2c_quantity,base.controls.oms_sap_oms_quantity,"数量差异"],["一致键尾差",exact.sap_amount,exact.oms_amount,null,"一致键数量",exact.sap_qty,exact.oms_qty,"数量金额一致"],["仅SAP",only.sap_amount,0,null,"仅SAP数量",only.sap_qty,0,"仅SAP"]]);for(let r=5;r<=7;r++)s.getRange(`D${r}`).formulas=[[`=C${r}-B${r}`]];body(s,"A5:H7");status(s,"H5:H7");s.getRange("B5:D7").setNumberFormat("#,##0.00");s.getRange("F5:G7").setNumberFormat("#,##0");
  write(s,9,0,[["匹配键覆盖率",null],["发票数量覆盖率",null]]);s.getRange("B10").formulas=[["=11776/(11776+80)"]];s.getRange("B11").formulas=[["=G5/F5"]];s.getRange("B10:B11").setNumberFormat("0.00%");body(s,"A10:B11");
  [24,18,18,18,26,18,18,22].forEach((w,i)=>s.getRange(`${col(i)}:${col(i)}`).format.columnWidth=w);s.freezePanes.freezeRows(4);
}

// 10 惠策内部
{
  const s=ws("10.惠策内部汇总");title(s,"惠策明细与店铺汇总内部核对","用途是惠策内部完整性控制，不作为OMS日结桥表。汇总笔数采用更细粒度，当前32列明细无法完全重建。","H");
  write(s,3,0,[["指标","数值","说明","","重建指标","数值","口径","结论"]]);header(s,"A4:H4");
  write(s,4,0,[["源汇总组合",158,"月度+平台+店铺","","完全一致组合",0,"全部指标一致","粒度不可直接重建"],["明细额外组合",40,"汇总表无对应组合","","差异组合",158,"两侧有记录但指标不同","粒度不可直接重建"],["全部核对组合",198,"158+40","","可直接复核率",null,"完全一致/全部组合","不可作为OMS匹配键"]]);s.getRange("F7").formulas=[["=F5/B7"]];body(s,"A5:H7");status(s,"H5:H7");s.getRange("B5:B7").setNumberFormat("#,##0");s.getRange("F5:F6").setNumberFormat("#,##0");s.getRange("F7").setNumberFormat("0.00%");
  s.getRange("A9:H9").merge();s.getRange("A9").values=[["字段判断"]];section(s,"A9:H9");s.getRange("A10:H12").merge();s.getRange("A10").values=[["惠策店铺汇总的“月汇总流水号”（ZQHZ…）不出现在明细清单，也不与OMS business_no/document_no相交；它是统计记录编号，不是业务汇总单号。当前惠策明细没有物料或汇总单号字段，因此惠策不能直接承担OMS日结逐单桥接。"]];s.getRange("A10:H12").format={fill:C.amber,font:{color:C.amberText},wrapText:true,verticalAlignment:"center",borders:{preset:"outside",style:"thin",color:C.line}};
  [24,18,32,3,24,18,28,28].forEach((w,i)=>s.getRange(`${col(i)}:${col(i)}`).format.columnWidth=w);s.freezePanes.freezeRows(4);
}

detail("4.旺店通-惠策明细","旺店通－惠策逐订单明细","异常优先展示15,000行；完整673万行见索引。",data.wdtHuice,{platform_order_no:24,internal_orders:28,reconcile_ids:30,result:20});
detail("6.旺店通-OMS明细","旺店通网店销售－OMS日结逐日店铺物料明细","异常优先展示20,000行；金额差≤0.01视为金额一致，数量差额单独保留。",data.wdtOms,{ship_date:14,shop:30,customer_name:36,material_code:16,result:22});
detail("9.OMS-SAP明细","OMS月结－SAP标准发票（2C）明细","逐OMS销售单号、物料编码、销售单位核对。",data.omsSap,{oms_sales_no:23,sap_invoice_nos:25,result:20});
detail("11.惠策内部明细","惠策明细－店铺汇总逐组合明细","按月度、平台、店铺列示源汇总与明细重建差额；仅用于内部完整性检查。",data.huiceInternal,{shop:32,result:16});
detail("12.客户店铺映射","OMS客户－旺店通店铺候选映射","评分由7个月物料金额结构、名称及平台一致性组成。高置信可自动使用，待复核/低置信应结合金额命中率审阅。",data.shopMap,{customer_name:42,wdt_shop:32,mapping_status:18});

// 13 索引
{
  const s=ws("13.完整明细索引");title(s,"完整明细与程序索引","工作簿嵌入审阅样本；完整CSV和可重复执行程序保留在本地核对目录。","F");
  write(s,3,0,[["核对链","完整行数","工作簿行数","完整文件","工作簿页面","说明"]]);header(s,"A4:F4");
  write(s,4,0,[
    ["旺店通－惠策",base.detail_exports.wdt_huice_detail.rows,data.wdtHuice.rows.length,path.join(oldDir,"wdt_huice_detail.csv"),"4.旺店通-惠策明细","异常优先"],
    ["旺店通－OMS日结",Object.values(salesResults).reduce((a,x)=>a+x.groups,0),data.wdtOms.rows.length,path.join(flowDir,"wdt_oms_sales_day_item_recon.csv"),"6.旺店通-OMS明细","异常优先"],
    ["OMS月结－SAP",base.detail_exports.oms_sap_detail.rows,data.omsSap.rows.length,path.join(oldDir,"oms_sap_detail.csv"),"9.OMS-SAP明细","全部嵌入"],
    ["惠策内部",base.detail_exports.huice_summary_detail.rows,data.huiceInternal.rows.length,path.join(oldDir,"huice_summary_detail.csv"),"11.惠策内部明细","全部嵌入"],
    ["客户店铺映射",data.shopMap.rows.length,data.shopMap.rows.length,path.join(flowDir,"wdt_oms_shop_map.csv"),"12.客户店铺映射","全部嵌入"],
  ]);body(s,"A5:F9");s.getRange("B5:C9").setNumberFormat("#,##0");[24,18,18,82,26,24].forEach((w,i)=>s.getRange(`${col(i)}:${col(i)}`).format.columnWidth=w);s.getRange("D5:D9").format.wrapText=true;s.freezePanes.freezeRows(4);
}

await fs.mkdir(outputDir,{recursive:true});const previews=path.join(outputDir,"_qa_previews");await fs.mkdir(previews,{recursive:true});
const checks=await wb.inspect({kind:"table",range:"1.数据流总览!A1:K15",include:"values,formulas",tableMaxRows:20,tableMaxCols:12,maxChars:9000});console.log("OVERVIEW\n"+checks.ndjson);
const check2=await wb.inspect({kind:"table",range:"5.旺店通-OMS汇总!A1:H18",include:"values,formulas",tableMaxRows:22,tableMaxCols:10,maxChars:8000});console.log("WDT_OMS\n"+check2.ndjson);
const errors=await wb.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",options:{useRegex:true,maxResults:300},summary:"formula errors",maxChars:5000});console.log("ERRORS\n"+errors.ndjson);
for(const name of sheets){const s=ws(name),used=s.getUsedRange(true),maxCols=Math.min(used.columnCount||8,name.includes("明细")?10:12),maxRows=name.includes("明细")?22:20;const blob=await wb.render({sheetName:name,range:`A1:${col(maxCols-1)}${maxRows}`,scale:1.15,format:"png"});await fs.writeFile(path.join(previews,`${name}.png`),new Uint8Array(await blob.arrayBuffer()));}
const out=await SpreadsheetFile.exportXlsx(wb);await out.save(outputFile);console.log(JSON.stringify({outputFile,sheets,previews},null,2));
