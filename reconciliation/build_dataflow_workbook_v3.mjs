import fs from "node:fs/promises";
import path from "node:path";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

const baseDir = "/Users/aatrox/Desktop/miaoke sales to c";
const oldDir = path.join(baseDir,"reconciliation/output_full");
const flowDir = path.join(baseDir,"reconciliation/output_flow_exploration");
const hcDir = path.join(baseDir,"reconciliation/output_huice_oms");
const outputDir = path.join(baseDir,"outputs/sales_toc_huice_oms_reconciliation_20260726");
const outputFile = path.join(outputDir,"销售ToC数据流核对底稿_V3_惠策对OMS_2025年12月至2026年6月.xlsx");
const readJson = async p => JSON.parse(await fs.readFile(p,"utf8"));
const base = await readJson(path.join(oldDir,"summary.json"));
const flow = await readJson(path.join(flowDir,"exploration_summary.json"));
const hc = await readJson(path.join(hcDir,"huice_oms_summary.json"));
const data = {
  wdtHuice: await readJson(path.join(oldDir,"wdt_huice_detail_workbook.json")),
  internal: await readJson(path.join(hcDir,"huice_detail_summary_month_shop_recon_v3_workbook.json")),
  material: await readJson(path.join(hcDir,"huice_shop_material_month_workbook.json")),
  hcOms: await readJson(path.join(hcDir,"huice_oms_month_item_recon_workbook.json")),
  shop: await readJson(path.join(hcDir,"huice_oms_month_shop_recon_workbook.json")),
  map: await readJson(path.join(hcDir,"huice_oms_shop_map_workbook.json")),
  omsSap: await readJson(path.join(oldDir,"oms_sap_detail_workbook.json")),
};
data.wdtHuice.rows=data.wdtHuice.rows.slice(0,15000);

const wb=Workbook.create();
const sheetNames=[
  "1.数据流总览","2.口径与匹配规则","3.订单物料桥接汇总","4.订单物料桥接明细",
  "5.惠策内部汇总","6.惠策内部明细","7.惠策店铺物料汇总",
  "8.惠策-OMS汇总","9.惠策-OMS明细","10.OMS-SAP汇总","11.OMS-SAP明细",
  "12.惠策OMS店铺映射","13.完整明细索引",
];
for(const n of sheetNames)wb.worksheets.add(n);
const C={navy:"#17365D",blue:"#2F75B5",pale:"#DDEBF7",pale2:"#EAF3F8",white:"#FFFFFF",text:"#203040",line:"#B4C6E7",gray:"#F2F2F2",green:"#E2F0D9",greenText:"#375623",amber:"#FFF2CC",amberText:"#7F6000",red:"#FCE4D6",redText:"#9C0006"};
const ws=n=>wb.worksheets.getItem(n);
const col=i=>{let n=i+1,s="";while(n){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26);}return s;};
const clean=v=>v===undefined||v===null||(typeof v==="number"&&!Number.isFinite(v))?null:v;
function write(s,r,c,rows){if(rows.length&&rows[0].length)s.getRangeByIndexes(r,c,rows.length,rows[0].length).values=rows.map(x=>x.map(clean));}
function title(s,t,sub,last="J"){s.showGridLines=false;s.getRange(`A1:${last}1`).merge();s.getRange("A1").values=[[t]];s.getRange(`A1:${last}1`).format={fill:C.navy,font:{bold:true,color:C.white,size:16},verticalAlignment:"center"};s.getRange(`A1:${last}1`).format.rowHeight=30;s.getRange(`A2:${last}2`).merge();s.getRange("A2").values=[[sub]];s.getRange(`A2:${last}2`).format={fill:C.pale2,font:{italic:true,color:C.text,size:10},wrapText:true,verticalAlignment:"center"};s.getRange(`A2:${last}2`).format.rowHeight=38;}
function header(s,r){s.getRange(r).format={fill:C.blue,font:{bold:true,color:C.white},horizontalAlignment:"center",verticalAlignment:"center",wrapText:true,borders:{preset:"all",style:"thin",color:C.line}};}
function section(s,r){s.getRange(r).format={fill:C.pale,font:{bold:true,color:C.navy,size:11},borders:{bottom:{style:"medium",color:C.blue}}};}
function body(s,r){s.getRange(r).format={font:{color:C.text,size:10},verticalAlignment:"center",borders:{insideHorizontal:{style:"thin",color:"#E7E6E6"}}};}
function status(s,r){const x=s.getRange(r);x.conditionalFormats.add("containsText",{text:"差异",format:{fill:C.red,font:{bold:true,color:C.redText}}});x.conditionalFormats.add("containsText",{text:"仅",format:{fill:C.amber,font:{bold:true,color:C.amberText}}});x.conditionalFormats.add("containsText",{text:"一致",format:{fill:C.green,font:{bold:true,color:C.greenText}}});x.conditionalFormats.add("containsText",{text:"低置信",format:{fill:C.red,font:{bold:true,color:C.redText}}});x.conditionalFormats.add("containsText",{text:"待复核",format:{fill:C.amber,font:{bold:true,color:C.amberText}}});}
function detail(name,t,sub,d,widths={}){const s=ws(name),last=col(d.headers.length-1),titleLast=col(Math.min(d.headers.length,10)-1);title(s,t,sub,titleLast);write(s,3,0,[d.headers]);header(s,`A4:${last}4`);for(let i=0;i<d.rows.length;i+=3000)write(s,4+i,0,d.rows.slice(i,i+3000));const end=4+d.rows.length;if(d.rows.length){body(s,`A5:${last}${end}`);for(const key of ["result","mapping_status","final_status"]){const ix=d.headers.indexOf(key);if(ix>=0)status(s,`${col(ix)}5:${col(ix)}${end}`);}}d.headers.forEach((h,i)=>{let w=widths[h]||(/name|shop|orders|invoice|reconcile/.test(h)?25:/amount|difference|receivable|cash/.test(h)?17:14);s.getRange(`${col(i)}:${col(i)}`).format.columnWidth=w;if(/amount|difference|receivable|cash|score|rate|share/.test(h))s.getRange(`${col(i)}5:${col(i)}${end}`).setNumberFormat(/rate|share/.test(h)?"0.00%":"#,##0.00");else if(/qty|count|rows|lines|groups|orders/.test(h))s.getRange(`${col(i)}5:${col(i)}${end}`).setNumberFormat("#,##0");});s.freezePanes.freezeRows(4);s.freezePanes.freezeColumns(Math.min(2,d.headers.length));}
const omsY005=flow.totals.find(x=>x.source==="OMS_Y005");
const exactSap=base.oms_sap_summary.find(x=>x.result==="数量金额一致");
const onlySap=base.oms_sap_summary.find(x=>x.result==="仅SAP");

// 1 数据流总览
{
 const s=ws("1.数据流总览");title(s,"销售 ToC 数据流核对总览（惠策对OMS版）","旺店通仅作为订单及物料桥接；惠策本期账单金额作为OMS日结的核对来源；OMS月结Y001继续与SAP标准发票（2C）核对。","K");
 s.getRange("A4:K4").merge();s.getRange("A4").values=[["修订后的数据流"]];section(s,"A4:K4");
 write(s,4,0,[["步骤","上游","下游","核对/桥接颗粒度","上游控制量","下游/已覆盖","覆盖率","差异","性质","明细页","结论"]]);header(s,"A5:K5");
 write(s,5,0,[
  [1,"旺店通订单行","惠策明细","原始单号=平台订单号；取得SAP物料",null,null,null,null,"物料桥接","4.订单物料桥接明细","不用于OMS金额结论"],
  [2,"惠策明细","惠策原始店铺汇总","账期月+平台+店铺（两侧唯一共同颗粒度）",null,null,null,null,"内部完整性","6.惠策内部明细","原始汇总无物料字段"],
  [3,"惠策明细+物料桥","惠策店铺物料重建","业务月+店铺+SAP物料",null,null,null,null,"派生汇总","7.惠策店铺物料汇总","供OMS核对使用"],
  [4,"惠策本期账单","OMS日结Y005","业务月+店铺/客户映射+SAP物料",null,null,null,null,"主核对","9.惠策-OMS明细","金额主核对；数量仅作桥接参考"],
  [5,"OMS月结Y001","SAP标准发票（2C）","OMS销售单号+物料+销售单位",null,null,null,null,"主核对","11.OMS-SAP明细","沿用已验证强主键"],
 ]);
 s.getRange("E6").formulas=[["='3.订单物料桥接汇总'!B5"]];s.getRange("F6").formulas=[["='3.订单物料桥接汇总'!B6"]];s.getRange("G6").formulas=[["='3.订单物料桥接汇总'!B7"]];s.getRange("H6").formulas=[["=E6-F6"]];
 s.getRange("E7").formulas=[["='5.惠策内部汇总'!B5"]];s.getRange("F7").formulas=[["='5.惠策内部汇总'!B6"]];s.getRange("G7").formulas=[["='5.惠策内部汇总'!B14/SUM('5.惠策内部汇总'!B12:B15)"]];s.getRange("H7").formulas=[["=E7-F7"]];
 s.getRange("E8").formulas=[["='3.订单物料桥接汇总'!F5"]];s.getRange("F8").formulas=[["='3.订单物料桥接汇总'!F6"]];s.getRange("G8").formulas=[["='3.订单物料桥接汇总'!F9"]];s.getRange("H8").formulas=[["=E8-F8"]];
 s.getRange("E9").formulas=[["='8.惠策-OMS汇总'!B5"]];s.getRange("F9").formulas=[["='8.惠策-OMS汇总'!F5"]];s.getRange("G9").formulas=[["='8.惠策-OMS汇总'!B7"]];s.getRange("H9").formulas=[["=F9-E9"]];
 s.getRange("E10").formulas=[["='10.OMS-SAP汇总'!B5"]];s.getRange("F10").formulas=[["='10.OMS-SAP汇总'!C5"]];s.getRange("G10").formulas=[["='10.OMS-SAP汇总'!B10"]];s.getRange("H10").formulas=[["=F10-E10"]];
 body(s,"A6:K10");status(s,"K6:K10");s.getRange("E6:F10").setNumberFormat("#,##0.00");s.getRange("G6:G10").setNumberFormat("0.00%");s.getRange("H6:H10").setNumberFormat("#,##0.00");
 s.getRange("A12:K14").merge();s.getRange("A12").values=[["重要限制：惠策明细与原始店铺汇总都没有物料、商品或数量字段。因此不能把两份原始文件直接按“店铺+物料”比较。本底稿通过平台订单号从旺店通取得物料后重建惠策店铺物料汇总；旺店通金额不进入惠策—OMS金额结论。"]];s.getRange("A12:K14").format={fill:C.amber,font:{bold:true,color:C.amberText},wrapText:true,verticalAlignment:"center",borders:{preset:"outside",style:"thin",color:C.line}};
 [8,22,23,34,18,18,14,18,18,26,30].forEach((w,i)=>s.getRange(`${col(i)}:${col(i)}`).format.columnWidth=w);s.freezePanes.freezeRows(5);
}

// 2 口径规则
{
 const s=ws("2.口径与匹配规则");title(s,"字段口径与分层匹配规则","先确认原始字段可用性，再进行订单物料桥接、月度归并和差异剔除；任何桥接字段均保留来源说明。","J");
 write(s,3,0,[["层级","数据链","颗粒度/键","金额口径","数量口径","判断条件","处理","可自动关闭","主要风险","复核页面"]]);header(s,"A4:J4");
 write(s,4,0,[
  [1,"旺店通→惠策","原始单号=平台订单号","不作为OMS金额","旺店通商品数量","订单号命中","补充SAP物料","否","订单下定/发货状态与账单状态不同","3-4"],
  [2,"惠策明细→原始店铺汇总","账期月+平台+店铺","净应收/净实收（含往期）","无","两侧总额及分类金额差≤0.01","内部完整性","仅完全一致组","汇总无物料字段","5-6"],
  [3,"惠策物料重建","业务月+店铺+SAP物料","本期应收、本期实收","桥接数量，非惠策原生","按旺店通行金额绝对值分摊；零金额按数量","形成派生汇总","否","83.28%订单月份可补物料","7"],
  [4,"惠策→OMS日结Y005","业务月+店铺/客户映射+SAP物料","OMS金额分别比较本期应收与本期实收","OMS数量与桥接数量分列","金额差≤0.01","一致组剔除，差异保留","仅金额一致组","结算时点、平台费用、退补及映射","8-9"],
  [5,"OMS月结Y001→SAP 2C","OMS销售单号+物料+单位","含税金额","双方原生数量","数量相同且金额差≤0.01","匹配后剔除","是","仅SAP记录保留","10-11"],
 ]);body(s,"A5:J9");status(s,"H5:H9");
 s.getRange("A11:J11").merge();s.getRange("A11").values=[["本期金额定义"]];section(s,"A11:J11");
 write(s,11,0,[["字段","公式/取值","目的","说明"],["惠策本期应收","正应收金额－负应收金额","对OMS发运金额","排除反复滚入后续账期的往期应收"],["惠策本期实收","收款金额（正实收）－退款金额（负实收）","辅助比较实际结算","可能包含结算时点差和平台费用影响"],["月份","业务日期月份；为空取账期结束月份","与OMS发运月比较","原始店铺汇总内部核对另用账期月份"]]);header(s,"A12:D12");body(s,"A13:D15");
 [8,24,31,24,22,24,26,18,35,18].forEach((w,i)=>s.getRange(`${col(i)}:${col(i)}`).format.columnWidth=w);s.freezePanes.freezeRows(4);
}

// 3 订单物料桥接
{
 const s=ws("3.订单物料桥接汇总");title(s,"旺店通订单用于惠策物料桥接","旺店通仅提供平台订单号、店铺、SAP物料及商品数量；OMS金额主核对使用惠策本期账单金额。","H");
 write(s,3,0,[["订单控制","数值","说明","","金额控制","数值","说明","结论"]]);header(s,"A4:H4");
 write(s,4,0,[["惠策订单月份组合",hc.huice_order_months,"业务月+店铺+平台订单号","","惠策本期应收",hc.huice_receivable,"总体","桥接总体"],["取得物料的订单月份",hc.mapped_order_months,"订单号连接旺店通","","已补物料本期应收",hc.mapped_receivable,"分摊前保持订单总额","可用于店铺物料汇总"],["订单月份覆盖率",null,"已取得物料/总体","","本期实收",hc.huice_cash,"总体","辅助金额口径"],["未取得物料订单月份",null,"总体-已取得物料","","已补物料本期实收",hc.mapped_cash,"分摊前保持订单总额","覆盖率单列"],["店铺物料月度组合",hc.shop_material_groups,"派生汇总行数","","应收金额覆盖率",null,"已补物料/总体","非100%完整"]]);
 s.getRange("B7").formulas=[["=B6/B5"]];s.getRange("B8").formulas=[["=B5-B6"]];s.getRange("F9").formulas=[["=F6/F5"]];body(s,"A5:H9");status(s,"H5:H9");s.getRange("B5:B6").setNumberFormat("#,##0");s.getRange("B7").setNumberFormat("0.00%");s.getRange("B8:B9").setNumberFormat("#,##0");s.getRange("F5:F8").setNumberFormat("#,##0.00");s.getRange("F9").setNumberFormat("0.00%");
 s.getRange("A11:H13").merge();s.getRange("A11").values=[["物料分摊规则：同一平台订单含多个物料时，按旺店通各物料“分摊后总价”的绝对值占比分配惠策本期应收/实收；若订单行金额合计为零，则按数量绝对值；金额和数量均为零时平均分配。该步骤只改变分析颗粒度，不改变订单层面的惠策金额合计。"]];s.getRange("A11:H13").format={fill:C.amber,font:{color:C.amberText},wrapText:true,verticalAlignment:"center",borders:{preset:"outside",style:"thin",color:C.line}};
 [25,18,34,3,25,18,32,28].forEach((w,i)=>s.getRange(`${col(i)}:${col(i)}`).format.columnWidth=w);s.freezePanes.freezeRows(4);
}

// 5 惠策内部汇总
{
 const s=ws("5.惠策内部汇总");title(s,"惠策明细与原始店铺汇总内部核对","原始店铺汇总没有物料字段，内部核对只能在账期月+平台+店铺共同颗粒度进行；店铺物料表由明细经订单桥接另行重建。","H");
 const ir=Object.fromEntries(hc.internal_results.map(x=>[x.result,x]));
 write(s,3,0,[["内部控制","数值","说明","","物料重建","数值","说明","结论"]]);header(s,"A4:H4");
 write(s,4,0,[["明细重建金额",Object.values(ir).reduce((a,x)=>a+(x.detail_receivable||0),0),"含往期净应收，按账期月","","店铺物料组合",hc.shop_material_groups,"业务月+店铺+SAP物料","派生而非原始汇总"],["原始汇总金额",Object.values(ir).reduce((a,x)=>a+(x.summary_receivable||0),0),"成功+差异应收+单边应收","","可补物料订单月份",hc.mapped_order_months,"平台订单号桥接","覆盖不完整"],["金额差异",null,"明细重建-原始汇总","","物料覆盖订单率",null,"可补物料/惠策订单月份","不得称原始汇总物料核对"],["完全一致组合",ir["金额一致"]?.groups||0,"金额差≤0.01","","原始物料字段数",0,"两份惠策原始文件均为0","无法直接店铺+物料"]]);
 s.getRange("B7").formulas=[["=B5-B6"]];s.getRange("F7").formulas=[["='3.订单物料桥接汇总'!B7"]];body(s,"A5:H8");status(s,"H5:H8");s.getRange("B5:B7").setNumberFormat("#,##0.00");s.getRange("B8:F8").setNumberFormat("#,##0");s.getRange("F7").setNumberFormat("0.00%");
 s.getRange("A10:H10").merge();s.getRange("A10").values=[["内部核对结果分布"]];section(s,"A10:H10");write(s,10,0,[["结果","组合数","明细净应收","原始汇总应收","差异","处理","明细页","说明"]]);header(s,"A11:H11");write(s,11,0,hc.internal_results.map(x=>[x.result,x.groups,x.detail_receivable,x.summary_receivable,x.receivable_difference,x.result==="金额一致"?"关闭":"保留复核","6.惠策内部明细",x.result.includes("仅")?"店铺名称/平台/文件覆盖不一致":""]));body(s,`A12:H${11+hc.internal_results.length}`);status(s,`A12:A${11+hc.internal_results.length}`);s.getRange("B12:B20").setNumberFormat("#,##0");s.getRange("C12:E20").setNumberFormat("#,##0.00");
 [25,18,34,18,25,18,34,34].forEach((w,i)=>s.getRange(`${col(i)}:${col(i)}`).format.columnWidth=w);s.freezePanes.freezeRows(4);
}

// 8 惠策-OMS
{
 const s=ws("8.惠策-OMS汇总");title(s,"惠策本期账单与OMS日结Y005核对汇总","主维度：业务月+惠策店铺/OMS客户映射+SAP物料。分别比较惠策本期应收、本期实收与OMS分摊金额；数量不作为惠策原生结论。","H");
 write(s,3,0,[["惠策金额控制","数值","说明","","OMS控制","数值","说明","结论"]]);header(s,"A4:H4");
 write(s,4,0,[["惠策本期应收",hc.huice_receivable,"全部惠策订单月份","","OMS Y005金额",omsY005.amount,"全部日结","两侧总体"],["已补物料本期应收",hc.mapped_receivable,"可进入物料核对","","OMS Y005数量",omsY005.quantity,"OMS原生数量","数量只做参考"],["应收桥接覆盖率",null,"已补物料/全部","","月店铺物料核对组",hc.recon_groups,"含仅一侧","差异池"],["惠策本期实收",hc.huice_cash,"辅助口径","","金额一致组",hc.receivable_exact_groups+hc.cash_exact_groups,"应收或实收差≤0.01","一致组较少"],["实收桥接覆盖率",null,"已补物料/全部","","低置信店铺映射",hc.low_confidence_shops,"须人工复核","不可自动关闭"]]);
 s.getRange("B7").formulas=[["='3.订单物料桥接汇总'!F9"]];s.getRange("B9").formulas=[["='3.订单物料桥接汇总'!F8/'3.订单物料桥接汇总'!F7"]];body(s,"A5:H9");status(s,"H5:H9");s.getRange("B5:B6").setNumberFormat("#,##0.00");s.getRange("B7").setNumberFormat("0.00%");s.getRange("B8").setNumberFormat("#,##0.00");s.getRange("B9").setNumberFormat("0.00%");s.getRange("F5").setNumberFormat("#,##0.00");s.getRange("F6:F9").setNumberFormat("#,##0");
 s.getRange("A11:H11").merge();s.getRange("A11").values=[["月度店铺物料结果分布"]];section(s,"A11:H11");write(s,11,0,[["结果","组数","惠策本期应收","惠策本期实收","OMS金额","桥接数量","OMS数量","处理"]]);header(s,"A12:H12");write(s,12,0,hc.item_results.map(x=>[x.result,x.groups,x.huice_receivable,x.huice_cash,x.oms_amount,x.bridge_qty,x.oms_qty,(x.result==="应收金额一致"||x.result==="实收金额一致")?"金额一致剔除":"保留差异"]));body(s,`A13:H${12+hc.item_results.length}`);status(s,`A13:A${12+hc.item_results.length}`);s.getRange("B13:B20").setNumberFormat("#,##0");s.getRange("C13:E20").setNumberFormat("#,##0.00");s.getRange("F13:G20").setNumberFormat("#,##0");
 let start=14+hc.item_results.length;s.getRange(`A${start}:H${start}`).merge();s.getRange(`A${start}`).values=[["按业务月份汇总"]];section(s,`A${start}:H${start}`);write(s,start,0,[["月份","组数","惠策本期应收","惠策本期实收","OMS金额","应收差异","实收差异","判断"]]);header(s,`A${start+1}:H${start+1}`);write(s,start+1,0,hc.month_results.map(x=>[x.huice_month,x.groups,x.huice_receivable,x.huice_cash,x.oms_amount,x.receivable_difference,x.cash_difference,Math.abs(x.receivable_difference)<=Math.abs(x.cash_difference)?"应收更接近":"实收更接近"]));body(s,`A${start+2}:H${start+1+hc.month_results.length}`);s.getRange(`B${start+2}:B${start+10}`).setNumberFormat("#,##0");s.getRange(`C${start+2}:G${start+10}`).setNumberFormat("#,##0.00");const noteRow=start+3+hc.month_results.length;s.getRange(`A${noteRow}:H${noteRow+1}`).merge();s.getRange(`A${noteRow}`).values=[["期间提示：现有惠策账单文件从2026年1月开始，2025年12月仅包含后续账期中落入12月业务日期的少量记录，不能视为完整的2025年12月惠策账单总体；该月差异应单独标记为资料覆盖差异。"]];s.getRange(`A${noteRow}:H${noteRow+1}`).format={fill:C.amber,font:{color:C.amberText,bold:true},wrapText:true,verticalAlignment:"center",borders:{preset:"outside",style:"thin",color:C.line}};
 [24,18,25,18,24,18,30,28].forEach((w,i)=>s.getRange(`${col(i)}:${col(i)}`).format.columnWidth=w);s.freezePanes.freezeRows(4);
}

// 10 OMS-SAP
{
 const s=ws("10.OMS-SAP汇总");title(s,"OMS月结与SAP标准发票（2C）核对汇总","沿用已验证规则：OMS Y001月结，主键为OMS销售单号，辅键为物料编码+销售单位。","H");
 write(s,3,0,[["金额指标","SAP","OMS","差异","数量指标","SAP","OMS","结论"]]);header(s,"A4:H4");write(s,4,0,[["一致键金额",exactSap.sap_amount,exactSap.oms_amount,null,"全量/匹配数量",base.controls.sap2c_quantity,base.controls.oms_sap_oms_quantity,"数量差异"],["一致键尾差",exactSap.sap_amount,exactSap.oms_amount,null,"一致键数量",exactSap.sap_qty,exactSap.oms_qty,"数量金额一致"],["仅SAP",onlySap.sap_amount,0,null,"仅SAP数量",onlySap.sap_qty,0,"仅SAP"]]);for(let r=5;r<=7;r++)s.getRange(`D${r}`).formulas=[[`=C${r}-B${r}`]];body(s,"A5:H7");status(s,"H5:H7");s.getRange("B5:D7").setNumberFormat("#,##0.00");s.getRange("F5:G7").setNumberFormat("#,##0");write(s,9,0,[["匹配键覆盖率",null],["发票数量覆盖率",null]]);s.getRange("B10").formulas=[["=11776/(11776+80)"]];s.getRange("B11").formulas=[["=G5/F5"]];s.getRange("B10:B11").setNumberFormat("0.00%");body(s,"A10:B11");[24,18,18,18,26,18,18,22].forEach((w,i)=>s.getRange(`${col(i)}:${col(i)}`).format.columnWidth=w);s.freezePanes.freezeRows(4);
}

detail("4.订单物料桥接明细","旺店通—惠策订单桥接明细","异常优先展示15,000行；该表验证平台订单号覆盖与金额差异，但仅物料归属进入惠策—OMS核对。",data.wdtHuice,{platform_order_no:25,internal_orders:30,reconcile_ids:30,result:22});
detail("6.惠策内部明细","惠策明细—原始店铺汇总内部明细","按账期月+平台+店铺核对；这是两份原始惠策文件唯一共同颗粒度。",data.internal,{shop:34,result:18});
detail("7.惠策店铺物料汇总","惠策月度店铺物料重建汇总","物料由平台订单号连接旺店通取得；金额来自惠策本期账单，旺店通金额不作为核对金额。",data.material,{huice_shop:34,customer_name:38,material_code:17,mapping_status:18});
detail("9.惠策-OMS明细","惠策本期账单—OMS日结月度店铺物料明细","分别展示惠策本期应收、本期实收、OMS金额及差异；桥接数量不得解释为惠策原生数量。",data.hcOms,{huice_shop:34,customer_name:38,material_code:17,result:20,mapping_status:18});
detail("11.OMS-SAP明细","OMS月结—SAP标准发票（2C）明细","逐OMS销售单号、物料编码、销售单位核对。",data.omsSap,{oms_sales_no:24,sap_invoice_nos:25,result:20});
detail("12.惠策OMS店铺映射","惠策店铺—旺店通店铺—OMS客户映射","先以共同平台订单号确定惠策店铺到旺店通店铺，再连接既有OMS客户映射；低置信不得自动关闭。",data.map,{huice_shop:34,wdt_shop:34,customer_name:42,final_status:18});

// 13 索引
{
 const s=ws("13.完整明细索引");title(s,"完整明细与程序索引","工作簿嵌入全部中间汇总及审阅明细；大体量订单桥接明细保留完整CSV。","F");
 write(s,3,0,[["数据链","完整行数","工作簿行数","完整文件","工作簿页面","说明"]]);header(s,"A4:F4");
 write(s,4,0,[
  ["旺店通—惠策订单桥接",base.detail_exports.wdt_huice_detail.rows,data.wdtHuice.rows.length,path.join(oldDir,"wdt_huice_detail.csv"),"4.订单物料桥接明细","旺店通仅作桥接"],
  ["惠策内部月店铺",data.internal.rows.length,data.internal.rows.length,path.join(hcDir,"huice_detail_summary_month_shop_recon_v3.csv"),"6.惠策内部明细","原始共同颗粒度"],
  ["惠策店铺物料重建",data.material.rows.length,data.material.rows.length,path.join(hcDir,"huice_shop_material_month.csv"),"7.惠策店铺物料汇总","派生汇总"],
  ["惠策—OMS月店铺物料",hc.recon_groups,data.hcOms.rows.length,path.join(hcDir,"huice_oms_month_item_recon.csv"),"9.惠策-OMS明细","全部嵌入"],
  ["惠策—OMS店铺映射",data.map.rows.length,data.map.rows.length,path.join(hcDir,"huice_oms_shop_map.csv"),"12.惠策OMS店铺映射","全部嵌入"],
  ["OMS月结—SAP",base.detail_exports.oms_sap_detail.rows,data.omsSap.rows.length,path.join(oldDir,"oms_sap_detail.csv"),"11.OMS-SAP明细","全部嵌入"],
 ]);body(s,"A5:F10");s.getRange("B5:C10").setNumberFormat("#,##0");[28,18,18,86,27,28].forEach((w,i)=>s.getRange(`${col(i)}:${col(i)}`).format.columnWidth=w);s.getRange("D5:D10").format.wrapText=true;s.freezePanes.freezeRows(4);
}

await fs.mkdir(outputDir,{recursive:true});const previews=path.join(outputDir,"_qa_previews");await fs.mkdir(previews,{recursive:true});
console.log("OVERVIEW\n"+(await wb.inspect({kind:"table",range:"1.数据流总览!A1:K14",include:"values,formulas",tableMaxRows:18,tableMaxCols:12,maxChars:9000})).ndjson);
console.log("HC_OMS\n"+(await wb.inspect({kind:"table",range:"8.惠策-OMS汇总!A1:H30",include:"values,formulas",tableMaxRows:35,tableMaxCols:10,maxChars:12000})).ndjson);
console.log("ERRORS\n"+(await wb.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",options:{useRegex:true,maxResults:300},summary:"formula errors",maxChars:5000})).ndjson);
for(const name of sheetNames){const s=ws(name),used=s.getUsedRange(true),maxCols=Math.min(used.columnCount||8,name.includes("明细")?10:12),maxRows=name.includes("明细")?22:24;const blob=await wb.render({sheetName:name,range:`A1:${col(maxCols-1)}${maxRows}`,scale:1.15,format:"png"});await fs.writeFile(path.join(previews,`${name}.png`),new Uint8Array(await blob.arrayBuffer()));}
const out=await SpreadsheetFile.exportXlsx(wb);await out.save(outputFile);console.log(JSON.stringify({outputFile,sheets:sheetNames,previews},null,2));
