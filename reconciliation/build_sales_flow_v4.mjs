import fs from "node:fs/promises";
import path from "node:path";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

const root="/Users/aatrox/Desktop/miaoke sales to c";
const dataDir=process.env.SALES_TOC_DATA_DIR||path.join(root,"reconciliation/output_flow_v4");
const outputDir=process.env.SALES_TOC_OUTPUT_DIR||path.join(root,"outputs/sales_toc_workpaper_20260101_20260630");
const outputFile=process.env.SALES_TOC_OUTPUT_FILE||path.join(outputDir,"销售ToC业务流程核对底稿_20260101-20260630.xlsx");
const read=async n=>JSON.parse(await fs.readFile(path.join(dataDir,n),"utf8"));
const S=await read("summary_v4.json");
const data={
 orderBill:await read("order_bill_recon_workbook.json"),billOms:await read("bill_oms_month_recon_workbook.json"),
 qty:await read("order_bill_oms_qty_recon_workbook.json"),omsSap:await read("oms_sap_field_map_workbook.json"),
 shop:await read("huice_shop_map_workbook.json"),
};
const wb=Workbook.create();
const names=["1.全局口径与总览","2.业务流程与单对单规则","3.订单-账单汇总","4.订单-账单明细","5.账单-OMS月结汇总","6.账单-OMS月结明细","7.数量核对汇总","8.数量核对明细","9.OMS月结-SAP汇总","10.OMS-SAP字段映射","11.月度全流程汇总","12.店铺客户映射","13.完整明细索引"];
for(const n of names)wb.worksheets.add(n);
const scope="统一核对期间：2026-01-01至2026-06-30（含首尾）";
const C={navy:"#17365D",blue:"#2F75B5",pale:"#DDEBF7",pale2:"#EAF3F8",white:"#FFFFFF",text:"#203040",line:"#B4C6E7",green:"#E2F0D9",greenText:"#375623",amber:"#FFF2CC",amberText:"#7F6000",red:"#FCE4D6",redText:"#9C0006"};
const ws=n=>wb.worksheets.getItem(n);
const col=i=>{let n=i+1,s="";while(n){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26);}return s;};
const clean=v=>v===undefined||v===null||(typeof v==="number"&&!Number.isFinite(v))?null:v;
function write(s,r,c,rows){if(rows.length&&rows[0].length)s.getRangeByIndexes(r,c,rows.length,rows[0].length).values=rows.map(x=>x.map(clean));}
function title(s,t,sub,last="J"){s.showGridLines=false;s.getRange(`A1:${last}1`).merge();s.getRange("A1").values=[[t]];s.getRange(`A1:${last}1`).format={fill:C.navy,font:{bold:true,color:C.white,size:16},verticalAlignment:"center"};s.getRange(`A1:${last}1`).format.rowHeight=30;s.getRange(`A2:${last}2`).merge();s.getRange("A2").values=[[`${scope}｜${sub}`]];s.getRange(`A2:${last}2`).format={fill:C.pale2,font:{italic:true,color:C.text,size:10},wrapText:true,verticalAlignment:"center"};s.getRange(`A2:${last}2`).format.rowHeight=38;}
function header(s,r){s.getRange(r).format={fill:C.blue,font:{bold:true,color:C.white},horizontalAlignment:"center",verticalAlignment:"center",wrapText:true,borders:{preset:"all",style:"thin",color:C.line}};}
function section(s,r){s.getRange(r).format={fill:C.pale,font:{bold:true,color:C.navy,size:11},borders:{bottom:{style:"medium",color:C.blue}}};}
function body(s,r){s.getRange(r).format={font:{color:C.text,size:10},verticalAlignment:"center",borders:{insideHorizontal:{style:"thin",color:"#E7E6E6"}}};}
function status(s,r){const x=s.getRange(r);x.conditionalFormats.add("containsText",{text:"差异",format:{fill:C.red,font:{bold:true,color:C.redText}}});x.conditionalFormats.add("containsText",{text:"未映射",format:{fill:C.red,font:{bold:true,color:C.redText}}});x.conditionalFormats.add("containsText",{text:"仅",format:{fill:C.amber,font:{bold:true,color:C.amberText}}});x.conditionalFormats.add("containsText",{text:"一致",format:{fill:C.green,font:{bold:true,color:C.greenText}}});x.conditionalFormats.add("containsText",{text:"待复核",format:{fill:C.amber,font:{bold:true,color:C.amberText}}});}
const headerZh={
 platform_order_no:"平台订单号",matchable:"可匹配标识",wdt_shop:"旺店通店铺",platform:"平台",huice_shop:"惠策店铺",
 ship_month:"发货月份",bill_month:"账单月份",internal_order_count:"内部订单数",huice_rows:"惠策明细行数",
 wdt_qty:"旺店通数量",wdt_amount:"旺店通金额",wdt_allocated_amount:"旺店通分摊金额",bill_receivable:"惠策本期应收",
 bill_cash:"惠策本期实收",receivable_difference:"应收差异",cash_difference:"实收差异",internal_orders:"旺店通内部单号",
 reconcile_ids:"惠策对账流水号",result:"核对结果",customer_code:"OMS客户编码",customer_name:"OMS客户名称",
 mapping_status:"映射置信度",mapping_source:"映射来源",bill_record_count:"账单记录数",success_count:"成功记录数",
 bill_success_amount:"对账成功金额",oms_docs:"OMS单据数",oms_qty:"OMS数量",oms_amount:"OMS金额",
 sap_assisted_qty:"SAP辅助数量",sap_assisted_amount:"SAP辅助金额",success_difference:"成功金额差异",
 sap_success_difference:"SAP辅助成功差异",source_rows:"源文件行数",material_code:"物料编码",billed_orders:"账单证据订单数",
 order_bill_qty:"订单账单数量",wdt_item_amount:"旺店通商品金额",qty_difference:"数量差异",oms_sales_no:"OMS销售单号",
 sales_unit:"销售单位",file_month:"SAP文件月份",outbound_month:"OMS出库月份",sap_invoice_nos:"SAP发票号",
 sap_rows:"SAP行数",oms_rows:"OMS行数",sap_qty:"SAP数量",quantity_difference:"数量差异",sap_amount:"SAP金额",
 amount_difference:"金额差异",mapped_qty:"映射数量",mapped_amount:"映射金额",mapping_result:"映射结果",source_result:"源核对结果"
};
function detail(name,t,sub,d,widths={},omit=[]){const keep=d.headers.map((h,i)=>({h,i})).filter(x=>!omit.includes(x.h)),headers=keep.map(x=>x.h),rows=keep.length===d.headers.length?d.rows:d.rows.map(r=>keep.map(x=>r[x.i])),displayHeaders=headers.map(h=>headerZh[h]?`${h}\n${headerZh[h]}`:h);const s=ws(name),last=col(headers.length-1),titleLast=col(Math.min(headers.length,10)-1);title(s,t,sub,titleLast);write(s,3,0,[displayHeaders]);header(s,`A4:${last}4`);s.getRange(`A4:${last}4`).format.rowHeight=34;for(let i=0;i<rows.length;i+=3000)write(s,4+i,0,rows.slice(i,i+3000));const end=4+rows.length;if(rows.length){body(s,`A5:${last}${end}`);for(const k of ["result","mapping_result"]){const ix=headers.indexOf(k);if(ix>=0)status(s,`${col(ix)}5:${col(ix)}${end}`);}}headers.forEach((h,i)=>{let w=widths[h]||(/name|shop|orders|invoice|reconcile/.test(h)?25:/amount|difference|receivable|cash/.test(h)?17:14);s.getRange(`${col(i)}:${col(i)}`).format.columnWidth=w;if(/amount|difference|receivable|cash/.test(h))s.getRange(`${col(i)}5:${col(i)}${end}`).setNumberFormat("#,##0.00");else if(/qty|count|rows|lines|groups|orders|keys/.test(h))s.getRange(`${col(i)}5:${col(i)}${end}`).setNumberFormat("#,##0");});s.freezePanes.freezeRows(4);s.freezePanes.freezeColumns(Math.min(2,headers.length));}
const by=(arr,key)=>Object.fromEntries(arr.map(x=>[x[key],x]));
const ob=by(S.order_bill_results,"result"),bo=by(S.bill_oms_results,"result"),qr=by(S.qty_results,"result"),os=by(S.oms_sap_results,"mapping_result"),ctl=S.controls;
const orderExact=(ob["单号应收金额一致"]?.groups||0)+(ob["单号实收金额一致"]?.groups||0);
const orderExactAmount=(ob["单号应收金额一致"]?.wdt_amount||0)+(ob["单号实收金额一致"]?.wdt_amount||0);
const orderMatchedBillAmount=(ob["单号应收金额一致"]?.bill_receivable||0)+(ob["单号实收金额一致"]?.bill_cash||0);
const orderExactQty=(ob["单号应收金额一致"]?.wdt_qty||0)+(ob["单号实收金额一致"]?.wdt_qty||0);
const orderReconAmount=S.order_bill_results.reduce((a,x)=>a+(x.wdt_amount||0),0);
const commonWdtQty=(qr["数量差异"]?.order_bill_qty||0)+(qr["数量一致"]?.order_bill_qty||0);
const commonOmsQty=(qr["数量差异"]?.oms_qty||0)+(qr["数量一致"]?.oms_qty||0);
const exactMap=os["双向字段一致"],sapOnly=os["SAP补数量金额"],omsOnly=os["仅OMS月结"]||{keys:0,oms_qty:0,oms_amount:0};
const dm=S.display_metrics||{};
const orderKeyCoverage=dm.order_key_matches/ctl.wdt_orders;
const huiceKeyCoverage=dm.order_key_matches/ctl.huice_orders;
const orderAmountExactCoverage=dm.order_amount_exact_groups/ctl.wdt_orders;
const billAmountRatio=ctl.huice_bill_success_amount/ctl.oms_month_amount;
const billMappingCoverage=ctl.mapped_bill_amount/ctl.huice_bill_receivable;
const billExactGroupRate=dm.bill_exact_groups/dm.bill_total_groups;
const mappingHighCoverage=dm.mapping_high_ar/ctl.huice_bill_receivable;
const qtyCommonGroupRate=dm.qty_exact_groups/dm.qty_common_groups;
const qtyNetRatio=commonWdtQty/commonOmsQty;
const bilateralSapKeyRate=exactMap.keys/(exactMap.keys+sapOnly.keys+omsOnly.keys);
const omsQtyCoverage=exactMap.oms_qty/ctl.oms_month_qty;
const sapQtyCoverage=exactMap.sap_qty/ctl.sap_full_qty;
const sapAmountCoverage=exactMap.sap_amount/ctl.sap_full_amount;

// 1 全局口径
{
 const s=ws(names[0]);title(s,"销售 ToC 核对底稿汇总", "旺店通订单→惠策店铺账单→OMS月结Y001→SAP标准发票（2C）；参考上年底稿集中展示金额和数量匹配。","K");
 s.getRange("A4:K4").merge();s.getRange("A4").values=[["金额匹配"]];section(s,"A4:K4");
 write(s,4,0,[["Sheet","参与匹配数据","匹配所用字段","核对方式","结论","旺店通金额","惠策金额","OMS月结金额","SAP发票金额","差异","匹配率"]]);header(s,"A5:K5");
 s.getRange("A5:K5").format.rowHeight=30;
 write(s,5,0,[
  ["1.订单—账单","旺店通订单—惠策明细","原始单号=平台订单号","逐单匹配","应收或实收金额一致",orderExactAmount,orderMatchedBillAmount,null,null,null,null],
  ["2.账单—OMS月结","惠策店铺汇总—OMS Y001","月份+店铺/OMS客户","总额及逐月店铺","成功金额/OMS月结金额",null,ctl.huice_bill_success_amount,ctl.oms_month_amount,null,null,null],
  ["3.OMS月结—SAP","OMS Y001—SAP标准发票（2C）","销售单号+物料+销售单位","逐键及总量","一致键数量金额相符",null,null,exactMap.oms_amount,exactMap.sap_amount,null,null],
 ]);
 s.getRange("J6").formulas=[["=G6-F6"]];s.getRange("K6").formulas=[[`=F6/${orderReconAmount}`]];
 s.getRange("J7").formulas=[["=H7-G7"]];s.getRange("K7").formulas=[["=G7/H7"]];
 s.getRange("J8").formulas=[["=I8-H8"]];s.getRange("K8").formulas=[[`=H8/${ctl.oms_month_amount}`]];
 body(s,"A6:K8");status(s,"E6:E8");s.getRange("F6:J8").setNumberFormat("#,##0.00;[Red](#,##0.00);-");s.getRange("K6:K8").setNumberFormat("0.00%");
 s.getRange("A10:K10").merge();s.getRange("A10").values=[["数量匹配"]];section(s,"A10:K10");
 write(s,10,0,[["Sheet","参与匹配数据","匹配所用字段","核对方式","结论","旺店通/订单账单数量","惠策数量","OMS月结数量","SAP发票数量","差异","匹配率"]]);header(s,"A11:K11");
 s.getRange("A11:K11").format.rowHeight=30;
 write(s,11,0,[
  ["1.订单账单—OMS","有账单证据的旺店通订单—OMS Y001","发货月+店铺/客户+物料","多对一归并","共同范围净数量比",commonWdtQty,null,commonOmsQty,null,null,null],
  ["2.OMS月结—SAP","OMS Y001—SAP标准发票（2C）","销售单号+物料+销售单位","逐键及总量","一致键数量相符",null,null,exactMap.oms_qty,exactMap.sap_qty,null,null],
 ]);
 s.getRange("J12").formulas=[["=H12-F12"]];s.getRange("K12").formulas=[["=F12/H12"]];s.getRange("J13").formulas=[["=I13-H13"]];s.getRange("K13").formulas=[[`=H13/${ctl.oms_month_qty}`]];
 body(s,"A12:K13");status(s,"E12:E13");s.getRange("F12:J13").setNumberFormat("#,##0;[Red](#,##0);-");s.getRange("K12:K13").setNumberFormat("0.00%");
 s.getRange("A15:K17").merge();s.getRange("A15").values=[["说明：惠策无商品数量，数量核对以惠策中出现的平台订单作为账单证据，提取对应旺店通商品数量后与OMS月结比较；订单账单—OMS的匹配率为共同范围净数量比，不替代逐组合数量一致率。金额、数量的完整差异明细分别见对应汇总页和明细页。"]];s.getRange("A15:K17").format={fill:C.amber,font:{color:C.amberText,bold:true},wrapText:true,verticalAlignment:"center",borders:{preset:"outside",style:"thin",color:"#C9A227"}};
 [18,30,31,20,24,18,18,18,18,18,15].forEach((w,i)=>s.getRange(`${col(i)}:${col(i)}`).format.columnWidth=w);s.freezePanes.freezeRows(5);
}

// 2 规则
{
 const s=ws(names[1]);title(s,"业务流程与单对单核对规则","每一环节独立钩稽，避免因下一环节缺字段而否定上一环节的有效匹配。","J");
 write(s,3,0,[["环节","数据A","数据B","强键/共同维度","金额比较","数量比较","匹配阈值","匹配后处理","限制","页面"]]);header(s,"A4:J4");
 write(s,4,0,[
  [1,"旺店通订单","惠策明细","旺店通原始单号=惠策平台订单号","订单金额分别对本期应收、本期实收","记录旺店通商品数量；惠策无商品数量","金额差≤0.01","一致组剔除","订单与账单状态时点不同","3-4"],
  [2,"惠策店铺汇总","OMS月结Y001","月份+惠策店铺/OMS客户映射","成功金额为主；应收、实收并列","账单笔数与OMS商品数量不可直接比较","金额差≤0.01","一致组剔除","平台费用、单边未关联、结算时点","5-6"],
  [3,"惠策有记录的订单数量","OMS月结Y001","发货月+店铺/客户+SAP物料","旺店通商品金额仅辅助","订单商品数量 vs OMS月结数量","数量差=0；总体比例另列","一致组剔除","惠策仅作为订单存在证据","7-8"],
  [4,"OMS月结Y001","SAP标准发票（2C）","OMS销售单号+物料+销售单位","含税金额","发票数量 vs OMS数量","金额差≤0.01且数量相等","生成双向映射",`${sapOnly.keys}个仅SAP、${omsOnly.keys}个仅OMS键保留`,"9-10"],
 ]);body(s,"A5:J8");
 s.getRange("A10:J10").merge();s.getRange("A10").values=[["SAP—OMS双向字段映射"]];section(s,"A10:J10");write(s,10,0,[["来源字段","补充到","映射后字段","使用条件","用途"],["OMS客户编码/客户名称/出库月份","SAP记录","店铺客户、业务月份","OMS销售单号+物料+单位匹配","按店铺/月度辅助汇总"],["SAP发票号/发票数量/含税金额","OMS记录","发票字段、SAP辅助金额","同一强键匹配","验证OMS金额数量及发票追踪"]]);header(s,"A11:E11");body(s,"A12:E13");
 [10,25,25,34,28,27,20,24,35,15].forEach((w,i)=>s.getRange(`${col(i)}:${col(i)}`).format.columnWidth=w);s.freezePanes.freezeRows(4);
}

// 3 订单账单汇总
{
 const s=ws(names[2]);title(s,"旺店通订单—惠策明细核对汇总","强键：旺店通原始单号=惠策平台订单号；金额分别比较惠策本期应收和本期实收。","H");
 write(s,3,0,[["订单覆盖指标","数值","说明","","金额/数量控制","数值","说明","结论"]]);header(s,"A4:H4");
 write(s,4,0,[["旺店通平台订单",ctl.wdt_orders,"按发货时间限定期间","","旺店通订单表头数量",S.order_bill_results.reduce((a,x)=>a+(x.wdt_qty||0),0),"非商品明细数量","控制总体"],["订单号已覆盖",dm.order_key_matches,"含金额一致及金额差异","","旺店通订单表头应收",S.order_bill_results.reduce((a,x)=>a+(x.wdt_amount||0),0),"与商品明细金额不同","控制总体"],["订单号+金额匹配",orderExact,"应收或实收金额一致","","匹配订单对应表头数量",orderExactQty,"有账单证据","匹配成功"],["旺店通订单号覆盖率",null,"订单号已覆盖/旺店通订单","","匹配订单金额",orderExactAmount,"匹配后可剔除","优先展示"],["订单号+金额匹配率",null,"金额匹配/旺店通订单","","匹配金额覆盖率",null,"匹配金额/订单表头应收","匹配指标"],["惠策订单号覆盖率",null,"订单号已覆盖/惠策订单","","惠策平台订单",ctl.huice_orders,"按业务日期限定期间","控制总体"]]);
 s.getRange("B8").formulas=[["=B6/B5"]];s.getRange("B9").formulas=[["=B7/B5"]];s.getRange("B10").formulas=[["=B6/F10"]];s.getRange("F9").formulas=[["=F8/F6"]];body(s,"A5:H10");status(s,"H5:H10");s.getRange("B5:B7").setNumberFormat("#,##0");s.getRange("B8:B10").setNumberFormat("0.00%");s.getRange("F5:F8").setNumberFormat("#,##0.00");s.getRange("F9").setNumberFormat("0.00%");s.getRange("F10").setNumberFormat("#,##0");
 s.getRange("A12:H12").merge();s.getRange("A12").values=[["结果分层"]];section(s,"A12:H12");write(s,12,0,[["结果","组数","旺店通数量","旺店通金额","惠策本期应收","惠策本期实收","处理","说明"]]);header(s,"A13:H13");write(s,13,0,S.order_bill_results.map(x=>[x.result,x.groups,x.wdt_qty,x.wdt_amount,x.bill_receivable,x.bill_cash,["单号应收金额一致","单号实收金额一致"].includes(x.result)?"匹配剔除":"保留差异",x.result==="仅账单"?"账单侧无期间内发货订单":""]));body(s,`A14:H${13+S.order_bill_results.length}`);status(s,`A14:A${13+S.order_bill_results.length}`);s.getRange("B14:C21").setNumberFormat("#,##0");s.getRange("D14:F21").setNumberFormat("#,##0.00");
 const pctSectionRow=15+S.order_bill_results.length,pctHeaderRow=pctSectionRow+1,pctFirstRow=pctHeaderRow+1,pctTotalRow=pctFirstRow+S.order_bill_results.length;
 s.getRange(`A${pctSectionRow}:H${pctSectionRow}`).merge();s.getRange(`A${pctSectionRow}`).values=[["金额比对百分比统计"]];section(s,`A${pctSectionRow}:H${pctSectionRow}`);
 write(s,pctHeaderRow-1,0,[["结果","旺店通金额","旺店通金额占比","惠策本期应收","应收金额占比","惠策本期实收","实收金额占比"]]);header(s,`A${pctHeaderRow}:G${pctHeaderRow}`);
 write(s,pctFirstRow-1,0,S.order_bill_results.map(x=>[x.result,x.wdt_amount,null,x.bill_receivable,null,x.bill_cash,null]));
 for(let r=pctFirstRow;r<pctTotalRow;r++){s.getRange(`C${r}`).formulas=[[`=IFERROR(B${r}/$B$${pctTotalRow},0)`]];s.getRange(`E${r}`).formulas=[[`=IFERROR(D${r}/$D$${pctTotalRow},0)`]];s.getRange(`G${r}`).formulas=[[`=IFERROR(F${r}/$F$${pctTotalRow},0)`]];}
 write(s,pctTotalRow-1,0,[["合计",null,null,null,null,null,null]]);for(const c of ["B","D","F"])s.getRange(`${c}${pctTotalRow}`).formulas=[[`=SUM(${c}${pctFirstRow}:${c}${pctTotalRow-1})`]];for(const c of ["C","E","G"])s.getRange(`${c}${pctTotalRow}`).formulas=[[`=IFERROR(${String.fromCharCode(c.charCodeAt(0)-1)}${pctTotalRow}/${String.fromCharCode(c.charCodeAt(0)-1)}${pctTotalRow},0)`]];
 body(s,`A${pctFirstRow}:G${pctTotalRow-1}`);status(s,`A${pctFirstRow}:A${pctTotalRow-1}`);section(s,`A${pctTotalRow}:G${pctTotalRow}`);s.getRange(`B${pctFirstRow}:B${pctTotalRow}`).setNumberFormat("#,##0.00");s.getRange(`D${pctFirstRow}:D${pctTotalRow}`).setNumberFormat("#,##0.00");s.getRange(`F${pctFirstRow}:F${pctTotalRow}`).setNumberFormat("#,##0.00");s.getRange(`C${pctFirstRow}:C${pctTotalRow}`).setNumberFormat("0.00%");s.getRange(`E${pctFirstRow}:E${pctTotalRow}`).setNumberFormat("0.00%");s.getRange(`G${pctFirstRow}:G${pctTotalRow}`).setNumberFormat("0.00%");
 [26,18,27,18,28,18,24,32].forEach((w,i)=>s.getRange(`${col(i)}:${col(i)}`).format.columnWidth=w);s.freezePanes.freezeRows(4);
}

// 5 账单OMS汇总
{
 const s=ws(names[4]);title(s,"惠策店铺账单—OMS月结Y001核对汇总","不补物料，直接按月份+店铺/OMS客户核对店铺账单金额；成功、应收、实收三种口径同时保留。","I");
 write(s,3,0,[["惠策账单口径","金额","与OMS差异","相对OMS","","OMS/SAP辅助","金额/数量","说明","判断"]]);header(s,"A4:I4");
 write(s,4,0,[["系统对账成功金额",ctl.huice_bill_success_amount,null,null,"","OMS月结金额",ctl.oms_month_amount,"Y001","优先展示总额比"],["应收口径",ctl.huice_bill_receivable,null,null,"","OMS月结数量",ctl.oms_month_qty,"OMS原生数量","分列"],["实收口径",ctl.huice_bill_cash,null,null,"","SAP一致键金额",ctl.sap_assisted_amount,"OMS-SAP一致键汇总","辅助验证"],["店铺应收映射金额",ctl.mapped_bill_amount,null,null,"","差异绝对值合计",dm.bill_gross_abs_success_diff,"不允许净额抵销","覆盖指标"]]);
 for(let r=5;r<=7;r++){s.getRange(`C${r}`).formulas=[[`=G5-B${r}`]];s.getRange(`D${r}`).formulas=[[`=B${r}/G5`]];}s.getRange("D8").formulas=[["=B8/B6"]];body(s,"A5:I8");status(s,"I5:I8");s.getRange("B5:C8").setNumberFormat("#,##0.00");s.getRange("D5:D8").setNumberFormat("0.00%");s.getRange("G5:G8").setNumberFormat("#,##0.00");
 s.getRange("A10:I10").merge();s.getRange("A10").values=[["月份+店铺结果分层"]];section(s,"A10:I10");write(s,10,0,[["结果","组合数","成功金额","应收金额","实收金额","OMS金额","SAP辅助金额","处理","风险"]]);header(s,"A11:I11");write(s,11,0,S.bill_oms_results.map(x=>[x.result,x.groups,x.bill_success_amount,x.bill_receivable,x.bill_cash,x.oms_amount,x.sap_assisted_amount,["成功金额一致","应收金额一致","实收金额一致","SAP辅助金额一致"].includes(x.result)?"匹配剔除":"保留差异",x.result==="店铺未映射"?"映射缺口":""]));body(s,`A12:I${11+S.bill_oms_results.length}`);status(s,`A12:A${11+S.bill_oms_results.length}`);s.getRange("B12:B19").setNumberFormat("#,##0");s.getRange("C12:G19").setNumberFormat("#,##0.00");
 [25,18,18,18,18,25,18,30,24].forEach((w,i)=>s.getRange(`${col(i)}:${col(i)}`).format.columnWidth=w);s.freezePanes.freezeRows(4);
}

// 7 数量核对
{
 const s=ws(names[6]);title(s,"订单—账单证据—OMS月结数量核对汇总","惠策没有商品数量；数量A取“平台订单号在惠策出现”的旺店通商品数量，数量B取OMS月结Y001数量。","H");
 write(s,3,0,[["总体数量控制","旺店通/订单账单","OMS月结","差异","","共同范围控制","数值","比例/结论"]]);header(s,"A4:H4");
 write(s,4,0,[["全量商品数量",ctl.wdt_qty,ctl.oms_month_qty,null,"","有账单证据订单数量",ctl.billed_wdt_qty,"账单订单范围"],["全量数量比",null,null,null,"","共同月份店铺物料-订单数量",commonWdtQty,"共同键"],["有账单证据数量/OMS",null,null,null,"","共同月份店铺物料-OMS数量",commonOmsQty,"共同键"],["共同组差异绝对值",dm.qty_gross_abs_diff,null,null,"","共同范围净数量差异",null,"不得以净额替代匹配"],["数量一致组合率",null,dm.qty_exact_groups,dm.qty_common_groups,"","共同范围净数量比",null,"优先展示净比例"]]);
 s.getRange("D5").formulas=[["=C5-B5"]];s.getRange("B6").formulas=[["=B5/C5"]];s.getRange("B7").formulas=[["=G5/C5"]];s.getRange("B9").formulas=[["=C9/D9"]];s.getRange("G8").formulas=[["=G7-G6"]];s.getRange("G9").formulas=[["=G6/G7"]];body(s,"A5:H9");status(s,"H5:H9");s.getRange("B5:D5").setNumberFormat("#,##0");s.getRange("B6:B7").setNumberFormat("0.00%");s.getRange("B8:D9").setNumberFormat("#,##0");s.getRange("B9").setNumberFormat("0.00%");s.getRange("G5:G8").setNumberFormat("#,##0");s.getRange("G9").setNumberFormat("0.00%");
 s.getRange("A11:H11").merge();s.getRange("A11").values=[["月份+店铺+物料结果分层"]];section(s,"A11:H11");write(s,11,0,[["结果","组数","账单订单数","订单商品数量","OMS数量","数量差异","处理","说明"]]);header(s,"A12:H12");write(s,12,0,S.qty_results.map(x=>[x.result,x.groups,x.billed_orders,x.order_bill_qty,x.oms_qty,x.qty_difference,x.result==="数量一致"?"匹配剔除":"保留差异",x.result==="数量差异"?"多对一归并后仍存在净差异":""]));body(s,`A13:H${12+S.qty_results.length}`);status(s,`A13:A${12+S.qty_results.length}`);s.getRange("B13:F20").setNumberFormat("#,##0");
 [27,18,18,18,18,27,18,32].forEach((w,i)=>s.getRange(`${col(i)}:${col(i)}`).format.columnWidth=w);s.freezePanes.freezeRows(4);
}

// 9 OMS-SAP
{
 const s=ws(names[8]);title(s,"OMS月结Y001—SAP标准发票（2C）核对汇总","高匹配记录用于双向字段映射：OMS补客户/店铺，SAP补发票号、发票数量及含税金额。","H");
 write(s,3,0,[["双向覆盖控制","数值","说明","","数量/金额控制","SAP","OMS","差异（OMS-SAP）"]]);header(s,"A4:H4");
 write(s,4,0,[["双向一致键",exactMap.keys,"销售单号+物料+单位","","一致键数量",exactMap.sap_qty,exactMap.oms_qty,null],["仅SAP键",sapOnly.keys,"OMS无对应细项","","一致键金额",exactMap.sap_amount,exactMap.oms_amount,null],["仅OMS键",omsOnly.keys,"SAP无对应细项","","仅SAP数量",sapOnly.sap_qty,0,null],["OMS数量覆盖率",null,"一致键OMS数量/OMS月结总体","","仅SAP金额",sapOnly.sap_amount,0,null],["双向键匹配率",null,"一致键/双方键并集","","仅OMS数量",0,omsOnly.oms_qty,null],["SAP数量覆盖率",null,"一致键SAP数量/SAP总体","","仅OMS金额",0,omsOnly.oms_amount,null],["SAP金额覆盖率",null,"一致键SAP金额/SAP总体","","全量数量",ctl.sap_full_qty,ctl.oms_month_qty,null],["完整性控制",null,"双方全量金额分别列示","","全量金额",ctl.sap_full_amount,ctl.oms_month_amount,null]]);
 s.getRange("B8").formulas=[["=G5/G11"]];s.getRange("B9").formulas=[["=B5/SUM(B5:B7)"]];s.getRange("B10").formulas=[["=F5/F11"]];s.getRange("B11").formulas=[["=F6/F12"]];for(let r=5;r<=12;r++)s.getRange(`H${r}`).formulas=[[`=G${r}-F${r}`]];body(s,"A5:H12");s.getRange("B5:B7").setNumberFormat("#,##0");s.getRange("B8:B11").setNumberFormat("0.00%");s.getRange("F5:H12").setNumberFormat("#,##0.00");
 s.getRange("A14:H16").merge();s.getRange("A14").values=[[`字段映射使用：对${exactMap.keys.toLocaleString("zh-CN")}个一致键执行双向补字段；${sapOnly.keys}个仅SAP键和${omsOnly.keys}个仅OMS键完整保留。优先展示OMS数量覆盖率，同时披露双向键匹配率及SAP数量、金额覆盖率。`]];s.getRange("A14:H16").format={fill:C.amber,font:{color:C.amberText,bold:true},wrapText:true,verticalAlignment:"center",borders:{preset:"outside",style:"thin",color:C.line}};
 [26,18,31,3,25,18,18,18].forEach((w,i)=>s.getRange(`${col(i)}:${col(i)}`).format.columnWidth=w);s.freezePanes.freezeRows(4);
}

// 11 月度全流程
{
 const s=ws(names[10]);title(s,"2026年1-6月月度全流程汇总","同一月份横向展示订单、店铺账单三口径、OMS月结、SAP一致键及SAP全量；各列不因总体接近而自动抵销。","M");
 write(s,3,0,[["月份","旺店通商品数量","旺店通商品金额","惠策成功金额","惠策应收","惠策实收","OMS月结数量","OMS月结金额","SAP一致键数量","SAP一致键金额","SAP全量数量","SAP全量金额","SAP全量-OMS金额差"]]);header(s,"A4:M4");
 write(s,4,0,S.monthly_flow.map(x=>[x.month,x.wdt_qty,x.wdt_amount,x.bill_success_amount,x.bill_receivable,x.bill_cash,x.oms_qty,x.oms_amount,x.sap_qty,x.sap_amount,x.sap_full_qty,x.sap_full_amount,null]));for(let r=5;r<=10;r++)s.getRange(`M${r}`).formulas=[[`=L${r}-H${r}`]];body(s,"A5:M10");s.getRange("B5:B10").setNumberFormat("#,##0");s.getRange("C5:F10").setNumberFormat("#,##0.00");s.getRange("G5:G10").setNumberFormat("#,##0");s.getRange("H5:H10").setNumberFormat("#,##0.00");s.getRange("I5:I10").setNumberFormat("#,##0");s.getRange("J5:J10").setNumberFormat("#,##0.00");s.getRange("K5:K10").setNumberFormat("#,##0");s.getRange("L5:M10").setNumberFormat("#,##0.00");
 write(s,12,0,[["合计",null,null,null,null,null,null,null,null,null,null,null,null]]);for(const c of ["B","C","D","E","F","G","H","I","J","K","L","M"])s.getRange(`${c}13`).formulas=[[`=SUM(${c}5:${c}10)`]];section(s,"A13:M13");s.getRange("B13:B13").setNumberFormat("#,##0");s.getRange("C13:F13").setNumberFormat("#,##0.00");s.getRange("G13:G13").setNumberFormat("#,##0");s.getRange("H13:H13").setNumberFormat("#,##0.00");s.getRange("I13:I13").setNumberFormat("#,##0");s.getRange("J13:J13").setNumberFormat("#,##0.00");s.getRange("K13:K13").setNumberFormat("#,##0");s.getRange("L13:M13").setNumberFormat("#,##0.00");
 [14,18,19,20,20,20,18,20,18,20,18,20,22].forEach((w,i)=>s.getRange(`${col(i)}:${col(i)}`).format.columnWidth=w);s.freezePanes.freezeRows(4);
}

detail(names[3],"旺店通订单—惠策明细逐订单核对","订单号匹配及本期应收/实收金额差异；异常优先展示15,000行。",data.orderBill,{platform_order_no:25,wdt_shop:32,huice_shop:34,internal_orders:30,reconcile_ids:30,result:22});
detail(names[5],"惠策店铺账单—OMS月结逐月店铺核对","直接按月份+店铺/OMS客户核对，不使用物料；三种惠策金额口径及SAP辅助金额同时列示。",data.billOms,{huice_shop:34,customer_name:40,result:20},["mapping_status","mapping_source"]);
detail(names[7],"有账单证据订单数量—OMS月结数量明细","按发货月+店铺/OMS客户+SAP物料多对一归并；旺店通数量与OMS月结数量分别列示。",data.qty,{wdt_shop:34,customer_name:40,material_code:17,result:22},["mapping_status"]);
detail(names[9],"OMS月结—SAP双向字段映射明细","一致键同时携带OMS客户/月份和SAP发票字段；仅SAP和仅OMS键完整保留，不强行推定缺失字段。",data.omsSap,{oms_sales_no:24,sap_invoice_nos:26,customer_name:40,mapping_result:22,source_result:20});
detail(names[11],"惠策店铺—OMS客户映射","展示惠策店铺与OMS客户的核对映射结果。",data.shop,{huice_shop:34,customer_name:40},["mapping_status","mapping_source"]);

// 13 索引
{
 const s=ws(names[12]);title(s,"完整明细及代码索引","工作簿嵌入汇总和审阅明细；大体量完整结果保留CSV。","F");
 write(s,3,0,[["核对环节","完整行数","工作簿行数","完整文件","页面","说明"]]);header(s,"A4:F4");
 const idx=[
  ["订单—账单",S.detail_rows.order_bill_recon,data.orderBill.rows.length,path.join(dataDir,"order_bill_recon.csv"),"4.订单-账单明细","异常优先"],
  ["账单—OMS月结",S.detail_rows.bill_oms_month_recon,data.billOms.rows.length,path.join(dataDir,"bill_oms_month_recon.csv"),"6.账单-OMS月结明细","全部嵌入"],
  ["订单账单—OMS数量",S.detail_rows.order_bill_oms_qty_recon,data.qty.rows.length,path.join(dataDir,"order_bill_oms_qty_recon.csv"),"8.数量核对明细","全部嵌入"],
  ["OMS—SAP字段映射",S.detail_rows.oms_sap_field_map,data.omsSap.rows.length,path.join(dataDir,"oms_sap_field_map.csv"),"10.OMS-SAP字段映射","全部嵌入"],
  ["店铺客户映射",S.detail_rows.huice_shop_map,data.shop.rows.length,path.join(dataDir,"huice_shop_map.csv"),"12.店铺客户映射","全部嵌入"],
 ];write(s,4,0,idx);body(s,"A5:F9");s.getRange("B5:C9").setNumberFormat("#,##0");[28,18,18,86,28,24].forEach((w,i)=>s.getRange(`${col(i)}:${col(i)}`).format.columnWidth=w);s.getRange("D5:D9").format.wrapText=true;s.freezePanes.freezeRows(4);
}

await fs.mkdir(outputDir,{recursive:true});const previews=path.join(outputDir,"_qa_previews");await fs.mkdir(previews,{recursive:true});
console.log("OVERVIEW\n"+(await wb.inspect({kind:"table",range:"1.全局口径与总览!A1:K16",include:"values,formulas",tableMaxRows:20,tableMaxCols:12,maxChars:12000})).ndjson);
console.log("QTY\n"+(await wb.inspect({kind:"table",range:"7.数量核对汇总!A1:H18",include:"values,formulas",tableMaxRows:22,tableMaxCols:10,maxChars:10000})).ndjson);
console.log("LOW_CONFIDENCE\n"+(await wb.inspect({kind:"match",searchTerm:"低置信",options:{useRegex:false,maxResults:100},summary:"visible low-confidence labels",maxChars:3000})).ndjson);
console.log("ERRORS\n"+(await wb.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",options:{useRegex:true,maxResults:300},summary:"formula errors",maxChars:5000})).ndjson);
for(const name of names){const s=ws(name),used=s.getUsedRange(true),maxCols=Math.min(used.columnCount||8,name.includes("明细")||name.includes("映射")?10:14),maxRows=name.includes("明细")||name.includes("映射")?22:30;const blob=await wb.render({sheetName:name,range:`A1:${col(maxCols-1)}${maxRows}`,scale:1.15,format:"png"});await fs.writeFile(path.join(previews,`${name}.png`),new Uint8Array(await blob.arrayBuffer()));}
const out=await SpreadsheetFile.exportXlsx(wb);await out.save(outputFile);console.log(JSON.stringify({outputFile,sheets:names,previews},null,2));
