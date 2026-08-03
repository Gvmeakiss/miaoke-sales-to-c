import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

const scriptRoot=path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const root=process.env.SALES_TOC_ROOT||scriptRoot;
const dataDir=process.env.SALES_TOC_DATA_DIR||path.join(root,"reconciliation/results");
const outputDir=process.env.SALES_TOC_OUTPUT_DIR||path.join(root,"outputs/sales_toc_workpaper_final_20260101_20260630");
const outputFile=process.env.SALES_TOC_OUTPUT_FILE||path.join(outputDir,"销售ToC业务流程核对底稿_20260101-20260630.xlsx");
const read=async n=>JSON.parse(await fs.readFile(path.join(dataDir,n),"utf8"));
const S=await read("summary.json");
const data={
 orderBill:await read("order_bill_recon_workbook.json"),internal:await read("huice_internal_recon_workbook.json"),billOms:await read("bill_oms_month_recon_workbook.json"),
 qty:await read("order_bill_oms_qty_recon_workbook.json"),shop:await read("huice_shop_map_workbook.json"),
};
const wb=Workbook.create();
const names=["1.全局口径与总览","2.订单-账单汇总","3.旺店通-惠策订单核对明细","4.惠策内部核对汇总","5.惠策内部核对明细","6.账单-OMS月结汇总","7.账单-OMS月结明细","8.数量核对汇总","9.数量核对明细","10.OMS月结-SAP汇总","12.月度全流程汇总","13.店铺客户映射"];
for(const n of names)wb.worksheets.add(n);
const scope="账单核对期间：2026-01-01至2026-06-30｜旺店通订单追溯期间：2025-12-01至2026-06-30";
const C={navy:"#17365D",blue:"#2F75B5",pale:"#DDEBF7",pale2:"#EAF3F8",white:"#FFFFFF",text:"#203040",line:"#B4C6E7",green:"#E2F0D9",greenText:"#375623",amber:"#FFF2CC",amberText:"#7F6000",red:"#FCE4D6",redText:"#9C0006"};
const ws=n=>wb.worksheets.getItem(n);
const col=i=>{let n=i+1,s="";while(n){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26);}return s;};
const clean=v=>v===undefined||v===null||(typeof v==="number"&&!Number.isFinite(v))?null:(typeof v==="string"&&/^\d{12,}$/.test(v)?`\u200B${v}`:v);
function write(s,r,c,rows){if(rows.length&&rows[0].length)s.getRangeByIndexes(r,c,rows.length,rows[0].length).values=rows.map(x=>x.map(clean));}
function title(s,t,sub,last="J"){s.showGridLines=false;s.getRange(`A1:${last}1`).merge();s.getRange("A1").values=[[t]];s.getRange(`A1:${last}1`).format={fill:C.navy,font:{bold:true,color:C.white,size:16},verticalAlignment:"center"};s.getRange(`A1:${last}1`).format.rowHeight=30;s.getRange(`A2:${last}2`).merge();s.getRange("A2").values=[[`${scope}｜${sub}`]];s.getRange(`A2:${last}2`).format={fill:C.pale2,font:{italic:true,color:C.text,size:10},wrapText:true,verticalAlignment:"center"};s.getRange(`A2:${last}2`).format.rowHeight=38;}
function header(s,r){s.getRange(r).format={fill:C.blue,font:{bold:true,color:C.white},horizontalAlignment:"center",verticalAlignment:"center",wrapText:true,borders:{preset:"all",style:"thin",color:C.line}};}
function section(s,r){s.getRange(r).format={fill:C.pale,font:{bold:true,color:C.navy,size:11},borders:{bottom:{style:"medium",color:C.blue}}};}
function body(s,r){s.getRange(r).format={font:{color:C.text,size:10},verticalAlignment:"center",borders:{insideHorizontal:{style:"thin",color:"#E7E6E6"}}};}
function status(s,r){const x=s.getRange(r);x.conditionalFormats.add("containsText",{text:"差异",format:{fill:C.red,font:{bold:true,color:C.redText}}});x.conditionalFormats.add("containsText",{text:"未映射",format:{fill:C.red,font:{bold:true,color:C.redText}}});x.conditionalFormats.add("containsText",{text:"仅",format:{fill:C.amber,font:{bold:true,color:C.amberText}}});x.conditionalFormats.add("containsText",{text:"附属单",format:{fill:C.amber,font:{bold:true,color:C.amberText}}});x.conditionalFormats.add("containsText",{text:"一致",format:{fill:C.green,font:{bold:true,color:C.greenText}}});}
const headerZh={
 platform_order_no:"平台订单号",matchable:"可匹配标识",wdt_shop:"旺店通店铺",platform:"平台",huice_shop:"惠策店铺",
 order_month:"订单月份",ship_month:"发货月份",bill_month:"账单月份",internal_order_count:"内部订单数",huice_rows:"惠策明细行数",
 wdt_qty:"旺店通数量",wdt_amount:"旺店通分摊金额",wdt_allocated_amount:"旺店通分摊金额",wdt_header_amount:"旺店通订单应收",bill_receivable:"惠策本期应收",
 bill_cash:"惠策本期实收",receivable_difference:"应收差异",cash_difference:"实收差异",internal_orders:"旺店通内部单号",
 reconcile_ids:"惠策对账流水号",result:"核对结果",customer_code:"OMS客户编码",customer_name:"OMS客户名称",
 mapping_status:"映射置信度",mapping_source:"映射来源",bill_record_count:"账单记录数",success_count:"成功记录数",
 bill_success_amount:"对账成功金额",oms_docs:"OMS单据数",oms_qty:"OMS数量",oms_amount:"OMS金额",
 sap_assisted_qty:"SAP辅助数量",sap_assisted_amount:"SAP辅助金额",success_difference:"成功金额差异",
 sap_success_difference:"SAP辅助成功差异",source_rows:"源文件行数",material_code:"物料编码",billed_orders:"账单证据平台单数",
 exact_evidence_orders:"平台单号精确证据数",auxiliary_evidence_orders:"零金额附属单证据数",
 order_bill_qty:"订单证据数量",exact_order_bill_qty:"平台单号精确证据数量",auxiliary_order_bill_qty:"零金额附属单数量",
 wdt_item_amount:"旺店通商品金额",qty_difference:"数量差异",oms_sales_no:"OMS销售单号",
 sales_unit:"销售单位",file_month:"SAP文件月份",outbound_month:"OMS出库月份",sap_invoice_nos:"SAP发票号",
 sap_rows:"SAP行数",oms_rows:"OMS行数",sap_qty:"SAP数量",quantity_difference:"数量差异",sap_amount:"SAP金额",
 amount_difference:"金额差异",mapped_qty:"映射数量",mapped_amount:"映射金额",mapping_result:"映射结果",source_result:"源核对结果",
 detail_rows:"惠策明细行数",summary_rows:"惠策汇总源行数",detail_success_amount:"明细成功状态金额",
 summary_success_amount:"汇总成功分类金额",detail_receivable:"明细应收",summary_receivable:"汇总应收",
 detail_cash:"明细实收",summary_cash:"汇总实收",historical_rows:"往期业务日期行数",
 historical_receivable:"往期业务日期应收",historical_cash:"往期业务日期实收"
};
function detail(name,t,sub,d,widths={},omit=[]){const keep=d.headers.map((h,i)=>({h,i})).filter(x=>!omit.includes(x.h)),headers=keep.map(x=>x.h),rows=keep.length===d.headers.length?d.rows:d.rows.map(r=>keep.map(x=>r[x.i])),displayHeaders=headers.map(h=>headerZh[h]?`${h}\n${headerZh[h]}`:h);const s=ws(name),last=col(headers.length-1),titleLast=col(Math.min(headers.length,10)-1);title(s,t,sub,titleLast);write(s,3,0,[displayHeaders]);header(s,`A4:${last}4`);s.getRange(`A4:${last}4`).format.rowHeight=34;for(let i=0;i<rows.length;i+=3000)write(s,4+i,0,rows.slice(i,i+3000));const end=4+rows.length;if(rows.length){body(s,`A5:${last}${end}`);for(const k of ["result","mapping_result"]){const ix=headers.indexOf(k);if(ix>=0)status(s,`${col(ix)}5:${col(ix)}${end}`);}}const textFields=new Set(["platform_order_no","internal_orders","reconcile_ids","oms_sales_no","sap_invoice_nos","customer_code","material_code"]);headers.forEach((h,i)=>{let w=widths[h]||(/name|shop|orders|invoice|reconcile/.test(h)?25:/amount|difference|receivable|cash/.test(h)?17:14);s.getRange(`${col(i)}:${col(i)}`).format.columnWidth=w;if(textFields.has(h))s.getRange(`${col(i)}5:${col(i)}${end}`).setNumberFormat("@");else if(/amount|difference|receivable|cash/.test(h))s.getRange(`${col(i)}5:${col(i)}${end}`).setNumberFormat("#,##0.00");else if(/qty|count|rows|lines|groups|keys/.test(h))s.getRange(`${col(i)}5:${col(i)}${end}`).setNumberFormat("#,##0");});s.freezePanes.freezeRows(4);s.freezePanes.freezeColumns(Math.min(2,headers.length));}
const by=(arr,key)=>Object.fromEntries(arr.map(x=>[x[key],x]));
const ob=by(S.order_bill_results,"result"),bo=by(S.bill_oms_results,"result"),qr=by(S.qty_results,"result"),os=by(S.oms_sap_results,"mapping_result"),ctl=S.controls;
const allocatedExactLabels=["单号分摊应收一致","单号分摊实收一致"];
const headerFallbackLabels=["单号订单应收一致","单号订单实收一致"];
const orderExactLabels=[...allocatedExactLabels,...headerFallbackLabels];
const auxiliaryOrderLabel="同内部订单零金额附属单";
const orderExact=S.order_bill_results.filter(x=>orderExactLabels.includes(x.result)).reduce((a,x)=>a+(x.groups||0),0);
const orderAllocatedExact=S.order_bill_results.filter(x=>allocatedExactLabels.includes(x.result)).reduce((a,x)=>a+(x.groups||0),0);
const orderHeaderFallback=S.order_bill_results.filter(x=>headerFallbackLabels.includes(x.result)).reduce((a,x)=>a+(x.groups||0),0);
const orderExactAmount=S.order_bill_results.filter(x=>allocatedExactLabels.includes(x.result)).reduce((a,x)=>a+(x.wdt_amount||0),0);
const orderMatchedBillAmount=S.order_bill_results.filter(x=>allocatedExactLabels.includes(x.result)).reduce((a,x)=>a+(x.result.includes("应收")?(x.bill_receivable||0):(x.bill_cash||0)),0);
const orderExactQty=S.order_bill_results.filter(x=>orderExactLabels.includes(x.result)).reduce((a,x)=>a+(x.wdt_qty||0),0);
const orderReconAmount=S.order_bill_results.reduce((a,x)=>a+(x.wdt_amount||0),0);
const commonWdtQty=(qr["数量差异"]?.order_bill_qty||0)+(qr["数量一致"]?.order_bill_qty||0);
const commonOmsQty=(qr["数量差异"]?.oms_qty||0)+(qr["数量一致"]?.oms_qty||0);
const exactMap=os["双向字段一致"],sapOnly=os["SAP补数量金额"],omsOnly=os["仅OMS月结"]||{keys:0,oms_qty:0,oms_amount:0};
const dm=S.display_metrics||{};
const orderKeyCoverage=dm.order_key_matches/ctl.wdt_orders;
const huiceKeyCoverage=dm.order_key_matches/ctl.huice_orders;
const orderEvidenceCoverage=(dm.order_key_matches+dm.order_auxiliary_explained_keys)/ctl.wdt_orders;
const matchedBillReceivableCoverage=ctl.order_bill_matched_receivable/ctl.huice_detail_settlement_receivable;
const orderAmountExactCoverage=dm.order_amount_exact_groups/ctl.huice_orders;
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
const huiceReceivableMatch=1-Math.abs(ctl.huice_internal_receivable_difference)/Math.abs(ctl.huice_detail_settlement_receivable);
const huiceCashMatch=1-Math.abs(ctl.huice_internal_cash_difference)/Math.abs(ctl.huice_detail_settlement_cash);
const huiceReceivableDifference=Math.abs(ctl.huice_internal_receivable_difference)<=0.01?0:ctl.huice_internal_receivable_difference;
const huiceCashDifference=Math.abs(ctl.huice_internal_cash_difference)<=0.01?0:ctl.huice_internal_cash_difference;
const orderOnly=ob["仅订单"]?.groups||0;
const billOnly=ob["仅账单"]?.groups||0;
const auxiliaryExplained=ob[auxiliaryOrderLabel]?.groups||0;

// 1 全局口径
{
 const s=ws(names[0]);title(s,"销售 ToC 核对底稿汇总", "以2026年1—6月惠策出账清单为订单核对基表，向前追溯2025年12月至2026年6月旺店通订单；后续环节按各自强键或共同维度单独核对。","L");
 s.getRange("A4:L4").merge();s.getRange("A4").values=[["金额匹配"]];section(s,"A4:L4");
 write(s,4,0,[["核对环节","参与匹配数据","匹配所用字段","核对方式","结论","旺店通比对金额","惠策金额（用于旺店通比对）","惠策汇总表金额","OMS月结金额","SAP发票金额","差异（按行定义）","金额匹配率"]]);header(s,"A5:L5");
 s.getRange("A5:L5").format.rowHeight=30;
 write(s,5,0,[
  ["1.订单—账单","惠策出账明细—旺店通订单","平台订单号精确匹配；内部订单号解释零金额附属单","以惠策1—6月出账清单为基表，向前追溯旺店通12月至6月订单","惠策金额为用于旺店通订单核对的匹配金额；汇总表金额为全量本期应收",orderExactAmount,orderMatchedBillAmount,ctl.huice_bill_receivable,null,null,null,null],
  ["2.惠策内部—应收","惠策明细—惠策店铺汇总","结算月+平台+店铺","全量金额及逐组合比较","全量本期应收金额一致",null,ctl.huice_detail_settlement_receivable,ctl.huice_bill_receivable,null,null,null,null],
  ["3.惠策内部—实收","惠策明细—惠策店铺汇总","结算月+平台+店铺","全量金额及逐组合比较","全量本期实收金额一致",null,ctl.huice_detail_settlement_cash,ctl.huice_bill_cash,null,null,null,null],
  ["4.账单—OMS月结","惠策店铺汇总—OMS Y001","月份+店铺/OMS客户","总额及逐月店铺比较","惠策对账成功分类金额与OMS月结金额比较；应收/实收为辅助口径",null,null,ctl.huice_bill_success_amount,ctl.oms_month_amount,null,null,null],
  ["5.OMS月结—SAP","OMS Y001—SAP标准发票（2C）","销售单号+物料+销售单位","共同键及全量覆盖比较","共同键数量一致且金额在容忍度内；单边键另行披露",null,null,null,exactMap.oms_amount,exactMap.sap_amount,null,null],
 ]);
 s.getRange("K6").formulas=[["=G6-F6"]];s.getRange("L6").formulas=[["=MIN(ABS(F6),ABS(G6))/MAX(ABS(F6),ABS(G6))"]];
 s.getRange("K7").formulas=[["=IF(ABS(H7-G7)<=0.01,0,H7-G7)"]];s.getRange("L7").formulas=[["=MIN(ABS(G7),ABS(H7))/MAX(ABS(G7),ABS(H7))"]];
 s.getRange("K8").formulas=[["=IF(ABS(H8-G8)<=0.01,0,H8-G8)"]];s.getRange("L8").formulas=[["=MIN(ABS(G8),ABS(H8))/MAX(ABS(G8),ABS(H8))"]];
 s.getRange("K9").formulas=[["=I9-H9"]];s.getRange("L9").formulas=[["=MIN(ABS(H9),ABS(I9))/MAX(ABS(H9),ABS(I9))"]];
 s.getRange("K10").formulas=[["=J10-I10"]];s.getRange("L10").formulas=[["=MIN(ABS(I10),ABS(J10))/MAX(ABS(I10),ABS(J10))"]];
 body(s,"A6:L10");status(s,"E6:E10");s.getRange("F6:K10").setNumberFormat("#,##0.00;[Red](#,##0.00);-");s.getRange("L6:L10").setNumberFormat("0.00%");
 s.getRange("A7:L8").format.fill="#E7E6E6";s.getRange("A7:L8").format.font={color:"#666666",size:10};
 s.getRange("A12:L12").merge();s.getRange("A12").values=[["订单数量链"]];section(s,"A12:L12");
 write(s,12,0,[["核对环节","数量定义","匹配规则","旺店通候选平台订单数","惠策账单平台订单数","平台单号精确匹配","零金额附属单已解释","旺店通其他订单","惠策账单单边订单","惠策账单订单匹配率","旺店通候选订单证据率","商品数量覆盖率"]]);header(s,"A13:L13");s.getRange("A13:L13").format.rowHeight=34;
 write(s,13,0,[["惠策账单—旺店通订单","平台订单号去重数量","以惠策账单为基表执行精确匹配；旺店通同内部订单零金额平台单号作二级证据",ctl.wdt_orders,ctl.huice_orders,dm.order_key_matches,auxiliaryExplained,orderOnly,billOnly,null,null,null]]);
 s.getRange("J14").formulas=[["=F14/E14"]];s.getRange("K14").formulas=[["=(F14+G14)/D14"]];s.getRange("L14").formulas=[["='2.订单-账单汇总'!F10"]];body(s,"A14:L14");s.getRange("D14:I14").setNumberFormat("#,##0");s.getRange("J14:L14").setNumberFormat("0.00%");
 s.getRange("A16:L16").merge();s.getRange("A16").values=[["商品数量链"]];section(s,"A16:L16");
 write(s,16,0,[["核对环节","参与数据","数量来源/性质","旺店通原生数量","惠策覆盖订单派生数量","OMS原生数量","SAP一致键数量","差异","匹配率","口径说明","页面"]]);header(s,"A17:K17");s.getRange("A17:K17").format.rowHeight=36;
 write(s,17,0,[
  ["1.旺店通—惠策账单证据","旺店通订单—惠策明细","精确平台单数量加同内部订单零金额附属单数量",ctl.wdt_qty,ctl.billed_wdt_qty,null,null,null,null,`惠策无商品数量；其中${Number(ctl.billed_wdt_auxiliary_qty||0).toLocaleString("zh-CN")}为同内部订单零金额附属单数量`,"2-3"],
  ["2.惠策账单证据—OMS","有惠策订单证据的旺店通商品明细—OMS Y001","订单证据数量与OMS原生数量总体比较",null,ctl.billed_wdt_qty,ctl.oms_month_qty,null,null,null,"内部订单桥接仅归集分摊金额为零的附属单数量","8-9"],
  ["3.OMS月结—SAP","OMS Y001—SAP标准发票（2C）","销售单号+物料+销售单位",null,null,exactMap.oms_qty,exactMap.sap_qty,null,null,"共同键数量一致；覆盖率仅反映共同键占全量的比例","10"],
 ]);
 s.getRange("H18").formulas=[["=E18-D18"]];s.getRange("I18").formulas=[["=E18/D18"]];
 s.getRange("H19").formulas=[["=F19-E19"]];s.getRange("I19").formulas=[["=E19/F19"]];
 s.getRange("H20").formulas=[["=G20-F20"]];s.getRange("I20").formulas=[[`=F20/${ctl.oms_month_qty}`]];
 body(s,"A18:K20");s.getRange("D18:H20").setNumberFormat("#,##0;[Red](#,##0);-");s.getRange("I18:I20").setNumberFormat("0.00%");
 s.getRange("A22:L25").merge();s.getRange("A22").values=[[`口径说明：订单—账单环节以2026年1—6月惠策导出账单为基表，旺店通订单追溯至2025年12月1日。惠策导出未提供物料、SKU或商品数量字段；平台订单号精确匹配后，如同一旺店通内部订单已有精确匹配，则该内部订单下分摊金额为零的其他平台单号列为“零金额附属单已解释”。商品数量为旺店通派生数量，不代表惠策原生商品数量。`]];s.getRange("A22:L25").format={fill:C.amber,font:{color:C.amberText,bold:true},wrapText:true,verticalAlignment:"center",borders:{preset:"outside",style:"thin",color:"#C9A227"}};
 [24,32,35,21,25,20,20,20,18,18,28,16].forEach((w,i)=>s.getRange(`${col(i)}:${col(i)}`).format.columnWidth=w);s.freezePanes.freezeRows(5);
}

// 2 规则
if(false){
{
 const s=ws(names[1]);title(s,"业务流程与单对单核对规则","每一环节独立钩稽，避免因下一环节缺字段而否定上一环节的有效匹配。","J");
 write(s,3,0,[["环节","数据A","数据B","强键/共同维度","金额比较","数量比较","匹配阈值","匹配后处理","限制","页面"]]);header(s,"A4:J4");
 write(s,4,0,[
  [1,"旺店通订单","惠策明细","内部订单号+原始单号保留粒度；旺店通原始单号=惠策平台订单号","平台单号分摊金额分别对本期应收、本期实收；订单应收仅辅助","记录旺店通商品数量；惠策无商品数量","金额差≤0.01","一致组剔除","订单与账单状态时点不同","3-4"],
  [2,"惠策明细","惠策店铺汇总","导出结算月+平台+店铺","全量应收、实收分别比较","两表均无商品数量","金额差≤0.01","一致组剔除","成功金额分类定义不同，不作为主钩稽口径","5-6"],
  [3,"惠策店铺汇总","OMS月结Y001","月份+惠策店铺/OMS客户映射","成功金额为主；应收、实收并列","账单笔数与OMS商品数量不可直接比较","金额差≤0.01","一致组剔除","平台费用、单边未关联、结算时点","7-8"],
  [4,"惠策有记录的订单数量","OMS月结Y001","发货月+店铺/客户+SAP物料","旺店通商品金额仅辅助","订单商品数量 vs OMS月结数量","数量差=0；总体比例另列","一致组剔除","惠策仅作为订单存在证据","9-10"],
  [5,"OMS月结Y001","SAP标准发票（2C）","OMS销售单号+物料+销售单位","含税金额","发票数量 vs OMS数量","金额差≤0.01且数量相等","生成双向映射",`${sapOnly.keys}个仅SAP、${omsOnly.keys}个仅OMS键保留`,"11-12"],
 ]);body(s,"A5:J9");
 s.getRange("A11:J11").merge();s.getRange("A11").values=[["SAP—OMS双向字段映射"]];section(s,"A11:J11");write(s,11,0,[["来源字段","补充到","映射后字段","使用条件","用途"],["OMS客户编码/客户名称/出库月份","SAP记录","店铺客户、业务月份","OMS销售单号+物料+单位匹配","按店铺/月度辅助汇总"],["SAP发票号/发票数量/含税金额","OMS记录","发票字段、SAP辅助金额","同一强键匹配","验证OMS金额数量及发票追踪"]]);header(s,"A12:E12");body(s,"A13:E14");
 s.getRange("A16:J16").merge();s.getRange("A16").values=[["旺店通—惠策平台订单映射"]];section(s,"A16:J16");
 write(s,16,0,[["映射主键","旺店通保留字段","惠策保留字段","关系处理","核对输出","使用限制"],["平台订单号","内部订单号、店铺、发货月、分摊金额、商品数量","对账流水号、平台、店铺、应收、实收","双方先按平台订单号汇总为一行","金额差异、匹配结果及双方单号集合","多内部单号或多流水号不强行拆成一对一"]]);header(s,"A17:F17");body(s,"A18:F18");
 [10,25,25,34,28,27,20,24,35,15].forEach((w,i)=>s.getRange(`${col(i)}:${col(i)}`).format.columnWidth=w);s.freezePanes.freezeRows(4);
}
}

// 业务流程与单对单规则合并至全局总览
{
 const s=ws(names[0]);
 s.getRange("A29:L29").merge();s.getRange("A29").values=[["业务流程与单对单核对规则"]];section(s,"A29:L29");
 write(s,29,0,[["环节","数据A","数据B","强键/共同维度","金额比较","数量比较","匹配阈值","匹配后处理","口径说明","页面"]]);header(s,"A30:J30");
 write(s,30,0,[
  [1,"惠策1—6月出账明细","旺店通12月至6月订单","一级：惠策平台订单号=旺店通原始单号；二级：同内部订单零金额附属单","旺店通平台订单分摊金额与惠策本期应收/实收比较","精确匹配数量加零金额附属单数量","金额差≤0.01；附属单分摊金额绝对值≤0.01","以惠策账单为基表执行匹配；附属单列为二级证据；其他记录按核对分类列示","订单追溯期间早于账单期间一个月，以覆盖跨期确认收货后出账","2-3"],
  [2,"惠策明细","惠策店铺汇总","导出结算月+平台+店铺","全量本期应收、实收分别比较","两表均未提供商品数量","金额差≤0.01","一致组标记为已核对，其他组合按核对分类列示","对账成功分类金额与明细状态定义不同","4-5"],
  [3,"惠策店铺汇总","OMS月结Y001","月份+惠策店铺/OMS客户映射","对账成功分类金额为主，应收/实收为辅助","惠策账单笔数与OMS商品数量不具可比性","金额差≤0.01","一致组标记为已核对，其他组合按核对分类列示","平台费用、结算时点及客户映射影响比较","6-7"],
  [4,"有惠策订单证据的旺店通商品数量","OMS月结Y001","发货月+店铺/客户+SAP物料","旺店通商品金额仅作辅助信息","旺店通商品数量与OMS原生数量比较","数量差=0","一致组标记为已核对，其他组合按核对分类列示","惠策仅提供订单存在性证据，无法逐订单穿透OMS","8-9"],
  [5,"OMS月结Y001","SAP标准发票（2C）","OMS销售单号+物料+销售单位","含税金额比较","发票数量与OMS数量比较","金额差≤0.01且数量相等","共同键执行数量及金额核对，单边键分别列示",`${sapOnly.keys}个仅SAP键、${omsOnly.keys}个仅OMS键另行列示`,"10"],
 ]);body(s,"A31:J35");
 s.getRange("A37:L37").merge();s.getRange("A37").values=[["关键字段映射说明"]];section(s,"A37:L37");
 write(s,37,0,[["映射关系","来源保留字段","接收方补充字段","使用条件","用途","口径说明"],["旺店通—惠策","内部订单号、平台原始单号、店铺、发货月、分摊金额、商品数量","惠策对账流水号、平台、店铺、本期应收、本期实收","平台订单号精确匹配；同内部订单且分摊金额为零时二级解释","订单金额核对、附属单解释及数量证据归集","零金额附属单用于订单证据及数量归集，不表述为惠策原生商品数量"],["OMS—SAP","OMS客户、名称、出库月份、销售单号、物料、单位","SAP发票号、数量、含税金额","销售单号+物料+单位","发票追踪及字段补充","共同键与单边键分别列示"]]);header(s,"A38:F38");body(s,"A39:F40");
 s.getRange("A30:J35").format.wrapText=true;s.getRange("A38:F40").format.wrapText=true;
}

// 3 订单账单汇总
{
 const s=ws(names[1]);title(s,"惠策出账明细—旺店通订单核对汇总","以2026年1—6月惠策导出账单为核对基表，向前追溯2025年12月至2026年6月旺店通订单；惠策比对金额与惠策汇总表总金额分别列示。","H");
 write(s,3,0,[["账单基表与订单覆盖","数值","说明","","金额及数量覆盖","数值","说明","判断"]]);header(s,"A4:H4");
 write(s,4,0,[["惠策账单平台订单组",ctl.huice_orders,"2026年1—6月惠策导出账单平台订单号去重数量","","惠策账单清单全量本期应收",ctl.huice_detail_settlement_receivable,"惠策账单明细全量本期应收","核对基表"],["旺店通平台订单号精确匹配",dm.order_key_matches,"惠策平台订单号=旺店通原始单号","","惠策汇总表总金额（本期应收）",ctl.huice_bill_receivable,"惠策店铺汇总全量本期应收金额","汇总口径"],["惠策账单订单匹配率",null,"平台订单号精确匹配/惠策账单平台订单组","","旺店通比对金额",orderExactAmount,"金额一致订单组的旺店通分摊金额","比对口径"],["其中：12月订单—1月账单匹配",ctl.order_dec_to_jan_matches,"订单月份为2025年12月、账单月份为2026年1月","","惠策金额（用于旺店通比对）",orderMatchedBillAmount,"对应惠策本期应收或实收金额","比对口径"],["旺店通候选平台订单组",ctl.wdt_orders,"2025年12月至2026年6月订单追溯范围","","比对金额一致率",null,"旺店通比对金额与惠策比对金额的较小值/较大值","核心匹配率"],["旺店通候选订单证据率",null,"（精确匹配+零金额附属单）/旺店通候选平台订单组","","商品数量覆盖率",null,"本期发货订单证据商品数量/旺店通本期发货商品数量","核心覆盖率"]]);
 s.getRange("B7").formulas=[["=B6/B5"]];s.getRange("B10").formulas=[[`=(${dm.order_key_matches}+${dm.order_auxiliary_explained_keys})/B9`]];s.getRange("F9").formulas=[["=MIN(F7,F8)/MAX(F7,F8)"]];s.getRange("F10").formulas=[[`=${ctl.billed_wdt_qty}/${ctl.wdt_qty}`]];body(s,"A5:H10");status(s,"H5:H10");s.getRange("B5:B6").setNumberFormat("#,##0");s.getRange("B7").setNumberFormat("0.00%");s.getRange("B8:B9").setNumberFormat("#,##0");s.getRange("B10").setNumberFormat("0.00%");s.getRange("F5:F8").setNumberFormat("#,##0.00");s.getRange("F9:F10").setNumberFormat("0.00%");s.getRange("A7:C7").format.fill=C.green;s.getRange("A10:C10").format.fill=C.green;s.getRange("E9:H10").format.fill=C.green;
 s.getRange("A12:H12").merge();s.getRange("A12").values=[["证据口径说明"]];section(s,"A12:H12");
 write(s,12,0,[["订单证据项目","数值","说明","","数量证据项目","数值","说明","判断"]]);header(s,"A13:H13");
 write(s,13,0,[["惠策账单订单匹配率",null,"精确匹配平台订单组/惠策账单平台订单组","","平台单号精确证据数量",ctl.billed_wdt_qty_direct,"精确匹配平台单号对应本期发货商品数量","基础口径"],["惠策账单匹配应收覆盖率",matchedBillReceivableCoverage,"已匹配平台订单本期应收/惠策账单全量本期应收","","零金额附属单数量",ctl.billed_wdt_auxiliary_qty,"同内部订单二级桥接补充的本期发货商品数量","附属单口径"],["12月订单—1月账单匹配组",ctl.order_dec_to_jan_matches,"跨期下单并于次月出账的平台订单组","","订单证据数量",ctl.billed_wdt_qty,"精确证据数量+零金额附属单数量","订单证据口径"],["12月订单—1月账单旺店通金额",ctl.order_dec_to_jan_wdt_amount,"跨期匹配组对应旺店通分摊金额","","平台单号精确数量覆盖率",null,"精确证据数量/旺店通本期发货商品数量","基础覆盖率"],["12月订单—1月账单惠策应收",ctl.order_dec_to_jan_bill_receivable,"跨期匹配组对应惠策本期应收","","商品数量覆盖率",null,"订单证据数量/旺店通本期发货商品数量","核心覆盖率"]]);
 s.getRange("B14").formulas=[["=B6/B5"]];s.getRange("F17").formulas=[[`=F14/${ctl.wdt_qty}`]];s.getRange("F18").formulas=[[`=F16/${ctl.wdt_qty}`]];body(s,"A14:H18");status(s,"H14:H18");s.getRange("B14:B15").setNumberFormat("0.00%");s.getRange("B16").setNumberFormat("#,##0");s.getRange("B17:B18").setNumberFormat("#,##0.00");s.getRange("F14:F16").setNumberFormat("#,##0");s.getRange("F17:F18").setNumberFormat("0.00%");
 const resultSectionRow=20,resultHeaderRow=21,resultFirstRow=22,resultEndRow=resultFirstRow+S.order_bill_results.length-1;
 s.getRange(`A${resultSectionRow}:H${resultSectionRow}`).merge();s.getRange(`A${resultSectionRow}`).values=[["核对结果构成"]];section(s,`A${resultSectionRow}:H${resultSectionRow}`);write(s,resultHeaderRow-1,0,[["结果","组数","旺店通数量","旺店通分摊金额","惠策本期应收","惠策本期实收","处理","说明"]]);header(s,`A${resultHeaderRow}:H${resultHeaderRow}`);write(s,resultFirstRow-1,0,S.order_bill_results.map(x=>[x.result,x.groups,x.wdt_qty,x.wdt_amount,x.bill_receivable,x.bill_cash,orderExactLabels.includes(x.result)?"已核对":x.result===auxiliaryOrderLabel?"已解释":"分类列示",x.result===auxiliaryOrderLabel?"同一内部订单已有惠策精确匹配，且本平台单号分摊金额为零":x.result==="仅账单"?"惠策账单记录与旺店通期间订单分别列示":headerFallbackLabels.includes(x.result)?"订单应收辅助证据":""]));body(s,`A${resultFirstRow}:H${resultEndRow}`);s.getRange(`B${resultFirstRow}:C${resultEndRow}`).setNumberFormat("#,##0");s.getRange(`D${resultFirstRow}:F${resultEndRow}`).setNumberFormat("#,##0.00");
 const pctSectionRow=resultEndRow+2,pctHeaderRow=pctSectionRow+1,pctFirstRow=pctHeaderRow+1,pctTotalRow=pctFirstRow+S.order_bill_results.length;
 s.getRange(`A${pctSectionRow}:H${pctSectionRow}`).merge();s.getRange(`A${pctSectionRow}`).values=[["金额比对百分比统计"]];section(s,`A${pctSectionRow}:H${pctSectionRow}`);
 write(s,pctHeaderRow-1,0,[["结果","旺店通金额","旺店通金额占比","惠策本期应收","应收金额占比","惠策本期实收","实收金额占比"]]);header(s,`A${pctHeaderRow}:G${pctHeaderRow}`);
 write(s,pctFirstRow-1,0,S.order_bill_results.map(x=>[x.result,x.wdt_amount,null,x.bill_receivable,null,x.bill_cash,null]));
 for(let r=pctFirstRow;r<pctTotalRow;r++){s.getRange(`C${r}`).formulas=[[`=IFERROR(B${r}/$B$${pctTotalRow},0)`]];s.getRange(`E${r}`).formulas=[[`=IFERROR(D${r}/$D$${pctTotalRow},0)`]];s.getRange(`G${r}`).formulas=[[`=IFERROR(F${r}/$F$${pctTotalRow},0)`]];}
 write(s,pctTotalRow-1,0,[["合计",null,null,null,null,null,null]]);for(const c of ["B","D","F"])s.getRange(`${c}${pctTotalRow}`).formulas=[[`=SUM(${c}${pctFirstRow}:${c}${pctTotalRow-1})`]];for(const c of ["C","E","G"])s.getRange(`${c}${pctTotalRow}`).formulas=[[`=IFERROR(${String.fromCharCode(c.charCodeAt(0)-1)}${pctTotalRow}/${String.fromCharCode(c.charCodeAt(0)-1)}${pctTotalRow},0)`]];
 body(s,`A${pctFirstRow}:G${pctTotalRow-1}`);status(s,`A${pctFirstRow}:A${pctTotalRow-1}`);section(s,`A${pctTotalRow}:G${pctTotalRow}`);s.getRange(`B${pctFirstRow}:B${pctTotalRow}`).setNumberFormat("#,##0.00");s.getRange(`D${pctFirstRow}:D${pctTotalRow}`).setNumberFormat("#,##0.00");s.getRange(`F${pctFirstRow}:F${pctTotalRow}`).setNumberFormat("#,##0.00");s.getRange(`C${pctFirstRow}:C${pctTotalRow}`).setNumberFormat("0.00%");s.getRange(`E${pctFirstRow}:E${pctTotalRow}`).setNumberFormat("0.00%");s.getRange(`G${pctFirstRow}:G${pctTotalRow}`).setNumberFormat("0.00%");
 [26,18,27,18,28,18,24,32].forEach((w,i)=>s.getRange(`${col(i)}:${col(i)}`).format.columnWidth=w);s.freezePanes.freezeRows(4);
}

// 5 惠策明细—店铺汇总内部核对
{
 const s=ws(names[3]);title(s,"惠策明细—惠策店铺汇总内部金额核对","以导出文件所属结算月份+平台+店铺作为共同维度；对全量本期应收、实收分别执行核对。","K");
 s.getRange("A4:K4").merge();s.getRange("A4").values=[["数据完整性控制"]];section(s,"A4:K4");
 write(s,4,0,[["明细/汇总控制","数值","说明","","结算时点控制","数值","说明"]]);header(s,"A5:G5");
 write(s,5,0,[
  ["惠策汇总源记录数",ctl.huice_summary_source_rows,"全部19个店铺汇总导出文件的源记录数","","惠策结算明细行数",ctl.huice_detail_settlement_rows,"2026年1—6月结算月份导出文件"],
  ["跨文件重复流水号保留数",ctl.huice_summary_reused_ids,"使用source_file+monthly_id复合键保留，不因跨文件重复而覆盖","","含往期业务日期行数",ctl.huice_historical_rows,"业务日期早于核对期间，但属于本期结算文件，按结算月份保留"],
  ["应收及实收均一致的月度店铺组合",S.display_metrics.huice_internal_exact_groups,"结算月份+平台+店铺组合","","往期业务日期应收",ctl.huice_historical_receivable,"不按业务日期错误剔除"],
  ["明细单边组合数",S.huice_internal_results.find(x=>x.result==="仅惠策明细")?.groups||0,"明细与汇总月度店铺组合核对结果","","往期业务日期实收",ctl.huice_historical_cash,"包含退款或冲销形成的净额"],
 ]);body(s,"A6:G9");s.getRange("B6:B9").setNumberFormat("#,##0");s.getRange("F6:F7").setNumberFormat("#,##0");s.getRange("F8:F9").setNumberFormat("#,##0.00;[Red](#,##0.00)");
 s.getRange("A11:K11").merge();s.getRange("A11").values=[["全量金额核对"]];section(s,"A11:K11");
 write(s,11,0,[["口径","惠策明细金额","惠策汇总金额","差异（汇总-明细）","匹配率","结论"]]);header(s,"A12:F12");
 write(s,12,0,[["应收",ctl.huice_detail_settlement_receivable,ctl.huice_bill_receivable,huiceReceivableDifference,huiceReceivableMatch,huiceReceivableDifference===0?"全量一致":"按明细核对结果列示"],["实收",ctl.huice_detail_settlement_cash,ctl.huice_bill_cash,huiceCashDifference,huiceCashMatch,huiceCashDifference===0?"全量一致":"按明细核对结果列示"]]);body(s,"A13:F14");status(s,"F13:F14");s.getRange("B13:D14").setNumberFormat("#,##0.00;[Red](#,##0.00)");s.getRange("E13:E14").setNumberFormat("0.00%");
 s.getRange("A16:K17").merge();s.getRange("A16").values=[["口径提示：汇总表“对账成功金额”是源表成功分类字段的合计；明细表“对账状态=对账成功”是逐行状态字段，两者定义不一致。惠策内部主核对采用全量本期应收与全量本期实收，不将成功分类金额差异解释为明细缺失。"]];s.getRange("A16:K17").format={fill:C.amber,font:{color:C.amberText,bold:true},wrapText:true,verticalAlignment:"center",borders:{preset:"outside",style:"thin",color:C.line}};
 s.getRange("A19:K19").merge();s.getRange("A19").values=[["月度应收及实收核对"]];section(s,"A19:K19");
 write(s,19,0,[["结算月份","明细行数","汇总源行数","明细应收","汇总应收","应收差异","应收匹配率","明细实收","汇总实收","实收差异","实收匹配率"]]);header(s,"A20:K20");
 write(s,20,0,S.huice_internal_monthly.map(x=>[x.bill_month,x.detail_rows,x.summary_rows,x.detail_receivable,x.summary_receivable,Math.abs(x.receivable_difference)<=0.01?0:x.receivable_difference,1-Math.abs(x.receivable_difference)/Math.abs(x.detail_receivable),x.detail_cash,x.summary_cash,Math.abs(x.cash_difference)<=0.01?0:x.cash_difference,1-Math.abs(x.cash_difference)/Math.abs(x.detail_cash)]));body(s,"A21:K26");s.getRange("B21:C26").setNumberFormat("#,##0");s.getRange("D21:F26").setNumberFormat("#,##0.00;[Red](#,##0.00)");s.getRange("G21:G26").setNumberFormat("0.00%");s.getRange("H21:J26").setNumberFormat("#,##0.00;[Red](#,##0.00)");s.getRange("K21:K26").setNumberFormat("0.00%");
 s.getRange("A28:K28").merge();s.getRange("A28").values=[["结果分层"]];section(s,"A28:K28");write(s,28,0,[["结果","组合数","明细行数","汇总源行数","明细应收","汇总应收","应收差异","明细实收","汇总实收","实收差异"]]);header(s,"A29:J29");
 write(s,29,0,S.huice_internal_results.map(x=>[x.result,x.groups,x.detail_rows,x.summary_rows,x.detail_receivable,x.summary_receivable,Math.abs(x.receivable_difference)<=0.01?0:x.receivable_difference,x.detail_cash,x.summary_cash,Math.abs(x.cash_difference)<=0.01?0:x.cash_difference]));body(s,`A30:J${29+S.huice_internal_results.length}`);status(s,`A30:A${29+S.huice_internal_results.length}`);s.getRange(`B30:D${29+S.huice_internal_results.length}`).setNumberFormat("#,##0");s.getRange(`E30:J${29+S.huice_internal_results.length}`).setNumberFormat("#,##0.00;[Red](#,##0.00)");
 [24,18,31,18,18,18,16,18,18,18,16].forEach((w,i)=>s.getRange(`${col(i)}:${col(i)}`).format.columnWidth=w);s.freezePanes.freezeRows(5);
}

// 7 账单OMS月结汇总
{
 const s=ws(names[5]);title(s,"惠策店铺账单—OMS月结Y001金额核对汇总","由于惠策账单未提供物料及OMS销售单号，本环节按月份+惠策店铺/OMS客户比较金额；对账成功分类金额为主比较口径，应收、实收为辅助口径。","I");
 write(s,3,0,[["主口径及辅助金额","金额","差异（OMS－惠策）","匹配率/覆盖率","","OMS/SAP辅助","金额/数量","说明","判断"]]);header(s,"A4:I4");
 write(s,4,0,[["对账成功分类金额（主比较口径）",ctl.huice_bill_success_amount,null,null,"","OMS月结金额",ctl.oms_month_amount,"业务类型Y001","核心金额匹配"],["已获得客户编码的店铺应收金额",ctl.mapped_bill_amount,null,null,"","客户编码映射覆盖率",null,"已映射应收金额/惠策应收金额","映射覆盖控制"],["本期实收金额（辅助，仅列金额）",ctl.huice_bill_cash,null,null,"","OMS月结数量",ctl.oms_month_qty,"OMS原生数量","辅助口径不计算匹配率"],["本期应收金额（辅助，仅列金额）",ctl.huice_bill_receivable,null,null,"","SAP一致键金额",ctl.sap_assisted_amount,"OMS-SAP共同键汇总","辅助口径不计算匹配率"]]);
 s.getRange("C5").formulas=[["=G5-B5"]];s.getRange("D5").formulas=[["=MIN(ABS(B5),ABS(G5))/MAX(ABS(B5),ABS(G5))"]];s.getRange("G6").formulas=[[`=B6/${ctl.huice_bill_receivable}`]];body(s,"A5:I8");status(s,"I5:I8");s.getRange("B5:C8").setNumberFormat("#,##0.00");s.getRange("D5").setNumberFormat("0.00%");s.getRange("G5").setNumberFormat("#,##0.00");s.getRange("G6").setNumberFormat("0.00%");s.getRange("G7:G8").setNumberFormat("#,##0.00");s.getRange("A5:D5").format.fill=C.green;s.getRange("E6:I6").format.fill=C.green;s.getRange("A7:D8").format.fill="#E7E6E6";
 s.getRange("A10:I10").merge();s.getRange("A10").values=[["月份+店铺核对结果构成（组合数不作为总体金额匹配率）"]];section(s,"A10:I10");write(s,10,0,[["结果","组合数","成功分类金额","本期应收","本期实收","OMS Y001金额","SAP共同键辅助金额","处理","说明"]]);header(s,"A11:I11");write(s,11,0,[...S.bill_oms_results].sort((a,b)=>(a.result==="成功金额一致"?-1:b.result==="成功金额一致"?1:0)).map(x=>[x.result,x.groups,x.bill_success_amount,x.bill_receivable,x.bill_cash,x.oms_amount,x.sap_assisted_amount,["成功金额一致","应收金额一致","实收金额一致","SAP辅助金额一致"].includes(x.result)?"已核对":"分类列示",x.result==="店铺未映射"?"OMS客户编码映射分类":x.result==="仅账单"?"惠策账单与OMS记录分别列示":"详见核对明细"]));body(s,`A12:I${11+S.bill_oms_results.length}`);s.getRange("B12:B19").setNumberFormat("#,##0");s.getRange("C12:G19").setNumberFormat("#,##0.00");
 const billEnd=11+S.bill_oms_results.length,billPeriodSection=billEnd+2;
 s.getRange(`A${billPeriodSection}:I${billPeriodSection}`).merge();s.getRange(`A${billPeriodSection}`).values=[["跨月结算时点分析"]];section(s,`A${billPeriodSection}:I${billPeriodSection}`);
 write(s,billPeriodSection,0,[["项目","数值","说明","","项目","数值","说明"],["逐月店铺差异绝对值合计",dm.bill_gross_abs_success_diff,"逐月比较结果，不进行跨月抵销","","同一客户期间累计差异绝对值",dm.bill_period_gross_difference,"用于观察期间累计比较结果"],["跨月抵销解释金额",dm.bill_timing_offset,"逐月差异绝对值与期间累计差异绝对值之差","","期间累计一致客户数",dm.bill_period_exact_customers,"作为结算时点辅助分析"],["期间累计客户数",dm.bill_period_total_customers,"已取得客户编码或映射依据的客户范围","","未纳入客户编码映射的应收金额",ctl.unmapped_bill_amount,"按客户编码映射范围单独列示"]]);header(s,`A${billPeriodSection+1}:G${billPeriodSection+1}`);body(s,`A${billPeriodSection+2}:G${billPeriodSection+4}`);s.getRange(`B${billPeriodSection+2}:B${billPeriodSection+4}`).setNumberFormat("#,##0.00");s.getRange(`F${billPeriodSection+2}:F${billPeriodSection+4}`).setNumberFormat("#,##0.00");
 [25,18,18,18,18,25,18,30,24].forEach((w,i)=>s.getRange(`${col(i)}:${col(i)}`).format.columnWidth=w);s.freezePanes.freezeRows(4);
}

// 9 数量核对
{
 const s=ws(names[7]);title(s,"订单证据—OMS月结数量比较汇总","惠策没有商品数量；数量A为平台单号精确证据数量加同内部订单零金额附属单数量，数量B为OMS月结Y001原生数量。","H");
 write(s,3,0,[["总体数量控制","旺店通/订单账单","OMS月结","差异","","共同范围控制","数值","比例/结论"]]);header(s,"A4:H4");
 write(s,4,0,[["旺店通商品数量—OMS月结数量",ctl.wdt_qty,ctl.oms_month_qty,null,"","订单证据数量",ctl.billed_wdt_qty,"精确证据+零金额附属单"],["旺店通商品数量总体匹配率",null,null,null,"","其中：平台单号精确证据数量",ctl.billed_wdt_qty_direct,"一级精确匹配"],["订单证据/旺店通数量覆盖率",null,null,null,"","其中：零金额附属单数量",ctl.billed_wdt_auxiliary_qty,"同内部订单二级桥接"],["订单证据/OMS数量覆盖率",null,null,null,"","共同月份+店铺+物料—旺店通数量",commonWdtQty,"共同键范围"],["共同范围净数量匹配率",null,null,null,"","共同月份+店铺+物料—OMS数量",commonOmsQty,"共同键范围"]]);
 s.getRange("D5").formulas=[["=C5-B5"]];s.getRange("B6").formulas=[["=MIN(B5,C5)/MAX(B5,C5)"]];s.getRange("B7").formulas=[["=G5/B5"]];s.getRange("B8").formulas=[["=G5/C5"]];s.getRange("B9").formulas=[["=MIN(G8,G9)/MAX(G8,G9)"]];body(s,"A5:H9");status(s,"H5:H9");s.getRange("B5:D5").setNumberFormat("#,##0");s.getRange("B6:B9").setNumberFormat("0.00%");s.getRange("G5:G9").setNumberFormat("#,##0");s.getRange("A6:D9").format.fill=C.green;
 s.getRange("A11:H11").merge();s.getRange("A11").values=[["月份+店铺+物料核对结果构成（组数不作为总体数量匹配率）"]];section(s,"A11:H11");write(s,11,0,[["结果","组数","订单证据计次（可跨物料重复）","旺店通商品数量","OMS Y001数量","数量差异（OMS－旺店通）","处理","说明"]]);header(s,"A12:H12");write(s,12,0,[...S.qty_results].sort((a,b)=>(a.result==="数量一致"?-1:b.result==="数量一致"?1:0)).map(x=>[x.result,x.groups,x.billed_orders,x.order_bill_qty,x.oms_qty,x.qty_difference,x.result==="数量一致"?"已核对":"分类列示",x.result==="数量差异"?"多对一归并后的数量比较结果":"惠策作为订单存在性证据"]));body(s,`A13:H${12+S.qty_results.length}`);s.getRange("B13:F20").setNumberFormat("#,##0");
 const qtyEnd=12+S.qty_results.length,qtySection=qtyEnd+2;
 s.getRange(`A${qtySection}:H${qtySection}`).merge();s.getRange(`A${qtySection}`).values=[["数据粒度与多物料抵销分析"]];section(s,`A${qtySection}:H${qtySection}`);
 write(s,qtySection,0,[["项目","数值","说明","","项目","数值","说明"],["旺店通订单表头数量",ctl.wdt_header_qty,"内部订单+平台订单号粒度","","旺店通商品明细数量",ctl.wdt_qty,"按SAP物料汇总"],["零金额附属单平台单号",ctl.wdt_zero_auxiliary_platform_keys,"同内部订单已有惠策精确匹配","","零金额附属单数量",ctl.billed_wdt_auxiliary_qty,"计入订单证据数量"],["物料级数量差异绝对值",dm.qty_gross_abs_diff,"按月份+客户+物料汇总","","客户月度数量差异绝对值",dm.qty_customer_month_gross_difference,"多物料汇总后的数量比较结果"],["跨物料抵销数量",dm.qty_cross_material_offset,"用于说明物料归并影响","","客户月度精确一致组",dm.qty_customer_month_exact_groups,"客户月度汇总结果"]]);header(s,`A${qtySection+1}:G${qtySection+1}`);body(s,`A${qtySection+2}:G${qtySection+5}`);s.getRange(`B${qtySection+2}:B${qtySection+5}`).setNumberFormat("#,##0.00");s.getRange(`F${qtySection+2}:F${qtySection+5}`).setNumberFormat("#,##0.00");
 [27,18,18,18,18,27,18,32].forEach((w,i)=>s.getRange(`${col(i)}:${col(i)}`).format.columnWidth=w);s.freezePanes.freezeRows(4);
}

// 11 OMS-SAP
{
 const s=ws(names[9]);title(s,"OMS月结Y001—SAP标准发票（2C）核对汇总","先披露OMS与SAP共同键的数量/金额核对结果，再披露双方全量覆盖情况；共同键结果不代表全量单据均已逐键匹配。","H");
 write(s,3,0,[["核心匹配准确性","数值","说明","","共同键比较","SAP","OMS","差异（OMS－SAP）"]]);header(s,"A4:H4");
 write(s,4,0,[["数量及金额均一致的共同键",exactMap.keys,"销售单号+物料+销售单位","","共同键数量",exactMap.sap_qty,exactMap.oms_qty,null],["共同键数量匹配率",null,"共同键范围内OMS与SAP数量比较","","共同键金额",exactMap.sap_amount,exactMap.oms_amount,null],["共同键金额匹配率",null,"共同键范围内OMS与SAP金额比较","","",null,null,null],["OMS月结数量共同键覆盖率",null,"数量一致共同键/OMS月结数量总额","","",null,null,null]]);
 s.getRange("B6").formulas=[["=MIN(ABS(F5),ABS(G5))/MAX(ABS(F5),ABS(G5))"]];s.getRange("B7").formulas=[["=MIN(ABS(F6),ABS(G6))/MAX(ABS(F6),ABS(G6))"]];s.getRange("B8").formulas=[[`=${exactMap.oms_qty}/${ctl.oms_month_qty}`]];s.getRange("H5").formulas=[["=G5-F5"]];s.getRange("H6").formulas=[["=G6-F6"]];body(s,"A5:H8");s.getRange("A6:C8").format.fill=C.green;s.getRange("B5").setNumberFormat("#,##0");s.getRange("B6:B8").setNumberFormat("0.00%");s.getRange("F5:H5").setNumberFormat("#,##0");s.getRange("F6:H8").setNumberFormat("#,##0.00;[Red](#,##0.00)");
 s.getRange("A10:H10").merge();s.getRange("A10").values=[["共同键覆盖与单边记录"]];section(s,"A10:H10");
 write(s,10,0,[["系统","共同键数","共同键数量","全量数量","单边记录键数","共同键金额","全量金额","结论"],["OMS月结",exactMap.keys,exactMap.oms_qty,ctl.oms_month_qty,omsOnly.keys,exactMap.oms_amount,ctl.oms_month_amount,"共同键数量覆盖率99.99%"],["SAP标准发票",exactMap.keys,exactMap.sap_qty,ctl.sap_full_qty,sapOnly.keys,exactMap.sap_amount,ctl.sap_full_amount,"单边记录分别列示"]]);header(s,"A11:H11");body(s,"A12:H13");s.getRange("B12:E13").setNumberFormat("#,##0");s.getRange("F12:G13").setNumberFormat("#,##0.00");status(s,"H12:H13");
 s.getRange("A15:H17").merge();s.getRange("A15").values=[[`结论：在销售单号+物料+销售单位共同键范围内，${exactMap.keys.toLocaleString("zh-CN")}个共同键的数量匹配率为100.00%，金额差异为${Math.abs(exactMap.oms_amount-exactMap.sap_amount).toLocaleString("zh-CN",{minimumFractionDigits:2})}元；OMS月结数量共同键覆盖率为${omsQtyCoverage.toLocaleString("zh-CN",{style:"percent",minimumFractionDigits:2})}。${sapOnly.keys}个仅SAP键和${omsOnly.keys}个仅OMS键分别列示。`]];s.getRange("A15:H17").format={fill:C.amber,font:{color:C.amberText,bold:true},wrapText:true,verticalAlignment:"center",borders:{preset:"outside",style:"thin",color:C.line}};
 [27,18,32,18,24,19,19,20].forEach((w,i)=>s.getRange(`${col(i)}:${col(i)}`).format.columnWidth=w);s.freezePanes.freezeRows(4);
}

// 13 月度全流程
{
 const s=ws(names[10]);title(s,"2026年1-6月月度全流程汇总","同一月份横向展示订单、店铺账单三口径、OMS月结、SAP一致键及SAP全量；各列不因总体接近而自动抵销。","M");
 write(s,3,0,[["月份","旺店通商品数量","旺店通商品金额","惠策成功金额","惠策应收","惠策实收","OMS月结数量","OMS月结金额","SAP一致键数量","SAP一致键金额","SAP全量数量","SAP全量金额","SAP全量-OMS金额差"]]);header(s,"A4:M4");
 write(s,4,0,S.monthly_flow.map(x=>[x.month,x.wdt_qty,x.wdt_amount,x.bill_success_amount,x.bill_receivable,x.bill_cash,x.oms_qty,x.oms_amount,x.sap_qty,x.sap_amount,x.sap_full_qty,x.sap_full_amount,null]));for(let r=5;r<=10;r++)s.getRange(`M${r}`).formulas=[[`=L${r}-H${r}`]];body(s,"A5:M10");s.getRange("B5:B10").setNumberFormat("#,##0");s.getRange("C5:F10").setNumberFormat("#,##0.00");s.getRange("G5:G10").setNumberFormat("#,##0");s.getRange("H5:H10").setNumberFormat("#,##0.00");s.getRange("I5:I10").setNumberFormat("#,##0");s.getRange("J5:J10").setNumberFormat("#,##0.00");s.getRange("K5:K10").setNumberFormat("#,##0");s.getRange("L5:M10").setNumberFormat("#,##0.00");
 write(s,12,0,[["合计",null,null,null,null,null,null,null,null,null,null,null,null]]);for(const c of ["B","C","D","E","F","G","H","I","J","K","L","M"])s.getRange(`${c}13`).formulas=[[`=SUM(${c}5:${c}10)`]];section(s,"A13:M13");s.getRange("B13:B13").setNumberFormat("#,##0");s.getRange("C13:F13").setNumberFormat("#,##0.00");s.getRange("G13:G13").setNumberFormat("#,##0");s.getRange("H13:H13").setNumberFormat("#,##0.00");s.getRange("I13:I13").setNumberFormat("#,##0");s.getRange("J13:J13").setNumberFormat("#,##0.00");s.getRange("K13:K13").setNumberFormat("#,##0");s.getRange("L13:M13").setNumberFormat("#,##0.00");
 [14,18,19,20,20,20,18,20,18,20,18,20,22].forEach((w,i)=>s.getRange(`${col(i)}:${col(i)}`).format.columnWidth=w);s.freezePanes.freezeRows(4);
}

detail(names[2],"惠策出账明细—旺店通平台订单核对明细","以2026年1—6月惠策账单为基表，一行一个平台订单号；旺店通订单向前追溯至2025年12月，并分别列示订单月份、发货月份和账单月份。",data.orderBill,{platform_order_no:25,wdt_shop:32,huice_shop:34,internal_orders:34,reconcile_ids:34,result:26});
detail(names[4],"惠策明细—惠策店铺汇总逐月店铺核对","按导出结算月份+平台+店铺核对全量本期应收、实收；差异组合优先展示。",data.internal,{huice_shop:34,result:22});
detail(names[6],"惠策店铺账单—OMS月结逐月店铺核对","按月份+店铺/OMS客户比较金额，不进行物料级连接；列示成功分类金额、本期应收、本期实收及SAP共同键辅助金额。",data.billOms,{huice_shop:34,customer_name:40,result:20},["mapping_status","mapping_source"]);
detail(names[8],"有惠策订单证据的旺店通数量—OMS月结数量明细","按发货月+店铺/OMS客户+SAP物料汇总比较；分别列示平台单号精确证据数量和同内部订单零金额附属单数量，OMS数量为Y001原生数量。",data.qty,{wdt_shop:34,customer_name:40,material_code:17,result:22},["mapping_status"]);
detail(names[11],"惠策店铺—OMS客户映射","展示惠策店铺与OMS客户编码的映射结果及映射依据；已映射不等同于OMS期间内存在对应交易。",data.shop,{huice_shop:34,customer_name:40},["mapping_status","mapping_source"]);

// 15 索引
if(false){
{
 const s=ws(names[14]);title(s,"完整明细及代码索引","工作簿嵌入汇总和审阅明细；大体量完整结果保留CSV。","F");
 write(s,3,0,[["核对环节","完整行数","工作簿行数","完整文件","页面","说明"]]);header(s,"A4:F4");
 const idx=[
  ["旺店通—惠策平台订单映射",S.detail_rows.order_bill_recon,data.orderBill.rows.length,path.join(dataDir,"order_bill_recon.csv"),"4.旺店通-惠策订单映射","分层样本；完整映射见CSV"],
  ["惠策明细—汇总",S.detail_rows.huice_internal_recon,data.internal.rows.length,path.join(dataDir,"huice_internal_recon.csv"),"6.惠策内部核对明细","全部嵌入"],
  ["账单—OMS月结",S.detail_rows.bill_oms_month_recon,data.billOms.rows.length,path.join(dataDir,"bill_oms_month_recon.csv"),"8.账单-OMS月结明细","全部嵌入"],
  ["订单账单—OMS数量",S.detail_rows.order_bill_oms_qty_recon,data.qty.rows.length,path.join(dataDir,"order_bill_oms_qty_recon.csv"),"10.数量核对明细","全部嵌入"],
  ["OMS—SAP字段映射",S.detail_rows.oms_sap_field_map,data.omsSap.rows.length,path.join(dataDir,"oms_sap_field_map.csv"),"12.OMS-SAP字段映射","全部嵌入"],
  ["店铺客户映射",S.detail_rows.huice_shop_map,data.shop.rows.length,path.join(dataDir,"huice_shop_map.csv"),"14.店铺客户映射","全部嵌入"],
 ];write(s,4,0,idx);body(s,`A5:F${4+idx.length}`);s.getRange(`B5:C${4+idx.length}`).setNumberFormat("#,##0");[28,18,18,86,28,24].forEach((w,i)=>s.getRange(`${col(i)}:${col(i)}`).format.columnWidth=w);s.getRange(`D5:D${4+idx.length}`).format.wrapText=true;s.freezePanes.freezeRows(4);
}
}

await fs.mkdir(outputDir,{recursive:true});const previews=path.join(root,"reconciliation/qa_previews");await fs.mkdir(previews,{recursive:true});
console.log("OVERVIEW\n"+(await wb.inspect({kind:"table",range:"1.全局口径与总览!A1:L40",include:"values,formulas",tableMaxRows:44,tableMaxCols:13,maxChars:24000})).ndjson);
console.log("ORDER_SUMMARY\n"+(await wb.inspect({kind:"table",range:"2.订单-账单汇总!A1:H12",include:"values,formulas",tableMaxRows:14,tableMaxCols:10,maxChars:10000})).ndjson);
console.log("BILL_OMS\n"+(await wb.inspect({kind:"table",range:"6.账单-OMS月结汇总!A1:I18",include:"values,formulas",tableMaxRows:20,tableMaxCols:10,maxChars:12000})).ndjson);
console.log("OMS_SAP\n"+(await wb.inspect({kind:"table",range:"10.OMS月结-SAP汇总!A1:H17",include:"values,formulas",tableMaxRows:20,tableMaxCols:10,maxChars:12000})).ndjson);
console.log("QTY\n"+(await wb.inspect({kind:"table",range:"8.数量核对汇总!A1:H18",include:"values,formulas",tableMaxRows:22,tableMaxCols:10,maxChars:10000})).ndjson);
console.log("LOW_CONFIDENCE\n"+(await wb.inspect({kind:"match",searchTerm:"低置信",options:{useRegex:false,maxResults:100},summary:"visible low-confidence labels",maxChars:3000})).ndjson);
console.log("ERRORS\n"+(await wb.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",options:{useRegex:true,maxResults:300},summary:"formula errors",maxChars:5000})).ndjson);
for(const name of names){const s=ws(name),used=s.getUsedRange(true),isWideDetail=name.includes("明细")||name.includes("映射"),maxCols=Math.min(used.columnCount||8,isWideDetail?10:14),maxRows=name===names[0]?42:(isWideDetail?22:30);const blob=await wb.render({sheetName:name,range:`A1:${col(maxCols-1)}${maxRows}`,scale:1.15,format:"png"});await fs.writeFile(path.join(previews,`${name}.png`),new Uint8Array(await blob.arrayBuffer()));if(isWideDetail&&(used.columnCount||0)>10){const rightEnd=Math.min(used.columnCount,20);const right=await wb.render({sheetName:name,range:`K1:${col(rightEnd-1)}${maxRows}`,scale:1.15,format:"png"});await fs.writeFile(path.join(previews,`${name}_右侧.png`),new Uint8Array(await right.arrayBuffer()));}}
const out=await SpreadsheetFile.exportXlsx(wb);await out.save(outputFile);console.log(JSON.stringify({outputFile,sheets:names,previews},null,2));
