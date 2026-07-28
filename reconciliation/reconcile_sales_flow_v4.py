#!/usr/bin/env python3
"""2026年1-6月销售ToC流程核对：订单→账单→OMS月结→SAP。

核对原则：业务流程顺序展示，各环节单独执行pairwise核对。
惠策不含商品数量，因此数量链使用“惠策已出现订单对应的旺店通商品数量”。
"""

from __future__ import annotations

import argparse
import csv
import json
import sqlite3
from pathlib import Path

ROOT = Path(__file__).resolve().parent
DB = ROOT / "work_full" / "reconciliation.db"
OUT = ROOT / "output_flow_v4"
START = "2026-01-01"
END_EXCLUSIVE = "2026-07-01"


def configure(conn: sqlite3.Connection) -> None:
    conn.executescript("PRAGMA journal_mode=OFF; PRAGMA synchronous=OFF; PRAGMA temp_store=MEMORY; PRAGMA cache_size=-400000;")


def build(conn: sqlite3.Connection) -> None:
    conn.executescript(f"""
    DROP TABLE IF EXISTS v4_wdt_platform;
    CREATE TABLE v4_wdt_platform AS
    SELECT CASE WHEN platform_order_no='' THEN '__WDT__'||order_no ELSE platform_order_no END platform_order_no,
      CASE WHEN platform_order_no='' THEN 0 ELSE 1 END matchable,
      MIN(shop) wdt_shop,MIN(substr(ship_time,1,7)) ship_month,
      COUNT(*) internal_order_count,SUM(receivable_amount) wdt_amount,SUM(allocated_total) wdt_allocated_amount,
      SUM(quantity) wdt_qty,GROUP_CONCAT(order_no,'|') internal_orders
    FROM wdt_order_dedup
    WHERE ship_time>='{START}' AND ship_time<'{END_EXCLUSIVE}'
    GROUP BY CASE WHEN platform_order_no='' THEN '__WDT__'||order_no ELSE platform_order_no END;
    CREATE UNIQUE INDEX idx_v4_wdt_platform ON v4_wdt_platform(platform_order_no);

    DROP TABLE IF EXISTS v4_huice_platform;
    CREATE TABLE v4_huice_platform AS
    SELECT CASE WHEN h.platform_order_no='' THEN '__HC__'||h.reconcile_id ELSE h.platform_order_no END platform_order_no,
      MIN(h.platform) platform,MIN(h.shop) huice_shop,
      MIN(substr(COALESCE(NULLIF(h.business_date,''),h.period_end),1,7)) bill_month,
      COUNT(*) huice_rows,SUM(c.current_receivable) bill_receivable,SUM(c.current_cash) bill_cash,
      GROUP_CONCAT(h.reconcile_id,'|') reconcile_ids
    FROM huice_detail h JOIN huice_current_amount c ON c.reconcile_id=h.reconcile_id
    WHERE COALESCE(NULLIF(h.business_date,''),h.period_end)>='{START}'
      AND COALESCE(NULLIF(h.business_date,''),h.period_end)<'{END_EXCLUSIVE}'
    GROUP BY CASE WHEN h.platform_order_no='' THEN '__HC__'||h.reconcile_id ELSE h.platform_order_no END;
    CREATE UNIQUE INDEX idx_v4_huice_platform ON v4_huice_platform(platform_order_no);

    DROP TABLE IF EXISTS v4_order_bill_recon;
    CREATE TABLE v4_order_bill_recon AS
    SELECT w.platform_order_no,w.matchable,w.wdt_shop,h.platform,h.huice_shop,w.ship_month,h.bill_month,
      w.internal_order_count,COALESCE(h.huice_rows,0) huice_rows,w.wdt_qty,w.wdt_amount,w.wdt_allocated_amount,
      COALESCE(h.bill_receivable,0) bill_receivable,COALESCE(h.bill_cash,0) bill_cash,
      w.wdt_amount-COALESCE(h.bill_receivable,0) receivable_difference,
      w.wdt_amount-COALESCE(h.bill_cash,0) cash_difference,w.internal_orders,COALESCE(h.reconcile_ids,'') reconcile_ids,
      CASE WHEN h.platform_order_no IS NULL AND w.matchable=0 THEN '旺店通原始单号为空'
           WHEN h.platform_order_no IS NULL THEN '仅订单'
           WHEN ABS(w.wdt_amount-h.bill_receivable)<=0.01 THEN '单号应收金额一致'
           WHEN ABS(w.wdt_amount-h.bill_cash)<=0.01 THEN '单号实收金额一致'
           ELSE '单号一致金额差异' END result
    FROM v4_wdt_platform w LEFT JOIN v4_huice_platform h ON h.platform_order_no=w.platform_order_no
    UNION ALL
    SELECT h.platform_order_no,1,'',h.platform,h.huice_shop,'',h.bill_month,0,h.huice_rows,0,0,0,
      h.bill_receivable,h.bill_cash,-h.bill_receivable,-h.bill_cash,'',h.reconcile_ids,'仅账单'
    FROM v4_huice_platform h LEFT JOIN v4_wdt_platform w ON w.platform_order_no=h.platform_order_no
    WHERE w.platform_order_no IS NULL;

    DROP TABLE IF EXISTS v4_huice_shop_bill;
    CREATE TABLE v4_huice_shop_bill AS
    SELECT CASE WHEN source_file LIKE '%1月%' THEN '2026-01' WHEN source_file LIKE '%2月%' THEN '2026-02'
                WHEN source_file LIKE '%3月%' THEN '2026-03' WHEN source_file LIKE '%4月%' THEN '2026-04'
                WHEN source_file LIKE '%5月%' THEN '2026-05' WHEN source_file LIKE '%6月%' THEN '2026-06' END bill_month,
      platform,shop huice_shop,
      SUM(success_count+mismatch_count+single_ar_count+single_cash_count) bill_record_count,
      SUM(success_count) success_count,SUM(success_amount) bill_success_amount,
      SUM(success_amount+mismatch_receivable+single_ar_amount) bill_receivable,
      SUM(success_amount+mismatch_cash+single_cash_amount) bill_cash,
      COUNT(*) source_rows
    FROM huice_summary
    WHERE source_file LIKE '%1月%' OR source_file LIKE '%2月%' OR source_file LIKE '%3月%'
       OR source_file LIKE '%4月%' OR source_file LIKE '%5月%' OR source_file LIKE '%6月%'
    GROUP BY 1,2,3;

    DROP TABLE IF EXISTS v4_huice_shop_map;
    CREATE TABLE v4_huice_shop_map AS
    SELECT b.bill_month,b.platform,b.huice_shop,
      COALESCE(m.customer_code,d.customer_code) customer_code,
      COALESCE(m.customer_name,d.customer_name) customer_name,
      COALESCE(m.final_status,CASE WHEN d.customer_code IS NOT NULL THEN d.mapping_status ELSE '未映射' END) mapping_status,
      CASE WHEN m.customer_code IS NOT NULL THEN '订单桥接映射'
           WHEN d.customer_code IS NOT NULL THEN '同名店铺映射' ELSE '未映射' END mapping_source,
      b.bill_record_count,b.bill_receivable,b.bill_cash,b.source_rows
      ,b.success_count,b.bill_success_amount
    FROM v4_huice_shop_bill b
    LEFT JOIN huice_oms_shop_map m ON m.platform=b.platform AND m.huice_shop=b.huice_shop
    LEFT JOIN wdt_oms_shop_map d ON d.wdt_shop=b.huice_shop;

    DROP TABLE IF EXISTS v4_oms_month_shop;
    CREATE TABLE v4_oms_month_shop AS
    SELECT outbound_month,customer_code,MIN(customer_name) customer_name,
      COUNT(DISTINCT document_no) oms_docs,SUM(item_num) oms_qty,SUM(share_amount) oms_amount,COUNT(*) oms_lines
    FROM oms_detail
    WHERE business_type='Y001' AND outbound_time>='{START}' AND outbound_time<'{END_EXCLUSIVE}'
    GROUP BY 1,2;

    DROP TABLE IF EXISTS v4_oms_sap_field_map;
    CREATE TABLE v4_oms_sap_field_map AS
    WITH od AS (
      SELECT document_no,item_code,sale_unit,MIN(outbound_month) outbound_month,MIN(customer_code) customer_code,
        MIN(customer_name) customer_name,SUM(item_num) oms_source_qty,SUM(share_amount) oms_source_amount,
        COUNT(*) oms_source_rows
      FROM oms_detail
      WHERE business_type='Y001' AND outbound_time>='{START}' AND outbound_time<'{END_EXCLUSIVE}'
      GROUP BY 1,2,3
    )
    SELECT r.oms_sales_no,r.material_code,r.sales_unit,r.file_month,od.outbound_month,r.sap_invoice_nos,
      od.customer_code,od.customer_name,r.sap_rows,r.oms_rows,
      r.invoice_qty sap_qty,COALESCE(od.oms_source_qty,r.oms_qty) oms_qty,r.quantity_difference,
      r.sap_amount,COALESCE(od.oms_source_amount,r.oms_amount) oms_amount,r.amount_difference,
      COALESCE(od.oms_source_qty,r.invoice_qty) mapped_qty,
      COALESCE(od.oms_source_amount,r.sap_amount) mapped_amount,
      CASE WHEN od.document_no IS NULL THEN 'SAP补数量金额'
           WHEN r.result='数量金额一致' THEN '双向字段一致'
           WHEN ABS(r.quantity_difference)<=0.000001 THEN '数量一致金额差异'
           ELSE '数量金额差异' END mapping_result,
      r.result source_result
    FROM oms_sap_recon r LEFT JOIN od
      ON od.document_no=r.oms_sales_no AND od.item_code=r.material_code AND od.sale_unit=r.sales_unit
    WHERE COALESCE(od.outbound_month,r.file_month) BETWEEN '2026-01' AND '2026-06'
    UNION ALL
    SELECT od.document_no,od.item_code,od.sale_unit,'',od.outbound_month,'',
      od.customer_code,od.customer_name,0,od.oms_source_rows,
      0,od.oms_source_qty,od.oms_source_qty,
      0,od.oms_source_amount,od.oms_source_amount,
      od.oms_source_qty,od.oms_source_amount,'仅OMS月结','仅OMS'
    FROM od LEFT JOIN oms_sap_recon r
      ON r.oms_sales_no=od.document_no AND r.material_code=od.item_code AND r.sales_unit=od.sale_unit
    WHERE r.oms_sales_no IS NULL;

    DROP TABLE IF EXISTS v4_oms_month_shop_sap;
    CREATE TABLE v4_oms_month_shop_sap AS
    SELECT outbound_month,customer_code,MIN(customer_name) customer_name,
      COUNT(DISTINCT oms_sales_no) mapped_docs,SUM(sap_qty) sap_qty,SUM(sap_amount) sap_amount,
      SUM(oms_qty) mapped_oms_qty,SUM(oms_amount) mapped_oms_amount,
      SUM(CASE WHEN mapping_result='双向字段一致' THEN 1 ELSE 0 END) exact_keys,COUNT(*) mapped_keys
    FROM v4_oms_sap_field_map
    WHERE customer_code<>'' GROUP BY 1,2;

    DROP TABLE IF EXISTS v4_bill_oms_month_recon;
    CREATE TABLE v4_bill_oms_month_recon AS
    SELECT h.bill_month,h.platform,h.huice_shop,h.customer_code,h.customer_name,h.mapping_status,h.mapping_source,
      h.bill_record_count,h.success_count,h.bill_success_amount,h.bill_receivable,h.bill_cash,
      COALESCE(o.oms_docs,0) oms_docs,COALESCE(o.oms_qty,0) oms_qty,COALESCE(o.oms_amount,0) oms_amount,
      COALESCE(s.sap_qty,0) sap_assisted_qty,COALESCE(s.sap_amount,0) sap_assisted_amount,
      COALESCE(o.oms_amount,0)-h.bill_success_amount success_difference,
      COALESCE(o.oms_amount,0)-h.bill_receivable receivable_difference,
      COALESCE(o.oms_amount,0)-h.bill_cash cash_difference,
      COALESCE(s.sap_amount,0)-h.bill_success_amount sap_success_difference,
      CASE WHEN h.customer_code IS NULL OR h.customer_code='' THEN '店铺未映射'
           WHEN o.customer_code IS NULL THEN '仅账单'
           WHEN ABS(o.oms_amount-h.bill_success_amount)<=0.01 THEN '成功金额一致'
           WHEN ABS(o.oms_amount-h.bill_receivable)<=0.01 THEN '应收金额一致'
           WHEN ABS(o.oms_amount-h.bill_cash)<=0.01 THEN '实收金额一致'
           WHEN ABS(COALESCE(s.sap_amount,0)-h.bill_success_amount)<=0.01 THEN 'SAP辅助金额一致'
           ELSE '金额差异' END result
    FROM v4_huice_shop_map h LEFT JOIN v4_oms_month_shop o
      ON o.outbound_month=h.bill_month AND o.customer_code=h.customer_code
    LEFT JOIN v4_oms_month_shop_sap s
      ON s.outbound_month=h.bill_month AND s.customer_code=h.customer_code
    UNION ALL
    SELECT o.outbound_month,'','',o.customer_code,o.customer_name,'OMS侧未对应','OMS客户',0,0,0,0,0,
      o.oms_docs,o.oms_qty,o.oms_amount,COALESCE(s.sap_qty,0),COALESCE(s.sap_amount,0),
      o.oms_amount,o.oms_amount,o.oms_amount,COALESCE(s.sap_amount,0),'仅OMS月结'
    FROM v4_oms_month_shop o LEFT JOIN v4_huice_shop_map h
      ON h.bill_month=o.outbound_month AND h.customer_code=o.customer_code
    LEFT JOIN v4_oms_month_shop_sap s
      ON s.outbound_month=o.outbound_month AND s.customer_code=o.customer_code
    WHERE h.huice_shop IS NULL;

    DROP TABLE IF EXISTS v4_billed_wdt_item;
    CREATE TABLE v4_billed_wdt_item AS
    SELECT substr(d.ship_time,1,7) ship_month,d.shop wdt_shop,i.material_code,i.unit,
      COUNT(DISTINCT d.platform_order_no) billed_orders,SUM(i.quantity) order_bill_qty,SUM(i.amount) wdt_item_amount,
      COUNT(*) item_groups
    FROM wdt_order_dedup d JOIN wdt_order_item i ON i.order_no=d.order_no
    WHERE d.ship_time>='{START}' AND d.ship_time<'{END_EXCLUSIVE}' AND d.platform_order_no<>''
      AND EXISTS (SELECT 1 FROM v4_huice_platform h WHERE h.platform_order_no=d.platform_order_no)
    GROUP BY 1,2,3,4;

    DROP TABLE IF EXISTS v4_order_bill_oms_qty_recon;
    CREATE TABLE v4_order_bill_oms_qty_recon AS
    WITH w AS (
      SELECT b.ship_month,b.wdt_shop,m.customer_code,m.customer_name,m.mapping_status,b.material_code,
        SUM(b.billed_orders) billed_orders,SUM(b.order_bill_qty) order_bill_qty,SUM(b.wdt_item_amount) wdt_item_amount
      FROM v4_billed_wdt_item b LEFT JOIN wdt_oms_shop_map m ON m.wdt_shop=b.wdt_shop
      GROUP BY 1,2,3,4,5,6
    ), o AS (
      SELECT outbound_month,customer_code,MIN(customer_name) customer_name,item_code,
        SUM(item_num) oms_qty,SUM(share_amount) oms_amount,COUNT(DISTINCT document_no) oms_docs
      FROM oms_detail WHERE business_type='Y001' AND outbound_time>='{START}' AND outbound_time<'{END_EXCLUSIVE}'
      GROUP BY 1,2,4
    )
    SELECT w.ship_month,w.wdt_shop,w.customer_code,w.customer_name,w.mapping_status,w.material_code,
      w.billed_orders,w.order_bill_qty,w.wdt_item_amount,COALESCE(o.oms_qty,0) oms_qty,COALESCE(o.oms_amount,0) oms_amount,
      COALESCE(o.oms_docs,0) oms_docs,COALESCE(o.oms_qty,0)-w.order_bill_qty qty_difference,
      CASE WHEN w.customer_code IS NULL THEN '店铺未映射'
           WHEN o.item_code IS NULL THEN '仅订单账单数量'
           WHEN ABS(o.oms_qty-w.order_bill_qty)<=0.000001 THEN '数量一致' ELSE '数量差异' END result
    FROM w LEFT JOIN o ON o.outbound_month=w.ship_month AND o.customer_code=w.customer_code
      AND o.item_code=w.material_code
    UNION ALL
    SELECT o.outbound_month,m.wdt_shop,o.customer_code,o.customer_name,m.mapping_status,o.item_code,
      0,0,0,o.oms_qty,o.oms_amount,o.oms_docs,o.oms_qty,'仅OMS月结数量'
    FROM o LEFT JOIN wdt_oms_shop_map m ON m.customer_code=o.customer_code
    LEFT JOIN w ON w.ship_month=o.outbound_month AND w.customer_code=o.customer_code
      AND w.material_code=o.item_code
    WHERE w.material_code IS NULL;
    """)
    conn.commit()


def rows(conn: sqlite3.Connection, query: str):
    cur = conn.execute(query)
    headers = [x[0] for x in cur.description]
    return headers, cur.fetchall()


def dict_rows(conn: sqlite3.Connection, query: str):
    headers, values = rows(conn, query)
    return [dict(zip(headers, value)) for value in values]


def scalar(conn: sqlite3.Connection, query: str):
    value = conn.execute(query).fetchone()[0]
    return value or 0


def export(conn: sqlite3.Connection) -> dict:
    OUT.mkdir(parents=True, exist_ok=True)
    queries = {
        "order_bill_recon": "SELECT * FROM v4_order_bill_recon ORDER BY CASE result WHEN '单号应收金额一致' THEN 3 WHEN '单号实收金额一致' THEN 3 ELSE 1 END,result,platform_order_no",
        "bill_oms_month_recon": "SELECT * FROM v4_bill_oms_month_recon ORDER BY CASE result WHEN '成功金额一致' THEN 3 WHEN '应收金额一致' THEN 3 WHEN '实收金额一致' THEN 3 WHEN 'SAP辅助金额一致' THEN 3 ELSE 1 END,result,bill_month,huice_shop",
        "order_bill_oms_qty_recon": "SELECT * FROM v4_order_bill_oms_qty_recon ORDER BY CASE result WHEN '数量一致' THEN 3 ELSE 1 END,result,ship_month,wdt_shop,material_code",
        "oms_sap_field_map": "SELECT * FROM v4_oms_sap_field_map ORDER BY CASE mapping_result WHEN '双向字段一致' THEN 2 ELSE 1 END,mapping_result,outbound_month,oms_sales_no,material_code",
        "huice_shop_map": "SELECT * FROM v4_huice_shop_map ORDER BY mapping_status,bill_month,platform,huice_shop",
    }
    limits = {"order_bill_recon":15000,"bill_oms_month_recon":20000,"order_bill_oms_qty_recon":20000,"oms_sap_field_map":20000,"huice_shop_map":20000}
    for name, query in queries.items():
        headers, values = rows(conn, query)
        with (OUT/f"{name}.csv").open("w",encoding="utf-8-sig",newline="") as f:
            writer=csv.writer(f);writer.writerow(headers);writer.writerows(values)
        sample=values[:limits[name]]
        (OUT/f"{name}_workbook.json").write_text(json.dumps({"headers":headers,"rows":sample},ensure_ascii=False),encoding="utf-8")

    summary = {
        "period_start": START,
        "period_end": "2026-06-30",
        "order_bill_results": dict_rows(conn,"SELECT result,COUNT(*) groups,SUM(wdt_qty) wdt_qty,SUM(wdt_amount) wdt_amount,SUM(bill_receivable) bill_receivable,SUM(bill_cash) bill_cash FROM v4_order_bill_recon GROUP BY result ORDER BY groups DESC"),
        "bill_oms_results": dict_rows(conn,"SELECT result,COUNT(*) groups,SUM(bill_record_count) bill_records,SUM(bill_success_amount) bill_success_amount,SUM(bill_receivable) bill_receivable,SUM(bill_cash) bill_cash,SUM(oms_qty) oms_qty,SUM(oms_amount) oms_amount,SUM(sap_assisted_qty) sap_assisted_qty,SUM(sap_assisted_amount) sap_assisted_amount FROM v4_bill_oms_month_recon GROUP BY result ORDER BY groups DESC"),
        "qty_results": dict_rows(conn,"SELECT result,COUNT(*) groups,SUM(billed_orders) billed_orders,SUM(order_bill_qty) order_bill_qty,SUM(oms_qty) oms_qty,SUM(qty_difference) qty_difference FROM v4_order_bill_oms_qty_recon GROUP BY result ORDER BY groups DESC"),
        "oms_sap_results": dict_rows(conn,"SELECT mapping_result,COUNT(*) keys,SUM(sap_qty) sap_qty,SUM(oms_qty) oms_qty,SUM(sap_amount) sap_amount,SUM(oms_amount) oms_amount FROM v4_oms_sap_field_map GROUP BY mapping_result ORDER BY keys DESC"),
        "monthly_flow": dict_rows(conn,"""
          WITH w AS (SELECT ship_month month,SUM(quantity) wdt_qty,SUM(amount) wdt_amount FROM wdt_item WHERE ship_month BETWEEN '2026-01' AND '2026-06' GROUP BY 1),
          h AS (SELECT bill_month month,SUM(bill_record_count) bill_records,SUM(bill_success_amount) bill_success_amount,SUM(bill_receivable) bill_receivable,SUM(bill_cash) bill_cash FROM v4_huice_shop_bill GROUP BY 1),
          o AS (SELECT outbound_month month,SUM(oms_qty) oms_qty,SUM(oms_amount) oms_amount FROM v4_oms_month_shop GROUP BY 1),
          s AS (SELECT outbound_month month,SUM(sap_qty) sap_qty,SUM(sap_amount) sap_amount FROM v4_oms_month_shop_sap GROUP BY 1),
          sf AS (SELECT file_month month,SUM(invoice_qty) sap_full_qty,SUM(tax_amount) sap_full_amount FROM sap2c
                 WHERE file_month BETWEEN '2026-01' AND '2026-06' GROUP BY 1)
          SELECT w.month,w.wdt_qty,w.wdt_amount,h.bill_records,h.bill_success_amount,h.bill_receivable,h.bill_cash,
            o.oms_qty,o.oms_amount,s.sap_qty,s.sap_amount,sf.sap_full_qty,sf.sap_full_amount
          FROM w LEFT JOIN h ON h.month=w.month LEFT JOIN o ON o.month=w.month
          LEFT JOIN s ON s.month=w.month LEFT JOIN sf ON sf.month=w.month ORDER BY w.month
        """),
        "controls": {
            "wdt_orders": scalar(conn,"SELECT COUNT(*) FROM v4_wdt_platform"),
            "wdt_qty": scalar(conn,"SELECT SUM(quantity) FROM wdt_item WHERE ship_month BETWEEN '2026-01' AND '2026-06'"),
            "wdt_amount": scalar(conn,"SELECT SUM(amount) FROM wdt_item WHERE ship_month BETWEEN '2026-01' AND '2026-06'"),
            "huice_orders": scalar(conn,"SELECT COUNT(*) FROM v4_huice_platform"),
            "huice_bill_records": scalar(conn,"SELECT SUM(bill_record_count) FROM v4_huice_shop_bill"),
            "huice_bill_success_amount": scalar(conn,"SELECT SUM(bill_success_amount) FROM v4_huice_shop_bill"),
            "huice_bill_receivable": scalar(conn,"SELECT SUM(bill_receivable) FROM v4_huice_shop_bill"),
            "huice_bill_cash": scalar(conn,"SELECT SUM(bill_cash) FROM v4_huice_shop_bill"),
            "oms_month_docs": scalar(conn,"SELECT SUM(oms_docs) FROM v4_oms_month_shop"),
            "oms_month_qty": scalar(conn,"SELECT SUM(oms_qty) FROM v4_oms_month_shop"),
            "oms_month_amount": scalar(conn,"SELECT SUM(oms_amount) FROM v4_oms_month_shop"),
            "sap_assisted_qty": scalar(conn,"SELECT SUM(sap_qty) FROM v4_oms_month_shop_sap"),
            "sap_assisted_amount": scalar(conn,"SELECT SUM(sap_amount) FROM v4_oms_month_shop_sap"),
            "sap_full_qty": scalar(conn,"SELECT SUM(invoice_qty) FROM sap2c WHERE file_month BETWEEN '2026-01' AND '2026-06'"),
            "sap_full_amount": scalar(conn,"SELECT SUM(tax_amount) FROM sap2c WHERE file_month BETWEEN '2026-01' AND '2026-06'"),
            "billed_wdt_qty": scalar(conn,"SELECT SUM(order_bill_qty) FROM v4_billed_wdt_item"),
            "mapped_bill_amount": scalar(conn,"SELECT SUM(bill_receivable) FROM v4_huice_shop_map WHERE customer_code<>''"),
            "unmapped_bill_amount": scalar(conn,"SELECT SUM(bill_receivable) FROM v4_huice_shop_map WHERE customer_code IS NULL OR customer_code=''"),
        },
        "display_metrics": {
            "order_key_matches": scalar(conn,"SELECT COUNT(*) FROM v4_order_bill_recon WHERE result IN ('单号应收金额一致','单号实收金额一致','单号一致金额差异')"),
            "order_amount_exact_groups": scalar(conn,"SELECT COUNT(*) FROM v4_order_bill_recon WHERE result IN ('单号应收金额一致','单号实收金额一致')"),
            "bill_exact_groups": scalar(conn,"SELECT COUNT(*) FROM v4_bill_oms_month_recon WHERE result='成功金额一致'"),
            "bill_total_groups": scalar(conn,"SELECT COUNT(*) FROM v4_bill_oms_month_recon"),
            "bill_gross_abs_success_diff": scalar(conn,"SELECT SUM(ABS(success_difference)) FROM v4_bill_oms_month_recon"),
            "bill_exact_success_amount": scalar(conn,"SELECT SUM(ABS(bill_success_amount)) FROM v4_bill_oms_month_recon WHERE result='成功金额一致'"),
            "mapping_high_ar": scalar(conn,"SELECT SUM(bill_receivable) FROM v4_huice_shop_map WHERE mapping_status='高置信'"),
            "mapping_review_ar": scalar(conn,"SELECT SUM(bill_receivable) FROM v4_huice_shop_map WHERE mapping_status='待复核'"),
            "mapping_low_ar": scalar(conn,"SELECT SUM(bill_receivable) FROM v4_huice_shop_map WHERE mapping_status='低置信'"),
            "mapping_unmapped_ar": scalar(conn,"SELECT SUM(bill_receivable) FROM v4_huice_shop_map WHERE mapping_status='未映射'"),
            "qty_common_groups": scalar(conn,"SELECT COUNT(*) FROM v4_order_bill_oms_qty_recon WHERE result IN ('数量一致','数量差异')"),
            "qty_exact_groups": scalar(conn,"SELECT COUNT(*) FROM v4_order_bill_oms_qty_recon WHERE result='数量一致'"),
            "qty_gross_abs_diff": scalar(conn,"SELECT SUM(ABS(qty_difference)) FROM v4_order_bill_oms_qty_recon WHERE result IN ('数量一致','数量差异')"),
            "qty_exact_amount": scalar(conn,"SELECT SUM(oms_qty) FROM v4_order_bill_oms_qty_recon WHERE result='数量一致'"),
        },
        "detail_rows": {name: scalar(conn,f"SELECT COUNT(*) FROM v4_{'order_bill_recon' if name=='order_bill_recon' else 'bill_oms_month_recon' if name=='bill_oms_month_recon' else 'order_bill_oms_qty_recon' if name=='order_bill_oms_qty_recon' else 'oms_sap_field_map' if name=='oms_sap_field_map' else 'huice_shop_map'}") for name in queries},
    }
    (OUT/"summary_v4.json").write_text(json.dumps(summary,ensure_ascii=False,indent=2),encoding="utf-8")
    return summary


def main() -> None:
    global OUT
    parser = argparse.ArgumentParser(description="生成销售ToC流程核对结果")
    parser.add_argument("--output-dir", type=Path, default=OUT, help="CSV/JSON输出目录")
    args = parser.parse_args()
    OUT = args.output_dir.resolve()
    conn=sqlite3.connect(DB);configure(conn);build(conn);summary=export(conn)
    print(json.dumps(summary,ensure_ascii=False,indent=2))


if __name__=="__main__":
    main()
