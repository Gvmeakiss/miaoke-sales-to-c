#!/usr/bin/env python3
"""以惠策实际账单为主线，借助旺店通订单行补充物料后与 OMS 日结核对。

原始惠策明细只有平台订单号、店铺和金额，不含物料与数量；旺店通仅作为
订单号->物料的维表使用，不再作为 OMS 金额核对的主数据源。
"""

from __future__ import annotations

import argparse
import csv
import importlib.util
import json
import sqlite3
from collections import defaultdict
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path
from typing import Dict, Optional, Tuple


ROOT = Path(__file__).resolve().parent
SPEC = importlib.util.spec_from_file_location("recon_base", ROOT / "reconcile_sales_toc.py")
BASE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader
SPEC.loader.exec_module(BASE)
HUICE_SUPPLEMENT_FILE = BASE.HUICE_SUPPLEMENT_FILE


def extract_order_items(task: Tuple[str, str, Optional[int]]) -> Tuple[str, str, int]:
    path_text, cache_text, max_rows = task
    path, cache = Path(path_text), Path(cache_text)
    if cache.exists() and cache.stat().st_size:
        return path.name, str(cache), sum(1 for _ in cache.open("r", encoding="utf-8"))
    header = BASE.read_header(path)
    fields = {
        "order_no": BASE.header_index(header, "订单编号"),
        "platform_order_no": BASE.header_index(header, "原始单号"),
        "sap_code": BASE.header_index(header, "SAP编码"),
        "product_code": BASE.header_index(header, "货品编号"),
        "unit": BASE.header_index(header, "单位"),
        "quantity": BASE.header_index(header, "数量"),
        "amount": BASE.header_index(header, "分摊后总价"),
    }
    selected = sorted({c for c in fields.values() if c})
    cache.parent.mkdir(parents=True, exist_ok=True)
    current_order = ""
    aggregates: Dict[Tuple[str, str, str, str], list[float]] = defaultdict(lambda: [0.0, 0.0, 0.0])
    emitted = 0

    def value(values, key):
        col = fields.get(key)
        return BASE.text(values.get(col)) if col else ""

    with cache.open("w", encoding="utf-8", newline="") as out:
        writer = csv.writer(out)

        def emit():
            nonlocal emitted
            for (order_no, platform_order_no, material_key, unit), nums in aggregates.items():
                writer.writerow([order_no, platform_order_no, material_key, unit, nums[0], nums[1], int(nums[2])])
                emitted += 1
            aggregates.clear()

        for _, values in BASE.iter_selected_rows(path, selected, max_rows=max_rows):
            order_no = value(values, "order_no")
            if not order_no or order_no == "订单编号":
                continue
            if current_order and order_no != current_order:
                emit()
            current_order = order_no
            sap = value(values, "sap_code")
            product = value(values, "product_code")
            material = sap or ("SKU:" + product if product else "")
            if not material:
                continue
            key = (order_no, value(values, "platform_order_no"), material, value(values, "unit"))
            aggregates[key][0] += BASE.as_number(value(values, "quantity"))
            aggregates[key][1] += BASE.as_number(value(values, "amount"))
            aggregates[key][2] += 1
        emit()
    return path.name, str(cache), emitted


def load_order_items(conn: sqlite3.Connection, input_dir: Path, work_dir: Path, workers: int, max_rows: Optional[int]) -> None:
    conn.executescript("""
    DROP TABLE IF EXISTS wdt_order_item_stage;
    CREATE TABLE wdt_order_item_stage(source_file TEXT,order_no TEXT,platform_order_no TEXT,material_key TEXT,unit TEXT,quantity REAL,amount REAL,line_count INTEGER);
    """)
    files = BASE.wdt_files(input_dir)
    cache_dir = work_dir / "wdt_order_item_platform_cache"
    tasks = [(str(p), str(cache_dir / f"{i:02d}_{p.stem}.csv"), max_rows) for i, p in enumerate(files, 1)]
    results = []
    with ProcessPoolExecutor(max_workers=max(1, workers)) as pool:
        futures = [pool.submit(extract_order_items, t) for t in tasks]
        for i, future in enumerate(as_completed(futures), 1):
            result = future.result(); results.append(result)
            print(f"订单物料抽取 {i}/{len(tasks)} {result[0]}：{result[2]:,}行", flush=True)
    for name, cache, _ in sorted(results):
        batch = []
        with Path(cache).open("r", encoding="utf-8", newline="") as stream:
            for row in csv.reader(stream):
                batch.append((name,row[0],row[1],row[2],row[3],float(row[4]),float(row[5]),int(row[6])))
                if len(batch) >= 50000:
                    conn.executemany("INSERT INTO wdt_order_item_stage VALUES(?,?,?,?,?,?,?,?)", batch); conn.commit(); batch.clear()
            if batch:
                conn.executemany("INSERT INTO wdt_order_item_stage VALUES(?,?,?,?,?,?,?,?)", batch); conn.commit()
        print(f"订单物料入库 {name}", flush=True)
    conn.executescript("""
    DROP TABLE IF EXISTS wdt_order_item;
    CREATE TABLE wdt_order_item AS
    WITH chosen_source AS (
      SELECT source_file,order_no,platform_order_no,
        ROW_NUMBER() OVER (
          PARTITION BY order_no,platform_order_no
          ORDER BY quantity DESC,line_count DESC,ship_time DESC,source_file DESC
        ) source_rank
      FROM wdt_orders_file
    )
    SELECT s.order_no,s.platform_order_no,
      CASE WHEN s.material_key LIKE 'SKU:%' THEN COALESCE(m.sap_code,s.material_key) ELSE s.material_key END material_code,
      s.unit,SUM(s.quantity) quantity,SUM(s.amount) amount,SUM(s.line_count) line_count
    FROM wdt_order_item_stage s
    JOIN chosen_source c ON c.source_file=s.source_file AND c.order_no=s.order_no
      AND c.platform_order_no=s.platform_order_no AND c.source_rank=1
    LEFT JOIN wdt_product_sap_best m ON substr(s.material_key,5)=m.product_code
    GROUP BY 1,2,3,4;
    CREATE INDEX idx_wdt_order_item_order ON wdt_order_item(order_no,platform_order_no);
    CREATE INDEX idx_wdt_order_item_platform ON wdt_order_item(platform_order_no);
    """)
    conn.commit()


def extract_huice_current_file(task: Tuple[str, str, Optional[int]]) -> Tuple[str, str, int]:
    path_text, cache_text, max_rows = task
    path, cache = Path(path_text), Path(cache_text)
    if cache.exists() and cache.stat().st_size:
        return path.name, str(cache), sum(1 for _ in cache.open("r", encoding="utf-8"))
    header = BASE.read_header(path)
    names = ["对账流水号", "正应收金额", "负应收金额", "收款金额（正实收）", "退款金额（负实收）"]
    columns = {name: BASE.header_index(header, name) for name in names}
    selected = sorted({c for c in columns.values() if c})
    cache.parent.mkdir(parents=True, exist_ok=True)
    count = 0
    with cache.open("w", encoding="utf-8", newline="") as out:
        writer = csv.writer(out)
        for _, values in BASE.iter_selected_rows(path, selected, max_rows=max_rows):
            def val(name):
                col = columns.get(name)
                return BASE.text(values.get(col)) if col else ""
            rid = val("对账流水号")
            if not rid or rid == "对账流水号":
                continue
            writer.writerow([rid,
                BASE.as_number(val("正应收金额"))-BASE.as_number(val("负应收金额")),
                BASE.as_number(val("收款金额（正实收）"))-BASE.as_number(val("退款金额（负实收）"))])
            count += 1
    return path.name, str(cache), count


def load_huice_current(conn: sqlite3.Connection, input_dir: Path, work_dir: Path, workers: int, max_rows: Optional[int]) -> None:
    conn.executescript("DROP TABLE IF EXISTS huice_current_amount; CREATE TABLE huice_current_amount(reconcile_id TEXT PRIMARY KEY,current_receivable REAL,current_cash REAL);")
    files = BASE.huice_detail_files(input_dir)
    cache_dir = work_dir / "huice_current_cache"
    tasks = [(str(p), str(cache_dir / BASE.stable_cache_name(p)), max_rows) for p in files]
    results = []
    with ProcessPoolExecutor(max_workers=max(1, workers)) as pool:
        futures = [pool.submit(extract_huice_current_file, t) for t in tasks]
        for i, future in enumerate(as_completed(futures), 1):
            result = future.result(); results.append(result)
            print(f"惠策本期金额抽取 {i}/{len(tasks)} {result[0]}：{result[2]:,}行", flush=True)
    for name, cache, _ in sorted(results, key=lambda item: (item[0] == HUICE_SUPPLEMENT_FILE, item[0])):
        batch = []
        with Path(cache).open("r", encoding="utf-8", newline="") as stream:
            for row in csv.reader(stream):
                batch.append((row[0],float(row[1]),float(row[2])))
                if len(batch)>=50000:
                    conn.executemany("INSERT OR IGNORE INTO huice_current_amount VALUES(?,?,?)",batch); conn.commit(); batch.clear()
            if batch:
                conn.executemany("INSERT OR IGNORE INTO huice_current_amount VALUES(?,?,?)",batch); conn.commit()
        print(f"惠策本期金额入库 {name}",flush=True)


def build_reconciliation(conn: sqlite3.Connection) -> dict:
    conn.executescript(f"""
    DROP TABLE IF EXISTS huice_order_month;
    CREATE TABLE huice_order_month AS
    SELECT CASE WHEN h.source_file='{HUICE_SUPPLEMENT_FILE}' THEN '2026-06'
                ELSE substr(COALESCE(NULLIF(business_date,''),period_end),1,7) END huice_month,
      platform,shop,platform_order_no,
      SUM(c.current_receivable) huice_receivable,SUM(c.current_cash) huice_cash,COUNT(*) huice_rows
    FROM huice_detail h JOIN huice_current_amount c ON c.reconcile_id=h.reconcile_id
    WHERE platform_order_no<>'' AND (
      h.source_file='{HUICE_SUPPLEMENT_FILE}' OR
      substr(COALESCE(NULLIF(business_date,''),period_end),1,7) BETWEEN '2025-12' AND '2026-06'
    )
    GROUP BY 1,2,3,4;
    CREATE INDEX idx_huice_order_month_order ON huice_order_month(platform_order_no);

    DROP TABLE IF EXISTS wdt_platform_order_item;
    CREATE TABLE wdt_platform_order_item AS
    SELECT d.platform_order_no,substr(d.ship_time,1,7) ship_month,d.shop,i.material_code,i.unit,
      SUM(i.quantity) wdt_qty,SUM(i.amount) wdt_amount,SUM(i.line_count) wdt_lines
    FROM wdt_order_dedup d JOIN wdt_order_item i
      ON i.order_no=d.order_no AND i.platform_order_no=d.platform_order_no
    WHERE d.platform_order_no<>'' AND substr(d.ship_time,1,7) BETWEEN '2025-12' AND '2026-06'
    GROUP BY 1,2,3,4,5;
    CREATE INDEX idx_wdt_platform_order_item_order ON wdt_platform_order_item(platform_order_no);

    DROP TABLE IF EXISTS huice_wdt_shop_bridge;
    CREATE TABLE huice_wdt_shop_bridge AS
    WITH x AS (
      SELECT h.platform,h.shop huice_shop,w.shop wdt_shop,COUNT(DISTINCT h.platform_order_no) matched_orders,
        SUM(ABS(h.huice_receivable)) matched_receivable
      FROM huice_order_month h JOIN wdt_platform_order_item w ON w.platform_order_no=h.platform_order_no
      GROUP BY 1,2,3
    ), r AS (
      SELECT *,ROW_NUMBER() OVER(PARTITION BY platform,huice_shop ORDER BY matched_orders DESC,matched_receivable DESC,wdt_shop) rn,
        SUM(matched_orders) OVER(PARTITION BY platform,huice_shop) total_orders
      FROM x
    )
    SELECT platform,huice_shop,wdt_shop,matched_orders,total_orders,
      1.0*matched_orders/NULLIF(total_orders,0) bridge_share,
      CASE WHEN 1.0*matched_orders/NULLIF(total_orders,0)>=0.90 THEN '高置信'
           WHEN 1.0*matched_orders/NULLIF(total_orders,0)>=0.70 THEN '待复核' ELSE '低置信' END bridge_status
    FROM r WHERE rn=1;

    DROP TABLE IF EXISTS huice_oms_shop_map;
    CREATE TABLE huice_oms_shop_map AS
    SELECT b.platform,b.huice_shop,b.wdt_shop,m.customer_code,m.customer_name,m.total_score oms_map_score,
      b.bridge_share,b.bridge_status,m.mapping_status oms_mapping_status,
      CASE WHEN b.bridge_status='高置信' AND m.mapping_status='高置信' THEN '高置信'
           WHEN m.customer_code IS NOT NULL AND b.bridge_status<>'低置信' AND m.mapping_status<>'低置信' THEN '待复核'
           ELSE '低置信' END final_status
    FROM huice_wdt_shop_bridge b
    LEFT JOIN wdt_oms_shop_map m ON m.wdt_shop=b.wdt_shop;

    DROP TABLE IF EXISTS huice_order_item_alloc;
    CREATE TABLE huice_order_item_alloc AS
    WITH joined AS (
      SELECT h.huice_month,h.platform,h.shop huice_shop,h.platform_order_no,h.huice_receivable,h.huice_cash,h.huice_rows,
        w.ship_month,w.shop wdt_shop,w.material_code,w.unit,w.wdt_qty,w.wdt_amount,w.wdt_lines,
        SUM(ABS(w.wdt_amount)) OVER(PARTITION BY h.huice_month,h.platform,h.shop,h.platform_order_no) abs_amount_total,
        SUM(ABS(w.wdt_qty)) OVER(PARTITION BY h.huice_month,h.platform,h.shop,h.platform_order_no) abs_qty_total,
        COUNT(*) OVER(PARTITION BY h.huice_month,h.platform,h.shop,h.platform_order_no) item_groups
      FROM huice_order_month h JOIN wdt_platform_order_item w ON w.platform_order_no=h.platform_order_no
    )
    SELECT *,
      CASE WHEN abs_amount_total>0 THEN ABS(wdt_amount)/abs_amount_total
           WHEN abs_qty_total>0 THEN ABS(wdt_qty)/abs_qty_total ELSE 1.0/item_groups END alloc_rate,
      huice_receivable*CASE WHEN abs_amount_total>0 THEN ABS(wdt_amount)/abs_amount_total
           WHEN abs_qty_total>0 THEN ABS(wdt_qty)/abs_qty_total ELSE 1.0/item_groups END allocated_receivable,
      huice_cash*CASE WHEN abs_amount_total>0 THEN ABS(wdt_amount)/abs_amount_total
           WHEN abs_qty_total>0 THEN ABS(wdt_qty)/abs_qty_total ELSE 1.0/item_groups END allocated_cash
    FROM joined;
    CREATE INDEX idx_huice_order_item_alloc_key ON huice_order_item_alloc(huice_month,huice_shop,material_code);

    DROP TABLE IF EXISTS huice_shop_material_month;
    CREATE TABLE huice_shop_material_month AS
    SELECT a.huice_month,a.platform,a.huice_shop,a.material_code,
      MIN(m.customer_code) customer_code,MIN(m.customer_name) customer_name,MIN(m.final_status) mapping_status,
      SUM(a.allocated_receivable) huice_receivable,SUM(a.allocated_cash) huice_cash,
      SUM(a.wdt_qty) bridge_qty,COUNT(DISTINCT a.platform_order_no) mapped_orders,SUM(a.huice_rows) huice_rows
    FROM huice_order_item_alloc a
    LEFT JOIN huice_oms_shop_map m ON m.platform=a.platform AND m.huice_shop=a.huice_shop AND m.wdt_shop=a.wdt_shop
    GROUP BY 1,2,3,4;

    DROP TABLE IF EXISTS huice_oms_month_item_recon;
    CREATE TABLE huice_oms_month_item_recon AS
    WITH h AS (
      SELECT huice_month,platform,huice_shop,customer_code,customer_name,mapping_status,material_code,
        SUM(huice_receivable) huice_receivable,SUM(huice_cash) huice_cash,SUM(bridge_qty) bridge_qty,
        SUM(mapped_orders) mapped_orders,SUM(huice_rows) huice_rows
      FROM huice_shop_material_month GROUP BY 1,2,3,4,5,6,7
    ), o AS (
      SELECT outbound_month,customer_code,MIN(customer_name) customer_name,item_code,
        SUM(item_num) oms_qty,SUM(share_amount) oms_amount,COUNT(DISTINCT document_no) oms_docs
      FROM oms_detail WHERE business_type='Y005' GROUP BY 1,2,4
    ), l AS (
      SELECT h.huice_month,h.platform,h.huice_shop,h.customer_code,h.customer_name,h.mapping_status,h.material_code,
        h.huice_receivable,h.huice_cash,h.bridge_qty,h.mapped_orders,h.huice_rows,
        COALESCE(o.oms_qty,0) oms_qty,COALESCE(o.oms_amount,0) oms_amount,COALESCE(o.oms_docs,0) oms_docs,
        COALESCE(o.oms_amount,0)-h.huice_receivable receivable_difference,
        COALESCE(o.oms_amount,0)-h.huice_cash cash_difference,
        CASE WHEN o.item_code IS NULL THEN '仅惠策'
             WHEN ABS(o.oms_amount-h.huice_receivable)<=0.01 THEN '应收金额一致'
             WHEN ABS(o.oms_amount-h.huice_cash)<=0.01 THEN '实收金额一致'
             ELSE '金额差异' END result
      FROM h LEFT JOIN o ON o.outbound_month=h.huice_month AND o.customer_code=h.customer_code AND o.item_code=h.material_code
    ) SELECT * FROM l
    UNION ALL
    SELECT o.outbound_month,'','',o.customer_code,o.customer_name,'未映射OMS',o.item_code,
      0,0,0,0,0,o.oms_qty,o.oms_amount,o.oms_docs,o.oms_amount,o.oms_amount,'仅OMS'
    FROM o LEFT JOIN h ON h.huice_month=o.outbound_month AND h.customer_code=o.customer_code AND h.material_code=o.item_code
    WHERE h.material_code IS NULL;

    DROP TABLE IF EXISTS huice_oms_month_shop_recon;
    CREATE TABLE huice_oms_month_shop_recon AS
    SELECT huice_month,platform,huice_shop,customer_code,customer_name,mapping_status,
      SUM(huice_receivable) huice_receivable,SUM(huice_cash) huice_cash,SUM(bridge_qty) bridge_qty,
      SUM(oms_qty) oms_qty,SUM(oms_amount) oms_amount,
      SUM(receivable_difference) receivable_difference,SUM(cash_difference) cash_difference,
      SUM(mapped_orders) mapped_orders,SUM(huice_rows) huice_rows,SUM(oms_docs) oms_docs,
      CASE WHEN ABS(SUM(receivable_difference))<=0.01 THEN '店铺月度应收一致'
           WHEN ABS(SUM(cash_difference))<=0.01 THEN '店铺月度实收一致' ELSE '店铺月度差异' END result
    FROM huice_oms_month_item_recon GROUP BY 1,2,3,4,5,6;

    DROP TABLE IF EXISTS huice_detail_summary_month_shop_recon_v3;
    CREATE TABLE huice_detail_summary_month_shop_recon_v3 AS
    WITH d AS (
      SELECT substr(h.period_end,1,7) month,h.platform,h.shop,
        COUNT(*) detail_rows,SUM(h.net_receivable) detail_receivable,SUM(h.net_cash) detail_cash
      FROM huice_detail h GROUP BY 1,2,3
    ), s AS (
      SELECT substr(reconcile_date,1,7) month,platform,shop,
        SUM(success_count) success_count,SUM(success_amount) success_amount,
        SUM(mismatch_count) mismatch_count,SUM(mismatch_receivable) mismatch_receivable,SUM(mismatch_cash) mismatch_cash,
        SUM(single_ar_count) single_ar_count,SUM(single_ar_amount) single_ar_amount,
        SUM(single_cash_count) single_cash_count,SUM(single_cash_amount) single_cash_amount
      FROM huice_summary GROUP BY 1,2,3
    )
    SELECT COALESCE(d.month,s.month) month,COALESCE(d.platform,s.platform) platform,COALESCE(d.shop,s.shop) shop,
      COALESCE(d.detail_rows,0) detail_rows,COALESCE(d.detail_receivable,0) detail_receivable,COALESCE(d.detail_cash,0) detail_cash,
      COALESCE(s.success_count,0) success_count,COALESCE(s.success_amount,0) success_amount,
      COALESCE(s.mismatch_count,0) mismatch_count,COALESCE(s.mismatch_receivable,0) mismatch_receivable,COALESCE(s.mismatch_cash,0) mismatch_cash,
      COALESCE(s.single_ar_count,0) single_ar_count,COALESCE(s.single_ar_amount,0) single_ar_amount,
      COALESCE(s.single_cash_count,0) single_cash_count,COALESCE(s.single_cash_amount,0) single_cash_amount,
      COALESCE(d.detail_receivable,0)-(COALESCE(s.success_amount,0)+COALESCE(s.mismatch_receivable,0)+COALESCE(s.single_ar_amount,0)) receivable_difference,
      COALESCE(d.detail_cash,0)-(COALESCE(s.success_amount,0)+COALESCE(s.mismatch_cash,0)+COALESCE(s.single_cash_amount,0)) cash_difference,
      CASE WHEN d.shop IS NULL THEN '仅店铺汇总' WHEN s.shop IS NULL THEN '仅明细重建'
           WHEN ABS(COALESCE(d.detail_receivable,0)-(COALESCE(s.success_amount,0)+COALESCE(s.mismatch_receivable,0)+COALESCE(s.single_ar_amount,0)))<=0.01
            AND ABS(COALESCE(d.detail_cash,0)-(COALESCE(s.success_amount,0)+COALESCE(s.mismatch_cash,0)+COALESCE(s.single_cash_amount,0)))<=0.01
           THEN '金额一致' ELSE '金额差异' END result
    FROM d LEFT JOIN s ON s.month=d.month AND s.platform=d.platform AND s.shop=d.shop
    UNION ALL
    SELECT s.month,s.platform,s.shop,0,0,0,s.success_count,s.success_amount,s.mismatch_count,s.mismatch_receivable,s.mismatch_cash,
      s.single_ar_count,s.single_ar_amount,s.single_cash_count,s.single_cash_amount,
      -(s.success_amount+s.mismatch_receivable+s.single_ar_amount),-(s.success_amount+s.mismatch_cash+s.single_cash_amount),'仅店铺汇总'
    FROM s LEFT JOIN d ON d.month=s.month AND d.platform=s.platform AND d.shop=s.shop WHERE d.shop IS NULL;
    """)
    conn.commit()

    def scalar(sql): return conn.execute(sql).fetchone()[0] or 0
    stats = {
        "huice_raw_rows": scalar("SELECT COUNT(*) FROM huice_detail"),
        "huice_order_months": scalar("SELECT COUNT(*) FROM huice_order_month"),
        "mapped_order_months": scalar("SELECT COUNT(DISTINCT huice_month||'|'||platform||'|'||huice_shop||'|'||platform_order_no) FROM huice_order_item_alloc"),
        "huice_receivable": scalar("SELECT SUM(huice_receivable) FROM huice_order_month"),
        "mapped_receivable": scalar("SELECT SUM(huice_receivable) FROM (SELECT huice_month,platform,huice_shop,platform_order_no,MAX(huice_receivable) huice_receivable FROM huice_order_item_alloc GROUP BY 1,2,3,4)"),
        "huice_cash": scalar("SELECT SUM(huice_cash) FROM huice_order_month"),
        "mapped_cash": scalar("SELECT SUM(huice_cash) FROM (SELECT huice_month,platform,huice_shop,platform_order_no,MAX(huice_cash) huice_cash FROM huice_order_item_alloc GROUP BY 1,2,3,4)"),
        "shop_material_groups": scalar("SELECT COUNT(*) FROM huice_shop_material_month"),
        "recon_groups": scalar("SELECT COUNT(*) FROM huice_oms_month_item_recon"),
        "receivable_exact_groups": scalar("SELECT COUNT(*) FROM huice_oms_month_item_recon WHERE result='应收金额一致'"),
        "cash_exact_groups": scalar("SELECT COUNT(*) FROM huice_oms_month_item_recon WHERE result='实收金额一致'"),
        "high_confidence_shops": scalar("SELECT COUNT(*) FROM huice_oms_shop_map WHERE final_status='高置信'"),
        "review_shops": scalar("SELECT COUNT(*) FROM huice_oms_shop_map WHERE final_status='待复核'"),
        "low_confidence_shops": scalar("SELECT COUNT(*) FROM huice_oms_shop_map WHERE final_status='低置信'"),
    }
    stats["order_mapping_rate"] = stats["mapped_order_months"] / stats["huice_order_months"] if stats["huice_order_months"] else 0
    stats["receivable_mapping_rate"] = stats["mapped_receivable"] / stats["huice_receivable"] if stats["huice_receivable"] else 0
    stats["cash_mapping_rate"] = stats["mapped_cash"] / stats["huice_cash"] if stats["huice_cash"] else 0
    return stats


def export_outputs(conn: sqlite3.Connection, output_dir: Path, stats: dict) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    exports = {
        "huice_oms_month_item_recon.csv": "SELECT * FROM huice_oms_month_item_recon ORDER BY CASE result WHEN '应收金额一致' THEN 3 WHEN '实收金额一致' THEN 2 ELSE 1 END,result,huice_month,huice_shop,material_code",
        "huice_oms_month_shop_recon.csv": "SELECT * FROM huice_oms_month_shop_recon ORDER BY CASE result WHEN '店铺月度应收一致' THEN 2 WHEN '店铺月度实收一致' THEN 2 ELSE 1 END,result,huice_month,huice_shop",
        "huice_oms_shop_map.csv": "SELECT * FROM huice_oms_shop_map ORDER BY final_status,platform,huice_shop",
        "huice_detail_summary_month_shop_recon_v3.csv": "SELECT * FROM huice_detail_summary_month_shop_recon_v3 ORDER BY CASE result WHEN '金额一致' THEN 2 ELSE 1 END,result,month,platform,shop",
        "huice_shop_material_month.csv": "SELECT * FROM huice_shop_material_month ORDER BY huice_month,platform,huice_shop,material_code",
    }
    for filename, query in exports.items():
        cur = conn.execute(query)
        with (output_dir / filename).open("w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f); writer.writerow([x[0] for x in cur.description]); writer.writerows(cur)
    samples = {
        "huice_oms_month_item_recon_workbook.json": (exports["huice_oms_month_item_recon.csv"] + " LIMIT 20000"),
        "huice_oms_month_shop_recon_workbook.json": exports["huice_oms_month_shop_recon.csv"],
        "huice_oms_shop_map_workbook.json": exports["huice_oms_shop_map.csv"],
        "huice_detail_summary_month_shop_recon_v3_workbook.json": exports["huice_detail_summary_month_shop_recon_v3.csv"],
        "huice_shop_material_month_workbook.json": exports["huice_shop_material_month.csv"],
    }
    for filename, query in samples.items():
        cur = conn.execute(query); headers = [x[0] for x in cur.description]; rows = cur.fetchall()
        (output_dir / filename).write_text(json.dumps({"headers":headers,"rows":rows},ensure_ascii=False),encoding="utf-8")
    def dict_rows(query):
        cur=conn.execute(query); headers=[x[0] for x in cur.description]
        return [dict(zip(headers,row)) for row in cur.fetchall()]
    stats["item_results"] = dict_rows("SELECT result,COUNT(*) groups,SUM(huice_receivable) huice_receivable,SUM(huice_cash) huice_cash,SUM(oms_amount) oms_amount,SUM(bridge_qty) bridge_qty,SUM(oms_qty) oms_qty FROM huice_oms_month_item_recon GROUP BY result ORDER BY groups DESC")
    stats["month_results"] = dict_rows("SELECT huice_month,COUNT(*) groups,SUM(huice_receivable) huice_receivable,SUM(huice_cash) huice_cash,SUM(oms_amount) oms_amount,SUM(receivable_difference) receivable_difference,SUM(cash_difference) cash_difference FROM huice_oms_month_item_recon GROUP BY huice_month ORDER BY huice_month")
    stats["internal_results"] = dict_rows("SELECT result,COUNT(*) groups,SUM(detail_receivable) detail_receivable,SUM(success_amount+mismatch_receivable+single_ar_amount) summary_receivable,SUM(receivable_difference) receivable_difference FROM huice_detail_summary_month_shop_recon_v3 GROUP BY result ORDER BY groups DESC")
    (output_dir / "huice_oms_summary.json").write_text(json.dumps(stats, ensure_ascii=False, indent=2), encoding="utf-8")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", type=Path, default=ROOT.parent / "input")
    parser.add_argument("--db", type=Path, default=ROOT / "work" / "reconciliation.db")
    parser.add_argument("--work", type=Path, default=ROOT / "work")
    parser.add_argument("--output", type=Path, default=ROOT / "intermediate" / "huice_oms")
    parser.add_argument("--workers", type=int, default=4)
    parser.add_argument("--max-rows", type=int)
    parser.add_argument("--reuse-order-items", action="store_true")
    parser.add_argument("--reuse-huice-current", action="store_true")
    args = parser.parse_args()
    conn = sqlite3.connect(args.db); BASE.configure_database(conn)
    if not args.reuse_order_items or not conn.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='wdt_order_item'").fetchone():
        load_order_items(conn,args.input,args.work,args.workers,args.max_rows)
    if not args.reuse_huice_current or not conn.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='huice_current_amount'").fetchone():
        load_huice_current(conn,args.input,args.work,args.workers,args.max_rows)
    stats = build_reconciliation(conn)
    export_outputs(conn,args.output,stats)
    print(json.dumps(stats,ensure_ascii=False,indent=2))


if __name__ == "__main__":
    main()
