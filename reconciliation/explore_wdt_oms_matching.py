#!/usr/bin/env python3
"""探索旺店通与OMS的候选钩稽关系。

核心思路：
1. 从旺店通原始文件抽取发货月、店铺、SAP编码、货品编号、单位、数量和分摊金额。
2. OMS只使用Y005（日结候选）与Y001（月结候选），避免两类单据重复计入。
3. 比较月+店铺/客户+物料、月+店铺/客户等多个粒度，并输出候选店铺映射评分。
"""

from __future__ import annotations

import argparse
import csv
import importlib.util
import json
import math
import re
import sqlite3
from collections import Counter, defaultdict
from concurrent.futures import ProcessPoolExecutor, as_completed
from difflib import SequenceMatcher
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple


ROOT = Path(__file__).resolve().parent
SPEC = importlib.util.spec_from_file_location("recon_base", ROOT / "reconcile_sales_toc.py")
BASE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader
SPEC.loader.exec_module(BASE)


def month_from_filename(name: str) -> str:
    match = re.search(r"(25|26)年(\d{1,2})月", name)
    return f"20{match.group(1)}-{int(match.group(2)):02d}" if match else ""


def extract_file(task: Tuple[str, str, Optional[int]]) -> Tuple[str, str, int]:
    path_text, cache_text, max_rows = task
    path, cache = Path(path_text), Path(cache_text)
    if cache.exists() and cache.stat().st_size:
        return path.name, str(cache), sum(1 for _ in cache.open("r", encoding="utf-8"))
    header = BASE.read_header(path)
    wanted = {
        "order_no": "订单编号", "shop": "店铺名称", "ship_time": "发货时间",
        "sap_code": "SAP编码", "product_code": "货品编号", "unit": "单位",
        "quantity": "数量", "amount": "分摊后总价", "order_type": "订单类型",
    }
    cols = {key: BASE.header_index(header, name) for key, name in wanted.items()}
    selected = sorted({col for col in cols.values() if col})
    fallback_month = month_from_filename(path.name)
    aggregates: Dict[Tuple[str, ...], List[float]] = defaultdict(lambda: [0.0, 0.0, 0.0])
    product_sap: Counter[Tuple[str, str]] = Counter()
    scanned = 0
    for _, values in BASE.iter_selected_rows(path, selected, max_rows=max_rows):
        def val(key: str) -> str:
            col = cols.get(key)
            return BASE.text(values.get(col)) if col else ""
        order_no = val("order_no")
        if not order_no or order_no == "订单编号":
            continue
        ship_date = val("ship_time")[:10]
        ship_month = BASE.month_from_date(val("ship_time")) or fallback_month
        shop = val("shop")
        sap_code = val("sap_code")
        product_code = val("product_code")
        unit = val("unit")
        order_type = val("order_type")
        if product_code and sap_code:
            product_sap[(product_code, sap_code)] += 1
        material_key = sap_code or ("SKU:" + product_code if product_code else "")
        if not ship_month or not shop or not material_key:
            continue
        key = (ship_date, ship_month, shop, material_key, unit, order_type)
        aggregates[key][0] += BASE.as_number(val("quantity"))
        aggregates[key][1] += BASE.as_number(val("amount"))
        aggregates[key][2] += 1
        scanned += 1
    cache.parent.mkdir(parents=True, exist_ok=True)
    with cache.open("w", encoding="utf-8", newline="") as out:
        writer = csv.writer(out)
        for key, measures in aggregates.items():
            writer.writerow(["A", *key, *measures])
        for (product_code, sap_code), count in product_sap.items():
            writer.writerow(["M", product_code, sap_code, count])
    return path.name, str(cache), scanned


def normalize_shop(value: str) -> str:
    value = BASE.normalize_name(value)
    return re.sub(r"抖音|抖店|快手|小店|天猫|淘宝|京东|拼多多|微信|视频号|小红书|店铺", "", value)


def platform(value: str) -> str:
    for token, canonical in [
        ("抖", "抖音"), ("快手", "快手"), ("天猫", "天猫"), ("淘宝", "天猫"),
        ("京东", "京东"), ("寻梦", "拼多多"), ("拼多多", "拼多多"),
        ("视频号", "视频号"), ("微信", "视频号"), ("小红书", "小红书"),
        ("鲸灵", "鲸灵"), ("有赞", "有赞"),
    ]:
        if token in value:
            return canonical
    return ""


def cosine(a: Dict[Tuple[str, str], float], b: Dict[Tuple[str, str], float]) -> float:
    common = set(a) & set(b)
    numerator = sum(a[k] * b[k] for k in common)
    da = math.sqrt(sum(v * v for v in a.values()))
    db = math.sqrt(sum(v * v for v in b.values()))
    return numerator / da / db if da and db else 0.0


def load_wdt(connection: sqlite3.Connection, input_dir: Path, work_dir: Path, workers: int, max_rows: Optional[int]) -> None:
    connection.executescript("""
    DROP TABLE IF EXISTS wdt_item_raw;
    CREATE TABLE wdt_item_raw(
      ship_date TEXT, ship_month TEXT, shop TEXT, material_key TEXT, unit TEXT, order_type TEXT,
      quantity REAL, amount REAL, line_count INTEGER
    );
    DROP TABLE IF EXISTS wdt_product_sap;
    CREATE TABLE wdt_product_sap(product_code TEXT, sap_code TEXT, occurrences INTEGER);
    """)
    cache_dir = work_dir / "wdt_item_day_cache_v2"
    files = BASE.wdt_files(input_dir)
    tasks = [(str(f), str(cache_dir / f"{i:02d}_{f.stem}.csv"), max_rows) for i, f in enumerate(files, 1)]
    results = []
    with ProcessPoolExecutor(max_workers=max(1, workers)) as pool:
        futures = {pool.submit(extract_file, task): task[0] for task in tasks}
        for i, future in enumerate(as_completed(futures), 1):
            result = future.result(); results.append(result)
            print(f"抽取 {i}/{len(tasks)} {result[0]}，有效商品行 {result[2]:,}", flush=True)
    for name, cache, _ in sorted(results):
        aggs, maps = [], []
        with Path(cache).open("r", encoding="utf-8", newline="") as stream:
            for row in csv.reader(stream):
                if row[0] == "A":
                    aggs.append((*row[1:7], float(row[7]), float(row[8]), int(float(row[9]))))
                else:
                    maps.append((row[1], row[2], int(row[3])))
        connection.executemany("INSERT INTO wdt_item_raw VALUES(?,?,?,?,?,?,?,?,?)", aggs)
        connection.executemany("INSERT INTO wdt_product_sap VALUES(?,?,?)", maps)
        connection.commit()
    connection.executescript("""
    DROP TABLE IF EXISTS wdt_product_sap_best;
    CREATE TABLE wdt_product_sap_best AS
    WITH x AS (SELECT product_code,sap_code,SUM(occurrences) n FROM wdt_product_sap GROUP BY 1,2),
    r AS (SELECT *,ROW_NUMBER() OVER(PARTITION BY product_code ORDER BY n DESC,sap_code) rn FROM x)
    SELECT product_code,sap_code,n FROM r WHERE rn=1;

    DROP TABLE IF EXISTS wdt_item;
    CREATE TABLE wdt_item AS
    SELECT ship_date,ship_month,shop,
      CASE WHEN material_key LIKE 'SKU:%' THEN COALESCE(m.sap_code,material_key) ELSE material_key END material_code,
      unit,order_type,SUM(quantity) quantity,SUM(amount) amount,SUM(line_count) line_count
    FROM wdt_item_raw w LEFT JOIN wdt_product_sap_best m ON substr(w.material_key,5)=m.product_code
    WHERE ship_month BETWEEN '2025-12' AND '2026-06'
    GROUP BY 1,2,3,4,5,6;
    CREATE INDEX IF NOT EXISTS idx_wdt_item_key ON wdt_item(ship_month,shop,material_code);
    CREATE INDEX IF NOT EXISTS idx_wdt_item_day_key ON wdt_item(ship_date,shop,material_code);
    """)
    connection.commit()


def build_candidates(connection: sqlite3.Connection, output_dir: Path) -> Dict[str, object]:
    shops = [r[0] for r in connection.execute("SELECT DISTINCT shop FROM wdt_item WHERE shop<>''")]
    customers = list(connection.execute("""
      SELECT DISTINCT customer_code,customer_name FROM oms_detail
      WHERE business_type='Y005' AND outbound_month BETWEEN '2025-12' AND '2026-06'
    """))
    w_vectors: Dict[str, Dict[Tuple[str, str], float]] = {}
    for shop in shops:
        w_vectors[shop] = {(m, item): amt for m, item, amt in connection.execute(
            "SELECT ship_month,material_code,SUM(amount) FROM wdt_item WHERE shop=? GROUP BY 1,2", (shop,)
        )}
    o_vectors: Dict[str, Dict[Tuple[str, str], float]] = {}
    for code, _ in customers:
        o_vectors[code] = {(m, item): amt for m, item, amt in connection.execute(
            "SELECT outbound_month,item_code,SUM(share_amount) FROM oms_detail WHERE business_type='Y005' AND customer_code=? GROUP BY 1,2", (code,)
        )}
    candidates = []
    for code, name in customers:
        p = platform(name)
        for shop in shops:
            sp = platform(shop)
            name_score = SequenceMatcher(None, normalize_shop(name), normalize_shop(shop)).ratio()
            vector_score = cosine(o_vectors[code], w_vectors[shop])
            platform_score = 1.0 if p and p == sp else (0.0 if p and sp and p != sp else 0.4)
            total = 0.50 * vector_score + 0.30 * name_score + 0.20 * platform_score
            candidates.append((code, name, shop, p, sp, name_score, vector_score, platform_score, total))
    candidates.sort(key=lambda r: (r[0], -r[-1]))
    best = []
    seen = set()
    for row in candidates:
        if row[0] not in seen:
            best.append(row); seen.add(row[0])
    connection.executescript("DROP TABLE IF EXISTS wdt_oms_shop_candidates; CREATE TABLE wdt_oms_shop_candidates(customer_code,customer_name,wdt_shop,oms_platform,wdt_platform,name_score,vector_score,platform_score,total_score);")
    connection.executemany("INSERT INTO wdt_oms_shop_candidates VALUES(?,?,?,?,?,?,?,?,?)", candidates)
    connection.executescript("""
    DROP TABLE IF EXISTS wdt_oms_shop_map;
    CREATE TABLE wdt_oms_shop_map AS
    WITH r AS (SELECT *,ROW_NUMBER() OVER(PARTITION BY customer_code ORDER BY total_score DESC,wdt_shop) rn FROM wdt_oms_shop_candidates)
    SELECT customer_code,customer_name,wdt_shop,name_score,vector_score,platform_score,total_score,
      CASE WHEN total_score>=0.72 THEN '高置信' WHEN total_score>=0.55 THEN '待复核' ELSE '低置信' END mapping_status
    FROM r WHERE rn=1;

    DROP TABLE IF EXISTS wdt_oms_item_recon;
    CREATE TABLE wdt_oms_item_recon AS
    WITH w AS (
      SELECT ship_month,shop,material_code,SUM(quantity) wdt_qty,SUM(amount) wdt_amount,SUM(line_count) wdt_lines
      FROM wdt_item GROUP BY 1,2,3
    ), o AS (
      SELECT outbound_month,customer_code,MIN(customer_name) customer_name,item_code,SUM(item_num) oms_qty,SUM(share_amount) oms_amount,
        COUNT(DISTINCT document_no) oms_docs,COUNT(*) oms_lines
      FROM oms_detail WHERE business_type='Y005' GROUP BY 1,2,4
    ), left_side AS (
      SELECT w.ship_month,w.shop,m.customer_code,m.customer_name,w.material_code,w.wdt_qty,w.wdt_amount,w.wdt_lines,
        o.oms_qty,o.oms_amount,o.oms_docs,o.oms_lines,m.total_score,m.mapping_status,
        COALESCE(o.oms_qty,0)-w.wdt_qty qty_difference,COALESCE(o.oms_amount,0)-w.wdt_amount amount_difference,
        CASE WHEN o.item_code IS NULL THEN '仅旺店通'
             WHEN ABS(o.oms_qty-w.wdt_qty)<=0.000001 AND ABS(o.oms_amount-w.wdt_amount)<=0.01 THEN '数量金额一致'
             WHEN ABS(o.oms_qty-w.wdt_qty)<=0.000001 THEN '数量一致金额差异'
             ELSE '数量金额差异' END result
      FROM w LEFT JOIN wdt_oms_shop_map m ON m.wdt_shop=w.shop
      LEFT JOIN o ON o.outbound_month=w.ship_month AND o.customer_code=m.customer_code AND o.item_code=w.material_code
    ) SELECT * FROM left_side
    UNION ALL
    SELECT o.outbound_month,m.wdt_shop,o.customer_code,o.customer_name,o.item_code,0,0,0,o.oms_qty,o.oms_amount,o.oms_docs,o.oms_lines,m.total_score,m.mapping_status,o.oms_qty,o.oms_amount,'仅OMS'
    FROM o JOIN wdt_oms_shop_map m ON m.customer_code=o.customer_code
    LEFT JOIN w ON w.ship_month=o.outbound_month AND w.shop=m.wdt_shop AND w.material_code=o.item_code
    WHERE w.material_code IS NULL;

    DROP TABLE IF EXISTS wdt_oms_shop_month_recon;
    CREATE TABLE wdt_oms_shop_month_recon AS
    SELECT ship_month,shop,customer_code,customer_name,mapping_status,total_score,
      SUM(wdt_qty) wdt_qty,SUM(oms_qty) oms_qty,SUM(qty_difference) qty_difference,
      SUM(wdt_amount) wdt_amount,SUM(oms_amount) oms_amount,SUM(amount_difference) amount_difference,
      SUM(CASE WHEN result='数量金额一致' THEN 1 ELSE 0 END) exact_item_groups,
      COUNT(*) item_groups,
      CASE WHEN ABS(SUM(qty_difference))<=0.000001 AND ABS(SUM(amount_difference))<=0.01 THEN '店铺月度数量金额一致'
           WHEN ABS(SUM(qty_difference))<=0.000001 THEN '店铺月度数量一致金额差异'
           ELSE '店铺月度差异' END result
    FROM wdt_oms_item_recon GROUP BY 1,2,3,4,5,6;

    DROP TABLE IF EXISTS wdt_oms_day_item_recon;
    CREATE TABLE wdt_oms_day_item_recon AS
    WITH w AS (
      SELECT ship_date,shop,material_code,SUM(quantity) wdt_qty,SUM(amount) wdt_amount,SUM(line_count) wdt_lines
      FROM wdt_item WHERE ship_date<>'' GROUP BY 1,2,3
    ), o AS (
      SELECT substr(outbound_time,1,10) outbound_date,customer_code,MIN(customer_name) customer_name,item_code,
        SUM(item_num) oms_qty,SUM(share_amount) oms_amount,COUNT(DISTINCT document_no) oms_docs,COUNT(*) oms_lines
      FROM oms_detail WHERE business_type='Y005' GROUP BY 1,2,4
    ), left_side AS (
      SELECT w.ship_date,w.shop,m.customer_code,m.customer_name,w.material_code,w.wdt_qty,w.wdt_amount,w.wdt_lines,
        o.oms_qty,o.oms_amount,o.oms_docs,o.oms_lines,m.total_score,m.mapping_status,
        COALESCE(o.oms_qty,0)-w.wdt_qty qty_difference,COALESCE(o.oms_amount,0)-w.wdt_amount amount_difference,
        CASE WHEN o.item_code IS NULL THEN '仅旺店通'
             WHEN ABS(o.oms_qty-w.wdt_qty)<=0.000001 AND ABS(o.oms_amount-w.wdt_amount)<=0.01 THEN '数量金额一致'
             WHEN ABS(o.oms_qty-w.wdt_qty)<=0.000001 THEN '数量一致金额差异'
             ELSE '数量金额差异' END result
      FROM w LEFT JOIN wdt_oms_shop_map m ON m.wdt_shop=w.shop
      LEFT JOIN o ON o.outbound_date=w.ship_date AND o.customer_code=m.customer_code AND o.item_code=w.material_code
    ) SELECT * FROM left_side
    UNION ALL
    SELECT o.outbound_date,m.wdt_shop,o.customer_code,o.customer_name,o.item_code,0,0,0,o.oms_qty,o.oms_amount,o.oms_docs,o.oms_lines,m.total_score,m.mapping_status,o.oms_qty,o.oms_amount,'仅OMS'
    FROM o JOIN wdt_oms_shop_map m ON m.customer_code=o.customer_code
    LEFT JOIN w ON w.ship_date=o.outbound_date AND w.shop=m.wdt_shop AND w.material_code=o.item_code
    WHERE w.material_code IS NULL;

    DROP TABLE IF EXISTS wdt_oms_sales_day_item_recon;
    CREATE TABLE wdt_oms_sales_day_item_recon AS
    WITH w AS (
      SELECT ship_date,shop,material_code,SUM(quantity) wdt_qty,SUM(amount) wdt_amount,SUM(line_count) wdt_lines
      FROM wdt_item WHERE ship_date<>'' AND order_type='网店销售' GROUP BY 1,2,3
    ), o AS (
      SELECT substr(outbound_time,1,10) outbound_date,customer_code,MIN(customer_name) customer_name,item_code,
        SUM(item_num) oms_qty,SUM(share_amount) oms_amount,COUNT(DISTINCT document_no) oms_docs,COUNT(*) oms_lines
      FROM oms_detail WHERE business_type='Y005' GROUP BY 1,2,4
    ), left_side AS (
      SELECT w.ship_date,w.shop,m.customer_code,m.customer_name,w.material_code,w.wdt_qty,w.wdt_amount,w.wdt_lines,
        o.oms_qty,o.oms_amount,o.oms_docs,o.oms_lines,m.total_score,m.mapping_status,
        COALESCE(o.oms_qty,0)-w.wdt_qty qty_difference,COALESCE(o.oms_amount,0)-w.wdt_amount amount_difference,
        CASE WHEN o.item_code IS NULL THEN '仅旺店通'
             WHEN ABS(o.oms_qty-w.wdt_qty)<=0.000001 AND ABS(o.oms_amount-w.wdt_amount)<=0.01 THEN '数量金额一致'
             WHEN ABS(o.oms_amount-w.wdt_amount)<=0.01 THEN '金额一致数量差异'
             WHEN ABS(o.oms_qty-w.wdt_qty)<=0.000001 THEN '数量一致金额差异'
             ELSE '数量金额差异' END result
      FROM w LEFT JOIN wdt_oms_shop_map m ON m.wdt_shop=w.shop
      LEFT JOIN o ON o.outbound_date=w.ship_date AND o.customer_code=m.customer_code AND o.item_code=w.material_code
    ) SELECT * FROM left_side
    UNION ALL
    SELECT o.outbound_date,m.wdt_shop,o.customer_code,o.customer_name,o.item_code,0,0,0,o.oms_qty,o.oms_amount,o.oms_docs,o.oms_lines,m.total_score,m.mapping_status,o.oms_qty,o.oms_amount,'仅OMS'
    FROM o JOIN wdt_oms_shop_map m ON m.customer_code=o.customer_code
    LEFT JOIN w ON w.ship_date=o.outbound_date AND w.shop=m.wdt_shop AND w.material_code=o.item_code
    WHERE w.material_code IS NULL;

    DROP TABLE IF EXISTS wdt_oms_order_type_recon;
    CREATE TABLE wdt_oms_order_type_recon AS
    WITH pairs(wdt_order_type,oms_business_type) AS (
      VALUES('网店销售','Y005'),('分销订单','Y051'),('样品发货','Z003'),
            ('退货损失','Z004'),('赠品','Z001'),('线下订单','Z006')
    ), w AS (
      SELECT ship_date d,order_type,material_code item,SUM(quantity) qty,SUM(amount) amount
      FROM wdt_item WHERE ship_date<>'' GROUP BY 1,2,3
    ), o AS (
      SELECT substr(outbound_time,1,10) d,business_type,item_code item,SUM(item_num) qty,SUM(share_amount) amount
      FROM oms_detail GROUP BY 1,2,3
    ), exact AS (
      SELECT p.wdt_order_type,p.oms_business_type,w.d,w.item,w.qty,w.amount
      FROM pairs p JOIN w ON w.order_type=p.wdt_order_type
      JOIN o ON o.business_type=p.oms_business_type AND o.d=w.d AND o.item=w.item
      WHERE ABS(o.qty-w.qty)<=0.000001 AND ABS(o.amount-w.amount)<=0.01
    ), totals AS (
      SELECT order_type,SUM(qty) wdt_qty,SUM(amount) wdt_amount FROM w GROUP BY 1
    )
    SELECT p.wdt_order_type,p.oms_business_type,t.wdt_qty,t.wdt_amount,
      COUNT(e.item) exact_groups,COALESCE(SUM(e.qty),0) exact_qty,
      COALESCE(SUM(e.qty),0)/NULLIF(t.wdt_qty,0) qty_coverage,
      COALESCE(SUM(e.amount),0) exact_amount,
      CASE WHEN ABS(t.wdt_amount)>0.001 THEN COALESCE(SUM(e.amount),0)/t.wdt_amount END amount_coverage
    FROM pairs p JOIN totals t ON t.order_type=p.wdt_order_type
    LEFT JOIN exact e ON e.wdt_order_type=p.wdt_order_type AND e.oms_business_type=p.oms_business_type
    GROUP BY 1,2,3,4;
    """)
    connection.commit()

    output_dir.mkdir(parents=True, exist_ok=True)
    exports = {
        "wdt_oms_shop_map.csv": "SELECT * FROM wdt_oms_shop_map ORDER BY total_score DESC",
        "wdt_oms_item_recon.csv": "SELECT * FROM wdt_oms_item_recon ORDER BY CASE result WHEN '数量金额一致' THEN 2 ELSE 1 END,result,ship_month,shop,material_code",
        "wdt_oms_day_item_recon.csv": "SELECT * FROM wdt_oms_day_item_recon ORDER BY CASE result WHEN '数量金额一致' THEN 2 ELSE 1 END,result,ship_date,shop,material_code",
        "wdt_oms_sales_day_item_recon.csv": "SELECT * FROM wdt_oms_sales_day_item_recon ORDER BY CASE result WHEN '数量金额一致' THEN 3 WHEN '金额一致数量差异' THEN 2 ELSE 1 END,result,ship_date,shop,material_code",
        "wdt_oms_order_type_recon.csv": "SELECT * FROM wdt_oms_order_type_recon ORDER BY wdt_amount DESC",
        "wdt_oms_shop_month_recon.csv": "SELECT * FROM wdt_oms_shop_month_recon ORDER BY result,ship_month,shop",
    }
    for filename, query in exports.items():
        cur = connection.execute(query)
        with (output_dir / filename).open("w", encoding="utf-8-sig", newline="") as out:
            writer = csv.writer(out); writer.writerow([d[0] for d in cur.description]); writer.writerows(cur)

    workbook_specs = {
        "wdt_oms_sales_day_item_recon_workbook.json": (
            """SELECT * FROM wdt_oms_sales_day_item_recon
               ORDER BY CASE result WHEN '数量金额一致' THEN 4 WHEN '金额一致数量差异' THEN 3
                                    WHEN '数量一致金额差异' THEN 2 ELSE 1 END,
                        ABS(amount_difference) DESC,ship_date,shop,material_code LIMIT 20000"""
        ),
        "wdt_oms_shop_map_workbook.json": "SELECT * FROM wdt_oms_shop_map ORDER BY total_score DESC",
        "wdt_oms_order_type_recon_workbook.json": "SELECT * FROM wdt_oms_order_type_recon ORDER BY wdt_amount DESC",
        "wdt_oms_shop_month_recon_workbook.json": "SELECT * FROM wdt_oms_shop_month_recon ORDER BY result,ship_month,shop",
    }
    for filename, query in workbook_specs.items():
        cur = connection.execute(query)
        headers = [d[0] for d in cur.description]
        payload = {"headers": headers, "rows": [list(row) for row in cur.fetchall()]}
        (output_dir / filename).write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")

    def rows(query: str) -> List[Dict[str, object]]:
        cur = connection.execute(query); cols = [d[0] for d in cur.description]
        return [dict(zip(cols, row)) for row in cur.fetchall()]
    result = {
        "totals": rows("""SELECT 'WDT' source,SUM(quantity) quantity,SUM(amount) amount FROM wdt_item
          UNION ALL SELECT 'OMS_Y005',SUM(item_num),SUM(share_amount) FROM oms_detail WHERE business_type='Y005'
          UNION ALL SELECT 'OMS_Y001',SUM(item_num),SUM(share_amount) FROM oms_detail WHERE business_type='Y001'"""),
        "mapping_status": rows("SELECT mapping_status,COUNT(*) customers,SUM(total_score)/COUNT(*) avg_score FROM wdt_oms_shop_map GROUP BY 1"),
        "item_results": rows("SELECT result,COUNT(*) groups,SUM(wdt_qty) wdt_qty,SUM(oms_qty) oms_qty,SUM(wdt_amount) wdt_amount,SUM(oms_amount) oms_amount FROM wdt_oms_item_recon GROUP BY 1"),
        "day_item_results": rows("SELECT result,COUNT(*) groups,SUM(wdt_qty) wdt_qty,SUM(oms_qty) oms_qty,SUM(wdt_amount) wdt_amount,SUM(oms_amount) oms_amount FROM wdt_oms_day_item_recon GROUP BY 1"),
        "sales_day_item_results": rows("SELECT result,COUNT(*) groups,SUM(wdt_qty) wdt_qty,SUM(oms_qty) oms_qty,SUM(wdt_amount) wdt_amount,SUM(oms_amount) oms_amount FROM wdt_oms_sales_day_item_recon GROUP BY 1"),
        "order_type_results": rows("SELECT * FROM wdt_oms_order_type_recon ORDER BY wdt_amount DESC"),
        "sales_controls": rows("""SELECT COUNT(*) wdt_groups,SUM(wdt_qty) wdt_qty,SUM(wdt_amount) wdt_amount,
          SUM(CASE WHEN oms_qty IS NOT NULL THEN 1 ELSE 0 END) linked_groups,
          SUM(CASE WHEN oms_qty IS NOT NULL AND ABS(amount_difference)<=0.01 THEN 1 ELSE 0 END) amount_exact_groups,
          SUM(CASE WHEN ABS(amount_difference)<=0.01 THEN wdt_amount ELSE 0 END) amount_exact,
          SUM(CASE WHEN ABS(amount_difference)<=0.01 THEN wdt_amount ELSE 0 END)/NULLIF(SUM(wdt_amount),0) amount_coverage,
          SUM(CASE WHEN ABS(qty_difference)<=0.000001 THEN wdt_qty ELSE 0 END) qty_exact,
          SUM(CASE WHEN ABS(qty_difference)<=0.000001 THEN wdt_qty ELSE 0 END)/NULLIF(SUM(wdt_qty),0) qty_coverage
          FROM wdt_oms_sales_day_item_recon WHERE wdt_qty>0"""),
        "shop_month_results": rows("SELECT result,COUNT(*) groups,SUM(wdt_qty) wdt_qty,SUM(oms_qty) oms_qty,SUM(wdt_amount) wdt_amount,SUM(oms_amount) oms_amount FROM wdt_oms_shop_month_recon GROUP BY 1"),
        "top_mappings": rows("SELECT * FROM wdt_oms_shop_map ORDER BY total_score DESC LIMIT 30"),
    }
    (output_dir / "exploration_summary.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return result


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", type=Path, default=ROOT.parent / "input")
    parser.add_argument("--database", type=Path, default=ROOT / "work" / "reconciliation.db")
    parser.add_argument("--work", type=Path, default=ROOT / "work")
    parser.add_argument("--output", type=Path, default=ROOT / "intermediate" / "exploration")
    parser.add_argument("--workers", type=int, default=4)
    parser.add_argument("--max-rows-per-file", type=int)
    parser.add_argument("--reuse-wdt", action="store_true")
    args = parser.parse_args()
    connection = sqlite3.connect(args.database)
    BASE.configure_database(connection)
    if not args.reuse_wdt:
        load_wdt(connection, args.input, args.work, args.workers, args.max_rows_per_file)
    result = build_candidates(connection, args.output)
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
