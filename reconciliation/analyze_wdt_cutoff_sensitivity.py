#!/usr/bin/env python3
"""旺店通订单期间敏感性分析。

保持正式2026-01至2026-06订单池及匹配结果不变，2025年12月及新增
2025年11月订单加载至临时表，分别测算2026年度、追溯至2025年12月及追溯至
2025年11月三个范围对惠策账单订单号覆盖和金额一致覆盖的影响。
"""

from __future__ import annotations

import csv
import json
import sqlite3
import sys
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "reconciliation"))

from reconcile_sales_toc import extract_wdt_file_to_csv, read_header  # noqa: E402

INPUT = ROOT / "input" / "旺店通订单清单"
NOVEMBER = INPUT / "25年11月"
DB = ROOT / "reconciliation" / "work" / "reconciliation.db"
WORK = ROOT / "reconciliation" / "work" / "wdt_november_platform_cache"
OUTPUT = ROOT / "reconciliation" / "results" / "wdt_cutoff_sensitivity.json"

CORE_FIELDS = {
    "订单号": ["订单编号", "原始单号"],
    "商品编码": ["货品编号", "商家编码"],
    "商品名称": ["货品名称"],
    "数量": ["数量"],
    "金额字段": ["应收金额", "分摊后总价"],
    "店铺名称": ["店铺名称"],
    "下单/支付时间": ["交易时间", "付款时间"],
    "状态字段": ["订单状态"],
}


def inspect_files() -> dict:
    reference = read_header(INPUT / "25年12月-1.xlsx")
    files = sorted(NOVEMBER.glob("*.xlsx"))
    headers = []
    errors = []
    for path in files:
        try:
            header = read_header(path)
            if not header:
                raise ValueError("未读取到表头")
            headers.append((path.name, header))
        except Exception as exc:  # pragma: no cover - audit output
            errors.append(f"{path.name}: {exc}")
    common = set.intersection(*(set(header) for _, header in headers)) if headers else set()
    new_union = set.union(*(set(header) for _, header in headers)) if headers else set()
    core = {label: all(field in common for field in fields) for label, fields in CORE_FIELDS.items()}
    missing = [field for field in reference if field not in new_union]
    added = [field for field in sorted(new_union) if field not in set(reference)]
    return {
        "files": [path.name for path in files],
        "file_count": len(files),
        "readable": bool(files) and not errors and len(headers) == len(files),
        "errors": errors,
        "reference_column_count": len(reference),
        "new_column_counts": sorted({len(header) for _, header in headers}),
        "all_new_files_same_schema": len({tuple(header) for _, header in headers}) <= 1,
        "exact_field_consistency": bool(headers) and all(header == reference for _, header in headers),
        "missing_fields": missing,
        "added_fields": added,
        "core_fields": core,
        "usable_for_matching": bool(headers) and all(core.values()),
        "note": "11月文件缺少SAP编码等非订单号及金额匹配必需字段；本次敏感性分析仅使用与现有逻辑相同的订单号、金额、数量、店铺及日期字段。",
    }


def extract_november() -> list[tuple[str, str, int]]:
    WORK.mkdir(parents=True, exist_ok=True)
    files = sorted(NOVEMBER.glob("*.xlsx"))
    tasks = [(str(path), str(WORK / f"{index:02d}_{path.stem}.csv"), None) for index, path in enumerate(files, 1)]
    results = []
    with ProcessPoolExecutor(max_workers=min(4, max(1, len(tasks)))) as executor:
        futures = [executor.submit(extract_wdt_file_to_csv, task) for task in tasks]
        for future in as_completed(futures):
            results.append(future.result())
    return sorted(results)


def load_november(conn: sqlite3.Connection, extracts: list[tuple[str, str, int]]) -> None:
    conn.executescript("""
    DROP TABLE IF EXISTS temp.sens_nov_file;
    CREATE TEMP TABLE sens_nov_file(
      source_file TEXT,order_no TEXT,platform_order_no TEXT,sub_order_no TEXT,shop TEXT,
      order_status TEXT,order_type TEXT,trade_time TEXT,payment_time TEXT,ship_time TEXT,
      receivable_amount REAL,allocated_total REAL,quantity REAL,line_count INTEGER,schema_columns INTEGER
    );
    """)
    insert = "INSERT INTO sens_nov_file VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
    for _, cache_path, _ in extracts:
        batch = []
        with Path(cache_path).open("r", encoding="utf-8", newline="") as stream:
            for row in csv.reader(stream):
                batch.append(row)
                if len(batch) >= 20000:
                    conn.executemany(insert, batch)
                    batch.clear()
            if batch:
                conn.executemany(insert, batch)
    conn.executescript("""
    DROP TABLE IF EXISTS temp.sens_nov_dedup;
    CREATE TEMP TABLE sens_nov_dedup AS
    SELECT order_no,platform_order_no,MAX(sub_order_no) sub_order_no,MAX(shop) shop,
      MAX(order_status) order_status,MAX(order_type) order_type,
      MIN(NULLIF(trade_time,'')) trade_time,MIN(NULLIF(payment_time,'')) payment_time,
      MAX(NULLIF(ship_time,'')) ship_time,MAX(receivable_amount) receivable_amount,
      MAX(allocated_total) allocated_total,MAX(quantity) quantity,MAX(line_count) line_count,
      COUNT(*) file_occurrences,GROUP_CONCAT(source_file,'|') source_files
    FROM sens_nov_file GROUP BY order_no,platform_order_no;

    DROP TABLE IF EXISTS temp.sens_wdt_combined;
    CREATE TEMP TABLE sens_wdt_combined AS
    SELECT order_no,platform_order_no,MAX(sub_order_no) sub_order_no,MAX(shop) shop,
      MAX(order_status) order_status,MAX(order_type) order_type,
      MIN(NULLIF(trade_time,'')) trade_time,MIN(NULLIF(payment_time,'')) payment_time,
      MAX(NULLIF(ship_time,'')) ship_time,MAX(receivable_amount) receivable_amount,
      MAX(allocated_total) allocated_total,MAX(quantity) quantity,MAX(line_count) line_count
    FROM (
      SELECT order_no,platform_order_no,sub_order_no,shop,order_status,order_type,trade_time,payment_time,ship_time,
        receivable_amount,allocated_total,quantity,line_count FROM wdt_order_dedup
      UNION ALL
      SELECT order_no,platform_order_no,sub_order_no,shop,order_status,order_type,trade_time,payment_time,ship_time,
        receivable_amount,allocated_total,quantity,line_count FROM sens_nov_dedup
    ) GROUP BY order_no,platform_order_no;
    CREATE INDEX temp.idx_sens_wdt_dates ON sens_wdt_combined(trade_time,payment_time,ship_time);
    CREATE INDEX temp.idx_sens_wdt_platform ON sens_wdt_combined(platform_order_no);
    """)


def scenario(conn: sqlite3.Connection, label: str, start: str) -> dict:
    conn.executescript("DROP TABLE IF EXISTS temp.sens_scope; DROP TABLE IF EXISTS temp.sens_platform;")
    conn.execute("""
      CREATE TEMP TABLE sens_scope AS
      SELECT * FROM sens_wdt_combined
      WHERE (trade_time>=? AND trade_time<'2026-07-01')
         OR (payment_time>=? AND payment_time<'2026-07-01')
         OR (ship_time>=? AND ship_time<'2026-07-01')
    """, (start, start, start))
    conn.executescript("""
      CREATE TEMP TABLE sens_platform AS
      SELECT CASE WHEN platform_order_no='' THEN '__WDT__'||order_no ELSE platform_order_no END platform_order_no,
        SUM(receivable_amount) wdt_header_amount,SUM(allocated_total) wdt_amount,SUM(quantity) wdt_qty
      FROM sens_scope
      GROUP BY CASE WHEN platform_order_no='' THEN '__WDT__'||order_no ELSE platform_order_no END;
      CREATE UNIQUE INDEX temp.idx_sens_platform ON sens_platform(platform_order_no);
    """)
    order_count, amount, qty = conn.execute(
        "SELECT COUNT(*),COALESCE(SUM(wdt_amount),0),COALESCE(SUM(wdt_qty),0) FROM sens_platform"
    ).fetchone()
    huice_orders, huice_cash = conn.execute(
        "SELECT COUNT(*),COALESCE(SUM(bill_cash),0) FROM v4_huice_platform"
    ).fetchone()
    key_matches = conn.execute(
        "SELECT COUNT(*) FROM v4_huice_platform h JOIN sens_platform w ON w.platform_order_no=h.platform_order_no"
    ).fetchone()[0]
    exact_orders, exact_cash, exact_wdt_amount = conn.execute("""
      SELECT COUNT(*),COALESCE(SUM(h.bill_cash),0),
        COALESCE(SUM(CASE WHEN ABS(w.wdt_amount-h.bill_cash)<=0.01
                          THEN w.wdt_amount ELSE w.wdt_header_amount END),0)
      FROM v4_huice_platform h JOIN sens_platform w ON w.platform_order_no=h.platform_order_no
      WHERE ABS(w.wdt_amount-h.bill_cash)<=0.01 OR ABS(w.wdt_header_amount-h.bill_cash)<=0.01
    """).fetchone()
    return {
        "range": label,
        "start_date": start,
        "end_date": "2026-06-30",
        "order_count": order_count or 0,
        "order_amount": amount or 0,
        "order_quantity": qty or 0,
        "huice_order_count": huice_orders or 0,
        "huice_cash": huice_cash or 0,
        "key_matched_orders": key_matches or 0,
        "amount_exact_orders": exact_orders or 0,
        "amount_exact_cash": exact_cash or 0,
        "amount_exact_wdt_amount": exact_wdt_amount or 0,
        "amount_match_rate": abs(exact_cash or 0) / abs(huice_cash or 1),
        "quantity_match_rate": (key_matches or 0) / (huice_orders or 1),
    }


def main() -> None:
    audit = inspect_files()
    if not audit["usable_for_matching"]:
        OUTPUT.write_text(json.dumps({"data_check": audit, "sensitivity": []}, ensure_ascii=False, indent=2), encoding="utf-8")
        raise SystemExit("新增11月文件缺少核心匹配字段，未执行敏感性分析")
    extracts = extract_november()
    conn = sqlite3.connect(DB)
    conn.executescript("PRAGMA journal_mode=OFF; PRAGMA synchronous=OFF; PRAGMA temp_store=FILE; PRAGMA cache_size=-300000;")
    load_november(conn, extracts)
    settings = [("2026年度", "2026-01-01"), ("2025.12-2026.06", "2025-12-01"), ("2025.11-2026.06", "2025-11-01")]
    sensitivity = [scenario(conn, label, start) for label, start in settings]
    for index, item in enumerate(sensitivity):
        previous = sensitivity[index - 1] if index else item
        item["incremental_order_count"] = item["order_count"] - previous["order_count"] if index else 0
        item["incremental_amount"] = item["order_amount"] - previous["order_amount"] if index else 0
        item["amount_rate_change"] = item["amount_match_rate"] - previous["amount_match_rate"] if index else 0
        item["quantity_rate_change"] = item["quantity_match_rate"] - previous["quantity_match_rate"] if index else 0
    conn.close()
    result = {
        "data_check": audit,
        "sensitivity": sensitivity,
        "definitions": {
            "order_count": "各范围内旺店通平台订单号数量；空平台单号按内部订单号独立计数",
            "order_amount": "各范围内旺店通平台订单分摊后总价allocated_total合计",
            "increment": "相较上一行期间口径的新增订单数量及新增金额",
            "amount_match_rate": "惠策金额一致订单实际实收/惠策明细全量实际实收",
            "amount_exact_wdt_amount": "金额一致订单对应的旺店通金额；优先使用allocated_total，订单头金额一致场景使用receivable_amount",
            "quantity_match_rate": "平台订单号命中的惠策账单订单数/惠策账单平台订单总数；属于订单数量覆盖率，不是商品数量匹配率",
        },
    }
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
