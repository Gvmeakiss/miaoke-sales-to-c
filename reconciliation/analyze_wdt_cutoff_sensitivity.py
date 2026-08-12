#!/usr/bin/env python3
"""旺店通订单期间敏感性分析。

正式匹配范围为2026-01-01至2026-06-30；仅保留向前追溯至2025年12月的
Cut-off补充测试。扩展范围不改写正式订单池，也不纳入正式三单匹配结果。
"""

from __future__ import annotations

import json
import sqlite3
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DB = ROOT / "reconciliation" / "work" / "reconciliation.db"
OUTPUT = ROOT / "reconciliation" / "results" / "wdt_cutoff_sensitivity.json"


def build_combined(conn: sqlite3.Connection) -> None:
    conn.executescript("""
    DROP TABLE IF EXISTS temp.sens_wdt_combined;
    CREATE TEMP TABLE sens_wdt_combined AS
    SELECT order_no,platform_order_no,MAX(sub_order_no) sub_order_no,MAX(shop) shop,
      MAX(order_status) order_status,MAX(order_type) order_type,
      MIN(NULLIF(trade_time,'')) trade_time,MIN(NULLIF(payment_time,'')) payment_time,
      MAX(NULLIF(ship_time,'')) ship_time,MAX(receivable_amount) receivable_amount,
      MAX(allocated_total) allocated_total,MAX(quantity) quantity,MAX(line_count) line_count
    FROM wdt_order_dedup GROUP BY order_no,platform_order_no;
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
    key_matches, participating_cash, participating_wdt_amount = conn.execute("""
      SELECT COUNT(*),COALESCE(SUM(h.bill_cash),0),
        COALESCE(SUM(CASE WHEN ABS(w.wdt_amount-h.bill_cash)<=0.01 THEN w.wdt_amount
                          WHEN ABS(w.wdt_header_amount-h.bill_cash)<=0.01 THEN w.wdt_header_amount
                          ELSE w.wdt_amount END),0)
      FROM v4_huice_platform h JOIN sens_platform w ON w.platform_order_no=h.platform_order_no
    """).fetchone()
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
        "participating_cash": participating_cash or 0,
        "participating_wdt_amount": participating_wdt_amount or 0,
        "amount_exact_orders": exact_orders or 0,
        "amount_exact_cash": exact_cash or 0,
        "amount_exact_wdt_amount": exact_wdt_amount or 0,
        "amount_match_rate": min(abs(exact_cash or 0), abs(exact_wdt_amount or 0)) / max(abs(exact_cash or 0), abs(exact_wdt_amount or 0), 1),
        "huice_coverage_rate": abs(exact_cash or 0) / max(abs(huice_cash or 0), 1),
        "wdt_coverage_rate": abs(exact_wdt_amount or 0) / max(abs(amount or 0), 1),
        "quantity_match_rate": (key_matches or 0) / (huice_orders or 1),
    }


def main() -> None:
    conn = sqlite3.connect(DB)
    conn.executescript("PRAGMA journal_mode=OFF; PRAGMA synchronous=OFF; PRAGMA temp_store=FILE; PRAGMA cache_size=-300000;")
    build_combined(conn)
    settings = [("2026年度", "2026-01-01"), ("2025.12-2026.06", "2025-12-01")]
    sensitivity = [scenario(conn, label, start) for label, start in settings]
    for index, item in enumerate(sensitivity):
        previous = sensitivity[index - 1] if index else item
        item["incremental_order_count"] = item["order_count"] - previous["order_count"] if index else 0
        item["incremental_amount"] = item["order_amount"] - previous["order_amount"] if index else 0
        item["huice_coverage_change"] = item["huice_coverage_rate"] - previous["huice_coverage_rate"] if index else 0
    conn.close()
    result = {
        "data_check": {"source": "reconciliation.db/wdt_order_dedup", "usable_for_matching": True},
        "sensitivity": sensitivity,
        "definitions": {
            "order_amount": "各范围内旺店通平台订单分摊金额allocated_total合计",
            "participating_amount": "惠策平台订单号与旺店通原始单号共同键范围内的双方金额，不要求金额一致",
            "actual_match_amount": "共同键范围内金额差异不超过0.01元的双方金额",
            "amount_match_rate": "双方实际匹配金额较小值/较大值",
            "coverage_rate": "各方实际匹配金额/各方执行总金额",
        },
    }
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
