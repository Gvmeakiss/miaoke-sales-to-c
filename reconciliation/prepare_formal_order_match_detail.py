#!/usr/bin/env python3
"""生成2026年度正式范围旺店通—惠策订单匹配明细CSV及控制数。"""

from __future__ import annotations

import csv
import json
import sqlite3
from pathlib import Path


ROOT = Path(__file__).resolve().parent.parent
DB = ROOT / "reconciliation" / "work" / "reconciliation.db"
OUT = ROOT / "reconciliation" / "results" / "formal_order_match_detail"

HEADERS = [
    "记录类型", "匹配状态", "惠策订单号", "旺店通订单号", "惠策店铺名称", "旺店通店铺名称",
    "商品编码（旺店通）", "商品名称（OMS物料名称辅助）", "账单日期（业务日期范围）", "订单日期",
    "惠策实际实收", "旺店通订单金额", "金额差异（惠策-旺店通）", "旺店通分摊金额",
    "旺店通表头金额", "旺店通数量", "惠策对账流水号", "惠策平台", "匹配字段/规则", "未匹配原因",
]

CATEGORIES = {
    "可匹配条目": {
        "file": "01_可匹配条目.csv",
        "where": "r.result IN ('单号分摊实收一致','单号订单实收一致')",
    },
    "仅账单未匹配": {
        "file": "02_仅账单及账单未成功匹配.csv",
        "where": "r.result IN ('仅账单','单号一致金额差异')",
    },
    "仅订单未匹配": {
        "file": "03_仅订单未匹配.csv",
        "where": "r.result IN ('仅订单','同内部订单零金额附属单')",
    },
}


def build_support_tables(conn: sqlite3.Connection) -> None:
    conn.executescript("""
    PRAGMA temp_store=FILE;
    PRAGMA cache_size=-500000;
    DROP TABLE IF EXISTS temp.detail_wdt_meta;
    CREATE TEMP TABLE detail_wdt_meta AS
    SELECT s.platform_order_no,
      MIN(NULLIF(s.order_scope_date,'')) order_date,
      GROUP_CONCAT(DISTINCT i.material_code) material_codes
    FROM v4_wdt_order_scope s
    LEFT JOIN wdt_order_item i
      ON i.order_no=s.order_no AND i.platform_order_no=s.platform_order_no
    WHERE s.platform_order_no<>''
    GROUP BY s.platform_order_no;
    CREATE UNIQUE INDEX temp.idx_detail_wdt_meta ON detail_wdt_meta(platform_order_no);

    DROP TABLE IF EXISTS temp.detail_huice_date;
    CREATE TEMP TABLE detail_huice_date AS
    SELECT platform_order_no,
      MIN(NULLIF(business_date,'')) first_bill_date,
      MAX(NULLIF(business_date,'')) last_bill_date
    FROM huice_detail
    WHERE platform_order_no<>''
    GROUP BY platform_order_no;
    CREATE UNIQUE INDEX temp.idx_detail_huice_date ON detail_huice_date(platform_order_no);
    """)


def material_names(conn: sqlite3.Connection) -> dict[str, str]:
    return {
        str(code): str(name or "")
        for code, name in conn.execute("""
          SELECT item_code,MIN(item_name)
          FROM oms_detail
          WHERE item_code<>''
          GROUP BY item_code
        """)
    }


def name_list(codes: str | None, mapping: dict[str, str]) -> str:
    if not codes:
        return "N/A"
    names = []
    seen = set()
    for code in str(codes).split(","):
        name = mapping.get(code, "")
        if name and name not in seen:
            names.append(name)
            seen.add(name)
    return "｜".join(names) if names else "N/A"


def output_query(where: str) -> str:
    return f"""
    SELECT r.result,r.matchable,r.platform_order_no,r.internal_orders,
      r.huice_shop,r.wdt_shop,m.material_codes,
      CASE WHEN h.first_bill_date IS NULL THEN r.bill_month
           WHEN h.first_bill_date=h.last_bill_date THEN h.first_bill_date
           ELSE h.first_bill_date||'~'||h.last_bill_date END bill_date,
      COALESCE(m.order_date,r.order_month) order_date,
      r.bill_cash,
      CASE WHEN r.result='单号订单实收一致' THEN r.wdt_header_amount ELSE r.wdt_amount END comparison_amount,
      r.wdt_amount,r.wdt_header_amount,r.wdt_qty,r.reconcile_ids,r.platform,
      CASE WHEN EXISTS (
        SELECT 1 FROM wdt_order_dedup d WHERE d.platform_order_no=r.platform_order_no
      ) THEN 1 ELSE 0 END exists_outside_formal_scope
    FROM v4_order_bill_recon r
    LEFT JOIN temp.detail_wdt_meta m ON m.platform_order_no=r.platform_order_no
    LEFT JOIN temp.detail_huice_date h ON h.platform_order_no=r.platform_order_no
    WHERE {where}
    """


def classify(row: tuple) -> list:
    (result, matchable, platform_order_no, internal_orders, huice_shop, wdt_shop, codes,
     bill_date, order_date, bill_cash, comparison_amount, allocated_amount, header_amount,
     wdt_qty, reconcile_ids, platform, exists_outside_formal_scope) = row
    bill_cash = float(bill_cash or 0)
    comparison_amount = float(comparison_amount or 0)
    difference = bill_cash - comparison_amount
    if result == "单号分摊实收一致":
        record_type, status = "可匹配条目", "订单号及分摊金额一致"
        rule, reason = "惠策平台订单号=旺店通原始单号；分摊后总价与惠策实际实收差异不超过0.01元", ""
    elif result == "单号订单实收一致":
        record_type, status = "可匹配条目", "订单号及订单头金额一致"
        rule, reason = "惠策平台订单号=旺店通原始单号；旺店通订单头金额与惠策实际实收差异不超过0.01元", ""
    elif result == "单号一致金额差异":
        record_type, status = "仅账单未匹配", "订单号已匹配、金额差异"
        rule, reason = "惠策平台订单号=旺店通原始单号；金额一致性未通过", "金额差异"
    elif result == "仅账单":
        record_type, status = "仅账单未匹配", "惠策账单无正式范围旺店通匹配"
        rule = "惠策平台订单号=旺店通原始单号"
        if not matchable or str(platform_order_no).startswith("__BILL__"):
            reason = "惠策平台订单号为空"
        else:
            reason = "旺店通订单日期不在2026正式范围" if exists_outside_formal_scope else "旺店通正式范围无对应订单"
    elif result == "同内部订单零金额附属单":
        record_type, status = "仅订单未匹配", "同内部订单零金额附属单"
        rule, reason = "同一旺店通内部订单下存在其他已匹配平台订单号", "零金额附属单未形成独立惠策账单"
    else:
        record_type, status = "仅订单未匹配", "旺店通订单无惠策账单匹配"
        rule = "惠策平台订单号=旺店通原始单号"
        reason = "旺店通平台订单号为空" if not matchable or str(platform_order_no).startswith("__WDT__") else "惠策账单无对应订单"
    huice_order_no = "" if str(platform_order_no).startswith("__") and record_type != "可匹配条目" else platform_order_no
    return [
        record_type, status, huice_order_no if bill_cash or reconcile_ids else "", internal_orders or "",
        huice_shop or "", wdt_shop or "", codes or "N/A", None, bill_date or "", order_date or "",
        bill_cash, comparison_amount, difference, float(allocated_amount or 0), float(header_amount or 0),
        float(wdt_qty or 0), reconcile_ids or "", platform or "", rule, reason,
    ]


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB)
    build_support_tables(conn)
    names = material_names(conn)
    controls = {}
    for label, config in CATEGORIES.items():
        path = OUT / config["file"]
        count = 0
        bill_total = 0.0
        wdt_total = 0.0
        with path.open("w", encoding="utf-8", newline="") as stream:
            writer = csv.writer(stream)
            writer.writerow(HEADERS)
            for source in conn.execute(output_query(config["where"])):
                record = classify(source)
                record[7] = name_list(record[6] if record[6] != "N/A" else None, names)
                writer.writerow(record)
                count += 1
                bill_total += record[10]
                wdt_total += record[11]
        check_rows, check_bill, check_wdt = conn.execute(f"""
          SELECT COUNT(*),COALESCE(SUM(r.bill_cash),0),
            COALESCE(SUM(CASE WHEN r.result='单号订单实收一致'
                              THEN r.wdt_header_amount ELSE r.wdt_amount END),0)
          FROM v4_order_bill_recon r
          WHERE {config['where']}
        """).fetchone()
        if count != check_rows:
            raise RuntimeError(f"{label}导出行数{count}与数据库控制数{check_rows}不一致")
        controls[label] = {
            "file": str(path), "rows": count,
            "huice_amount": round(float(check_bill or 0), 2),
            "wdt_amount": round(float(check_wdt or 0), 2),
        }
        print(json.dumps({"category": label, **controls[label]}, ensure_ascii=False), flush=True)
    summary = {
        "scope": "2026-01-01至2026-06-30",
        "matching_rule": "惠策平台订单号=旺店通原始单号；可匹配条目进一步要求金额差异不超过0.01元",
        "categories": controls,
        "control_totals": {
            "huice_orders": 5769505,
            "huice_amount": 275333617.04,
            "wdt_orders": 4728414,
            "wdt_amount": 268424590.67,
            "successful_orders": 4204532,
            "successful_huice_amount": 243122621.11,
            "successful_wdt_amount": 243122622.15,
        },
        "notes": [
            "2025年11月及12月Cut-off扩展数据未纳入本明细。",
            "惠策账单清单无商品编码及商品名称；商品编码取自旺店通商品明细，商品名称由OMS物料名称按商品编码辅助映射。",
            "“仅账单未匹配”区域包含订单号已命中但金额一致性未通过的记录，并单独标注为金额差异。",
        ],
    }
    (OUT / "正式范围明细控制数.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    conn.close()


if __name__ == "__main__":
    main()
