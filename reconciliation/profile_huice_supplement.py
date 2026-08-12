#!/usr/bin/env python3
"""Profile the late-arriving Huice bill supplement before appending it."""

from __future__ import annotations

import json
import sqlite3
from pathlib import Path

import reconcile_sales_toc as base


ROOT = Path(__file__).resolve().parent.parent
INPUT = ROOT / "input" / "惠策系统对账单清单"
DB = ROOT / "reconciliation" / "work" / "reconciliation.db"
OUT = ROOT / "reconciliation" / "results" / "huice_supplement_merge_audit.json"
SUPPLEMENT = INPUT / base.HUICE_SUPPLEMENT_FILE


def main() -> None:
    regular = [path for path in base.huice_detail_files(ROOT / "input") if path.name != base.HUICE_SUPPLEMENT_FILE]
    supplement_header = base.read_header(SUPPLEMENT)
    reference_header = base.read_header(regular[0])
    required = [
        "对账流水号", "平台", "店铺", "对账状态", "账期开始日期", "账期结束日期", "业务日期",
        "平台订单号", "正应收金额", "负应收金额", "收款金额（正实收）", "退款金额（负实收）",
    ]
    columns = {name: base.header_index(supplement_header, name) for name in required}
    selected = sorted({index for index in columns.values() if index})

    conn = sqlite3.connect(DB)
    conn.execute("DROP TABLE IF EXISTS temp.supplement_profile")
    conn.execute("""
      CREATE TEMP TABLE supplement_profile(
        reconcile_id TEXT PRIMARY KEY, platform_order_no TEXT, business_date TEXT,
        current_receivable REAL, current_cash REAL
      )
    """)
    rows_seen = 0
    duplicate_rows = 0
    missing_reconcile_id = 0
    batch = []
    for _, values in base.iter_selected_rows(SUPPLEMENT, selected):
        def value(name: str) -> str:
            index = columns.get(name)
            return base.text(values.get(index)) if index else ""

        reconcile_id = value("对账流水号")
        if not reconcile_id or reconcile_id == "对账流水号":
            missing_reconcile_id += 1
            continue
        rows_seen += 1
        record = (
            reconcile_id,
            value("平台订单号"),
            value("业务日期"),
            base.as_number(value("正应收金额")) - base.as_number(value("负应收金额")),
            base.as_number(value("收款金额（正实收）")) - base.as_number(value("退款金额（负实收）")),
        )
        try:
            conn.execute("INSERT INTO supplement_profile VALUES(?,?,?,?,?)", record)
        except sqlite3.IntegrityError:
            duplicate_rows += 1
    conn.commit()

    unique_rows, order_count, receivable, cash, min_date, max_date = conn.execute("""
      SELECT COUNT(*),COUNT(DISTINCT NULLIF(platform_order_no,'')),SUM(current_receivable),SUM(current_cash),
             MIN(NULLIF(business_date,'')),MAX(NULLIF(business_date,''))
      FROM supplement_profile
    """).fetchone()
    overlap_rows, overlap_receivable, overlap_cash = conn.execute("""
      SELECT COUNT(*),SUM(s.current_receivable),SUM(s.current_cash)
      FROM supplement_profile s JOIN huice_detail h ON h.reconcile_id=s.reconcile_id
      WHERE h.source_file<>?
    """, (base.HUICE_SUPPLEMENT_FILE,)).fetchone()
    new_rows, new_receivable, new_cash = conn.execute("""
      SELECT COUNT(*),SUM(s.current_receivable),SUM(s.current_cash)
      FROM supplement_profile s LEFT JOIN huice_detail h
        ON h.reconcile_id=s.reconcile_id AND h.source_file<>?
      WHERE h.reconcile_id IS NULL
    """, (base.HUICE_SUPPLEMENT_FILE,)).fetchone()
    ingested_rows, ingested_receivable, ingested_cash = conn.execute("""
      SELECT COUNT(*),SUM(c.current_receivable),SUM(c.current_cash)
      FROM huice_detail h JOIN huice_current_amount c ON c.reconcile_id=h.reconcile_id
      WHERE h.source_file=?
    """, (base.HUICE_SUPPLEMENT_FILE,)).fetchone()

    result = {
        "supplement_file": str(SUPPLEMENT),
        "deduplication_key": "对账流水号",
        "deduplication_precedence": "现有月度惠策账单优先；补充文件仅追加未出现的对账流水号",
        "date_policy": "补充文件不执行账期或日期范围过滤；按2026年6月补充导出批次参与匹配",
        "schema": {
            "supplement_columns": len(supplement_header),
            "reference_columns": len(reference_header),
            "same_header": supplement_header == reference_header,
            "missing_required_fields": [name for name, index in columns.items() if not index],
            "additional_fields": [name for name in supplement_header if name not in reference_header],
        },
        "profile": {
            "rows_seen": rows_seen,
            "unique_reconcile_ids": unique_rows,
            "internal_duplicate_rows": duplicate_rows,
            "missing_reconcile_id_rows": missing_reconcile_id,
            "unique_platform_orders": order_count,
            "business_date_min": min_date,
            "business_date_max": max_date,
            "current_receivable": round(float(receivable or 0), 2),
            "current_cash": round(float(cash or 0), 2),
        },
        "merge_impact_against_regular_exports": {
            "overlap_existing_rows": overlap_rows or 0,
            "overlap_current_receivable": round(float(overlap_receivable or 0), 2),
            "overlap_current_cash": round(float(overlap_cash or 0), 2),
            "new_unique_rows": new_rows or 0,
            "new_current_receivable": round(float(new_receivable or 0), 2),
            "new_current_cash": round(float(new_cash or 0), 2),
        },
        "post_merge_database": {
            "ingested_supplement_rows": ingested_rows or 0,
            "ingested_current_receivable": round(float(ingested_receivable or 0), 2),
            "ingested_current_cash": round(float(ingested_cash or 0), 2),
        },
    }
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    conn.close()


if __name__ == "__main__":
    main()
