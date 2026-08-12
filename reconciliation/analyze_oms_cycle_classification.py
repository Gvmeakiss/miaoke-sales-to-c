#!/usr/bin/env python3
"""按权威事务码表刷新OMS日/月结分类并输出控制数。"""

from __future__ import annotations

import csv
import json
import sqlite3
from datetime import datetime
from pathlib import Path

from oms_transaction_codes import (
    OMS_CODE_MAP,
    OMS_STANDARD_SETTLEMENT_CODES,
    SOURCE_FILE,
    sql_list,
)


ROOT = Path(__file__).resolve().parent
DB = ROOT / "work" / "reconciliation.db"
OUT_JSON = ROOT / "results" / "oms_cycle_classification.json"
OUT_CSV = ROOT / "results" / "oms_cycle_classification_by_code.csv"


def refresh_reference(conn: sqlite3.Connection) -> None:
    conn.executescript(
        """
        DROP VIEW IF EXISTS oms_detail_classified;
        DROP TABLE IF EXISTS oms_transaction_code_map;
        CREATE TABLE oms_transaction_code_map (
          business_type TEXT PRIMARY KEY,
          settlement_cycle TEXT NOT NULL,
          source_file TEXT NOT NULL
        );
        """
    )
    rows = [
        (code, cycle, SOURCE_FILE.name)
        for cycle, codes in OMS_CODE_MAP.items()
        for code in codes
    ]
    conn.executemany("INSERT INTO oms_transaction_code_map VALUES (?,?,?)", rows)
    conn.executescript(
        """
        CREATE VIEW oms_detail_classified AS
        SELECT o.*, COALESCE(m.settlement_cycle,'未分类') AS settlement_cycle
        FROM oms_detail o
        LEFT JOIN oms_transaction_code_map m ON m.business_type=o.business_type;
        """
    )
    conn.commit()


def main() -> None:
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    refresh_reference(conn)
    records = [dict(row) for row in conn.execute(
        """
        SELECT business_type,settlement_cycle,COUNT(*) row_count,
          COUNT(DISTINCT document_no) document_count,
          SUM(item_num) quantity,SUM(share_amount) amount
        FROM oms_detail_classified
        WHERE outbound_time>='2026-01-01' AND outbound_time<'2026-07-01'
        GROUP BY business_type,settlement_cycle
        ORDER BY CASE settlement_cycle WHEN '月结' THEN 1 WHEN '日结' THEN 2 ELSE 3 END,business_type
        """
    )]
    totals = {
        row["settlement_cycle"]: {
            "row_count": row["row_count"],
            "document_count": row["document_count"],
            "quantity": row["quantity"],
            "amount": row["amount"],
        }
        for row in conn.execute(
            """
            SELECT settlement_cycle,COUNT(*) row_count,COUNT(DISTINCT document_no) document_count,
              SUM(item_num) quantity,SUM(share_amount) amount
            FROM oms_detail_classified
            WHERE outbound_time>='2026-01-01' AND outbound_time<'2026-07-01'
            GROUP BY settlement_cycle
            """
        )
    }
    standard_sql = sql_list(OMS_STANDARD_SETTLEMENT_CODES)
    standard = dict(conn.execute(
        f"""
        SELECT COUNT(*) row_count,COUNT(DISTINCT document_no) document_count,
          SUM(item_num) quantity,SUM(share_amount) amount
        FROM oms_detail
        WHERE business_type IN ({standard_sql}) AND outbound_time>='2026-01-01' AND outbound_time<'2026-07-01'
        """
    ).fetchone())
    conn.close()

    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "period": "2026-01-01至2026-06-30",
        "source_file": str(SOURCE_FILE),
        "definitions": {
            "日结": list(OMS_CODE_MAP["日结"]),
            "月结": list(OMS_CODE_MAP["月结"]),
            "SAP标准发票专项核对": list(OMS_STANDARD_SETTLEMENT_CODES),
            "cycle_type_field": "OMS原字段结算类型，不作为日结/月结分类依据",
        },
        "totals": totals,
        "y001_standard_settlement": standard,
        "non_y001_month_amount": (totals.get("月结", {}).get("amount") or 0) - (standard.get("amount") or 0),
        "by_code": records,
    }
    OUT_JSON.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    with OUT_CSV.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=["business_type", "settlement_cycle", "row_count", "document_count", "quantity", "amount"])
        writer.writeheader()
        writer.writerows(records)
    print(json.dumps(payload, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
