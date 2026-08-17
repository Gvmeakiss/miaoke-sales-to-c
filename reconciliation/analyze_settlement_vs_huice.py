#!/usr/bin/env python3
"""核对结算中心发货对账明细与惠策账单明细。

强键为双方原始字段“平台订单号”。分析区分：
1. 订单号能否串联；
2. 串联后金额是否一致；
3. 同月匹配与跨期匹配；
4. 正式2026年1—6月范围与向前追溯2025年12月范围。

脚本只读取客户原始文件及现有SQLite标准化库，临时聚合库在退出时自动删除。
"""

from __future__ import annotations

import argparse
import csv
import json
import math
import re
import sqlite3
import zipfile
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path

from refresh_settlement_oms_workpaper import (
    COLS,
    discover,
    file_digest,
    parse_row,
    rows_from_stream,
    signature,
)

ROOT = Path(__file__).resolve().parents[1]
SOURCE_DB = ROOT / "reconciliation/work/reconciliation.db"
OUTPUT = ROOT / "reconciliation/results/settlement_vs_huice_reconciliation.json"
EXCEPTION_CSV = ROOT / "outputs/sales_toc_workpaper_final_20260101_20260630/发货对账-惠策账单差异及单边明细_20260101-20260630.csv"
ANALYSIS_DB = ROOT / "reconciliation/work/settlement_huice_analysis.db"
KEEP_MONTHS = {"2025-12", *{f"2026-{m:02d}" for m in range(1, 8)}}
FORMAL_MONTHS = {f"2026-{m:02d}" for m in range(1, 7)}
FORMAL_WITH_DEC = {"2025-12", *FORMAL_MONTHS}


def number(value: object) -> float:
    try:
        result = float(value or 0)
        return result if math.isfinite(result) else 0.0
    except (TypeError, ValueError):
        return 0.0


def norm(value: object) -> str:
    text = str(value or "").strip().replace("\u200b", "")
    return text[:-2] if text.endswith(".0") and text[:-2].isdigit() else text


def month_of(value: object) -> str:
    text = str(value or "").strip()
    match = re.search(r"(20\d{2})[-/.年](\d{1,2})", text)
    if match:
        return f"{int(match.group(1)):04d}-{int(match.group(2)):02d}"
    try:
        serial = float(text)
        if 40000 <= serial <= 60000:
            return (datetime(1899, 12, 30) + timedelta(days=serial)).strftime("%Y-%m")
    except ValueError:
        pass
    return ""


def configure(connection: sqlite3.Connection) -> None:
    connection.executescript(
        """
        PRAGMA journal_mode=OFF;
        PRAGMA synchronous=OFF;
        PRAGMA temp_store=FILE;
        PRAGMA cache_size=-524288;
        PRAGMA locking_mode=EXCLUSIVE;
        """
    )


def build_huice(connection: sqlite3.Connection) -> None:
    print("[1/4] 聚合惠策账单至平台订单号及导出月份粒度", flush=True)
    connection.execute("ATTACH DATABASE ? AS src", (str(SOURCE_DB),))
    connection.executescript(
        """
        CREATE TABLE huice_order AS
        SELECT h.platform_order_no,
               GROUP_CONCAT(DISTINCT h.platform) AS huice_platforms,
               GROUP_CONCAT(DISTINCT h.shop) AS huice_shops,
               GROUP_CONCAT(DISTINCT CASE
                 WHEN h.source_file LIKE '%1月%' THEN '2026-01'
                 WHEN h.source_file LIKE '%2月%' THEN '2026-02'
                 WHEN h.source_file LIKE '%3月%' THEN '2026-03'
                 WHEN h.source_file LIKE '%4月%' THEN '2026-04'
                 WHEN h.source_file LIKE '%5月%' THEN '2026-05'
                 WHEN h.source_file LIKE '%6月%' THEN '2026-06'
               END) AS huice_bill_months,
               COUNT(*) AS huice_rows,
               SUM(c.current_cash) AS huice_cash,
               SUM(CASE WHEN h.reconcile_status='对账成功' THEN c.current_cash ELSE 0 END) AS huice_success_cash,
               SUM(CASE WHEN h.reconcile_status='对账失败' THEN c.current_cash ELSE 0 END) AS huice_failed_cash,
               SUM(CASE WHEN h.reconcile_status='对账成功' THEN 1 ELSE 0 END) AS huice_success_rows,
               SUM(CASE WHEN h.reconcile_status='对账失败' THEN 1 ELSE 0 END) AS huice_failed_rows
        FROM src.huice_detail h
        JOIN src.huice_current_amount c ON c.reconcile_id=h.reconcile_id
        WHERE NULLIF(TRIM(h.platform_order_no),'') IS NOT NULL
        GROUP BY h.platform_order_no;
        CREATE UNIQUE INDEX idx_huice_order ON huice_order(platform_order_no);

        CREATE TABLE huice_order_month AS
        SELECT CASE
                 WHEN h.source_file LIKE '%1月%' THEN '2026-01'
                 WHEN h.source_file LIKE '%2月%' THEN '2026-02'
                 WHEN h.source_file LIKE '%3月%' THEN '2026-03'
                 WHEN h.source_file LIKE '%4月%' THEN '2026-04'
                 WHEN h.source_file LIKE '%5月%' THEN '2026-05'
                 WHEN h.source_file LIKE '%6月%' THEN '2026-06'
               END AS bill_month,
               h.platform_order_no,
               COUNT(*) AS huice_rows,
               SUM(c.current_cash) AS huice_cash,
               SUM(CASE WHEN h.reconcile_status='对账成功' THEN c.current_cash ELSE 0 END) AS huice_success_cash
        FROM src.huice_detail h
        JOIN src.huice_current_amount c ON c.reconcile_id=h.reconcile_id
        WHERE NULLIF(TRIM(h.platform_order_no),'') IS NOT NULL
        GROUP BY 1,h.platform_order_no;
        CREATE UNIQUE INDEX idx_huice_order_month ON huice_order_month(bill_month,platform_order_no);
        """
    )
    connection.commit()
    print("      惠策订单聚合完成", flush=True)


def merge_file_aggregate(
    connection: sqlite3.Connection,
    aggregates: dict[tuple[str, str], list[float]],
) -> None:
    sql = """
        INSERT INTO settlement_order_month(
          settlement_month,platform_order_no,all_rows,success_rows,failed_rows,all_amount,success_amount,failed_amount
        ) VALUES(?,?,?,?,?,?,?,?)
        ON CONFLICT(settlement_month,platform_order_no) DO UPDATE SET
          all_rows=all_rows+excluded.all_rows,
          success_rows=success_rows+excluded.success_rows,
          failed_rows=failed_rows+excluded.failed_rows,
          all_amount=all_amount+excluded.all_amount,
          success_amount=success_amount+excluded.success_amount,
          failed_amount=failed_amount+excluded.failed_amount
    """
    connection.executemany(
        sql,
        (
            (month, order_no, int(v[0]), int(v[1]), int(v[2]), v[3], v[4], v[5])
            for (month, order_no), v in aggregates.items()
        ),
    )
    connection.commit()


def load_settlement(connection: sqlite3.Connection) -> dict[str, object]:
    print("[2/4] 流式读取发货对账明细并按平台订单号聚合", flush=True)
    connection.executescript(
        """
        CREATE TABLE settlement_order_month(
          settlement_month TEXT NOT NULL,
          platform_order_no TEXT NOT NULL,
          all_rows INTEGER,
          success_rows INTEGER,
          failed_rows INTEGER,
          all_amount REAL,
          success_amount REAL,
          failed_amount REAL,
          PRIMARY KEY(settlement_month,platform_order_no)
        ) WITHOUT ROWID;
        """
    )
    files, _ = discover()
    previous_hashes: dict[str, str] = {}
    prior_signatures: set[int] = set()
    active_group = None
    file_stats = []
    blank = defaultdict(lambda: {"rows": 0, "success_rows": 0, "amount": 0.0, "success_amount": 0.0})
    for file_no, info in enumerate(files, start=1):
        path = info["path"]
        if info["start"] != active_group:
            prior_signatures.clear()
            active_group = info["start"]
        digest = file_digest(path)
        if digest in previous_hashes:
            file_stats.append({"file": path.name, "status": "文件内容重复，未纳入", "duplicate_of": previous_hashes[digest]})
            continue
        previous_hashes[digest] = path.name
        aggregates: dict[tuple[str, str], list[float]] = defaultdict(lambda: [0, 0, 0, 0.0, 0.0, 0.0])
        current_signatures: set[int] = set()
        stats = {"file": path.name, "status": "已纳入", "rows": 0, "kept_rows": 0, "duplicate_rows_removed": 0}
        with zipfile.ZipFile(path) as archive:
            sheet_names = sorted(
                (name for name in archive.namelist() if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", name)),
                key=lambda name: int(re.search(r"sheet(\d+)", name).group(1)),
            )
            for sheet_name in sheet_names:
                with archive.open(sheet_name) as stream:
                    rows = rows_from_stream(stream)
                    try:
                        next(rows)
                    except StopIteration:
                        continue
                    for xml in rows:
                        stats["rows"] += 1
                        row = parse_row(xml)
                        month = month_of(row.get("C"))
                        if month not in KEEP_MONTHS:
                            continue
                        sig = signature(row)
                        if sig in prior_signatures:
                            stats["duplicate_rows_removed"] += 1
                            continue
                        current_signatures.add(sig)
                        stats["kept_rows"] += 1
                        order_no = norm(row.get("E"))
                        amount = number(row.get("W"))
                        success = row.get("P") == "对账成功"
                        if not order_no:
                            blank[month]["rows"] += 1
                            blank[month]["amount"] += amount
                            if success:
                                blank[month]["success_rows"] += 1
                                blank[month]["success_amount"] += amount
                            continue
                        target = aggregates[(month, order_no)]
                        target[0] += 1
                        target[3] += amount
                        if success:
                            target[1] += 1
                            target[4] += amount
                        else:
                            target[2] += 1
                            target[5] += amount
        merge_file_aggregate(connection, aggregates)
        prior_signatures.update(current_signatures)
        file_stats.append(stats)
        print(
            f"      {file_no}/{len(files)} {path.name}: 原始{stats['rows']:,}行，纳入{stats['kept_rows']:,}行，"
            f"订单月键{len(aggregates):,}",
            flush=True,
        )
    connection.execute("CREATE INDEX idx_settlement_order ON settlement_order_month(platform_order_no)")
    connection.commit()
    return {"files": file_stats, "blank_platform_order": dict(blank)}


def scalar(connection: sqlite3.Connection, sql: str, params: tuple = ()) -> float:
    result = connection.execute(sql, params).fetchone()[0]
    return result or 0


def scope_result(connection: sqlite3.Connection, name: str, months: set[str]) -> dict[str, object]:
    placeholders = ",".join("?" for _ in months)
    params = tuple(sorted(months))
    connection.execute("DROP TABLE IF EXISTS scope_settlement")
    connection.execute(
        f"""
        CREATE TEMP TABLE scope_settlement AS
        SELECT platform_order_no,SUM(all_rows) all_rows,SUM(success_rows) success_rows,
               SUM(failed_rows) failed_rows,SUM(all_amount) all_amount,
               SUM(success_amount) success_amount,SUM(failed_amount) failed_amount
        FROM settlement_order_month
        WHERE settlement_month IN ({placeholders})
        GROUP BY platform_order_no
        """,
        params,
    )
    connection.execute("CREATE UNIQUE INDEX temp.idx_scope_settlement ON scope_settlement(platform_order_no)")
    row = connection.execute(
        """
        SELECT
          COUNT(*) settlement_orders,
          SUM(CASE WHEN s.success_rows>0 THEN 1 ELSE 0 END) settlement_success_orders,
          SUM(s.success_amount) settlement_success_amount,
          SUM(ABS(s.success_amount)) settlement_success_abs_amount,
          SUM(CASE WHEN h.platform_order_no IS NOT NULL THEN 1 ELSE 0 END) common_orders,
          SUM(CASE WHEN s.success_rows>0 AND h.platform_order_no IS NOT NULL THEN 1 ELSE 0 END) common_success_orders,
          SUM(CASE WHEN h.platform_order_no IS NOT NULL THEN s.success_amount ELSE 0 END) common_settlement_amount,
          SUM(CASE WHEN h.platform_order_no IS NOT NULL THEN h.huice_cash ELSE 0 END) common_huice_cash,
          SUM(CASE WHEN h.platform_order_no IS NOT NULL THEN h.huice_success_cash ELSE 0 END) common_huice_success_cash,
          SUM(CASE WHEN s.success_rows>0 AND h.platform_order_no IS NOT NULL AND ABS(s.success_amount-h.huice_cash)<=0.01 THEN 1 ELSE 0 END) exact_all_status_orders,
          SUM(CASE WHEN s.success_rows>0 AND h.platform_order_no IS NOT NULL AND ABS(s.success_amount-h.huice_success_cash)<=0.01 THEN 1 ELSE 0 END) exact_success_status_orders,
          SUM(CASE WHEN h.platform_order_no IS NOT NULL AND ABS(s.success_amount-h.huice_cash)<=0.01 THEN ABS(s.success_amount) ELSE 0 END) exact_all_status_abs_amount,
          SUM(CASE WHEN h.platform_order_no IS NOT NULL AND ABS(s.success_amount-h.huice_success_cash)<=0.01 THEN ABS(s.success_amount) ELSE 0 END) exact_success_status_abs_amount,
          SUM(CASE WHEN h.platform_order_no IS NULL THEN 1 ELSE 0 END) settlement_only_orders,
          SUM(CASE WHEN h.platform_order_no IS NULL THEN s.success_amount ELSE 0 END) settlement_only_amount
        FROM scope_settlement s
        LEFT JOIN huice_order h ON h.platform_order_no=s.platform_order_no
        """
    ).fetchone()
    keys = [
        "settlement_orders", "settlement_success_orders", "settlement_success_amount", "settlement_success_abs_amount",
        "common_orders", "common_success_orders", "common_settlement_amount", "common_huice_cash", "common_huice_success_cash",
        "exact_all_status_orders", "exact_success_status_orders", "exact_all_status_abs_amount", "exact_success_status_abs_amount",
        "settlement_only_orders", "settlement_only_amount",
    ]
    result = {key: (value or 0) for key, value in zip(keys, row)}
    huice = connection.execute(
        """
        SELECT COUNT(*) huice_orders,SUM(huice_cash) huice_cash,
               SUM(CASE WHEN s.platform_order_no IS NOT NULL THEN 1 ELSE 0 END) huice_common_orders,
               SUM(CASE WHEN s.platform_order_no IS NOT NULL THEN h.huice_cash ELSE 0 END) huice_common_cash
        FROM huice_order h LEFT JOIN scope_settlement s ON s.platform_order_no=h.platform_order_no
        """
    ).fetchone()
    result.update(dict(zip(["huice_orders", "huice_cash", "huice_common_orders", "huice_common_cash"], (x or 0 for x in huice))))
    result["scope"] = name
    result["months"] = sorted(months)
    result["settlement_order_hit_rate"] = result["common_orders"] / result["settlement_orders"] if result["settlement_orders"] else None
    result["huice_order_hit_rate"] = result["huice_common_orders"] / result["huice_orders"] if result["huice_orders"] else None
    result["exact_amount_rate_all_huice_status"] = result["exact_all_status_orders"] / result["common_success_orders"] if result["common_success_orders"] else None
    result["exact_amount_rate_success_huice_status"] = result["exact_success_status_orders"] / result["common_success_orders"] if result["common_success_orders"] else None
    result["exact_amount_coverage_all_huice_status"] = result["exact_all_status_abs_amount"] / result["settlement_success_abs_amount"] if result["settlement_success_abs_amount"] else None
    result["exact_amount_coverage_success_huice_status"] = result["exact_success_status_abs_amount"] / result["settlement_success_abs_amount"] if result["settlement_success_abs_amount"] else None
    return result


def same_month_result(connection: sqlite3.Connection) -> list[dict[str, object]]:
    rows = connection.execute(
        """
        SELECT s.settlement_month,
               COUNT(*) settlement_success_orders,
               SUM(s.success_amount) settlement_success_amount,
               SUM(CASE WHEN h.platform_order_no IS NOT NULL THEN 1 ELSE 0 END) common_orders,
               SUM(CASE WHEN h.platform_order_no IS NOT NULL THEN s.success_amount ELSE 0 END) common_settlement_amount,
               SUM(CASE WHEN h.platform_order_no IS NOT NULL THEN h.huice_cash ELSE 0 END) common_huice_cash,
               SUM(CASE WHEN s.success_rows>0 AND h.platform_order_no IS NOT NULL AND ABS(s.success_amount-h.huice_cash)<=0.01 THEN 1 ELSE 0 END) exact_orders,
               SUM(CASE WHEN s.success_rows>0 THEN 1 ELSE 0 END) success_orders,
               SUM(CASE WHEN s.success_rows>0 AND h.platform_order_no IS NOT NULL THEN 1 ELSE 0 END) common_success_orders
        FROM settlement_order_month s
        LEFT JOIN huice_order_month h
          ON h.bill_month=s.settlement_month AND h.platform_order_no=s.platform_order_no
        WHERE s.settlement_month BETWEEN '2026-01' AND '2026-06' AND s.success_rows>0
        GROUP BY s.settlement_month ORDER BY s.settlement_month
        """
    ).fetchall()
    result = []
    for row in rows:
        keys = ["month", "settlement_success_orders", "settlement_success_amount", "common_orders", "common_settlement_amount", "common_huice_cash", "exact_orders", "success_orders", "common_success_orders"]
        item = dict(zip(keys, row))
        item["order_hit_rate"] = item["common_orders"] / item["settlement_success_orders"] if item["settlement_success_orders"] else None
        item["success_order_hit_rate"] = item["common_success_orders"] / item["success_orders"] if item["success_orders"] else None
        item["exact_amount_rate"] = item["exact_orders"] / item["common_success_orders"] if item["common_success_orders"] else None
        result.append(item)
    return result


def build_formal_reconciliation(connection: sqlite3.Connection) -> None:
    """建立正式2026年1—6月成功发货对账订单与惠策全状态净实收的订单级核对表。"""
    connection.executescript(
        """
        DROP TABLE IF EXISTS formal_settlement_order;
        CREATE TABLE formal_settlement_order AS
        SELECT platform_order_no,
               GROUP_CONCAT(DISTINCT settlement_month) AS settlement_months,
               SUM(success_rows) AS settlement_rows,
               SUM(success_amount) AS settlement_amount
        FROM settlement_order_month
        WHERE settlement_month BETWEEN '2026-01' AND '2026-06'
        GROUP BY platform_order_no
        HAVING SUM(success_rows)>0;
        CREATE UNIQUE INDEX idx_formal_settlement_order ON formal_settlement_order(platform_order_no);

        DROP TABLE IF EXISTS settlement_huice_recon;
        CREATE TABLE settlement_huice_recon AS
        SELECT CASE WHEN h.platform_order_no IS NULL THEN '仅发货对账'
                    WHEN ABS(s.settlement_amount-h.huice_success_cash)<=0.01 THEN '金额一致'
                    ELSE '金额差异' END AS result,
               s.platform_order_no,
               s.settlement_months,
               COALESCE(h.huice_bill_months,'') AS huice_bill_months,
               COALESCE(h.huice_platforms,'') AS huice_platforms,
               COALESCE(h.huice_shops,'') AS huice_shops,
               s.settlement_rows,
               COALESCE(h.huice_rows,0) AS huice_rows,
               COALESCE(h.huice_success_rows,0) AS huice_success_rows,
               COALESCE(h.huice_failed_rows,0) AS huice_failed_rows,
               s.settlement_amount,
               COALESCE(h.huice_success_cash,0) AS huice_success_cash,
               COALESCE(h.huice_cash,0) AS huice_all_status_cash,
               s.settlement_amount-COALESCE(h.huice_success_cash,0) AS amount_difference,
               CASE WHEN h.platform_order_no IS NULL THEN NULL
                    WHEN MAX(ABS(s.settlement_amount),ABS(h.huice_success_cash))=0 THEN 1.0
                    ELSE MIN(ABS(s.settlement_amount),ABS(h.huice_success_cash))/MAX(ABS(s.settlement_amount),ABS(h.huice_success_cash)) END AS amount_match_rate,
               CASE WHEN h.platform_order_no IS NULL THEN '发货对账平台订单号在惠策账单中不存在'
                    WHEN ABS(s.settlement_amount-h.huice_success_cash)<=0.01 THEN '平台订单号一致且双方对账成功金额差不超过0.01元'
                    WHEN s.settlement_months<>COALESCE(h.huice_bill_months,'') THEN '平台订单号一致但金额存在差异；月份不同，需结合跨期结算复核'
                    ELSE '平台订单号一致但订单汇总金额存在差异' END AS reason
        FROM formal_settlement_order s
        LEFT JOIN huice_order h ON h.platform_order_no=s.platform_order_no
        UNION ALL
        SELECT '仅惠策账单',h.platform_order_no,'',h.huice_bill_months,h.huice_platforms,h.huice_shops,
               0,h.huice_rows,h.huice_success_rows,h.huice_failed_rows,
               0,h.huice_success_cash,h.huice_cash,-h.huice_success_cash,NULL,
               '惠策账单平台订单号在2026年1—6月发货对账成功订单中不存在'
        FROM huice_order h
        LEFT JOIN formal_settlement_order s ON s.platform_order_no=h.platform_order_no
        WHERE s.platform_order_no IS NULL;
        CREATE INDEX idx_settlement_huice_result ON settlement_huice_recon(result);
        CREATE INDEX idx_settlement_huice_order ON settlement_huice_recon(platform_order_no);
        """
    )
    connection.commit()


def category_result(connection: sqlite3.Connection) -> list[dict[str, object]]:
    rows = connection.execute(
        """
        SELECT result,COUNT(*) order_count,SUM(settlement_rows) settlement_rows,SUM(huice_rows) huice_rows,
               SUM(settlement_amount) settlement_amount,SUM(huice_success_cash) huice_success_cash,
               SUM(huice_all_status_cash) huice_all_status_cash,
               SUM(amount_difference) amount_difference
        FROM settlement_huice_recon
        GROUP BY result
        ORDER BY CASE result WHEN '金额一致' THEN 1 WHEN '金额差异' THEN 2 WHEN '仅发货对账' THEN 3 ELSE 4 END
        """
    ).fetchall()
    result = []
    for row in rows:
        item = dict(zip(["result", "order_count", "settlement_rows", "huice_rows", "settlement_amount", "huice_success_cash", "huice_all_status_cash", "amount_difference"], row))
        denominator = max(abs(item["settlement_amount"] or 0), abs(item["huice_success_cash"] or 0))
        item["amount_match_rate"] = min(abs(item["settlement_amount"] or 0), abs(item["huice_success_cash"] or 0)) / denominator if denominator else None
        result.append(item)
    return result


def export_exception_detail(connection: sqlite3.Connection, path: Path) -> dict[str, object]:
    headers = [
        "result", "platform_order_no", "settlement_months", "huice_bill_months", "huice_platforms", "huice_shops",
        "settlement_rows", "huice_rows", "huice_success_rows", "huice_failed_rows", "settlement_amount", "huice_success_cash",
        "huice_all_status_cash",
        "amount_difference", "amount_match_rate", "reason",
    ]
    sql = """
        SELECT result,platform_order_no,settlement_months,huice_bill_months,huice_platforms,huice_shops,
               settlement_rows,huice_rows,huice_success_rows,huice_failed_rows,settlement_amount,huice_success_cash,
               huice_all_status_cash,
               amount_difference,amount_match_rate,reason
        FROM settlement_huice_recon
        WHERE result<>'金额一致'
        ORDER BY CASE result WHEN '金额差异' THEN 1 WHEN '仅发货对账' THEN 2 ELSE 3 END,
                 ABS(amount_difference) DESC,platform_order_no
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    row_count = 0
    settlement_amount = 0.0
    huice_cash = 0.0
    preview_rows = []
    with path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.writer(stream)
        writer.writerow(headers)
        cursor = connection.execute(sql)
        while True:
            batch = cursor.fetchmany(10000)
            if not batch:
                break
            writer.writerows(batch)
            for row in batch:
                row_count += 1
                settlement_amount += number(row[10])
                huice_cash += number(row[11])
                if len(preview_rows) < 10000:
                    preview_rows.append(list(row))
    return {
        "file": path.name,
        "row_count": row_count,
        "settlement_amount": settlement_amount,
        "huice_cash": huice_cash,
        "headers": headers,
        "preview_rows": preview_rows,
        "preview_limit": 10000,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", type=Path, default=OUTPUT)
    parser.add_argument("--exception-csv", type=Path, default=EXCEPTION_CSV)
    parser.add_argument("--analysis-db", type=Path, default=ANALYSIS_DB)
    args = parser.parse_args()
    if not SOURCE_DB.exists():
        raise SystemExit(f"标准化数据库不存在：{SOURCE_DB}")
    args.analysis_db.parent.mkdir(parents=True, exist_ok=True)
    args.analysis_db.unlink(missing_ok=True)
    try:
        connection = sqlite3.connect(str(args.analysis_db))
        configure(connection)
        build_huice(connection)
        source_control = load_settlement(connection)
        print("[3/4] 计算订单号覆盖率、金额一致率及跨期影响", flush=True)
        build_formal_reconciliation(connection)
        scopes = [
            scope_result(connection, "正式范围：发货对账账期2026-01至2026-06", FORMAL_MONTHS),
            scope_result(connection, "跨期参考：发货对账账期2025-12至2026-06", FORMAL_WITH_DEC),
        ]
        categories = category_result(connection)
        exception_detail = export_exception_detail(connection, args.exception_csv)
        result = {
            "generated_at": datetime.now().isoformat(timespec="seconds"),
            "sources": {
                "settlement": "input/对账明细（to oms 月结）/发货对账明细*.xlsx",
                "huice": "input/惠策系统对账单清单/历史账期对账结果明细*.xlsx（含6月-3补充文件）",
            },
            "grain": {
                "settlement": "原始一行一条发货/退款商品明细；核对时按账期结束月份+平台订单号汇总",
                "huice": "原始一行一条对账结果；核对时按导出月份+平台订单号汇总",
            },
            "keys": {
                "strong_key": "发货对账明细.平台订单号 = 惠策账单明细.平台订单号",
                "invalid_candidate": "发货对账明细.汇总单号（FHDZS前缀）不等于惠策.对账流水号（DZ前缀）",
                "item_level_limit": "惠策账单明细无商品编码及数量，不能直接执行商品行级核对",
            },
            "amounts": {
                "settlement": "发货对账明细.收款金额；主分析仅使用对账状态=对账成功的订单汇总金额",
                "huice_all_status": "惠策当前收款金额（正实收）-退款金额（负实收），包含成功及失败状态",
                "huice_success_status": "同上，但仅对账状态=对账成功",
                "exact_tolerance": "平台订单号汇总后金额差绝对值不超过0.01元",
            },
            "scopes": scopes,
            "category_summary": categories,
            "same_month": same_month_result(connection),
            "exception_detail": exception_detail,
            "source_control": source_control,
        }
        print("[4/4] 写出核对结果", flush=True)
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        connection.close()
        print(f"完成：{args.output}", flush=True)
    finally:
        args.analysis_db.unlink(missing_ok=True)


if __name__ == "__main__":
    main()
