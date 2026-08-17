#!/usr/bin/env python3
"""销售 ToC 全链路核对。

输入：旺店通订单、惠策对账明细/汇总、OMS SQL、SAP 发票清单。
输出：SQLite 核对库、分项明细 CSV、供 Excel 构建器读取的 summary.json。

设计原则：
1. 大型 xlsx 使用 XML 流式读取，不一次性载入内存。
2. 旺店通商品行先压缩到内部订单，再汇总至平台原始单号。
3. OMS 通过 document_no = SAP.OMS销售单号识别已开票子集。
4. 完整明细保存在 CSV；Excel 工作簿按容量纳入全部或优先异常明细。
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import difflib
import io
import json
import os
import re
import sqlite3
import sys
import unicodedata
import zipfile
from collections import defaultdict
from concurrent.futures import ProcessPoolExecutor, as_completed
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, Iterable, Iterator, List, Optional, Sequence, Tuple

from oms_transaction_codes import OMS_CODE_MAP, SOURCE_FILE as OMS_CODE_SOURCE

HUICE_SUPPLEMENT_FILE = "历史账期对账结果明细6月-3.xlsx"

# 客户确认的SAP录入更正：仅修正用于OMS—SAP钩稽的销售单号，
# 原始SAP文件保持不变，并在sap_oms_sales_no_correction_log中保留审计轨迹。
SAP_OMS_SALES_NO_CORRECTIONS = {
    "Y83078d072e7ac5620e599_1": "Y83078d072e7ac5620e599",
}


def correct_sap_oms_sales_no(value: Optional[str]) -> str:
    raw_value = value or ""
    return SAP_OMS_SALES_NO_CORRECTIONS.get(raw_value, raw_value)


def huice_detail_files(input_dir: Path) -> List[Path]:
    """Return regular Huice exports first and the late-arriving supplement last.

    Both downstream tables use INSERT OR IGNORE on 对账流水号, so this order
    preserves the existing export when the same bill is present in both files.
    """
    files = sorted((input_dir / "惠策系统对账单清单").glob("*.xlsx"))
    return sorted(files, key=lambda path: (path.name == HUICE_SUPPLEMENT_FILE, path.name))


def stable_cache_name(path: Path) -> str:
    digest = hashlib.sha1(path.name.encode("utf-8")).hexdigest()[:10]
    return f"{path.stem}__{digest}.csv"

from lxml import etree


NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
ROW_TAG = f"{{{NS_MAIN}}}row"


def log(message: str) -> None:
    print(f"[{datetime.now():%H:%M:%S}] {message}", flush=True)


def col_number(cell_ref: str) -> int:
    match = re.match(r"([A-Z]+)", cell_ref or "")
    if not match:
        return 0
    result = 0
    for char in match.group(1):
        result = result * 26 + ord(char) - 64
    return result


def workbook_first_sheet_path(zf: zipfile.ZipFile) -> str:
    workbook = etree.fromstring(zf.read("xl/workbook.xml"))
    relationships = etree.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_map = {node.get("Id"): node.get("Target") for node in relationships}
    sheet = workbook.find(f".//{{{NS_MAIN}}}sheet")
    if sheet is None:
        raise ValueError("工作簿不存在工作表")
    relation_id = sheet.get(f"{{{NS_REL_DOC}}}id")
    target = rel_map[relation_id]
    if target.startswith("/"):
        return target.lstrip("/")
    if target.startswith("xl/"):
        return target
    return "xl/" + target.lstrip("/")


def load_shared_strings(zf: zipfile.ZipFile) -> List[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    strings: List[str] = []
    with zf.open("xl/sharedStrings.xml") as stream:
        for _, element in etree.iterparse(
            stream,
            events=("end",),
            tag=f"{{{NS_MAIN}}}si",
            huge_tree=True,
            recover=True,
        ):
            strings.append("".join(element.itertext()))
            element.clear()
            while element.getprevious() is not None:
                del element.getparent()[0]
    return strings


def cell_value(cell: etree._Element, shared_strings: Sequence[str]) -> Optional[str]:
    cell_type = cell.get("t")
    if cell_type == "inlineStr":
        return "".join(cell.itertext())
    value_node = cell.find(f"{{{NS_MAIN}}}v")
    if value_node is None:
        return None
    value = value_node.text
    if cell_type == "s" and value is not None:
        try:
            return shared_strings[int(value)]
        except (ValueError, IndexError):
            return value
    return value


def read_header(path: Path) -> List[str]:
    with zipfile.ZipFile(path) as zf:
        shared_strings = load_shared_strings(zf)
        sheet_path = workbook_first_sheet_path(zf)
        with zf.open(sheet_path) as stream:
            for _, row in etree.iterparse(
                stream, events=("end",), tag=ROW_TAG, huge_tree=True, recover=True
            ):
                cells: Dict[int, str] = {}
                for cell in row:
                    if etree.QName(cell).localname != "c":
                        continue
                    cells[col_number(cell.get("r"))] = cell_value(cell, shared_strings) or ""
                max_column = max(cells, default=0)
                return [cells.get(index, "") for index in range(1, max_column + 1)]
    return []


def iter_selected_rows(
    path: Path,
    selected_columns: Sequence[int],
    max_rows: Optional[int] = None,
) -> Iterator[Tuple[int, Dict[int, Optional[str]]]]:
    selected = set(selected_columns)
    with zipfile.ZipFile(path) as zf:
        shared_strings = load_shared_strings(zf)
        sheet_path = workbook_first_sheet_path(zf)
        with zf.open(sheet_path) as stream:
            data_rows = 0
            for _, row in etree.iterparse(
                stream, events=("end",), tag=ROW_TAG, huge_tree=True, recover=True
            ):
                row_number = int(row.get("r") or 0)
                values: Dict[int, Optional[str]] = {}
                for cell in row:
                    if etree.QName(cell).localname != "c":
                        continue
                    column = col_number(cell.get("r"))
                    if column in selected:
                        values[column] = cell_value(cell, shared_strings)
                row.clear()
                while row.getprevious() is not None:
                    del row.getparent()[0]
                if row_number == 1:
                    continue
                data_rows += 1
                yield row_number, values
                if max_rows and data_rows >= max_rows:
                    break


def header_index(header: Sequence[str], name: str, occurrence: int = 1) -> Optional[int]:
    found = 0
    for index, value in enumerate(header, start=1):
        if value == name:
            found += 1
            if found == occurrence:
                return index
    return None


def as_number(value: Optional[str]) -> float:
    if value is None:
        return 0.0
    text = str(value).strip().replace(",", "")
    if not text or text.upper() in {"NULL", "N/A", "NA"}:
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def text(value: Optional[str]) -> str:
    return "" if value is None else str(value).strip()


def month_from_date(value: str) -> str:
    if not value:
        return ""
    match = re.search(r"(20\d{2})[-/.年](\d{1,2})", value)
    if match:
        return f"{int(match.group(1)):04d}-{int(match.group(2)):02d}"
    return ""


def next_day(value: str) -> str:
    try:
        return (datetime.strptime(value[:10], "%Y-%m-%d") + timedelta(days=1)).strftime("%Y-%m-%d")
    except Exception:
        return value[:10]


def configure_database(connection: sqlite3.Connection) -> None:
    connection.executescript(
        """
        PRAGMA journal_mode=OFF;
        PRAGMA synchronous=OFF;
        PRAGMA temp_store=MEMORY;
        PRAGMA cache_size=-400000;
        PRAGMA locking_mode=EXCLUSIVE;
        """
    )


def create_schema(connection: sqlite3.Connection) -> None:
    connection.executescript(
        """
        CREATE TABLE IF NOT EXISTS stage_meta (
            stage TEXT PRIMARY KEY,
            completed_at TEXT,
            detail TEXT
        );

        CREATE TABLE IF NOT EXISTS wdt_orders_file (
            source_file TEXT,
            order_no TEXT,
            platform_order_no TEXT,
            sub_order_no TEXT,
            shop TEXT,
            order_status TEXT,
            order_type TEXT,
            trade_time TEXT,
            payment_time TEXT,
            ship_time TEXT,
            receivable_amount REAL,
            allocated_total REAL,
            quantity REAL,
            line_count INTEGER,
            schema_columns INTEGER,
            PRIMARY KEY (source_file, order_no, platform_order_no)
        );

        CREATE TABLE IF NOT EXISTS huice_detail (
            reconcile_id TEXT PRIMARY KEY,
            platform TEXT,
            shop TEXT,
            reconcile_status TEXT,
            period_start TEXT,
            period_end TEXT,
            business_date TEXT,
            platform_order_no TEXT,
            net_receivable REAL,
            net_cash REAL,
            positive_difference REAL,
            reverse_difference REAL,
            computed_difference REAL,
            reconcile_result TEXT,
            failure_result TEXT,
            source_file TEXT
        );

        CREATE TABLE IF NOT EXISTS huice_summary (
            monthly_id TEXT,
            reconcile_date TEXT,
            platform TEXT,
            shop TEXT,
            success_count REAL,
            success_amount REAL,
            mismatch_count REAL,
            mismatch_receivable REAL,
            mismatch_cash REAL,
            single_ar_count REAL,
            single_ar_amount REAL,
            single_cash_count REAL,
            single_cash_amount REAL,
            source_file TEXT,
            PRIMARY KEY (source_file, monthly_id)
        );

        CREATE TABLE IF NOT EXISTS sap2c (
            file_month TEXT,
            oms_sales_no TEXT,
            material_code TEXT,
            sales_unit TEXT,
            sap_invoice_no TEXT,
            invoice_qty REAL,
            tax_amount REAL,
            net_amount REAL,
            tax_value REAL,
            row_count INTEGER,
            PRIMARY KEY (file_month, oms_sales_no, material_code, sales_unit, sap_invoice_no)
        );

        CREATE TABLE IF NOT EXISTS oms_detail (
            item_code TEXT,
            item_name TEXT,
            item_num REAL,
            sale_unit TEXT,
            delivery_amount REAL,
            share_amount REAL,
            remark TEXT,
            business_no TEXT,
            document_no TEXT,
            business_type TEXT,
            customer_code TEXT,
            customer_name TEXT,
            outbound_time TEXT,
            outbound_month TEXT,
            warehouse_code TEXT,
            warehouse_name TEXT,
            create_time TEXT,
            cycle_type TEXT,
            PRIMARY KEY (document_no, item_code, sale_unit, warehouse_code, outbound_time)
        );

        CREATE TABLE IF NOT EXISTS customer_shop_map (
            customer_code TEXT PRIMARY KEY,
            customer_name TEXT,
            mapped_platform TEXT,
            mapped_shop TEXT,
            match_score REAL,
            mapping_status TEXT
        );
        """
    )


def refresh_oms_transaction_classification(connection: sqlite3.Connection) -> None:
    """按事务码表建立日/月结分类；OMS原字段cycle_type不用于该判断。"""
    connection.executescript(
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
    connection.executemany(
        "INSERT INTO oms_transaction_code_map VALUES (?,?,?)",
        [
            (code, cycle, OMS_CODE_SOURCE.name)
            for cycle, codes in OMS_CODE_MAP.items()
            for code in codes
        ],
    )
    connection.executescript(
        """
        CREATE VIEW oms_detail_classified AS
        SELECT o.*,COALESCE(m.settlement_cycle,'未分类') settlement_cycle
        FROM oms_detail o
        LEFT JOIN oms_transaction_code_map m ON m.business_type=o.business_type;
        """
    )
    connection.commit()


def ensure_huice_summary_composite_key(connection: sqlite3.Connection) -> None:
    """月汇总流水号仅在单个导出文件内唯一，旧单字段主键会覆盖跨文件记录。"""
    info = connection.execute("PRAGMA table_info(huice_summary)").fetchall()
    primary_key = [row[1] for row in sorted((row for row in info if row[5]), key=lambda row: row[5])]
    expected = ["source_file", "monthly_id"]
    if not info or primary_key == expected:
        return
    connection.executescript(
        """
        DROP TABLE IF EXISTS huice_summary;
        CREATE TABLE huice_summary (
            monthly_id TEXT,
            reconcile_date TEXT,
            platform TEXT,
            shop TEXT,
            success_count REAL,
            success_amount REAL,
            mismatch_count REAL,
            mismatch_receivable REAL,
            mismatch_cash REAL,
            single_ar_count REAL,
            single_ar_amount REAL,
            single_cash_count REAL,
            single_cash_amount REAL,
            source_file TEXT,
            PRIMARY KEY (source_file, monthly_id)
        );
        DELETE FROM stage_meta WHERE stage='huice_summary';
        """
    )
    connection.commit()


def ensure_wdt_order_composite_key(connection: sqlite3.Connection) -> None:
    """旺店通内部订单编号会跨多个平台原始单号复用，必须保留平台单号粒度。"""
    info = connection.execute("PRAGMA table_info(wdt_orders_file)").fetchall()
    primary_key = [row[1] for row in sorted((row for row in info if row[5]), key=lambda row: row[5])]
    expected = ["source_file", "order_no", "platform_order_no"]
    if not info or primary_key == expected:
        return
    connection.executescript(
        """
        DROP TABLE IF EXISTS wdt_orders_file;
        CREATE TABLE wdt_orders_file (
            source_file TEXT,
            order_no TEXT,
            platform_order_no TEXT,
            sub_order_no TEXT,
            shop TEXT,
            order_status TEXT,
            order_type TEXT,
            trade_time TEXT,
            payment_time TEXT,
            ship_time TEXT,
            receivable_amount REAL,
            allocated_total REAL,
            quantity REAL,
            line_count INTEGER,
            schema_columns INTEGER,
            PRIMARY KEY (source_file, order_no, platform_order_no)
        );
        DELETE FROM stage_meta WHERE stage='wdt';
        """
    )
    connection.commit()


def stage_complete(connection: sqlite3.Connection, stage: str) -> bool:
    return connection.execute("SELECT 1 FROM stage_meta WHERE stage=?", (stage,)).fetchone() is not None


def mark_stage(connection: sqlite3.Connection, stage: str, detail: str = "") -> None:
    connection.execute(
        "INSERT OR REPLACE INTO stage_meta(stage, completed_at, detail) VALUES(?,?,?)",
        (stage, datetime.now().isoformat(timespec="seconds"), detail),
    )
    connection.commit()


def wdt_files(input_dir: Path) -> List[Path]:
    files = sorted((input_dir / "旺店通订单清单").glob("*.xlsx"))
    # 归档文件与26年3月-1完全相同，保留月份文件。
    return [path for path in files if "已归档订单明细" not in path.name]


def extract_wdt_file_to_csv(task: Tuple[str, str, Optional[int]]) -> Tuple[str, str, int]:
    """并行工作进程：把一个旺店通分卷压缩到“内部订单+平台单号”级CSV。

    一个旺店通内部订单可能包含多个平台原始单号（常见于批量/手工导入单）。
    表头金额、商品数量及分摊金额必须保留到平台单号粒度，不能全部挂在首个
    平台单号下。
    """
    path_text, cache_text, max_rows = task
    path = Path(path_text)
    cache_path = Path(cache_text)
    if cache_path.exists() and cache_path.stat().st_size > 0:
        with cache_path.open("r", encoding="utf-8", newline="") as stream:
            count = sum(1 for _ in stream)
        return path.name, str(cache_path), count
    header = read_header(path)
    fields = {
        "order_no": header_index(header, "订单编号"),
        "shop": header_index(header, "店铺名称"),
        "platform_order_no": header_index(header, "原始单号"),
        "sub_order_no": header_index(header, "原始子单号"),
        "order_status": header_index(header, "订单状态"),
        "order_type": header_index(header, "订单类型"),
        "trade_time": header_index(header, "交易时间"),
        "payment_time": header_index(header, "付款时间"),
        "ship_time": header_index(header, "发货时间"),
        "receivable_amount": header_index(header, "应收金额"),
        "allocated_total": header_index(header, "分摊后总价"),
        "quantity": header_index(header, "数量"),
    }
    if not fields["order_no"]:
        return path.name, str(cache_path), 0
    selected = sorted({column for column in fields.values() if column})
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    current_order = ""
    aggregates: Dict[Tuple[str, str], Dict[str, object]] = {}
    count = 0
    with cache_path.open("w", encoding="utf-8", newline="") as output:
        writer = csv.writer(output)

        def emit() -> None:
            nonlocal count
            if not aggregates:
                return
            for aggregate in aggregates.values():
                writer.writerow(
                    [
                        path.name,
                        aggregate["order_no"], aggregate["platform_order_no"], aggregate["sub_order_no"],
                        aggregate["shop"], aggregate["order_status"], aggregate["order_type"],
                        aggregate["trade_time"], aggregate["payment_time"], aggregate["ship_time"],
                        aggregate["receivable_amount"], aggregate["allocated_total"], aggregate["quantity"],
                        aggregate["line_count"], len(header),
                    ]
                )
                count += 1
            aggregates.clear()

        def value(values: Dict[int, Optional[str]], key: str) -> str:
            column = fields[key]
            return text(values.get(column)) if column else ""

        for _, values in iter_selected_rows(path, selected, max_rows=max_rows):
            order_no = value(values, "order_no")
            if not order_no:
                continue
            if order_no != current_order:
                emit()
                current_order = order_no
            platform_order_no = value(values, "platform_order_no")
            key = (order_no, platform_order_no)
            aggregate = aggregates.get(key)
            if aggregate is None:
                aggregate = {
                    "order_no": order_no,
                    "platform_order_no": platform_order_no,
                    "sub_order_no": value(values, "sub_order_no"),
                    "shop": value(values, "shop"),
                    "order_status": value(values, "order_status"),
                    "order_type": value(values, "order_type"),
                    "trade_time": value(values, "trade_time"),
                    "payment_time": value(values, "payment_time"),
                    "ship_time": value(values, "ship_time"),
                    "receivable_amount": as_number(value(values, "receivable_amount")),
                    "allocated_total": 0.0,
                    "quantity": 0.0,
                    "line_count": 0,
                }
                aggregates[key] = aggregate
            else:
                aggregate["receivable_amount"] = max(
                    float(aggregate["receivable_amount"]),
                    as_number(value(values, "receivable_amount")),
                )
            aggregate["allocated_total"] = float(aggregate["allocated_total"]) + as_number(value(values, "allocated_total"))
            aggregate["quantity"] = float(aggregate["quantity"]) + as_number(value(values, "quantity"))
            aggregate["line_count"] = int(aggregate["line_count"]) + 1
        emit()
    return path.name, str(cache_path), count


def extract_wdt(
    connection: sqlite3.Connection,
    input_dir: Path,
    work_dir: Path,
    max_rows: Optional[int],
    workers: int,
) -> None:
    stage = "wdt"
    if stage_complete(connection, stage):
        log("跳过旺店通：已完成")
        return
    connection.execute("DELETE FROM wdt_orders_file")
    connection.commit()
    insert_sql = """
        INSERT INTO wdt_orders_file VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        ON CONFLICT(source_file,order_no,platform_order_no) DO UPDATE SET
          sub_order_no=MAX(wdt_orders_file.sub_order_no,excluded.sub_order_no),
          shop=MAX(wdt_orders_file.shop,excluded.shop),
          order_status=MAX(wdt_orders_file.order_status,excluded.order_status),
          order_type=MAX(wdt_orders_file.order_type,excluded.order_type),
          trade_time=MIN(wdt_orders_file.trade_time,excluded.trade_time),
          payment_time=MIN(wdt_orders_file.payment_time,excluded.payment_time),
          ship_time=MAX(wdt_orders_file.ship_time,excluded.ship_time),
          receivable_amount=MAX(wdt_orders_file.receivable_amount,excluded.receivable_amount),
          allocated_total=wdt_orders_file.allocated_total+excluded.allocated_total,
          quantity=wdt_orders_file.quantity+excluded.quantity,
          line_count=wdt_orders_file.line_count+excluded.line_count,
          schema_columns=MAX(wdt_orders_file.schema_columns,excluded.schema_columns)
    """
    files = wdt_files(input_dir)
    cache_dir = work_dir / "wdt_order_platform_cache"
    tasks = [
        (str(path), str(cache_dir / f"{index:02d}_{path.stem}.csv"), max_rows)
        for index, path in enumerate(files, start=1)
    ]
    results = []
    with ProcessPoolExecutor(max_workers=max(1, workers)) as executor:
        futures = {executor.submit(extract_wdt_file_to_csv, task): task[0] for task in tasks}
        for completed, future in enumerate(as_completed(futures), start=1):
            name, cache_path, count = future.result()
            results.append((name, cache_path, count))
            log(f"旺店通并行抽取 {completed}/{len(tasks)}: {name}，订单{count:,}")
    total_orders = 0
    for name, cache_path, count in sorted(results):
        batch = []
        with Path(cache_path).open("r", encoding="utf-8", newline="") as stream:
            for row in csv.reader(stream):
                batch.append(tuple(row))
                if len(batch) >= 20000:
                    connection.executemany(insert_sql, batch)
                    connection.commit()
                    batch.clear()
            if batch:
                connection.executemany(insert_sql, batch)
                connection.commit()
        total_orders += count
        log(f"旺店通缓存入库：{name}")
    mark_stage(connection, stage, f"orders_file={total_orders}")


def extract_huice_file_to_csv(task: Tuple[str, str, Optional[int]]) -> Tuple[str, str, int]:
    """并行工作进程：标准化一个惠策明细分卷。"""
    path_text, cache_text, max_rows = task
    path = Path(path_text)
    cache_path = Path(cache_text)
    if cache_path.exists() and cache_path.stat().st_size > 0:
        with cache_path.open("r", encoding="utf-8", newline="") as stream:
            count = sum(1 for _ in stream)
        return path.name, str(cache_path), count
    header = read_header(path)
    names = [
        "对账流水号", "平台", "店铺", "对账状态", "账期开始日期", "账期结束日期",
        "业务日期", "平台订单号", "正应收金额", "往期正应收金额", "负应收金额",
        "收款金额（正实收）", "往期收款金额", "退款金额（负实收）", "往期退款金额",
        "正向差异金额", "逆向差异金额", "往期负应收金额", "对账结果", "对账失败结果",
    ]
    columns = {name: header_index(header, name) for name in names}
    selected = sorted({column for column in columns.values() if column})
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    count = 0
    with cache_path.open("w", encoding="utf-8", newline="") as output:
        writer = csv.writer(output)
        for _, values in iter_selected_rows(path, selected, max_rows=max_rows):
            def val(name: str) -> str:
                column = columns[name]
                return text(values.get(column)) if column else ""
            reconcile_id = val("对账流水号")
            if not reconcile_id or reconcile_id == "对账流水号":
                continue
            net_receivable = as_number(val("正应收金额")) + as_number(val("往期正应收金额")) - as_number(val("负应收金额")) - as_number(val("往期负应收金额"))
            net_cash = as_number(val("收款金额（正实收）")) + as_number(val("往期收款金额")) - as_number(val("退款金额（负实收）")) - as_number(val("往期退款金额"))
            positive_difference = as_number(val("正向差异金额"))
            reverse_difference = as_number(val("逆向差异金额"))
            writer.writerow(
                [
                    reconcile_id, val("平台"), val("店铺"), val("对账状态"),
                    val("账期开始日期"), val("账期结束日期"), val("业务日期"), val("平台订单号"),
                    net_receivable, net_cash, positive_difference, reverse_difference,
                    net_receivable - net_cash, val("对账结果"), val("对账失败结果"), path.name,
                ]
            )
            count += 1
    return path.name, str(cache_path), count


def extract_huice_detail(
    connection: sqlite3.Connection,
    input_dir: Path,
    work_dir: Path,
    max_rows: Optional[int],
    workers: int,
) -> None:
    stage = "huice_detail"
    if stage_complete(connection, stage):
        log("跳过惠策明细：已完成")
        return
    connection.execute("DELETE FROM huice_detail")
    connection.commit()
    insert_sql = """
        INSERT OR IGNORE INTO huice_detail VALUES
        (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    """
    files = huice_detail_files(input_dir)
    cache_dir = work_dir / "huice_detail_cache"
    tasks = [(str(path), str(cache_dir / stable_cache_name(path)), max_rows) for path in files]
    results = []
    with ProcessPoolExecutor(max_workers=max(1, workers)) as executor:
        futures = {executor.submit(extract_huice_file_to_csv, task): task[0] for task in tasks}
        for completed, future in enumerate(as_completed(futures), start=1):
            name, cache_path, count = future.result()
            results.append((name, cache_path, count))
            log(f"惠策明细并行抽取 {completed}/{len(tasks)}: {name}，记录{count:,}")
    total = 0
    for name, cache_path, count in sorted(results, key=lambda item: (item[0] == HUICE_SUPPLEMENT_FILE, item[0])):
        batch = []
        with Path(cache_path).open("r", encoding="utf-8", newline="") as stream:
            for row in csv.reader(stream):
                batch.append(tuple(row))
                if len(batch) >= 20000:
                    connection.executemany(insert_sql, batch)
                    connection.commit()
                    batch.clear()
            if batch:
                connection.executemany(insert_sql, batch)
                connection.commit()
        total += count
        log(f"惠策明细缓存入库：{name}")
    mark_stage(connection, stage, f"rows_seen={total}")


def extract_huice_summary(connection: sqlite3.Connection, input_dir: Path, max_rows: Optional[int]) -> None:
    stage = "huice_summary"
    if stage_complete(connection, stage):
        log("跳过惠策汇总：已完成")
        return
    connection.execute("DELETE FROM huice_summary")
    connection.commit()
    files = sorted((input_dir / "惠策系统对账单汇总").glob("*"))
    files = [path for path in files if path.is_file() and "xlsx" in path.name.lower()]
    insert_sql = "INSERT OR REPLACE INTO huice_summary VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
    total = 0
    for file_number, path in enumerate(files, start=1):
        batch = []
        for _, values in iter_selected_rows(path, list(range(1, 14)), max_rows=max_rows):
            monthly_id = text(values.get(1))
            if not monthly_id or monthly_id == "月汇总流水号":
                continue
            batch.append(
                (
                    monthly_id,
                    text(values.get(2)),
                    text(values.get(3)),
                    text(values.get(4)),
                    as_number(values.get(5)),
                    as_number(values.get(6)),
                    as_number(values.get(7)),
                    as_number(values.get(8)),
                    as_number(values.get(9)),
                    as_number(values.get(10)),
                    as_number(values.get(11)),
                    as_number(values.get(12)),
                    as_number(values.get(13)),
                    path.name,
                )
            )
            total += 1
        if batch:
            connection.executemany(insert_sql, batch)
            connection.commit()
        log(f"惠策汇总 {file_number}/{len(files)}: {path.name}")
    mark_stage(connection, stage, f"rows={total}")


def extract_sap2c(connection: sqlite3.Connection, input_dir: Path, max_rows: Optional[int]) -> None:
    stage = "sap2c"
    if stage_complete(connection, stage):
        log("跳过SAP 2C：已完成")
        return
    connection.execute("DELETE FROM sap2c")
    connection.commit()
    files = sorted((input_dir / "发票清单：26.01.01-26.06.30").glob("*.XLSX"))
    insert_sql = """
        INSERT INTO sap2c VALUES (?,?,?,?,?,?,?,?,?,?)
        ON CONFLICT(file_month, oms_sales_no, material_code, sales_unit, sap_invoice_no)
        DO UPDATE SET
          invoice_qty=invoice_qty+excluded.invoice_qty,
          tax_amount=tax_amount+excluded.tax_amount,
          net_amount=net_amount+excluded.net_amount,
          tax_value=tax_value+excluded.tax_value,
          row_count=row_count+excluded.row_count
    """
    total = 0
    for file_number, path in enumerate(files, start=1):
        # 固定104列结构：A OMS销售单号、F SAP发票号、O 类型描述、AF物料、AQ数量、AR单位、AW/AX/AY金额。
        selected = [1, 6, 15, 32, 43, 44, 49, 50, 51]
        batch = []
        file_month = path.stem
        for _, values in iter_selected_rows(path, selected, max_rows=max_rows):
            invoice_type = text(values.get(15))
            if "标准发票（2C)" != invoice_type:
                continue
            batch.append(
                (
                    file_month,
                    text(values.get(1)),
                    text(values.get(32)),
                    text(values.get(44)),
                    text(values.get(6)),
                    as_number(values.get(43)),
                    as_number(values.get(49)),
                    as_number(values.get(50)),
                    as_number(values.get(51)),
                    1,
                )
            )
            total += 1
            if len(batch) >= 10000:
                connection.executemany(insert_sql, batch)
                connection.commit()
                batch.clear()
        if batch:
            connection.executemany(insert_sql, batch)
            connection.commit()
        log(f"SAP 2C {file_number}/{len(files)}: {path.name}")
    mark_stage(connection, stage, f"rows={total}")


def parse_sql_values(line: str) -> Optional[List[str]]:
    stripped = line.strip()
    if not stripped.startswith("VALUES ("):
        return None
    body = stripped[len("VALUES (") :]
    if body.endswith(");"):
        body = body[:-2]
    return [
        value.strip()
        for value in next(
            csv.reader(io.StringIO(body), delimiter=",", quotechar="'", skipinitialspace=True, doublequote=True)
        )
    ]


def extract_oms(connection: sqlite3.Connection, input_dir: Path, max_rows: Optional[int]) -> None:
    stage = "oms"
    if stage_complete(connection, stage):
        log("跳过OMS：已完成")
        return
    connection.execute("DELETE FROM oms_detail")
    connection.commit()
    sql_path = input_dir / "OMS 系统日结月结查询记录：25年12月到26年6月2C单据.sql"
    insert_sql = "INSERT OR REPLACE INTO oms_detail VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
    batch = []
    total = 0
    with sql_path.open("r", encoding="utf-16") as stream:
        for line in stream:
            row = parse_sql_values(line)
            if not row or len(row) != 17:
                continue
            outbound_time = text(row[12])
            batch.append(
                (
                    text(row[0]), text(row[1]), as_number(row[2]), text(row[3]),
                    as_number(row[4]), as_number(row[5]), text(row[6]), text(row[7]),
                    text(row[8]), text(row[9]), text(row[10]), text(row[11]),
                    outbound_time, month_from_date(outbound_time), text(row[13]), text(row[14]),
                    text(row[15]), text(row[16]),
                )
            )
            total += 1
            if len(batch) >= 20000:
                connection.executemany(insert_sql, batch)
                connection.commit()
                batch.clear()
            if max_rows and total >= max_rows:
                break
    if batch:
        connection.executemany(insert_sql, batch)
        connection.commit()
    mark_stage(connection, stage, f"rows={total}")


def normalize_name(value: str) -> str:
    normalized = unicodedata.normalize("NFKC", value or "").lower()
    normalized = re.sub(
        r"有限公司|有限责任公司|信息技术|电子商务|网络科技|世纪|官方|旗舰店|专卖店|食品|奶酪|妙可蓝多|蒙牛|技术|商贸",
        "",
        normalized,
    )
    normalized = re.sub(r"[\s()（）\-—_·.,，。/]+", "", normalized)
    return normalized


def platform_keyword(value: str) -> str:
    for keyword in ["天猫", "淘宝", "京东", "拼多多", "抖音", "快手", "微信", "视频号", "小红书", "唯品会"]:
        if keyword in value:
            return keyword
    return ""


def build_customer_mapping(connection: sqlite3.Connection) -> None:
    stage = "customer_mapping"
    if stage_complete(connection, stage):
        return
    connection.execute("DELETE FROM customer_shop_map")
    customers = connection.execute(
        "SELECT customer_code, MIN(customer_name) FROM oms_detail GROUP BY customer_code"
    ).fetchall()
    shops = connection.execute(
        "SELECT platform, shop FROM huice_detail WHERE shop<>'' GROUP BY platform, shop"
    ).fetchall()
    rows = []
    for customer_code, customer_name in customers:
        customer_name = customer_name or ""
        parenthetical = " ".join(re.findall(r"[（(]([^）)]+)[）)]", customer_name))
        customer_normalized = normalize_name(parenthetical or customer_name)
        customer_platform = platform_keyword(customer_name)
        best: Optional[Tuple[float, str, str]] = None
        for platform, shop in shops:
            shop_normalized = normalize_name(shop)
            if not customer_normalized or not shop_normalized:
                score = 0.0
            else:
                score = difflib.SequenceMatcher(None, customer_normalized, shop_normalized).ratio()
                if customer_normalized in shop_normalized or shop_normalized in customer_normalized:
                    score = max(score, 0.85)
            shop_platform = platform_keyword(f"{platform}{shop}")
            if customer_platform and shop_platform:
                score += 0.12 if customer_platform == shop_platform else -0.20
            if best is None or score > best[0]:
                best = (score, platform, shop)
        score, mapped_platform, mapped_shop = best or (0.0, "", "")
        if score >= 0.72:
            status = "自动高置信"
        elif score >= 0.48:
            status = "自动待复核"
        else:
            status = "未映射"
            mapped_platform = ""
            mapped_shop = ""
        rows.append((customer_code, customer_name, mapped_platform, mapped_shop, score, status))
    connection.executemany("INSERT OR REPLACE INTO customer_shop_map VALUES (?,?,?,?,?,?)", rows)
    connection.commit()
    mark_stage(connection, stage, f"rows={len(rows)}")


def materialize_analysis_tables(connection: sqlite3.Connection) -> None:
    log("建立核对中间表")
    connection.executescript(
        """
        DROP TABLE IF EXISTS wdt_order_dedup;
        CREATE TABLE wdt_order_dedup AS
        SELECT
          order_no,
          MAX(platform_order_no) AS platform_order_no,
          MAX(sub_order_no) AS sub_order_no,
          MAX(shop) AS shop,
          MAX(order_status) AS order_status,
          MAX(order_type) AS order_type,
          MIN(NULLIF(trade_time,'')) AS trade_time,
          MIN(NULLIF(payment_time,'')) AS payment_time,
          MAX(NULLIF(ship_time,'')) AS ship_time,
          MAX(receivable_amount) AS receivable_amount,
          MAX(allocated_total) AS allocated_total,
          MAX(quantity) AS quantity,
          MAX(line_count) AS line_count,
          COUNT(*) AS file_occurrences,
          GROUP_CONCAT(source_file, '|') AS source_files
        FROM wdt_orders_file
        GROUP BY order_no, platform_order_no;
        CREATE UNIQUE INDEX IF NOT EXISTS idx_wdt_order_pair ON wdt_order_dedup(order_no,platform_order_no);
        CREATE INDEX IF NOT EXISTS idx_wdt_order_platform ON wdt_order_dedup(platform_order_no);

        DROP TABLE IF EXISTS wdt_platform;
        CREATE TABLE wdt_platform AS
        SELECT
          CASE WHEN platform_order_no='' THEN '__WDT__'||order_no ELSE platform_order_no END AS platform_order_no,
          CASE WHEN platform_order_no='' THEN 0 ELSE 1 END AS matchable,
          MIN(shop) AS shop,
          COUNT(*) AS internal_order_count,
          SUM(receivable_amount) AS wdt_amount,
          SUM(allocated_total) AS wdt_allocated_amount,
          SUM(quantity) AS wdt_quantity,
          MIN(substr(COALESCE(NULLIF(payment_time,''),trade_time),1,7)) AS wdt_month,
          GROUP_CONCAT(order_no, '|') AS internal_orders
        FROM wdt_order_dedup
        GROUP BY CASE WHEN platform_order_no='' THEN '__WDT__'||order_no ELSE platform_order_no END;
        CREATE UNIQUE INDEX IF NOT EXISTS idx_wdt_platform_key ON wdt_platform(platform_order_no);

        DROP TABLE IF EXISTS huice_platform;
        CREATE TABLE huice_platform AS
        SELECT
          CASE WHEN platform_order_no='' THEN '__HC__'||reconcile_id ELSE platform_order_no END AS platform_order_no,
          MIN(platform) AS platform,
          MIN(shop) AS shop,
          COUNT(*) AS huice_row_count,
          SUM(net_receivable) AS huice_net_receivable,
          SUM(net_cash) AS huice_net_cash,
          SUM(computed_difference) AS huice_computed_difference,
          MIN(substr(COALESCE(NULLIF(business_date,''),period_end),1,7)) AS huice_month,
          GROUP_CONCAT(reconcile_id, '|') AS reconcile_ids
        FROM huice_detail
        GROUP BY CASE WHEN platform_order_no='' THEN '__HC__'||reconcile_id ELSE platform_order_no END;
        CREATE UNIQUE INDEX IF NOT EXISTS idx_huice_platform_key ON huice_platform(platform_order_no);

        DROP TABLE IF EXISTS wdt_huice_recon;
        CREATE TABLE wdt_huice_recon AS
        SELECT
          w.platform_order_no,
          w.matchable,
          w.shop AS wdt_shop,
          h.platform AS huice_platform,
          h.shop AS huice_shop,
          w.wdt_month,
          h.huice_month,
          w.internal_order_count,
          h.huice_row_count,
          w.wdt_amount,
          w.wdt_allocated_amount,
          h.huice_net_receivable,
          h.huice_net_cash,
          w.wdt_amount-h.huice_net_receivable AS amount_difference,
          w.wdt_quantity,
          w.internal_orders,
          h.reconcile_ids,
          CASE
            WHEN h.platform_order_no IS NULL AND w.matchable=0 THEN '旺店通原始单号为空'
            WHEN h.platform_order_no IS NULL THEN '仅旺店通'
            WHEN ABS(w.wdt_amount-h.huice_net_receivable)<=0.01 THEN '单号金额一致'
            ELSE '单号一致金额差异'
          END AS result
        FROM wdt_platform w
        LEFT JOIN huice_platform h ON w.platform_order_no=h.platform_order_no
        UNION ALL
        SELECT
          h.platform_order_no,
          1,
          '',
          h.platform,
          h.shop,
          '',
          h.huice_month,
          0,
          h.huice_row_count,
          0,
          0,
          h.huice_net_receivable,
          h.huice_net_cash,
          -h.huice_net_receivable,
          0,
          '',
          h.reconcile_ids,
          '仅惠策'
        FROM huice_platform h
        LEFT JOIN wdt_platform w ON w.platform_order_no=h.platform_order_no
        WHERE w.platform_order_no IS NULL;
        CREATE INDEX IF NOT EXISTS idx_wdt_huice_result ON wdt_huice_recon(result);

        DROP TABLE IF EXISTS huice_rebuilt_summary;
        CREATE TABLE huice_rebuilt_summary AS
        SELECT
          date(period_end,'+1 day') AS reconcile_date,
          platform,
          shop,
          SUM(CASE WHEN reconcile_result LIKE '%成功%' OR reconcile_status LIKE '%成功%' THEN 1 ELSE 0 END) AS success_count,
          SUM(CASE WHEN reconcile_result LIKE '%成功%' OR reconcile_status LIKE '%成功%' THEN net_receivable ELSE 0 END) AS success_amount,
          SUM(CASE WHEN reconcile_result LIKE '%不一致%' OR reconcile_status LIKE '%不一致%' THEN 1 ELSE 0 END) AS mismatch_count,
          SUM(CASE WHEN reconcile_result LIKE '%不一致%' OR reconcile_status LIKE '%不一致%' THEN net_receivable ELSE 0 END) AS mismatch_receivable,
          SUM(CASE WHEN reconcile_result LIKE '%不一致%' OR reconcile_status LIKE '%不一致%' THEN net_cash ELSE 0 END) AS mismatch_cash,
          SUM(CASE WHEN (ABS(net_receivable)>0.000001 AND ABS(net_cash)<=0.000001) THEN 1 ELSE 0 END) AS single_ar_count,
          SUM(CASE WHEN (ABS(net_receivable)>0.000001 AND ABS(net_cash)<=0.000001) THEN net_receivable ELSE 0 END) AS single_ar_amount,
          SUM(CASE WHEN (ABS(net_cash)>0.000001 AND ABS(net_receivable)<=0.000001) THEN 1 ELSE 0 END) AS single_cash_count,
          SUM(CASE WHEN (ABS(net_cash)>0.000001 AND ABS(net_receivable)<=0.000001) THEN net_cash ELSE 0 END) AS single_cash_amount
        FROM huice_detail
        GROUP BY date(period_end,'+1 day'), platform, shop;

        DROP TABLE IF EXISTS huice_summary_recon;
        CREATE TABLE huice_summary_recon AS
        WITH source AS (
          SELECT reconcile_date,platform,shop,
            SUM(success_count) success_count,SUM(success_amount) success_amount,
            SUM(mismatch_count) mismatch_count,SUM(mismatch_receivable) mismatch_receivable,SUM(mismatch_cash) mismatch_cash,
            SUM(single_ar_count) single_ar_count,SUM(single_ar_amount) single_ar_amount,
            SUM(single_cash_count) single_cash_count,SUM(single_cash_amount) single_cash_amount
          FROM huice_summary GROUP BY reconcile_date,platform,shop
        )
        SELECT
          s.reconcile_date,s.platform,s.shop,
          s.success_count source_success_count,r.success_count rebuilt_success_count,s.success_count-r.success_count diff_success_count,
          s.success_amount source_success_amount,r.success_amount rebuilt_success_amount,s.success_amount-r.success_amount diff_success_amount,
          s.mismatch_count source_mismatch_count,r.mismatch_count rebuilt_mismatch_count,s.mismatch_count-r.mismatch_count diff_mismatch_count,
          s.mismatch_receivable source_mismatch_receivable,r.mismatch_receivable rebuilt_mismatch_receivable,s.mismatch_receivable-r.mismatch_receivable diff_mismatch_receivable,
          s.mismatch_cash source_mismatch_cash,r.mismatch_cash rebuilt_mismatch_cash,s.mismatch_cash-r.mismatch_cash diff_mismatch_cash,
          s.single_ar_count source_single_ar_count,r.single_ar_count rebuilt_single_ar_count,s.single_ar_count-r.single_ar_count diff_single_ar_count,
          s.single_ar_amount source_single_ar_amount,r.single_ar_amount rebuilt_single_ar_amount,s.single_ar_amount-r.single_ar_amount diff_single_ar_amount,
          s.single_cash_count source_single_cash_count,r.single_cash_count rebuilt_single_cash_count,s.single_cash_count-r.single_cash_count diff_single_cash_count,
          s.single_cash_amount source_single_cash_amount,r.single_cash_amount rebuilt_single_cash_amount,s.single_cash_amount-r.single_cash_amount diff_single_cash_amount,
          CASE WHEN
            ABS(s.success_count-r.success_count)<=0.001 AND ABS(s.success_amount-r.success_amount)<=0.01 AND
            ABS(s.mismatch_count-r.mismatch_count)<=0.001 AND ABS(s.mismatch_receivable-r.mismatch_receivable)<=0.01 AND ABS(s.mismatch_cash-r.mismatch_cash)<=0.01 AND
            ABS(s.single_ar_count-r.single_ar_count)<=0.001 AND ABS(s.single_ar_amount-r.single_ar_amount)<=0.01 AND
            ABS(s.single_cash_count-r.single_cash_count)<=0.001 AND ABS(s.single_cash_amount-r.single_cash_amount)<=0.01
          THEN '一致' ELSE '差异' END AS result
        FROM source s LEFT JOIN huice_rebuilt_summary r
          ON s.reconcile_date=r.reconcile_date AND s.platform=r.platform AND s.shop=r.shop
        UNION ALL
        SELECT
          r.reconcile_date,r.platform,r.shop,
          0,r.success_count,-r.success_count,0,r.success_amount,-r.success_amount,
          0,r.mismatch_count,-r.mismatch_count,0,r.mismatch_receivable,-r.mismatch_receivable,
          0,r.mismatch_cash,-r.mismatch_cash,0,r.single_ar_count,-r.single_ar_count,
          0,r.single_ar_amount,-r.single_ar_amount,0,r.single_cash_count,-r.single_cash_count,
          0,r.single_cash_amount,-r.single_cash_amount,'仅明细重建'
        FROM huice_rebuilt_summary r LEFT JOIN source s
          ON s.reconcile_date=r.reconcile_date AND s.platform=r.platform AND s.shop=r.shop
        WHERE s.shop IS NULL;

        DROP TABLE IF EXISTS sap_oms_sales_no_correction_log;
        CREATE TABLE sap_oms_sales_no_correction_log AS
        SELECT oms_sales_no original_oms_sales_no,
          correct_sap_oms_sales_no(oms_sales_no) corrected_oms_sales_no,
          GROUP_CONCAT(DISTINCT sap_invoice_no) sap_invoice_nos,
          SUM(row_count) affected_rows,SUM(invoice_qty) affected_quantity,
          SUM(tax_amount) affected_amount,
          '客户确认：SAP的OMS销售单号录入错误' correction_basis
        FROM sap2c
        WHERE correct_sap_oms_sales_no(oms_sales_no)<>oms_sales_no
        GROUP BY oms_sales_no,correct_sap_oms_sales_no(oms_sales_no);

        DROP TABLE IF EXISTS sap2c_key;
        CREATE TABLE sap2c_key AS
        SELECT correct_sap_oms_sales_no(oms_sales_no) oms_sales_no,material_code,sales_unit,
          SUM(invoice_qty) invoice_qty,SUM(tax_amount) sap_amount,SUM(net_amount) sap_net_amount,SUM(tax_value) sap_tax,
          SUM(row_count) sap_rows,MIN(file_month) file_month,GROUP_CONCAT(DISTINCT sap_invoice_no) sap_invoice_nos
        FROM sap2c GROUP BY correct_sap_oms_sales_no(oms_sales_no),material_code,sales_unit;
        CREATE INDEX IF NOT EXISTS idx_sap_doc ON sap2c_key(oms_sales_no);

        DROP TABLE IF EXISTS oms_key;
        CREATE TABLE oms_key AS
        SELECT document_no,item_code,sale_unit,
          SUM(item_num) oms_qty,SUM(share_amount) oms_amount,COUNT(*) oms_rows,
          MIN(outbound_month) outbound_month,MIN(customer_code) customer_code,MIN(customer_name) customer_name
        FROM oms_detail GROUP BY document_no,item_code,sale_unit;
        CREATE INDEX IF NOT EXISTS idx_oms_doc ON oms_key(document_no);

        DROP TABLE IF EXISTS oms_sap_recon;
        CREATE TABLE oms_sap_recon AS
        SELECT
          s.oms_sales_no,s.material_code,s.sales_unit,s.file_month,o.outbound_month,
          s.sap_invoice_nos,s.sap_rows,o.oms_rows,
          s.invoice_qty,o.oms_qty,o.oms_qty-s.invoice_qty quantity_difference,
          s.sap_amount,o.oms_amount,o.oms_amount-s.sap_amount amount_difference,
          s.sap_net_amount,s.sap_tax,
          CASE
            WHEN o.document_no IS NULL THEN '仅SAP'
            WHEN ABS(o.oms_qty-s.invoice_qty)<=0.000001 AND ABS(o.oms_amount-s.sap_amount)<=0.01 THEN '数量金额一致'
            WHEN ABS(o.oms_qty-s.invoice_qty)<=0.000001 THEN '数量一致金额差异'
            ELSE '数量金额差异'
          END result
        FROM sap2c_key s LEFT JOIN oms_key o
          ON s.oms_sales_no=o.document_no AND s.material_code=o.item_code AND s.sales_unit=o.sale_unit
        UNION ALL
        SELECT
          o.document_no,o.item_code,o.sale_unit,'',o.outbound_month,'',0,o.oms_rows,
          0,o.oms_qty,o.oms_qty,0,o.oms_amount,o.oms_amount,0,0,'仅OMS'
        FROM oms_key o LEFT JOIN sap2c_key s
          ON s.oms_sales_no=o.document_no AND s.material_code=o.item_code AND s.sales_unit=o.sale_unit
        WHERE s.oms_sales_no IS NULL AND EXISTS (SELECT 1 FROM sap2c_key sd WHERE sd.oms_sales_no=o.document_no);
        CREATE INDEX IF NOT EXISTS idx_oms_sap_result ON oms_sap_recon(result);

        DROP TABLE IF EXISTS oms_huice_month_shop;
        CREATE TABLE oms_huice_month_shop AS
        WITH sap_docs AS (SELECT DISTINCT oms_sales_no FROM sap2c_key WHERE oms_sales_no<>''),
        oms_pool AS (
          SELECT o.outbound_month,m.mapped_platform,m.mapped_shop,m.mapping_status,m.match_score,
            o.customer_code,MIN(o.customer_name) customer_name,
            SUM(o.item_num) oms_qty,SUM(o.share_amount) oms_amount,COUNT(*) oms_rows
          FROM oms_detail o
          LEFT JOIN sap_docs s ON s.oms_sales_no=o.document_no
          LEFT JOIN customer_shop_map m ON m.customer_code=o.customer_code
          WHERE s.oms_sales_no IS NULL
          GROUP BY o.outbound_month,m.mapped_platform,m.mapped_shop,m.mapping_status,m.match_score,o.customer_code
        ),
        huice AS (
          SELECT substr(COALESCE(NULLIF(business_date,''),period_end),1,7) huice_month,platform,shop,
            SUM(net_receivable) huice_net_receivable,SUM(net_cash) huice_net_cash,COUNT(*) huice_rows
          FROM huice_detail GROUP BY substr(COALESCE(NULLIF(business_date,''),period_end),1,7),platform,shop
        )
        SELECT
          o.outbound_month,o.mapped_platform,o.mapped_shop,o.customer_code,o.customer_name,
          o.mapping_status,o.match_score,o.oms_rows,o.oms_qty,o.oms_amount,
          COALESCE(h.huice_rows,0) huice_rows,COALESCE(h.huice_net_receivable,0) huice_net_receivable,
          COALESCE(h.huice_net_cash,0) huice_net_cash,
          o.oms_amount-COALESCE(h.huice_net_receivable,0) amount_difference,
          CASE
            WHEN o.mapped_shop IS NULL OR o.mapped_shop='' THEN '店铺未映射'
            WHEN h.shop IS NULL THEN '仅OMS'
            WHEN ABS(o.oms_amount-h.huice_net_receivable)<=0.01 THEN '金额一致'
            ELSE '金额差异'
          END result
        FROM oms_pool o LEFT JOIN huice h
          ON h.huice_month=o.outbound_month AND h.platform=o.mapped_platform AND h.shop=o.mapped_shop;
        """
    )
    connection.commit()


def query_one(connection: sqlite3.Connection, sql: str, params: Sequence = ()) -> Tuple:
    return connection.execute(sql, params).fetchone()


def rows_as_dicts(connection: sqlite3.Connection, sql: str, params: Sequence = ()) -> List[Dict]:
    cursor = connection.execute(sql, params)
    columns = [item[0] for item in cursor.description]
    return [dict(zip(columns, row)) for row in cursor.fetchall()]


def export_query_csv(connection: sqlite3.Connection, sql: str, path: Path) -> int:
    path.parent.mkdir(parents=True, exist_ok=True)
    cursor = connection.execute(sql)
    headers = [item[0] for item in cursor.description]
    count = 0
    with path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.writer(stream)
        writer.writerow(headers)
        while True:
            rows = cursor.fetchmany(10000)
            if not rows:
                break
            writer.writerows(rows)
            count += len(rows)
    return count


def export_query_json_limited(
    connection: sqlite3.Connection, sql: str, path: Path, limit: int
) -> Dict[str, object]:
    cursor = connection.execute(f"SELECT * FROM ({sql}) LIMIT ?", (limit,))
    headers = [item[0] for item in cursor.description]
    rows = [list(row) for row in cursor.fetchall()]
    payload = {"headers": headers, "rows": rows, "included_rows": len(rows), "limit": limit}
    path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    return {"path": str(path), "included_rows": len(rows), "limit": limit}


def build_summary(connection: sqlite3.Connection, output_dir: Path) -> Dict:
    output_dir.mkdir(parents=True, exist_ok=True)
    log("导出完整核对明细 CSV")
    detail_exports = {}
    export_specs = {
        "wdt_huice_detail": "SELECT * FROM wdt_huice_recon ORDER BY CASE result WHEN '单号金额一致' THEN 2 ELSE 1 END, result, platform_order_no",
        "huice_summary_detail": "SELECT * FROM huice_summary_recon ORDER BY result DESC,reconcile_date,platform,shop",
        "oms_sap_detail": "SELECT * FROM oms_sap_recon ORDER BY CASE result WHEN '数量金额一致' THEN 2 ELSE 1 END,result,oms_sales_no,material_code",
        "oms_huice_detail": "SELECT * FROM oms_huice_month_shop ORDER BY CASE result WHEN '金额一致' THEN 2 ELSE 1 END,result,outbound_month,mapped_shop",
        "customer_shop_mapping": "SELECT * FROM customer_shop_map ORDER BY mapping_status,customer_code",
    }
    for name, sql in export_specs.items():
        path = output_dir / f"{name}.csv"
        count = export_query_csv(connection, sql, path)
        json_limit = 50000 if name == "wdt_huice_detail" else 200000
        json_path = output_dir / f"{name}_workbook.json"
        workbook_export = export_query_json_limited(connection, sql, json_path, json_limit)
        detail_exports[name] = {
            "path": str(path),
            "rows": count,
            "workbook_json": workbook_export["path"],
            "workbook_rows": workbook_export["included_rows"],
            "workbook_limit": workbook_export["limit"],
        }
        log(f"导出 {name}: {count:,}行")

    wdt_huice_status = rows_as_dicts(
        connection,
        """SELECT result,COUNT(*) order_count,SUM(wdt_amount) wdt_amount,
                  SUM(huice_net_receivable) huice_net_receivable,SUM(amount_difference) amount_difference
           FROM wdt_huice_recon GROUP BY result ORDER BY result""",
    )
    huice_summary_status = rows_as_dicts(
        connection,
        "SELECT result,COUNT(*) group_count FROM huice_summary_recon GROUP BY result ORDER BY result",
    )
    oms_sap_status = rows_as_dicts(
        connection,
        """SELECT result,COUNT(*) key_count,SUM(invoice_qty) sap_qty,SUM(oms_qty) oms_qty,
                  SUM(sap_amount) sap_amount,SUM(oms_amount) oms_amount,
                  SUM(quantity_difference) quantity_difference,SUM(amount_difference) amount_difference
           FROM oms_sap_recon GROUP BY result ORDER BY result""",
    )
    oms_huice_status = rows_as_dicts(
        connection,
        """SELECT result,COUNT(*) group_count,SUM(oms_amount) oms_amount,
                  SUM(huice_net_receivable) huice_net_receivable,SUM(amount_difference) amount_difference
           FROM oms_huice_month_shop GROUP BY result ORDER BY result""",
    )

    wdt_total = query_one(connection, "SELECT COUNT(*),SUM(wdt_amount) FROM wdt_platform")
    huice_total = query_one(connection, "SELECT COUNT(*),SUM(huice_net_receivable),SUM(huice_net_cash) FROM huice_platform")
    wdt_matched = query_one(
        connection,
        "SELECT COUNT(*),SUM(wdt_amount),SUM(huice_net_receivable) FROM wdt_huice_recon WHERE result IN ('单号金额一致','单号一致金额差异')",
    )
    hc_summary_total = query_one(connection, "SELECT COUNT(*),SUM(CASE WHEN result='一致' THEN 1 ELSE 0 END) FROM huice_summary_recon")
    sap_total = query_one(connection, "SELECT SUM(invoice_qty),SUM(tax_amount),SUM(row_count) FROM sap2c")
    oms_total = query_one(connection, "SELECT SUM(item_num),SUM(share_amount),COUNT(*) FROM oms_detail")
    oms_sap_matched = query_one(
        connection,
        "SELECT SUM(invoice_qty),SUM(oms_qty),SUM(sap_amount),SUM(oms_amount),SUM(CASE WHEN result='数量金额一致' THEN 1 ELSE 0 END),COUNT(*) FROM oms_sap_recon",
    )
    oms_pool = query_one(
        connection,
        """SELECT SUM(item_num),SUM(share_amount),COUNT(*) FROM oms_detail o
           WHERE NOT EXISTS (SELECT 1 FROM sap2c_key s WHERE s.oms_sales_no=o.document_no)""",
    )
    mapping = query_one(
        connection,
        "SELECT COUNT(*),SUM(CASE WHEN mapping_status='自动高置信' THEN 1 ELSE 0 END),SUM(CASE WHEN mapping_status='未映射' THEN 1 ELSE 0 END) FROM customer_shop_map",
    )

    correction_log = rows_as_dicts(
        connection,
        "SELECT * FROM sap_oms_sales_no_correction_log ORDER BY original_oms_sales_no",
    )

    summary = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "parameters": {
            "wdt_huice_amount_tolerance": 0.01,
            "huice_summary_amount_tolerance": 0.01,
            "oms_sap_line_amount_tolerance": 0.01,
            "oms_sap_total_amount_tolerance": 2.0,
            "date_scope": "2025-12至2026-06（SAP/惠策主体为2026-01至2026-06）",
            "sap_filter": "发票类型描述=标准发票（2C)",
            "oms_sap_key": "OMS.document_no=SAP.OMS销售单号；再加物料编码+销售单位",
            "wdt_huice_key": "旺店通.原始单号=惠策.平台订单号",
        },
        "controls": {
            "wdt_platform_orders": wdt_total[0] or 0,
            "wdt_amount": wdt_total[1] or 0,
            "huice_platform_orders": huice_total[0] or 0,
            "huice_net_receivable": huice_total[1] or 0,
            "huice_net_cash": huice_total[2] or 0,
            "wdt_huice_matched_orders": wdt_matched[0] or 0,
            "wdt_huice_matched_wdt_amount": wdt_matched[1] or 0,
            "wdt_huice_matched_huice_amount": wdt_matched[2] or 0,
            "huice_summary_groups": hc_summary_total[0] or 0,
            "huice_summary_consistent_groups": hc_summary_total[1] or 0,
            "sap2c_quantity": sap_total[0] or 0,
            "sap2c_amount": sap_total[1] or 0,
            "sap2c_rows": sap_total[2] or 0,
            "oms_total_quantity": oms_total[0] or 0,
            "oms_total_amount": oms_total[1] or 0,
            "oms_total_rows": oms_total[2] or 0,
            "oms_sap_sap_quantity": oms_sap_matched[0] or 0,
            "oms_sap_oms_quantity": oms_sap_matched[1] or 0,
            "oms_sap_sap_amount": oms_sap_matched[2] or 0,
            "oms_sap_oms_amount": oms_sap_matched[3] or 0,
            "oms_sap_exact_keys": oms_sap_matched[4] or 0,
            "oms_sap_total_keys": oms_sap_matched[5] or 0,
            "oms_huice_pool_quantity": oms_pool[0] or 0,
            "oms_huice_pool_amount": oms_pool[1] or 0,
            "oms_huice_pool_rows": oms_pool[2] or 0,
            "mapping_customers": mapping[0] or 0,
            "mapping_high_confidence": mapping[1] or 0,
            "mapping_unmapped": mapping[2] or 0,
        },
        "wdt_huice_summary": wdt_huice_status,
        "huice_summary_recon": huice_summary_status,
        "oms_sap_summary": oms_sap_status,
        "oms_huice_summary": oms_huice_status,
        "sap_oms_sales_no_corrections": correction_log,
        "detail_exports": detail_exports,
    }
    (output_dir / "summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    return summary


def main() -> None:
    parser = argparse.ArgumentParser(description="销售 ToC 全链路核对")
    parser.add_argument("--input", type=Path, default=Path("input"))
    parser.add_argument("--work", type=Path, default=Path("reconciliation/work"))
    parser.add_argument("--output", type=Path, default=Path("reconciliation/output"))
    parser.add_argument("--max-rows-per-file", type=int, default=0, help="测试模式：每个xlsx最多读取的数据行数")
    parser.add_argument("--workers", type=int, default=min(4, os.cpu_count() or 1), help="大型分卷并行抽取进程数")
    parser.add_argument("--rebuild", action="store_true", help="删除已有数据库并全量重跑")
    args = parser.parse_args()
    args.work.mkdir(parents=True, exist_ok=True)
    args.output.mkdir(parents=True, exist_ok=True)
    database_path = args.work / "reconciliation.db"
    if args.rebuild and database_path.exists():
        database_path.unlink()
    connection = sqlite3.connect(database_path)
    connection.create_function(
        "correct_sap_oms_sales_no",
        1,
        correct_sap_oms_sales_no,
        deterministic=True,
    )
    configure_database(connection)
    create_schema(connection)
    ensure_wdt_order_composite_key(connection)
    ensure_huice_summary_composite_key(connection)
    max_rows = args.max_rows_per_file or None

    extract_wdt(connection, args.input, args.work, max_rows, args.workers)
    extract_huice_detail(connection, args.input, args.work, max_rows, args.workers)
    extract_huice_summary(connection, args.input, max_rows)
    extract_sap2c(connection, args.input, max_rows)
    extract_oms(connection, args.input, max_rows)
    refresh_oms_transaction_classification(connection)
    build_customer_mapping(connection)
    materialize_analysis_tables(connection)
    summary = build_summary(connection, args.output)
    log(f"完成。summary.json={args.output / 'summary.json'}")
    log(json.dumps(summary["controls"], ensure_ascii=False))
    connection.close()


if __name__ == "__main__":
    main()
