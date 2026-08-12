#!/usr/bin/env python3
"""生成SAP标准发票（2C）与OMS月结Y001标准结算的金额匹配明细JSON。"""

from __future__ import annotations

import json
import sqlite3
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path

from reconcile_sales_toc import as_number, iter_selected_rows, text
from oms_transaction_codes import OMS_STANDARD_SETTLEMENT_CODES, sql_list

ROOT = Path(__file__).resolve().parent.parent
DB = ROOT / "reconciliation" / "work" / "reconciliation.db"
SAP_DIR = ROOT / "input" / "发票清单：26.01.01-26.06.30"
OUTPUT = ROOT / "reconciliation" / "results" / "oms_sap_amount_match_detail.json"
OMS_STANDARD_SETTLEMENT_SQL = sql_list(OMS_STANDARD_SETTLEMENT_CODES)


def iso_date(value: object) -> str:
    raw = text(value)
    if not raw:
        return ""
    try:
        serial = float(raw)
        if 30000 <= serial <= 60000:
            return (datetime(1899, 12, 30) + timedelta(days=serial)).date().isoformat()
    except ValueError:
        pass
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y%m%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(raw, fmt).date().isoformat()
        except ValueError:
            continue
    return raw


def collect_sap_attributes() -> dict[tuple[str, str, str], dict[str, set[str]]]:
    attributes: dict[tuple[str, str, str], dict[str, set[str]]] = defaultdict(
        lambda: {"invoices": set(), "creation_dates": set(), "posting_dates": set(), "customer_codes": set(), "customer_names": set()}
    )
    selected = [1, 6, 15, 16, 17, 26, 27, 32, 44]
    for path in sorted(SAP_DIR.glob("*.XLSX")):
        for _, values in iter_selected_rows(path, selected):
            if text(values.get(15)) != "标准发票（2C)":
                continue
            key = (text(values.get(1)), text(values.get(32)), text(values.get(44)))
            item = attributes[key]
            pairs = (
                ("invoices", text(values.get(6))),
                ("creation_dates", iso_date(values.get(26))),
                ("posting_dates", iso_date(values.get(27))),
                ("customer_codes", text(values.get(16))),
                ("customer_names", text(values.get(17))),
            )
            for field, value in pairs:
                if value:
                    item[field].add(value)
    return attributes


def joined(values: set[str]) -> str:
    return " | ".join(sorted(values))


def main() -> None:
    attributes = collect_sap_attributes()
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    rows = []
    categories = {name: {"keys": 0, "sap_amount": 0.0, "oms_amount": 0.0} for name in (
        "完全匹配", "金额差异", "SAP存在但OMS不存在", "OMS存在但SAP不存在"
    )}
    query = """
      SELECT oms_sales_no,material_code,sales_unit,file_month,outbound_month,sap_invoice_nos,
        customer_code,customer_name,sap_qty,oms_qty,sap_amount,oms_amount,amount_difference,
        mapping_result,source_result
      FROM v4_oms_sap_field_map
      ORDER BY CASE mapping_result WHEN '双向字段一致' THEN 1 WHEN '数量一致金额差异' THEN 2
        WHEN '数量金额差异' THEN 2 WHEN 'SAP补数量金额' THEN 3 ELSE 4 END,
        outbound_month,oms_sales_no,material_code,sales_unit
    """
    for record in conn.execute(query):
        if record["mapping_result"] == "双向字段一致":
            category = "完全匹配"
        elif record["mapping_result"] in ("数量一致金额差异", "数量金额差异"):
            category = "金额差异"
        elif record["mapping_result"] == "SAP补数量金额":
            category = "SAP存在但OMS不存在"
        else:
            category = "OMS存在但SAP不存在"
        key = (record["oms_sales_no"] or "", record["material_code"] or "", record["sales_unit"] or "")
        attr = attributes.get(key, {})
        sap_amount = as_number(record["sap_amount"])
        oms_amount = as_number(record["oms_amount"])
        categories[category]["keys"] += 1
        categories[category]["sap_amount"] += sap_amount
        categories[category]["oms_amount"] += oms_amount
        rows.append({
            "match_category": category,
            "sap_invoice_no": joined(attr.get("invoices", set())) or (record["sap_invoice_nos"] or ""),
            "sap_invoice_date": joined(attr.get("creation_dates", set())),
            "sap_posting_date": joined(attr.get("posting_dates", set())),
            "sap_customer_code": joined(attr.get("customer_codes", set())),
            "sap_customer_name": joined(attr.get("customer_names", set())),
            "sap_invoice_amount": sap_amount,
            "oms_monthly_document_no": record["oms_sales_no"] or "",
            "oms_monthly_amount": oms_amount,
            "material_code": record["material_code"] or "",
            "sales_unit": record["sales_unit"] or "",
            "match_fields": "OMS销售单号+物料编码+销售单位",
            "amount_difference": sap_amount - oms_amount,
            "sap_quantity": as_number(record["sap_qty"]),
            "oms_quantity": as_number(record["oms_qty"]),
            "source_mapping_result": record["mapping_result"] or "",
        })
    controls = dict(conn.execute(f"""
      SELECT
        (SELECT COALESCE(SUM(tax_amount),0) FROM sap2c WHERE file_month BETWEEN '2026-01' AND '2026-06') sap_execution_total,
        (SELECT COALESCE(SUM(share_amount),0) FROM oms_detail WHERE business_type IN ({OMS_STANDARD_SETTLEMENT_SQL}) AND outbound_time>='2026-01-01' AND outbound_time<'2026-07-01') oms_execution_total
    """).fetchone())
    conn.close()
    exact = categories["完全匹配"]
    difference = categories["金额差异"]
    summary = {
        **controls,
        "sap_participating_amount": exact["sap_amount"] + difference["sap_amount"],
        "oms_participating_amount": exact["oms_amount"] + difference["oms_amount"],
        "sap_actual_match_amount": exact["sap_amount"],
        "oms_actual_match_amount": exact["oms_amount"],
    }
    summary["actual_match_difference"] = summary["sap_actual_match_amount"] - summary["oms_actual_match_amount"]
    summary["amount_match_rate"] = min(abs(summary["sap_actual_match_amount"]), abs(summary["oms_actual_match_amount"])) / max(abs(summary["sap_actual_match_amount"]), abs(summary["oms_actual_match_amount"]), 1)
    summary["sap_coverage_rate"] = abs(summary["sap_actual_match_amount"]) / max(abs(summary["sap_execution_total"]), 1)
    summary["oms_coverage_rate"] = abs(summary["oms_actual_match_amount"]) / max(abs(summary["oms_execution_total"]), 1)
    result = {
        "period": "2026-01-01至2026-06-30",
        "definitions": {
            "sap_invoice_date": "SAP开票清单字段：发票创建日期",
            "sap_invoice_amount": "SAP开票清单字段：含税金额；仅标准发票（2C）",
            "oms_monthly_amount": "OMS日结月结查询记录字段：share_amount；月结事务码中的Y001标准结算子集",
            "participating_amount": "OMS销售单号+物料编码+销售单位在双方均存在的共同键金额，包括完全匹配和金额差异",
            "actual_match_amount": "共同键且SAP含税金额与OMS share_amount差异不超过0.01元的金额",
        },
        "summary": summary,
        "categories": categories,
        "rows": rows,
    }
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT.write_text(json.dumps(result, ensure_ascii=False), encoding="utf-8")
    print(json.dumps({"output": str(OUTPUT), "rows": len(rows), "summary": summary, "categories": categories}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
