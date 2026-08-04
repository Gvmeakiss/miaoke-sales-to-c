#!/usr/bin/env python3
"""将正式范围匹配明细按每个CSV不超过100万条拆分，并生成审计校验清单。"""

from __future__ import annotations

import csv
import hashlib
import json
from decimal import Decimal
from pathlib import Path


ROOT = Path(__file__).resolve().parent.parent
SOURCE = ROOT / "reconciliation" / "results" / "formal_order_match_detail"
OUTPUT = ROOT / "outputs" / "sales_toc_workpaper_final_20260101_20260630" / "2026年度正式范围旺店通订单匹配明细" / "明细数据"
ROWS_PER_FILE = 1_000_000

CATEGORIES = [
    ("可匹配条目", SOURCE / "01_可匹配条目.csv", 4_204_532, Decimal("243122621.11"), Decimal("243122622.15")),
    ("仅账单未匹配", SOURCE / "02_仅账单及账单未成功匹配.csv", 1_564_973, Decimal("32210995.93"), Decimal("23114287.69")),
    ("仅订单未匹配", SOURCE / "03_仅订单未匹配.csv", 127_158, Decimal("0.00"), Decimal("2367485.25")),
]


def checksum(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def split_category(label: str, source: Path, expected_rows: int, expected_huice: Decimal, expected_wdt: Decimal) -> dict:
    category_dir = OUTPUT / label
    category_dir.mkdir(parents=True, exist_ok=True)
    files = []
    total_rows = 0
    total_huice = Decimal("0")
    total_wdt = Decimal("0")
    with source.open("r", encoding="utf-8", newline="") as input_stream:
        reader = csv.reader(input_stream)
        headers = next(reader)
        part = 0
        writer = None
        output_stream = None
        part_rows = 0
        part_huice = Decimal("0")
        part_wdt = Decimal("0")
        part_start = 1

        def close_part() -> None:
            nonlocal output_stream, writer, part_rows, part_huice, part_wdt
            if output_stream is None:
                return
            output_stream.close()
            part_end = part_start + part_rows - 1
            final_name = f"{label}_第{part:02d}部分_{part_start:07d}-{part_end:07d}.csv"
            temp_path = category_dir / f".{label}_第{part:02d}部分.tmp.csv"
            final_path = category_dir / final_name
            temp_path.replace(final_path)
            files.append({
                "file": str(final_path), "rows": part_rows,
                "start": part_start, "end": part_end,
                "huice_amount": float(part_huice), "wdt_amount": float(part_wdt),
                "bytes": final_path.stat().st_size, "sha256": checksum(final_path),
            })
            output_stream = None
            writer = None

        for row in reader:
            if writer is None or part_rows >= ROWS_PER_FILE:
                close_part()
                part += 1
                part_start = total_rows + 1
                part_rows = 0
                part_huice = Decimal("0")
                part_wdt = Decimal("0")
                temp_path = category_dir / f".{label}_第{part:02d}部分.tmp.csv"
                output_stream = temp_path.open("w", encoding="utf-8-sig", newline="")
                writer = csv.writer(output_stream)
                writer.writerow(headers)
            writer.writerow(row)
            huice = Decimal(row[10] or "0")
            wdt = Decimal(row[11] or "0")
            part_rows += 1
            total_rows += 1
            part_huice += huice
            part_wdt += wdt
            total_huice += huice
            total_wdt += wdt
        close_part()

    verification = {
        "rows_match": total_rows == expected_rows,
        "huice_amount_match": total_huice.quantize(Decimal("0.01")) == expected_huice,
        "wdt_amount_match": total_wdt.quantize(Decimal("0.01")) == expected_wdt,
        "all_files_within_limit": all(item["rows"] <= ROWS_PER_FILE for item in files),
    }
    if not all(verification.values()):
        raise RuntimeError(f"{label}拆分校验失败：{verification}")
    result = {
        "category": label, "rows": total_rows,
        "huice_amount": float(total_huice.quantize(Decimal("0.01"))),
        "wdt_amount": float(total_wdt.quantize(Decimal("0.01"))),
        "files": files, "verification": verification,
    }
    print(json.dumps({"category": label, "rows": total_rows, "parts": len(files)}, ensure_ascii=False), flush=True)
    return result


def main() -> None:
    OUTPUT.mkdir(parents=True, exist_ok=True)
    categories = [split_category(*item) for item in CATEGORIES]
    manifest = {
        "scope": "2026-01-01至2026-06-30",
        "rows_per_file": ROWS_PER_FILE,
        "categories": categories,
        "totals": {
            "union_rows": sum(item["rows"] for item in categories),
            "huice_orders": 5_769_505,
            "wdt_orders": 4_728_414,
            "successful_orders": 4_204_532,
            "successful_huice_amount": 243_122_621.11,
            "successful_wdt_amount": 243_122_622.15,
        },
        "verification": {
            "union_rows_match": sum(item["rows"] for item in categories) == 5_896_663,
            "all_category_checks_pass": all(all(item["verification"].values()) for item in categories),
        },
    }
    (OUTPUT.parent / "明细拆分索引及校验结果.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8"
    )


if __name__ == "__main__":
    main()
