#!/usr/bin/env python3
"""从已抽取缓存按“内部订单号+平台原始单号”粒度重载旺店通订单。"""

from pathlib import Path
import sqlite3

from reconcile_sales_toc import (
    configure_database,
    create_schema,
    ensure_wdt_order_composite_key,
    extract_wdt,
)


def main() -> None:
    database = Path("reconciliation/work/reconciliation.db")
    connection = sqlite3.connect(database)
    configure_database(connection)
    create_schema(connection)
    ensure_wdt_order_composite_key(connection)
    connection.execute("DELETE FROM stage_meta WHERE stage='wdt'")
    connection.commit()
    extract_wdt(
        connection,
        Path("input"),
        Path("reconciliation/work"),
        None,
        4,
    )
    connection.executescript(
        """
        DROP TABLE IF EXISTS wdt_order_dedup;
        CREATE TABLE wdt_order_dedup AS
        SELECT order_no,platform_order_no,MAX(sub_order_no) sub_order_no,MAX(shop) shop,
          MAX(order_status) order_status,MAX(order_type) order_type,
          MIN(NULLIF(trade_time,'')) trade_time,MIN(NULLIF(payment_time,'')) payment_time,
          MAX(NULLIF(ship_time,'')) ship_time,MAX(receivable_amount) receivable_amount,
          MAX(allocated_total) allocated_total,MAX(quantity) quantity,MAX(line_count) line_count,
          COUNT(*) file_occurrences,GROUP_CONCAT(source_file,'|') source_files
        FROM wdt_orders_file GROUP BY order_no,platform_order_no;
        CREATE UNIQUE INDEX idx_wdt_order_pair ON wdt_order_dedup(order_no,platform_order_no);
        CREATE INDEX idx_wdt_order_platform ON wdt_order_dedup(platform_order_no);

        DROP TABLE IF EXISTS wdt_order_item;
        CREATE TABLE wdt_order_item AS
        WITH chosen_source AS (
          SELECT source_file,order_no,platform_order_no,
            ROW_NUMBER() OVER (
              PARTITION BY order_no,platform_order_no
              ORDER BY quantity DESC,line_count DESC,ship_time DESC,source_file DESC
            ) source_rank
          FROM wdt_orders_file
        )
        SELECT s.order_no,s.platform_order_no,
          CASE WHEN s.material_key LIKE 'SKU:%' THEN COALESCE(m.sap_code,s.material_key) ELSE s.material_key END material_code,
          s.unit,SUM(s.quantity) quantity,SUM(s.amount) amount,SUM(s.line_count) line_count
        FROM wdt_order_item_stage s
        JOIN chosen_source c ON c.source_file=s.source_file AND c.order_no=s.order_no
          AND c.platform_order_no=s.platform_order_no AND c.source_rank=1
        LEFT JOIN wdt_product_sap_best m ON substr(s.material_key,5)=m.product_code
        GROUP BY 1,2,3,4;
        CREATE INDEX idx_wdt_order_item_order ON wdt_order_item(order_no,platform_order_no);
        CREATE INDEX idx_wdt_order_item_platform ON wdt_order_item(platform_order_no);
        """
    )
    connection.commit()
    print("loaded", connection.execute("SELECT COUNT(*) FROM wdt_orders_file").fetchone()[0])
    connection.close()


if __name__ == "__main__":
    main()
