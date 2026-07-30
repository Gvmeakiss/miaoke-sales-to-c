# 销售 ToC 核对项目

本项目用于核对 2026-01-01 至 2026-06-30 的销售 ToC 数据链：

`旺店通订单 → 惠策对账明细 → 惠策店铺汇总 → OMS 月结 Y001 → SAP 标准发票（2C）`

当前目录仅保留正式版本、原始输入、可复现脚本、当前数据库及完整核对结果。历史 V2/V3/V4/V4.1、修复版、优化测试版、样本数据库、旧检查日志和 PKL 缓存已清理。

## 1. 项目结构

```text
miaoke sales to c/
├── input/                         原始资料，只读保留
├── reconciliation/
│   ├── reconcile_sales_toc.py    原始资料入库及基础标准化
│   ├── explore_wdt_oms_matching.py 生成旺店通店铺—OMS客户候选映射
│   ├── reconcile_huice_oms.py    生成惠策—OMS桥接映射及中间核对
│   ├── reload_wdt_orders.py      按内部订单号+平台订单号重载旺店通粒度
│   ├── reconcile_sales_flow.py   生成正式核对结果
│   ├── build_sales_workpaper.mjs 生成正式Excel底稿
│   ├── work/                     当前SQLite数据库及抽取缓存，可重建
│   ├── results/                  完整CSV及工作簿JSON结果，可重建
│   ├── intermediate/             全量重跑时生成的中间输出，可删除
│   └── qa_previews/              生成底稿时的页面检查图，可删除
└── outputs/
    └── sales_toc_workpaper_final_20260101_20260630/
        ├── 销售ToC业务流程核对底稿_20260101-20260630.xlsx
        ├── 销售ToC核对财务报告说明_20260101-20260630.txt
        └── 客户待补充材料清单_销售ToC_20260101-20260630.txt
```

## 2. 数据范围与来源

- 核对期间：2026-01-01 至 2026-06-30，含首尾。
- 旺店通：`input/旺店通订单清单/`，订单表头及商品明细。
- 惠策明细：`input/惠策系统对账单清单/`，共 15 份订单级对账明细。
- 惠策汇总：`input/惠策系统对账单汇总/`，共 18 份店铺级月度汇总。
- OMS：`input/OMS 系统日结月结查询记录：25年12月到26年6月2C单据.sql`。
- SAP：`input/发票清单：26.01.01-26.06.30/`。

原始资料不随核对过程修改。当前惠策明细和汇总导出均没有物料、SKU或商品数量字段。

## 3. 核对逻辑

### 3.1 旺店通订单—惠策明细

- 强键：`旺店通原始单号 = 惠策平台订单号`。
- 映射粒度：一行一个平台订单号，同时保留旺店通内部订单号集合、惠策对账流水号集合、店铺及月份。
- 金额口径：旺店通平台订单分摊金额分别与惠策本期应收、本期实收比较；订单应收仅作为辅助证据。
- 完整平台订单并集为 5,805,366 行，保存在 `reconciliation/results/order_bill_recon.csv`。
- Excel 第 4 页分层展示 15,000 行：一致 6,000 行、金额差异 3,000 行、双方单边各 3,000 行。

### 3.2 惠策明细—惠策店铺汇总

- 共同维度：导出结算月份 + 平台 + 店铺。
- 主口径：全量应收和全量实收分别核对。
- 汇总“对账成功金额”与明细“对账状态=对账成功”定义不同，不作为内部主钩稽口径。

### 3.3 惠策店铺汇总—OMS 月结

- 惠策没有物料字段，按月份 + 惠策店铺/OMS客户映射核对金额。
- 惠策对账成功分类金额为主，应收、实收作为并列辅助口径。
- 同一客户跨月累计仅解释结算时点，不替代逐月匹配。
- 不执行“OMS未开票池—惠策”核对。

### 3.4 数量链

- 惠策平台订单数是惠策原生订单数量。
- “惠策覆盖订单派生数量”由平台订单号连接旺店通商品明细取得，不代表惠策原生商品数量。
- 派生订单数量与 OMS 按发货月份 + 店铺/客户 + SAP物料进行多对一核对。
- OMS 与 SAP 按销售单号 + 物料编码 + 销售单位核对原生数量。

### 3.5 OMS 月结—SAP

- 强键：OMS销售单号 + 物料编码 + 销售单位。
- 同时比较数量和含税金额。
- 一致键用于双向补充客户、月份、发票号等字段；双方单边记录完整保留。

## 4. 关键结果

- 旺店通平台订单 4,701,512；惠策平台订单 5,683,439；共同订单 4,579,585。
- 旺店通订单号覆盖率 97.41%；惠策订单号覆盖率 80.58%；金额一致订单率 93.02%。
- 订单分摊金额精确匹配 251,472,703.57 元，分摊金额覆盖率 94.25%。
- 惠策明细—汇总应收匹配率 99.98%，实收匹配率 99.97%；差异集中于 2026 年 6 月单一店铺组合。
- 惠策成功分类金额 283,744,335.83 元；OMS月结金额 280,247,849.59 元；总体金额匹配率 98.75%。
- 旺店通商品数量 17,858,650；惠策覆盖订单派生数量 17,306,938；OMS月结数量 17,870,519。
- OMS—SAP一致键双方数量 17,868,882；OMS数量覆盖率 99.99%。

## 5. 正式交付物

- 主底稿：`outputs/sales_toc_workpaper_final_20260101_20260630/销售ToC业务流程核对底稿_20260101-20260630.xlsx`
- 财务报告文字说明：同目录 `销售ToC核对财务报告说明_20260101-20260630.txt`
- 资料限制说明：同目录 `客户待补充材料清单_销售ToC_20260101-20260630.txt`
- 完整明细：`reconciliation/results/`

`results/` 中主要文件：

- `summary.json`：全局控制数及各环节汇总。
- `order_bill_recon.csv`：旺店通—惠策完整平台订单映射。
- `huice_internal_recon.csv`：惠策明细—店铺汇总核对。
- `bill_oms_month_recon.csv`：惠策店铺汇总—OMS月结核对。
- `order_bill_oms_qty_recon.csv`：订单账单证据—OMS数量核对。
- `oms_sap_field_map.csv`：OMS—SAP双向字段映射。
- `huice_shop_map.csv`：惠策店铺—OMS客户映射。

## 6. 重跑方式

### 6.1 使用现有数据库刷新正式结果

```bash
python3 reconciliation/reload_wdt_orders.py
python3 reconciliation/reconcile_sales_flow.py --output-dir reconciliation/results
```

生成 Excel 时需使用 Codex 工作区提供的 Node.js 与 `@oai/artifact-tool` 依赖。构建脚本会在临时目录建立依赖链接，完成后自动清理：

```bash
WORKSPACE_NODE='<workspace-node>' \
WORKSPACE_NODE_MODULES='<workspace-node_modules>' \
  reconciliation/build_workpaper.sh
```

### 6.2 从原始资料全量重建

```bash
python3 reconciliation/reconcile_sales_toc.py \
  --input input --work reconciliation/work \
  --output reconciliation/intermediate/base --rebuild

python3 reconciliation/explore_wdt_oms_matching.py
python3 reconciliation/reconcile_huice_oms.py
python3 reconciliation/reload_wdt_orders.py
python3 reconciliation/reconcile_sales_flow.py --output-dir reconciliation/results
```

全量重跑体量较大。完成后应检查：

1. `reconciliation/results/summary.json` 控制数是否与本README关键结果一致。
2. Excel公式错误扫描是否为0。
3. 15个工作表页面预览是否完整、无截断。
4. XLSX压缩结构检查是否通过。

## 7. 维护原则

- `input/`、正式底稿及当前数据库不得作为临时目录使用。
- 新测试统一写入 `reconciliation/intermediate/`，验证后可整体删除。
- 不再新增 V2/V3/V4 等并行版本目录；正式结果直接覆盖 `reconciliation/results/` 和正式底稿。
- 需保留历史版本时，应在项目外归档，不放回当前工作目录。
