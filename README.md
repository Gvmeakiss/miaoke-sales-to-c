# 销售 ToC 数据核对（Miaoke） 🧮

> 核对 Miaoke 销售 ToC 业务链中旺店通订单、惠策平台账单、OMS 月结与 SAP 开票数据，输出内部全链路核对底稿（Python 数据处理 + Node.js 底稿构建）。

[![Language](https://img.shields.io/badge/language-Python%20%7C%20Node.js-blue)](https://github.com/Gvmeakiss/miaoke-sales-to-c) [![License](https://img.shields.io/badge/license-MIT-green)](https://github.com/Gvmeakiss/miaoke-sales-to-c/blob/main/LICENSE) [![Domain](https://img.shields.io/badge/domain-Audit%20Analytics-orange)](https://github.com/Gvmeakiss/miaoke-sales-to-c)

## 📌 项目简介

本仓库用于核对 Miaoke 销售 ToC 业务链条的数据一致性：从旺店通订单，到惠策平台账单，到 OMS 月结（业务类型 `Y001`），再到 SAP 标准发票（2C）。业务流程按 `旺店通订单 → 惠策账单 → OMS 月结 Y001 → SAP 标准发票` 组织，但核对采用相邻系统单对单（pairwise）方式执行。技术栈为 **Python**（数据抽取、钩稽、SQLite 中间库、CSV/JSON 导出）加 **Node.js**（基于 `@oai/artifact-tool` 构建带格式的多 Sheet Excel 底稿）。全部为内部核对与审阅辅助，不自动下审计/会计结论。

## ✨ 功能特性

- **四段 pairwise 核对**：
  - 旺店通订单 → 惠策明细：主键 `旺店通原始单号 = 惠策平台订单号`，金额比较应收/实收；
  - 惠策店铺账单 → OMS 月结：主维度 `月份 + 店铺/客户映射`，主口径惠策对账成功金额 vs OMS 月结金额；
  - 订单—账单证据 → OMS 月结数量：`数量 A`（惠策出现的旺店通商品数量）vs `数量 B`（OMS 月结 Y001 原生数量）；
  - OMS 月结 → SAP 发票：主键 `OMS销售单号 + 物料编码 + 销售单位`，比较数量与含税金额并双向补充字段。
- **流式读取大 Excel**：`reconcile_sales_toc.py` 用 `lxml` 以 XML 流式读取 `.xlsx`，避免一次性载入内存；旺店通商品行先压缩到内部订单再汇总至平台原始单号。
- **SQLite 中间库**：`reconcile_sales_flow_v4.py` / `reconcile_sales_toc.py` 将抽取结果写入 `work_full/reconciliation.db`，再导出分项明细 CSV 与 `summary.json`。
- **惠策↔OMS 核对**：`reconcile_huice_oms.py` 的 `extract_order_items` / `load_order_items` / `extract_huice_current_file` / `load_huice_current` / `build_reconciliation` / `export_outputs` 完成订单项与惠策账单抽取与核对。
- **店铺映射探索**：`explore_wdt_oms_matching.py` 的 `load_wdt` / `build_candidates` / `normalize_shop` / `platform` / `cosine` 用余弦相似度探索旺店通与 OMS 店铺映射。
- **订单财务分析**：`订单数据分析_财务版.py` 的 `load_and_standardize` / `_periodize` / `_aggregate_finance` / `build_outputs` / `write_excel` 按期间与财务口径聚合订单数据。
- **多版底稿构建（Node.js）**：`build_sales_flow_v4.mjs`（13 张表「业务流程核对底稿」）、`build_dataflow_workbook.mjs` / `build_dataflow_workbook_v3.mjs`（数据流核对底稿 V2/V3）、`build_reconciliation_workbook.mjs`（12 张表「全链路核对底稿」，读取 `summary.json` 与各 `*_workbook.json`）；均扫描 `#REF!`/`#DIV/0!`/`#VALUE!`/`#NAME?`/`#N/A` 并生成 `_qa_previews/`。
- **明确阈值**：代码中金额一致阈值为绝对差异 `≤ 0.01` 元（`reconcile_sales_flow_v4.py` 多处 `ABS(...) <= 0.01`），数量一致阈值为差异 `≤ 0.000001`。

## 📂 目录结构

```
miaoke-sales-to-c/
├── README.md
├── requirements.txt              # pandas>=2.0 / numpy>=1.24 / openpyxl>=3.1 / lxml>=4.9
├── 订单数据分析_财务版.py          # 订单财务口径聚合与 Excel 导出（load_and_standardize/_periodize/build_outputs）
├── reconciliation/
│   ├── reconcile_sales_toc.py    # ToC 全链路核对：流式读 xlsx → SQLite → CSV + summary.json
│   ├── reconcile_huice_oms.py    # 惠策↔OMS 月结核对
│   ├── explore_wdt_oms_matching.py # 旺店通↔OMS 店铺映射探索（余弦相似度）
│   ├── reconcile_sales_flow_v4.py # 2026H1 主流程（订单→账单→OMS月结→SAP），写 work_full/reconciliation.db、output_flow_v4/
│   ├── build_sales_flow_v4.mjs    # 生成 13 张表「业务流程核对底稿」Excel
│   ├── build_dataflow_workbook.mjs / build_dataflow_workbook_v3.mjs # 数据流核对底稿 V2/V3
│   └── build_reconciliation_workbook.mjs # 生成 12 张表「全链路核对底稿」Excel
└── LICENSE
```

## 🔧 环境要求

- Python 3.8+（代码使用 `from __future__ import annotations`、`pathlib`、类型注解、`lxml` 流式解析）
- Node.js 18+（Excel 构建脚本依赖 `@oai/artifact-tool`，普通 Node 环境可能需 Codex 工作区提供的依赖运行环境）
- Python 依赖见 `requirements.txt`：`pandas>=2.0`、`numpy>=1.24`、`openpyxl>=3.1`、`lxml>=4.9`

## 🚀 安装

```bash
git clone https://github.com/Gvmeakiss/miaoke-sales-to-c.git
cd miaoke-sales-to-c
pip install -r requirements.txt
# Node 侧需能解析 @oai/artifact-tool（见下方重跑说明）
```

## 💡 快速开始 / 使用示例

主流程（2026 年 1–6 月）：

```bash
# 1) Python：抽取、钩稽、写 SQLite 与 CSV/JSON
python3 reconciliation/reconcile_sales_flow_v4.py

# 2) Node：根据 output_flow_v4 产物生成 Excel 底稿
node reconciliation/build_sales_flow_v4.mjs
# 可用环境变量覆盖输出位置：
#   SALES_TOC_OUTPUT_DIR / SALES_TOC_OUTPUT_FILE
```

其它可单独运行的脚本：

```bash
python3 reconciliation/reconcile_sales_toc.py        # ToC 全链路核对（流式读 xlsx）
python3 reconciliation/reconcile_huice_oms.py        # 惠策↔OMS 月结核对
python3 reconciliation/explore_wdt_oms_matching.py   # 旺店通↔OMS 店铺映射探索
python3 订单数据分析_财务版.py                        # 订单财务口径分析
```

重跑后校验 Excel 完整性（示例文件名以实际输出为准）：

```bash
unzip -t 'outputs/sales_toc_flow_v4_20260101_20260630/销售ToC业务流程核对底稿_V4_20260101-20260630.xlsx'
```

## 🧠 核心逻辑（方法论）

1. **抽取与标准化**：`reconcile_sales_toc.py` / `reconcile_huice_oms.py` 用 `lxml` 流式读取旺店通订单、惠策明细/汇总、OMS SQL、SAP 发票清单（`input/` 下 `旺店通订单清单/`、`惠策系统对账单清单/`、`惠策系统对账单汇总/`、`OMS ...2C单据.sql`、`发票清单：26.01.01-26.06.30/`），写入 `work_full/reconciliation.db`。
2. **分段 pairwise 钩稽**（`reconcile_sales_flow_v4.py`，期间 `START="2026-01-01"` 至 `END_EXCLUSIVE="2026-07-01"`）：
   - 订单→账单：按 `旺店通原始单号 = 惠策平台订单号`，金额差异 `ABS(...) <= 0.01` 判为一致；
   - 账单→OMS 月结：按 `月份 + 店铺/客户映射`，惠策成功金额 vs OMS 月结金额；
   - 数量核对：惠策证据数量 vs OMS 月结 Y001 数量，差异 `<= 0.000001` 判为数量一致；
   - OMS 月结→SAP：按 `OMS销售单号 + 物料编码 + 销售单位` 比较数量与含税金额，双向补充字段，仅 SAP 存在的键保留为例外。
3. **导出**：`export()` 写出 `output_flow_v4/` 下分项 CSV（`order_bill_recon.csv`、`bill_oms_month_recon.csv`、`order_bill_oms_qty_recon.csv`、`oms_sap_field_map.csv`、`huice_shop_map.csv`）与 `summary_v4.json`。
4. **底稿构建**：`build_sales_flow_v4.mjs` 等读取 CSV/JSON 生成带审计格式的 Excel，并扫描公式错误生成 `_qa_previews/`。程序仅呈现差异与匹配率，不自动构成会计调整或审计结论。

## 📋 输入与输出

- **输入**：`input/` 下客户导出的旺店通订单、惠策对账明细/汇总、OMS 日结月结 SQL（仅 `Y001` 月结）、SAP 2026 年 1–6 月发票清单（仅标准发票 2C）。
- **中间数据**：`reconciliation/work_full/reconciliation.db`（SQLite，运行时生成）。
- **输出**：
  - `reconciliation/output_flow_v4/`：完整 CSV 与汇总 `summary_v4.json`（订单—账单约 576 万行、账单—OMS 月结、数量、OMS—SAP 字段映射、店铺—客户映射等）；
  - `outputs/`：Excel 核对底稿（主底稿「销售ToC业务流程核对底稿_20260101-20260630.xlsx」含 13 张表；数据流/全链路底稿各版本）；Excel 对订单—账单明细仅嵌入异常优先的有限行，完整明细以 CSV 为准。

## ⚠️ 注意事项

- 数据脱敏：仓库不含真实客户业务数据，示例与说明均为脱敏/合成数据；实际运行需将客户导出文件放入 `input/`。
- 口径说明：核对期间、金额/数量阈值、业务范围（Y001、2C）以代码与仓库核对口径说明为准，本 README 仅作说明。
- 重要限制：惠策无商品级物料与数量，不能直接参与商品数量钩稽；订单、账单、发运、开票业务时点不同，跨月与退款可能形成合理差异；多对一汇总会提高可核对范围但可能抵销单笔差异，完整异常明细仍需保留。
- 审计结论：本项目结果用于财务核对与异常定位，不自动构成会计调整、收入确认或审计结论。

## 🔗 相关仓库

- https://github.com/Gvmeakiss/miaoke-sales-to-b-2026
- https://github.com/Gvmeakiss/miaoke-sales-to-b-2025
- https://github.com/Gvmeakiss/sales-three-match-miaoke-2026
- https://github.com/Gvmeakiss/sales-three-match-newhope-2026

## 📄 License

MIT（详见仓库 `LICENSE`）。

---

<div align="center">

*Disclaimer: Personal project and personal views. Not affiliated with or endorsed by KPMG or any client.*<br>
*本仓库为个人项目与个人观点，与任何前/现雇主及客户无关。*

</div>
