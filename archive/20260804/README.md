# 2026-08-04 项目过程归档说明

本目录保存正式交付版本形成过程中产生、但不应包含在Audit Team交付包中的过程资料。归档操作未修改或移动`input/`内的客户原始资料，也未改变正式核对结果。

## 归档内容

| 目录 | 内容 | 归档原因 | 恢复方式 |
|---|---|---|---|
| `过程检查/` | Excel检查日志及各Sheet页面预览 | 仅用于公式和版式QA，不属于正式底稿 | 可直接移回原位置，或重新运行底稿QA生成 |
| `中间结果/formal_order_match_detail/` | 拆分前的三类全量CSV及控制数 | 已由正式交付目录内8个CSV分片及校验索引替代 | 重新运行`reconciliation/prepare_formal_order_match_detail.py`可生成 |
| `弃用试验脚本/` | 大体量明细转Excel的试验脚本 | 受单工作簿单元格体量限制，最终交付采用CSV分片 | 仅供历史追溯，不纳入正式重跑流程 |
| `系统缓存/` | `.DS_Store`及Python字节码 | 系统自动生成，与业务结果无关 | 无需恢复，可由系统重新生成 |

## 当前正式位置

- 正式交付：`outputs/sales_toc_workpaper_final_20260101_20260630/`
- 可复现代码：`reconciliation/`
- 当前核对结果：`reconciliation/results/`
- 可复算数据库及抽取缓存：`reconciliation/work/`

本归档目录应只读保留。若后续形成新正式版本，请新建独立日期目录，不覆盖本次归档。
