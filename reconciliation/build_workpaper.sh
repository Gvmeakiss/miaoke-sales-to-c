#!/usr/bin/env bash
set -euo pipefail

project_root="$(cd "$(dirname "$0")/.." && pwd)"
: "${WORKSPACE_PYTHON:=python3}"
: "${WORKSPACE_NODE:?请设置 WORKSPACE_NODE 为 Codex 工作区 Node.js 路径}"
: "${WORKSPACE_NODE_MODULES:?请设置 WORKSPACE_NODE_MODULES 为 Codex 工作区 node_modules 路径}"

if [[ "${SKIP_SETTLEMENT_REFRESH:-0}" != "1" ]]; then
  "$WORKSPACE_PYTHON" "$project_root/reconciliation/refresh_settlement_oms_workpaper.py"
fi

settlement_huice_result="$project_root/reconciliation/results/settlement_vs_huice_reconciliation.json"
if [[ "${SKIP_SETTLEMENT_HUICE_REFRESH:-0}" != "1" ]]; then
  refresh_settlement_huice=0
  if [[ ! -f "$settlement_huice_result" ]]; then
    refresh_settlement_huice=1
  elif [[ "$project_root/reconciliation/analyze_settlement_vs_huice.py" -nt "$settlement_huice_result" ]] || \
       [[ "$project_root/reconciliation/work/reconciliation.db" -nt "$settlement_huice_result" ]]; then
    refresh_settlement_huice=1
  elif find "$project_root/input/对账明细（to oms 月结）" "$project_root/input/惠策系统对账单清单" \
       -maxdepth 1 -type f -name '*.xlsx' -newer "$settlement_huice_result" -print -quit | grep -q .; then
    refresh_settlement_huice=1
  fi
  if [[ "$refresh_settlement_huice" == "1" ]]; then
    PYTHONPATH="$project_root/reconciliation" "$WORKSPACE_PYTHON" "$project_root/reconciliation/analyze_settlement_vs_huice.py"
  fi
fi

task_tmp="$(mktemp -d)"
cleanup() {
  test ! -L "$task_tmp/node_modules" || unlink "$task_tmp/node_modules"
  test ! -f "$task_tmp/build_sales_workpaper.mjs" || unlink "$task_tmp/build_sales_workpaper.mjs"
  rmdir "$task_tmp"
}
trap cleanup EXIT

ln -s "$WORKSPACE_NODE_MODULES" "$task_tmp/node_modules"
cp "$project_root/reconciliation/build_sales_workpaper.mjs" "$task_tmp/build_sales_workpaper.mjs"
cd "$task_tmp"
SALES_TOC_ROOT="$project_root" "$WORKSPACE_NODE" "$task_tmp/build_sales_workpaper.mjs"
