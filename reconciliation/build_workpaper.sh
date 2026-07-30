#!/usr/bin/env bash
set -euo pipefail

project_root="$(cd "$(dirname "$0")/.." && pwd)"
: "${WORKSPACE_NODE:?请设置 WORKSPACE_NODE 为 Codex 工作区 Node.js 路径}"
: "${WORKSPACE_NODE_MODULES:?请设置 WORKSPACE_NODE_MODULES 为 Codex 工作区 node_modules 路径}"

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
