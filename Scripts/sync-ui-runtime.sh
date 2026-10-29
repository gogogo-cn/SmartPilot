#!/usr/bin/env bash
# ============================================================
# sync-ui-runtime.sh — 把 UI/ 同步到全部运行时目录并刷新缓存版本号
# ============================================================
# 【为什么必须有这个脚本】2026-10-05 晚上连续 4 轮"改了 JS 但线上没生效"，
# 根因是 WebView2 按 <script src="...js?v=版本号"> 做缓存：
#   1. 旧 sync-ui-files.ps1 只同步 UI/js/*.js 顶层文件，
#      **不覆盖 chat-v2/ 等子目录**，也不同步 Hosts 下两处运行时目录；
#   2. 即使把 JS 文件复制过去了，index.html 里的 ?v= 参数没变，
#      WebView2 仍加载旧版本 —— 文件是新的、内存里是旧的，最难查。
# 本脚本一次性解决：递归同步 + 统一刷新 ?v= +逐字节校验。
#
# 用法：bash Scripts/sync-ui-runtime.sh
# 之后重开 SmartPilot 任务面板即可生效（不必重启 Office）。
# ============================================================
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
SRC="$ROOT/UI"

# 全部运行时目录（顺序无关）
TARGETS=(
  "$ROOT/bin/Debug/UI"
  "$ROOT/bin/Release/UI"
  "$ROOT/Hosts/PptAIVSTOAddIn/bin/Debug/UI"
  "$ROOT/Hosts/WordAIVSTOAddIn/bin/Debug/UI"
)

# 生成 12 位时间戳版本号：vYYYYMMDDHHMM
STAMP="$(date +%Y%m%d%H%M)"
VER="v${STAMP}"

echo "==> 缓存版本号: $VER"

# 1) 先刷新源 index.html 的所有 ?v= 参数
#只替换形如 ?v=v202609262700 的 12 位版本号，不动其它 query
sed -i -E "s/\?v=v[0-9]{12}/?v=${VER}/g" "$SRC/index.html"
# 同时刷新内联 window.__APP_VERSION__（状态栏/关于面板直接读它，不随 ?v= 走，
# 停在旧版本会误导排查）。.ps1 版 2026-10-12 已有此逻辑而 .sh 漏了 → 两边行为
# 分叉，2026-10-18 补齐（此前 .sh 同步后 __APP_VERSION__ 一直停在旧版本号）。
sed -i -E "s/(window\.__APP_VERSION__[[:space:]]*=[[:space:]]*)'v[0-9]{12}'/\1'${VER}'/g" "$SRC/index.html"
echo "==> 源 index.html ?v= 与 __APP_VERSION__ 已刷新"

# 2) 递归同步到每个目标目录（--delete 会删掉目标里源已删除的文件，避免旧文件残留）
for T in "${TARGETS[@]}"; do
  if [ ! -d "$T" ]; then
    echo "    跳过（不存在）: $T"
    continue
  fi
  mkdir -p "$T"
  # 【不要用 rsync --delete】运行时目录里存在源UI/ 没有的文件（构建产物、
  # 其它模块的 html、.bak 备份等），--delete 会把它们全删掉。
  # 这里只做"覆盖式复制"：源有的→覆盖，源没有的→保留。
  if command -v rsync >/dev/null 2>&1; then
    rsync -a "$SRC/" "$T/"
  else
    cp -r "$SRC/." "$T/"
  fi
  # 源 index.html 已在第1 步刷成统一版本号，直接复制即可。
  # 【不要在这里再对目标 sed 一次】—— 各目标单独刷会产生不同版本号，
  # 既让"4处一致性"检查永远失败，也会让浏览器按 query 缓存出多份副本。
  echo "    已同步: $T"
done

# 3) 逐字节校验：源与每个目标的关键文件必须完全一致
echo "==> 校验一致性"
FILES=(
  "index.html"
  "js/chat-v2/agent-loop.js"
  "js/chat-v2/services/tool-registry.js"
  "js/chat-v2/services/tool-registry.util.js"
)
FAILED=0
for T in "${TARGETS[@]}"; do
  [ -d "$T" ] || continue
  for F in "${FILES[@]}"; do
    if [ ! -f "$T/$F" ]; then
      echo "    缺失: $T/$F"; FAILED=1; continue
    fi
    # 忽略 CRLF 差异只比内容哈希过于脆弱，这里直接 diff（去行尾 CR）
    if ! diff -q <(tr -d '\r' < "$SRC/$F") <(tr -d '\r' < "$T/$F") >/dev/null 2>&1; then
      echo "    不一致: $T/$F"; FAILED=1
    fi
  done
done

if [ "$FAILED" -ne 0 ]; then
  echo "!! 校验失败：存在缺失或不一致文件"
  exit 1
fi

echo "==> 全部一致（$VER）"
echo "==> 现在重开 SmartPilot 任务面板即可生效。"
