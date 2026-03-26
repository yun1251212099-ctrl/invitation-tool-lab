#!/usr/bin/env bash
# 启动 Cloudflare 持久隧道（需已完成 cloudflared tunnel login + create + route dns）
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
CFG="${CLOUDFLARED_CONFIG:-}"

if [[ -z "$CFG" ]]; then
  if [[ -f "$HOME/.cloudflared/config.yml" ]] && [[ -s "$HOME/.cloudflared/config.yml" ]]; then
    CFG="$HOME/.cloudflared/config.yml"
  elif [[ -f "$ROOT/cloudflared/config.local.yml" ]]; then
    CFG="$ROOT/cloudflared/config.local.yml"
  fi
fi

if [[ -z "$CFG" ]] || [[ ! -f "$CFG" ]]; then
  echo "未找到 Cloudflare 隧道配置。" >&2
  echo "请先阅读: $ROOT/scripts/LONG_TERM_PUBLIC_ACCESS.md" >&2
  echo "或设置环境变量: export CLOUDFLARED_CONFIG=/path/to/config.yml" >&2
  exit 1
fi

echo "Using config: $CFG"
exec cloudflared --config "$CFG" tunnel run
