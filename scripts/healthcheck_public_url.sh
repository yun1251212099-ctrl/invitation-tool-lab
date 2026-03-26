#!/usr/bin/env bash
# 健康检查：本机端口 + 本地 HTTP +（可选）公网 URL
set -euo pipefail
PORT="${STREAMLIT_PORT:-8599}"
LOCAL_URL="${LOCAL_URL:-http://127.0.0.1:${PORT}}"
PUBLIC_URL="${PUBLIC_URL:-}"

fail() { echo "FAIL: $*" >&2; exit 1; }

if ! command -v curl >/dev/null 2>&1; then
  fail "需要 curl"
fi

if ! lsof -nP -iTCP:"$PORT" -sTCP:LISTEN >/dev/null 2>&1; then
  fail "端口 ${PORT} 未监听（请先启动 Streamlit）"
fi

http_code() {
  curl -sS -o /dev/null -w "%{http_code}" --max-time 20 "$1" || echo "000"
}

c=$(http_code "$LOCAL_URL")
if [[ "$c" =~ ^(200|301|302|303|304)$ ]]; then
  echo "OK local ${LOCAL_URL} HTTP ${c}"
else
  fail "本地 ${LOCAL_URL} 返回 HTTP ${c}（预期 200/302/301/304）"
fi

if [[ -n "$PUBLIC_URL" ]]; then
  c2=$(http_code "$PUBLIC_URL")
  if [[ "$c2" =~ ^(200|301|302|303|304)$ ]]; then
    echo "OK public ${PUBLIC_URL} HTTP ${c2}"
  else
    fail "公网 ${PUBLIC_URL} 返回 HTTP ${c2}"
  fi
fi

echo "healthcheck passed"
