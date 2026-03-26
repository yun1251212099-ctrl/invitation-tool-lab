#!/usr/bin/env bash
# 稳定启动 Streamlit（无交互、可写 credentials，避免 ~/.streamlit 权限问题）
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
PYTHON="${PYTHON:-}"
if [[ -z "$PYTHON" ]] && [[ -x "$ROOT/.venv/bin/python" ]]; then
  PYTHON="$ROOT/.venv/bin/python"
fi
PYTHON="${PYTHON:-python3}"
export STREAMLIT_HOME="${STREAMLIT_HOME:-$ROOT/.streamlit_home}"
mkdir -p "$STREAMLIT_HOME/.streamlit"
if [[ ! -f "$STREAMLIT_HOME/.streamlit/credentials.toml" ]]; then
  printf '%s\n' '[general]' 'email = ""' >"$STREAMLIT_HOME/.streamlit/credentials.toml"
fi
export STREAMLIT_BROWSER_GATHER_USAGE_STATS="${STREAMLIT_BROWSER_GATHER_USAGE_STATS:-false}"
cd "$ROOT"
exec "$PYTHON" -m streamlit run app.py --server.port "${STREAMLIT_PORT:-8599}" --server.address 0.0.0.0 "$@"
