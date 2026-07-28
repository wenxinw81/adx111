#!/usr/bin/env bash
set -euo pipefail

PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
if [[ -z "${PYTHON_BIN:-}" ]]; then
  if [[ -x "$PROJECT_DIR/.venv/bin/python3" ]]; then
    PYTHON_BIN="$PROJECT_DIR/.venv/bin/python3"
  elif [[ -x "$HOME/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3" ]]; then
    PYTHON_BIN="$HOME/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3"
  else
    PYTHON_BIN="python3"
  fi
fi

cd "$PROJECT_DIR"
exec "$PYTHON_BIN" -m adx_report_agent.web_app --config configs/agent.direct.example.json --host "${ADX_WEB_HOST:-0.0.0.0}" --port "${ADX_WEB_PORT:-8787}" "$@"
