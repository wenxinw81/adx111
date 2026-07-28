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
"$PYTHON_BIN" -m pip install -e .
"$PYTHON_BIN" -m pip install pyinstaller
ADX_PACKAGE_VARIANT=direct "$PYTHON_BIN" -m PyInstaller --clean --noconfirm packaging/ADXReportAgent.spec
ADX_PACKAGE_VARIANT=ssh "$PYTHON_BIN" -m PyInstaller --clean --noconfirm packaging/ADXReportAgent.spec
cd dist
rm -f ADXReportAgent-Direct-macOS.zip ADXReportAgent-SSH-macOS.zip
zip -qr ADXReportAgent-Direct-macOS.zip ADXReportAgent-Direct.app
zip -qr ADXReportAgent-SSH-macOS.zip ADXReportAgent-SSH.app
echo "Built dist/ADXReportAgent-Direct-macOS.zip"
echo "Built dist/ADXReportAgent-SSH-macOS.zip"
