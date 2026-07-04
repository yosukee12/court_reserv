#!/bin/bash
set -euo pipefail

ZIP_PATH="/Users/yosuke/develop/chatgpt/court_reserv.zip"

rm -f "$ZIP_PATH"

zip -r "$ZIP_PATH" . \
  -x ".git/*" \
  -x "*/.git/*" \
  -x ".DS_Store" \
  -x "*/.DS_Store" \
  -x ".idea/*" \
  -x "*/.idea/*" \
  -x "__pycache__/*" \
  -x "*/__pycache__/*" \
  -x "*.pyc" \
  -x "*/.pytest_cache/*" \
  -x ".mypy_cache/*" \
  -x "*/.mypy_cache/*" \
  -x ".ruff_cache/*" \
  -x "*/.ruff_cache/*" \
  -x ".venv/*" \
  -x "*/.venv/*" \
  -x "venv/*" \
  -x "*/venv/*" \
  -x ".env" \
  -x "*/.env" \
  -x ".env.*" \
  -x "*/.env.*" \
  -x "config.local.ini" \
  -x "*/config.local.ini" \
  -x "local.settings.json" \
  -x "*/local.settings.json" \
  -x "docs/_build/*" \
  -x "*/docs/_build/*" \
  -x "logs/*" \
  -x "*/logs/*" \
  -x "*.log" \
  -x "*.csv" \
  -x "*.xlsx" \
  -x "make_zip.sh"

echo "Created: $ZIP_PATH"