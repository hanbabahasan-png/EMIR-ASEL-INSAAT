#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$(realpath "$0")")"
if [[ ! -x ".venv/bin/python" ]]; then
  echo "Sanal ortam bulunamadı. README_TR.md içindeki Linux kurulum adımlarını uygulayın." >&2
  exit 1
fi
.venv/bin/python scripts/preflight.py
exec .venv/bin/python -m streamlit run app.py
