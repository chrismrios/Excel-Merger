#!/usr/bin/env bash
set -e

# Go to script dir
cd "$(dirname "$0")"

# Create venv if missing
if [ ! -d ".venv" ]; then
  python3 -m venv .venv
fi

# Activate venv
source .venv/bin/activate

# Upgrade pip quietly
python -m pip install --upgrade pip >/dev/null 2>&1 || true

# Install deps if missing
python - <<'PY'
import pkgutil, sys
reqs = ['streamlit','pandas','openpyxl','rapidfuzz','xlrd']
missing = [r for r in reqs if not pkgutil.find_loader(r)]
sys.exit(1 if missing else 0)
PY
if [ $? -ne 0 ]; then
  pip install -r requirements.txt
fi

# Launch
streamlit run app.py --server.headless true