#!/usr/bin/env bash
# run.sh - Linux / EC2 launcher. Always writes both outputs:
#   reports/GT_Report_<from>_to_<to>.xlsx   full history (+ live TWS data)
#   reports/daily/GT_Report_<day>.xlsx      one workbook per trading day
# Extra arguments are passed to the script, e.g.
#   bash run.sh --from 2026-10-01 --to 2026-10-31
#   bash run.sh --no-tws
set -euo pipefail
cd "$(dirname "$0")"

VENV_DIR=".venv"
SCRIPT="gt_flex_report.py"

if [ ! -f ".env" ]; then
    echo "Error: .env file not found. Copy .env.example to .env and fill in credentials."
    exit 1
fi

if ! command -v python3 &>/dev/null; then
    echo "python3 not found. Installing..."
    sudo apt-get update -y && sudo apt-get install -y python3 python3-pip python3-venv
fi

if [ ! -d "$VENV_DIR" ]; then
    echo "Creating virtual environment..."
    python3 -m venv "$VENV_DIR"
fi

source "$VENV_DIR/bin/activate"
pip install --upgrade pip --quiet
pip install -r requirements.txt --quiet

echo "Running $SCRIPT..."
python "$SCRIPT" --daily "$@"
