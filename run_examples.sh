#!/usr/bin/env bash
set -euo pipefail

python3 scripts/price_daily.py --input config/market_price_run.example.json
python3 scripts/manager_weekly.py --input config/manager_weekly_run.example.json
