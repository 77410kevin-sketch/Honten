#!/bin/bash
# 週六 09:00 由 launchd 觸發，跑 BBU 週報 dashboard
set -e
export PATH="/opt/homebrew/bin:/usr/bin:/bin:$PATH"
cd /Users/kevin/Downloads/kevin-Agant
/opt/homebrew/bin/python3 weekly_dashboard.py
