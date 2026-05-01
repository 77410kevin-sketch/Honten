#!/bin/bash
# 每月 1 號 09:03 由 launchd 觸發，跑鴻騰產業分析月報
set -e
export PATH="/opt/homebrew/bin:/usr/bin:/bin:/Applications/cmux.app/Contents/Resources/bin:$PATH"
cd /Users/kevin/Projects/HonTen/automations

echo "===== $(date '+%Y-%m-%d %H:%M:%S') 月報 routine 啟動 ====="
claude -p \
  --permission-mode bypassPermissions \
  --output-format text \
  --max-budget-usd 5 \
  --model claude-sonnet-4-6 \
  --allowedTools "Bash WebSearch WebFetch Read Write Edit Glob Grep" \
  < prompts/monthly_industry.md
echo "===== $(date '+%Y-%m-%d %H:%M:%S') 月報 routine 完成 ====="
