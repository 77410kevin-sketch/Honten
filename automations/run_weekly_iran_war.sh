#!/bin/bash
# 每週日 09:00 由 launchd 觸發，跑美伊戰爭情勢週報
set -e
export PATH="/opt/homebrew/bin:/usr/bin:/bin:/Applications/cmux.app/Contents/Resources/bin:$PATH"
cd /Users/kevin/Projects/HonTen/automations

echo "===== $(date '+%Y-%m-%d %H:%M:%S') 美伊週報 routine 啟動 ====="
claude -p \
  --permission-mode bypassPermissions \
  --output-format text \
  --max-budget-usd 3 \
  --model claude-sonnet-4-6 \
  --allowedTools "Bash WebSearch WebFetch Read Write Edit Glob Grep" \
  < prompts/weekly_iran_war.md
echo "===== $(date '+%Y-%m-%d %H:%M:%S') 美伊週報 routine 完成 ====="
