#!/bin/bash
# 自動備份腳本：有變更就 commit + push 到 GitHub（所有分支）

REPO="$(git -C "$(dirname "$0")" rev-parse --show-toplevel 2>/dev/null || echo "/Users/USER/Documents/GitHub/Honten")"
LOG="$REPO/.auto_push.log"

cd "$REPO" || exit 1

CURRENT_BRANCH=$(git rev-parse --abbrev-ref HEAD)
TIMESTAMP=$(date '+%Y-%m-%d %H:%M')

# 目前分支有變更就 commit
if git status --porcelain | grep -q .; then
  git add .
  git commit -m "自動備份 $TIMESTAMP"
  echo "[$TIMESTAMP] commit on $CURRENT_BRANCH" >> "$LOG"
fi

# 推送所有有 tracking 的本地分支
git push --all origin >> "$LOG" 2>&1 && \
  echo "[$TIMESTAMP] push all branches OK" >> "$LOG" || \
  echo "[$TIMESTAMP] push failed" >> "$LOG"
