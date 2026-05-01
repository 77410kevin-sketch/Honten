#!/bin/bash
# 月報 PREVIEW 經 Kevin 審核通過後，由本腳本推 GitHub + LINE
# 用法: ./publish_monthly.sh
set -e

PREVIEW_HTML="/tmp/monthly_industry_preview.html"
TODAY=$(date +%Y-%m-%d)

if [ ! -f "$PREVIEW_HTML" ]; then
    echo "❌ 找不到 PREVIEW HTML：$PREVIEW_HTML"
    exit 1
fi

echo "===== $(date '+%Y-%m-%d %H:%M:%S') 月報 PUBLISH 啟動 ====="
echo "HTML: $PREVIEW_HTML ($(wc -c < $PREVIEW_HTML) bytes)"

# 用 claude CLI 跑「只推 GitHub + LINE」的精簡 prompt
/Applications/cmux.app/Contents/Resources/bin/claude -p \
  --permission-mode bypassPermissions \
  --output-format text \
  --max-budget-usd 1 \
  --model claude-sonnet-4-6 \
  --allowedTools "Bash Read" \
<<EOF
你是 Kevin 的助理。Kevin 已審核 \`/tmp/monthly_industry_preview.html\` 通過。
請執行以下兩件事：

## 1. PUT 到 GitHub
- REPO: 77410kevin-sketch/honten-docs
- FILE: industry_analysis.html (master)
- GITHUB_PAT: <REDACTED-GH-PAT>

用 Python urllib：先 GET sha → 讀 \`/tmp/monthly_industry_preview.html\` → base64 encode → PUT 帶 sha。
取回 commit_url。失敗印 HTTP code + 內文。

## 2. LINE Multicast 推 4 人
- LINE Channel ID: 2009761937、Channel Secret: 1d1f82a7cf36aa1fcaf19eb5c01e1bc6
- 4 人：Alice U1e1afcc43fa9644dc26140b9735ccf4d、Kevin U31e3a20e0fcddc89faa25cc79e3b70f5、Jerry Uc5dd0084639462d3fc3e06d3dcbee447、James U09d987a36342404af0b67e1246d10e4b
- PAGES URL: https://77410kevin-sketch.github.io/honten-docs/industry_analysis.html

OAuth 取 token：POST https://api.line.me/v2/oauth/accessToken
form data: grant_type=client_credentials&client_id={Channel ID}&client_secret={Channel Secret}

訊息內容由 \`/tmp/monthly_preview_summary.txt\` 讀取（如果沒有則自己依 HTML 內容組）。
multicast endpoint: POST https://api.line.me/v2/bot/message/multicast

## 3. 最終輸出
✅ 成功：commit_url + LINE OK
❌ 失敗：HTTP code + 錯誤
EOF

echo ""
echo "===== $(date '+%Y-%m-%d %H:%M:%S') 月報 PUBLISH 結束 ====="
