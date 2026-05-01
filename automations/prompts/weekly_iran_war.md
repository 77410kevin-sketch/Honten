你是 Kevin 的智能助理。請執行以下任務：生成美伊戰爭情勢週報 **HTML 檔**，推 GitHub，LINE 通知 Kevin。

====================================================================
## 重要原則（違反則失敗）
====================================================================

1. 禁止 curl，一律 Python urllib（UTF-8 安全、錯誤訊息明確）。
2. **HTML 用 git CLI 推**（已 clone 在 `/Users/kevin/Projects/HonTen/honten-docs/`），不要用 GitHub API（PAT 是 read-only）。
3. **LINE 通知永遠都要發**（成功→報連結；失敗→報錯誤）。
4. 失敗顯示實際 HTTP code + 回應內文。

====================================================================
## 步驟 0：取得今天日期
====================================================================

使用 Bash：`date +%Y-%m-%d` → 記為 TODAY。
計算衝突天數：從 2026-02-28 起算到 TODAY。
NEXT_SUN = TODAY + 7 days（用 Python `datetime` 算）。

====================================================================
## 步驟 1：搜尋最新資料（WebSearch）
====================================================================

各執行一次（共 4 次）：
- 「美伊戰爭 最新情勢」+ TODAY
- 「US Iran war latest news」+ TODAY
- 「美伊戰爭 油價 台灣 經濟影響」+ TODAY
- 「Iran ceasefire negotiations」+ TODAY

整理重點事件、油價數據、台灣影響、外交進展，至少蒐集 5 個可用來源連結。

====================================================================
## 步驟 2：生成 HTML（Bootstrap 5 單頁）
====================================================================

寫入 `/Users/kevin/Projects/HonTen/honten-docs/iran_war_weekly.html`

結構：
1. `<head>`：UTF-8 / viewport / Bootstrap 5 CDN（jsdelivr）/ Bootstrap Icons CDN / Noto Sans TC 字型 / 自訂 CSS
2. Header：「⚔️ 美伊戰爭情勢週報」+ 副標題顯示「📅 報告日期：{TODAY} ・ 衝突第 N 天 ・ 🔄 下次更新：{NEXT_SUN}」
3. 頂部 callout（藍色 alert-info）：本週重點 3-5 條 bullet
4. **七大章節**（每章 `<h2>` + 內容）：
   - 一、戰爭背景與本週摘要（段落）
   - 二、當前軍事情勢（段落 + 表格：指標 / 現況 / 備註）
   - 三、外交談判動態（段落 + 重點 bullet）
   - 四、全球經濟衝擊（表格：項目 / 數據 / 說明）
   - 五、台灣受衝擊評估（表格：面向 / 現況 / 風險評估，含 CPI / 央行 / 科技產業 / 航運）
   - 六、後市情勢研判（**短期 1-3 月** + **中期 6-12 月** 兩段，分子標題）
   - 七、資料來源（無序清單，每筆 `<a href>` 含時間戳）
5. Footer：「🤖 自動由 Claude Agent 於 {TODAY} 生成 ・ 下次更新預計 {NEXT_SUN}」

CSS 重點：
- `body { font-family: 'Noto Sans TC', sans-serif; max-width: 960px; margin: 0 auto; padding: 2rem 1rem; }`
- `h2 { border-left: 4px solid #dc3545; padding-left: 0.75rem; margin-top: 2.5rem; }`
- `table { font-size: 0.92rem; }`
- 表格用 `class="table table-bordered table-hover"`

HTML 必須含 `<!DOCTYPE html>` 和 `</html>`。

====================================================================
## 步驟 3：git push 到 honten-docs
====================================================================

```bash
cd /Users/kevin/Projects/HonTen/honten-docs
git add iran_war_weekly.html
git commit -m "美伊戰爭情勢週報 {TODAY}"
git push
```

取回 commit hash 寫到 `/tmp/iran_war_commit.txt`。

若 `nothing to commit`（同日重跑）→ 跳過 push，但仍發 LINE 通知（用上次的 Pages URL）。

====================================================================
## 步驟 4：LINE 推送（給 Kevin 個人）
====================================================================

憑證：
- LINE_TOKEN: `pdCKG8MPsnqX+7NDRbp0eZE24ZkT5tm6M2EctFgB+VF5tfFkca3m7YrwRNPtzcKN0tum0gDbi8Gtx1mAUfV97qLyLCn8PYI6AS5mX3yWgC1z8OO9+LfH1KwE5ndlxM8MDFA2Eb4/OLe4rOtTjyb1IgdB04t89/1O/w1cDnyilFU=`
- LINE_USER_ID: `U2abebd9337887d1b430072ae19e5ad7c`
- PAGES_URL: `https://77410kevin-sketch.github.io/honten-docs/iran_war_weekly.html`

成功訊息：
```
⚔️ 美伊戰爭情勢週報 {TODAY}
衝突第 {N} 天（自 2026-02-28 起）

📌 本週重點：
• {重點 1，30 字內}
• {重點 2，30 字內}
• {重點 3，30 字內}

💰 油價：{布蘭特原油現價}/桶，{變動} vs 上週
🇹🇼 台灣影響：{1-2 句重點}

👀 完整報告：
{PAGES_URL}

🔄 下次更新：{NEXT_SUN}（週日）
```

失敗訊息：
```
⚠️ 美伊週報 {TODAY} 推送失敗
錯誤：{HTTP code + 摘要}
```

```python
import urllib.request, urllib.error, json
LINE_TOKEN = "pdCKG8MPsnqX+7NDRbp0eZE24ZkT5tm6M2EctFgB+VF5tfFkca3m7YrwRNPtzcKN0tum0gDbi8Gtx1mAUfV97qLyLCn8PYI6AS5mX3yWgC1z8OO9+LfH1KwE5ndlxM8MDFA2Eb4/OLe4rOtTjyb1IgdB04t89/1O/w1cDnyilFU="
LINE_USER_ID = "U2abebd9337887d1b430072ae19e5ad7c"
text = "..."  # 組成功訊息
payload = {"to": LINE_USER_ID, "messages": [{"type": "text", "text": text}]}
req = urllib.request.Request(
    "https://api.line.me/v2/bot/message/push",
    data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
    headers={"Authorization": f"Bearer {LINE_TOKEN}", "Content-Type": "application/json"},
    method="POST")
try:
    with urllib.request.urlopen(req, timeout=15) as r:
        print(f"LINE OK HTTP {r.status}")
except urllib.error.HTTPError as e:
    print(f"LINE HTTPError {e.code}: {e.read().decode()[:300]}")
```

====================================================================
## 步驟 5：最終輸出（一句話）
====================================================================

- 成功：`✅ 週報完成：{Pages URL}　📦 commit {hash}　📱 LINE 已通知`
- HTML 寫入失敗：`⚠️ HTML 寫入失敗：{錯誤}`
- git push 失敗：`⚠️ git push 失敗：{錯誤}　📱 LINE 已通知`
- LINE 失敗：`⚠️ LINE 失敗 HTTP {code}`

核心提醒：
1. **HTML 用 git CLI 推**（不要用 GitHub API PAT）
2. **不需要 Notion**，整個流程是 WebSearch → HTML → git → LINE
3. LINE 訊息控制在 1500 字內
4. HTML 必須含 `<!DOCTYPE html>` 和 `</html>`
5. 衝突天數計算：`(date(TODAY) - date(2026-02-28)).days`
