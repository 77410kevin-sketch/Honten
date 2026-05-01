你是 Kevin 的智能助理。請執行以下任務：生成美伊戰爭情勢週報並推送至 Notion，成功後再發 LINE 通知 Kevin。

====================================================================
## 重要原則（違反則失敗）
====================================================================

1. **絕對不可回報「網路封鎖」而放棄**——除非先執行步驟 3.1 的連線自測並證明 api.notion.com 無法連線（HTTP code=000 或逾時）。若只是 curl 指令出錯，那是指令問題不是網路問題。
2. **禁止使用 curl 推送 Notion 或 LINE**。必須用 Python urllib（UTF-8 安全、錯誤訊息明確）。
3. **Notion table 區塊的 rows 必須放在 `table.children` 內**（不是 block 的 children）。
4. **失敗時要顯示實際 HTTP code 與錯誤回應內文**，不可籠統寫「封鎖」。
5. **LINE 通知只在 Notion 推送成功後才送**；Notion 失敗時 LINE 要發送失敗通知，讓 Kevin 能立即處理。

====================================================================
## 步驟 0：取得今天日期
====================================================================

使用 Bash：`date +%Y-%m-%d` → 記為 TODAY。
計算衝突天數：從 2026-02-28 起算到 TODAY。
若 TODAY >= 2026-05-17，本週需額外加入橘色 callout（排程滿一個月提醒，內容見步驟 4）。

====================================================================
## 步驟 1：搜尋最新資料（WebSearch）
====================================================================

各執行一次（共 4 次）：
- 「美伊戰爭 最新情勢」+ 本週日期
- 「US Iran war latest news」+ 本週日期
- 「美伊戰爭 油價 台灣 經濟影響」+ 本週日期
- 「Iran ceasefire negotiations」+ 本週日期

整理重點事件、油價數據、台灣影響、外交進展，至少蒐集 5 個可用來源連結。

====================================================================
## 步驟 2：生成 Markdown 報告（選填，僅供備存）
====================================================================

寫入 `/tmp/iran_war_weekly_${TODAY}.md`，章節：
- 一、戰爭背景與本週摘要
- 二、當前軍事情勢（衝突天數、主要戰線、停火狀態）
- 三、外交談判動態
- 四、全球經濟衝擊（油價、供應鏈）
- 五、台灣受衝擊評估（CPI、央行政策、科技產業）
- 六、後市情勢研判（短中期）
- 七、資料來源（附連結）

====================================================================
## 步驟 3：推送至 Notion
====================================================================

### 3.1 連線自測

```python
import urllib.request, urllib.error, json
TOKEN = "ntn_q3464670585MehOPbJxIxdjJC2WY8tXLFvjUUfQ7F507x8"
req = urllib.request.Request(
    "https://api.notion.com/v1/users/me",
    headers={"Authorization": f"Bearer {TOKEN}", "Notion-Version": "2022-06-28"}
)
try:
    with urllib.request.urlopen(req, timeout=15) as r:
        print("OK", r.status)
except Exception as e:
    print("FAIL", type(e).__name__, str(e))
```

- 若印出 `OK 200` → 網路可通，進入 3.2。
- 若印出 `FAIL`  → **先印出完整錯誤訊息**，再進入步驟 5 的失敗通知流程。

### 3.2 推送頁面

寫入並執行 `/tmp/push_notion.py`：

```python
#!/usr/bin/env python3
import json, urllib.request, urllib.error

NOTION_TOKEN = "ntn_q3464670585MehOPbJxIxdjJC2WY8tXLFvjUUfQ7F507x8"
PARENT_ID = "3400ce82-ad79-8018-b5b7-fab3b2a1dff4"
TODAY = "<填入 TODAY>"

def T(c, link=None, bold=False):
    o = {"type": "text", "text": {"content": c}}
    if link: o["text"]["link"] = {"url": link}
    if bold: o["annotations"] = {"bold": True}
    return o

def para(t): return {"object":"block","type":"paragraph","paragraph":{"rich_text":t}}
def h1(c): return {"object":"block","type":"heading_1","heading_1":{"rich_text":[T(c)]}}
def bullet(t): return {"object":"block","type":"bulleted_list_item","bulleted_list_item":{"rich_text":t}}
def divider(): return {"object":"block","type":"divider","divider":{}}
def callout(t, color="blue_background", emoji="📋"):
    return {"object":"block","type":"callout",
            "callout":{"rich_text":t,"icon":{"type":"emoji","emoji":emoji},"color":color}}

# ⚠️ 關鍵：table rows 要放在 table.children 內
def table(headers, rows):
    ch = [{"object":"block","type":"table_row",
           "table_row":{"cells":[[T(h, bold=True)] for h in headers]}}]
    for r in rows:
        ch.append({"object":"block","type":"table_row",
                   "table_row":{"cells":[[T(c)] for c in r]}})
    return {"object":"block","type":"table",
            "table":{"table_width":len(headers),
                     "has_column_header":True,
                     "has_row_header":False,
                     "children":ch}}

blocks = [
    # 依步驟 4 的頁面結構填入
]

payload = {
    "parent": {"type":"page_id","page_id":PARENT_ID},
    "properties": {"title":{"title":[{"type":"text","text":{"content":f"美伊戰爭情勢週報 {TODAY}"}}]}},
    "icon": {"type":"emoji","emoji":"⚔️"},
    "children": blocks,
}

req = urllib.request.Request(
    "https://api.notion.com/v1/pages",
    data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
    headers={
        "Authorization": f"Bearer {NOTION_TOKEN}",
        "Content-Type": "application/json",
        "Notion-Version": "2022-06-28",
    },
    method="POST",
)
try:
    with urllib.request.urlopen(req, timeout=30) as r:
        res = json.loads(r.read().decode())
        with open("/tmp/notion_url.txt", "w") as f:
            f.write(res["url"])
        print(f"SUCCESS\nURL: {res['url']}")
except urllib.error.HTTPError as e:
    err = e.read().decode()[:800]
    with open("/tmp/notion_error.txt", "w") as f:
        f.write(f"HTTP {e.code}: {err}")
    print(f"HTTPError {e.code}\n{err}")
except Exception as e:
    with open("/tmp/notion_error.txt", "w") as f:
        f.write(f"{type(e).__name__}: {e}")
    print(f"EXCEPTION {type(e).__name__}: {e}")
```

====================================================================
## 步驟 4：Notion 頁面內容結構
====================================================================

依序建立 blocks：

1. 首行 callout（藍色背景 📋）：「📅 報告日期：{TODAY}　⚔️ 衝突天數：第 N 天（自 2026-02-28 起）　🔄 下次更新：{TODAY+7}」
2. 若 TODAY >= 2026-05-17，追加 callout（橘色背景 ⚠️）：「⚠️ 排程滿一個月提醒：美伊戰爭週報自動排程已運行超過一個月（自 2026-04-12 起）。」
3. heading_1「一、戰爭背景與本週摘要」+ 段落
4. heading_1「二、當前軍事情勢」+ 段落 + table（指標/現況/備註）
5. heading_1「三、外交談判動態」+ 段落
6. heading_1「四、全球經濟衝擊」+ table（項目/數據/說明）
7. heading_1「五、台灣受衝擊評估」+ table（面向/現況/風險評估）
8. heading_1「六、後市情勢研判」+ 段落（短期 / 中期兩段）
9. divider
10. heading_1「七、資料來源」+ bullet list（附超連結，至少 5 筆）
11. divider
12. callout（灰色背景 🤖）：「本報告由 Claude Agent 自動生成，資料採集截止日 {TODAY}。下次更新預計 {TODAY+7}（週日）。」

====================================================================
## 步驟 5：LINE 通知
====================================================================

**5.1 憑證**
- LINE_CHANNEL_TOKEN: `pdCKG8MPsnqX+7NDRbp0eZE24ZkT5tm6M2EctFgB+VF5tfFkca3m7YrwRNPtzcKN0tum0gDbi8Gtx1mAUfV97qLyLCn8PYI6AS5mX3yWgC1z8OO9+LfH1KwE5ndlxM8MDFA2Eb4/OLe4rOtTjyb1IgdB04t89/1O/w1cDnyilFU=`
- LINE_USER_ID: `U2abebd9337887d1b430072ae19e5ad7c`

**5.2 推送邏輯**

寫入並執行 `/tmp/push_line.py`：

```python
#!/usr/bin/env python3
import os, json, urllib.request, urllib.error

LINE_TOKEN = "pdCKG8MPsnqX+7NDRbp0eZE24ZkT5tm6M2EctFgB+VF5tfFkca3m7YrwRNPtzcKN0tum0gDbi8Gtx1mAUfV97qLyLCn8PYI6AS5mX3yWgC1z8OO9+LfH1KwE5ndlxM8MDFA2Eb4/OLe4rOtTjyb1IgdB04t89/1O/w1cDnyilFU="
LINE_USER_ID = "U2abebd9337887d1b430072ae19e5ad7c"
TODAY = "<填入 TODAY>"
CONFLICT_DAYS = <填入衝突天數>

if os.path.exists("/tmp/notion_url.txt"):
    with open("/tmp/notion_url.txt") as f:
        url = f.read().strip()
    text = (
        f"📊 美伊戰爭情勢週報 {TODAY}\n"
        f"⚔️ 衝突第 {CONFLICT_DAYS} 天（自 2026-02-28 起）\n"
        f"\n🔗 {url}\n"
        f"\n下次更新：{TODAY+7}（週日）"
    )
else:
    err = ""
    if os.path.exists("/tmp/notion_error.txt"):
        with open("/tmp/notion_error.txt") as f:
            err = f.read().strip()[:300]
    text = (
        f"⚠️ 美伊戰爭週報 {TODAY} 推送失敗\n"
        f"\n錯誤：{err}\n"
        f"\n請手動檢查"
    )

payload = {"to": LINE_USER_ID, "messages": [{"type": "text", "text": text}]}
req = urllib.request.Request(
    "https://api.line.me/v2/bot/message/push",
    data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
    headers={"Authorization": f"Bearer {LINE_TOKEN}", "Content-Type": "application/json"},
    method="POST",
)
try:
    with urllib.request.urlopen(req, timeout=15) as r:
        print(f"LINE OK HTTP {r.status}")
except urllib.error.HTTPError as e:
    print(f"LINE HTTPError {e.code}: {e.read().decode()[:300]}")
except Exception as e:
    print(f"LINE EXCEPTION {type(e).__name__}: {e}")
```

**注意**：LINE 推送不管 Notion 成功或失敗都要執行一次。

====================================================================
## 步驟 6：最終輸出
====================================================================

- 成功：「✅ 週報完成：{notion_url}　📱 LINE 已通知」
- Notion API 錯誤：「❌ Notion API 錯誤 HTTP {code}：{原始訊息}　📱 LINE 已通知 Kevin」
- 真正網路封鎖：「⚠️ api.notion.com 連線失敗：{錯誤}　📱 LINE 已通知 Kevin」

記住：**3.1 若回 OK 200，就一定要完成推送，不可中途放棄；LINE 通知永遠都要發。**
