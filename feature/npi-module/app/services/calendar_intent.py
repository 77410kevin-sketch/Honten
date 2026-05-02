"""
LINE 訊息意圖解析（Claude API 抽欄位）

回傳結構：
{
  "intent": "ROOM | CAR | LEAVE | OUTING | QUERY | CANCEL | BIND | HELP | UNKNOWN",
  "start_at": "ISO 8601 或 null",
  "end_at": "ISO 8601 或 null",
  "all_day": bool,
  "resource_name": str | null,   # 「大會議室」「小藍」等原文
  "leave_type": "ANNUAL|SICK|PERSONAL" | null,
  "title": str,
  "customer": str | null,
  "notes": str,
  "query_type": "BALANCE|SCHEDULE" | null,
  "confidence": 0.0~1.0,
  "reasoning": "短解釋"
}
"""
import json
import logging
import os
import re
from datetime import datetime, timedelta
from typing import Optional
from urllib import request as _urlreq, error as _urlerr

logger = logging.getLogger(__name__)

ANTHROPIC_KEY = os.getenv("ANTHROPIC_API_KEY", "").strip()
ANTHROPIC_MODEL = os.getenv("ANTHROPIC_MODEL_FAST", "claude-haiku-4-5-20251001")

SYSTEM_PROMPT = """你是鴻騰電子行事曆助理。從使用者的中文訊息中抽取結構化資訊，**只回 JSON、不要其他文字**。

支援 intent：
- ROOM    : 借會議室
- CAR     : 借公務車
- LEAVE   : 請假
- OUTING  : 業務外出
- QUERY   : 查詢（特休餘額／行事曆）
- CANCEL  : 取消最近一筆預約／請假
- BIND    : 綁定 LINE 帳號（訊息含「綁定」「bind」+ 帳號名稱）
- HELP    : 求助
- UNKNOWN : 無法判斷

所有時間請以 ISO 8601 字串輸出（YYYY-MM-DDTHH:MM:00）。
時間最小單位為 30 分鐘，請對齊到 :00 或 :30。
公司可借資源：大會議室、小會議室、小藍（公務車）、小綠（公務車）。
假別 leave_type 限：ANNUAL（特休）、SICK（病假）、PERSONAL（事假）。

回傳格式（缺值請用 null，不是省略 key）：
{
  "intent": "...",
  "start_at": "...",
  "end_at": "...",
  "all_day": true|false,
  "resource_name": "...",
  "leave_type": "...",
  "title": "...",
  "customer": "...",
  "notes": "",
  "query_type": "...",
  "confidence": 0.85,
  "reasoning": "..."
}"""


def _build_user_prompt(text: str, today: datetime, user_name: str) -> str:
    next_mon = today + timedelta(days=(7 - today.weekday()) % 7 or 7)
    return f"""今天是 {today.strftime('%Y-%m-%d')}（週{['一','二','三','四','五','六','日'][today.weekday()]}）。
下週一是 {next_mon.strftime('%Y-%m-%d')}。
使用者名稱：{user_name}

訊息：「{text}」

範例：
輸入「借大會議室10點到12點 開週會」
→ {{"intent":"ROOM","start_at":"{today.strftime('%Y-%m-%d')}T10:00:00","end_at":"{today.strftime('%Y-%m-%d')}T12:00:00","all_day":false,"resource_name":"大會議室","leave_type":null,"title":"週會","customer":null,"notes":"","query_type":null,"confidence":0.95,"reasoning":"明確會議室+時段"}}

輸入「下週一請特休」
→ {{"intent":"LEAVE","start_at":"{next_mon.strftime('%Y-%m-%d')}T09:00:00","end_at":"{next_mon.strftime('%Y-%m-%d')}T18:00:00","all_day":true,"resource_name":null,"leave_type":"ANNUAL","title":"特休","customer":null,"notes":"","query_type":null,"confidence":0.9,"reasoning":"請特休全天"}}

輸入「我的特休還剩幾天」
→ {{"intent":"QUERY","start_at":null,"end_at":null,"all_day":false,"resource_name":null,"leave_type":"ANNUAL","title":"","customer":null,"notes":"","query_type":"BALANCE","confidence":0.95,"reasoning":"查餘額"}}

輸入「外出拜訪金士頓 2-4點」
→ {{"intent":"OUTING","start_at":"{today.strftime('%Y-%m-%d')}T14:00:00","end_at":"{today.strftime('%Y-%m-%d')}T16:00:00","all_day":false,"resource_name":null,"leave_type":null,"title":"拜訪金士頓","customer":"金士頓","notes":"","query_type":null,"confidence":0.92,"reasoning":"外出+客戶"}}

請輸出 JSON："""


def _call_claude(system: str, user: str) -> Optional[str]:
    if not ANTHROPIC_KEY:
        logger.warning("[Intent] ANTHROPIC_API_KEY 未設定，無法呼叫 Claude")
        return None
    payload = {
        "model": ANTHROPIC_MODEL,
        "max_tokens": 600,
        "system": system,
        "messages": [{"role": "user", "content": user}],
    }
    try:
        req = _urlreq.Request(
            "https://api.anthropic.com/v1/messages",
            data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
            headers={
                "x-api-key": ANTHROPIC_KEY,
                "anthropic-version": "2023-06-01",
                "Content-Type": "application/json",
            },
            method="POST",
        )
        with _urlreq.urlopen(req, timeout=15) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            blocks = data.get("content") or []
            for b in blocks:
                if b.get("type") == "text":
                    return b.get("text", "").strip()
            return None
    except _urlerr.HTTPError as e:
        body = e.read().decode("utf-8", errors="ignore")[:300] if e.fp else ""
        logger.error(f"[Intent HTTPError] {e.code}: {body}")
        return None
    except Exception as e:
        logger.error(f"[Intent ERROR] {e}")
        return None


def _extract_json(text: str) -> Optional[dict]:
    """從 LLM 輸出抽 JSON（容錯：可能被 ```json 包覆）"""
    if not text:
        return None
    m = re.search(r"\{[\s\S]*\}", text)
    if not m:
        return None
    try:
        return json.loads(m.group(0))
    except json.JSONDecodeError:
        return None


def parse_intent(raw_text: str, user_name: str = "", today: Optional[datetime] = None) -> dict:
    """
    主入口：解析 LINE 文字 → intent dict。
    若 ANTHROPIC_KEY 未設或 API 失敗 → 回傳 intent=UNKNOWN 但不拋例外。
    """
    text = (raw_text or "").strip()
    if not text:
        return {"intent": "UNKNOWN", "reasoning": "空訊息", "confidence": 0.0}
    today = today or datetime.now()

    out = _call_claude(SYSTEM_PROMPT, _build_user_prompt(text, today, user_name or ""))
    if not out:
        return {"intent": "UNKNOWN", "reasoning": "LLM 無回應", "confidence": 0.0}
    j = _extract_json(out)
    if not j:
        return {"intent": "UNKNOWN", "reasoning": f"LLM 回非 JSON：{out[:100]}", "confidence": 0.0}

    # 補預設欄位
    j.setdefault("intent", "UNKNOWN")
    for k in ("start_at", "end_at", "resource_name", "leave_type",
              "customer", "query_type", "reasoning"):
        j.setdefault(k, None)
    j.setdefault("all_day", False)
    j.setdefault("title", "")
    j.setdefault("notes", "")
    j.setdefault("confidence", 0.0)
    return j
