"""
LINE Messaging API client（純標準函式庫，無 SDK 依賴）
- reply_message(reply_token, text)：webhook 立即回覆
- push_to_user(user_id, text)：對個人主動推播
- push_to_group(group_id, text)：對群組推播
- verify_signature(body, signature, channel_secret)：驗 webhook 簽章
"""
import base64
import hashlib
import hmac
import json
import logging
import os
from typing import Optional
from urllib import request as _urlreq, error as _urlerr

logger = logging.getLogger(__name__)

LINE_TOKEN = os.getenv("LINE_CHANNEL_ACCESS_TOKEN", "").strip()
LINE_SECRET = os.getenv("LINE_CHANNEL_SECRET", "").strip()


def is_configured() -> bool:
    return bool(LINE_TOKEN)


def _post(url: str, payload: dict) -> bool:
    """通用 POST helper。無 token → dry-run console log。"""
    if not LINE_TOKEN:
        logger.info(f"[LINE dry-run] {url} payload={json.dumps(payload, ensure_ascii=False)[:200]}")
        print(f"\n📱 [LINE dry-run] {url}\n{'-'*40}\n{payload}\n{'-'*40}\n")
        return False
    try:
        req = _urlreq.Request(
            url,
            data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
            headers={
                "Authorization": f"Bearer {LINE_TOKEN}",
                "Content-Type": "application/json",
            },
            method="POST",
        )
        with _urlreq.urlopen(req, timeout=10) as resp:
            if 200 <= resp.status < 300:
                return True
            body = resp.read().decode("utf-8", errors="ignore")[:300]
            logger.error(f"[LINE FAIL] {resp.status}: {body}")
            return False
    except _urlerr.HTTPError as e:
        body = e.read().decode("utf-8", errors="ignore")[:300] if e.fp else ""
        logger.error(f"[LINE HTTPError] {e.code}: {body}")
        return False
    except Exception as e:
        logger.error(f"[LINE ERROR] {e}")
        return False


def reply_message(reply_token: str, text: str) -> bool:
    """webhook 內立即回覆（reply_token 一次性、有效期 1 分鐘）"""
    return _post("https://api.line.me/v2/bot/message/reply", {
        "replyToken": reply_token,
        "messages": [{"type": "text", "text": text[:4900]}],
    })


def push_to_user(line_user_id: str, text: str) -> bool:
    if not line_user_id:
        return False
    return _post("https://api.line.me/v2/bot/message/push", {
        "to": line_user_id,
        "messages": [{"type": "text", "text": text[:4900]}],
    })


def push_to_group(group_id: str, text: str) -> bool:
    if not group_id:
        return False
    return _post("https://api.line.me/v2/bot/message/push", {
        "to": group_id,
        "messages": [{"type": "text", "text": text[:4900]}],
    })


def verify_signature(body: bytes, signature_header: str) -> bool:
    """驗 LINE webhook X-Line-Signature。
    若未設定 LINE_CHANNEL_SECRET 視為 dev 模式直接通過（記 warning）。
    """
    if not LINE_SECRET:
        logger.warning("[LINE] LINE_CHANNEL_SECRET 未設定，略過簽章驗證（僅限 dev）")
        return True
    if not signature_header:
        return False
    digest = hmac.new(LINE_SECRET.encode("utf-8"), body, hashlib.sha256).digest()
    expected = base64.b64encode(digest).decode("utf-8")
    return hmac.compare_digest(expected, signature_header)
