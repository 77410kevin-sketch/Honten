"""
行事曆 LINE Webhook + dev 測試入口

實際 LINE webhook：POST /calendar/line/webhook
dev 測試入口（直接送純文字）：POST /calendar/line/test {text, username|line_user_id}
"""
import json
import logging
from datetime import datetime, timedelta
from typing import Optional

from fastapi import APIRouter, Depends, Request, HTTPException, Form
from fastapi.responses import JSONResponse, PlainTextResponse
from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy.orm import selectinload

from app.database import get_db
from app.models.user import User, Role
from app.models.calendar import (
    CalendarEvent, CalendarResource, EventType, EventStatus, ResourceType,
    LineMessageLog,
)
from app.services import line_bot
from app.services.calendar_intent import parse_intent
from app.routes.calendar import (
    LEAVE_LABEL, TYPE_LABEL, find_conflicts, _calc_leave_days, _pick_approver,
)

logger = logging.getLogger(__name__)
router = APIRouter(prefix="/calendar/line")
from fastapi.templating import Jinja2Templates
templates = Jinja2Templates(directory="app/templates")


# ── 工具：根據 line_user_id 找 User ──
async def _find_user_by_line(db: AsyncSession, line_user_id: str) -> Optional[User]:
    if not line_user_id:
        return None
    r = await db.execute(select(User).where(User.line_user_id == line_user_id))
    return r.scalars().first()


async def _find_user_by_username(db: AsyncSession, username: str) -> Optional[User]:
    if not username:
        return None
    r = await db.execute(select(User).where(User.username == username.lower()))
    return r.scalars().first()


async def _find_resource_by_name(db: AsyncSession, name: str) -> Optional[CalendarResource]:
    if not name:
        return None
    r = await db.execute(select(CalendarResource).where(
        CalendarResource.name == name.strip(),
        CalendarResource.active == True,
    ))
    row = r.scalars().first()
    if row:
        return row
    # 模糊匹配（包含關鍵字）
    r = await db.execute(select(CalendarResource).where(CalendarResource.active == True))
    rows = list(r.scalars().all())
    for row in rows:
        if row.name in name or name in row.name:
            return row
    return None


def _parse_iso(s: Optional[str]) -> Optional[datetime]:
    if not s:
        return None
    try:
        return datetime.fromisoformat(s.replace("Z", "+00:00").split("+")[0])
    except Exception:
        return None


def _snap30(dt: datetime) -> datetime:
    """對齊 30 分鐘"""
    if not dt:
        return dt
    m = dt.minute
    snapped = round(m / 30) * 30
    if snapped == 60:
        return dt.replace(minute=0, second=0, microsecond=0) + timedelta(hours=1)
    return dt.replace(minute=snapped, second=0, microsecond=0)


# ── 業務邏輯：處理一則 LINE 文字訊息 ──
async def handle_text_message(
    db: AsyncSession,
    raw_text: str,
    line_user_id: Optional[str] = None,
    user: Optional[User] = None,
) -> str:
    """主處理函式：回傳要回覆給使用者的純文字。"""
    text = (raw_text or "").strip()
    if not text:
        return "請輸入要做什麼，例如：「借大會議室 10:00 到 11:30」"

    # ── 0. 帳號綁定流程（不需 user）──
    if text.lower().startswith(("綁定", "bind ", "/bind ")):
        parts = text.split(maxsplit=1)
        if len(parts) < 2:
            return "請輸入：綁定 <你的系統帳號>，例如：綁定 sales01"
        username = parts[1].strip().lower()
        target = await _find_user_by_username(db, username)
        if not target:
            return f"❌ 找不到帳號「{username}」"
        if not line_user_id:
            return "❌ 無法取得您的 LINE userId（dev 測試請提供 line_user_id）"
        target.line_user_id = line_user_id
        await db.commit()
        return f"✅ 已綁定：{target.display_name}（{target.username}）\n之後可直接打字操作行事曆。"

    # ── 1. 必須已綁定 ──
    if not user:
        return "請先綁定您的系統帳號：\n輸入「綁定 <帳號>」例如：綁定 sales01\n（密碼帳號可詢問管理員）"

    # ── 2. 內建快速指令 ──
    low = text.lower()
    if low in ("help", "?", "幫助", "說明"):
        return _help_text()
    if "特休" in text and ("剩" in text or "剩餘" in text or "幾天" in text):
        return await _query_leave_balance(db, user)
    if low in ("今天", "今日", "today") or "今天行事曆" in text:
        return await _list_today_events(db, user)
    if low in ("本週", "這週", "本周") or "週" in text and "行事曆" in text:
        return await _list_week_events(db, user)
    if text.startswith(("取消", "/cancel")):
        return await _cancel_latest(db, user)

    # ── 3. LLM 抽 intent ──
    parsed = parse_intent(text, user_name=user.display_name)

    # 紀錄到 LineMessageLog
    log = LineMessageLog(
        user_id=user.id, line_user_id=line_user_id,
        raw_text=text, intent=parsed.get("intent"),
        parsed_json=json.dumps(parsed, ensure_ascii=False),
    )
    db.add(log)

    intent = parsed.get("intent")
    if intent == "QUERY":
        if parsed.get("query_type") == "BALANCE":
            await db.commit()
            return await _query_leave_balance(db, user)
        await db.commit()
        return await _list_today_events(db, user)
    if intent == "CANCEL":
        await db.commit()
        return await _cancel_latest(db, user)
    if intent == "HELP" or intent == "UNKNOWN":
        await db.commit()
        return f"⚠️ 無法解析您的訊息。{_help_text()}"

    # ROOM / CAR / LEAVE / OUTING → 建立事件
    s = _parse_iso(parsed.get("start_at"))
    e = _parse_iso(parsed.get("end_at"))
    if not s or not e:
        await db.commit()
        return "⚠️ 無法判斷時段，請明確說出開始與結束時間。例：「借大會議室 5/8 10:00 到 11:30」"
    s, e = _snap30(s), _snap30(e)
    if e <= s:
        await db.commit()
        return "⚠️ 結束時間需大於開始時間"

    title = (parsed.get("title") or "").strip()

    if intent in ("ROOM", "CAR"):
        et = EventType.ROOM if intent == "ROOM" else EventType.CAR
        res = await _find_resource_by_name(db, parsed.get("resource_name") or "")
        if not res:
            available = await _list_resources_text(db, ResourceType.ROOM if intent == "ROOM" else ResourceType.CAR)
            await db.commit()
            return f"⚠️ 找不到資源「{parsed.get('resource_name') or '?'}」\n目前可借：{available}"
        if res.type != (ResourceType.ROOM if intent == "ROOM" else ResourceType.CAR):
            await db.commit()
            return f"⚠️ 「{res.name}」不是{TYPE_LABEL[et.value]}"

        # 衝突檢測
        conflicts = await find_conflicts(db, et, res.id, s, e)
        if conflicts:
            c = conflicts[0]
            owner = c.user.display_name if c.user else "?"
            await db.commit()
            return (
                f"⚠️ 該時段已被借走\n"
                f"「{res.name}」{c.start_at.strftime('%m-%d %H:%M')}–{c.end_at.strftime('%H:%M')}\n"
                f"借用人：{owner}\n"
                f"請改其它時段或資源。"
            )

        ev = CalendarEvent(
            event_type=et, title=title or f"{res.name} 預約",
            start_at=s, end_at=e, all_day=bool(parsed.get("all_day")),
            user_id=user.id, resource_id=res.id,
            status=EventStatus.APPROVED,
            customer_name=(parsed.get("customer") or "").strip() or None,
            notes=(parsed.get("notes") or "").strip() or None,
            raw_text=text, parsed_json=json.dumps(parsed, ensure_ascii=False),
        )
        db.add(ev)
        await db.commit()
        await db.refresh(ev)
        log.result_event_id = ev.id
        await db.commit()
        return (
            f"✅ 已預約「{res.name}」\n"
            f"{s.strftime('%m-%d %H:%M')}–{e.strftime('%H:%M')}\n"
            f"標題：{ev.title}"
        )

    if intent == "LEAVE":
        leave_code = parsed.get("leave_type") or "ANNUAL"
        if leave_code not in LEAVE_LABEL:
            leave_code = "ANNUAL"
        days = _calc_leave_days(s, e)
        approver = await _pick_approver(db, user, days)
        if not approver:
            await db.commit()
            return "⚠️ 找不到簽核人，請聯絡管理員"

        ev = CalendarEvent(
            event_type=EventType.LEAVE,
            title=title or f"{LEAVE_LABEL[leave_code]} {days:g} 天",
            start_at=s, end_at=e, all_day=bool(parsed.get("all_day")),
            user_id=user.id,
            status=EventStatus.PENDING,
            leave_type=leave_code,
            approver_id=approver.id,
            notes=(parsed.get("notes") or "").strip() or None,
            raw_text=text, parsed_json=json.dumps(parsed, ensure_ascii=False),
        )
        db.add(ev)
        await db.commit()
        await db.refresh(ev)
        log.result_event_id = ev.id
        await db.commit()

        # 推播給簽核人
        if approver.line_user_id:
            line_bot.push_to_user(approver.line_user_id, (
                f"📩 待簽核請假\n"
                f"申請人：{user.display_name}\n"
                f"假別：{LEAVE_LABEL[leave_code]}\n"
                f"時段：{s.strftime('%m-%d %H:%M')}–{e.strftime('%m-%d %H:%M')}（{days:g} 天）\n"
                f"請至系統審核：/calendar/approvals"
            ))

        approver_role = "總經理" if approver.role == Role.ADMIN else "單位主管"
        return (
            f"📤 已送出請假申請\n"
            f"假別：{LEAVE_LABEL[leave_code]}\n"
            f"時段：{s.strftime('%m-%d %H:%M')}–{e.strftime('%m-%d %H:%M')}（{days:g} 天）\n"
            f"簽核者：{approver.display_name}（{approver_role}）"
        )

    if intent == "OUTING":
        ev = CalendarEvent(
            event_type=EventType.OUTING,
            title=title or "業務外出",
            start_at=s, end_at=e, all_day=bool(parsed.get("all_day")),
            user_id=user.id,
            status=EventStatus.APPROVED,
            customer_name=(parsed.get("customer") or "").strip() or None,
            notes=(parsed.get("notes") or "").strip() or None,
            raw_text=text, parsed_json=json.dumps(parsed, ensure_ascii=False),
        )
        db.add(ev)
        await db.commit()
        await db.refresh(ev)
        log.result_event_id = ev.id
        await db.commit()
        return (
            f"✅ 已登記外出：{ev.title}\n"
            f"{s.strftime('%m-%d %H:%M')}–{e.strftime('%H:%M')}"
        )

    await db.commit()
    return _help_text()


# ── 子查詢／指令 ──
def _help_text() -> str:
    return (
        "📖 用法範例：\n"
        "• 借大會議室 10:00 到 11:30 開週會\n"
        "• 借小藍 明天 9:00 到 12:00 拜訪客戶\n"
        "• 下週一請特休\n"
        "• 外出拜訪金士頓 14:00-16:00\n"
        "• 我的特休還剩幾天\n"
        "• 今天行事曆\n"
        "• 取消 → 取消最近一筆未完成預約"
    )


async def _query_leave_balance(db: AsyncSession, user: User) -> str:
    from app.models.calendar import LeaveBalance
    year = datetime.now().year
    r = await db.execute(select(LeaveBalance).where(
        LeaveBalance.user_id == user.id, LeaveBalance.year == year
    ))
    rows = list(r.scalars().all())
    if not rows:
        return f"📅 {year} 年特休：尚未設定。請聯絡人事建立額度。"
    out = [f"📅 {year} 年假別餘額（{user.display_name}）"]
    for b in rows:
        remain = b.total_days - b.used_days
        out.append(f"• {LEAVE_LABEL.get(b.leave_code, b.leave_code)}：剩 {remain} / 共 {b.total_days} 天")
    return "\n".join(out)


async def _list_today_events(db: AsyncSession, user: User) -> str:
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    tomorrow = today + timedelta(days=1)
    r = await db.execute(
        select(CalendarEvent)
        .options(selectinload(CalendarEvent.resource))
        .where(
            CalendarEvent.user_id == user.id,
            CalendarEvent.start_at < tomorrow,
            CalendarEvent.end_at > today,
            CalendarEvent.status != EventStatus.CANCELLED,
        )
        .order_by(CalendarEvent.start_at)
    )
    events = list(r.scalars().all())
    if not events:
        return f"📅 今天（{today.strftime('%m-%d')}）沒有行程"
    lines = [f"📅 今天（{today.strftime('%m-%d')}）行程"]
    for e in events:
        res = f"｜{e.resource.name}" if e.resource else ""
        lines.append(f"• {e.start_at.strftime('%H:%M')}–{e.end_at.strftime('%H:%M')} {TYPE_LABEL.get(e.event_type.value, '')}{res}：{e.title}")
    return "\n".join(lines)


async def _list_week_events(db: AsyncSession, user: User) -> str:
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    monday = today - timedelta(days=today.weekday())
    sunday = monday + timedelta(days=7)
    r = await db.execute(
        select(CalendarEvent)
        .options(selectinload(CalendarEvent.resource))
        .where(
            CalendarEvent.user_id == user.id,
            CalendarEvent.start_at < sunday,
            CalendarEvent.end_at > monday,
            CalendarEvent.status != EventStatus.CANCELLED,
        )
        .order_by(CalendarEvent.start_at)
    )
    events = list(r.scalars().all())
    if not events:
        return "📅 本週沒有行程"
    lines = ["📅 本週行程"]
    for e in events:
        res = f"｜{e.resource.name}" if e.resource else ""
        lines.append(f"• {e.start_at.strftime('%m-%d %H:%M')}–{e.end_at.strftime('%H:%M')} {TYPE_LABEL.get(e.event_type.value, '')}{res}：{e.title}")
    return "\n".join(lines)


async def _cancel_latest(db: AsyncSession, user: User) -> str:
    r = await db.execute(
        select(CalendarEvent)
        .where(
            CalendarEvent.user_id == user.id,
            CalendarEvent.status.in_([EventStatus.APPROVED, EventStatus.PENDING]),
            CalendarEvent.start_at >= datetime.now(),
        )
        .order_by(CalendarEvent.created_at.desc())
        .limit(1)
    )
    ev = r.scalars().first()
    if not ev:
        return "找不到可取消的未來預約"
    ev.status = EventStatus.CANCELLED
    await db.commit()
    return f"✅ 已取消：{TYPE_LABEL.get(ev.event_type.value, '')} {ev.title}（{ev.start_at.strftime('%m-%d %H:%M')}）"


async def _list_resources_text(db: AsyncSession, rtype: ResourceType) -> str:
    r = await db.execute(select(CalendarResource).where(
        CalendarResource.type == rtype, CalendarResource.active == True
    ).order_by(CalendarResource.code))
    return "、".join(x.name for x in r.scalars().all()) or "（無）"


# ── HTTP Endpoints ──

@router.post("/webhook")
async def webhook(request: Request, db: AsyncSession = Depends(get_db)):
    """LINE Messaging API webhook 入口"""
    body = await request.body()
    sig = request.headers.get("x-line-signature", "")
    if not line_bot.verify_signature(body, sig):
        raise HTTPException(401, "Invalid signature")
    try:
        payload = json.loads(body.decode("utf-8") or "{}")
    except json.JSONDecodeError:
        raise HTTPException(400, "Bad JSON")

    for ev in payload.get("events", []):
        if ev.get("type") != "message":
            continue
        msg = ev.get("message") or {}
        if msg.get("type") != "text":
            continue
        text = msg.get("text", "")
        line_user_id = (ev.get("source") or {}).get("userId")
        reply_token = ev.get("replyToken")
        user = await _find_user_by_line(db, line_user_id) if line_user_id else None
        try:
            answer = await handle_text_message(db, text, line_user_id, user)
        except Exception as ex:
            logger.exception("handle_text_message failed")
            answer = f"⚠️ 系統錯誤：{ex}"
        if reply_token:
            line_bot.reply_message(reply_token, answer)
        elif line_user_id:
            line_bot.push_to_user(line_user_id, answer)

    return PlainTextResponse("OK")


@router.get("/console")
async def console_page(
    request: Request,
    db: AsyncSession = Depends(get_db),
):
    """LINE 機器人控制台（管理 + dev 測試）"""
    from app.services.auth import get_current_user
    current_user = await get_current_user(request)
    if current_user.role != Role.ADMIN:
        raise HTTPException(403, "僅管理員可進入")
    r = await db.execute(select(User).where(User.is_active == True).order_by(User.role, User.username))
    users = list(r.scalars().all())
    log_r = await db.execute(
        select(LineMessageLog)
        .options(selectinload(LineMessageLog.user) if hasattr(LineMessageLog, "user") else selectinload(LineMessageLog.user_id))
        .order_by(LineMessageLog.id.desc())
        .limit(30)
    ) if False else await db.execute(
        select(LineMessageLog).order_by(LineMessageLog.id.desc()).limit(30)
    )
    logs = list(log_r.scalars().all())
    user_map = {u.id: u for u in users}
    return templates.TemplateResponse("calendar/line_console.html", {
        "request": request, "user": current_user,
        "users": users, "logs": logs, "user_map": user_map,
        "line_configured": line_bot.is_configured(),
    })


@router.post("/admin/bind")
async def admin_bind(
    request: Request,
    db: AsyncSession = Depends(get_db),
    user_id: int = Form(...),
    line_user_id: str = Form(""),
):
    from app.services.auth import get_current_user
    current_user = await get_current_user(request)
    if current_user.role != Role.ADMIN:
        raise HTTPException(403, "僅管理員可進入")
    r = await db.execute(select(User).where(User.id == user_id))
    u = r.scalars().first()
    if not u:
        raise HTTPException(404, "找不到使用者")
    u.line_user_id = (line_user_id or "").strip() or None
    await db.commit()
    from fastapi.responses import RedirectResponse
    return RedirectResponse("/calendar/line/console", status_code=303)


@router.post("/test")
async def line_test(
    request: Request,
    db: AsyncSession = Depends(get_db),
    text: str = Form(...),
    username: Optional[str] = Form(None),
    line_user_id: Optional[str] = Form(None),
):
    """dev 測試入口：模擬 LINE 文字訊息進來。
    可以指定 username 用於模擬該帳號（會 fallback 找該 user 的 line_user_id）。
    """
    user = None
    if username:
        user = await _find_user_by_username(db, username)
        if user and not line_user_id:
            line_user_id = user.line_user_id
    elif line_user_id:
        user = await _find_user_by_line(db, line_user_id)

    answer = await handle_text_message(db, text, line_user_id, user)
    return JSONResponse({
        "user": user.username if user else None,
        "reply": answer,
    })
