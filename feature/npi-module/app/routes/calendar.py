"""
行事曆模組 — P1（會議室／公務車預約 + 衝突檢測）
P2 接 LINE webhook，P3 加請假審核流。
"""
from datetime import datetime, timedelta
from typing import Optional

from fastapi import APIRouter, Depends, Request, Form, HTTPException, Query
from fastapi.responses import HTMLResponse, RedirectResponse, JSONResponse
from fastapi.templating import Jinja2Templates
from sqlalchemy import select, and_, or_
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy.orm import selectinload

from app.database import get_db
from app.models.user import User, Role
from app.models.calendar import (
    CalendarEvent, CalendarResource, EventType, EventStatus, ResourceType,
)
from app.services.auth import get_current_user

router = APIRouter(prefix="/calendar")
templates = Jinja2Templates(directory="app/templates")


# ── 顏色 / 標籤對應（前端 FullCalendar 用）────────
TYPE_COLOR = {
    EventType.LEAVE.value:      "#6c757d",
    EventType.CAR.value:        "#0d6efd",
    EventType.ROOM.value:       "#198754",
    EventType.OUTING.value:     "#fd7e14",
    EventType.ATTENDANCE.value: "#6610f2",
}
TYPE_LABEL = {
    EventType.LEAVE.value:      "請假",
    EventType.CAR.value:        "公務車",
    EventType.ROOM.value:       "會議室",
    EventType.OUTING.value:     "業務外出",
    EventType.ATTENDANCE.value: "出勤",
}
STATUS_LABEL = {
    EventStatus.DRAFT.value:     "草稿",
    EventStatus.PENDING.value:   "待審核",
    EventStatus.APPROVED.value:  "已核准",
    EventStatus.REJECTED.value:  "退回",
    EventStatus.CANCELLED.value: "取消",
}


def _parse_dt(s: str) -> datetime:
    """接受 'YYYY-MM-DDTHH:MM' 或 'YYYY-MM-DD'，回 datetime"""
    if not s:
        raise ValueError("時間不可為空")
    s = s.strip().replace(" ", "T")
    if len(s) == 10:
        return datetime.strptime(s, "%Y-%m-%d")
    if len(s) == 16:
        return datetime.strptime(s, "%Y-%m-%dT%H:%M")
    if len(s) == 19:
        return datetime.strptime(s, "%Y-%m-%dT%H:%M:%S")
    raise ValueError(f"時間格式無法解析：{s}")


async def find_conflicts(
    db: AsyncSession,
    event_type: EventType,
    resource_id: Optional[int],
    start_at: datetime,
    end_at: datetime,
    exclude_event_id: Optional[int] = None,
) -> list[CalendarEvent]:
    """查與該時段重疊且尚未取消／退回的同資源事件"""
    if not resource_id:
        return []
    stmt = (
        select(CalendarEvent)
        .options(selectinload(CalendarEvent.user), selectinload(CalendarEvent.resource))
        .where(
            CalendarEvent.event_type == event_type,
            CalendarEvent.resource_id == resource_id,
            CalendarEvent.status.in_([EventStatus.APPROVED, EventStatus.PENDING]),
            CalendarEvent.start_at < end_at,
            CalendarEvent.end_at > start_at,
        )
    )
    if exclude_event_id:
        stmt = stmt.where(CalendarEvent.id != exclude_event_id)
    r = await db.execute(stmt)
    return list(r.scalars().all())


# ── 月曆首頁 ─────────────────────────────────────
@router.get("/", response_class=HTMLResponse)
async def calendar_index(
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    res_q = await db.execute(
        select(CalendarResource).where(CalendarResource.active == True).order_by(CalendarResource.type, CalendarResource.code)
    )
    resources = list(res_q.scalars().all())
    return templates.TemplateResponse("calendar/index.html", {
        "request": request, "user": current_user,
        "resources": resources,
        "type_color": TYPE_COLOR,
        "type_label": TYPE_LABEL,
    })


# ── FullCalendar 事件 API ───────────────────────
@router.get("/api/events")
async def api_events(
    request: Request,
    start: str = Query(...),
    end: str = Query(...),
    type: Optional[str] = Query(None),
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    try:
        start_dt = _parse_dt(start[:19])
        end_dt   = _parse_dt(end[:19])
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=400)

    stmt = (
        select(CalendarEvent)
        .options(selectinload(CalendarEvent.user), selectinload(CalendarEvent.resource))
        .where(
            CalendarEvent.start_at < end_dt,
            CalendarEvent.end_at > start_dt,
            CalendarEvent.status != EventStatus.CANCELLED,
        )
        .order_by(CalendarEvent.start_at)
    )
    if type:
        try:
            stmt = stmt.where(CalendarEvent.event_type == EventType(type))
        except ValueError:
            pass
    r = await db.execute(stmt)
    events = list(r.scalars().all())

    out = []
    for e in events:
        t = e.event_type.value
        title_parts = [TYPE_LABEL.get(t, t)]
        if e.resource:
            title_parts.append(e.resource.name)
        title_parts.append(e.title or "")
        if e.user:
            title_parts.append(f"({e.user.display_name})")
        title = " · ".join(p for p in title_parts if p)
        out.append({
            "id": e.id,
            "title": title,
            "start": e.start_at.isoformat(),
            "end":   e.end_at.isoformat(),
            "allDay": bool(e.all_day),
            "color": TYPE_COLOR.get(t, "#6c757d"),
            "extendedProps": {
                "event_type": t,
                "type_label": TYPE_LABEL.get(t, t),
                "status": e.status.value,
                "status_label": STATUS_LABEL.get(e.status.value, e.status.value),
                "resource": e.resource.name if e.resource else None,
                "owner": e.user.display_name if e.user else "—",
                "owner_id": e.user_id,
                "notes": e.notes or "",
                "customer": e.customer_name or "",
                "url": f"/calendar/events/{e.id}",
            },
        })
    return out


# ── 新增事件（網頁表單）───────────────────────
@router.get("/events/new", response_class=HTMLResponse)
async def new_event_page(
    request: Request,
    type: Optional[str] = Query(None),
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    res_q = await db.execute(
        select(CalendarResource).where(CalendarResource.active == True).order_by(CalendarResource.type, CalendarResource.code)
    )
    resources = list(res_q.scalars().all())
    return templates.TemplateResponse("calendar/new.html", {
        "request": request, "user": current_user,
        "resources": resources,
        "preset_type": type,
    })


LEAVE_LABEL = {"ANNUAL": "特休", "SICK": "病假", "PERSONAL": "事假"}


def _calc_leave_days(start_at: datetime, end_at: datetime) -> float:
    """計算請假天數（簡化：日曆天，向上取至 0.5 天）"""
    delta = end_at - start_at
    days = delta.total_seconds() / 86400.0
    return round(days * 2) / 2 if days < 1 else round(days, 1)


async def _pick_approver(db: AsyncSession, applicant: User, days: float) -> Optional[User]:
    """≤3 天 → 單位主管（同 BU 的 BU Head）；>3 天 → 總經理（admin）"""
    if days <= 3:
        # 同 BU 的 BU Head
        if applicant.bu:
            r = await db.execute(
                select(User).where(User.role == Role.BU, User.bu == applicant.bu, User.is_active == True)
            )
            u = r.scalars().first()
            if u:
                return u
        # fallback：任一 BU Head
        r = await db.execute(select(User).where(User.role == Role.BU, User.is_active == True))
        u = r.scalars().first()
        if u:
            return u
    # > 3 天 或無 BU Head → 總經理（admin）
    r = await db.execute(select(User).where(User.role == Role.ADMIN, User.is_active == True))
    return r.scalars().first()


@router.post("/events/new")
async def create_event(
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
    event_type: str = Form(...),
    title: str = Form(...),
    start_at: str = Form(...),
    end_at: str = Form(...),
    all_day: Optional[str] = Form(None),
    resource_id: Optional[int] = Form(None),
    leave_type: Optional[str] = Form(None),
    customer_name: Optional[str] = Form(None),
    notes: Optional[str] = Form(None),
):
    is_ajax = "application/json" in (request.headers.get("accept") or "").lower()

    def _err(code: int, msg: str):
        if is_ajax:
            return JSONResponse({"ok": False, "error": msg}, status_code=code)
        raise HTTPException(code, msg)

    try:
        et = EventType(event_type)
    except ValueError:
        return _err(400, "未知的事件類型")

    try:
        s = _parse_dt(start_at)
        e = _parse_dt(end_at)
    except Exception as ex:
        return _err(400, f"時間格式錯誤：{ex}")
    if e <= s:
        return _err(400, "結束時間必須大於開始時間")

    rid = resource_id if resource_id else None
    approver_id = None
    leave_code = None

    if et in (EventType.ROOM, EventType.CAR):
        if not rid:
            return _err(400, f"{TYPE_LABEL[et.value]} 必須選擇資源")
        conflicts = await find_conflicts(db, et, rid, s, e)
        if conflicts:
            c = conflicts[0]
            msg = (
                f"⚠️ 該時段已被借走：「{c.resource.name}」"
                f" {c.start_at.strftime('%Y-%m-%d %H:%M')}–{c.end_at.strftime('%H:%M')} "
                f"by {c.user.display_name if c.user else '?'}"
            )
            return _err(409, msg)

    if et == EventType.LEAVE:
        if not leave_type or leave_type not in LEAVE_LABEL:
            return _err(400, "請選擇假別（特休／病假／事假）")
        leave_code = leave_type
        days = _calc_leave_days(s, e)
        approver = await _pick_approver(db, current_user, days)
        if not approver:
            return _err(400, "找不到簽核人，請聯絡管理員設定")
        approver_id = approver.id
        rid = None
        if title.strip() in ("", LEAVE_LABEL[leave_code]):
            title = f"{LEAVE_LABEL[leave_code]} {days:g} 天"

    status = EventStatus.PENDING if et == EventType.LEAVE else EventStatus.APPROVED

    ev = CalendarEvent(
        event_type=et,
        title=title.strip(),
        start_at=s, end_at=e,
        all_day=bool(all_day),
        user_id=current_user.id,
        resource_id=rid,
        status=status,
        leave_type=leave_code,
        approver_id=approver_id,
        customer_name=(customer_name or "").strip() or None,
        notes=(notes or "").strip() or None,
    )
    db.add(ev)
    await db.commit()
    await db.refresh(ev)
    if is_ajax:
        approver_name = None
        if approver_id:
            ar = await db.execute(select(User).where(User.id == approver_id))
            au = ar.scalars().first()
            approver_name = au.display_name if au else None
        return JSONResponse({
            "ok": True, "id": ev.id,
            "redirect": f"/calendar/events/{ev.id}",
            "status": status.value,
            "approver": approver_name,
        })
    return RedirectResponse(f"/calendar/events/{ev.id}", status_code=303)


# ── 待我簽核 ────────────────────────────────────
@router.get("/approvals", response_class=HTMLResponse)
async def my_approvals(
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    if current_user.role not in (Role.BU, Role.ADMIN):
        raise HTTPException(403, "您不是簽核者")
    stmt = (
        select(CalendarEvent)
        .options(selectinload(CalendarEvent.user), selectinload(CalendarEvent.approver))
        .where(
            CalendarEvent.event_type == EventType.LEAVE,
            CalendarEvent.status == EventStatus.PENDING,
            CalendarEvent.approver_id == current_user.id,
        )
        .order_by(CalendarEvent.created_at.desc())
    )
    r = await db.execute(stmt)
    pendings = list(r.scalars().all())
    # 歷史紀錄
    hist_stmt = (
        select(CalendarEvent)
        .options(selectinload(CalendarEvent.user), selectinload(CalendarEvent.approver))
        .where(
            CalendarEvent.event_type == EventType.LEAVE,
            CalendarEvent.status.in_([EventStatus.APPROVED, EventStatus.REJECTED, EventStatus.CANCELLED]),
            CalendarEvent.approver_id == current_user.id,
        )
        .order_by(CalendarEvent.updated_at.desc())
        .limit(30)
    )
    hr = await db.execute(hist_stmt)
    history = list(hr.scalars().all())
    return templates.TemplateResponse("calendar/approvals.html", {
        "request": request, "user": current_user,
        "pendings": pendings, "history": history,
        "leave_label": LEAVE_LABEL,
        "status_label": STATUS_LABEL,
    })


@router.post("/events/{event_id}/approve")
async def approve_event(
    event_id: int,
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    r = await db.execute(select(CalendarEvent).where(CalendarEvent.id == event_id))
    ev = r.scalars().first()
    if not ev:
        raise HTTPException(404, "事件不存在")
    if ev.status != EventStatus.PENDING:
        raise HTTPException(400, "此事件已不是待審狀態")
    if ev.approver_id != current_user.id and current_user.role != Role.ADMIN:
        raise HTTPException(403, "您不是此事件的簽核人")
    ev.status = EventStatus.APPROVED
    ev.approved_at = datetime.utcnow()
    if not ev.approver_id:
        ev.approver_id = current_user.id
    await db.commit()
    return RedirectResponse(f"/calendar/events/{ev.id}", status_code=303)


@router.post("/events/{event_id}/reject")
async def reject_event(
    event_id: int,
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
    reason: str = Form(""),
):
    r = await db.execute(select(CalendarEvent).where(CalendarEvent.id == event_id))
    ev = r.scalars().first()
    if not ev:
        raise HTTPException(404, "事件不存在")
    if ev.status != EventStatus.PENDING:
        raise HTTPException(400, "此事件已不是待審狀態")
    if ev.approver_id != current_user.id and current_user.role != Role.ADMIN:
        raise HTTPException(403, "您不是此事件的簽核人")
    ev.status = EventStatus.REJECTED
    ev.reject_reason = (reason or "").strip() or None
    if not ev.approver_id:
        ev.approver_id = current_user.id
    await db.commit()
    return RedirectResponse(f"/calendar/events/{ev.id}", status_code=303)


# ── 事件詳細 ─────────────────────────────────────
@router.get("/events/{event_id}", response_class=HTMLResponse)
async def event_detail(
    event_id: int,
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    r = await db.execute(
        select(CalendarEvent)
        .options(
            selectinload(CalendarEvent.user),
            selectinload(CalendarEvent.approver),
            selectinload(CalendarEvent.resource),
        )
        .where(CalendarEvent.id == event_id)
    )
    ev = r.scalars().first()
    if not ev:
        raise HTTPException(404, "事件不存在")
    return templates.TemplateResponse("calendar/detail.html", {
        "request": request, "user": current_user,
        "ev": ev,
        "type_label": TYPE_LABEL,
        "status_label": STATUS_LABEL,
        "type_color": TYPE_COLOR,
    })


@router.post("/events/{event_id}/cancel")
async def cancel_event(
    event_id: int,
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    r = await db.execute(select(CalendarEvent).where(CalendarEvent.id == event_id))
    ev = r.scalars().first()
    if not ev:
        raise HTTPException(404, "事件不存在")
    if ev.user_id != current_user.id and current_user.role != Role.ADMIN:
        raise HTTPException(403, "只有申請人或管理員可取消")
    ev.status = EventStatus.CANCELLED
    await db.commit()
    return RedirectResponse(f"/calendar/events/{ev.id}", status_code=303)
