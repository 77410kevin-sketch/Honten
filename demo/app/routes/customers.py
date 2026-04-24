"""客戶主檔（業務工作區 — 可從 ERP 同步）"""
from fastapi import APIRouter, Depends, Request, Form, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.templating import Jinja2Templates
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select
from datetime import datetime

from app.database import get_db
from app.models.user import User, Role
from app.models.customer import Customer
from app.services.auth import get_current_user
from app.services import erp_client

router = APIRouter(prefix="/customers")
templates = Jinja2Templates(directory="app/templates")

# 業務 / 管理員 能維護客戶主檔
_MANAGE_ROLES = (Role.SALES, Role.ADMIN)


def _can_manage(user: User) -> bool:
    return user.role in _MANAGE_ROLES


@router.get("/", response_class=HTMLResponse)
async def list_customers(
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    r = await db.execute(select(Customer).order_by(Customer.source.desc(), Customer.name))
    items = r.scalars().all()
    return templates.TemplateResponse("customers/list.html", {
        "request": request, "user": current_user,
        "items": items,
        "can_manage": _can_manage(current_user),
        "erp_status": erp_client.erp_status(),
    })


@router.get("/new", response_class=HTMLResponse)
async def new_customer_page(
    request: Request,
    current_user: User = Depends(get_current_user),
):
    if not _can_manage(current_user):
        raise HTTPException(status_code=403)
    return templates.TemplateResponse("customers/new.html", {
        "request": request, "user": current_user,
    })


@router.post("/new")
async def create_customer(
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
    name:    str = Form(...),
    contact: str = Form(""),
    email:   str = Form(""),
    phone:   str = Form(""),
    address: str = Form(""),
    bu:      str = Form(""),
    tax_id:  str = Form(""),
    memo:    str = Form(""),
):
    if not _can_manage(current_user):
        raise HTTPException(status_code=403)
    c = Customer(
        name=name.strip(),
        contact=contact or None,
        email=email or None,
        phone=phone or None,
        address=address or None,
        bu=bu or None,
        tax_id=tax_id or None,
        memo=memo or None,
        source="manual",
    )
    db.add(c)
    await db.commit()
    return RedirectResponse(url="/customers/", status_code=303)


@router.get("/{cust_id}/edit", response_class=HTMLResponse)
async def edit_customer_page(
    cust_id: int, request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    if not _can_manage(current_user):
        raise HTTPException(status_code=403)
    c = await db.get(Customer, cust_id)
    if not c:
        raise HTTPException(status_code=404)
    return templates.TemplateResponse("customers/edit.html", {
        "request": request, "user": current_user, "item": c,
    })


@router.post("/{cust_id}/edit")
async def update_customer(
    cust_id: int,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
    name:    str = Form(...),
    contact: str = Form(""),
    email:   str = Form(""),
    phone:   str = Form(""),
    address: str = Form(""),
    bu:      str = Form(""),
    tax_id:  str = Form(""),
    memo:    str = Form(""),
    is_active: str = Form("on"),
):
    if not _can_manage(current_user):
        raise HTTPException(status_code=403)
    c = await db.get(Customer, cust_id)
    if not c:
        raise HTTPException(status_code=404)
    c.name    = name.strip()
    c.contact = contact or None
    c.email   = email or None
    c.phone   = phone or None
    c.address = address or None
    c.bu      = bu or None
    c.tax_id  = tax_id or None
    c.memo    = memo or None
    c.is_active = (is_active == "on" or is_active == "true")
    await db.commit()
    return RedirectResponse(url="/customers/", status_code=303)


@router.post("/_sync-erp")
async def sync_from_erp(
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    """從 ERP 同步客戶主檔（依 erp_code 去重，新增或更新）。
    目前走 Stub backend；ERP 上線後切換 `.env` 的 `ERP_BACKEND=sqlserver` 即可。
    """
    if not _can_manage(current_user):
        raise HTTPException(status_code=403)
    st = erp_client.erp_status()
    if not st["connected"]:
        raise HTTPException(status_code=503, detail="ERP 尚未連線，請在 .env 設定 ERP_* 變數")

    rows = erp_client.fetch_customers_from_erp()
    added, updated = 0, 0
    for row in rows:
        r = await db.execute(select(Customer).where(Customer.erp_code == row.erp_code))
        c = r.scalars().first()
        if c:
            c.name = row.name
            c.contact = row.contact
            c.email = row.email
            c.phone = row.phone
            c.address = row.address
            c.bu = row.bu
            c.tax_id = row.tax_id
            c.is_active = row.is_active
            c.source = "erp"
            c.updated_at = datetime.utcnow()
            updated += 1
        else:
            db.add(Customer(
                erp_code=row.erp_code, name=row.name,
                contact=row.contact, email=row.email, phone=row.phone,
                address=row.address, bu=row.bu, tax_id=row.tax_id,
                is_active=row.is_active, source="erp",
            ))
            added += 1
    await db.commit()
    url = f"/customers/?sync=ok&added={added}&updated={updated}"
    return RedirectResponse(url=url, status_code=303)
