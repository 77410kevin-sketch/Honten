"""供應商主檔後台管理（供 NPI 模組派發詢價使用）"""
from fastapi import APIRouter, Depends, Request, Form, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.templating import Jinja2Templates
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select

from datetime import datetime

from app.database import get_db
from app.models.user import User, Role
from app.models.supplier import Supplier, SupplierType
from app.services.auth import get_current_user
from app.services import erp_client

router    = APIRouter(prefix="/suppliers")
templates = Jinja2Templates(directory="app/templates")

# 只有工程/工程主管/管理員可以維護供應商主檔
_MANAGE_ROLES = (Role.ENGINEER, Role.ENG_MGR, Role.ADMIN)


def _can_manage(user: User) -> bool:
    return user.role in _MANAGE_ROLES


@router.get("/", response_class=HTMLResponse)
async def list_suppliers(
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    r = await db.execute(select(Supplier).order_by(Supplier.type, Supplier.name))
    items = r.scalars().all()
    return templates.TemplateResponse("suppliers/list.html", {
        "request": request, "user": current_user,
        "items": items, "can_manage": _can_manage(current_user),
        "SupplierType": SupplierType,
        "erp_status": erp_client.erp_status(),
    })


@router.post("/_sync-erp")
async def sync_suppliers_from_erp(
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    """從 ERP 同步供應商主檔（依名稱去重）。Stub 模式。"""
    if not _can_manage(current_user):
        raise HTTPException(status_code=403)
    st = erp_client.erp_status()
    if not st["connected"]:
        raise HTTPException(status_code=503, detail="ERP 尚未連線")
    rows = erp_client.fetch_suppliers_from_erp()
    added, updated = 0, 0
    for row in rows:
        r = await db.execute(select(Supplier).where(Supplier.name == row.name))
        sup = r.scalars().first()
        sup_type = SupplierType.INTERNAL if row.type == "廠內" else SupplierType.EXTERNAL
        if sup:
            sup.contact = row.contact
            sup.email = row.email
            sup.phone = row.phone
            sup.type = sup_type
            sup.is_active = row.is_active
            updated += 1
        else:
            db.add(Supplier(
                name=row.name, type=sup_type,
                contact=row.contact, email=row.email, phone=row.phone,
                is_active=row.is_active,
            ))
            added += 1
    await db.commit()
    return RedirectResponse(url=f"/suppliers/?sync=ok&added={added}&updated={updated}", status_code=303)


@router.get("/new", response_class=HTMLResponse)
async def new_supplier_page(
    request: Request,
    current_user: User = Depends(get_current_user),
):
    if not _can_manage(current_user):
        raise HTTPException(status_code=403, detail="無權限")
    return templates.TemplateResponse("suppliers/new.html", {
        "request": request, "user": current_user,
        "SupplierType": SupplierType,
    })


@router.post("/new")
async def create_supplier(
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
    name:    str = Form(...),
    type:    str = Form("外部"),
    contact: str = Form(""),
    email:   str = Form(""),
    phone:   str = Form(""),
    memo:    str = Form(""),
):
    if not _can_manage(current_user):
        raise HTTPException(status_code=403)
    sup = Supplier(
        name=name.strip(),
        type=SupplierType(type) if type in (t.value for t in SupplierType) else SupplierType.EXTERNAL,
        contact=contact or None,
        email=email or None,
        phone=phone or None,
        memo=memo or None,
    )
    db.add(sup)
    await db.commit()
    return RedirectResponse(url="/suppliers/", status_code=303)


@router.get("/{sup_id}/edit", response_class=HTMLResponse)
async def edit_supplier_page(
    sup_id: int, request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    if not _can_manage(current_user):
        raise HTTPException(status_code=403)
    sup = await db.get(Supplier, sup_id)
    if not sup:
        raise HTTPException(status_code=404)
    return templates.TemplateResponse("suppliers/edit.html", {
        "request": request, "user": current_user,
        "item": sup, "SupplierType": SupplierType,
    })


@router.post("/{sup_id}/edit")
async def update_supplier(
    sup_id: int,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
    name:    str = Form(...),
    type:    str = Form("外部"),
    contact: str = Form(""),
    email:   str = Form(""),
    phone:   str = Form(""),
    memo:    str = Form(""),
    is_active: str = Form("on"),
):
    if not _can_manage(current_user):
        raise HTTPException(status_code=403)
    sup = await db.get(Supplier, sup_id)
    if not sup:
        raise HTTPException(status_code=404)
    sup.name    = name.strip()
    sup.type    = SupplierType(type) if type in (t.value for t in SupplierType) else sup.type
    sup.contact = contact or None
    sup.email   = email or None
    sup.phone   = phone or None
    sup.memo    = memo or None
    sup.is_active = (is_active == "on" or is_active == "true")
    await db.commit()
    return RedirectResponse(url="/suppliers/", status_code=303)


@router.post("/{sup_id}/toggle")
async def toggle_supplier(
    sup_id: int,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    if not _can_manage(current_user):
        raise HTTPException(status_code=403)
    sup = await db.get(Supplier, sup_id)
    if not sup:
        raise HTTPException(status_code=404)
    sup.is_active = not sup.is_active
    await db.commit()
    return RedirectResponse(url="/suppliers/", status_code=303)
