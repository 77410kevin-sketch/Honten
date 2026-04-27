"""QC 異常管理系統（NCR）路由

流程（草案，後續再迭代）：
  DRAFT (品保填寫 IPC)
    → PENDING_DISPOSITION (品保下處理判斷：退貨/實驗/特採)
    → PENDING_RCA (Mail 通知 + 根因分析)
    → PENDING_IMPROVEMENT (制定長期改善方案 — 圖面/SOP/SIP)
    → LINKED_ECN (若需修訂圖面/SOP/SIP，開 ECN 連結進去)
    → CLOSED
"""
import os, uuid, json, logging, mimetypes
from datetime import datetime
from typing import List

from fastapi import APIRouter, Depends, Request, Form, UploadFile, File, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse, FileResponse
from fastapi.templating import Jinja2Templates
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select, or_
from sqlalchemy.orm import selectinload

from app.database import get_db
from app.models.user import User, Role
from app.models.qc_exception import (
    QCException, QCExceptionDocument, QCExceptionApproval,
    QCExceptionStatus, QCDisposition, QCExceptionStage,
    QCDocType, QCEventDateType, QCSourceType,
)
from app.services.auth import get_current_user
from app.services import qc_notification as qc_notif

router    = APIRouter(prefix="/qc-exceptions")
templates = Jinja2Templates(directory="app/templates")


def _fromjson_filter(s):
    if not s:
        return None
    try:
        return json.loads(s)
    except Exception:
        return None
templates.env.filters["fromjson"] = _fromjson_filter

UPLOAD_BASE = "uploads"
ATTACH_CATEGORIES = ["異常照片", "實驗報告", "圖面", "Rework SOP", "其它"]

_QC_ROLES    = (Role.QC, Role.ADMIN)                                     # 處理判斷 / RCA / 改善方案 專屬
_CREATE_ROLES = (Role.QC, Role.PROD_MGR, Role.PC, Role.ASSISTANT, Role.ADMIN)  # 建單權限：品保 + 產線主管 + 生管 + 業助
_VIEW_ROLES  = (Role.QC, Role.ENGINEER, Role.ENG_MGR, Role.PURCHASE,
                Role.PROD_MGR, Role.PC, Role.ASSISTANT, Role.WAREHOUSE,
                Role.BU, Role.ADMIN)


# ── 共用 helper ─────────────────────────────────

def _upload_dir(form_pk: int) -> str:
    p = os.path.join(UPLOAD_BASE, f"qc_{form_pk}")
    os.makedirs(p, exist_ok=True)
    return p


async def _next_form_id(db: AsyncSession) -> str:
    today = datetime.now().strftime("%Y%m%d")
    prefix = f"NCR-{today}-"
    r = await db.execute(select(QCException).where(QCException.form_id.like(f"{prefix}%")))
    rows = list(r.scalars().all())
    seq = len(rows) + 1
    return f"{prefix}{seq:03d}"


async def _get_or_404(form_id: str, db: AsyncSession) -> QCException:
    r = await db.execute(
        select(QCException)
        .options(
            selectinload(QCException.creator),
            selectinload(QCException.assigned_qc),
            selectinload(QCException.dispositioner),
            selectinload(QCException.linked_ecn),
            selectinload(QCException.documents),
            selectinload(QCException.approvals).selectinload(QCExceptionApproval.approver),
        )
        .where(QCException.form_id == form_id)
    )
    f = r.scalars().first()
    if not f:
        raise HTTPException(status_code=404, detail="QC 異常單不存在")
    return f


def _log(form: QCException, user: User, action: str,
         from_s: QCExceptionStatus | None, to_s: QCExceptionStatus | None,
         comment: str = "", reject_target: str | None = None):
    return QCExceptionApproval(
        form_id_fk=form.id, approver_id=user.id, action=action,
        comment=comment or None, reject_target=reject_target,
        from_status=from_s.value if from_s else None,
        to_status=to_s.value if to_s else None,
    )


def _docs_by_cat(docs):
    out = {}
    for d in (docs or []):
        out.setdefault(d.category or "其它", []).append(d)
    return out


async def _save_attachments(db, form_pk, user_id, files, categories):
    if not files:
        return
    upload_dir = _upload_dir(form_pk)
    for i, uf in enumerate(files):
        if not uf or not uf.filename:
            continue
        content = await uf.read()
        if not content:
            continue
        ext = os.path.splitext(uf.filename)[1] or ".bin"
        saved = f"{uuid.uuid4().hex}{ext}"
        with open(os.path.join(upload_dir, saved), "wb") as f:
            f.write(content)
        cat = (categories[i] if i < len(categories) else "其它")
        if cat not in ATTACH_CATEGORIES:
            cat = "其它"
        db.add(QCExceptionDocument(
            form_id_fk=form_pk, filename=saved, original_name=uf.filename,
            category=cat, uploaded_by=user_id,
        ))


# ── 列表 ────────────────────────────────────────

@router.get("/", response_class=HTMLResponse)
async def list_qc(
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    if current_user.role not in _VIEW_ROLES:
        raise HTTPException(status_code=403, detail="您的角色無權限存取 QC 異常管理")
    q = (select(QCException)
         .options(selectinload(QCException.creator), selectinload(QCException.linked_ecn))
         .order_by(QCException.created_at.desc()))
    r = await db.execute(q)
    forms = list(r.scalars().all())
    return templates.TemplateResponse("qc_exceptions/list.html", {
        "request": request, "user": current_user,
        "forms": forms,
        "QCExceptionStatus": QCExceptionStatus,
        "QCDisposition": QCDisposition,
    })


# ── 新建（顯示 IPC 異常資訊表單） ────────────────

@router.get("/new", response_class=HTMLResponse)
async def new_qc_page(
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    if current_user.role not in _CREATE_ROLES:
        raise HTTPException(status_code=403, detail="僅品保 / 產線主管 / 業助可新建 QC 異常單")
    return templates.TemplateResponse("qc_exceptions/new.html", {
        "request": request, "user": current_user,
        "QCExceptionStage": QCExceptionStage,
        "QCDocType": QCDocType,
        "QCEventDateType": QCEventDateType,
        "QCSourceType": QCSourceType,
        "ATTACH_CATEGORIES": ATTACH_CATEGORIES,
    })


def _parse_defect_items(fd):
    """從 form data 解析多列異常項目（抽樣→不良→自動算總不良率）"""
    causes = fd.getlist("defect_cause")
    types_csvs = fd.getlist("defect_types_csv")
    sqs = fd.getlist("sample_qty")
    dqs = fd.getlist("defect_qty")
    def _int(s):
        try: return int(s)
        except (TypeError, ValueError): return None
    items, total_dq, total_sq = [], 0, 0
    for i, c in enumerate(causes):
        c = (c or "").strip()
        if not c:
            continue
        types_str = types_csvs[i] if i < len(types_csvs) else ""
        types = [t.strip() for t in (types_str or "").split(",") if t.strip()]
        sq = _int(sqs[i] if i < len(sqs) else None)
        dq = _int(dqs[i] if i < len(dqs) else None)
        items.append({"cause": c, "types": types, "sample_qty": sq, "defect_qty": dq})
        if dq is not None: total_dq += dq
        if sq is not None: total_sq += sq
    rate = (total_dq / total_sq) if total_sq > 0 else None
    return items, total_dq, total_sq, rate


@router.post("/new")
async def create_qc(
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    if current_user.role not in _CREATE_ROLES:
        raise HTTPException(status_code=403)
    fd = await request.form()
    def g(k, d=""):
        v = fd.get(k); return v.strip() if isinstance(v, str) else d
    def _int(s):
        try: return int(s)
        except (TypeError, ValueError): return None

    part_no = g("part_no")
    if not part_no:
        raise HTTPException(status_code=400, detail="品號為必填")
    items, total_dq, total_sq, rate = _parse_defect_items(fd)
    if not items:
        raise HTTPException(status_code=400, detail="至少需填一筆異常原因")

    try: st_enum = QCExceptionStage(g("stage", "IQC"))
    except ValueError: st_enum = QCExceptionStage.IQC
    try: dt_enum = QCDocType(g("doc_type", "RECEIVE"))
    except ValueError: dt_enum = QCDocType.RECEIVE
    try: edt_enum = QCEventDateType(g("event_date_type", "RECEIVE"))
    except ValueError: edt_enum = QCEventDateType.RECEIVE
    try: src_enum = QCSourceType(g("source_type", "SUPPLIER"))
    except ValueError: src_enum = QCSourceType.SUPPLIER

    submit_action = g("submit_action", "draft")
    form_id = await _next_form_id(db)
    initial_status = (QCExceptionStatus.PENDING_DISPOSITION
                      if submit_action == "submit"
                      else QCExceptionStatus.DRAFT)
    qc = QCException(
        form_id=form_id, status=initial_status,
        part_no=part_no,
        doc_type=dt_enum, receive_doc_no=g("receive_doc_no") or None,
        event_date_type=edt_enum, receive_date=g("receive_date") or None,
        stage=st_enum, source_type=src_enum,
        supplier_name=g("supplier_name") or None,
        receive_qty=_int(fd.get("receive_qty")),
        defect_cause=items[0]["cause"],
        measurement_data=None,
        defect_qty=total_dq or None, sample_qty=total_sq or None, defect_rate=rate,
        defect_items_json=json.dumps(items, ensure_ascii=False),
        created_by=current_user.id,
        assigned_qc_id=(current_user.id if current_user.role in _QC_ROLES else None),
    )
    db.add(qc)
    await db.commit()
    await db.refresh(qc)

    attach_files = [f for f in fd.getlist("attach_files")
                    if hasattr(f, "filename") and f.filename]
    attach_categories = fd.getlist("attach_categories")
    if attach_files:
        await _save_attachments(db, qc.id, current_user.id, attach_files, attach_categories)
        await db.commit()
    db.add(_log(qc, current_user,
                "SUBMIT" if initial_status != QCExceptionStatus.DRAFT else "CREATE",
                None, initial_status, "建立 QC 異常單"))
    await db.commit()
    # 送出（非草稿）才通知 LINE 群組 + 相關角色，避免草稿就吵到大家
    if initial_status != QCExceptionStatus.DRAFT:
        try:
            await qc_notif.notify_exception_created(
                db, qc, creator_name=(current_user.display_name or current_user.username))
        except Exception:
            logging.exception("notify_exception_created failed")
    return RedirectResponse(url=f"/qc-exceptions/{form_id}", status_code=303)


# ── 編輯（建單者於 DRAFT 狀態可修改 IPC 異常資訊） ──

@router.get("/{form_id}/edit", response_class=HTMLResponse)
async def edit_qc_page(
    form_id: str, request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    form = await _get_or_404(form_id, db)
    if form.status != QCExceptionStatus.DRAFT:
        raise HTTPException(status_code=400, detail="僅 DRAFT 狀態可編輯")
    is_creator = (form.created_by == current_user.id)
    if not (is_creator or current_user.role == Role.ADMIN):
        raise HTTPException(status_code=403, detail="僅建單者或 admin 可編輯")
    return templates.TemplateResponse("qc_exceptions/edit.html", {
        "request": request, "user": current_user, "form": form,
        "QCExceptionStage": QCExceptionStage,
        "QCDocType": QCDocType,
        "QCEventDateType": QCEventDateType,
        "QCSourceType": QCSourceType,
        "ATTACH_CATEGORIES": ATTACH_CATEGORIES,
        "docs_by_cat": _docs_by_cat(form.documents),
    })


@router.post("/{form_id}/edit")
async def update_qc(
    form_id: str,
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    form = await _get_or_404(form_id, db)
    if form.status != QCExceptionStatus.DRAFT:
        raise HTTPException(status_code=400, detail="僅 DRAFT 狀態可編輯")
    is_creator = (form.created_by == current_user.id)
    if not (is_creator or current_user.role == Role.ADMIN):
        raise HTTPException(status_code=403)

    fd = await request.form()
    def g(k, d=""):
        v = fd.get(k); return v.strip() if isinstance(v, str) else d
    def _int(s):
        try: return int(s)
        except (TypeError, ValueError): return None

    part_no = g("part_no")
    if part_no: form.part_no = part_no

    items, total_dq, total_sq, rate = _parse_defect_items(fd)
    if items:
        form.defect_cause = items[0]["cause"]
        form.defect_qty = total_dq or None
        form.sample_qty = total_sq or None
        form.defect_rate = rate
        form.defect_items_json = json.dumps(items, ensure_ascii=False)
    form.measurement_data = None  # 已棄用

    form.receive_doc_no = g("receive_doc_no") or None
    form.receive_date   = g("receive_date") or None
    form.supplier_name  = g("supplier_name") or None
    form.receive_qty    = _int(fd.get("receive_qty"))

    try: form.stage = QCExceptionStage(g("stage", "IQC"))
    except ValueError: pass
    try: form.doc_type = QCDocType(g("doc_type", "RECEIVE"))
    except ValueError: pass
    try: form.event_date_type = QCEventDateType(g("event_date_type", "RECEIVE"))
    except ValueError: pass
    try: form.source_type = QCSourceType(g("source_type", "SUPPLIER"))
    except ValueError: pass

    attach_files = [f for f in fd.getlist("attach_files")
                    if hasattr(f, "filename") and f.filename]
    attach_categories = fd.getlist("attach_categories")
    submit_action = g("submit_action", "save")
    if attach_files:
        await _save_attachments(db, form.id, current_user.id, attach_files, attach_categories)

    if submit_action == "resubmit":
        old = form.status
        form.status = QCExceptionStatus.PENDING_DISPOSITION
        db.add(_log(form, current_user, "RESUBMIT", old, form.status, "補件後重送品保判斷"))
        form.updated_at = datetime.utcnow()
        await db.commit()
        # 再送一次 LINE 群組通知
        try:
            await qc_notif.notify_exception_created(
                db, form, creator_name=(current_user.display_name or current_user.username))
        except Exception:
            logging.exception("resubmit notify failed")
    else:
        form.updated_at = datetime.utcnow()
        await db.commit()
    return RedirectResponse(url=f"/qc-exceptions/{form_id}", status_code=303)


@router.post("/{form_id}/delete-doc/{doc_id}")
async def delete_doc(
    form_id: str, doc_id: int,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    form = await _get_or_404(form_id, db)
    if form.status != QCExceptionStatus.DRAFT:
        raise HTTPException(status_code=400, detail="僅 DRAFT 狀態可刪附件")
    if not (form.created_by == current_user.id or current_user.role == Role.ADMIN):
        raise HTTPException(status_code=403)
    doc = await db.get(QCExceptionDocument, doc_id)
    if not doc or doc.form_id_fk != form.id:
        raise HTTPException(status_code=404)
    fp = os.path.join(UPLOAD_BASE, f"qc_{form.id}", doc.filename)
    if os.path.exists(fp):
        try: os.remove(fp)
        except Exception: pass
    await db.delete(doc)
    await db.commit()
    return RedirectResponse(url=f"/qc-exceptions/{form_id}/edit", status_code=303)


# ── 詳情 ────────────────────────────────────────

@router.get("/{form_id}", response_class=HTMLResponse)
async def detail_qc(
    form_id: str, request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    if current_user.role not in _VIEW_ROLES:
        raise HTTPException(status_code=403)
    form = await _get_or_404(form_id, db)
    transition_combo = {
        "DRAFT→PENDING_DISPOSITION":               ("品保送審",       "primary"),
        "PENDING_DISPOSITION→PENDING_IMPROVEMENT": ("品保下處理判斷 → 改善方案", "info"),
        "PENDING_DISPOSITION→PENDING_RCA":         ("品保下處理判斷（舊）", "info"),
        "PENDING_RCA→PENDING_IMPROVEMENT":         ("併入改善方案（舊）",   "info"),
        "PENDING_IMPROVEMENT→LINKED_ECN":          ("綁入 ECN",       "warning"),
        "PENDING_IMPROVEMENT→CLOSED":              ("結案",           "dark"),
        "LINKED_ECN→CLOSED":                       ("ECN 已結案 → 結案", "dark"),
    }
    # 自動 lookup 廠商主檔的 contact / email，server-side 帶入模板
    # 廠內單位收件清單（hard-coded）+ 廠商主檔清單，供 mail 收件人下拉選擇
    INTERNAL_STATIONS = [
        ("品保 QC",       "qa@honten.local"),
        ("IQC 進料檢驗",  "iqc@honten.local"),
        ("IPQC 製程檢驗", "ipqc@honten.local"),
        ("OQC 出貨檢驗",  "oqc@honten.local"),
        ("品檢",          "inspect@honten.local"),
        ("雷雕課",        "laser@honten.local"),
        ("CNC 課",        "cnc@honten.local"),
        ("組裝課",        "asm@honten.local"),
        ("生管 PMC",      "pmc@honten.local"),
        ("採購",          "purchase@honten.local"),
        ("業助",          "assist@honten.local"),
    ]
    STAGE_TO_EMAIL = {
        "IQC": "iqc@honten.local", "IPQC": "ipqc@honten.local", "OQC": "oqc@honten.local",
        "INSPECTION": "inspect@honten.local", "LASER": "laser@honten.local",
        "CNC": "cnc@honten.local", "ASSEMBLY": "asm@honten.local",
    }
    from app.models.supplier import Supplier
    rs2 = await db.execute(
        select(Supplier).where(Supplier.is_active == True).order_by(Supplier.name)
    )
    suppliers = [{"id": s.id, "name": s.name, "email": s.email or "",
                  "contact": s.contact or ""} for s in rs2.scalars().all() if s.email]

    # 找對應該單異常廠商的 email/contact
    sup_contact, sup_email = "", ""
    if form.supplier_name:
        for s in suppliers:
            if form.supplier_name.strip() in s["name"]:
                sup_contact = s["contact"]; sup_email = s["email"]; break

    # 預選 email 邏輯：
    #   SUPPLIER → 帶該廠商
    #   INTERNAL → 帶對應工段
    #   CUSTOMER（客訴）→ 不預選，由品保自選
    src = form.source_type.value if form.source_type else "SUPPLIER"
    preselect_email = ""
    if src == "SUPPLIER":
        preselect_email = sup_email
    elif src == "INTERNAL":
        preselect_email = STAGE_TO_EMAIL.get(form.stage.value if form.stage else "", "")

    return templates.TemplateResponse("qc_exceptions/detail.html", {
        "request": request, "user": current_user, "form": form,
        "docs_by_cat": _docs_by_cat(form.documents),
        "transition_combo": transition_combo,
        "QCExceptionStatus": QCExceptionStatus,
        "QCDisposition": QCDisposition,
        "QCExceptionStage": QCExceptionStage,
        "ATTACH_CATEGORIES": ATTACH_CATEGORIES,
        "qc_supplier_mail_tpl": qc_notif.build_supplier_mail_template(form, sup_contact),
        "qc_internal_stations": INTERNAL_STATIONS,
        "qc_active_suppliers":  suppliers,
        "qc_preselect_email":   preselect_email,
    })


# ── 供應商主檔查詢（建單/處理判斷時 auto-fill 用）───
@router.get("/api/supplier-lookup")
async def supplier_lookup(
    name: str,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    """依名稱模糊查供應商主檔，回傳 contact / email / phone 給前端 auto-fill"""
    if current_user.role not in _VIEW_ROLES:
        raise HTTPException(status_code=403)
    n = (name or "").strip()
    if not n:
        return {"matches": []}
    from app.models.supplier import Supplier
    r = await db.execute(
        select(Supplier).where(
            Supplier.is_active == True,
            Supplier.name.like(f"%{n}%"),
        ).order_by(Supplier.name).limit(10)
    )
    return {"matches": [
        {"id": s.id, "name": s.name, "contact": s.contact or "",
         "email": s.email or "", "phone": s.phone or ""}
        for s in r.scalars().all()
    ]}


# ── 附件預覽 ────────────────────────────────────

@router.get("/doc/preview/{doc_id}")
async def preview_doc(
    doc_id: int,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    if current_user.role not in _VIEW_ROLES:
        raise HTTPException(status_code=403)
    doc = await db.get(QCExceptionDocument, doc_id)
    if not doc:
        raise HTTPException(status_code=404)
    fp = os.path.join(UPLOAD_BASE, f"qc_{doc.form_id_fk}", doc.filename)
    if not os.path.exists(fp):
        raise HTTPException(status_code=404, detail="檔案不存在")
    mime, _ = mimetypes.guess_type(fp)
    return FileResponse(fp, media_type=(mime or "application/octet-stream"),
                        filename=doc.original_name,
                        headers={"Content-Disposition": f'inline; filename="{doc.filename}"'})


# ── 狀態流轉 ────────────────────────────────────

@router.post("/{form_id}/return-to-previous")
async def return_to_previous(
    form_id: str,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
    comment: str = Form(""),
):
    """品保／後續站，退回上一站要求補資料"""
    form = await _get_or_404(form_id, db)
    if current_user.role not in _QC_ROLES:
        raise HTTPException(status_code=403)
    # 反向狀態映射：把目前狀態退回前一站（PENDING_RCA 已併入 IMPROVEMENT）
    back_map = {
        QCExceptionStatus.PENDING_DISPOSITION: QCExceptionStatus.DRAFT,
        QCExceptionStatus.PENDING_IMPROVEMENT: QCExceptionStatus.PENDING_DISPOSITION,
        QCExceptionStatus.LINKED_ECN:          QCExceptionStatus.PENDING_IMPROVEMENT,
    }
    new_st = back_map.get(form.status)
    if not new_st:
        raise HTTPException(status_code=400, detail="目前狀態無法退回前一站")
    if not comment.strip():
        raise HTTPException(status_code=400, detail="退回原因為必填")
    old = form.status
    form.status = new_st
    form.updated_at = datetime.utcnow()
    db.add(_log(form, current_user, "RETURN_PREV", old, new_st,
                f"退回前一站：{comment.strip()[:200]}"))
    await db.commit()
    return RedirectResponse(url=f"/qc-exceptions/{form_id}", status_code=303)


@router.post("/{form_id}/disposition")
async def set_disposition(
    form_id: str,
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
    note: str = Form(""),
):
    """品保下處理判斷（多選） → 進 PENDING_RCA

    接收欄位：
      - dispositions[]：勾選的處理方式（可多選）
      - note：總體說明
      - 退回供應商：supplier_mail_to/cc/subject/body（可空，填了會存以便後續寄）
      - 實驗測試：lab_test_qty / lab_test_conditions / lab_test_due_date
      - 特採允收：sa_need_sorting / sa_need_rework / sa_station / sa_defect_handling
    """
    form = await _get_or_404(form_id, db)
    if current_user.role not in _QC_ROLES:
        raise HTTPException(status_code=403)
    if form.status not in (QCExceptionStatus.PENDING_DISPOSITION, QCExceptionStatus.DRAFT):
        raise HTTPException(status_code=400, detail="目前狀態無法下處理判斷")

    fd = await request.form()
    picked = [x for x in fd.getlist("dispositions") if x in [d.value for d in QCDisposition]]
    if not picked:
        raise HTTPException(status_code=400, detail="請至少勾選一個處理方式")

    def _int(s):
        try: return int(s)
        except (TypeError, ValueError): return None

    old = form.status
    form.dispositions_json = json.dumps(picked, ensure_ascii=False)
    # 主要 disposition 取第一個（向下相容 list view / mail 顯示）
    try:
        form.disposition = QCDisposition(picked[0])
    except ValueError:
        form.disposition = None
    form.disposition_note = note.strip() or None
    form.disposition_at = datetime.utcnow()
    form.disposition_by = current_user.id

    # 立即處理 — 通知信（對象可為「供應商」或「工站」）
    target = (fd.get("rts_target_type") or "").strip().upper()
    if target in ("SUPPLIER", "STATION"):
        form.rts_target_type = target
    form.supplier_mail_to      = (fd.get("supplier_mail_to") or "").strip() or None
    form.supplier_mail_cc      = (fd.get("supplier_mail_cc") or "").strip() or None
    form.supplier_mail_subject = (fd.get("supplier_mail_subject") or "").strip() or None
    form.supplier_mail_body    = (fd.get("supplier_mail_body") or "").strip() or None

    # A. 退貨 — 補貨需求說明（給採購/生管）+ 司機載回
    if "RETURN_TO_SUPPLIER" in picked:
        form.rts_replenish_note = (fd.get("rts_replenish_note") or "").strip() or None
        form.rts_pickup_required = bool(fd.get("rts_pickup_required"))
        form.rts_pickup_note = (fd.get("rts_pickup_note") or "").strip() or None

    # B. 處理方式 — 子類別多選（NO_ACTION 與其他互斥；廠內/客戶端可同時）
    if "SPECIAL_ACCEPT" in picked:
        subs = [s for s in fd.getlist("sa_subtypes")
                if s in ("NO_ACTION", "SORTING", "REWORK", "CUST_SORTING", "CUST_REWORK")]
        if not subs:
            subs = ["NO_ACTION"]
        # NO_ACTION 與其他互斥
        if "NO_ACTION" in subs and len(subs) > 1:
            subs = [s for s in subs if s != "NO_ACTION"]
        form.sa_subtypes_json = json.dumps(subs, ensure_ascii=False)
        form.sa_subtype = subs[0]  # 主類別（向下相容）
        form.sa_need_sorting = ("SORTING" in subs)
        form.sa_need_rework  = ("REWORK" in subs)
        # 由品保填寫
        form.sa_defect_handling = (fd.get("sa_defect_handling") or "").strip() or None
        # SORTING — 數量留空待 sorting 單位回填，但建立時也接受品保預填
        if "SORTING" in subs:
            form.sa_sorting_pass_qty = _int(fd.get("sa_sorting_pass_qty"))
            form.sa_sorting_fail_qty = _int(fd.get("sa_sorting_fail_qty"))
        # REWORK — 小批驗證內容 + 樣品測試需求
        if "REWORK" in subs:
            form.sa_rework_note      = (fd.get("sa_rework_note") or "").strip() or None
            form.lab_test_qty        = _int(fd.get("lab_test_qty"))
            form.lab_test_conditions = (fd.get("lab_test_conditions") or "").strip() or None
            form.lab_test_due_date   = (fd.get("lab_test_due_date") or "").strip() or None
        # B4 客戶端 Sorting / B5 客戶端 Rework — 工時與人力
        def _float(s):
            try: return float(s)
            except (TypeError, ValueError): return None
        if "CUST_SORTING" in subs:
            form.sa_cust_sorting_hours   = _float(fd.get("sa_cust_sorting_hours"))
            form.sa_cust_sorting_workers = _int(fd.get("sa_cust_sorting_workers"))
        if "CUST_REWORK" in subs:
            form.sa_cust_rework_hours   = _float(fd.get("sa_cust_rework_hours"))
            form.sa_cust_rework_workers = _int(fd.get("sa_cust_rework_workers"))
        if "CUST_SORTING" in subs or "CUST_REWORK" in subs:
            form.sa_cust_note = (fd.get("sa_cust_note") or "").strip() or None
        # Rework SOP 附件（B3 / B5 用）— 品保上傳，自動歸類「Rework SOP」
        sop_files = []
        for f in fd.getlist("rework_sop_files"):
            if hasattr(f, "filename") and f.filename:
                sop_files.append(f)
        if sop_files:
            await _save_attachments(db, form.id, current_user.id,
                                    sop_files, ["Rework SOP"] * len(sop_files))

    # C. 橫向展開 — 多列盤點單（JSON list）
    if "HORIZONTAL_EXPANSION" in picked:
        rows = []
        part_nos = fd.getlist("inv_part_no")
        cust_qs  = fd.getlist("inv_customer_qty")
        in_qs    = fd.getlist("inv_inhouse_qty")
        sup_qs   = fd.getlist("inv_supplier_qty")
        decs     = fd.getlist("inv_decision")
        for i, pn in enumerate(part_nos):
            pn = (pn or "").strip()
            if not pn:
                continue
            rows.append({
                "part_no": pn,
                "customer_qty": _int(cust_qs[i] if i < len(cust_qs) else None),
                "inhouse_qty":  _int(in_qs[i]   if i < len(in_qs)   else None),
                "supplier_qty": _int(sup_qs[i]  if i < len(sup_qs)  else None),
                "decision":     ((decs[i] if i < len(decs) else "") or "").strip() or None,
            })
        form.he_inventory_data = json.dumps(rows, ensure_ascii=False) if rows else None
        # 同時聚合到舊欄位（顯示總和，向下相容）
        form.he_customer_qty = sum((r["customer_qty"] or 0) for r in rows) or None
        form.he_inhouse_qty  = sum((r["inhouse_qty"]  or 0) for r in rows) or None
        form.he_supplier_qty = sum((r["supplier_qty"] or 0) for r in rows) or None

    # 舊版 LAB_TEST 還在 picked 也照存（避免歷史單破壞）
    if "LAB_TEST" in picked and "SPECIAL_ACCEPT" not in picked:
        form.lab_test_qty        = _int(fd.get("lab_test_qty"))
        form.lab_test_conditions = (fd.get("lab_test_conditions") or "").strip() or None
        form.lab_test_due_date   = (fd.get("lab_test_due_date") or "").strip() or None

    notify_pc = (fd.get("notify_pc") == "1")
    # 「送出 + 通知生管」按鈕：只觸發 send-to-prod，狀態維持在品保判斷階段
    # 主送出按鈕：才推進到 PENDING_IMPROVEMENT
    if not notify_pc:
        form.status = QCExceptionStatus.PENDING_IMPROVEMENT
    form.updated_at = datetime.utcnow()
    summary = "+".join(picked)
    if notify_pc:
        action_lbl, log_note = "SEND_TO_PC", f"暫存處理判斷 + 送生管：{summary}"
    else:
        action_lbl, log_note = "DISPOSITION", f"處理判斷：{summary}｜{note.strip()[:120]}"
    db.add(_log(form, current_user, action_lbl, old, form.status, log_note))
    await db.commit()

    # 通知 pipeline
    try:
        # 「送出 + 通知生管」按鈕：只觸發送生管，不發整體 disposition / mail / 退貨通知
        if not notify_pc:
            await qc_notif.notify_disposition(
                db, form, disposer_name=(current_user.display_name or current_user.username))
            # Step 1：通知信（給供應商或工站）
            if form.supplier_mail_to and form.supplier_mail_body:
                await qc_notif.send_supplier_mail(form)
                form.supplier_mail_sent_at = datetime.utcnow()
                await db.commit()
            # Step 2：A. 退貨 → 額外通知採購（進貨）/ 生管（製程）
            if "RETURN_TO_SUPPLIER" in picked:
                await qc_notif.notify_return_to_supplier(db, form)
        # notify_pc=1 + 有 Sorting/Rework → 送生管（狀態維持 PENDING_DISPOSITION）
        if notify_pc and (form.sa_need_sorting or form.sa_need_rework):
            form.sa_sent_to_prod_at = datetime.utcnow()
            await db.commit()
            pc_msg = (f"📋 【特採允收 — 待生管處理】{form.form_id}\n"
                      f"品號：{form.part_no}　異常：{form.defect_cause}\n"
                      f"處理：{('Sorting ' if form.sa_need_sorting else '')}"
                      f"{('Rework ' if form.sa_need_rework else '')}\n"
                      f"請填執行單位 + 安排執行後回填數量。\n"
                      f"系統：/qc-exceptions/{form.form_id}")
            await qc_notif._send_line_group(qc_notif.LINE_QC_GROUP, pc_msg)
            await qc_notif._ntf._notify_roles(db, [Role.PC], pc_msg)
    except Exception:
        logging.exception("notify_disposition pipeline failed")
    return RedirectResponse(url=f"/qc-exceptions/{form_id}", status_code=303)


# ── 立即處理 v3：送生管 / Sorting/Rework 回填 / 盤點單回填 ─────

@router.post("/{form_id}/send-to-prod")
async def send_to_prod(
    form_id: str,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    """品保把特採允收 (SORTING / REWORK) 的單送給生管，由生管填執行單位 + 安排回填數量"""
    form = await _get_or_404(form_id, db)
    if current_user.role not in _QC_ROLES:
        raise HTTPException(status_code=403)
    if not (form.sa_need_sorting or form.sa_need_rework):
        raise HTTPException(status_code=400, detail="本單未勾 Sorting / Rework，無需送生管")
    form.sa_sent_to_prod_at = datetime.utcnow()
    db.add(_log(form, current_user, "SEND_TO_PROD", form.status, form.status,
                "送生管：請填執行單位 + 回填 sorting/rework 數量"))
    await db.commit()
    msg = (f"📋 【特採允收 — 待生管處理】{form.form_id}\n"
           f"品號：{form.part_no}　異常：{form.defect_cause}\n"
           f"處理：{('Sorting ' if form.sa_need_sorting else '')}{('Rework ' if form.sa_need_rework else '')}\n"
           f"請填執行單位 + 安排執行後回填數量。\n"
           f"系統：/qc-exceptions/{form.form_id}")
    try:
        await qc_notif._send_line_group(qc_notif.LINE_QC_GROUP, msg)
        await qc_notif._ntf._notify_roles(db, [Role.PC], msg)
    except Exception:
        logging.exception("send_to_prod notify failed")
    return RedirectResponse(url=f"/qc-exceptions/{form_id}", status_code=303)


@router.post("/{form_id}/save-sa-fillback")
async def save_sa_fillback(
    form_id: str,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
    sa_station: str = Form(""),                  # 生管填
    sa_sorting_pass_qty: str = Form(""),         # sorting 單位回填
    sa_sorting_fail_qty: str = Form(""),
    sa_rework_pass_qty:  str = Form(""),         # rework 單位回填
    sa_rework_fail_qty:  str = Form(""),
    sa_rework_defect_handling: str = Form(""),   # Rework 後不良品處理方式
    sa_rework_result:    str = Form(""),         # （舊）rework 結果回報
):
    """生管 / 品保 / admin 回填 SA 處理結果（彙整表：Sorting 良/不良 + Rework 良/不良/不良品處理）"""
    form = await _get_or_404(form_id, db)
    if current_user.role not in (Role.PC, Role.QC, Role.ADMIN):
        raise HTTPException(status_code=403, detail="僅 生管 / 品保 可回填")
    def _int(s):
        try: return int(s)
        except (TypeError, ValueError): return None
    if sa_station.strip():
        form.sa_station = sa_station.strip()
    s_pass = _int(sa_sorting_pass_qty); s_fail = _int(sa_sorting_fail_qty)
    if s_pass is not None or s_fail is not None:
        form.sa_sorting_pass_qty = s_pass
        form.sa_sorting_fail_qty = s_fail
        form.sa_sorting_filled_at = datetime.utcnow()
    r_pass = _int(sa_rework_pass_qty); r_fail = _int(sa_rework_fail_qty)
    if r_pass is not None or r_fail is not None:
        form.sa_rework_pass_qty = r_pass
        form.sa_rework_fail_qty = r_fail
        form.sa_rework_filled_at = datetime.utcnow()
    if sa_rework_defect_handling.strip():
        form.sa_rework_defect_handling = sa_rework_defect_handling.strip()
    if sa_rework_result.strip():
        form.sa_rework_result = sa_rework_result.strip()
        form.sa_rework_filled_at = datetime.utcnow()
    db.add(_log(form, current_user, "SA_FILLBACK", form.status, form.status,
                f"回填：站別={form.sa_station or '—'} sort={s_pass}/{s_fail} rework={r_pass}/{r_fail}"))
    await db.commit()
    try:
        msg = (f"✅ 【特採處理結果回填】{form.form_id}\n品號：{form.part_no}\n"
               f"站別：{form.sa_station or '—'}\n"
               f"Sorting 良/不良：{s_pass if s_pass is not None else '—'}/{s_fail if s_fail is not None else '—'}\n"
               f"Rework 良/不良：{r_pass if r_pass is not None else '—'}/{r_fail if r_fail is not None else '—'}")
        await qc_notif._ntf._notify_roles(db, [Role.QC], msg)
    except Exception:
        pass
    return RedirectResponse(url=f"/qc-exceptions/{form_id}", status_code=303)


@router.post("/{form_id}/save-inventory")
async def save_inventory(
    form_id: str, request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    """橫向展開盤點單回填 — 各角色各填自己欄位
       業助 → customer_qty / 倉管 → inhouse_qty / 採購 → supplier_qty / 品保 → decision
    """
    form = await _get_or_404(form_id, db)
    if current_user.role not in (Role.ASSISTANT, Role.WAREHOUSE, Role.PURCHASE,
                                 Role.QC, Role.ADMIN):
        raise HTTPException(status_code=403)
    fd = await request.form()
    def _int(s):
        try: return int(s)
        except (TypeError, ValueError): return None
    # 載入既有
    try:
        existing = json.loads(form.he_inventory_data) if form.he_inventory_data else []
    except Exception:
        existing = []
    new_part_nos = fd.getlist("inv_part_no")
    new_cust = fd.getlist("inv_customer_qty")
    new_in   = fd.getlist("inv_inhouse_qty")
    new_sup  = fd.getlist("inv_supplier_qty")
    new_dec  = fd.getlist("inv_decision")
    # 角色決定可寫欄位
    can_cust = current_user.role in (Role.ASSISTANT, Role.QC, Role.ADMIN)
    can_in   = current_user.role in (Role.WAREHOUSE, Role.QC, Role.ADMIN)
    can_sup  = current_user.role in (Role.PURCHASE, Role.QC, Role.ADMIN)
    can_dec  = current_user.role in (Role.QC, Role.ADMIN)
    # 重組 rows — 以新提交的 part_no 為主
    rows = []
    for i, pn in enumerate(new_part_nos):
        pn = (pn or "").strip()
        if not pn:
            continue
        # 同 part_no 沿用既有資料
        old = next((r for r in existing if r.get("part_no") == pn), {})
        rows.append({
            "part_no": pn,
            "customer_qty": (_int(new_cust[i]) if can_cust and i < len(new_cust) else old.get("customer_qty")),
            "inhouse_qty":  (_int(new_in[i])   if can_in   and i < len(new_in)   else old.get("inhouse_qty")),
            "supplier_qty": (_int(new_sup[i])  if can_sup  and i < len(new_sup)  else old.get("supplier_qty")),
            "decision":     (((new_dec[i] if i < len(new_dec) else "") or "").strip() or None
                             if can_dec else old.get("decision")),
        })
    form.he_inventory_data = json.dumps(rows, ensure_ascii=False) if rows else None
    form.he_customer_qty = sum((r["customer_qty"] or 0) for r in rows) or None
    form.he_inhouse_qty  = sum((r["inhouse_qty"]  or 0) for r in rows) or None
    form.he_supplier_qty = sum((r["supplier_qty"] or 0) for r in rows) or None
    form.updated_at = datetime.utcnow()
    db.add(_log(form, current_user, "INVENTORY_FILLBACK", form.status, form.status,
                f"盤點回填（{current_user.role.value}）"))
    await db.commit()
    return RedirectResponse(url=f"/qc-exceptions/{form_id}", status_code=303)


@router.post("/{form_id}/save-improvement")
async def save_improvement(
    form_id: str,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
    notify_mail_to:   str = Form(""),
    notify_mail_cc:   str = Form(""),
    root_cause:       str = Form(""),
    need_drawing_rev: str = Form(""),
    need_sop_rev:     str = Form(""),
    need_sip_rev:     str = Form(""),
    improvement_plan: str = Form(""),
    advance:          str = Form(""),  # "ecn" / "close"
):
    """改善方案（合併 Mail 通知 + 根因分析）：可選擇推進 LINKED_ECN 或直接結案"""
    form = await _get_or_404(form_id, db)
    if current_user.role not in _QC_ROLES:
        raise HTTPException(status_code=403)
    # PENDING_RCA 為舊狀態（已併入 IMPROVEMENT），仍接受編輯避免舊資料卡死
    if form.status not in (QCExceptionStatus.PENDING_IMPROVEMENT,
                           QCExceptionStatus.LINKED_ECN,
                           QCExceptionStatus.PENDING_RCA):
        raise HTTPException(status_code=400, detail="目前狀態無法編輯改善方案")
    # 通知 + 根因
    form.notify_mail_to = notify_mail_to.strip() or None
    form.notify_mail_cc = notify_mail_cc.strip() or None
    new_root = root_cause.strip() or None
    if new_root and not form.root_cause:
        form.notify_sent_at = datetime.utcnow()  # 第一次填根因時記錄
    form.root_cause = new_root
    # 改善方案
    form.need_drawing_rev = (need_drawing_rev == "1")
    form.need_sop_rev     = (need_sop_rev == "1")
    form.need_sip_rev     = (need_sip_rev == "1")
    form.improvement_plan = improvement_plan.strip() or None
    # 把 PENDING_RCA 老資料順手推進
    if form.status == QCExceptionStatus.PENDING_RCA:
        old = form.status
        form.status = QCExceptionStatus.PENDING_IMPROVEMENT
        db.add(_log(form, current_user, "MIGRATE_TO_IMPROVE", old, form.status,
                    "舊資料併入改善方案階段"))
    if advance == "ecn" and form.status == QCExceptionStatus.PENDING_IMPROVEMENT:
        old = form.status
        form.status = QCExceptionStatus.LINKED_ECN
        db.add(_log(form, current_user, "TO_ECN", old, form.status,
                    "需修訂圖面/SOP/SIP，待開 ECN"))
    elif advance == "close":
        old = form.status
        form.status = QCExceptionStatus.CLOSED
        db.add(_log(form, current_user, "CLOSE", old, form.status, "結案"))
    form.updated_at = datetime.utcnow()
    await db.commit()
    return RedirectResponse(url=f"/qc-exceptions/{form_id}", status_code=303)


@router.post("/{form_id}/link-ecn")
async def link_ecn(
    form_id: str,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
    ecn_form_id: str = Form(...),  # PCNForm.form_id (e.g. PCN-20260424-001)
):
    """把已建立的 ECN 表單綁進來"""
    form = await _get_or_404(form_id, db)
    if current_user.role not in _QC_ROLES:
        raise HTTPException(status_code=403)
    from app.models.pcn_form import PCNForm
    r = await db.execute(select(PCNForm).where(PCNForm.form_id == ecn_form_id.strip()))
    ecn = r.scalars().first()
    if not ecn:
        raise HTTPException(status_code=404, detail="找不到該 ECN 表單")
    form.linked_ecn_form_id = ecn.id
    if form.status == QCExceptionStatus.PENDING_IMPROVEMENT:
        old = form.status
        form.status = QCExceptionStatus.LINKED_ECN
        db.add(_log(form, current_user, "LINK_ECN", old, form.status,
                    f"綁定 ECN {ecn.form_id}"))
    form.updated_at = datetime.utcnow()
    await db.commit()
    return RedirectResponse(url=f"/qc-exceptions/{form_id}", status_code=303)


@router.post("/{form_id}/close")
async def close_qc(
    form_id: str,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
    comment: str = Form(""),
):
    form = await _get_or_404(form_id, db)
    if current_user.role not in _QC_ROLES:
        raise HTTPException(status_code=403)
    if form.status == QCExceptionStatus.CLOSED:
        raise HTTPException(status_code=400, detail="已結案")
    old = form.status
    form.status = QCExceptionStatus.CLOSED
    form.updated_at = datetime.utcnow()
    db.add(_log(form, current_user, "CLOSE", old, form.status, comment.strip() or "結案"))
    await db.commit()
    return RedirectResponse(url=f"/qc-exceptions/{form_id}", status_code=303)
