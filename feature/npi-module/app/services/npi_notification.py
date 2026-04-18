"""
NPI 模組通知服務（Demo 版 — 以 console log 模擬實際 LINE/SMTP 呼叫）

正式環境替換重點：
- _send_line_push: LINE Messaging API push / broadcast
- _send_mail: SMTP 或 Google Workspace API
- _copy_to_nas: mount NAS 路徑 shutil.copy 或透過 rsync
"""
import os
import shutil
import logging
from datetime import datetime, timedelta
from typing import Iterable

from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy.orm import selectinload

from app.models.user import User, Role
from app.models.npi_form import NPIForm, NPISupplierInvite, NPIDocument
from app.models.supplier import Supplier

logger = logging.getLogger(__name__)

# NAS 根路徑（正式環境從 .env 讀，Demo 寫本機）
NAS_ROOT = os.getenv("NPI_NAS_ROOT", "nas_npi")
UPLOAD_BASE = "uploads"
# 未回覆幾天後自動跟催
REMINDER_DAYS = int(os.getenv("NPI_REMINDER_DAYS", "2"))


# ── 低層 helpers ─────────────────────────────────

def _send_line_push(target: str, message: str):
    logger.info(f"[LINE] → {target}: {message}")
    print(f"\n📱 [LINE] {target}\n   {message}\n")


def _send_mail(to_addr: str, subject: str, body: str, attachments: Iterable[str] = (),
               cc: Iterable[str] = ()):
    """寄送 mail。

    - 若 .env 設定 `SMTP_HOST` → 實際透過 smtplib 寄出（含附件、CC）
    - 否則退回 Demo 模式，print 到 console（維持開發測試不中斷）

    所需環境變數：
        SMTP_HOST, SMTP_PORT (default 587), SMTP_USER, SMTP_PASSWORD
        SMTP_FROM (default = SMTP_USER)
        SMTP_FROM_NAME (default 鴻騰電子 NPI 系統)
        SMTP_SSL (true = 465 direct SSL；預設 false = 587 STARTTLS)
        SMTP_REPLY_TO (optional)
    """
    att_names = ", ".join(os.path.basename(a) for a in attachments) or "—"
    cc_list = [c for c in cc if c]

    host = os.getenv("SMTP_HOST", "").strip()
    if not host:
        # Demo 模式：不寄，只 log
        logger.info(f"[MAIL dry-run] → {to_addr} | {subject} | att={att_names}")
        cc_str = (", ".join(cc_list)) if cc_list else "—"
        print(f"\n📧 [MAIL dry-run] To: {to_addr}  CC: {cc_str}\n"
              f"   Subject: {subject}\n   Attachments: {att_names}\n"
              f"   Body: {body[:120]}...\n")
        return

    port = int(os.getenv("SMTP_PORT", "587"))
    user = os.getenv("SMTP_USER", "").strip()
    password = os.getenv("SMTP_PASSWORD", "")
    from_addr = os.getenv("SMTP_FROM", user).strip()
    from_name = os.getenv("SMTP_FROM_NAME", "鴻騰電子 NPI 系統").strip()
    use_ssl = os.getenv("SMTP_SSL", "false").lower() in ("1", "true", "yes")
    reply_to = os.getenv("SMTP_REPLY_TO", "").strip()

    try:
        import smtplib
        import mimetypes
        from email.message import EmailMessage
        from email.utils import formataddr

        msg = EmailMessage()
        msg["From"] = formataddr((from_name, from_addr)) if from_name else from_addr
        msg["To"] = to_addr
        if cc_list:
            msg["Cc"] = ", ".join(cc_list)
        if reply_to:
            msg["Reply-To"] = reply_to
        msg["Subject"] = subject
        msg.set_content(body)

        for path in attachments:
            if not path or not os.path.exists(path):
                logger.warning(f"[MAIL] skip missing attachment: {path}")
                continue
            mime, _ = mimetypes.guess_type(path)
            maintype, subtype = (mime or "application/octet-stream").split("/", 1)
            with open(path, "rb") as f:
                msg.add_attachment(f.read(), maintype=maintype, subtype=subtype,
                                   filename=os.path.basename(path))

        if use_ssl:
            server = smtplib.SMTP_SSL(host, port, timeout=30)
        else:
            server = smtplib.SMTP(host, port, timeout=30)
            server.ehlo()
            server.starttls()
            server.ehlo()
        if user and password:
            server.login(user, password)
        recipients = [to_addr] + cc_list
        server.send_message(msg, from_addr=from_addr, to_addrs=recipients)
        server.quit()
        logger.info(f"[MAIL sent] → {to_addr} (cc={len(cc_list)}) | {subject} | att={att_names}")
    except Exception as e:
        logger.error(f"[MAIL FAILED] to={to_addr} err={e}")
        # 不 raise — 避免寄信失敗卡住整個流程；用 print 讓 demo 也能看到
        print(f"\n❌ [MAIL FAILED] To: {to_addr} | {subject}\n   Error: {e}\n")


async def _users_by_role(db: AsyncSession, role: Role) -> list[User]:
    r = await db.execute(select(User).where(User.role == role, User.is_active == True))
    return list(r.scalars().all())


async def _notify_roles(db: AsyncSession, roles: Iterable[Role], message: str):
    for role in roles:
        for u in await _users_by_role(db, role):
            tag = u.line_user_id or f"{u.display_name}({u.role.value})"
            _send_line_push(tag, message)


# ── NPI 流程節點通知 ─────────────────────────────

async def notify_sales_submitted(db: AsyncSession, form: NPIForm):
    """業務 DRAFT → ENG_DISPATCH：通知工程"""
    msg = f"【NPI-RFQ 待工程排製程】{form.form_id} - {form.customer_name}/{form.product_name}"
    await _notify_roles(db, [Role.ENGINEER, Role.ENG_MGR], msg)


_DEFAULT_RFQ_TEMPLATE = (
    "Hi {contact}，\n\n"
    "請依附件圖檔與提供的材質 {material}、MOQ {moq} 協助報價。\n\n"
    "案件資訊：\n"
    " 詢價單號：{form_id}\n"
    " 製程：{process}\n"
    " 對應圖面：{drawing}\n"
    " 客戶回覆期限：{due_date}\n\n"
    "若有疑問請直接回信聯絡，謝謝！\n\n"
    "— 鴻騰電子 NPI 採購"
)


def _render_rfq_body(template: str, *, form, invite, supplier, material, moq) -> str:
    """以變數替換組信件內文。未知變數保持原樣，不報錯。"""
    values = {
        "contact":  supplier.contact or "您好",
        "supplier": supplier.name or "",
        "material": material or "—",
        "moq":      moq if moq not in (None, "") else "—",
        "process":  invite.process_name or "—",
        "drawing":  invite.drawing.original_name if invite.drawing else "共用圖面",
        "form_id":  form.form_id,
        "product":  form.product_name or "",
        "customer": form.customer_name or "",
        "due_date": form.rfq_due_date or "—",
    }
    out = template or _DEFAULT_RFQ_TEMPLATE
    for k, v in values.items():
        out = out.replace("{" + k + "}", str(v))
    return out


def _collect_drawing_meta(invites):
    """從所有 invites 組出 {drawing_id: {material, qty}} map + fallback。"""
    m: dict = {}
    for i in invites:
        key = i.drawing_doc_id or 0
        if key not in m:
            m[key] = {"material": None, "qty": None}
        if i.material and not m[key]["material"]:
            m[key]["material"] = i.material
        if i.qty and not m[key]["qty"]:
            m[key]["qty"] = i.qty
    fb_mat = next((v["material"] for v in m.values() if v["material"]), None)
    fb_qty = next((v["qty"] for v in m.values() if v["qty"]), None)
    return m, fb_mat, fb_qty


async def notify_quotes_dispatched(
    db: AsyncSession,
    form: NPIForm,
    invites: list[NPISupplierInvite],
    *,
    merge: bool = False,
):
    """工程派發詢價：向每家供應商寄 mail（第一次）+ 通知業務 / 工程主管。
    - 支援使用者自訂的信件模板（form.eng_process_note），依供應商 / 圖替換變數。
    - merge=True：同一供應商若有多筆 invite → 合併成一封信，列多項目 + 合併附件。
    """
    # CC 對象：業務（建單者） + 工程（指派者）
    cc_list = []
    if form.creator and getattr(form.creator, "email", None):
        cc_list.append(form.creator.email)
    if form.assigned_eng and getattr(form.assigned_eng, "email", None):
        cc_list.append(form.assigned_eng.email)

    drawing_meta, fallback_mat, fallback_qty = _collect_drawing_meta(invites)
    template = (form.eng_process_note or "").strip() or _DEFAULT_RFQ_TEMPLATE

    def _attachments_for(inv):
        paths = []
        if inv.drawing:
            src = os.path.join(UPLOAD_BASE, f"npi_{form.id}", inv.drawing.filename)
            if os.path.exists(src):
                paths.append(src)
        else:
            for d in form.documents:
                if d.category == "圖面":
                    p = os.path.join(UPLOAD_BASE, f"npi_{form.id}", d.filename)
                    if os.path.exists(p):
                        paths.append(p)
        return paths

    def _meta_for(inv):
        meta = drawing_meta.get(inv.drawing_doc_id or 0, {})
        return (meta.get("material") or fallback_mat,
                meta.get("qty") or fallback_qty)

    if merge:
        # 依 supplier_id 群組
        groups: dict = {}
        for inv in invites:
            groups.setdefault(inv.supplier_id, []).append(inv)

        for sid, group in groups.items():
            first = group[0]
            sup: Supplier | None = first.supplier
            if not sup or not sup.email:
                logger.warning(f"Supplier {sid} 沒有 email，略過")
                continue

            # 單筆 → 走單項模板邏輯（與非 merge 模式一致）
            if len(group) == 1:
                mat, qty = _meta_for(first)
                subject = (f"【鴻騰電子 RFQ 詢價】{form.form_id} - "
                           f"{form.product_name}{(' / ' + first.process_name) if first.process_name else ''}")
                body = _render_rfq_body(template, form=form, invite=first, supplier=sup,
                                        material=mat, moq=qty)
                atts = _attachments_for(first)
                _send_mail(sup.email, subject, body, atts, cc=cc_list)
                first.first_sent_at = datetime.utcnow()
                continue

            # 多筆 → 合併：模板中的單項變數（製程/圖/材質/MOQ）改成「見下列項目」
            merged_template = (template
                               .replace("{process}", "（見下列項目）")
                               .replace("{drawing}", "（見下列項目）")
                               .replace("{material}", "（見下列項目）")
                               .replace("{moq}", "（見下列項目）"))
            intro = _render_rfq_body(merged_template, form=form, invite=first, supplier=sup,
                                     material="（見下列項目）", moq="（見下列項目）")

            lines = ["", "━━━━━━━━━━━━━━━━━━━━━━━", "本次詢價項目：", ""]
            att_set: list = []
            att_seen: set = set()
            for idx, inv in enumerate(group, 1):
                mat, qty = _meta_for(inv)
                drawing_name = inv.drawing.original_name if inv.drawing else "共用圖面"
                lines.append(f"{idx}. 製程：{inv.process_name or '—'}")
                lines.append(f"   對應圖面：{drawing_name}")
                lines.append(f"   材質：{mat or '—'}　MOQ：{qty if qty not in (None, '') else '—'}")
                lines.append("")
                for p in _attachments_for(inv):
                    if p not in att_seen:
                        att_seen.add(p)
                        att_set.append(p)
            body = intro + "\n".join(lines)

            # 主旨以項目數表示
            subject = (f"【鴻騰電子 RFQ 詢價】{form.form_id} - "
                       f"{form.product_name}（共 {len(group)} 項製程）")

            _send_mail(sup.email, subject, body, att_set, cc=cc_list)
            now = datetime.utcnow()
            for inv in group:
                inv.first_sent_at = now
    else:
        # 原始：逐筆寄送
        for inv in invites:
            sup: Supplier | None = inv.supplier
            if not sup or not sup.email:
                logger.warning(f"Supplier {inv.supplier_id} 沒有 email，略過")
                continue
            mat, qty = _meta_for(inv)
            subject = (f"【鴻騰電子 RFQ 詢價】{form.form_id} - "
                       f"{form.product_name}{(' / ' + inv.process_name) if inv.process_name else ''}")
            body = _render_rfq_body(template, form=form, invite=inv, supplier=sup,
                                    material=mat, moq=qty)
            _send_mail(sup.email, subject, body, _attachments_for(inv), cc=cc_list)
            inv.first_sent_at = datetime.utcnow()
    # 業務 + 工程主管
    await _notify_roles(
        db, [Role.SALES, Role.ENG_MGR],
        f"【NPI-RFQ 已派發詢價】{form.form_id} 共 {len(invites)} 家供應商",
    )


async def notify_remind_overdue(db: AsyncSession, form: NPIForm, inv: NPISupplierInvite):
    """2 天未回覆自動跟催（由定時任務觸發）"""
    sup = inv.supplier
    if not sup or not sup.email:
        return
    subject = f"【跟催 RFQ 報價】{form.form_id} - {form.product_name}"
    body = (
        f"您好 {sup.contact or ''}，\n"
        f"本案詢價信已於 {inv.first_sent_at:%Y-%m-%d} 發出，尚未收到報價回覆，\n"
        f"煩請儘速提供報價，謝謝。"
    )
    _send_mail(sup.email, subject, body)
    inv.last_reminder_at = datetime.utcnow()
    inv.reminder_count = (inv.reminder_count or 0) + 1


async def auto_remind_non_responders(db: AsyncSession) -> int:
    """排程入口：掃全部 QUOTING 狀態的單，對超過 REMINDER_DAYS 未回的供應商發跟催信"""
    from app.models.npi_form import NPIFormStatus
    now = datetime.utcnow()
    r = await db.execute(
        select(NPIForm)
        .where(NPIForm.status == NPIFormStatus.QUOTING)
        .options(selectinload(NPIForm.invites).selectinload(NPISupplierInvite.supplier),
                 selectinload(NPIForm.documents))
    )
    forms = r.scalars().all()
    sent = 0
    for f in forms:
        for inv in f.invites:
            if inv.replied_at:
                continue
            if not inv.first_sent_at:
                continue
            last = inv.last_reminder_at or inv.first_sent_at
            if now - last >= timedelta(days=REMINDER_DAYS):
                await notify_remind_overdue(db, f, inv)
                sent += 1
    if sent:
        await db.commit()
    return sent


async def notify_quote_replied(db: AsyncSession, form: NPIForm, inv: NPISupplierInvite):
    """供應商回覆報價後 → 通知工程 + 業務；報價檔自動落 NAS"""
    # 將報價附件複製到 NAS
    nas_dir = _ensure_nas_dir(form, "RFQ_Quotes")
    _copy_invite_quote_files_to_nas(form, inv, nas_dir)
    sup_name = inv.supplier.name if inv.supplier else "(未指定)"
    msg = f"【RFQ 供應商回覆】{form.form_id} - {sup_name} 金額 {inv.quote_amount or '—'}"
    await _notify_roles(db, [Role.ENGINEER, Role.SALES], msg)


async def notify_sales_cost_analysis_done(db: AsyncSession, form: NPIForm):
    """業務完成成本分析與客戶報價單 → 發送客戶 mail"""
    if form.customer_email:
        subject = f"【鴻騰電子 RFQ 報價回覆】{form.form_id} - {form.product_name}"
        body = (
            f"{form.customer_name} 您好，\n\n"
            f"針對貴司詢價案件，謹附上本公司之成本分析與正式報價單，敬請查收。\n"
            f"如有任何問題請隨時告知。\n\n"
            f"鴻騰電子 業務部 敬上"
        )
        att = []
        for d in form.documents:
            if d.category in ("成本分析表", "客戶報價單"):
                att.append(os.path.join(UPLOAD_BASE, f"npi_{form.id}", d.filename))
        _send_mail(form.customer_email, subject, body, att)
    await _notify_roles(db, [Role.BU, Role.SALES, Role.ENGINEER],
                       f"【RFQ 已發送客戶報價】{form.form_id} - {form.customer_name}")


async def notify_npi_started(db: AsyncSession, form: NPIForm):
    """客戶確定開發 → 通知工程選供應商並開 ERP 模具請購單"""
    await _notify_roles(
        db, [Role.ENGINEER, Role.ENG_MGR],
        f"【NPI 成案啟動】{form.form_id} - 請工程依成本分析選供應商並開 ERP 模具請購單",
    )


async def notify_npi_submit_bu(db: AsyncSession, form: NPIForm):
    """工程送 BU 核准"""
    await _notify_roles(
        db, [Role.BU],
        f"【NPI 待 BU 核准】{form.form_id} - {form.customer_name}/{form.product_name} 預估模具成本 {form.mould_cost_est or '—'}",
    )


async def notify_npi_approved(db: AsyncSession, form: NPIForm):
    """BU 核准 NPI → 通知採購議價"""
    await _notify_roles(
        db, [Role.PURCHASE],
        f"【NPI 待採購議價】{form.form_id} - 預估模具成本 {form.mould_cost_est or '—'}",
    )


async def notify_npi_rejected(db: AsyncSession, form: NPIForm, target: str):
    await _notify_roles(
        db, [Role.ENGINEER, Role.SALES],
        f"【NPI 退回 {target}】{form.form_id} - 請調整後重送",
    )


async def notify_quote_approved(db: AsyncSession, form: NPIForm):
    """BU 核准報價後：把成本分析 + 客戶報價單 HTML 落地 NAS 的 RFQ_Quote 資料夾。

    存入：
    - cost_analysis/quote_summary.json（內部：完整試算 + 利潤 + BU 評語）
    - cost_analysis/*（成本分析類附件，內部用）
    - supplier_quotes/*（供應商報價單，內部參考）
    - customer_quote/{form_id}_報價單.html（對外版本，業務可直接寄給客戶）
    - customer_quote/*（業務已上傳的客戶報價相關附件）
    """
    import json as _json
    nas_dir = _ensure_nas_dir(form, "RFQ_Quote")
    # 1. 內部：成本分析 JSON（含利潤/供應商資訊）
    internal_dir = os.path.join(nas_dir, "internal_cost_analysis")
    os.makedirs(internal_dir, exist_ok=True)
    summary = {
        "form_id": form.form_id,
        "customer_name": form.customer_name,
        "product_name": form.product_name,
        "product_model": form.product_model,
        "quoted_unit_price": form.quoted_unit_price,
        "cost_analysis_note": form.cost_analysis_note,
        "bu_quote_note": form.bu_quote_note,
        "quote_cost_data": _safe_parse_json(form.quote_cost_data),
        "approved_at": datetime.utcnow().isoformat(),
    }
    try:
        with open(os.path.join(internal_dir, "quote_summary.json"), "w", encoding="utf-8") as f:
            _json.dump(summary, f, ensure_ascii=False, indent=2)
    except Exception as e:
        logger.warning(f"write quote_summary.json failed: {e}")
    # 2. 附件分流：成本分析 → internal；供應商報價 → supplier_quotes；客戶報價 → customer_quote
    cat_to_dir = {
        "成本分析表": os.path.join(nas_dir, "internal_cost_analysis"),
        "供應商報價": os.path.join(nas_dir, "supplier_quotes"),
        "客戶報價單": os.path.join(nas_dir, "customer_quote"),
    }
    for d in form.documents:
        sub = cat_to_dir.get(d.category or "")
        if not sub:
            continue
        os.makedirs(sub, exist_ok=True)
        src = os.path.join(UPLOAD_BASE, f"npi_{form.id}", d.filename)
        if not os.path.exists(src):
            continue
        try:
            shutil.copy2(src, os.path.join(sub, d.original_name))
        except Exception as e:
            logger.warning(f"copy to NAS failed: {e}")
    # 3. 產出「對外客戶報價單」HTML 並落地
    try:
        from jinja2 import Environment, FileSystemLoader, select_autoescape
        env = Environment(loader=FileSystemLoader("app/templates"),
                          autoescape=select_autoescape(["html"]))
        tmpl = env.get_template("npi_forms/customer_quote.html")
        html = tmpl.render(form=form, quote_data=_safe_parse_json(form.quote_cost_data) or {},
                           now=datetime.utcnow())
        out_dir = os.path.join(nas_dir, "customer_quote")
        os.makedirs(out_dir, exist_ok=True)
        out_path = os.path.join(out_dir, f"{form.form_id}_客戶報價單.html")
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(html)
    except Exception as e:
        logger.warning(f"render customer_quote.html to NAS failed: {e}")
    # 4. 通知
    await _notify_roles(
        db, [Role.SALES, Role.BU],
        f"【報價核准並歸檔】{form.form_id} - 成本分析（內部）與客戶報價單（對外）已存入 NAS：{nas_dir}",
    )


def _safe_parse_json(s: str | None):
    if not s:
        return None
    try:
        import json as _json
        return _json.loads(s)
    except Exception:
        return s


async def notify_npi_closed(db: AsyncSession, form: NPIForm):
    """結案：資料匯入 NAS，mail 業務 + BU 主管"""
    nas_dir = _ensure_nas_dir(form, "Closure")
    _copy_all_docs_to_nas(form, nas_dir)

    # mail 業務 + BU
    recipients: list[User] = []
    recipients += await _users_by_role(db, Role.SALES)
    recipients += await _users_by_role(db, Role.BU)
    subject = f"【NPI 結案通知】{form.form_id} - {form.customer_name}/{form.product_name}"
    body = (
        f"本案 NPI 已完成結案流程。\n"
        f"客戶：{form.customer_name}\n"
        f"產品：{form.product_name} / 型號：{form.product_model or '—'}\n"
        f"最終模具成本：{form.mould_cost_final or '—'}\n"
        f"NAS 資料夾：{nas_dir}\n"
    )
    for u in recipients:
        addr = (u.username + "@honten.local") if not getattr(u, "email", None) else u.email  # type: ignore
        _send_mail(addr, subject, body)

    await _notify_roles(db, [Role.SALES, Role.BU, Role.ENGINEER, Role.PURCHASE],
                       f"【NPI 已結案】{form.form_id} - 資料已匯入 NAS")


# ── NAS 匯出 ────────────────────────────────────

def _ensure_nas_dir(form: NPIForm, sub: str) -> str:
    path = os.path.join(NAS_ROOT, form.form_id, sub)
    os.makedirs(path, exist_ok=True)
    return path


def _copy_invite_quote_files_to_nas(form: NPIForm, inv: NPISupplierInvite, nas_dir: str):
    sup_name = (inv.supplier.name if inv.supplier else f"supplier_{inv.supplier_id}").replace("/", "_")
    for d in form.documents:
        if d.invite_id_fk != inv.id:
            continue
        src = os.path.join(UPLOAD_BASE, f"npi_{form.id}", d.filename)
        if not os.path.exists(src):
            continue
        dst = os.path.join(nas_dir, f"{sup_name}__{d.original_name}")
        try:
            shutil.copy2(src, dst)
        except Exception as e:
            logger.warning(f"copy to NAS failed: {e}")


def _copy_all_docs_to_nas(form: NPIForm, nas_dir: str):
    for d in form.documents:
        src = os.path.join(UPLOAD_BASE, f"npi_{form.id}", d.filename)
        if not os.path.exists(src):
            continue
        cat = (d.category or "其它").replace("/", "_")
        sub = os.path.join(nas_dir, cat)
        os.makedirs(sub, exist_ok=True)
        try:
            shutil.copy2(src, os.path.join(sub, d.original_name))
        except Exception as e:
            logger.warning(f"copy to NAS failed: {e}")
