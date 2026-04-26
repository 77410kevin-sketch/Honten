"""QC 異常管理通知服務

通知策略：
1. 若環境變數設定 `LINE_CHANNEL_ACCESS_TOKEN` + `LINE_QC_GROUP_ID`，
   會透過 LINE Messaging API 推到 QC 異常群組
2. 同時對相關角色（品保 / 工程 / 採購 / 產線主管 / 業助 / BU）做角色推播 fallback
3. 兩個都沒設定 → console log（dry-run 模式）

需要 .env：
    LINE_CHANNEL_ACCESS_TOKEN=...
    LINE_QC_GROUP_ID=Cxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
"""
import os
import logging
from typing import Iterable

from app.models.qc_exception import (
    QCException, QCDocType, QCEventDateType, QCExceptionStage,
)
from app.models.user import Role
from app.services import npi_notification as _ntf

logger = logging.getLogger(__name__)

LINE_TOKEN     = os.getenv("LINE_CHANNEL_ACCESS_TOKEN", "").strip()
LINE_QC_GROUP  = os.getenv("LINE_QC_GROUP_ID", "").strip()

_DOC_LBL  = {"RECEIVE": "進貨單號", "PROCESS": "製程單號", "SHIP_DC": "出貨 D/C"}
_DATE_LBL = {"RECEIVE": "進貨日期", "PRODUCE": "生產日期",
             "SHIP": "出貨日期", "COMPLAINT": "客訴日期"}
_STAGE_LBL = {"IQC": "IQC", "IPQC": "IPQC", "OQC": "OQC", "INSPECTION": "品檢",
              "LASER": "雷雕", "CNC": "CNC", "ASSEMBLY": "組裝", "OTHER": "其他"}


def _send_line_group(group_id: str, message: str) -> bool:
    """送到 LINE 群組；無 token/group 則 console log（dry-run）。回傳是否真的送出。"""
    if not (LINE_TOKEN and group_id):
        logger.info(f"[LINE GROUP dry-run] target={group_id or '(unset)'}: {message[:100]}")
        print(f"\n📱 [LINE GROUP dry-run] {group_id or '(未設定 LINE_QC_GROUP_ID)'}\n"
              f"───────────────────────\n{message}\n───────────────────────\n")
        return False
    try:
        import json as _json
        from urllib import request as _urlreq, error as _urlerr
        req = _urlreq.Request(
            "https://api.line.me/v2/bot/message/push",
            data=_json.dumps({
                "to": group_id,
                "messages": [{"type": "text", "text": message[:4900]}],
            }).encode("utf-8"),
            headers={
                "Authorization": f"Bearer {LINE_TOKEN}",
                "Content-Type": "application/json",
            },
            method="POST",
        )
        with _urlreq.urlopen(req, timeout=10) as resp:
            if resp.status == 200:
                logger.info(f"[LINE group push OK] group={group_id}")
                return True
            body = resp.read().decode("utf-8", errors="ignore")[:200]
            logger.error(f"[LINE group push FAIL] {resp.status}: {body}")
            return False
    except _urlerr.HTTPError as e:
        body = e.read().decode("utf-8", errors="ignore")[:200] if e.fp else ""
        logger.error(f"[LINE group push HTTPError] {e.code}: {body}")
        return False
    except Exception as e:
        logger.error(f"[LINE group push ERROR] {e}")
        return False


def build_exception_message(form: QCException, creator_name: str = "") -> str:
    """組成 LINE 推播訊息文字（純文字，emoji 加強可讀性）"""
    doc_label = _DOC_LBL.get(form.doc_type.value if form.doc_type else "", "單號")
    date_label = _DATE_LBL.get(form.event_date_type.value if form.event_date_type else "", "日期")
    stage_label = _STAGE_LBL.get(form.stage.value if form.stage else "", form.stage.value if form.stage else "—")

    rate_str = "—"
    if form.defect_rate is not None:
        rate_str = f"{form.defect_qty or 0} / {form.sample_qty or 0} = {form.defect_rate * 100:.1f}%"

    lines = [
        "🚨 【QC 異常通知】",
        f"單號：{form.form_id}",
        f"品號：{form.part_no}",
        f"{doc_label}：{form.receive_doc_no or '—'}",
        f"{date_label}：{form.receive_date or '—'}",
        f"工段／廠商：{stage_label} ／ {form.supplier_name or '—'}",
        f"數量：{form.receive_qty or '—'} pcs",
        "",
        f"❗ 異常原因：{form.defect_cause}",
        f"📐 量測數據：{form.measurement_data or '—'}",
        f"📊 不良率：{rate_str}",
        "",
        f"建立者：{creator_name or '—'}",
        f"請至系統處理：/qc-exceptions/{form.form_id}",
    ]
    return "\n".join(lines)


# 品保 + 工程 + 採購 + 產線主管 + 生管 + 業助 + BU
_NOTIFY_ROLES = (Role.QC, Role.ENGINEER, Role.ENG_MGR, Role.PURCHASE,
                 Role.PROD_MGR, Role.PC, Role.ASSISTANT, Role.BU)


async def notify_exception_created(db, form: QCException, creator_name: str = ""):
    """異常單建立後通知 — LINE 群組 + 個別角色 fallback"""
    msg = build_exception_message(form, creator_name)
    # 1) LINE 群組推播
    _send_line_group(LINE_QC_GROUP, msg)
    # 2) 個別相關角色推播（即使群組也送）
    try:
        await _ntf._notify_roles(db, _NOTIFY_ROLES, msg)
    except Exception as e:
        logger.error(f"[QC notify roles error] {e}")


async def notify_return_to_supplier(db, form: QCException):
    """A. 退貨 → 依異常品來源路由通知：
       進貨單號 (RECEIVE)  → 採購（原物料）
       製程單號 (PROCESS)  → 生管（製成品）
       出貨 D/C (SHIP_DC)  → 採購 + 生管
    """
    dt = form.doc_type.value if form.doc_type else "RECEIVE"
    if dt == "PROCESS":
        roles = [Role.PC]
        target_label = "生管（製成品）"
    elif dt == "SHIP_DC":
        roles = [Role.PURCHASE, Role.PC]
        target_label = "採購 + 生管（出貨）"
    else:
        roles = [Role.PURCHASE]
        target_label = "採購（原物料）"
    extra = f"\n補貨資訊請求：{form.rts_replenish_note}" if form.rts_replenish_note else ""
    msg = (f"📤 【退貨通知 — {target_label}】{form.form_id}\n"
           f"品號：{form.part_no}　廠商：{form.supplier_name or '—'}\n"
           f"單據：{form.receive_doc_no or '—'}（{dt}）\n"
           f"異常：{form.defect_cause}\n"
           f"請進行退貨／後續處理。{extra}\n"
           f"系統：/qc-exceptions/{form.form_id}")
    _send_line_group(LINE_QC_GROUP, msg)
    try:
        await _ntf._notify_roles(db, roles, msg)
    except Exception as e:
        logger.error(f"[QC return-to-supplier notify error] {e}")


async def send_supplier_mail(form: QCException):
    """寄退貨通知信給供應商（含異常照片附件）。SMTP 未設則 dry-run。"""
    if not (form.supplier_mail_to and form.supplier_mail_body):
        return False
    cc = [c.strip() for c in (form.supplier_mail_cc or "").split(",") if c.strip()]
    # 收集異常照片附件路徑
    att_paths = []
    try:
        for d in (form.documents or []):
            if (d.category or "") == "異常照片":
                p = os.path.join("uploads", f"qc_{form.id}", d.filename)
                if os.path.exists(p):
                    att_paths.append(p)
    except Exception:
        pass
    subj = form.supplier_mail_subject or f"【鴻騰電子 QC 異常通知】{form.form_id} - {form.part_no}"
    _ntf._send_mail(form.supplier_mail_to, subj, form.supplier_mail_body,
                    attachments=att_paths, cc=cc)
    return True


def build_supplier_mail_template(form: QCException, contact_name: str = "") -> str:
    """產出退貨給供應商的預設 mail 範本（讓品保編輯）

    強調：要說明「退貨」、要帶「異常廠商」名稱（不是工段名）
    """
    rate_str = "—"
    if form.defect_rate is not None:
        rate_str = f"{form.defect_qty or 0} / {form.sample_qty or 0} ({form.defect_rate*100:.1f}%)"
    supplier = form.supplier_name or "（廠商）"
    greeting = f"{supplier} {contact_name or '聯絡人'} 您好：" if contact_name \
               else f"{supplier} 您好："
    return (
        f"{greeting}\n\n"
        f"本公司於進料檢驗 貴司（{supplier}）所出貨之品號 {form.part_no} 時發現異常，\n"
        f"經 IQC 確認本批不符合允收規格，**將辦理退貨**，請協助分析原因並回覆「異常分析報告」。\n\n"
        f"━━━━━━━━━━━━━━━━━━━━━━━\n"
        f"異常單號：{form.form_id}\n"
        f"異常廠商：{supplier}\n"
        f"品號：{form.part_no}\n"
        f"進貨單號：{form.receive_doc_no or '—'}\n"
        f"進貨日期：{form.receive_date or '—'}\n"
        f"異常原因：{form.defect_cause}\n"
        f"量測數據：{form.measurement_data or '—'}\n"
        f"異常抽驗比例：{rate_str}\n"
        f"━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        f"處理方式：本批整批退回 貴司，麻煩於收到後 3 個工作天內回覆異常分析報告，\n"
        f"內容請包含：5Why 分析、根本原因、暫時對策、永久對策、相同批號是否需擴大調查。\n\n"
        f"附件：異常照片如附\n\n"
        f"如有任何疑問請回信至本信件。\n\n"
        f"鴻騰電子 品保部 敬上"
    )


async def notify_disposition(db, form: QCException, disposer_name: str = ""):
    """品保下處理判斷後通知"""
    d_lbl = {"RETURN_TO_SUPPLIER": "退貨", "LAB_TEST": "實驗測試",
             "SPECIAL_ACCEPT": "特採允收"}.get(
        form.disposition.value if form.disposition else "", "—")
    msg = (f"✅ 【QC 處理判斷】{form.form_id}\n"
           f"品號：{form.part_no}\n"
           f"判定：{d_lbl}\n"
           f"判定人：{disposer_name or '—'}\n"
           f"備註：{form.disposition_note or '—'}\n"
           f"系統：/qc-exceptions/{form.form_id}")
    _send_line_group(LINE_QC_GROUP, msg)
    try:
        await _ntf._notify_roles(db, _NOTIFY_ROLES, msg)
    except Exception as e:
        logger.error(f"[QC disposition notify error] {e}")
