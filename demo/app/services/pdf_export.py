"""
CC Package PDF 生成服務
BU 核准後將表單資料 + 附件整合成單一 PDF
"""
import io, os, json, logging
from datetime import datetime
from pathlib import Path

logger = logging.getLogger(__name__)

FONT_PATH   = "/Library/Fonts/Arial Unicode.ttf"
UPLOAD_BASE = "uploads"


def _register_font():
    """註冊支援中文的字型"""
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    try:
        pdfmetrics.registerFont(TTFont("AUnicode", FONT_PATH))
        return "AUnicode"
    except Exception:
        # fallback: STHeiti
        heiti = "/System/Library/Fonts/STHeiti Light.ttc"
        try:
            pdfmetrics.registerFont(TTFont("STHeiti", heiti, subfontIndex=0))
            return "STHeiti"
        except Exception:
            return "Helvetica"


def generate_cc_pdf(form, inventory_rows=None) -> bytes:
    """
    產生 CC Package PDF，回傳 bytes
    包含：表單基本資料、核准記錄、庫存盤點、附件（圖片嵌入）
    """
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                    Table, TableStyle, Image as RLImage,
                                    HRFlowable, PageBreak)
    from reportlab.lib.enums import TA_LEFT, TA_CENTER

    font_name = _register_font()

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=2*cm, rightMargin=2*cm,
        topMargin=2*cm, bottomMargin=2*cm,
        title=f"PCN/ECN CC Package — {form.form_id}",
    )

    # ── 樣式 ──────────────────────────────────────
    normal = ParagraphStyle("normal", fontName=font_name, fontSize=9,
                             leading=14, spaceAfter=4)
    title_style = ParagraphStyle("title", fontName=font_name, fontSize=16,
                                  leading=20, textColor=colors.HexColor("#1a56db"),
                                  spaceAfter=6)
    h2 = ParagraphStyle("h2", fontName=font_name, fontSize=11,
                          leading=15, textColor=colors.HexColor("#374151"),
                          spaceBefore=12, spaceAfter=6,
                          borderPad=4)
    small = ParagraphStyle("small", fontName=font_name, fontSize=8,
                             leading=12, textColor=colors.grey)
    cell_style = ParagraphStyle("cell", fontName=font_name, fontSize=8, leading=12)

    def P(text, style=normal):
        return Paragraph(str(text or "—"), style)

    def tbl_style(header_color="#1a56db"):
        return TableStyle([
            ("BACKGROUND",  (0,0), (-1,0), colors.HexColor(header_color)),
            ("TEXTCOLOR",   (0,0), (-1,0), colors.white),
            ("FONTNAME",    (0,0), (-1,-1), font_name),
            ("FONTSIZE",    (0,0), (-1,-1), 8),
            ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#f9fafb")]),
            ("GRID",        (0,0), (-1,-1), 0.4, colors.HexColor("#d1d5db")),
            ("VALIGN",      (0,0), (-1,-1), "MIDDLE"),
            ("LEFTPADDING", (0,0), (-1,-1), 6),
            ("RIGHTPADDING",(0,0), (-1,-1), 6),
            ("TOPPADDING",  (0,0), (-1,-1), 4),
            ("BOTTOMPADDING",(0,0), (-1,-1), 4),
        ])

    story = []

    # ── 封面標題 ──────────────────────────────────
    story.append(P(f"鴻騰電子 PCN/ECN CC Package", title_style))
    story.append(P(f"單號：{form.form_id}　｜　核准日期：{datetime.now().strftime('%Y-%m-%d %H:%M')}", small))
    story.append(HRFlowable(width="100%", thickness=1.5,
                             color=colors.HexColor("#1a56db"), spaceAfter=10))

    # ── 基本資訊 ──────────────────────────────────
    story.append(P("▌ 基本資訊", h2))

    status_map = {
        "APPROVED": "已核准", "CLOSED": "已結案",
        "RETURNED": "已退回", "DRAFT": "草稿",
    }
    info_data = [
        ["欄位", "內容"],
        ["單號",         form.form_id],
        ["類型",         "PCN 開發轉量產" if form.type.value == "PCN" else "ECN 產品工程變更"],
        ["狀態",         status_map.get(form.status.value, form.status.value)],
        ["廠內產品料號", form.product_name],
        ["機種名稱",     form.product_model or "—"],
        ["提出人/部門",  form.department or "—"],
        ["提案日期",     form.effective_date or "—"],
        ["建立時間",     form.created_at.strftime("%Y-%m-%d %H:%M") if form.created_at else "—"],
        ["建單者",       form.creator.display_name if form.creator else "—"],
    ]
    if form.change_types:
        try:
            ct = "、".join(json.loads(form.change_types))
        except Exception:
            ct = form.change_types
        info_data.append(["ECN 變更類型", ct])

    t = Table([[P(r[0], cell_style), P(r[1], cell_style)] for r in info_data],
              colWidths=[4*cm, 12*cm])
    t.setStyle(tbl_style("#1a56db"))
    story.append(t)

    # ── 變更說明 ──────────────────────────────────
    story.append(Spacer(1, 0.3*cm))
    story.append(P("▌ 變更說明", h2))
    story.append(P(form.change_description or "—", normal))
    if form.change_reason:
        story.append(P(f"變更原因：{form.change_reason}", normal))

    # ── 庫存盤點 ──────────────────────────────────
    if inventory_rows:
        story.append(P("▌ 庫存盤點結果", h2))
        inv_data = [["#", "舊料號", "站別/工序", "庫存量", "處理方式", "備註"]]
        for i, row in enumerate(inventory_rows, 1):
            inv_data.append([
                str(i),
                row.get("old_pn", "—"),
                row.get("station", "—"),
                row.get("qty", "—"),
                row.get("action", "—"),
                row.get("remark", ""),
            ])
        t2 = Table([[P(c, cell_style) for c in r] for r in inv_data],
                   colWidths=[0.8*cm, 3.5*cm, 3*cm, 2*cm, 3*cm, 3.7*cm])
        t2.setStyle(tbl_style("#198754"))
        story.append(t2)

    # ── 核准記錄 ──────────────────────────────────
    story.append(P("▌ 審核記錄", h2))
    action_map = {
        "SUBMIT": "送審", "APPROVE": "核准", "REJECT": "退回",
        "ENG_CONFIRM": "工程確認", "ECN_QC_CONFIRM": "品保確認",
        "WH_CONFIRM": "倉管盤點", "QC_RESUBMIT": "品保重送",
        "ENG_RESUBMIT": "工程重送", "CLOSE": "結案",
    }
    apv_data = [["時間", "人員", "動作", "意見"]]
    for apv in (form.approvals or []):
        apv_data.append([
            apv.created_at.strftime("%m/%d %H:%M") if apv.created_at else "—",
            apv.approver.display_name if apv.approver else "—",
            action_map.get(apv.action, apv.action),
            apv.comment or "",
        ])
    if len(apv_data) == 1:
        apv_data.append(["—", "—", "—", "尚無記錄"])
    t3 = Table([[P(c, cell_style) for c in r] for r in apv_data],
               colWidths=[2.5*cm, 3*cm, 2.5*cm, 8*cm])
    t3.setStyle(tbl_style("#374151"))
    story.append(t3)

    # ── 附件 ──────────────────────────────────────
    if form.documents:
        story.append(P("▌ 附件", h2))
        img_exts = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"}
        for doc in form.documents:
            doc_path = Path(UPLOAD_BASE) / doc.filename
            ext = Path(doc.original_name).suffix.lower()
            story.append(P(f"📎 {doc.original_name}（{doc.category or '附件'}）", normal))
            if ext in img_exts and doc_path.exists():
                try:
                    from PIL import Image as PILImage
                    with PILImage.open(doc_path) as pimg:
                        w, h = pimg.size
                    # 縮放至最大 14cm 寬
                    max_w = 14 * cm
                    scale = min(max_w / w, (10*cm) / h, 1.0)
                    story.append(RLImage(str(doc_path), width=w*scale, height=h*scale))
                except Exception as e:
                    story.append(P(f"  ⚠ 圖片無法嵌入：{e}", small))
            elif ext == ".pdf" and doc_path.exists():
                story.append(P(f"  （PDF 附件請參閱原始檔案：{doc.filename}）", small))
            story.append(Spacer(1, 0.2*cm))

    # ── 頁腳 ──────────────────────────────────────
    story.append(Spacer(1, 1*cm))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.grey))
    story.append(P(f"本文件由鴻騰電子 PCN/ECN 系統自動產生｜{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", small))

    doc.build(story)
    return buf.getvalue()


def save_cc_pdf(form, inventory_rows=None) -> str:
    """生成 CC PDF 並儲存，回傳儲存路徑"""
    try:
        pdf_bytes = generate_cc_pdf(form, inventory_rows)
        out_dir = Path(UPLOAD_BASE) / form.form_id
        out_dir.mkdir(parents=True, exist_ok=True)
        out_path = out_dir / f"cc_package_{form.form_id}.pdf"
        out_path.write_bytes(pdf_bytes)
        logger.info(f"[PDF] CC Package 已生成：{out_path}")
        return str(out_path)
    except Exception as e:
        logger.error(f"[PDF] 生成失敗：{e}")
        return ""
