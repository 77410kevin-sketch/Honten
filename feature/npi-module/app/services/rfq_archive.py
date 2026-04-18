"""RFQ 報價歸檔 PDF 產生器 — 用於結案時自動落 NAS、或業務下載。

PDF 結構（2 頁）：
  P1: ① 客戶/產品 + ② 供應商派發 + ③ 業務成本試算
  P2: ④ 實際客戶報價單（含條款、簽核）

呼叫入口：
  build_archive_pdf(form, invites, quote_data, creator_name, bu_head_name, out_path)
  archive_filename(form_id, quote_data) → "REF-{第一個機種}.pdf"
"""
import os
import re
import json
from datetime import datetime

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak,
)

# ── 字型：macOS TTC；Linux/其他環境退回 reportlab 內建 CID 字型 ──
_FONT_CANDIDATES = [
    ("/System/Library/Fonts/STHeiti Light.ttc", 1),
    ("/System/Library/Fonts/STHeiti Medium.ttc", 1),
    ("/System/Library/Fonts/Supplemental/Songti.ttc", 1),
    ("/System/Library/Fonts/PingFang.ttc", 0),
]
_FONT_CANDIDATES_B = [
    ("/System/Library/Fonts/STHeiti Medium.ttc", 1),
    ("/System/Library/Fonts/STHeiti Light.ttc", 1),
]
_FONTS_READY = False
FONT = "HTFont"
FONT_B = "HTFontB"


def _register_first(name, candidates):
    for path, idx in candidates:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont(name, path, subfontIndex=idx))
                return True
            except Exception:
                continue
    return False


def _ensure_fonts():
    global FONT, FONT_B, _FONTS_READY
    if _FONTS_READY:
        return
    if not _register_first(FONT, _FONT_CANDIDATES):
        from reportlab.pdfbase.cidfonts import UnicodeCIDFont
        pdfmetrics.registerFont(UnicodeCIDFont("MSung-Light"))
        FONT = "MSung-Light"
    if not _register_first(FONT_B, _FONT_CANDIDATES_B):
        FONT_B = FONT
    _FONTS_READY = True


def _style(name, size=10, bold=False, align=0, leading=None, color=colors.black):
    return ParagraphStyle(
        name=name, fontName=(FONT_B if bold else FONT), fontSize=size,
        leading=leading or size * 1.25, alignment=align, textColor=color,
        wordWrap="CJK",
    )


def _P(text, size=8.5, bold=False, align=0, color=colors.black):
    if text is None or text == "":
        text = "—"
    return Paragraph(str(text), _style(f"p{size}{bold}{align}", size, bold, align, color=color))


def _fmt(v, dash="—"):
    if v is None or v == "":
        return dash
    try:
        return f"{float(v):,.2f}"
    except Exception:
        return str(v)


def _section_title(text):
    return Paragraph(text, _style("sec", 14, True, color=colors.HexColor("#0d6efd")))


def _header_table(form, creator_name, bu_head_name):
    def kv(label, value):
        return [_P(label, 9, True), _P(value, 9)]
    data = [
        kv("客戶", form.get("customer_name")) + kv("報價單號", form.get("form_id")),
        kv("聯絡人", form.get("customer_contact")) + kv("日期", datetime.utcnow().strftime("%Y-%m-%d")),
        kv("Email", form.get("customer_email")) + kv("事業部", form.get("bu")),
        kv("產品", form.get("product_name")) + kv("型號", form.get("product_model")),
        kv("建單業務", creator_name or "—") + kv("業務主管", bu_head_name or "—"),
        kv("規格摘要", form.get("spec_summary")) + kv("業務補充", form.get("sales_note")),
    ]
    t = Table(data, colWidths=[22 * mm, 63 * mm, 22 * mm, 63 * mm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#f8f9fa")),
        ("BACKGROUND", (2, 0), (2, -1), colors.HexColor("#f8f9fa")),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    return t


def _invites_table(invites):
    header = [_P(x, 9, True, align=1) for x in
              ["#", "供應商", "製程", "材質", "數量(MOQ)", "單價", "模治具", "交期(天)", "選用"]]
    data = [header]
    for idx, inv in enumerate(invites, 1):
        data.append([
            _P(str(idx), 8.5, align=1),
            _P(inv.get("supplier_name"), 8.5),
            _P(inv.get("process_name"), 8.5, align=1),
            _P(inv.get("material"), 8.5, align=1),
            _P(str(inv.get("qty")) if inv.get("qty") else "—", 8.5, align=1),
            _P(_fmt(inv.get("quote_amount")), 8.5, align=2),
            _P(_fmt(inv.get("tooling_cost")), 8.5, align=2),
            _P(str(inv.get("lead_time_days")) if inv.get("lead_time_days") else "—", 8.5, align=1),
            _P("✓" if inv.get("is_selected") else "", 8.5, align=1),
        ])
    t = Table(data, colWidths=[8*mm, 34*mm, 22*mm, 20*mm, 20*mm, 22*mm, 22*mm, 16*mm, 10*mm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e9ecef")),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    return t


def _cost_table(q):
    cols = q.get("columns", []) or []
    rows = q.get("rows", []) or []
    defect = q.get("defect_rate", 0) or 0
    oh = q.get("overhead_rate", 0) or 0
    qa = q.get("qa_ship_rate", 0) or 0

    def L(text, bold=False, align=0):
        return _P(text, 8.5, bold, align)
    def N(v):
        return _P(_fmt(v), 8.5, align=2)

    header = [L("項目 / 製程", True, 1)] + [
        L(c.get("label") or f"方案{i+1}", True, 1) for i, c in enumerate(cols)
    ]
    data = [header]
    for r in rows:
        row = [L(r.get("process", ""))]
        for p in r.get("prices", []):
            row.append(N(p))
        data.append(row)
    data.append([L("製程成本小計", True)] + [N(c.get("subtotal")) for c in cols])
    data.append([L(f"不良率 ({defect*100:.1f}%)")] + [N(c.get("defect_amount")) for c in cols])
    data.append([L(f"管銷 ({oh*100:.1f}%)")] + [N(c.get("overhead_amount")) for c in cols])
    data.append([L(f"品包運 ({qa*100:.1f}%)")] + [N(c.get("qa_ship_amount")) for c in cols])
    data.append([L("成本合計（不含模治具）", True)] + [N(c.get("cost_total")) for c in cols])
    data.append([L("利潤 (每欄 %)")] + [
        _P(f"{(c.get('profit_rate') or 0)*100:.1f}% / {_fmt(c.get('profit_amount'))}",
          8.5, align=2) for c in cols
    ])
    data.append([L("最終報價單價", True)] + [N(c.get("quote")) for c in cols])
    data.append([L("模治具成本")] + [N(c.get("tooling_cost")) for c in cols])
    data.append([L("模治具利潤 (每欄 %)")] + [
        _P(f"{(c.get('tooling_profit_rate') or 0)*100:.1f}% / {_fmt(c.get('tooling_profit_amount'))}",
          8.5, align=2) for c in cols
    ])
    data.append([L("模治具報價（含利潤）", True)] + [N(c.get("tooling_quote")) for c in cols])

    ncols = len(cols)
    label_w = 55 * mm
    remaining = 180 * mm - label_w
    col_w = [label_w] + [remaining / max(ncols, 1) for _ in range(ncols)]
    t = Table(data, colWidths=col_w)

    body_end = len(rows)
    sub_row = body_end + 1
    final_row = sub_row + 7
    tooling_quote_row = final_row + 3
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e9ecef")),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("BACKGROUND", (0, sub_row), (-1, sub_row), colors.HexColor("#dee2e6")),
        ("BACKGROUND", (0, final_row), (-1, final_row), colors.HexColor("#fff3cd")),
        ("BACKGROUND", (0, final_row+1), (-1, tooling_quote_row-1), colors.HexColor("#cff4fc")),
        ("BACKGROUND", (0, tooling_quote_row), (-1, tooling_quote_row), colors.HexColor("#fff3cd")),
    ]))
    return t


def _quote_sheet(form, q, creator_name, bu_head_name):
    cols = (q.get("columns") if q else None) or []
    elems = []
    elems.append(Paragraph("鴻騰電子股份有限公司 QUOTATION", _style("qt", 16, True, align=1)))
    elems.append(Paragraph("地址：新北市樹林區三俊街154號 ｜ 電話：(02) 2688-2150",
                            _style("addr", 9, align=1, color=colors.grey)))
    elems.append(Spacer(1, 6 * mm))

    def kv(label, value):
        return [_P(label, 9, True), _P(value, 9)]
    info = [
        kv("客戶", form.get("customer_name")) + kv("報價單號", form.get("form_id")),
        kv("收件人", form.get("customer_contact")) + kv("日期", datetime.utcnow().strftime("%Y-%m-%d")),
        kv("Email", form.get("customer_email")) + [_P("", 9), _P("", 9)],
    ]
    info_t = Table(info, colWidths=[22 * mm, 68 * mm, 22 * mm, 62 * mm])
    info_t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#f8f9fa")),
        ("BACKGROUND", (2, 0), (2, -1), colors.HexColor("#f8f9fa")),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    elems.append(info_t)
    elems.append(Spacer(1, 4 * mm))

    elems.append(Paragraph("報價明細", _style("qs", 11, True)))
    elems.append(Spacer(1, 2 * mm))

    header = [_P(x, 9, True, align=1) for x in
              ["#", "機種 / 料號", "產品說明", "材質", "數量", "單價 (NTD)", "總價 (NTD)"]]
    rows = [header]
    total = 0.0
    shared_mat = form.get("_shared_mat") or "—"
    shared_qty = form.get("_shared_qty") or 1
    for idx, c in enumerate(cols, 1):
        quote = c.get("quote") or 0
        line_total = quote * (shared_qty if isinstance(shared_qty, (int, float)) else 1)
        total += line_total
        rows.append([
            _P(str(idx), 8.5, align=1),
            _P(c.get("label"), 8.5, True),
            _P(form.get("product_name"), 8.5),
            _P(shared_mat, 8.5, align=1),
            _P(str(shared_qty), 8.5, align=2),
            _P(_fmt(quote), 8.5, align=2),
            _P(_fmt(line_total), 8.5, True, align=2),
        ])
        t_quote = c.get("tooling_quote") or c.get("tooling_cost") or 0
        if t_quote:
            rows.append([
                _P(f"{idx}-T", 8.5, align=1),
                _P(f"{c.get('label') or '—'}（模治具費用）", 8.5),
                _P("一次性模治具攤提 / 買斷", 8.5),
                _P("—", 8.5, align=1),
                _P("1", 8.5, align=2),
                _P(_fmt(t_quote), 8.5, align=2),
                _P(_fmt(t_quote), 8.5, True, align=2),
            ])
            total += t_quote
    rows.append([_P("", 8.5)] * 5 + [
        _P("總計 (NTD)", 9, True, align=2),
        _P(_fmt(total), 9, True, align=2),
    ])
    qt = Table(rows, colWidths=[8*mm, 32*mm, 48*mm, 18*mm, 18*mm, 24*mm, 28*mm])
    qt.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e9ecef")),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#fff3cd")),
    ]))
    elems.append(qt)
    elems.append(Spacer(1, 5 * mm))

    elems.append(Paragraph("報價條款", _style("tt", 10, True)))
    for line in [
        "1. 本報價金額以新台幣計算，未含營業稅。",
        "2. 報價有效期：自報價日起 30 日內。",
        "3. 付款條件：月結 60 天，或依雙方協議。",
        "4. 交貨地點：貴公司指定地點，海外運費另議。",
        "5. 若有任何規格／數量調整，請書面通知，本公司將重新估算報價。",
        "6. 模具費用（如適用）另議，由客戶分期攤提或一次性買斷。",
    ]:
        elems.append(Paragraph(line, _style("tli", 9, leading=14)))
    elems.append(Spacer(1, 6 * mm))

    sig = [["業務窗口", "業務主管", "客戶確認"],
           [creator_name or "—", bu_head_name or "—", ""]]
    sig_t = Table(sig, colWidths=[58 * mm, 58 * mm, 58 * mm], rowHeights=[8*mm, 18*mm])
    sig_t.setStyle(TableStyle([
        ("FONT", (0, 0), (-1, 0), FONT_B, 10),
        ("FONT", (0, 1), (-1, 1), FONT, 10),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("VALIGN", (0, 1), (-1, 1), "MIDDLE"),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#f8f9fa")),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
    ]))
    elems.append(sig_t)
    return elems


def _safe_filename(text):
    cleaned = re.sub(r'[\\/:*?"<>|]+', "_", str(text or "未命名")).strip()
    return cleaned or "未命名"


def archive_filename(form_id, quote_data):
    """REF-{第一個機種label}.pdf（若抓不到則 REF-{form_id}.pdf）"""
    first_label = form_id
    cols = (quote_data or {}).get("columns") or []
    if cols and cols[0].get("label"):
        first_label = cols[0]["label"]
    return f"REF-{_safe_filename(first_label)}.pdf"


def build_archive_pdf(form, invites, quote_data, creator_name, bu_head_name, out_path):
    """產出歸檔 PDF 到 out_path。

    form: dict with customer_name / customer_contact / customer_email / form_id /
          product_name / product_model / spec_summary / bu / sales_note
    invites: list[dict] with supplier_name / process_name / material / qty /
             quote_amount / tooling_cost / lead_time_days / is_selected
    quote_data: dict (quote_cost_data JSON)
    """
    _ensure_fonts()
    os.makedirs(os.path.dirname(out_path), exist_ok=True) if os.path.dirname(out_path) else None

    doc = SimpleDocTemplate(
        out_path, pagesize=A4,
        leftMargin=15 * mm, rightMargin=15 * mm,
        topMargin=15 * mm, bottomMargin=15 * mm,
        title=f"{form.get('form_id')} RFQ 歸檔",
    )
    story = []

    story.append(Paragraph(f"RFQ 報價歸檔 — {form.get('form_id')}",
                           _style("h", 14, True, color=colors.HexColor("#0d6efd"))))
    story.append(Paragraph(f"產出日期：{datetime.utcnow().strftime('%Y-%m-%d %H:%M')}",
                           _style("sub", 8, color=colors.grey)))
    story.append(Spacer(1, 3 * mm))

    story.append(_section_title("① 客戶 / 產品資訊"))
    story.append(Spacer(1, 2 * mm))
    story.append(_header_table(form, creator_name, bu_head_name))
    story.append(Spacer(1, 4 * mm))

    story.append(_section_title("② 供應商派發與回報報價"))
    story.append(Spacer(1, 2 * mm))
    if invites:
        story.append(_invites_table(invites))
    else:
        story.append(Paragraph("（無派發記錄）", _style("na", 9, color=colors.grey)))
    story.append(Spacer(1, 4 * mm))

    story.append(_section_title("③ 業務成本試算表"))
    story.append(Spacer(1, 2 * mm))
    if (quote_data or {}).get("columns"):
        story.append(_cost_table(quote_data))
    else:
        story.append(Paragraph("（無成本試算資料）", _style("na", 9, color=colors.grey)))
    story.append(PageBreak())

    story.append(_section_title("④ 實際客戶報價單"))
    story.append(Spacer(1, 3 * mm))
    story.extend(_quote_sheet(form, quote_data or {}, creator_name, bu_head_name))

    doc.build(story)
    return out_path
