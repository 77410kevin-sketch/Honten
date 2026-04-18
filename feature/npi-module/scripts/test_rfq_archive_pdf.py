"""測試：把一張 RFQ 單打包成單一 PDF（客戶資訊 + 供應商派發 + 成本試算 + 客戶報價單）。

用法：
    python scripts/test_rfq_archive_pdf.py RFQ-20260418-013
    → 產出 /tmp/sample_rfq_archive.pdf
"""
import os
import sys
import json
import sqlite3
from datetime import datetime

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm, cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak,
    KeepTogether,
)

# ── 字型：優先使用 macOS 系統內建繁中 TTC，避免缺字 ──
_FONT_CANDIDATES = [
    ("/System/Library/Fonts/STHeiti Light.ttc", 1),    # 繁中黑體
    ("/System/Library/Fonts/STHeiti Medium.ttc", 1),
    ("/System/Library/Fonts/Supplemental/Songti.ttc", 1),
    ("/System/Library/Fonts/PingFang.ttc", 0),
]
_FONT_CANDIDATES_B = [
    ("/System/Library/Fonts/STHeiti Medium.ttc", 1),
    ("/System/Library/Fonts/STHeiti Light.ttc", 1),
]

def _register_first(name, candidates):
    for path, idx in candidates:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont(name, path, subfontIndex=idx))
                return True
            except Exception:
                continue
    return False

FONT = "HTFont"
FONT_B = "HTFontB"
if not _register_first(FONT, _FONT_CANDIDATES):
    # 退回 CID — 可能會有缺字
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    pdfmetrics.registerFont(UnicodeCIDFont("MSung-Light"))
    FONT = "MSung-Light"
if not _register_first(FONT_B, _FONT_CANDIDATES_B):
    FONT_B = FONT

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB = os.path.join(ROOT, "data", "demo.db")


def style(name, size=10, bold=False, align=0, leading=None, color=colors.black):
    return ParagraphStyle(
        name=name, fontName=(FONT_B if bold else FONT), fontSize=size,
        leading=leading or size * 1.25, alignment=align, textColor=color,
        wordWrap="CJK",
    )


def P(text, size=8.5, bold=False, align=0, color=colors.black):
    """Table cell 用：可自動換行的 Paragraph"""
    if text is None or text == "":
        text = "—"
    return Paragraph(str(text), style(f"p{size}{bold}{align}", size, bold, align, color=color))


def load_form(form_id):
    con = sqlite3.connect(DB)
    con.row_factory = sqlite3.Row
    f = con.execute(
        "SELECT * FROM npi_forms WHERE form_id=?", (form_id,)
    ).fetchone()
    if not f:
        print(f"找不到 {form_id}")
        sys.exit(1)
    invites = con.execute(
        "SELECT i.*, s.name AS supplier_name, s.contact AS supplier_contact "
        "FROM npi_supplier_invites i LEFT JOIN suppliers s ON s.id=i.supplier_id "
        "WHERE i.form_id_fk=?",
        (f["id"],),
    ).fetchall()
    creator = con.execute(
        "SELECT display_name FROM users WHERE id=?", (f["created_by"],)
    ).fetchone()
    bu_head = con.execute(
        "SELECT display_name FROM users WHERE role='BU' AND bu=? LIMIT 1",
        (f["bu"],),
    ).fetchone() if f["bu"] else None
    con.close()
    return f, invites, creator, bu_head


def fmt_num(v, dash="—"):
    if v is None or v == "":
        return dash
    try:
        return f"{float(v):,.2f}"
    except Exception:
        return str(v)


def section_title(text):
    return Paragraph(text, style("sec", 14, True, color=colors.HexColor("#0d6efd")))


def build_header_table(form, creator, bu_head):
    def kv(label, value):
        return [P(label, 9, True), P(value, 9)]
    data = [
        kv("客戶", form["customer_name"]) + kv("報價單號", form["form_id"]),
        kv("聯絡人", form["customer_contact"]) + kv("日期", datetime.utcnow().strftime("%Y-%m-%d")),
        kv("Email", form["customer_email"]) + kv("事業部", form["bu"]),
        kv("產品", form["product_name"]) + kv("型號", form["product_model"]),
        kv("建單業務", creator["display_name"] if creator else "—") +
          kv("業務主管", bu_head["display_name"] if bu_head else "—"),
        kv("規格摘要", form["spec_summary"]) + kv("業務補充", form["sales_note"]),
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


def build_invites_table(invites):
    header = [P(x, 9, True, align=1) for x in
              ["#", "供應商", "製程", "材質", "數量(MOQ)", "單價", "模治具", "交期(天)", "選用"]]
    data = [header]
    for idx, inv in enumerate(invites, 1):
        data.append([
            P(str(idx), 8.5, align=1),
            P(inv["supplier_name"], 8.5),
            P(inv["process_name"], 8.5, align=1),
            P(inv["material"], 8.5, align=1),
            P(str(inv["qty"]) if inv["qty"] else "—", 8.5, align=1),
            P(fmt_num(inv["quote_amount"]), 8.5, align=2),
            P(fmt_num(inv["tooling_cost"]), 8.5, align=2),
            P(str(inv["lead_time_days"]) if inv["lead_time_days"] else "—", 8.5, align=1),
            P("✓" if inv["is_selected"] else "", 8.5, align=1),
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


def build_cost_table(q):
    """q = quote_cost_data dict"""
    cols = q.get("columns", [])
    rows = q.get("rows", [])
    defect = q.get("defect_rate", 0)
    oh = q.get("overhead_rate", 0)
    qa = q.get("qa_ship_rate", 0)
    def L(text, bold=False, align=0):  # 左欄 label
        return P(text, 8.5, bold, align)
    def N(v):  # 右欄 number
        return P(fmt_num(v), 8.5, align=2)

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
        P(f"{(c.get('profit_rate') or 0)*100:.1f}% / {fmt_num(c.get('profit_amount'))}",
          8.5, align=2) for c in cols
    ])
    data.append([L("最終報價單價", True)] + [N(c.get("quote")) for c in cols])
    data.append([L("模治具成本")] + [N(c.get("tooling_cost")) for c in cols])
    data.append([L("模治具利潤 (每欄 %)")] + [
        P(f"{(c.get('tooling_profit_rate') or 0)*100:.1f}% / {fmt_num(c.get('tooling_profit_amount'))}",
          8.5, align=2) for c in cols
    ])
    data.append([L("模治具報價（含利潤）", True)] + [N(c.get("tooling_quote")) for c in cols])

    ncols = len(cols)
    label_w = 55 * mm
    remaining = 180 * mm - label_w
    col_w = [label_w] + [remaining / max(ncols, 1) for _ in range(ncols)]
    t = Table(data, colWidths=col_w)

    body_start = 1
    body_end = body_start + len(rows) - 1
    sub_row = body_end + 1
    final_row = sub_row + 7
    tooling_quote_row = final_row + 3

    style_cmds = [
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
    ]
    t.setStyle(TableStyle(style_cmds))
    return t


def build_quote_sheet(form, q, bu_head, creator):
    """實際報價單樣式 — 單價 / 總價 / 模治具"""
    cols = q.get("columns", []) if q else []
    invites = []
    shared_mat = ""
    shared_qty = ""
    elems = []

    title_center = style("qt_title", 16, True, align=1)
    elems.append(Paragraph("鴻騰電子股份有限公司 QUOTATION", title_center))
    elems.append(Paragraph("地址：新北市樹林區三俊街154號 ｜ 電話：(02) 2688-2150",
                            style("qt_addr", 9, align=1, color=colors.grey)))
    elems.append(Spacer(1, 6 * mm))

    # 客戶資訊
    def kv(label, value):
        return [P(label, 9, True), P(value, 9)]
    info = [
        kv("客戶", form["customer_name"]) + kv("報價單號", form["form_id"]),
        kv("收件人", form["customer_contact"]) + kv("日期", datetime.utcnow().strftime("%Y-%m-%d")),
        kv("Email", form["customer_email"]) + [P("", 9), P("", 9)],
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

    # 報價明細表
    elems.append(Paragraph("報價明細", style("qt_sec", 11, True)))
    elems.append(Spacer(1, 2 * mm))
    header = [P(x, 9, True, align=1) for x in
              ["#", "機種 / 料號", "產品說明", "材質", "數量", "單價 (NTD)", "總價 (NTD)"]]
    rows = [header]
    total = 0.0
    for idx, c in enumerate(cols, 1):
        mat = shared_mat or "—"
        qty = shared_qty or 1
        quote = c.get("quote") or 0
        line_total = quote * (qty if isinstance(qty, (int, float)) else 1)
        total += line_total
        rows.append([
            P(str(idx), 8.5, align=1),
            P(c.get("label"), 8.5, True),
            P(form["product_name"], 8.5),
            P(mat, 8.5, align=1),
            P(str(qty), 8.5, align=2),
            P(fmt_num(quote), 8.5, align=2),
            P(fmt_num(line_total), 8.5, True, align=2),
        ])
        t_quote = c.get("tooling_quote") or c.get("tooling_cost") or 0
        if t_quote:
            rows.append([
                P(f"{idx}-T", 8.5, align=1),
                P(f"{c.get('label') or '—'}（模治具費用）", 8.5),
                P("一次性模治具攤提 / 買斷", 8.5),
                P("—", 8.5, align=1),
                P("1", 8.5, align=2),
                P(fmt_num(t_quote), 8.5, align=2),
                P(fmt_num(t_quote), 8.5, True, align=2),
            ])
            total += t_quote
    rows.append([P("", 8.5), P("", 8.5), P("", 8.5), P("", 8.5), P("", 8.5),
                 P("總計 (NTD)", 9, True, align=2),
                 P(fmt_num(total), 9, True, align=2)])
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

    # 條款
    terms = [
        "1. 本報價金額以新台幣計算，未含營業稅。",
        "2. 報價有效期：自報價日起 30 日內。",
        "3. 付款條件：月結 60 天，或依雙方協議。",
        "4. 交貨地點：貴公司指定地點，海外運費另議。",
        "5. 若有任何規格／數量調整，請書面通知，本公司將重新估算報價。",
        "6. 模具費用（如適用）另議，由客戶分期攤提或一次性買斷。",
    ]
    elems.append(Paragraph("報價條款", style("t_title", 10, True)))
    for line in terms:
        elems.append(Paragraph(line, style("t_li", 9, leading=14)))
    elems.append(Spacer(1, 6 * mm))

    # 簽核
    sig = [["業務窗口", "業務主管", "客戶確認"],
           [creator["display_name"] if creator else "—",
            bu_head["display_name"] if bu_head else "—", ""]]
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


def build_pdf(form_id, out_path):
    form, invites, creator, bu_head = load_form(form_id)
    quote_data = {}
    if form["quote_cost_data"]:
        try:
            quote_data = json.loads(form["quote_cost_data"])
        except Exception:
            pass

    doc = SimpleDocTemplate(
        out_path, pagesize=A4,
        leftMargin=15 * mm, rightMargin=15 * mm,
        topMargin=15 * mm, bottomMargin=15 * mm,
        title=f"{form_id} RFQ 歸檔",
    )
    story = []

    # 頁 1：客戶/產品 + 供應商派發 + 業務成本試算
    story.append(Paragraph(f"RFQ 報價歸檔 — {form_id}",
                           style("h", 14, True, color=colors.HexColor("#0d6efd"))))
    story.append(Paragraph(f"產出日期：{datetime.utcnow().strftime('%Y-%m-%d %H:%M')}",
                           style("sub", 8, color=colors.grey)))
    story.append(Spacer(1, 3 * mm))

    story.append(section_title("① 客戶 / 產品資訊"))
    story.append(Spacer(1, 2 * mm))
    story.append(build_header_table(form, creator, bu_head))
    story.append(Spacer(1, 4 * mm))

    story.append(section_title("② 供應商派發與回報報價"))
    story.append(Spacer(1, 2 * mm))
    if invites:
        story.append(build_invites_table(invites))
    else:
        story.append(Paragraph("（無派發記錄）", style("na", 9, color=colors.grey)))
    story.append(Spacer(1, 4 * mm))

    story.append(section_title("③ 業務成本試算表"))
    story.append(Spacer(1, 2 * mm))
    if quote_data.get("columns"):
        story.append(build_cost_table(quote_data))
    else:
        story.append(Paragraph("（無成本試算資料）", style("na", 9, color=colors.grey)))
    story.append(PageBreak())

    # 頁 2：實際客戶報價單
    story.append(section_title("④ 實際客戶報價單"))
    story.append(Spacer(1, 3 * mm))
    story.extend(build_quote_sheet(form, quote_data, bu_head, creator))

    doc.build(story)
    print(f"✅ 產出 PDF：{out_path}")


def _safe_filename(text):
    """把字串轉為安全檔名（移除 / \\ : * ? " < > | 等）"""
    import re
    cleaned = re.sub(r'[\\/:*?"<>|]+', "_", str(text or "未命名")).strip()
    return cleaned or "未命名"


def archive_filename(form_id):
    """REF-{第一個機種label}.pdf（若抓不到則 REF-{form_id}.pdf）"""
    con = sqlite3.connect(DB)
    con.row_factory = sqlite3.Row
    row = con.execute("SELECT quote_cost_data FROM npi_forms WHERE form_id=?", (form_id,)).fetchone()
    con.close()
    first_label = form_id
    if row and row["quote_cost_data"]:
        try:
            q = json.loads(row["quote_cost_data"])
            cols = q.get("columns") or []
            if cols and cols[0].get("label"):
                first_label = cols[0]["label"]
        except Exception:
            pass
    return f"REF-{_safe_filename(first_label)}.pdf"


if __name__ == "__main__":
    fid = sys.argv[1] if len(sys.argv) > 1 else "RFQ-20260418-013"
    if len(sys.argv) > 2:
        out = sys.argv[2]
    else:
        out = os.path.join("/tmp", archive_filename(fid))
    build_pdf(fid, out)
