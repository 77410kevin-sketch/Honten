"""測試合併寄送 — 模擬同一供應商 2 筆派發合併為一封信。"""
import os
import sys
import tempfile
from types import SimpleNamespace
from dotenv import load_dotenv

load_dotenv(dotenv_path=os.path.join(os.path.dirname(__file__), "..", ".env"))

if len(sys.argv) < 2:
    print("用法：python scripts/test_rfq_merge_mail.py <收件 email>")
    sys.exit(1)

recipient = sys.argv[1]

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from app.services.npi_notification import _render_rfq_body, _DEFAULT_RFQ_TEMPLATE, _send_mail

tmp = tempfile.mkdtemp()
png = bytes.fromhex(
    "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4"
    "890000000d49444154789c6364f8cf000000030001014efec24f0000000049454e44ae426082"
)
img1 = os.path.join(tmp, "IMG_3366_1.PNG")
img2 = os.path.join(tmp, "IMG_3366_2.PNG")
open(img1, "wb").write(png)
open(img2, "wb").write(png)

supplier = SimpleNamespace(name="久盛塑膠", contact="張經理", email=recipient)
form = SimpleNamespace(
    form_id="NPI-TEST-MERGE",
    customer_name="景利（Jingli）",
    product_name="BBU 電池外殼",
    product_model="",
    rfq_due_date="2026-04-30",
    eng_process_note=None,
    id=0,
)
# 模擬合併分支（直接複製 notify_quotes_dispatched 的合併邏輯）
template = _DEFAULT_RFQ_TEMPLATE
merged_template = (template
                   .replace("{process}", "（見下列項目）")
                   .replace("{drawing}", "（見下列項目）")
                   .replace("{material}", "（見下列項目）")
                   .replace("{moq}", "（見下列項目）"))
first = SimpleNamespace(
    process_name="沖壓",
    drawing=SimpleNamespace(original_name="IMG_3366_1.PNG"),
    drawing_doc_id=1,
)
intro = _render_rfq_body(merged_template, form=form, invite=first, supplier=supplier,
                         material="（見下列項目）", moq="（見下列項目）")

items = [
    ("沖壓", "IMG_3366_1.PNG", "AL6063", 2000),
    ("鋁擠", "IMG_3366_2.PNG", "SUS304", 6000),
]
lines = ["", "━━━━━━━━━━━━━━━━━━━━━━━", "本次詢價項目：", ""]
for i, (proc, drw, mat, qty) in enumerate(items, 1):
    lines += [
        f"{i}. 製程：{proc}",
        f"   對應圖面：{drw}",
        f"   材質：{mat}　MOQ：{qty}",
        "",
    ]
body = intro + "\n".join(lines)
subject = f"【鴻騰電子 RFQ 詢價】{form.form_id} - {form.product_name}（共 {len(items)} 項製程）"

print("=" * 60)
print(f"寄件人：{os.getenv('SMTP_FROM_NAME')} <{os.getenv('SMTP_FROM')}>")
print(f"收件人：{recipient}")
print(f"主旨：{subject}")
print(f"附件：{img1}, {img2}")
print("─" * 60)
print("內文：")
print(body)
print("=" * 60)

try:
    _send_mail(recipient, subject, body, attachments=[img1, img2])
    print("✅ 寄送完成")
except Exception as e:
    print(f"❌ 寄送失敗：{type(e).__name__}: {e}")
    import traceback; traceback.print_exc()
    sys.exit(1)
