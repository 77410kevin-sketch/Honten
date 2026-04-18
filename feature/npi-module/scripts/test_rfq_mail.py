"""測試 RFQ 詢價信寄送

用法：
    # 在專案根目錄執行（需先填好 .env 內的 SMTP_* 變數）
    python scripts/test_rfq_mail.py <收件 email>

    # 例：
    python scripts/test_rfq_mail.py 77410kevin@company.com

會產生一組假 NPI 單資料，使用真實的 notify_quotes_dispatched 邏輯
（變數替換、附件打包）透過 SMTP 寄出，用來確認內文與附件格式。
"""
import os
import sys
import tempfile
from types import SimpleNamespace
from dotenv import load_dotenv

# 載入專案根目錄 .env
load_dotenv(dotenv_path=os.path.join(os.path.dirname(__file__), "..", ".env"))

# 測試前置檢查
required = ["SMTP_HOST", "SMTP_USER", "SMTP_PASSWORD", "SMTP_FROM"]
missing = [k for k in required if not os.getenv(k) or "請填" in os.getenv(k, "")]
if missing:
    print(f"❌ .env 以下變數尚未填寫：{', '.join(missing)}")
    print(f"   請編輯 .env 後重試。")
    sys.exit(1)

if len(sys.argv) < 2:
    print("用法：python scripts/test_rfq_mail.py <收件 email>")
    sys.exit(1)

recipient = sys.argv[1]

# 確保能 import app.services
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from app.services.npi_notification import _render_rfq_body, _DEFAULT_RFQ_TEMPLATE, _send_mail

# 準備測試用的假圖檔（PNG 1x1）
tmp_dir = tempfile.mkdtemp()
sample_attach = os.path.join(tmp_dir, "IMG_3366.PNG")
# 最小合法 PNG（1x1 紅點）
png_bytes = bytes.fromhex(
    "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4"
    "890000000d49444154789c6364f8cf000000030001014efec24f0000000049454e44ae426082"
)
with open(sample_attach, "wb") as f:
    f.write(png_bytes)

# 建立假的 form / invite / supplier 物件（SimpleNamespace 就夠用）
supplier = SimpleNamespace(
    name="久盛塑膠",
    contact="陳經理",
    email=recipient,  # ← 測試時寄給自己
)
invite = SimpleNamespace(
    process_name="沖壓",
    drawing=SimpleNamespace(original_name="IMG_3366.PNG"),
    drawing_doc_id=1,
    material="SUS304",
    qty=1000,
    first_sent_at=None,
    supplier_id=1,
)
form = SimpleNamespace(
    form_id="NPI-TEST-001",
    customer_name="景利（Jingli）",
    product_name="BBU 電池外殼",
    product_model="HT-BBU-5.5K",
    rfq_due_date="2026-04-30",
    eng_process_note=None,  # None 會使用預設模板
    form_id_fk=0,
    id=0,
)

# 使用系統預設模板（同使用者在 modal 看到的預設）
template = _DEFAULT_RFQ_TEMPLATE
body = _render_rfq_body(
    template,
    form=form,
    invite=invite,
    supplier=supplier,
    material=invite.material,
    moq=invite.qty,
)

subject = f"【鴻騰電子 RFQ 詢價（測試）】{form.form_id} - {form.product_name} / {invite.process_name}"

print("=" * 60)
print(f"寄件人：{os.getenv('SMTP_FROM_NAME')} <{os.getenv('SMTP_FROM')}>")
print(f"收件人：{recipient}")
print(f"主旨：{subject}")
print(f"附件：{sample_attach}")
print("─" * 60)
print("內文：")
print(body)
print("=" * 60)
print("開始寄送...")

try:
    _send_mail(recipient, subject, body, attachments=[sample_attach])
    print("✅ 寄送完成，請到收件匣查看（可能進垃圾信箱）")
except Exception as e:
    print(f"❌ 寄送失敗：{type(e).__name__}: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)
