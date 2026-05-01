from sqlalchemy import Column, Integer, String, Enum, DateTime, Text, ForeignKey, Float, Boolean
from sqlalchemy.orm import relationship
from datetime import datetime
import enum
from app.database import Base


class QCExceptionStatus(str, enum.Enum):
    DRAFT                = "DRAFT"                # 品保草稿（IPC 異常資訊填寫中）
    PENDING_DISPOSITION  = "PENDING_DISPOSITION"  # 待品保下處理判斷（退貨/實驗/特採）
    PENDING_RCA          = "PENDING_RCA"          # 待 Mail 通知 + 根因分析
    PENDING_IMPROVEMENT  = "PENDING_IMPROVEMENT"  # 待制定長期改善方案（圖面/SOP/SIP/ECN）
    LINKED_ECN           = "LINKED_ECN"           # 已開 ECN，等對應 ECN 結案
    CLOSED               = "CLOSED"               # 結案


class QCDisposition(str, enum.Enum):
    RETURN_TO_SUPPLIER   = "RETURN_TO_SUPPLIER"    # 退貨
    LAB_TEST             = "LAB_TEST"              # 實驗測試（舊版相容，新版併入 SA Rework）
    SPECIAL_ACCEPT       = "SPECIAL_ACCEPT"        # 特採允收（含 NO_ACTION / SORTING / REWORK 子類）
    HORIZONTAL_EXPANSION = "HORIZONTAL_EXPANSION"  # 橫向展開（盤點 + 判定）


class QCExceptionStage(str, enum.Enum):
    """異常發生工段"""
    IQC        = "IQC"        # 進料檢驗
    IPQC       = "IPQC"       # 製程檢驗
    OQC        = "OQC"        # 出貨檢驗
    INSPECTION = "INSPECTION" # 品檢
    LASER      = "LASER"      # 雷雕
    CNC        = "CNC"        # CNC
    ASSEMBLY   = "ASSEMBLY"   # 組裝
    OTHER      = "OTHER"


class QCDocType(str, enum.Enum):
    """單號類型 — 進貨單號 / 製程單號 / 出貨 D/C"""
    RECEIVE  = "RECEIVE"   # 進貨單號
    PROCESS  = "PROCESS"   # 製程單號
    SHIP_DC  = "SHIP_DC"   # 出貨 D/C


class QCEventDateType(str, enum.Enum):
    """日期類型 — 進貨/生產/出貨/客訴"""
    RECEIVE   = "RECEIVE"   # 進貨日期
    PRODUCE   = "PRODUCE"   # 生產日期
    SHIP      = "SHIP"      # 出貨日期
    COMPLAINT = "COMPLAINT" # 客訴日期


class QCSourceType(str, enum.Enum):
    """異常來源類型 — 廠商 / 客戶 / 廠內"""
    SUPPLIER = "SUPPLIER"   # 廠商（外部供應商）
    CUSTOMER = "CUSTOMER"   # 客戶（客訴端）
    INTERNAL = "INTERNAL"   # 廠內（內部工站/單位）


class QCException(Base):
    __tablename__ = "qc_exceptions"

    id                = Column(Integer, primary_key=True, index=True)
    form_id           = Column(String(30), unique=True, nullable=False)  # NCR-YYYYMMDD-NNN
    status            = Column(Enum(QCExceptionStatus),
                               default=QCExceptionStatus.DRAFT, nullable=False)

    # ── IPC 異常資訊（業務首頁示例）─────────────────
    part_no           = Column(String(80),  nullable=False)   # 品號 KS04P
    # 單號（類型 + 號碼）
    doc_type          = Column(Enum(QCDocType),
                               default=QCDocType.RECEIVE, nullable=True)
    receive_doc_no    = Column(String(80),  nullable=True)    # 單號號碼
    lot_no            = Column(String(80),  nullable=True)    # 批號（保留 DB 欄位但 UI 已隱藏）
    # 日期（類型 + 值）
    event_date_type   = Column(Enum(QCEventDateType),
                               default=QCEventDateType.RECEIVE, nullable=True)
    receive_date      = Column(String(20),  nullable=True)    # 日期值
    stage             = Column(Enum(QCExceptionStage),
                               default=QCExceptionStage.IQC, nullable=False)
    source_type       = Column(Enum(QCSourceType),
                               default=QCSourceType.SUPPLIER, nullable=True)  # 廠商/客戶/廠內
    supplier_name     = Column(String(120), nullable=True)    # 名稱（廠商 展倚 / 客戶 景利 / 廠內 CNC 課）
    receive_qty       = Column(Integer,     nullable=True)    # 數量
    defect_cause      = Column(Text,        nullable=False)   # 主要異常原因（向下相容，取多列 list 第一筆）
    measurement_data  = Column(Text,        nullable=True)    # （已棄用，UI 取消，DB 保留）
    defect_qty        = Column(Integer,     nullable=True)    # 不良總數（多列加總）
    sample_qty        = Column(Integer,     nullable=True)    # 抽樣總數（多列加總）
    defect_rate       = Column(Float,       nullable=True)    # 整批不良率（總不良/總抽樣）
    defect_items_json = Column(Text,        nullable=True)    # 異常多列 JSON list:
    # [{"cause":"總長過長","types":["EXTERIOR","DIMENSION"],"defect_qty":40,"sample_qty":315}]

    # ── 品保處理判斷 ────────────────────────────────
    disposition       = Column(Enum(QCDisposition), nullable=True)   # 主要處理（向下相容）
    dispositions_json = Column(Text, nullable=True)                  # 多選 JSON list（向下相容）
    actions_json      = Column(Text, nullable=True)                  # v3 多卡片 JSON list
    # actions_json 結構：
    # [{"id":"uuid","type":"RTS|B1|B2|B3|B4|B5|HE",
    #   "fields":{...各 type 自己的 fields...},
    #   "created_at":"...", "sent_at":"...", "sent_by":1,
    #   "replies":[{"unit":"PURCHASE","at":"...","by":1,"note":"已收到"}]}]
    disposition_note  = Column(Text, nullable=True)                  # 處理判斷說明
    disposition_at    = Column(DateTime, nullable=True)
    disposition_by    = Column(Integer, ForeignKey("users.id"), nullable=True)

    # 立即處理 — 通知信（對象可為「供應商」或「工站」）
    rts_target_type        = Column(String(20), nullable=True)   # SUPPLIER | STATION
    rts_replenish_note     = Column(Text, nullable=True)         # 補貨資訊請求（給採購/生管）
    rts_pickup_required    = Column(Boolean, default=False)      # A 退貨：是否需安排司機載回
    rts_pickup_note        = Column(Text, nullable=True)         # 司機載回備註（地點/聯絡人/時間）
    supplier_mail_to       = Column(Text, nullable=True)
    supplier_mail_cc       = Column(Text, nullable=True)
    supplier_mail_subject  = Column(String(200), nullable=True)
    supplier_mail_body     = Column(Text, nullable=True)
    supplier_mail_sent_at  = Column(DateTime, nullable=True)

    # 實驗測試（舊欄位保留，新版併入 SA Rework 流程）
    lab_test_qty          = Column(Integer, nullable=True)
    lab_test_conditions   = Column(Text, nullable=True)
    lab_test_due_date     = Column(String(20), nullable=True)
    linked_sample_request_no = Column(String(50), nullable=True)

    # 特採允收 — 子類別（v2 多選 JSON list；v1 單一 sa_subtype 保留向下相容）
    sa_subtype          = Column(String(20), nullable=True)       # 舊：NO_ACTION | SORTING | REWORK
    sa_subtypes_json    = Column(Text, nullable=True)             # 新：["SORTING","REWORK"] / ["NO_ACTION"]
    sa_need_sorting     = Column(Boolean, default=False)
    sa_need_rework      = Column(Boolean, default=False)
    # 由品保填寫
    sa_defect_handling  = Column(Text, nullable=True)              # 不良品處理方式（品保填）
    # 由生管填寫（送生管按鈕觸發後可填）
    sa_station          = Column(String(50), nullable=True)        # 執行站別/單位（生管填）
    sa_sent_to_prod_at  = Column(DateTime, nullable=True)          # 已送生管時間
    # Sorting 結果回填
    sa_sorting_pass_qty = Column(Integer, nullable=True)
    sa_sorting_fail_qty = Column(Integer, nullable=True)
    sa_sorting_filled_at = Column(DateTime, nullable=True)
    # Rework 內容
    sa_rework_note      = Column(Text, nullable=True)
    sa_rework_result    = Column(Text, nullable=True)              # rework + 樣品測試 結果回報
    sa_rework_filled_at = Column(DateTime, nullable=True)
    sa_rework_pass_qty  = Column(Integer, nullable=True)           # Rework 後良品數
    sa_rework_fail_qty  = Column(Integer, nullable=True)           # Rework 後不良品數
    sa_rework_defect_handling = Column(Text, nullable=True)        # Rework 後不良品處理方式（取代 sa_defect_handling UI）
    # 客戶端執行（B4/B5）— 需計算工時與人力（業務協同）
    sa_cust_sorting_hours    = Column(Float, nullable=True)
    sa_cust_sorting_workers  = Column(Integer, nullable=True)
    sa_cust_rework_hours     = Column(Float, nullable=True)
    sa_cust_rework_workers   = Column(Integer, nullable=True)
    sa_cust_note             = Column(Text, nullable=True)         # 客戶端地點/聯絡人/排程

    # 橫向展開 — v1 單列（向下相容） + v2 多列盤點單 JSON
    he_customer_qty     = Column(Integer, nullable=True)
    he_inhouse_qty      = Column(Integer, nullable=True)
    he_supplier_qty     = Column(Integer, nullable=True)
    he_decision         = Column(Text, nullable=True)
    he_inventory_data   = Column(Text, nullable=True)              # JSON list:
    # [{"part_no":"KS04P","customer_qty":5000,"inhouse_qty":12000,"supplier_qty":8000,"decision":"..."}]

    # ── Mail 通知 + 根因分析 ────────────────────────
    notify_mail_to    = Column(Text, nullable=True)            # CSV 收件人 email
    notify_mail_cc    = Column(Text, nullable=True)            # CSV cc
    notify_sent_at    = Column(DateTime, nullable=True)
    root_cause        = Column(Text, nullable=True)            # 根因分析

    # ── 長期改善方案 + ECN 綁定 ─────────────────────
    need_drawing_rev  = Column(Boolean, default=False)         # 需修訂圖面
    need_sop_rev      = Column(Boolean, default=False)         # 需修訂 SOP
    need_sip_rev      = Column(Boolean, default=False)         # 需修訂 SIP
    improvement_plan  = Column(Text, nullable=True)            # 長期改善內容
    linked_ecn_form_id = Column(Integer, ForeignKey("pcn_forms.id"), nullable=True)  # 對應 ECN

    # ── 系統欄位 ────────────────────────────────────
    reject_to         = Column(String(50), nullable=True)
    created_by        = Column(Integer, ForeignKey("users.id"), nullable=False)
    assigned_qc_id    = Column(Integer, ForeignKey("users.id"), nullable=True)
    created_at        = Column(DateTime, default=datetime.utcnow)
    updated_at        = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    creator      = relationship("User", foreign_keys=[created_by])
    assigned_qc  = relationship("User", foreign_keys=[assigned_qc_id])
    dispositioner = relationship("User", foreign_keys=[disposition_by])
    linked_ecn   = relationship("PCNForm", foreign_keys=[linked_ecn_form_id])

    documents    = relationship("QCExceptionDocument", back_populates="form",
                                cascade="all, delete-orphan",
                                order_by="QCExceptionDocument.uploaded_at")
    approvals    = relationship("QCExceptionApproval", back_populates="form",
                                cascade="all, delete-orphan",
                                order_by="QCExceptionApproval.created_at")


class QCExceptionDocument(Base):
    __tablename__ = "qc_exception_documents"

    id            = Column(Integer, primary_key=True, index=True)
    form_id_fk    = Column(Integer, ForeignKey("qc_exceptions.id"), nullable=False)
    filename      = Column(String(255), nullable=False)
    original_name = Column(String(255), nullable=False)
    category      = Column(String(50), nullable=True)   # 異常照片/實驗報告/圖面/其它
    uploaded_by   = Column(Integer, ForeignKey("users.id"), nullable=True)
    uploaded_at   = Column(DateTime, default=datetime.utcnow)

    form     = relationship("QCException", back_populates="documents")
    uploader = relationship("User", foreign_keys=[uploaded_by])


class QCExceptionApproval(Base):
    __tablename__ = "qc_exception_approvals"

    id            = Column(Integer, primary_key=True, index=True)
    form_id_fk    = Column(Integer, ForeignKey("qc_exceptions.id"), nullable=False)
    approver_id   = Column(Integer, ForeignKey("users.id"), nullable=False)
    action        = Column(String(30), nullable=False)
    comment       = Column(Text, nullable=True)
    reject_target = Column(String(50), nullable=True)
    from_status   = Column(String(50), nullable=True)
    to_status     = Column(String(50), nullable=True)
    created_at    = Column(DateTime, default=datetime.utcnow)

    form     = relationship("QCException", back_populates="approvals")
    approver = relationship("User", foreign_keys=[approver_id])
