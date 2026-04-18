from sqlalchemy import Column, Integer, String, Boolean, DateTime, Text
from datetime import datetime
from app.database import Base


class Customer(Base):
    """客戶主檔。可手動建立，或從 ERP 同步（以 erp_code 去重）。"""
    __tablename__ = "customers"

    id          = Column(Integer, primary_key=True, index=True)
    erp_code    = Column(String(50), nullable=True, unique=True, index=True)   # ERP 代碼（從 ERP 同步時填）
    name        = Column(String(200), nullable=False, index=True)
    contact     = Column(String(100), nullable=True)
    email       = Column(String(200), nullable=True)
    phone       = Column(String(50), nullable=True)
    address     = Column(String(300), nullable=True)
    bu          = Column(String(50), nullable=True)   # 儲能事業部 / 消費性事業部
    tax_id      = Column(String(30), nullable=True)
    memo        = Column(Text, nullable=True)
    is_active   = Column(Boolean, default=True)
    source      = Column(String(20), default="manual")  # manual / erp
    created_at  = Column(DateTime, default=datetime.utcnow)
    updated_at  = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
