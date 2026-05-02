import enum
from datetime import datetime
from sqlalchemy import (
    Column, Integer, String, Boolean, Enum, DateTime, Text,
    ForeignKey, Index,
)
from sqlalchemy.orm import relationship
from app.database import Base


class ResourceType(str, enum.Enum):
    ROOM = "ROOM"   # 會議室
    CAR  = "CAR"    # 公務車


class EventType(str, enum.Enum):
    LEAVE      = "LEAVE"       # 請假
    CAR        = "CAR"         # 公務車租借
    ROOM       = "ROOM"        # 會議室租借
    OUTING     = "OUTING"      # 業務外出
    ATTENDANCE = "ATTENDANCE"  # 出勤


class EventStatus(str, enum.Enum):
    DRAFT     = "DRAFT"
    PENDING   = "PENDING"   # 請假待主管審核
    APPROVED  = "APPROVED"
    REJECTED  = "REJECTED"
    CANCELLED = "CANCELLED"


class CalendarResource(Base):
    """會議室／公務車主檔"""
    __tablename__ = "calendar_resources"

    id        = Column(Integer, primary_key=True, index=True)
    type      = Column(Enum(ResourceType), nullable=False)
    code      = Column(String(40), unique=True, nullable=False)   # ROOM-1 / CAR-1
    name      = Column(String(80), nullable=False)                # 大會議室 / TOYOTA-3168
    capacity  = Column(Integer, nullable=True)                    # 會議室人數
    location  = Column(String(120), nullable=True)
    notes     = Column(Text, nullable=True)
    active    = Column(Boolean, default=True, nullable=False)
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)


class CalendarEvent(Base):
    """統一事件表（請假/車/室/外出/出勤）"""
    __tablename__ = "calendar_events"

    id          = Column(Integer, primary_key=True, index=True)
    event_type  = Column(Enum(EventType), nullable=False, index=True)
    title       = Column(String(200), nullable=False)
    start_at    = Column(DateTime, nullable=False, index=True)
    end_at      = Column(DateTime, nullable=False, index=True)
    all_day     = Column(Boolean, default=False, nullable=False)

    user_id     = Column(Integer, ForeignKey("users.id"), nullable=False, index=True)
    resource_id = Column(Integer, ForeignKey("calendar_resources.id"), nullable=True, index=True)

    status      = Column(Enum(EventStatus), default=EventStatus.APPROVED, nullable=False, index=True)

    # 請假審核
    leave_type    = Column(String(40), nullable=True)   # ANNUAL/SICK/PERSONAL/...
    approver_id   = Column(Integer, ForeignKey("users.id"), nullable=True)
    approved_at   = Column(DateTime, nullable=True)
    reject_reason = Column(Text, nullable=True)

    # 業務外出
    customer_name = Column(String(120), nullable=True)
    contact_phone = Column(String(40), nullable=True)

    notes       = Column(Text, nullable=True)

    # LINE 來源紀錄
    raw_text    = Column(Text, nullable=True)
    parsed_json = Column(Text, nullable=True)

    created_at  = Column(DateTime, default=datetime.utcnow, nullable=False)
    updated_at  = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow, nullable=False)

    user      = relationship("User", foreign_keys=[user_id])
    approver  = relationship("User", foreign_keys=[approver_id])
    resource  = relationship("CalendarResource")


# 對 (resource_id, start_at, end_at) 加 composite index 加速衝突查詢
Index("ix_calevent_resource_time", CalendarEvent.resource_id, CalendarEvent.start_at, CalendarEvent.end_at)


class LeaveType(Base):
    """假別主檔"""
    __tablename__ = "calendar_leave_types"

    id   = Column(Integer, primary_key=True, index=True)
    code = Column(String(20), unique=True, nullable=False)   # ANNUAL/SICK/PERSONAL...
    name = Column(String(40), nullable=False)
    is_paid = Column(Boolean, default=True, nullable=False)
    max_days_per_year = Column(Integer, nullable=True)
    color = Column(String(20), nullable=True)


class LeaveBalance(Base):
    """個人假別餘額（年度）"""
    __tablename__ = "calendar_leave_balances"

    id          = Column(Integer, primary_key=True, index=True)
    user_id     = Column(Integer, ForeignKey("users.id"), nullable=False, index=True)
    year        = Column(Integer, nullable=False, index=True)
    leave_code  = Column(String(20), nullable=False)
    total_days  = Column(Integer, nullable=False, default=0)
    used_days   = Column(Integer, nullable=False, default=0)

    user = relationship("User")


Index("ix_leave_balance_user_year", LeaveBalance.user_id, LeaveBalance.year, LeaveBalance.leave_code, unique=True)


class LineMessageLog(Base):
    """LINE 訊息與解析結果（除錯用）"""
    __tablename__ = "calendar_line_logs"

    id           = Column(Integer, primary_key=True, index=True)
    user_id      = Column(Integer, ForeignKey("users.id"), nullable=True, index=True)
    line_user_id = Column(String(100), nullable=True)
    raw_text     = Column(Text, nullable=False)
    intent       = Column(String(20), nullable=True)
    parsed_json  = Column(Text, nullable=True)
    result_event_id = Column(Integer, ForeignKey("calendar_events.id"), nullable=True)
    error_msg    = Column(Text, nullable=True)
    created_at   = Column(DateTime, default=datetime.utcnow, nullable=False)
