"""ERP 連接口（抽象介面 + Stub 實作）

此模組定義 ERP 資料來源的統一介面，目前提供 **Stub 版本**（讀本機假資料），
未來 ERP 可連線時，只要把 `_BACKEND` 改成真實 SQL Server 實作，或在
`.env` 設定 `ERP_BACKEND=sqlserver`，其他程式完全不用改。

連線資訊（.env）：
    ERP_BACKEND=stub          # 目前用；未來改 sqlserver
    ERP_HOST=192.168.0.201
    ERP_PORT=1433
    ERP_DATABASE=HTE2026
    ERP_USER=ht_sys
    ERP_PASSWORD=***
    ERP_CUSTOMER_VIEW=v_customer_master   # 尚未指定 → 填入後 stub 會自動停用
    ERP_SUPPLIER_VIEW=v_supplier_master
    ERP_PROCESS_VIEW=v_process_code

實作切換：
    - Stub：呼叫 _stub_* 函式，回本機假資料（開發/Demo 用）
    - SQLServer：呼叫 _sqlserver_* 函式（未實作，pyodbc pending）
"""
from __future__ import annotations
import os
import logging
from dataclasses import dataclass, asdict
from typing import Protocol

logger = logging.getLogger(__name__)


# ── 資料傳輸物件（DTO）──────────────────────
@dataclass
class ERPCustomer:
    erp_code:   str              # ERP 客戶代碼（主鍵）
    name:       str              # 客戶名稱
    contact:    str | None = None
    email:      str | None = None
    phone:      str | None = None
    address:    str | None = None
    bu:         str | None = None   # 儲能事業部 / 消費性事業部
    tax_id:     str | None = None
    is_active:  bool = True

    def as_dict(self) -> dict:
        return asdict(self)


@dataclass
class ERPSupplier:
    erp_code:   str
    name:       str
    type:       str = "外部"        # 廠內 / 外部
    contact:    str | None = None
    email:      str | None = None
    phone:      str | None = None
    tax_id:     str | None = None
    is_active:  bool = True

    def as_dict(self) -> dict:
        return asdict(self)


@dataclass
class ERPProcess:
    code:        str              # 製程代碼，例 CNC001
    name:        str              # 製程名稱
    category:    str | None = None # 例：機加工 / 表面處理 / 熱處理
    default_supplier_ids: list[str] | None = None  # 建議供應商（ERP 代碼）

    def as_dict(self) -> dict:
        return asdict(self)


# ── 後端介面（Protocol）──────────────────────
class ERPBackend(Protocol):
    """任何 ERP backend 必須實作以下三個方法。"""
    def fetch_customers(self) -> list[ERPCustomer]: ...
    def fetch_suppliers(self) -> list[ERPSupplier]: ...
    def fetch_processes(self) -> list[ERPProcess]: ...
    def is_connected(self) -> bool: ...


# ── Stub 實作（本機假資料）──────────────────────
class _StubBackend:
    """開發/Demo 用：固定回傳示範資料，不呼叫任何外部服務。"""
    _CUSTOMERS = [
        ERPCustomer("C001", "景利（Jingli）",      "林副總", "lin@jingli.com.tw",   "02-1111-2222", "新北市", "儲能事業部", "12345678"),
        ERPCustomer("C002", "愛爾蘭金士頓",        "Kevin",  "kevin@airkingston.ie","+353-1-222-3333", "Dublin", "消費性事業部", "IE1234567"),
        ERPCustomer("C003", "佐茂（Zuomao）",      "王廠長", "wang@zuomao.cn",      "+86-755-8888-9999", "深圳", "消費性事業部", None),
        ERPCustomer("C004", "嘉彰（Jiazhang）",    "陳經理", "chen@jiazhang.com",   "03-5678-9012", "新竹", "儲能事業部", None),
        ERPCustomer("C005", "艾思科（Ai-Si-Ke）",  "Robin",  "robin@ai-si-ke.com",  "02-3456-7890", "台北", "消費性事業部", None),
    ]
    _SUPPLIERS = [
        ERPSupplier("S001", "新北方模具",  "外部", "王經理", "w@sd.com.tw",      "02-1234-5678"),
        ERPSupplier("S002", "豐隆精密",    "外部", "李廠長", "li@fx.com.tw",     "03-2345-6789"),
        ERPSupplier("S003", "昌泰五金",    "外部", "陳業務", "chen@ct.com.tw",   "04-3456-7890"),
        ERPSupplier("S004", "久盛塑膠",    "外部", "張經理", "js@jiusheng.com.tw","02-4567-8901"),
        ERPSupplier("S005", "台達電（Delta）","外部","張副理", "zhang@delta.com.tw","02-8797-2088"),
        ERPSupplier("S006", "凌陽（Sunplus）","外部","陳協理", "chen@sunplus.com", "03-578-6005"),
        ERPSupplier("I001", "機加工課",    "廠內", "陳課長", "mach@honten.local","分機 2301"),
        ERPSupplier("I002", "模具課",      "廠內", "林課長", "mold@honten.local","分機 2401"),
    ]
    _PROCESSES = [
        ERPProcess("P001", "CNC 加工",          "機加工",   ["S001", "S002"]),
        ERPProcess("P002", "沖壓",              "機加工",   ["S003"]),
        ERPProcess("P003", "銑床",              "機加工",   ["I001"]),
        ERPProcess("P004", "車床",              "機加工",   ["I001", "S002"]),
        ERPProcess("P005", "線割",              "機加工",   ["S001"]),
        ERPProcess("P006", "熱處理",            "熱處理",   ["S002"]),
        ERPProcess("P007", "表面處理（陽極）",  "表面處理", ["S003"]),
        ERPProcess("P008", "表面處理（電鍍）",  "表面處理", ["S003"]),
        ERPProcess("P009", "噴砂",              "表面處理", ["S003"]),
        ERPProcess("P010", "烤漆",              "表面處理", ["S003"]),
        ERPProcess("P011", "射出成型",          "塑膠成型", ["S004"]),
        ERPProcess("P012", "壓鑄",              "鑄造",     ["S002"]),
        ERPProcess("P013", "模具製作",          "模具",     ["S001", "I002"]),
        ERPProcess("P014", "雷射切割",          "機加工",   ["S001"]),
        ERPProcess("P015", "組裝",              "組裝",     ["I001"]),
        ERPProcess("P016", "包裝",              "組裝",     ["I001"]),
    ]

    def fetch_customers(self) -> list[ERPCustomer]:
        return list(self._CUSTOMERS)

    def fetch_suppliers(self) -> list[ERPSupplier]:
        return list(self._SUPPLIERS)

    def fetch_processes(self) -> list[ERPProcess]:
        return list(self._PROCESSES)

    def is_connected(self) -> bool:
        return True   # Stub 永遠「連線」


# ── SQL Server 實作（預留，尚未實作）────────────
class _SQLServerBackend:
    """透過 pyodbc 連 ERP SQL Server 的實作（待 ERP 開通後填入）。

    預期實作：
        conn = pyodbc.connect(
            f"DRIVER={{ODBC Driver 17 for SQL Server}};"
            f"SERVER={os.getenv('ERP_HOST')},{os.getenv('ERP_PORT')};"
            f"DATABASE={os.getenv('ERP_DATABASE')};"
            f"UID={os.getenv('ERP_USER')};PWD={os.getenv('ERP_PASSWORD')}"
        )
        cur = conn.cursor()
        cur.execute(f"SELECT ... FROM {os.getenv('ERP_CUSTOMER_VIEW')}")
        return [ERPCustomer(**row) for row in cur.fetchall()]
    """
    def fetch_customers(self) -> list[ERPCustomer]:
        raise NotImplementedError("SQL Server ERP 實作未就緒，請於 erp_client._SQLServerBackend 完成")

    def fetch_suppliers(self) -> list[ERPSupplier]:
        raise NotImplementedError()

    def fetch_processes(self) -> list[ERPProcess]:
        raise NotImplementedError()

    def is_connected(self) -> bool:
        for key in ("ERP_HOST", "ERP_DATABASE", "ERP_USER", "ERP_PASSWORD"):
            if not os.getenv(key):
                return False
        return True


# ── 單例切換 ──────────────────────────────────────
def _select_backend() -> ERPBackend:
    backend = os.getenv("ERP_BACKEND", "stub").lower()
    if backend == "sqlserver":
        return _SQLServerBackend()
    return _StubBackend()


_BACKEND: ERPBackend = _select_backend()


# ── 對外 API ─────────────────────────────────────
def fetch_customers_from_erp() -> list[ERPCustomer]:
    """取得 ERP 客戶清單。Stub 模式回示範資料。"""
    try:
        return _BACKEND.fetch_customers()
    except NotImplementedError:
        logger.warning("ERP backend 尚未實作 fetch_customers，回空陣列")
        return []


def fetch_suppliers_from_erp() -> list[ERPSupplier]:
    try:
        return _BACKEND.fetch_suppliers()
    except NotImplementedError:
        logger.warning("ERP backend 尚未實作 fetch_suppliers，回空陣列")
        return []


def fetch_processes_from_erp() -> list[ERPProcess]:
    try:
        return _BACKEND.fetch_processes()
    except NotImplementedError:
        logger.warning("ERP backend 尚未實作 fetch_processes，回空陣列")
        return []


def erp_status() -> dict:
    """回傳目前 ERP 連接口狀態（UI 頁首用）"""
    backend_name = os.getenv("ERP_BACKEND", "stub").lower()
    return {
        "backend": backend_name,
        "is_stub": backend_name == "stub",
        "connected": _BACKEND.is_connected(),
    }
