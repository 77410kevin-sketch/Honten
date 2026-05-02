"""
Step 0 — 先跑這支，確認 View 欄位名稱（尤其是「區塊」叫什麼）
結果會印在 console，不匯出。
"""
import os
import pyodbc
import pandas as pd
from dotenv import load_dotenv

load_dotenv()

conn_str = (
    f"DRIVER={{{os.getenv('ERP_ODBC_DRIVER', 'ODBC Driver 17 for SQL Server')}}};"
    f"SERVER={os.getenv('ERP_SERVER')},{os.getenv('ERP_PORT', '1433')};"
    f"DATABASE={os.getenv('ERP_DATABASE')};"
    f"UID={os.getenv('ERP_USER')};"
    f"PWD={os.getenv('ERP_PASSWORD')};"
    "TrustServerCertificate=yes;"
    "timeout=10;"
)


def get_columns(view_name: str) -> pd.DataFrame:
    conn = pyodbc.connect(conn_str)
    try:
        sql = """
        SELECT COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = ?
        ORDER BY ORDINAL_POSITION
        """
        return pd.read_sql(sql, conn, params=(view_name,))
    finally:
        conn.close()


def preview(view_name: str, n: int = 3) -> pd.DataFrame:
    conn = pyodbc.connect(conn_str)
    try:
        df = pd.read_sql(f"SELECT TOP {n} * FROM {view_name}", conn)
        for col in df.select_dtypes(include='object').columns:
            df[col] = df[col].str.rstrip()
        return df
    finally:
        conn.close()


if __name__ == "__main__":
    VIEWS_TO_CHECK = [
        "v_ht_customer_order_lines",
        "v_ht_customer_order",
        # 若有預計明細的專屬 View 再加進來，例如：
        # "v_ht_order_forecast",
    ]

    for v in VIEWS_TO_CHECK:
        print(f"\n{'='*60}")
        print(f"  View: {v}")
        print(f"{'='*60}")
        try:
            cols = get_columns(v)
            print(cols.to_string(index=False))
            print(f"\n-- 前 3 筆預覽 --")
            df = preview(v)
            pd.set_option('display.max_columns', None)
            pd.set_option('display.width', 200)
            print(df.to_string(index=False))
        except Exception as e:
            print(f"  ⚠️  讀取失敗: {e}")

    print("\n\n完成！把上面的欄位名稱告訴 Claude，他會幫你填進分析腳本。")
