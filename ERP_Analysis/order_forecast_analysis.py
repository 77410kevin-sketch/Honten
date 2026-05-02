"""
訂單預計明細 — 年度分析圖表
目標：客戶 × 月份 × 區塊 三維分析

使用前請先跑 step0_explore_columns.py 確認欄位名稱，
然後調整下方 CONFIG 區的欄位對應。
"""
import os
import pyodbc
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from dotenv import load_dotenv

load_dotenv()

# ─────────────────────────────────────────────
# CONFIG：根據 step0 印出的欄位名稱填這裡
# ─────────────────────────────────────────────
YEAR = 2026                          # 要分析的年度

VIEW_NAME    = "v_ht_customer_order_lines"  # 訂單明細 View

COL_DATE     = "order_date"          # 訂單日期欄位
COL_CUSTOMER = "customer_name"       # 客戶名稱欄位
COL_AMOUNT   = "total_amount"        # 金額欄位（若是數量改成 quantity）
COL_BLOCK    = "area_code"           # 區塊/業務區域欄位（跑完 step0 再填正確名稱）

TOP_N_CUSTOMERS = 10                 # 只顯示前 N 大客戶
# ─────────────────────────────────────────────

conn_str = (
    f"DRIVER={{{os.getenv('ERP_ODBC_DRIVER', 'ODBC Driver 17 for SQL Server')}}};"
    f"SERVER={os.getenv('ERP_SERVER')},{os.getenv('ERP_PORT', '1433')};"
    f"DATABASE={os.getenv('ERP_DATABASE')};"
    f"UID={os.getenv('ERP_USER')};"
    f"PWD={os.getenv('ERP_PASSWORD')};"
    "TrustServerCertificate=yes;"
    "timeout=30;"
)


def run_query(sql: str, params: tuple = ()) -> pd.DataFrame:
    conn = pyodbc.connect(conn_str)
    try:
        df = pd.read_sql(sql, conn, params=params)
        for col in df.select_dtypes(include='object').columns:
            df[col] = df[col].str.rstrip()
        return df
    finally:
        conn.close()


def fetch_data() -> pd.DataFrame:
    sql = f"""
    SELECT
        {COL_DATE}   AS order_date,
        {COL_CUSTOMER} AS customer,
        {COL_AMOUNT}   AS amount,
        {COL_BLOCK}    AS block
    FROM {VIEW_NAME}
    WHERE YEAR({COL_DATE}) = ?
      AND {COL_AMOUNT} IS NOT NULL
      AND {COL_AMOUNT} > 0
    """
    print(f"正在撈 {YEAR} 年資料…")
    df = run_query(sql, params=(YEAR,))
    df['order_date'] = pd.to_datetime(df['order_date'])
    df['month'] = df['order_date'].dt.month
    df['month_label'] = df['order_date'].dt.strftime('%Y-%m')
    print(f"✅ 撈到 {len(df):,} 筆")
    return df


def get_top_customers(df: pd.DataFrame) -> list:
    top = (df.groupby('customer')['amount']
             .sum()
             .nlargest(TOP_N_CUSTOMERS)
             .index.tolist())
    return top


def build_dashboard(df: pd.DataFrame):
    top_customers = get_top_customers(df)
    df_top = df[df['customer'].isin(top_customers)].copy()

    # ── 月彙總 ──
    monthly = (df.groupby('month_label')['amount']
                 .sum()
                 .reset_index()
                 .sort_values('month_label'))

    # ── 客戶彙總 ──
    customer_total = (df.groupby('customer')['amount']
                        .sum()
                        .reset_index()
                        .sort_values('amount', ascending=False)
                        .head(TOP_N_CUSTOMERS))

    # ── 區塊彙總 ──
    block_total = (df.groupby('block')['amount']
                     .sum()
                     .reset_index()
                     .sort_values('amount', ascending=False))

    # ── 客戶 × 月份（Heat Map）──
    pivot = (df_top.groupby(['customer', 'month_label'])['amount']
                   .sum()
                   .reset_index()
                   .pivot(index='customer', columns='month_label', values='amount')
                   .fillna(0))
    pivot = pivot.loc[top_customers]   # 保持排序

    # ── 區塊 × 月份趨勢 ──
    block_monthly = (df.groupby(['block', 'month_label'])['amount']
                       .sum()
                       .reset_index()
                       .sort_values('month_label'))

    # ═══════════════ 建圖 ═══════════════
    fig = make_subplots(
        rows=3, cols=2,
        subplot_titles=(
            f"{YEAR} 年月別訂單金額趨勢",
            f"前 {TOP_N_CUSTOMERS} 大客戶訂單總額",
            "區塊訂單占比",
            f"區塊 × 月份趨勢",
            f"前 {TOP_N_CUSTOMERS} 大客戶 × 月份 Heatmap",
            "",
        ),
        specs=[
            [{"type": "xy"},    {"type": "xy"}],
            [{"type": "domain"}, {"type": "xy"}],
            [{"colspan": 2, "type": "xy"}, None],
        ],
        row_heights=[0.28, 0.28, 0.44],
        vertical_spacing=0.10,
        horizontal_spacing=0.08,
    )

    # 1. 月別趨勢（折線）
    fig.add_trace(
        go.Scatter(
            x=monthly['month_label'], y=monthly['amount'],
            mode='lines+markers', name='月別金額',
            line=dict(color='#2b6cb0', width=2),
            marker=dict(size=7),
        ),
        row=1, col=1
    )

    # 2. 前 N 大客戶（橫條）
    fig.add_trace(
        go.Bar(
            x=customer_total['amount'],
            y=customer_total['customer'],
            orientation='h', name='客戶',
            marker_color='#2f855a',
            text=customer_total['amount'].apply(lambda x: f"{x:,.0f}"),
            textposition='outside',
        ),
        row=1, col=2
    )

    # 3. 區塊占比（圓餅）
    fig.add_trace(
        go.Pie(
            labels=block_total['block'],
            values=block_total['amount'],
            name='區塊',
            hole=0.35,
            textinfo='label+percent',
        ),
        row=2, col=1
    )

    # 4. 區塊 × 月份（折線群組）
    for blk in block_monthly['block'].unique():
        bdf = block_monthly[block_monthly['block'] == blk]
        fig.add_trace(
            go.Scatter(
                x=bdf['month_label'], y=bdf['amount'],
                mode='lines+markers', name=f"區塊:{blk}",
                line=dict(width=2),
            ),
            row=2, col=2
        )

    # 5. Heatmap（客戶 × 月份）
    fig.add_trace(
        go.Heatmap(
            z=pivot.values,
            x=pivot.columns.tolist(),
            y=pivot.index.tolist(),
            colorscale='Blues',
            text=[[f"{v:,.0f}" for v in row] for row in pivot.values],
            texttemplate="%{text}",
            textfont=dict(size=10),
            showscale=True,
            name='',
        ),
        row=3, col=1
    )

    fig.update_layout(
        title_text=f"訂單預計明細 — {YEAR} 年度分析",
        title_font_size=20,
        height=1400,
        showlegend=True,
        legend=dict(orientation='v', x=1.02, y=0.5),
        font=dict(family="Microsoft JhengHei, Arial", size=12),
        plot_bgcolor='#f8f9fa',
        paper_bgcolor='#ffffff',
    )

    return fig


def export_summary(df: pd.DataFrame):
    os.makedirs("output", exist_ok=True)
    path = f"output/訂單預計明細_{YEAR}年分析.xlsx"
    with pd.ExcelWriter(path, engine='openpyxl') as writer:
        # 原始資料
        df.to_excel(writer, sheet_name='原始資料', index=False)
        # 月別
        monthly = (df.groupby('month_label')['amount']
                     .sum().reset_index().sort_values('month_label'))
        monthly.to_excel(writer, sheet_name='月別彙總', index=False)
        # 客戶
        cust = (df.groupby('customer')['amount']
                  .sum().reset_index().sort_values('amount', ascending=False))
        cust.to_excel(writer, sheet_name='客戶彙總', index=False)
        # 區塊
        block = (df.groupby('block')['amount']
                   .sum().reset_index().sort_values('amount', ascending=False))
        block.to_excel(writer, sheet_name='區塊彙總', index=False)
        # 客戶×月份
        pivot_full = (df.groupby(['customer', 'month_label'])['amount']
                        .sum().reset_index()
                        .pivot(index='customer', columns='month_label', values='amount')
                        .fillna(0))
        pivot_full.to_excel(writer, sheet_name='客戶×月份')
    print(f"✅ Excel 已匯出：{path}")


if __name__ == "__main__":
    df = fetch_data()
    if df.empty:
        print("⚠️  查無資料，請確認 YEAR 設定與欄位名稱是否正確")
    else:
        export_summary(df)
        fig = build_dashboard(df)
        html_path = f"output/訂單預計明細_{YEAR}年圖表.html"
        fig.write_html(html_path)
        print(f"✅ 互動式圖表已匯出：{html_path}")
        fig.show()   # 直接在瀏覽器開啟
