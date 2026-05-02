"""
成本分析：PURR16（採購成本）+ MOCR34（委外製程成本）
輸出：Excel + HTML 儀表板

PURR16 → v_ht_purchase_order    （廠商 × 料號 × 金額）
MOCR34 → v_ht_manufacturing_order（委外淨金額）
         v_ht_mo_routing          （委外廠商 × 製程）
"""
import os, sys, warnings, argparse, json
from datetime import date, timedelta
import pyodbc, pandas as pd
from dotenv import load_dotenv

warnings.filterwarnings('ignore')
load_dotenv()

# ── CONFIG ────────────────────────────────────────────────────
DATE_FROM = os.getenv("COST_DATE_FROM", "2026-01-01")
DATE_TO   = os.getenv("COST_DATE_TO",   "2026-12-31")

# doc_type 篩選（採購）
PO_TYPES  = os.getenv("PO_ORDER_TYPES", "3302").split(",")   # 3302=一般採購單

CHART_COLORS = [
    "#1b5e20","#2e7d32","#388e3c","#43a047","#66bb6a",
    "#0d47a1","#1565c0","#1976d2","#2196f3","#64b5f6",
    "#f57f17","#f9a825","#fbc02d","#fdd835","#fff176",
    "#b71c1c","#c62828","#d32f2f","#e53935","#ef9a9a",
    "#4a148c","#6a1b9a","#7b1fa2","#ab47bc","#ce93d8",
    "#e65100","#ef6c00","#f57c00","#fb8c00","#ffa726",
    "#006064","#00838f","#0097a7","#26c6da","#80deea",
]


# ── 連線 ──────────────────────────────────────────────────────
def _conn_str():
    d  = os.getenv("ERP_ODBC_DRIVER","SQL Server")
    s  = os.getenv("ERP_SERVER","");  p = os.getenv("ERP_PORT","1433")
    db = os.getenv("ERP_DATABASE",""); u=os.getenv("ERP_USER",""); pw=os.getenv("ERP_PASSWORD","")
    miss=[k for k,v in{"ERP_SERVER":s,"ERP_DATABASE":db,"ERP_USER":u,"ERP_PASSWORD":pw}.items() if not v]
    if miss: print(f"[ERROR] .env 缺少：{miss}"); sys.exit(1)
    return (f"DRIVER={{{d}}};SERVER={s},{p};DATABASE={db};"
            f"UID={u};PWD={pw};TrustServerCertificate=yes;timeout=30;")

def get_conn(): return pyodbc.connect(_conn_str())


# ── PURR16：撈採購成本 ────────────────────────────────────────
def fetch_purr16() -> pd.DataFrame:
    ph  = ",".join("?"*len(PO_TYPES))
    sql = f"""
    SELECT order_no, doc_type, order_date,
           vendor_code, vendor_name,
           product_code, product_name,
           qty, unit_price, amount,
           delivery_date, delivered_qty
    FROM   v_ht_purchase_order
    WHERE  order_date >= ? AND order_date <= ?
      AND  doc_type IN ({ph})
    ORDER BY order_date, vendor_code, product_code
    """
    params = [DATE_FROM, DATE_TO] + PO_TYPES
    print(f"[PURR16] 採購成本 {DATE_FROM}~{DATE_TO}，doc_type={PO_TYPES}")
    conn = get_conn()
    try:    df = pd.read_sql(sql, conn, params=params)
    finally: conn.close()
    for c in df.select_dtypes(include='str').columns:
        df[c] = df[c].str.rstrip()
    df['order_date']    = pd.to_datetime(df['order_date'],    errors='coerce')
    df['delivery_date'] = pd.to_datetime(df['delivery_date'], errors='coerce')
    df['月份'] = df['order_date'].dt.strftime('%Y-%m')
    print(f"[PURR16] 撈到 {len(df):,} 筆")
    return df


# ── MOCR34：撈委外製程成本 ───────────────────────────────────
def fetch_mocr34() -> tuple[pd.DataFrame, pd.DataFrame]:
    """回傳 (df_mo, df_routing)"""
    # 製令主檔（含委外淨金額）
    sql_mo = """
    SELECT doc_type, order_no, order_date,
           product_code, product_name, product_spec,
           plan_qty, done_qty,
           outsource_in_amount, outsource_return_amount, outsource_net_amount
    FROM   v_ht_manufacturing_order
    WHERE  order_date >= ? AND order_date <= ?
      AND  outsource_net_amount > 0
    ORDER BY order_date, product_code
    """
    # 製令途程（廠商資訊）
    sql_rt = """
    SELECT r.doc_type, r.order_no, r.process_seq,
           r.process_code, r.process_name,
           r.line_vendor_code, r.line_vendor_name,
           r.plan_end_date, r.input_qty, r.done_qty
    FROM   v_ht_mo_routing r
    INNER JOIN v_ht_manufacturing_order m
           ON r.order_no = m.order_no
    WHERE  m.order_date >= ? AND m.order_date <= ?
      AND  m.outsource_net_amount > 0
      AND  r.line_vendor_code NOT IN ('001','002','003','004','005','006','007','008','009')
    ORDER BY r.order_no, r.process_seq
    """
    print(f"[MOCR34] 委外製程成本 {DATE_FROM}~{DATE_TO}")
    conn = get_conn()
    try:
        df_mo = pd.read_sql(sql_mo, conn, params=[DATE_FROM, DATE_TO])
        df_rt = pd.read_sql(sql_rt, conn, params=[DATE_FROM, DATE_TO])
    finally:
        conn.close()
    for df in [df_mo, df_rt]:
        for c in df.select_dtypes(include='str').columns:
            df[c] = df[c].str.rstrip()
    df_mo['order_date'] = pd.to_datetime(df_mo['order_date'], errors='coerce')
    df_mo['月份'] = df_mo['order_date'].dt.strftime('%Y-%m')
    print(f"[MOCR34] 製令 {len(df_mo):,} 筆，途程 {len(df_rt):,} 筆")
    return df_mo, df_rt


# ── 廠商×月份 pivot helper ───────────────────────────────────
def make_vendor_pivot(df_src: pd.DataFrame, vendor_col: str, amount_col: str, month_col: str):
    months   = sorted(df_src[month_col].dropna().unique().tolist())
    vendors  = (df_src.groupby(vendor_col)[amount_col].sum()
                      .sort_values(ascending=False).index.tolist())
    pivot = (df_src.groupby([month_col, vendor_col])[amount_col].sum()
                   .unstack(fill_value=0)
                   .reindex(index=months, columns=vendors, fill_value=0))
    return months, vendors, pivot


# ── HTML 生成 ─────────────────────────────────────────────────
def fmt_amt(v):
    if v >= 1_000_000: return f"{v/1_000_000:.1f}M"
    if v >= 10_000:    return f"{v/10_000:.1f}萬"
    return f"{v:,.0f}"

def pivot_table_html(months, vendors, pivot, amount_label="金額"):
    th  = 'style="background:#1b5e20;color:#fff;padding:6px 10px;white-space:nowrap;text-align:right;border:1px solid #388e3c;"'
    th0 = 'style="background:#1b5e20;color:#fff;padding:6px 10px;white-space:nowrap;text-align:left;border:1px solid #388e3c;"'
    td  = 'style="padding:5px 10px;text-align:right;border:1px solid #e0e0e0;white-space:nowrap;"'
    tdm = 'style="padding:5px 10px;font-weight:600;border:1px solid #e0e0e0;background:#f5f5f5;"'
    tdt = 'style="padding:5px 10px;text-align:right;border:1px solid #388e3c;background:#c8e6c9;font-weight:700;"'
    tdg = 'style="padding:5px 10px;text-align:right;border:1px solid #e0e0e0;background:#e8f5e9;font-weight:700;"'

    cols_html = "".join(f'<th {th}>{v}</th>' for v in vendors)
    header = f'<thead><tr><th {th0}>月份</th>{cols_html}<th {th}>月合計</th></tr></thead>'

    rows_html = ""; grand = {v: 0 for v in vendors}; grand_total = 0
    for m in months:
        row_total = 0; cells = ""
        for v in vendors:
            val = float(pivot.loc[m, v]) if m in pivot.index else 0
            grand[v] += val; row_total += val
            cells += f'<td {td}>{"−" if val==0 else f"{val:,.0f}"}</td>'
        grand_total += row_total
        rows_html += f'<tr><td {tdm}>{m}</td>{cells}<td {tdt}>{row_total:,.0f}</td></tr>'

    grand_cells = "".join(f'<td {tdg}>{grand[v]:,.0f}</td>' for v in vendors)
    rows_html += f'<tr><td {tdg}>總計</td>{grand_cells}<td {tdg}>{grand_total:,.0f}</td></tr>'

    return (f'<div style="overflow-x:auto;">'
            f'<table class="table table-sm table-bordered" style="font-size:.80rem;">'
            f'{header}<tbody>{rows_html}</tbody></table></div>')


def export_html(df_po: pd.DataFrame, df_mo: pd.DataFrame, df_rt: pd.DataFrame) -> str:
    os.makedirs("output", exist_ok=True)
    tag  = f"{DATE_FROM.replace('-','')}_{DATE_TO.replace('-','')}"
    path = f"output/成本分析儀表板_{tag}.html"
    today_str = date.today().strftime("%Y/%m/%d")

    # ── PURR16 資料 ──
    po_months, po_vendors, po_pivot = make_vendor_pivot(df_po, 'vendor_name', 'amount', '月份')
    po_total     = df_po['amount'].sum()
    po_n_vendors = df_po['vendor_name'].nunique()
    po_n_parts   = df_po['product_code'].nunique()

    # 廠商彙總
    po_vendor_sum = (df_po.groupby(['vendor_code','vendor_name'])
                          .agg(採購金額=('amount','sum'), 筆數=('order_no','count'))
                          .reset_index().sort_values('採購金額', ascending=False))
    # 料號彙總
    po_part_sum   = (df_po.groupby(['product_code','product_name'])
                          .agg(採購金額=('amount','sum'), 數量=('qty','sum'))
                          .reset_index().sort_values('採購金額', ascending=False))

    # Chart.js datasets（PURR16 堆疊長條）
    po_datasets = []
    for i, vend in enumerate(po_vendors[:30]):   # 最多顯示 30 個廠商
        color = CHART_COLORS[i % len(CHART_COLORS)]
        po_datasets.append({
            "label": vend,
            "data":  [round(float(po_pivot.loc[m, vend]), 0) if m in po_pivot.index else 0
                      for m in po_months],
            "backgroundColor": color, "stack": "s",
        })

    # ── MOCR34 資料 ──
    mo_total       = df_mo['outsource_net_amount'].sum()
    mo_n_products  = df_mo['product_code'].nunique()
    mo_n_vendors   = df_rt['line_vendor_name'].nunique() if not df_rt.empty else 0

    # 製令×月份 產品金額
    mo_product_sum = (df_mo.groupby(['product_code','product_name'])
                           .agg(委外金額=('outsource_net_amount','sum'),
                                製令數=('order_no','count'))
                           .reset_index().sort_values('委外金額', ascending=False))

    mo_monthly = (df_mo.groupby('月份')['outsource_net_amount']
                       .sum().reset_index().sort_values('月份')
                       .rename(columns={'outsource_net_amount':'委外金額'}))

    # 途程廠商彙總（由製令主檔關聯）
    if not df_rt.empty:
        df_rt_mo = df_rt.merge(
            df_mo[['order_no','outsource_net_amount','月份']],
            on='order_no', how='left')
        rt_vendor_sum = (df_rt_mo.groupby(['line_vendor_code','line_vendor_name'])
                                  .agg(關聯製令數=('order_no','nunique'))
                                  .reset_index().sort_values('關聯製令數', ascending=False))
    else:
        rt_vendor_sum = pd.DataFrame(columns=['line_vendor_code','line_vendor_name','關聯製令數'])

    # ── Pivot HTML ──
    po_pivot_html = pivot_table_html(po_months, po_vendors[:20], po_pivot, "採購金額")

    # ── 廠商排行 HTML（PURR16）──
    po_vendor_rows = "".join(
        f'<tr><td>{r.vendor_code}</td><td>{r.vendor_name}</td>'
        f'<td class="text-end fw-bold">{r.採購金額:,.0f}</td>'
        f'<td class="text-end">{r.筆數}</td></tr>'
        for _, r in po_vendor_sum.head(30).iterrows()
    )
    # ── 料號排行 HTML ──
    po_part_rows = "".join(
        f'<tr><td>{r.product_code}</td><td>{r.product_name}</td>'
        f'<td class="text-end">{r.數量:,.0f}</td>'
        f'<td class="text-end fw-bold">{r.採購金額:,.0f}</td></tr>'
        for _, r in po_part_sum.head(50).iterrows()
    )
    # ── 製令產品排行 HTML ──
    mo_product_rows = "".join(
        f'<tr><td>{r.product_code}</td><td>{r.product_name}</td>'
        f'<td class="text-end">{r.製令數}</td>'
        f'<td class="text-end fw-bold">{r.委外金額:,.0f}</td></tr>'
        for _, r in mo_product_sum.head(50).iterrows()
    )
    # ── 途程廠商 HTML ──
    rt_vendor_rows = "".join(
        f'<tr><td>{r.line_vendor_code}</td><td>{r.line_vendor_name}</td>'
        f'<td class="text-end">{r.關聯製令數}</td></tr>'
        for _, r in rt_vendor_sum.head(30).iterrows()
    )
    # ── 月別委外金額 HTML ──
    mo_monthly_rows = "".join(
        f'<tr><td>{r.月份}</td><td class="text-end fw-bold">{r.委外金額:,.0f}</td></tr>'
        for _, r in mo_monthly.iterrows()
    )

    # JSON for Charts
    js_po_months   = json.dumps(po_months, ensure_ascii=False)
    js_po_datasets = json.dumps(po_datasets, ensure_ascii=False)
    js_mo_months   = json.dumps(mo_monthly['月份'].tolist(), ensure_ascii=False)
    js_mo_data     = json.dumps([round(float(v),0) for v in mo_monthly['委外金額'].tolist()])
    # 採購廠商圓餅
    top_po_vend    = po_vendor_sum.head(15)
    js_pie_labels  = json.dumps(top_po_vend['vendor_name'].tolist(), ensure_ascii=False)
    js_pie_data    = json.dumps([round(float(v),0) for v in top_po_vend['採購金額'].tolist()])
    js_pie_colors  = json.dumps(CHART_COLORS[:len(top_po_vend)])

    html = f"""<!DOCTYPE html>
<html lang="zh-Hant">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>成本分析儀表板 — PURR16 + MOCR34</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.3/dist/chart.umd.min.js"></script>
<style>
  body{{font-family:"Microsoft JhengHei",Arial,sans-serif;background:#f4f6f9;}}
  .header{{background:linear-gradient(135deg,#1a237e 0%,#283593 60%,#1565c0 100%);
           color:#fff;padding:24px 32px;border-radius:0 0 16px 16px;margin-bottom:24px;}}
  .header h1{{font-size:1.5rem;font-weight:700;margin:0;}}
  .header .sub{{font-size:.85rem;opacity:.8;margin-top:4px;}}
  .kpi-card{{background:#fff;border-radius:12px;padding:20px 24px;
             box-shadow:0 2px 8px rgba(0,0,0,.07);border-left:5px solid;}}
  .kpi-card .val{{font-size:1.7rem;font-weight:700;}}
  .kpi-card .lbl{{font-size:.8rem;color:#666;margin-top:2px;}}
  .section-title{{font-size:1rem;font-weight:700;color:#1565c0;margin-bottom:12px;
                  border-left:4px solid #1565c0;padding-left:10px;}}
  .card-box{{background:#fff;border-radius:12px;padding:20px;
             box-shadow:0 2px 8px rgba(0,0,0,.07);margin-bottom:16px;}}
  .nav-tabs .nav-link.active{{color:#1565c0;border-bottom:3px solid #1565c0;font-weight:700;}}
  .nav-tabs .nav-link{{color:#555;font-weight:600;}}
  .badge-tag{{background:#e3f2fd;color:#1565c0;border-radius:6px;padding:3px 8px;font-size:.78rem;}}
</style>
</head>
<body>

<div class="header">
  <h1>成本分析儀表板 — PURR16 採購成本 ＋ MOCR34 委外製程成本</h1>
  <div class="sub">期間：{DATE_FROM} ～ {DATE_TO} ｜ 更新：{today_str}</div>
</div>

<div class="container-fluid px-4">

  <!-- KPI -->
  <div class="row g-3 mb-4">
    <div class="col-6 col-md-3">
      <div class="kpi-card" style="border-color:#1565c0;">
        <div class="val" style="color:#1565c0;">{fmt_amt(po_total)}</div>
        <div class="lbl">採購總金額（NT$）<span class="badge-tag">PURR16</span></div>
      </div>
    </div>
    <div class="col-6 col-md-3">
      <div class="kpi-card" style="border-color:#0097a7;">
        <div class="val" style="color:#0097a7;">{po_n_vendors}</div>
        <div class="lbl">採購廠商數 <span class="badge-tag">PURR16</span></div>
      </div>
    </div>
    <div class="col-6 col-md-3">
      <div class="kpi-card" style="border-color:#e65100;">
        <div class="val" style="color:#e65100;">{fmt_amt(mo_total)}</div>
        <div class="lbl">委外製程總金額（NT$）<span class="badge-tag">MOCR34</span></div>
      </div>
    </div>
    <div class="col-6 col-md-3">
      <div class="kpi-card" style="border-color:#6a1b9a;">
        <div class="val" style="color:#6a1b9a;">{mo_n_products}</div>
        <div class="lbl">委外產品種類 <span class="badge-tag">MOCR34</span></div>
      </div>
    </div>
  </div>

  <!-- Tabs -->
  <ul class="nav nav-tabs mb-3" id="mainTab">
    <li class="nav-item"><a class="nav-link active" data-bs-toggle="tab" href="#tab1">
      採購成本（PURR16）</a></li>
    <li class="nav-item"><a class="nav-link" data-bs-toggle="tab" href="#tab2">
      委外製程成本（MOCR34）</a></li>
  </ul>

  <div class="tab-content">

    <!-- ════ Tab 1：PURR16 採購成本 ════ -->
    <div class="tab-pane fade show active" id="tab1">

      <div class="row g-3 mb-3">
        <!-- 月別堆疊長條 -->
        <div class="col-12 col-xl-8">
          <div class="card-box">
            <div class="section-title">月別採購金額（依廠商堆疊）</div>
            <canvas id="poBarChart" style="max-height:380px;"></canvas>
          </div>
        </div>
        <!-- 廠商圓餅 -->
        <div class="col-12 col-xl-4">
          <div class="card-box">
            <div class="section-title">廠商占比（前15）</div>
            <canvas id="poPieChart" style="max-height:380px;"></canvas>
          </div>
        </div>
      </div>

      <div class="row g-3 mb-3">
        <!-- 廠商排行 -->
        <div class="col-12 col-lg-5">
          <div class="card-box">
            <div class="section-title">廠商採購金額排行</div>
            <div style="overflow-y:auto;max-height:420px;">
              <table class="table table-sm table-hover" style="font-size:.82rem;">
                <thead class="table-dark sticky-top">
                  <tr><th>廠商代號</th><th>廠商名稱</th><th class="text-end">採購金額</th><th class="text-end">筆數</th></tr>
                </thead>
                <tbody>{po_vendor_rows}</tbody>
              </table>
            </div>
          </div>
        </div>
        <!-- 料號排行 -->
        <div class="col-12 col-lg-7">
          <div class="card-box">
            <div class="section-title">進料料號金額排行（前50）</div>
            <div style="overflow-y:auto;max-height:420px;">
              <table class="table table-sm table-hover" style="font-size:.82rem;">
                <thead class="table-dark sticky-top">
                  <tr><th>料號</th><th>品名</th><th class="text-end">數量</th><th class="text-end">採購金額</th></tr>
                </thead>
                <tbody>{po_part_rows}</tbody>
              </table>
            </div>
          </div>
        </div>
      </div>

      <!-- Pivot 表 -->
      <div class="card-box">
        <div class="section-title">月別 × 廠商 Pivot（採購金額）</div>
        {po_pivot_html}
      </div>

    </div><!-- /tab1 -->

    <!-- ════ Tab 2：MOCR34 委外製程成本 ════ -->
    <div class="tab-pane fade" id="tab2">

      <div class="row g-3 mb-3">
        <!-- 月別委外金額 -->
        <div class="col-12 col-lg-7">
          <div class="card-box">
            <div class="section-title">月別委外製程金額趨勢</div>
            <canvas id="moBarChart" style="max-height:340px;"></canvas>
          </div>
        </div>
        <!-- 月別表格 -->
        <div class="col-12 col-lg-5">
          <div class="card-box">
            <div class="section-title">月別委外金額</div>
            <table class="table table-sm table-hover" style="font-size:.82rem;">
              <thead class="table-dark">
                <tr><th>月份</th><th class="text-end">委外淨金額（NT$）</th></tr>
              </thead>
              <tbody>{mo_monthly_rows}</tbody>
            </table>
          </div>
        </div>
      </div>

      <div class="row g-3 mb-3">
        <!-- 產品金額排行 -->
        <div class="col-12 col-lg-7">
          <div class="card-box">
            <div class="section-title">委外產品金額排行（前50）</div>
            <div style="overflow-y:auto;max-height:400px;">
              <table class="table table-sm table-hover" style="font-size:.82rem;">
                <thead class="table-dark sticky-top">
                  <tr><th>品號</th><th>品名</th><th class="text-end">製令數</th><th class="text-end">委外淨金額</th></tr>
                </thead>
                <tbody>{mo_product_rows}</tbody>
              </table>
            </div>
          </div>
        </div>
        <!-- 委外廠商 -->
        <div class="col-12 col-lg-5">
          <div class="card-box">
            <div class="section-title">委外廠商（途程）</div>
            <div style="overflow-y:auto;max-height:400px;">
              <table class="table table-sm table-hover" style="font-size:.82rem;">
                <thead class="table-dark sticky-top">
                  <tr><th>廠商代號</th><th>廠商名稱</th><th class="text-end">關聯製令數</th></tr>
                </thead>
                <tbody>{rt_vendor_rows}</tbody>
              </table>
            </div>
          </div>
        </div>
      </div>

    </div><!-- /tab2 -->

  </div><!-- tab-content -->
</div>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>
<script>
function fmtY(v){{
  if(v>=1e6) return '$'+(v/1e6).toFixed(1)+'M';
  if(v>=1e4) return '$'+(v/1e4).toFixed(0)+'萬';
  return '$'+v.toLocaleString();
}}

// PURR16 堆疊長條
new Chart(document.getElementById('poBarChart'),{{
  type:'bar',
  data:{{labels:{js_po_months}, datasets:{js_po_datasets}}},
  options:{{
    responsive:true, maintainAspectRatio:true,
    plugins:{{
      legend:{{position:'right',labels:{{font:{{family:'Microsoft JhengHei'}},boxWidth:12,padding:6}}}},
      tooltip:{{callbacks:{{label:c=>`${{c.dataset.label}}: NT$ ${{c.parsed.y.toLocaleString()}}`}}}}
    }},
    scales:{{
      x:{{stacked:true,ticks:{{font:{{family:'Microsoft JhengHei'}}}}}},
      y:{{stacked:true,ticks:{{callback:fmtY,font:{{family:'Microsoft JhengHei'}}}},grid:{{color:'#f0f0f0'}}}}
    }}
  }}
}});

// PURR16 廠商圓餅
new Chart(document.getElementById('poPieChart'),{{
  type:'doughnut',
  data:{{labels:{js_pie_labels},datasets:[{{data:{js_pie_data},backgroundColor:{js_pie_colors},borderWidth:2}}]}},
  options:{{
    responsive:true,
    plugins:{{
      legend:{{position:'bottom',labels:{{font:{{family:'Microsoft JhengHei'}},boxWidth:12,padding:6}}}},
      tooltip:{{callbacks:{{label:c=>`${{c.label}}: NT$ ${{c.parsed.toLocaleString()}}`}}}}
    }}
  }}
}});

// MOCR34 月別長條
new Chart(document.getElementById('moBarChart'),{{
  type:'bar',
  data:{{labels:{js_mo_months},datasets:[{{
    label:'委外淨金額',
    data:{js_mo_data},
    backgroundColor:'rgba(230,81,0,0.75)',
    borderColor:'#e65100',borderWidth:2,borderRadius:4
  }}]}},
  options:{{
    responsive:true,
    plugins:{{
      legend:{{display:false}},
      tooltip:{{callbacks:{{label:c=>`NT$ ${{c.parsed.y.toLocaleString()}}`}}}}
    }},
    scales:{{
      x:{{ticks:{{font:{{family:'Microsoft JhengHei'}}}}}},
      y:{{ticks:{{callback:fmtY,font:{{family:'Microsoft JhengHei'}}}},grid:{{color:'#f0f0f0'}}}}
    }}
  }}
}});
</script>
</body>
</html>"""

    with open(path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"[OK] HTML：{path}")
    return path


# ── Excel ─────────────────────────────────────────────────────
def export_excel(df_po: pd.DataFrame, df_mo: pd.DataFrame, df_rt: pd.DataFrame) -> str:
    os.makedirs("output", exist_ok=True)
    tag  = f"{DATE_FROM.replace('-','')}_{DATE_TO.replace('-','')}"
    path = f"output/成本分析_{tag}.xlsx"
    with pd.ExcelWriter(path, engine='openpyxl') as w:
        df_po.to_excel(w, sheet_name='PURR16_採購明細', index=False)
        (df_po.groupby(['vendor_code','vendor_name'])
              .agg(採購金額=('amount','sum'), 筆數=('order_no','count'))
              .reset_index().sort_values('採購金額', ascending=False)
              .to_excel(w, sheet_name='PURR16_廠商彙總', index=False))
        (df_po.groupby(['product_code','product_name'])
              .agg(採購金額=('amount','sum'), 數量=('qty','sum'))
              .reset_index().sort_values('採購金額', ascending=False)
              .to_excel(w, sheet_name='PURR16_料號彙總', index=False))
        df_mo.to_excel(w, sheet_name='MOCR34_製令明細', index=False)
        (df_mo.groupby(['product_code','product_name'])
              .agg(委外金額=('outsource_net_amount','sum'), 製令數=('order_no','count'))
              .reset_index().sort_values('委外金額', ascending=False)
              .to_excel(w, sheet_name='MOCR34_產品彙總', index=False))
        if not df_rt.empty:
            df_rt.to_excel(w, sheet_name='MOCR34_途程廠商', index=False)
    print(f"[OK] Excel：{path}")
    return path


# ── 主程式 ────────────────────────────────────────────────────
if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--explore", action="store_true")
    parser.add_argument("--date-from", default=DATE_FROM)
    parser.add_argument("--date-to",   default=DATE_TO)
    args = parser.parse_args()

    if args.date_from != DATE_FROM: os.environ["COST_DATE_FROM"] = args.date_from
    if args.date_to   != DATE_TO:   os.environ["COST_DATE_TO"]   = args.date_to

    if args.explore:
        conn = get_conn()
        for v in ["v_ht_purchase_order","v_ht_manufacturing_order","v_ht_mo_routing"]:
            print(f"\n== {v} ==")
            df = pd.read_sql(
                f"SELECT COLUMN_NAME,DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS "
                f"WHERE TABLE_NAME='{v}' ORDER BY ORDINAL_POSITION", conn)
            print(df.to_string(index=False))
        conn.close()
    else:
        df_po          = fetch_purr16()
        df_mo, df_rt   = fetch_mocr34()
        if df_po.empty and df_mo.empty:
            print("[WARN] 查無資料")
        else:
            export_excel(df_po, df_mo, df_rt)
            export_html(df_po, df_mo, df_rt)
