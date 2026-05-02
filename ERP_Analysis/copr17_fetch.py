"""
COPR17 — 訂單預計出貨明細表
輸出：Excel + HTML 互動儀表板
"""
import os, sys, warnings, argparse, json
from datetime import date, timedelta
import pyodbc, pandas as pd
from dotenv import load_dotenv

warnings.filterwarnings('ignore')
load_dotenv()

# ── CONFIG ────────────────────────────────────────────────────
TABLE_NAME  = os.getenv("ERP_COPR17_TABLE",  "v_ht_customer_order_lines")
USD_TO_NTD  = float(os.getenv("ERP_USD_RATE", "29.5"))
ORDER_TYPES = os.getenv("ERP_ORDER_TYPES", "2201,2202,2203").split(",")
DATE_FROM   = os.getenv("ERP_DATE_FROM", "2025-12-25")
DATE_TO     = os.getenv("ERP_DATE_TO",   "2026-12-31")

C = {
    "doc_type":      "doc_type",
    "order_no":      "order_no",
    "order_seq":     "line_no",
    "customer_code": "customer_code",
    "customer":      "customer_name",
    "order_date":    "order_date",
    "item_code":     "part_no",
    "item_name":     "product_name",
    "unit":          "unit",
    "qty_ordered":   "order_qty",
    "delivery_date": "delivery_date",
    "currency":      "currency",
    "unit_price":    "unit_price",
    "amount":        "amount_ntd",
}

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
    d = os.getenv("ERP_ODBC_DRIVER","SQL Server")
    s = os.getenv("ERP_SERVER","");  p = os.getenv("ERP_PORT","1433")
    db= os.getenv("ERP_DATABASE",""); u=os.getenv("ERP_USER",""); pw=os.getenv("ERP_PASSWORD","")
    miss = [k for k,v in {"ERP_SERVER":s,"ERP_DATABASE":db,"ERP_USER":u,"ERP_PASSWORD":pw}.items() if not v]
    if miss: print(f"[ERROR] .env 缺少：{miss}"); sys.exit(1)
    return f"DRIVER={{{d}}};SERVER={s},{p};DATABASE={db};UID={u};PWD={pw};TrustServerCertificate=yes;timeout=30;"

def get_conn(): return pyodbc.connect(_conn_str())

def this_week():
    today = date.today()
    start = today - timedelta(days=today.weekday())
    return str(start), str(start + timedelta(days=6))


# ── 撈資料 ────────────────────────────────────────────────────
def fetch_copr17() -> pd.DataFrame:
    sel   = ", ".join(f"{v} AS {k}" for k,v in C.items())
    ph    = ",".join("?"*len(ORDER_TYPES))
    sql   = (f"SELECT {sel} FROM {TABLE_NAME}\n"
             f"WHERE {C['delivery_date']} >= ? AND {C['delivery_date']} <= ?\n"
             f"  AND {C['doc_type']} IN ({ph})\n"
             f"ORDER BY {C['delivery_date']}, {C['order_no']}, {C['order_seq']}")
    params = [DATE_FROM, DATE_TO] + ORDER_TYPES
    print(f"[INFO] 交貨日期 {DATE_FROM}~{DATE_TO}，doc_type={ORDER_TYPES}")
    conn = get_conn()
    try:    df = pd.read_sql(sql, conn, params=params)
    finally: conn.close()
    for c in df.select_dtypes(include='str').columns:
        df[c] = df[c].str.rstrip()
    df['delivery_date'] = pd.to_datetime(df['delivery_date'], errors='coerce')
    df['order_date']    = pd.to_datetime(df['order_date'],    errors='coerce')
    df['amount_twd']    = df.apply(
        lambda r: r['amount']*USD_TO_NTD if str(r.get('currency','')).strip()=='USD' else r['amount'], axis=1)
    df['月份'] = df['delivery_date'].dt.strftime('%Y-%m')
    print(f"[OK] 撈到 {len(df):,} 筆")
    return df


# ── Excel ─────────────────────────────────────────────────────
def export_excel(df: pd.DataFrame, week_start: str, week_end: str) -> str:
    os.makedirs("output", exist_ok=True)
    tag  = f"{DATE_FROM.replace('-','')}_{DATE_TO.replace('-','')}"
    path = f"output/COPR17_訂單預計出貨明細_{tag}.xlsx"
    df_week = df[(df['order_date']>=week_start)&(df['order_date']<=week_end)].copy()
    with pd.ExcelWriter(path, engine='openpyxl') as w:
        detail_cols = [c for c in ['doc_type','order_no','order_seq','customer_code','customer',
            'order_date','item_code','item_name','unit','qty_ordered','delivery_date',
            'currency','unit_price','amount','amount_twd','月份'] if c in df.columns]
        df[detail_cols].to_excel(w, sheet_name='訂單預計出貨明細', index=False)
        mc=(df.groupby(['月份','customer_code','customer'])['amount_twd'].sum().reset_index()
              .rename(columns={'amount_twd':'台幣金額'})
              .sort_values(['月份','台幣金額'],ascending=[True,False]))
        mc.to_excel(w, sheet_name='月別×客戶彙總', index=False)
        (df.groupby('月份').agg(訂單筆數=('order_no','count'),台幣金額=('amount_twd','sum'))
           .reset_index().sort_values('月份').to_excel(w, sheet_name='月別彙總', index=False))
        (df.groupby(['customer_code','customer']).agg(訂單筆數=('order_no','count'),台幣金額=('amount_twd','sum'))
           .reset_index().sort_values('台幣金額',ascending=False).to_excel(w, sheet_name='客戶彙總', index=False))
        if not df_week.empty:
            wc=[c for c in ['order_date','order_no','order_seq','customer_code','customer',
                'item_code','item_name','unit','qty_ordered','delivery_date',
                'currency','unit_price','amount_twd'] if c in df_week.columns]
            df_week[wc].sort_values(['customer','amount_twd'],ascending=[True,False]).to_excel(w,sheet_name='本週新單',index=False)
    print(f"[OK] Excel：{path}")
    return path


# ── HTML 儀表板 ───────────────────────────────────────────────
def fmt_twd(v: float) -> str:
    """格式化台幣（萬元）"""
    if v >= 1_000_000: return f"{v/1_000_000:.1f}M"
    if v >= 10_000:    return f"{v/10_000:.1f}萬"
    return f"{v:,.0f}"

def export_html(df: pd.DataFrame, week_start: str, week_end: str) -> str:
    os.makedirs("output", exist_ok=True)
    tag  = f"{DATE_FROM.replace('-','')}_{DATE_TO.replace('-','')}"
    path = f"output/COPR17_儀表板_{tag}.html"

    # ── 資料準備 ──
    months   = sorted(df['月份'].dropna().unique().tolist())
    # 客戶依總金額大到小排序
    cust_order = (df.groupby('customer')['amount_twd'].sum()
                    .sort_values(ascending=False).index.tolist())

    # pivot: 月份 × 客戶
    pivot = (df.groupby(['月份','customer'])['amount_twd'].sum()
               .unstack(fill_value=0)
               .reindex(index=months, columns=cust_order, fill_value=0))

    # Chart.js datasets
    datasets = []
    for i, cust in enumerate(cust_order):
        color = CHART_COLORS[i % len(CHART_COLORS)]
        datasets.append({
            "label": cust,
            "data":  [round(float(pivot.loc[m, cust]), 0) if m in pivot.index else 0 for m in months],
            "backgroundColor": color,
            "stack": "s",
        })

    # 月別合計
    monthly_totals = [round(float(pivot.loc[m].sum()), 0) if m in pivot.index else 0 for m in months]

    # 總覽數字
    total_twd   = df['amount_twd'].sum()
    n_customers = df['customer'].nunique()
    n_orders    = df['order_no'].nunique()

    # 本週新單
    df_week = df[(df['order_date']>=week_start)&(df['order_date']<=week_end)].copy()
    week_total  = df_week['amount_twd'].sum()
    week_count  = len(df_week)

    # 本週展開資料（依客戶分組）
    week_by_cust: dict = {}
    for cust, grp in df_week.groupby('customer'):
        rows = []
        for _, r in grp.iterrows():
            rows.append({
                "order_no":    str(r['order_no']).strip(),
                "item_code":   str(r['item_code']).strip(),
                "item_name":   str(r['item_name']).strip(),
                "qty":         int(r['qty_ordered']) if pd.notna(r['qty_ordered']) else 0,
                "currency":    str(r['currency']).strip(),
                "unit_price":  float(r['unit_price']) if pd.notna(r['unit_price']) else 0,
                "amount_twd":  round(float(r['amount_twd']), 0),
                "delivery":    str(r['delivery_date'].date()) if pd.notna(r['delivery_date']) else "",
            })
        week_by_cust[cust] = {
            "total": round(float(grp['amount_twd'].sum()), 0),
            "rows":  rows,
        }
    week_by_cust_sorted = dict(sorted(week_by_cust.items(), key=lambda x: -x[1]['total']))

    # 客戶彙總（月份展開）
    cust_monthly = {}
    for cust in cust_order:
        cust_df = df[df['customer']==cust]
        monthly_vals = {}
        for m in months:
            v = float(pivot.loc[m, cust]) if m in pivot.index else 0
            if v > 0: monthly_vals[m] = round(v, 0)
        cust_monthly[cust] = {
            "total": round(float(cust_df['amount_twd'].sum()), 0),
            "monthly": monthly_vals,
        }

    # pivot 表 HTML
    def make_pivot_table():
        th_style = 'style="background:#1b5e20;color:#fff;padding:8px 12px;white-space:nowrap;text-align:right;border:1px solid #388e3c;"'
        th_first = 'style="background:#1b5e20;color:#fff;padding:8px 12px;white-space:nowrap;text-align:left;border:1px solid #388e3c;"'
        td_style = 'style="padding:6px 12px;text-align:right;border:1px solid #e0e0e0;white-space:nowrap;"'
        td_month = 'style="padding:6px 12px;font-weight:600;border:1px solid #e0e0e0;background:#f5f5f5;"'
        td_total_row='style="padding:6px 12px;text-align:right;border:1px solid #e0e0e0;background:#e8f5e9;font-weight:700;"'
        td_total_col='style="padding:6px 12px;text-align:right;border:1px solid #388e3c;background:#c8e6c9;font-weight:700;"'

        cols_html = "".join(f'<th {th_style}>{c}</th>' for c in cust_order)
        header = f'<thead><tr><th {th_first}>月份</th>{cols_html}<th {th_style}>月合計</th></tr></thead>'

        rows_html = ""
        grand_cols = {c: 0 for c in cust_order}
        grand_total = 0
        for m in months:
            row_total = 0
            cells = ""
            for c in cust_order:
                v = float(pivot.loc[m, c]) if m in pivot.index else 0
                grand_cols[c] += v; row_total += v
                txt = f"{v:,.0f}" if v else "-"
                cells += f'<td {td_style}>{txt}</td>'
            grand_total += row_total
            rows_html += f'<tr><td {td_month}>{m}</td>{cells}<td {td_total_col}>{row_total:,.0f}</td></tr>'

        # 總計列
        grand_cells = "".join(f'<td {td_total_row}>{grand_cols[c]:,.0f}</td>' for c in cust_order)
        rows_html += f'<tr><td {td_total_row}>總計</td>{grand_cells}<td {td_total_row}>{grand_total:,.0f}</td></tr>'

        return f'<table class="table table-sm table-bordered" style="font-size:0.82rem;">{header}<tbody>{rows_html}</tbody></table>'

    pivot_table_html = make_pivot_table()

    # JSON for JS
    js_months   = json.dumps(months,   ensure_ascii=False)
    js_datasets = json.dumps(datasets, ensure_ascii=False)
    js_monthly_totals = json.dumps(monthly_totals)
    js_week     = json.dumps(week_by_cust_sorted, ensure_ascii=False)
    js_cust_monthly = json.dumps(cust_monthly, ensure_ascii=False)

    today_str = date.today().strftime("%Y/%m/%d")

    # 本週展開卡片
    week_accordion = ""
    for i, (cust, info) in enumerate(week_by_cust_sorted.items()):
        color = CHART_COLORS[i % len(CHART_COLORS)]
        rows_html = "".join(
            f'<tr><td>{r["order_no"]}</td><td>{r["item_code"]}</td><td>{r["item_name"]}</td>'
            f'<td class="text-end">{r["qty"]:,}</td><td>{r["currency"]}</td>'
            f'<td class="text-end">{r["unit_price"]:,.3f}</td>'
            f'<td class="text-end fw-bold">{r["amount_twd"]:,.0f}</td>'
            f'<td>{r["delivery"]}</td></tr>'
            for r in info['rows']
        )
        week_accordion += f"""
        <div class="accordion-item">
          <h2 class="accordion-header">
            <button class="accordion-button {'collapsed' if i>0 else ''}" type="button"
              data-bs-toggle="collapse" data-bs-target="#wk{i}">
              <span class="badge me-2" style="background:{color}">●</span>
              <strong>{cust}</strong>
              <span class="ms-3 text-muted" style="font-size:.85rem;">{len(info['rows'])} 筆</span>
              <span class="ms-auto fw-bold" style="color:{color}">NT$ {info['total']:,.0f}</span>
            </button>
          </h2>
          <div id="wk{i}" class="accordion-collapse collapse {'show' if i==0 else ''}">
            <div class="accordion-body p-2">
              <table class="table table-sm table-hover mb-0" style="font-size:.82rem;">
                <thead class="table-dark">
                  <tr><th>訂單號</th><th>品號</th><th>品名</th><th class="text-end">數量</th>
                      <th>幣別</th><th class="text-end">單價</th><th class="text-end">台幣金額</th><th>交貨日</th></tr>
                </thead>
                <tbody>{rows_html}</tbody>
              </table>
            </div>
          </div>
        </div>"""

    # 客戶月份展開卡片
    cust_accordion = ""
    for i, (cust, info) in enumerate(cust_monthly.items()):
        if info['total'] == 0: continue
        color = CHART_COLORS[i % len(CHART_COLORS)]
        month_rows = "".join(
            f'<tr><td>{m}</td><td class="text-end fw-bold">NT$ {v:,.0f}</td></tr>'
            for m, v in info['monthly'].items()
        )
        cust_accordion += f"""
        <div class="accordion-item">
          <h2 class="accordion-header">
            <button class="accordion-button collapsed" type="button"
              data-bs-toggle="collapse" data-bs-target="#c{i}">
              <span class="badge me-2" style="background:{color}">●</span>
              <strong>{cust}</strong>
              <span class="ms-auto fw-bold" style="color:{color}">NT$ {info['total']:,.0f}</span>
            </button>
          </h2>
          <div id="c{i}" class="accordion-collapse collapse">
            <div class="accordion-body p-2">
              <table class="table table-sm mb-0" style="font-size:.82rem;max-width:300px;">
                <thead class="table-dark"><tr><th>月份</th><th class="text-end">台幣金額</th></tr></thead>
                <tbody>{month_rows}</tbody>
              </table>
            </div>
          </div>
        </div>"""

    # ── HTML ──
    html = f"""<!DOCTYPE html>
<html lang="zh-Hant">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>COPR17 訂單預計出貨明細 — 儀表板</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.3/dist/chart.umd.min.js"></script>
<style>
  body{{font-family:"Microsoft JhengHei",Arial,sans-serif;background:#f4f6f9;}}
  .header{{background:linear-gradient(135deg,#1b5e20 0%,#2e7d32 60%,#388e3c 100%);
           color:#fff;padding:24px 32px;border-radius:0 0 16px 16px;margin-bottom:24px;}}
  .header h1{{font-size:1.5rem;font-weight:700;margin:0;}}
  .header .sub{{font-size:.85rem;opacity:.8;margin-top:4px;}}
  .kpi-card{{background:#fff;border-radius:12px;padding:20px 24px;box-shadow:0 2px 8px rgba(0,0,0,.07);
             border-left:5px solid;}}
  .kpi-card .val{{font-size:1.7rem;font-weight:700;}}
  .kpi-card .lbl{{font-size:.8rem;color:#666;margin-top:2px;}}
  .nav-tabs .nav-link{{color:#555;font-weight:600;}}
  .nav-tabs .nav-link.active{{color:#1b5e20;border-bottom:3px solid #1b5e20;}}
  .chart-wrap{{background:#fff;border-radius:12px;padding:20px;box-shadow:0 2px 8px rgba(0,0,0,.07);}}
  .pivot-wrap{{background:#fff;border-radius:12px;padding:16px;box-shadow:0 2px 8px rgba(0,0,0,.07);overflow-x:auto;}}
  .accordion-button:not(.collapsed){{background:#e8f5e9;color:#1b5e20;}}
  .section-title{{font-size:1rem;font-weight:700;color:#1b5e20;margin-bottom:12px;
                  border-left:4px solid #1b5e20;padding-left:10px;}}
  .badge-date{{background:#e8f5e9;color:#1b5e20;border-radius:6px;padding:4px 10px;font-size:.8rem;}}
</style>
</head>
<body>

<div class="header">
  <h1>COPR17 訂單預計出貨明細 — 分析儀表板</h1>
  <div class="sub">
    交貨日期：{DATE_FROM} ～ {DATE_TO} ｜
    統計單別：{', '.join(ORDER_TYPES)} ｜
    USD 匯率：{USD_TO_NTD} ｜
    更新：{today_str}
  </div>
</div>

<div class="container-fluid px-4">

  <!-- KPI 卡片 -->
  <div class="row g-3 mb-4">
    <div class="col-6 col-md-3">
      <div class="kpi-card" style="border-color:#1b5e20;">
        <div class="val text-success">{fmt_twd(total_twd)}</div>
        <div class="lbl">台幣總金額（NT$）</div>
      </div>
    </div>
    <div class="col-6 col-md-3">
      <div class="kpi-card" style="border-color:#1565c0;">
        <div class="val" style="color:#1565c0;">{n_customers}</div>
        <div class="lbl">客戶數</div>
      </div>
    </div>
    <div class="col-6 col-md-3">
      <div class="kpi-card" style="border-color:#f9a825;">
        <div class="val" style="color:#f57f17;">{n_orders:,}</div>
        <div class="lbl">訂單張數</div>
      </div>
    </div>
    <div class="col-6 col-md-3">
      <div class="kpi-card" style="border-color:#c62828;">
        <div class="val" style="color:#c62828;">{week_count}</div>
        <div class="lbl">本週新單筆數 <span class="badge-date">{week_start}~{week_end}</span></div>
      </div>
    </div>
  </div>

  <!-- Tab -->
  <ul class="nav nav-tabs mb-3" id="mainTab">
    <li class="nav-item"><a class="nav-link active" data-bs-toggle="tab" href="#tab1">月別×客戶</a></li>
    <li class="nav-item"><a class="nav-link" data-bs-toggle="tab" href="#tab2">客戶彙總</a></li>
    <li class="nav-item"><a class="nav-link" data-bs-toggle="tab" href="#tab3">本週新單
      <span class="badge rounded-pill bg-danger ms-1">{week_count}</span></a></li>
  </ul>

  <div class="tab-content">

    <!-- Tab 1 月別×客戶 -->
    <div class="tab-pane fade show active" id="tab1">
      <div class="row g-3 mb-3">
        <div class="col-12">
          <div class="chart-wrap">
            <div class="section-title">月別×客戶 台幣金額（堆疊長條圖）</div>
            <canvas id="barChart" style="max-height:420px;"></canvas>
          </div>
        </div>
      </div>
      <div class="pivot-wrap">
        <div class="section-title">月別×客戶 Pivot（台幣）</div>
        {pivot_table_html}
      </div>
    </div>

    <!-- Tab 2 客戶彙總 -->
    <div class="tab-pane fade" id="tab2">
      <div class="row g-3 mb-3">
        <div class="col-12 col-lg-5">
          <div class="chart-wrap">
            <div class="section-title">客戶金額占比</div>
            <canvas id="pieChart" style="max-height:380px;"></canvas>
          </div>
        </div>
        <div class="col-12 col-lg-7">
          <div class="chart-wrap">
            <div class="section-title">月別趨勢</div>
            <canvas id="lineChart" style="max-height:380px;"></canvas>
          </div>
        </div>
      </div>
      <div class="accordion" id="custAccordion">
        <div class="section-title">客戶月份展開</div>
        {cust_accordion}
      </div>
    </div>

    <!-- Tab 3 本週新單 -->
    <div class="tab-pane fade" id="tab3">
      <div class="alert alert-success d-flex align-items-center mb-3" style="border-radius:10px;">
        <span>本週（{week_start} ～ {week_end}）共 <strong>{week_count}</strong> 筆新單，
        台幣合計 <strong>NT$ {week_total:,.0f}</strong></span>
      </div>
      {'<div class="text-center text-muted py-5">本週無新下單資料</div>' if not week_by_cust_sorted else
      f'<div class="accordion" id="weekAccordion">{week_accordion}</div>'}
    </div>

  </div><!-- tab-content -->
</div><!-- container -->

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>
<script>
const months   = {js_months};
const datasets = {js_datasets};
const monthly  = {js_monthly_totals};
const weekData = {js_week};
const custMonthly = {js_cust_monthly};

// 格式化軸標籤
function fmtY(v){{
  if(v>=1e6) return '$'+(v/1e6).toFixed(1)+'M';
  if(v>=1e4) return '$'+(v/1e4).toFixed(0)+'萬';
  return '$'+v.toLocaleString();
}}

// 1. 堆疊長條圖
new Chart(document.getElementById('barChart'),{{
  type:'bar',
  data:{{labels:months, datasets}},
  options:{{
    responsive:true, maintainAspectRatio:true,
    plugins:{{
      legend:{{position:'right', labels:{{font:{{family:'Microsoft JhengHei'}},boxWidth:12,padding:8}}}},
      tooltip:{{
        callbacks:{{
          label:ctx=>`${{ctx.dataset.label}}: NT$ ${{ctx.parsed.y.toLocaleString()}}`
        }}
      }}
    }},
    scales:{{
      x:{{stacked:true, ticks:{{font:{{family:'Microsoft JhengHei'}}}}}},
      y:{{stacked:true, ticks:{{callback:fmtY, font:{{family:'Microsoft JhengHei'}}}},
         grid:{{color:'#f0f0f0'}}}}
    }}
  }}
}});

// 2. 圓餅圖（客戶占比）
const custTotals = Object.entries(custMonthly).map(([k,v])=>v.total).filter(v=>v>0);
const custLabels = Object.entries(custMonthly).filter(([k,v])=>v.total>0).map(([k])=>k);
const pieColors  = {json.dumps(CHART_COLORS)}.slice(0,custLabels.length);
new Chart(document.getElementById('pieChart'),{{
  type:'doughnut',
  data:{{labels:custLabels, datasets:[{{data:custTotals,backgroundColor:pieColors,borderWidth:2}}]}},
  options:{{
    responsive:true,
    plugins:{{
      legend:{{position:'right',labels:{{font:{{family:'Microsoft JhengHei'}},boxWidth:12}}}},
      tooltip:{{callbacks:{{label:ctx=>`${{ctx.label}}: NT$ ${{ctx.parsed.toLocaleString()}}`}}}}
    }}
  }}
}});

// 3. 月別趨勢折線圖
new Chart(document.getElementById('lineChart'),{{
  type:'bar',
  data:{{
    labels:months,
    datasets:[{{
      label:'月別台幣合計',
      data:monthly,
      backgroundColor:'rgba(27,94,32,0.75)',
      borderColor:'#1b5e20',
      borderWidth:2,
      borderRadius:4,
    }}]
  }},
  options:{{
    responsive:true,
    plugins:{{
      legend:{{display:false}},
      tooltip:{{callbacks:{{label:ctx=>`NT$ ${{ctx.parsed.y.toLocaleString()}}`}}}}
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


# ── 主程式 ────────────────────────────────────────────────────
if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--explore", action="store_true")
    args = parser.parse_args()
    if args.explore:
        conn = get_conn()
        df_c = pd.read_sql(
            f"SELECT COLUMN_NAME,DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS "
            f"WHERE TABLE_NAME='{TABLE_NAME}' ORDER BY ORDINAL_POSITION", conn)
        print(df_c.to_string(index=False)); conn.close()
    else:
        df = fetch_copr17()
        if df.empty:
            print("[WARN] 查無資料")
        else:
            ws, we = this_week()
            export_excel(df, ws, we)
            export_html(df, ws, we)
            print(f"\n本週（{ws}~{we}）新單：{len(df[(df['order_date']>=ws)&(df['order_date']<=we)]):,} 筆")
