"""
鴻騰電子 — 業績 + 成本 整合儀表板
Tab 1：月別×客戶（營業額 COPR17）
Tab 2：客戶彙總
Tab 3：本週新單
Tab 4：採購成本（PURR16）
Tab 5：委外製程成本（MOCR34）
"""
import os, sys, warnings, json
from datetime import date, timedelta
import pyodbc, pandas as pd
from dotenv import load_dotenv

warnings.filterwarnings('ignore')
load_dotenv()

# ── CONFIG ────────────────────────────────────────────────────
# 營業額
REV_DATE_FROM  = os.getenv("ERP_DATE_FROM",    "2025-12-25")
REV_DATE_TO    = os.getenv("ERP_DATE_TO",       "2026-12-31")
REV_TYPES      = os.getenv("ERP_ORDER_TYPES",   "2201,2202,2203").split(",")
USD_TO_NTD     = float(os.getenv("ERP_USD_RATE","29.5"))
# 成本
COST_DATE_FROM = os.getenv("COST_DATE_FROM",    "2025-12-01")   # 改為 25年12月起
COST_DATE_TO   = os.getenv("COST_DATE_TO",      "2026-12-31")
PO_TYPES       = os.getenv("PO_ORDER_TYPES",    "3302").split(",")
PO_USD_RATE    = float(os.getenv("PO_USD_RATE", "32.0"))   # 採購美金匯率
# 美金廠商代號清單（amount 欄是 USD，需乘匯率換台幣）
PO_USD_VENDORS = set(os.getenv("PO_USD_VENDORS", "101196").split(","))  # 101196=翼高

# ── 各客戶結帳日（日 > 結帳日 → 歸下月）─────────────────────
DEFAULT_BILLING_DAY = 25
BILLING_DAY_MAP: dict = {
    "KS1": 30,   # 金士頓科技
    "KS":  30,   # 金士頓
}

def calc_billing_month(delivery_date, customer_code: str) -> str:
    if pd.isna(delivery_date): return ""
    bday = BILLING_DAY_MAP.get(str(customer_code).strip(), DEFAULT_BILLING_DAY)
    d = pd.Timestamp(delivery_date)
    if d.day > bday:
        return f"{d.year+1}-01" if d.month == 12 else f"{d.year}-{d.month+1:02d}"
    return f"{d.year}-{d.month:02d}"

COLORS = [
    "#1b5e20","#2e7d32","#388e3c","#43a047","#66bb6a",
    "#0d47a1","#1565c0","#1976d2","#2196f3","#64b5f6",
    "#f57f17","#f9a825","#fbc02d","#fdd835",
    "#b71c1c","#c62828","#d32f2f","#e53935",
    "#4a148c","#6a1b9a","#7b1fa2","#ab47bc",
    "#e65100","#ef6c00","#f57c00","#fb8c00",
    "#006064","#00838f","#0097a7","#26c6da",
]

def _conn():
    d=os.getenv("ERP_ODBC_DRIVER","SQL Server"); s=os.getenv("ERP_SERVER","")
    p=os.getenv("ERP_PORT","1433"); db=os.getenv("ERP_DATABASE","")
    u=os.getenv("ERP_USER",""); pw=os.getenv("ERP_PASSWORD","")
    miss=[k for k,v in{"ERP_SERVER":s,"ERP_DATABASE":db,"ERP_USER":u,"ERP_PASSWORD":pw}.items() if not v]
    if miss: print(f"[ERROR] .env 缺少：{miss}"); sys.exit(1)
    return pyodbc.connect(f"DRIVER={{{d}}};SERVER={s},{p};DATABASE={db};"
                          f"UID={u};PWD={pw};TrustServerCertificate=yes;timeout=30;")

def this_week():
    t=date.today(); s=t-timedelta(days=t.weekday())
    return str(s), str(s+timedelta(days=6))

def rstrip_df(df):
    for c in df.select_dtypes(include='str').columns:
        df[c]=df[c].str.rstrip()
    return df

def fmt(v):
    if v>=1_000_000: return f"{v/1_000_000:.1f}M"
    if v>=10_000:    return f"{v/10_000:.1f}萬"
    return f"{v:,.0f}"

# ── 撈資料 ────────────────────────────────────────────────────
def fetch_revenue():
    ph=",".join("?"*len(REV_TYPES))
    sql=(f"SELECT doc_type,order_no,line_no AS order_seq,customer_code,customer_name AS customer,"
         f"order_date,part_no AS item_code,product_name AS item_name,unit,order_qty AS qty_ordered,"
         f"delivery_date,currency,unit_price,amount_ntd AS amount "
         f"FROM v_ht_customer_order_lines "
         f"WHERE delivery_date>=? AND delivery_date<=? AND doc_type IN ({ph}) "
         f"ORDER BY delivery_date,order_no,line_no")
    conn=_conn()
    try: df=pd.read_sql(sql,conn,params=[REV_DATE_FROM,REV_DATE_TO]+REV_TYPES)
    finally: conn.close()
    rstrip_df(df)
    df['delivery_date']=pd.to_datetime(df['delivery_date'],errors='coerce')
    df['order_date']=pd.to_datetime(df['order_date'],errors='coerce')
    df['amount_twd']=df.apply(lambda r:r['amount']*USD_TO_NTD if str(r.get('currency','')).strip()=='USD' else r['amount'],axis=1)
    # 用結帳年月（依客戶結帳日計算），不用預交日期月份
    df['月份']=df.apply(lambda r: calc_billing_month(r['delivery_date'], r['customer_code']), axis=1)
    print(f"[COPR17] {len(df):,} 筆，結帳日預設={DEFAULT_BILLING_DAY}日")
    return df

def fetch_purchase():
    ph=",".join("?"*len(PO_TYPES))
    sql=(f"SELECT order_no,doc_type,order_date,vendor_code,vendor_name,"
         f"product_code,product_name,qty,unit_price,amount,delivery_date,delivered_qty "
         f"FROM v_ht_purchase_order "
         f"WHERE order_date>=? AND order_date<=? AND doc_type IN ({ph}) "
         f"ORDER BY order_date,vendor_code")
    conn=_conn()
    try: df=pd.read_sql(sql,conn,params=[COST_DATE_FROM,COST_DATE_TO]+PO_TYPES)
    finally: conn.close()
    rstrip_df(df)
    df['order_date']    = pd.to_datetime(df['order_date'],    errors='coerce')
    df['delivery_date'] = pd.to_datetime(df['delivery_date'], errors='coerce')
    df['月份'] = df['order_date'].dt.strftime('%Y-%m')
    # 台幣換算：美金廠商 × PO_USD_RATE，其餘直接用 amount
    df['amount_ntd'] = df.apply(
        lambda r: r['amount'] * PO_USD_RATE
                  if str(r.get('vendor_code','')).strip() in PO_USD_VENDORS
                  else r['amount'],
        axis=1
    )
    df['currency'] = df['vendor_code'].apply(
        lambda v: 'USD' if str(v).strip() in PO_USD_VENDORS else 'NTD'
    )
    usd_rows = (df['vendor_code'].apply(lambda v: str(v).strip() in PO_USD_VENDORS)).sum()
    print(f"[PURR16] {len(df):,} 筆（其中 USD 廠商 {usd_rows} 筆，匯率 {PO_USD_RATE}）")
    return df

def fetch_outsource():
    sql_mo=(f"SELECT doc_type,order_no,order_date,product_code,product_name,"
            f"plan_qty,done_qty,outsource_in_amount,outsource_return_amount,outsource_net_amount "
            f"FROM v_ht_manufacturing_order "
            f"WHERE order_date>=? AND order_date<=? AND outsource_net_amount>0 "
            f"ORDER BY order_date,product_code")
    sql_rt=(f"SELECT r.order_no,r.process_seq,r.process_code,r.process_name,"
            f"r.line_vendor_code,r.line_vendor_name,r.plan_end_date,r.input_qty,r.done_qty "
            f"FROM v_ht_mo_routing r "
            f"INNER JOIN v_ht_manufacturing_order m ON r.order_no=m.order_no "
            f"WHERE m.order_date>=? AND m.order_date<=? AND m.outsource_net_amount>0 "
            f"AND r.line_vendor_code NOT IN ('001','002','003','004','005','006','007','008','009') "
            f"ORDER BY r.order_no,r.process_seq")
    conn=_conn()
    try:
        df_mo=pd.read_sql(sql_mo,conn,params=[COST_DATE_FROM,COST_DATE_TO])
        df_rt=pd.read_sql(sql_rt,conn,params=[COST_DATE_FROM,COST_DATE_TO])
    finally: conn.close()
    for df in [df_mo,df_rt]: rstrip_df(df)
    df_mo['order_date']=pd.to_datetime(df_mo['order_date'],errors='coerce')
    df_mo['月份']=df_mo['order_date'].dt.strftime('%Y-%m')
    print(f"[MOCR34] 製令 {len(df_mo):,} 筆，途程 {len(df_rt):,} 筆")
    return df_mo, df_rt

# ── HTML 元件 ─────────────────────────────────────────────────
def pivot_html(months, vendors, pivot):
    th ='style="background:#1b5e20;color:#fff;padding:6px 10px;white-space:nowrap;text-align:right;border:1px solid #388e3c;"'
    th0='style="background:#1b5e20;color:#fff;padding:6px 10px;white-space:nowrap;text-align:left;border:1px solid #388e3c;"'
    td ='style="padding:5px 10px;text-align:right;border:1px solid #e0e0e0;white-space:nowrap;"'
    tdm='style="padding:5px 10px;font-weight:600;border:1px solid #e0e0e0;background:#f5f5f5;"'
    tdt='style="padding:5px 10px;text-align:right;border:1px solid #388e3c;background:#c8e6c9;font-weight:700;"'
    tdg='style="padding:5px 10px;text-align:right;border:1px solid #e0e0e0;background:#e8f5e9;font-weight:700;"'
    header=f'<thead><tr><th {th0}>月份</th>{"".join(f"<th {th}>{v}</th>" for v in vendors)}<th {th}>月合計</th></tr></thead>'
    body=""; gc={v:0 for v in vendors}; gt=0
    for m in months:
        rt=0; cells=""
        for v in vendors:
            val=float(pivot.loc[m,v]) if m in pivot.index else 0
            gc[v]+=val; rt+=val
            cells+=f'<td {td}>{"−" if val==0 else f"{val:,.0f}"}</td>'
        gt+=rt
        body+=f'<tr><td {tdm}>{m}</td>{cells}<td {tdt}>{rt:,.0f}</td></tr>'
    body+=f'<tr><td {tdg}>總計</td>{"".join(f"<td {tdg}>{gc[v]:,.0f}</td>" for v in vendors)}<td {tdg}>{gt:,.0f}</td></tr>'
    return f'<div style="overflow-x:auto;"><table class="table table-sm table-bordered" style="font-size:.80rem;"><thead>{header}</thead><tbody>{body}</tbody></table></div>'

def po_pivot_html(months, vendors, pivot):
    th ='style="background:#1565c0;color:#fff;padding:6px 10px;white-space:nowrap;text-align:right;border:1px solid #1976d2;"'
    th0='style="background:#1565c0;color:#fff;padding:6px 10px;white-space:nowrap;text-align:left;border:1px solid #1976d2;"'
    td ='style="padding:5px 10px;text-align:right;border:1px solid #e0e0e0;white-space:nowrap;"'
    tdm='style="padding:5px 10px;font-weight:600;border:1px solid #e0e0e0;background:#f5f5f5;"'
    tdt='style="padding:5px 10px;text-align:right;border:1px solid #1976d2;background:#bbdefb;font-weight:700;"'
    tdg='style="padding:5px 10px;text-align:right;border:1px solid #e0e0e0;background:#e3f2fd;font-weight:700;"'
    header=f'<thead><tr><th {th0}>月份</th>{"".join(f"<th {th}>{v}</th>" for v in vendors)}<th {th}>月合計</th></tr></thead>'
    body=""; gc={v:0 for v in vendors}; gt=0
    for m in months:
        rt=0; cells=""
        for v in vendors:
            val=float(pivot.loc[m,v]) if m in pivot.index else 0
            gc[v]+=val; rt+=val
            cells+=f'<td {td}>{"−" if val==0 else f"{val:,.0f}"}</td>'
        gt+=rt
        body+=f'<tr><td {tdm}>{m}</td>{cells}<td {tdt}>{rt:,.0f}</td></tr>'
    body+=f'<tr><td {tdg}>總計</td>{"".join(f"<td {tdg}>{gc[v]:,.0f}</td>" for v in vendors)}<td {tdg}>{gt:,.0f}</td></tr>'
    return f'<div style="overflow-x:auto;"><table class="table table-sm table-bordered" style="font-size:.80rem;"><thead>{header}</thead><tbody>{body}</tbody></table></div>'

# ── 主輸出 ────────────────────────────────────────────────────
def generate(df_rev, df_po, df_mo, df_rt):
    os.makedirs("output", exist_ok=True)
    path = f"output/鴻騰整合儀表板_{date.today().strftime('%Y%m%d')}.html"
    today_str = date.today().strftime("%Y/%m/%d")
    ws, we = this_week()

    # ══ 營業額資料 ══
    rev_months  = sorted(df_rev['月份'].dropna().unique().tolist())
    cust_order  = (df_rev.groupby('customer')['amount_twd'].sum()
                         .sort_values(ascending=False).index.tolist())
    rev_pivot   = (df_rev.groupby(['月份','customer'])['amount_twd'].sum()
                         .unstack(fill_value=0)
                         .reindex(index=rev_months, columns=cust_order, fill_value=0))
    rev_datasets= [{"label":c,"data":[round(float(rev_pivot.loc[m,c]),0) if m in rev_pivot.index else 0 for m in rev_months],
                    "backgroundColor":COLORS[i%len(COLORS)],"stack":"s"}
                   for i,c in enumerate(cust_order)]
    rev_total   = df_rev['amount_twd'].sum()
    rev_n_cust  = df_rev['customer'].nunique()
    rev_n_ord   = df_rev['order_no'].nunique()
    df_week     = df_rev[(df_rev['order_date']>=ws)&(df_rev['order_date']<=we)].copy()
    week_total  = df_week['amount_twd'].sum()
    week_count  = len(df_week)

    # 月別趨勢
    rev_monthly_totals=[round(float(rev_pivot.loc[m].sum()),0) if m in rev_pivot.index else 0 for m in rev_months]

    # 客戶圓餅
    cust_sum=(df_rev.groupby('customer')['amount_twd'].sum().sort_values(ascending=False).head(15))
    pie_labels=cust_sum.index.tolist(); pie_data=[round(float(v),0) for v in cust_sum.values]

    # 客戶月份展開 accordion
    cust_monthly_data={c:{"total":round(float(df_rev[df_rev.customer==c]['amount_twd'].sum()),0),
                          "monthly":{m:round(float(rev_pivot.loc[m,c]),0) for m in rev_months if m in rev_pivot.index and rev_pivot.loc[m,c]>0}}
                       for c in cust_order}
    cust_acc=""
    for i,(c,info) in enumerate(cust_monthly_data.items()):
        if info['total']==0: continue
        col=COLORS[i%len(COLORS)]
        mrows="".join(f'<tr><td>{m}</td><td class="text-end fw-bold">NT$ {v:,.0f}</td></tr>' for m,v in info['monthly'].items())
        cust_acc+=f"""<div class="accordion-item">
          <h2 class="accordion-header"><button class="accordion-button collapsed" type="button"
            data-bs-toggle="collapse" data-bs-target="#rc{i}">
            <span class="badge me-2" style="background:{col}">●</span><strong>{c}</strong>
            <span class="ms-auto fw-bold" style="color:{col}">NT$ {info['total']:,.0f}</span>
          </button></h2>
          <div id="rc{i}" class="accordion-collapse collapse">
            <div class="accordion-body p-2">
              <table class="table table-sm mb-0" style="font-size:.82rem;max-width:280px;">
                <thead class="table-dark"><tr><th>月份</th><th class="text-end">台幣金額</th></tr></thead>
                <tbody>{mrows}</tbody></table></div></div></div>"""

    # 本週 accordion
    week_by_cust={}
    for cust,grp in df_week.groupby('customer'):
        rows=[{"order_no":str(r['order_no']).strip(),"item_code":str(r['item_code']).strip(),
               "item_name":str(r['item_name']).strip(),"qty":int(r['qty_ordered']) if pd.notna(r['qty_ordered']) else 0,
               "currency":str(r['currency']).strip(),"unit_price":float(r['unit_price']) if pd.notna(r['unit_price']) else 0,
               "amount_twd":round(float(r['amount_twd']),0),"delivery":str(r['delivery_date'].date()) if pd.notna(r['delivery_date']) else ""}
              for _,r in grp.iterrows()]
        week_by_cust[cust]={"total":round(float(grp['amount_twd'].sum()),0),"rows":rows}
    week_by_cust=dict(sorted(week_by_cust.items(),key=lambda x:-x[1]['total']))
    week_acc=""
    for i,(c,info) in enumerate(week_by_cust.items()):
        col=COLORS[i%len(COLORS)]
        rrows="".join(f'<tr><td>{r["order_no"]}</td><td>{r["item_code"]}</td><td>{r["item_name"]}</td>'
                      f'<td class="text-end">{r["qty"]:,}</td><td>{r["currency"]}</td>'
                      f'<td class="text-end">{r["unit_price"]:,.3f}</td>'
                      f'<td class="text-end fw-bold">{r["amount_twd"]:,.0f}</td>'
                      f'<td>{r["delivery"]}</td></tr>' for r in info['rows'])
        week_acc+=f"""<div class="accordion-item">
          <h2 class="accordion-header"><button class="accordion-button {'collapsed' if i>0 else ''}" type="button"
            data-bs-toggle="collapse" data-bs-target="#wk{i}">
            <span class="badge me-2" style="background:{col}">●</span><strong>{c}</strong>
            <span class="ms-3 text-muted" style="font-size:.85rem;">{len(info['rows'])} 筆</span>
            <span class="ms-auto fw-bold" style="color:{col}">NT$ {info['total']:,.0f}</span>
          </button></h2>
          <div id="wk{i}" class="accordion-collapse collapse {'show' if i==0 else ''}">
            <div class="accordion-body p-2">
              <table class="table table-sm table-hover mb-0" style="font-size:.82rem;">
                <thead class="table-dark"><tr><th>訂單號</th><th>品號</th><th>品名</th><th class="text-end">數量</th>
                  <th>幣別</th><th class="text-end">單價</th><th class="text-end">台幣金額</th><th>交貨日</th></tr></thead>
                <tbody>{rrows}</tbody></table></div></div></div>"""

    rev_pivot_html = pivot_html(rev_months, cust_order[:20], rev_pivot)

    # ══ 採購成本資料 ══
    po_months_list = sorted(df_po['月份'].dropna().unique().tolist())
    po_vend_order  = (df_po.groupby('vendor_name')['amount_ntd'].sum()
                           .sort_values(ascending=False).index.tolist())
    po_pivot_data  = (df_po.groupby(['月份','vendor_name'])['amount_ntd'].sum()
                           .unstack(fill_value=0)
                           .reindex(index=po_months_list, columns=po_vend_order, fill_value=0))
    po_datasets    = [{"label":v,"data":[round(float(po_pivot_data.loc[m,v]),0) if m in po_pivot_data.index else 0 for m in po_months_list],
                       "backgroundColor":COLORS[i%len(COLORS)],"stack":"s"}
                      for i,v in enumerate(po_vend_order[:30])]
    po_total       = df_po['amount_ntd'].sum()
    po_n_vendors   = df_po['vendor_name'].nunique()
    po_vendor_sum  = (df_po.groupby(['vendor_code','vendor_name','currency'])
                           .agg(採購金額=('amount_ntd','sum'),筆數=('order_no','count'))
                           .reset_index().sort_values('採購金額',ascending=False))
    po_part_sum    = (df_po.groupby(['product_code','product_name'])
                           .agg(採購金額=('amount_ntd','sum'),數量=('qty','sum'))
                           .reset_index().sort_values('採購金額',ascending=False))
    po_vendor_rows = "".join(
        f'<tr><td>{r.vendor_code}</td><td>{r.vendor_name}</td>'
        f'<td class="text-center"><span class="badge" style="background:{"#e65100" if r.currency=="USD" else "#1565c0"}">'
        f'{r.currency}</span></td>'
        f'<td class="text-end fw-bold">{r.採購金額:,.0f}</td><td class="text-end">{r.筆數}</td></tr>'
        for _,r in po_vendor_sum.head(30).iterrows())
    po_part_rows   = "".join(f'<tr><td>{r.product_code}</td><td>{r.product_name}</td>'
                             f'<td class="text-end">{r.數量:,.0f}</td>'
                             f'<td class="text-end fw-bold">{r.採購金額:,.0f}</td></tr>'
                             for _,r in po_part_sum.head(50).iterrows())
    po_piv_html    = po_pivot_html(po_months_list, po_vend_order[:20], po_pivot_data)
    top_po=po_vendor_sum.head(15)
    po_pie_labels=top_po['vendor_name'].tolist(); po_pie_data=[round(float(v),0) for v in top_po['採購金額'].tolist()]

    # 採購本週下單
    df_po_week    = df_po[(df_po['order_date']>=ws)&(df_po['order_date']<=we)].copy()
    po_week_total = df_po_week['amount_ntd'].sum()
    po_week_count = len(df_po_week)
    po_week_vend_sum = (df_po_week.groupby(['vendor_code','vendor_name'])
                                  .agg(採購金額=('amount_ntd','sum'),筆數=('order_no','count'))
                                  .reset_index().sort_values('採購金額',ascending=False))
    # 採購廠商月份展開 accordion（依金額大到小）
    po_vend_acc = ""
    _po_vend_totals = (df_po.groupby('vendor_name')['amount_ntd']
                            .sum().sort_values(ascending=False))
    for i, vend in enumerate(_po_vend_totals.index):
        grp = df_po[df_po['vendor_name']==vend]
        col = COLORS[i % len(COLORS)]
        vtotal = round(float(grp['amount_ntd'].sum()), 0)
        curr = grp['currency'].iloc[0] if 'currency' in grp.columns else 'NTD'
        mrows = "".join(
            f'<tr><td>{m}</td><td class="text-end fw-bold">NT$ {round(float(v),0):,.0f}</td></tr>'
            for m, v in grp.groupby('月份')['amount_ntd'].sum().sort_index().items()
        )
        po_vend_acc += f"""<div class="accordion-item">
          <h2 class="accordion-header"><button class="accordion-button collapsed" type="button"
            data-bs-toggle="collapse" data-bs-target="#pv{i}">
            <span class="badge me-2" style="background:{col}">●</span><strong>{vend}</strong>
            <span class="ms-2"><span class="badge" style="background:{"#e65100" if curr=="USD" else "#1565c0"};font-size:.7rem">{curr}</span></span>
            <span class="ms-3 text-muted" style="font-size:.85rem;">{len(grp)} 筆</span>
            <span class="ms-auto fw-bold" style="color:{col}">NT$ {vtotal:,.0f}</span>
          </button></h2>
          <div id="pv{i}" class="accordion-collapse collapse">
            <div class="accordion-body p-2">
              <table class="table table-sm mb-0" style="font-size:.82rem;max-width:280px;">
                <thead class="table-dark"><tr><th>月份</th><th class="text-end">台幣金額</th></tr></thead>
                <tbody>{mrows}</tbody></table></div></div></div>"""

    # 本週採購 accordion
    po_week_acc = ""
    for i,(vend,grp) in enumerate(
        sorted(df_po_week.groupby('vendor_name'),
               key=lambda x: -x[1]['amount_ntd'].sum())):
        col=COLORS[i%len(COLORS)]
        vrows="".join(f'<tr><td>{r.order_no}</td><td>{r.product_code}</td><td>{r.product_name}</td>'
                      f'<td class="text-end">{r.qty:,.0f}</td>'
                      f'<td class="text-end">{r.unit_price:,.3f}</td>'
                      f'<td class="text-end fw-bold">{r.amount_ntd:,.0f}</td>'
                      f'<td>{str(pd.Timestamp(r.delivery_date).date()) if pd.notna(r.delivery_date) else ""}</td></tr>'
                      for _,r in grp.sort_values('amount_ntd',ascending=False).iterrows())
        vtotal=round(float(grp['amount_ntd'].sum()),0)
        po_week_acc+=f"""<div class="accordion-item">
          <h2 class="accordion-header"><button class="accordion-button {'collapsed' if i>0 else ''}" type="button"
            data-bs-toggle="collapse" data-bs-target="#pw{i}">
            <span class="badge me-2" style="background:{col}">●</span><strong>{vend}</strong>
            <span class="ms-3 text-muted" style="font-size:.85rem;">{len(grp)} 筆</span>
            <span class="ms-auto fw-bold" style="color:{col}">NT$ {vtotal:,.0f}</span>
          </button></h2>
          <div id="pw{i}" class="accordion-collapse collapse {'show' if i==0 else ''}">
            <div class="accordion-body p-2">
              <table class="table table-sm table-hover mb-0" style="font-size:.82rem;">
                <thead class="table-dark"><tr><th>採購單號</th><th>料號</th><th>品名</th>
                  <th class="text-end">數量</th><th class="text-end">單價</th>
                  <th class="text-end">金額</th><th>交貨日</th></tr></thead>
                <tbody>{vrows}</tbody></table></div></div></div>"""

    # ══ 委外製程資料 ══
    mo_total      = df_mo['outsource_net_amount'].sum()
    mo_n_products = df_mo['product_code'].nunique()
    mo_monthly    = (df_mo.groupby('月份')['outsource_net_amount'].sum()
                          .reset_index().sort_values('月份')
                          .rename(columns={'outsource_net_amount':'委外金額'}))
    mo_product_sum= (df_mo.groupby(['product_code','product_name'])
                          .agg(委外金額=('outsource_net_amount','sum'),製令數=('order_no','count'))
                          .reset_index().sort_values('委外金額',ascending=False))

    # 廠商成本分析：依途程步數比例分配 outsource_net_amount
    if not df_rt.empty:
        # 每張製令的外部步數
        rt_steps = (df_rt.groupby('order_no').size().reset_index(name='total_ext_steps'))
        # 每張製令每個廠商的步數
        rt_vend_steps = (df_rt.groupby(['order_no','line_vendor_code','line_vendor_name'])
                               .size().reset_index(name='vend_steps'))
        rt_vend_steps = rt_vend_steps.merge(rt_steps, on='order_no')
        rt_vend_steps = rt_vend_steps.merge(
            df_mo[['order_no','outsource_net_amount','月份']], on='order_no', how='left')
        rt_vend_steps['委外分配金額'] = (
            rt_vend_steps['vend_steps'] / rt_vend_steps['total_ext_steps']
            * rt_vend_steps['outsource_net_amount'])
        # 廠商彙總
        mo_vend_sum = (rt_vend_steps.groupby(['line_vendor_code','line_vendor_name'])
                                    .agg(委外金額=('委外分配金額','sum'),
                                         關聯製令數=('order_no','nunique'))
                                    .reset_index().sort_values('委外金額',ascending=False))
        # 廠商月份
        mo_vend_monthly = (rt_vend_steps.groupby(['line_vendor_name','月份'])['委外分配金額']
                                        .sum().reset_index())
        mo_vend_months  = sorted(mo_vend_monthly['月份'].dropna().unique().tolist())
        mo_vend_order   = mo_vend_sum['line_vendor_name'].tolist()
        mo_vend_pivot   = (mo_vend_monthly.groupby(['月份','line_vendor_name'])['委外分配金額']
                                          .sum().unstack(fill_value=0)
                                          .reindex(index=mo_vend_months, columns=mo_vend_order, fill_value=0))
        mo_vend_datasets= [{"label":v,
                            "data":[round(float(mo_vend_pivot.loc[m,v]),0) if m in mo_vend_pivot.index else 0
                                    for m in mo_vend_months],
                            "backgroundColor":COLORS[i%len(COLORS)],"stack":"s"}
                           for i,v in enumerate(mo_vend_order[:20])]
        # 廠商月份 accordion
        # 依總金額大到小排序
        _vend_totals = (rt_vend_steps.groupby('line_vendor_name')['委外分配金額']
                                     .sum().sort_values(ascending=False))
        mo_vend_acc=""
        for i,(vend,grp) in enumerate((rt_vend_steps
                                       .set_index('line_vendor_name')
                                       .loc[_vend_totals.index]
                                       .reset_index()
                                       .groupby('line_vendor_name',sort=False))):
            col=COLORS[i%len(COLORS)]
            total=round(float(grp['委外分配金額'].sum()),0)
            mrows="".join(f'<tr><td>{m}</td><td class="text-end fw-bold">NT$ {round(float(v),0):,.0f}</td></tr>'
                          for m,v in grp.groupby('月份')['委外分配金額'].sum().sort_index().items())
            mo_vend_acc+=f"""<div class="accordion-item">
              <h2 class="accordion-header"><button class="accordion-button collapsed" type="button"
                data-bs-toggle="collapse" data-bs-target="#mv{i}">
                <span class="badge me-2" style="background:{col}">●</span><strong>{vend}</strong>
                <span class="ms-3 text-muted" style="font-size:.85rem;">{grp['order_no'].nunique()} 張製令</span>
                <span class="ms-auto fw-bold" style="color:{col}">NT$ {total:,.0f}</span>
              </button></h2>
              <div id="mv{i}" class="accordion-collapse collapse">
                <div class="accordion-body p-2">
                  <table class="table table-sm mb-0" style="font-size:.82rem;max-width:280px;">
                    <thead class="table-dark"><tr><th>月份</th><th class="text-end">分配委外金額</th></tr></thead>
                    <tbody>{mrows}</tbody></table></div></div></div>"""
        mo_vend_rows="".join(
            f'<tr><td>{r.line_vendor_code}</td><td>{r.line_vendor_name}</td>'
            f'<td class="text-end">{r.關聯製令數}</td>'
            f'<td class="text-end fw-bold">{r.委外金額:,.0f}</td></tr>'
            for _,r in mo_vend_sum.head(30).iterrows())
    else:
        mo_vend_sum=pd.DataFrame(); mo_vend_months=[]; mo_vend_order=[]
        mo_vend_pivot=pd.DataFrame(); mo_vend_datasets=[]; mo_vend_acc=""
        mo_vend_rows=""

    mo_monthly_rows ="".join(f'<tr><td>{r.月份}</td><td class="text-end fw-bold">{r.委外金額:,.0f}</td></tr>' for _,r in mo_monthly.iterrows())
    mo_product_rows ="".join(f'<tr><td>{r.product_code}</td><td>{r.product_name}</td>'
                             f'<td class="text-end">{r.製令數}</td>'
                             f'<td class="text-end fw-bold">{r.委外金額:,.0f}</td></tr>'
                             for _,r in mo_product_sum.head(50).iterrows())
    # 廠商 pivot HTML（橘色系）
    def mo_pivot_html_fn(months, vendors, pivot):
        th ='style="background:#e65100;color:#fff;padding:6px 10px;white-space:nowrap;text-align:right;border:1px solid #f57c00;"'
        th0='style="background:#e65100;color:#fff;padding:6px 10px;white-space:nowrap;text-align:left;border:1px solid #f57c00;"'
        td ='style="padding:5px 10px;text-align:right;border:1px solid #e0e0e0;white-space:nowrap;"'
        tdm='style="padding:5px 10px;font-weight:600;border:1px solid #e0e0e0;background:#f5f5f5;"'
        tdt='style="padding:5px 10px;text-align:right;border:1px solid #f57c00;background:#ffe0b2;font-weight:700;"'
        tdg='style="padding:5px 10px;text-align:right;border:1px solid #e0e0e0;background:#fff3e0;font-weight:700;"'
        header=f'<thead><tr><th {th0}>月份</th>{"".join(f"<th {th}>{v}</th>" for v in vendors)}<th {th}>月合計</th></tr></thead>'
        body=""; gc={v:0 for v in vendors}; gt=0
        for m in months:
            rt=0; cells=""
            for v in vendors:
                val=float(pivot.loc[m,v]) if m in pivot.index and v in pivot.columns else 0
                gc[v]+=val; rt+=val
                cells+=f'<td {td}>{"−" if val==0 else f"{val:,.0f}"}</td>'
            gt+=rt
            body+=f'<tr><td {tdm}>{m}</td>{cells}<td {tdt}>{rt:,.0f}</td></tr>'
        body+=f'<tr><td {tdg}>總計</td>{"".join(f"<td {tdg}>{gc[v]:,.0f}</td>" for v in vendors)}<td {tdg}>{gt:,.0f}</td></tr>'
        return f'<div style="overflow-x:auto;"><table class="table table-sm table-bordered" style="font-size:.80rem;"><thead>{header}</thead><tbody>{body}</tbody></table></div>'
    mo_vend_piv_html = mo_pivot_html_fn(mo_vend_months, mo_vend_order[:20], mo_vend_pivot) if mo_vend_months else ""

    po_week_table_html = ""  # 本週廠商小計已移除

    po_week_acc_html = (
        '<div class="text-muted small mb-3">本週無採購下單</div>'
        if not po_week_acc else
        f'<div class="accordion mb-3" id="poWeekAcc">'
        f'<div class="sec sec-blue mb-2">本週採購下單明細（依廠商）</div>'
        f'{po_week_acc}</div>'
    )

    # JSON
    J=lambda x:json.dumps(x,ensure_ascii=False)
    js_rev_months=J(rev_months); js_rev_ds=J(rev_datasets); js_rev_mt=json.dumps(rev_monthly_totals)
    js_pie_l=J(pie_labels); js_pie_d=json.dumps(pie_data); js_pie_c=J(COLORS[:len(pie_labels)])
    js_po_m=J(po_months_list); js_po_ds=J(po_datasets)
    js_po_pie_l=J(po_pie_labels); js_po_pie_d=json.dumps(po_pie_data); js_po_pie_c=J(COLORS[:len(po_pie_labels)])
    js_mo_m=J(mo_monthly['月份'].tolist()); js_mo_d=json.dumps([round(float(v),0) for v in mo_monthly['委外金額'].tolist()])
    js_mo_vend_m=J(mo_vend_months); js_mo_vend_ds=J(mo_vend_datasets)

    html=f"""<!DOCTYPE html>
<html lang="zh-Hant">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>鴻騰電子 — 整合儀表板</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.3/dist/chart.umd.min.js"></script>
<style>
body{{font-family:"Microsoft JhengHei",Arial,sans-serif;background:#f4f6f9;}}
.header{{background:linear-gradient(135deg,#1a237e 0%,#1b5e20 100%);
         color:#fff;padding:22px 32px;border-radius:0 0 16px 16px;margin-bottom:20px;}}
.header h1{{font-size:1.45rem;font-weight:700;margin:0;}}
.header .sub{{font-size:.82rem;opacity:.8;margin-top:4px;}}
.kpi{{background:#fff;border-radius:12px;padding:18px 20px;
      box-shadow:0 2px 8px rgba(0,0,0,.07);border-left:5px solid;}}
.kpi .val{{font-size:1.6rem;font-weight:700;}}
.kpi .lbl{{font-size:.78rem;color:#666;margin-top:2px;}}
.card-box{{background:#fff;border-radius:12px;padding:18px;
           box-shadow:0 2px 8px rgba(0,0,0,.07);margin-bottom:16px;}}
.sec{{font-size:.95rem;font-weight:700;margin-bottom:10px;padding-left:9px;border-left:4px solid;}}
.sec-green{{color:#1b5e20;border-color:#1b5e20;}}
.sec-blue{{color:#1565c0;border-color:#1565c0;}}
.sec-orange{{color:#e65100;border-color:#e65100;}}
.nav-tabs .nav-link{{color:#555;font-weight:600;font-size:.9rem;}}
.nav-tabs .nav-link.active{{font-weight:700;}}
.tab-rev  .nav-link.active{{color:#1b5e20;border-bottom:3px solid #1b5e20;}}
.tab-cost .nav-link.active{{color:#1565c0;border-bottom:3px solid #1565c0;}}
.accordion-button:not(.collapsed){{background:#f1f8e9;}}
.divider{{height:3px;background:linear-gradient(90deg,#1b5e20,#1565c0,#e65100);
          border-radius:2px;margin:24px 0;}}
</style>
</head>
<body>
<div class="header">
  <h1>鴻騰電子 — 業績 ＋ 成本 整合儀表板</h1>
  <div class="sub">
    營業額期間：{REV_DATE_FROM} ～ {REV_DATE_TO}（USD@{USD_TO_NTD}）｜
    成本期間：{COST_DATE_FROM} ～ {COST_DATE_TO}｜
    更新：{today_str}
  </div>
</div>

<div class="container-fluid px-4">

  <!-- KPI 總覽 -->
  <div class="row g-3 mb-4">
    <div class="col-6 col-md col-xl">
      <div class="kpi" style="border-color:#1b5e20;">
        <div class="val text-success">{fmt(rev_total)}</div>
        <div class="lbl">營業額台幣合計</div>
      </div>
    </div>
    <div class="col-6 col-md col-xl">
      <div class="kpi" style="border-color:#43a047;">
        <div class="val" style="color:#388e3c;">{rev_n_cust}</div>
        <div class="lbl">客戶數</div>
      </div>
    </div>
    <div class="col-6 col-md col-xl">
      <div class="kpi" style="border-color:#f9a825;">
        <div class="val" style="color:#f57f17;">{week_count}</div>
        <div class="lbl">本週新單筆數</div>
      </div>
    </div>
    <div class="col-6 col-md col-xl">
      <div class="kpi" style="border-color:#1565c0;">
        <div class="val" style="color:#1565c0;">{fmt(po_total)}</div>
        <div class="lbl">採購總金額</div>
      </div>
    </div>
    <div class="col-6 col-md col-xl">
      <div class="kpi" style="border-color:#e65100;">
        <div class="val" style="color:#e65100;">{fmt(mo_total)}</div>
        <div class="lbl">委外製程總金額</div>
      </div>
    </div>
  </div>

  <!-- 主 Tab -->
  <ul class="nav nav-tabs mb-0" id="mainTab" style="border-bottom:2px solid #dee2e6;">
    <li class="nav-item"><a class="nav-link active" data-bs-toggle="tab" href="#t1"
      style="color:#1b5e20;">📦 月別×客戶</a></li>
    <li class="nav-item"><a class="nav-link" data-bs-toggle="tab" href="#t2"
      style="">👥 客戶彙總</a></li>
    <li class="nav-item"><a class="nav-link" data-bs-toggle="tab" href="#t3">
      🆕 本週新單 <span class="badge rounded-pill bg-danger ms-1">{week_count}</span></a></li>
    <li class="nav-item"><a class="nav-link" data-bs-toggle="tab" href="#t4"
      style="">🏭 採購成本</a></li>
    <li class="nav-item"><a class="nav-link" data-bs-toggle="tab" href="#t5"
      style="">🔧 委外製程</a></li>
  </ul>

  <div class="tab-content pt-3">

    <!-- ══ Tab 1：月別×客戶 ══ -->
    <div class="tab-pane fade show active" id="t1">
      <div class="row g-3 mb-3">
        <div class="col-12">
          <div class="card-box">
            <div class="sec sec-green">月別×客戶 台幣金額（堆疊長條圖）</div>
            <canvas id="revBar" style="max-height:420px;"></canvas>
          </div>
        </div>
      </div>
      <div class="card-box">
        <div class="sec sec-green">月別×客戶 Pivot（台幣）</div>
        {rev_pivot_html}
      </div>
    </div>

    <!-- ══ Tab 2：客戶彙總 ══ -->
    <div class="tab-pane fade" id="t2">
      <div class="row g-3 mb-3">
        <div class="col-12 col-lg-5">
          <div class="card-box">
            <div class="sec sec-green">客戶金額占比（前15）</div>
            <canvas id="revPie" style="max-height:360px;"></canvas>
          </div>
        </div>
        <div class="col-12 col-lg-7">
          <div class="card-box">
            <div class="sec sec-green">月別營業額趨勢</div>
            <canvas id="revLine" style="max-height:360px;"></canvas>
          </div>
        </div>
      </div>
      <div class="card-box">
        <div class="sec sec-green">客戶月份展開</div>
        <div class="accordion" id="custAcc">{cust_acc}</div>
      </div>
    </div>

    <!-- ══ Tab 3：本週新單 ══ -->
    <div class="tab-pane fade" id="t3">
      <div class="alert alert-success mb-3" style="border-radius:10px;">
        本週（{ws} ～ {we}）共 <strong>{week_count}</strong> 筆新單，
        台幣合計 <strong>NT$ {week_total:,.0f}</strong>
      </div>
      {'<div class="text-center text-muted py-4">本週無新下單資料</div>' if not week_by_cust else
      f'<div class="accordion" id="weekAcc">{week_acc}</div>'}
    </div>

    <!-- ══ Tab 4：採購成本 PURR16 ══ -->
    <div class="tab-pane fade" id="t4">
      <!-- 本週採購 KPI -->
      <div class="alert alert-primary d-flex align-items-center mb-3" style="border-radius:10px;">
        本週（{ws} ～ {we}）採購下單 <strong class="mx-2">{po_week_count}</strong> 筆，
        合計 <strong>NT$ {po_week_total:,.0f}</strong>
      </div>
      <!-- 本週廠商展開 -->
      {po_week_acc_html}
      <!-- 廠商本週小計表 -->
      {po_week_table_html}
      <hr class="my-3">
      <div class="row g-3 mb-3">
        <div class="col-12 col-xl-8">
          <div class="card-box">
            <div class="sec sec-blue">月別採購金額（依廠商堆疊）</div>
            <canvas id="poBar" style="max-height:380px;"></canvas>
          </div>
        </div>
        <div class="col-12 col-xl-4">
          <div class="card-box">
            <div class="sec sec-blue">廠商占比（前15）</div>
            <canvas id="poPie" style="max-height:380px;"></canvas>
          </div>
        </div>
      </div>
      <div class="row g-3 mb-3">
        <div class="col-12 col-lg-5">
          <div class="card-box">
            <div class="sec sec-blue">廠商採購金額排行</div>
            <div style="overflow-y:auto;max-height:400px;">
              <table class="table table-sm table-hover" style="font-size:.82rem;">
                <thead class="table-dark sticky-top"><tr>
                  <th>廠商代號</th><th>廠商名稱</th><th class="text-end">採購金額</th><th class="text-end">筆數</th>
                </tr></thead>
                <tbody>{po_vendor_rows}</tbody>
              </table>
            </div>
          </div>
        </div>
        <div class="col-12 col-lg-7">
          <div class="card-box">
            <div class="sec sec-blue">進料料號金額排行（前50）</div>
            <div style="overflow-y:auto;max-height:400px;">
              <table class="table table-sm table-hover" style="font-size:.82rem;">
                <thead class="table-dark sticky-top"><tr>
                  <th>料號</th><th>品名</th><th class="text-end">數量</th><th class="text-end">採購金額</th>
                </tr></thead>
                <tbody>{po_part_rows}</tbody>
              </table>
            </div>
          </div>
        </div>
      </div>
      <div class="card-box mb-3">
        <div class="sec sec-blue">月別 × 廠商 Pivot（採購金額，台幣）</div>
        {po_piv_html}
      </div>
      <div class="card-box">
        <div class="sec sec-blue">廠商月份展開（依金額大到小）</div>
        <div class="accordion" id="poVendAcc">{po_vend_acc}</div>
      </div>
    </div>

    <!-- ══ Tab 5：委外製程 MOCR34 ══ -->
    <div class="tab-pane fade" id="t5">
      <div class="row g-3 mb-3">
        <div class="col-12 col-xl-8">
          <div class="card-box">
            <div class="sec sec-orange">月別委外金額（依廠商堆疊，依途程步數比例分配）</div>
            <canvas id="moVendBar" style="max-height:380px;"></canvas>
          </div>
        </div>
        <div class="col-12 col-xl-4">
          <div class="card-box">
            <div class="sec sec-orange">月別委外總金額趨勢</div>
            <canvas id="moBar" style="max-height:380px;"></canvas>
          </div>
        </div>
      </div>
      <div class="row g-3 mb-3">
        <div class="col-12 col-lg-5">
          <div class="card-box">
            <div class="sec sec-orange">廠商委外金額排行（分配後）</div>
            <div style="overflow-y:auto;max-height:400px;">
              <table class="table table-sm table-hover" style="font-size:.82rem;">
                <thead class="table-dark sticky-top"><tr>
                  <th>廠商代號</th><th>廠商名稱</th><th class="text-end">關聯製令數</th><th class="text-end">分配委外金額</th>
                </tr></thead>
                <tbody>{mo_vend_rows}</tbody>
              </table>
            </div>
          </div>
        </div>
        <div class="col-12 col-lg-7">
          <div class="card-box">
            <div class="sec sec-orange">委外產品金額排行（前50）</div>
            <div style="overflow-y:auto;max-height:400px;">
              <table class="table table-sm table-hover" style="font-size:.82rem;">
                <thead class="table-dark sticky-top"><tr>
                  <th>品號</th><th>品名</th><th class="text-end">製令數</th><th class="text-end">委外淨金額</th>
                </tr></thead>
                <tbody>{mo_product_rows}</tbody>
              </table>
            </div>
          </div>
        </div>
      </div>
      <div class="card-box mb-3">
        <div class="sec sec-orange">月別 × 廠商 Pivot（分配委外金額）</div>
        {mo_vend_piv_html}
      </div>
      <div class="card-box">
        <div class="sec sec-orange">廠商月份展開</div>
        <div class="accordion" id="moVendAcc">{mo_vend_acc}</div>
      </div>
    </div>

  </div><!-- tab-content -->
</div>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>
<script>
function fmtY(v){{if(v>=1e6)return'$'+(v/1e6).toFixed(1)+'M';if(v>=1e4)return'$'+(v/1e4).toFixed(0)+'萬';return'$'+v.toLocaleString();}}
const font='Microsoft JhengHei';

new Chart(document.getElementById('revBar'),{{type:'bar',
  data:{{labels:{js_rev_months},datasets:{js_rev_ds}}},
  options:{{responsive:true,plugins:{{legend:{{position:'right',labels:{{font:{{family:font}},boxWidth:12,padding:6}}}},
    tooltip:{{callbacks:{{label:c=>`${{c.dataset.label}}: NT$ ${{c.parsed.y.toLocaleString()}}`}}}}}},
    scales:{{x:{{stacked:true,ticks:{{font:{{family:font}}}}}},y:{{stacked:true,ticks:{{callback:fmtY,font:{{family:font}}}},grid:{{color:'#f0f0f0'}}}}}}}}}});

new Chart(document.getElementById('revPie'),{{type:'doughnut',
  data:{{labels:{js_pie_l},datasets:[{{data:{js_pie_d},backgroundColor:{js_pie_c},borderWidth:2}}]}},
  options:{{responsive:true,plugins:{{legend:{{position:'bottom',labels:{{font:{{family:font}},boxWidth:12,padding:6}}}},
    tooltip:{{callbacks:{{label:c=>`${{c.label}}: NT$ ${{c.parsed.toLocaleString()}}`}}}}}}}}}});

new Chart(document.getElementById('revLine'),{{type:'bar',
  data:{{labels:{js_rev_months},datasets:[{{label:'月別台幣合計',data:{js_rev_mt},
    backgroundColor:'rgba(27,94,32,0.75)',borderColor:'#1b5e20',borderWidth:2,borderRadius:4}}]}},
  options:{{responsive:true,plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{label:c=>`NT$ ${{c.parsed.y.toLocaleString()}}`}}}}}},
    scales:{{x:{{ticks:{{font:{{family:font}}}}}},y:{{ticks:{{callback:fmtY,font:{{family:font}}}},grid:{{color:'#f0f0f0'}}}}}}}}}});

new Chart(document.getElementById('poBar'),{{type:'bar',
  data:{{labels:{js_po_m},datasets:{js_po_ds}}},
  options:{{responsive:true,plugins:{{legend:{{position:'right',labels:{{font:{{family:font}},boxWidth:12,padding:6}}}},
    tooltip:{{callbacks:{{label:c=>`${{c.dataset.label}}: NT$ ${{c.parsed.y.toLocaleString()}}`}}}}}},
    scales:{{x:{{stacked:true,ticks:{{font:{{family:font}}}}}},y:{{stacked:true,ticks:{{callback:fmtY,font:{{family:font}}}},grid:{{color:'#f0f0f0'}}}}}}}}}});

new Chart(document.getElementById('poPie'),{{type:'doughnut',
  data:{{labels:{js_po_pie_l},datasets:[{{data:{js_po_pie_d},backgroundColor:{js_po_pie_c},borderWidth:2}}]}},
  options:{{responsive:true,plugins:{{legend:{{position:'bottom',labels:{{font:{{family:font}},boxWidth:12,padding:6}}}},
    tooltip:{{callbacks:{{label:c=>`${{c.label}}: NT$ ${{c.parsed.toLocaleString()}}`}}}}}}}}}});

new Chart(document.getElementById('moVendBar'),{{type:'bar',
  data:{{labels:{js_mo_vend_m},datasets:{js_mo_vend_ds}}},
  options:{{responsive:true,plugins:{{legend:{{position:'right',labels:{{font:{{family:font}},boxWidth:12,padding:6}}}},
    tooltip:{{callbacks:{{label:c=>`${{c.dataset.label}}: NT$ ${{c.parsed.y.toLocaleString()}}`}}}}}},
    scales:{{x:{{stacked:true,ticks:{{font:{{family:font}}}}}},y:{{stacked:true,ticks:{{callback:fmtY,font:{{family:font}}}},grid:{{color:'#f0f0f0'}}}}}}}}}});

new Chart(document.getElementById('moBar'),{{type:'bar',
  data:{{labels:{js_mo_m},datasets:[{{label:'委外淨金額',data:{js_mo_d},
    backgroundColor:'rgba(230,81,0,0.75)',borderColor:'#e65100',borderWidth:2,borderRadius:4}}]}},
  options:{{responsive:true,plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{label:c=>`NT$ ${{c.parsed.y.toLocaleString()}}`}}}}}},
    scales:{{x:{{ticks:{{font:{{family:font}}}}}},y:{{ticks:{{callback:fmtY,font:{{family:font}}}},grid:{{color:'#f0f0f0'}}}}}}}}}});
</script>
</body></html>"""

    with open(path,'w',encoding='utf-8') as f: f.write(html)
    print(f"[OK] HTML：{path}")
    return path


if __name__=="__main__":
    print("撈取所有資料中...")
    df_rev       = fetch_revenue()
    df_po        = fetch_purchase()
    df_mo, df_rt = fetch_outsource()
    generate(df_rev, df_po, df_mo, df_rt)
