#!/usr/bin/env python3
"""
build_wifi_ytd_dashboard.py
Generates wifi-ytd-dashboard.html
from GA4 YTD CSV exports (Jan 1 2025 – Apr 30 2026).
"""

import csv, json, os, re
from collections import defaultdict
import openpyxl

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ── Input files ───────────────────────────────────────────────────────────
# Channel Acquisition — quarterly pulls + monthly delta files for active quarter.
# Re-pulled May 2026 with corrected GA4 Key Event config (reservation event only).
ACQ_CSVS = [os.path.join(SCRIPT_DIR, f) for f in [
    '2025-Q1-Channel-Acquisition.csv',
    '2025-Q2-Channel-Acquisition.csv',
    '2025-Q3-Channel-Acquisition.csv',
    '2025-Q4-Channel-Acquisition.csv',
    '2026-Q1-Channel-Acquisition.csv',
    '2026-04-Channel-Acquisition.csv',
]]
LP_CSVS  = [os.path.join(SCRIPT_DIR, f) for f in [
    '2025-Q1-Landing-Pages.csv',
    '2025-Q2-Landing-Pages.csv',
    '2025-Q3-Landing-Pages.csv',
    '2025-Q4-Landing-Pages.csv',
    '2026-Q1-Landing-Pages.csv',
    '2026-04-Landing-Pages.csv',
]]
DB_XLSX  = os.path.join(SCRIPT_DIR, 'My Place Hotels Database - Master.xlsx')
LOGO_TXT = os.path.join(SCRIPT_DIR, 'logo_uri.txt')

# ── Output file ───────────────────────────────────────────────────────────
OUT_FILE = os.path.join(SCRIPT_DIR, 'wifi-ytd-dashboard.html')

# ── Constants ─────────────────────────────────────────────────────────────
WIFI_CH   = "Property WiFi"   # matched via startswith to handle varying trailing spaces
CUTOFF_25 = "20250430"           # YTD same-period end date for 2025
CUTOFF_26 = "20260430"           # YTD end date for 2026

REMOVED_HOTELS = {
    '/wifi/north-aurora-il',
    '/wifi/spokane-valley',
    '/wifi/spokane-valley-wa',
    '/wifi/avondale-az',
}

WIFI_PAGE_MERGES = {
    '/wifi/aberdeen-sd-2': '/wifi/aberdeen-sd',
    '/wifi/wi-fi-page-buckeye-az': '/wifi/buckeye-az',
    '/wifi/idaho-falls': '/wifi/idaho-falls-id',
    '/wifi/tucson-south': '/wifi/tucson-south-az',
    '/wifi/wi-fi-page-template-2': None,
}

WIFI_NAME_OVERRIDES = {
    '/wifi/monaca-pa':   'My Place Hotel-Pittsburgh North/Monaca, PA',
    '/wifi/missoula-mt': 'My Place Hotel-Missoula, MT',
    '/wifi/augusta-ga':  'My Place Hotel-Augusta, GA',
}

# ── Helpers ───────────────────────────────────────────────────────────────
def load_logo():
    try:
        with open(LOGO_TXT) as f:
            return f.read().strip()
    except FileNotFoundError:
        return ''

def is_artifact(path):
    return '\u2026' in path or '\u201d' in path or '\u201c' in path

def fmt_n(n):
    if n >= 1e6:   return f'{n/1e6:.2f}M'
    elif n >= 1e3: return f'{n/1e3:.1f}K'
    else:          return f'{round(n):,}'

def yoy_badge(curr, prev):
    if prev == 0:
        return '<span style="font-size:11px;color:#5A6B84;">New</span>'
    pct = (curr - prev) / prev * 100
    if pct >= 0:
        color, bg, arrow = '#2E8B57', 'rgba(46,139,87,0.1)', '▲'
    else:
        color, bg, arrow = '#CC3333', 'rgba(204,51,51,0.1)', '▼'
    return f'<span style="font-size:11px;font-weight:700;color:{color};background:{bg};padding:2px 8px;border-radius:10px;">{arrow} {abs(pct):.1f}%</span>'

def read_csv(filepath):
    with open(filepath, encoding='utf-8-sig') as f:
        lines = [l for l in f if not l.startswith('#')]
    return list(csv.DictReader(lines))

def _infer_date_from_filename(fp):
    """If a monthly delta CSV has no Date column, synthesize one from its filename.
       Pattern: 2026-04-... → returns '20260430' (last day of month) so it
       passes both same-period cutoff filters and month-bucket logic."""
    m = re.search(r'(\d{4})-(\d{2})', os.path.basename(fp))
    if not m:
        return None
    yyyy, mm = m.group(1), m.group(2)
    last_days = {'01':'31','02':'28','03':'31','04':'30','05':'31','06':'30',
                 '07':'31','08':'31','09':'30','10':'31','11':'30','12':'31'}
    return f'{yyyy}{mm}{last_days.get(mm, "28")}'

def read_multi_csv(filepaths):
    """Read and merge multiple CSV files with the same column structure.
       If a file lacks a Date column (monthly aggregate export), synthesize
       a Date from the filename so downstream parsers can bucket it correctly."""
    all_rows = []
    for fp in filepaths:
        chunk = read_csv(fp)
        if chunk and 'Date' not in chunk[0]:
            inferred = _infer_date_from_filename(fp)
            if inferred:
                for r in chunk:
                    r['Date'] = inferred
                print(f'    {os.path.basename(fp)}: {len(chunk):,} rows  (Date synthesized = {inferred})')
            else:
                print(f'    {os.path.basename(fp)}: {len(chunk):,} rows  (no Date column, could not infer)')
        else:
            print(f'    {os.path.basename(fp)}: {len(chunk):,} rows')
        all_rows.extend(chunk)
    return all_rows

# ── Load hotel names from Excel ───────────────────────────────────────────
def load_hotel_names():
    hotel_names = {}
    hotel_slug_index = {}
    try:
        wb = openpyxl.load_workbook(DB_XLSX, read_only=True, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or len(row) < 13:
                continue
            name = row[1]     # Column B — Hotel Name
            url  = row[12]    # Column M — Property URL (shifted from L→M in May 2026 layout after Pilot col added at I)
            if not name or not url:
                continue
            url = str(url).strip()
            m = re.search(r'(/locations/[^\s?#]+)', url)
            if not m:
                continue
            path = m.group(1).rstrip('/')
            hotel_names[path] = str(name).strip()
            slug = path.split('/')[-1]
            hotel_slug_index[slug] = path
        wb.close()
    except Exception as e:
        print(f'Warning: could not load hotel names: {e}')
    return hotel_names, hotel_slug_index

def resolve_wifi_name(wifi_path, hotel_names, hotel_slug_index):
    if wifi_path in WIFI_NAME_OVERRIDES:
        return WIFI_NAME_OVERRIDES[wifi_path]
    slug = wifi_path.replace('/wifi/', '')
    hotel_path = f'/locations/my-place-hotel-{slug}'
    if hotel_path in hotel_names:
        return hotel_names[hotel_path]
    if slug in hotel_slug_index:
        return hotel_names[hotel_slug_index[slug]]
    parts = slug.split('-')
    if len(parts) >= 2:
        state = parts[-1].upper()
        city  = ' '.join(p.capitalize() for p in parts[:-1])
        return f'My Place Hotel-{city}, {state}'
    return wifi_path

# ── Parse acquisition CSV ─────────────────────────────────────────────────
def parse_acq(rows):
    # Full-year 2025 monthly WiFi sessions (for trend chart)
    mo_wifi_25_full = {m: 0.0 for m in range(1, 13)}
    # 2026 YTD monthly WiFi sessions
    mo_wifi_26 = defaultdict(float)
    # Same-period totals
    wifi_sp_25 = 0.0
    wifi_sp_26 = 0.0

    for row in rows:
        channel = row.get('Session Updated Channel Group', '').strip()
        date    = str(row.get('Date', '')).strip()
        if len(date) != 8:
            continue
        if not channel.startswith(WIFI_CH):
            continue
        year  = int(date[:4])
        month = int(date[4:6])
        if year not in (2025, 2026):
            continue
        try:
            sess = float(row.get('Sessions') or 0)
        except (ValueError, TypeError):
            continue

        if year == 2025 and month in range(1, 13):
            mo_wifi_25_full[month] += sess
            if date <= CUTOFF_25:
                wifi_sp_25 += sess
        elif year == 2026 and date <= CUTOFF_26:
            mo_wifi_26[month] += sess
            wifi_sp_26 += sess

    return mo_wifi_25_full, mo_wifi_26, wifi_sp_25, wifi_sp_26

# ── Parse landing page CSV ────────────────────────────────────────────────
def parse_lp(rows, hotel_names, hotel_slug_index):
    # Same-period WiFi page views/users
    wifi_sp = defaultdict(lambda: {'v25': 0, 'v26': 0, 'u25': 0, 'u26': 0})

    for row in rows:
        raw_path = str(row.get('Landing page + query string', '')).strip()
        path  = raw_path.split('?')[0]
        date  = str(row.get('Date', '')).strip()
        if len(date) != 8:
            continue
        year  = int(date[:4])
        month = int(date[4:6])
        if year not in (2025, 2026):
            continue
        if not path.startswith('/wifi/'):
            continue
        if is_artifact(path) or path == '(not set)':
            continue

        p = WIFI_PAGE_MERGES.get(path, path)
        if p is None or p in REMOVED_HOTELS:
            continue

        try:
            views = int(float(row.get('Views') or 0))
            users = int(float(row.get('Active users') or 0))
        except (ValueError, TypeError):
            continue

        if year == 2025 and date <= CUTOFF_25:
            wifi_sp[p]['v25'] += views
            wifi_sp[p]['u25'] += users
        elif year == 2026 and date <= CUTOFF_26:
            wifi_sp[p]['v26'] += views
            wifi_sp[p]['u26'] += users

    # Build sorted list
    wifi_list = []
    for wifi_path, d in wifi_sp.items():
        label = resolve_wifi_name(wifi_path, hotel_names, hotel_slug_index)
        wifi_list.append({
            'label': label, 'path': wifi_path,
            'v25': round(d['v25']), 'v26': round(d['v26']),
            'u25': round(d['u25']), 'u26': round(d['u26']),
        })
    wifi_list.sort(key=lambda x: x['v25'] + x['v26'], reverse=True)

    return wifi_list

# ── CSS ───────────────────────────────────────────────────────────────────
CSS = """
:root {
  --bg:#EEF1F6; --card:#FFFFFF; --border:#D0D8E8;
  --text:#1C355E; --muted:#5A6B84; --navy:#1C355E;
  --orange:#FF5F00; --green:#2E8B57;
  --C25:#4A7FC1; --C26:#FF5F00;
}
* { margin:0; padding:0; box-sizing:border-box; }
body { font-family:Arial,'Helvetica Neue',Helvetica,sans-serif;
       background:var(--bg); color:var(--text); min-height:100vh; font-size:13px; }
.brand-bar { background:#FF5F00; height:5px; }
.header { background:#1C355E; padding:22px 40px 18px; border-bottom:3px solid #FF5F00; }
.header-inner { max-width:1400px; margin:0 auto;
                display:flex; justify-content:space-between; align-items:flex-end; }
.header-logo { display:flex; align-items:center; gap:16px; }
.header-eyebrow { font-size:10px; font-weight:600; color:#FFD080;
                  text-transform:uppercase; letter-spacing:2.5px; margin-bottom:4px; }
.header h1 { font-size:20px; font-weight:700; color:#fff; line-height:1.1; }
.header-subtitle { font-size:11px; color:rgba(255,255,255,0.65); margin-top:4px; }
.header-meta { text-align:right; display:flex; flex-direction:column; align-items:flex-end; gap:6px; }
.portal-btn { display:inline-flex; align-items:center; gap:5px; background:rgba(255,255,255,0.12);
              border:1px solid rgba(255,255,255,0.3); color:rgba(255,255,255,0.85);
              font-size:11px; font-weight:600; padding:4px 10px; border-radius:5px;
              text-decoration:none; transition:background 0.2s; white-space:nowrap; }
.portal-btn:hover { background:rgba(255,95,0,0.4); }
.updated-pill { background:rgba(255,255,255,0.12); border:1px solid rgba(255,255,255,0.25);
                color:rgba(255,255,255,0.7); font-size:10px; padding:3px 10px; border-radius:20px; }
.main { max-width:1400px; margin:0 auto; padding:24px 40px 48px; }
.info-bar { border-radius:8px; padding:10px 16px; margin-bottom:22px; font-size:11px;
            color:#5A6B84; display:flex; align-items:center; gap:10px;
            background:rgba(255,95,0,0.05); border:1px solid rgba(255,95,0,0.2); }
.info-bar strong { color:#FF5F00; }
.kpi-row { display:grid; grid-template-columns:repeat(4,1fr); gap:14px; margin-bottom:24px; }
.kpi-card { background:#fff; border:1px solid var(--border); border-radius:10px;
            padding:14px 16px; box-shadow:0 1px 4px rgba(28,53,94,0.06);
            border-top:3px solid #FF5F00; }
.kpi-label { font-size:9px; text-transform:uppercase; letter-spacing:0.8px;
             color:#8ba0bf; font-weight:700; margin-bottom:5px; }
.kpi-value { font-size:20px; font-weight:700; color:#1C355E; margin-bottom:3px; line-height:1; }
.kpi-sub   { font-size:11px; color:#5A6B84; margin-bottom:5px; }
.kpi-yoy  { margin-top:3px; }
.section-hdr { display:flex; align-items:center; gap:10px; margin:24px 0 14px; }
.section-hdr h2 { font-size:13px; font-weight:700; color:#1C355E; white-space:nowrap; }
.section-hdr .desc { font-size:11px; color:#8ba0bf; font-style:italic; }
.section-divider { flex:1; height:1px; background:linear-gradient(to right,#D0D8E8,transparent); }
.chart-card { background:#fff; border:1px solid var(--border); border-radius:10px;
              padding:18px 20px; box-shadow:0 1px 4px rgba(28,53,94,0.06); margin-bottom:16px; }
.chart-card h3 { font-size:11px; font-weight:700; color:var(--muted); text-transform:uppercase;
                 letter-spacing:0.6px; border-bottom:2px solid #FF5F00;
                 padding-bottom:7px; margin-bottom:5px; }
.chart-desc { font-size:10px; color:#8ba0bf; font-style:italic; margin-bottom:10px; line-height:1.5; }
.legend-row { display:flex; gap:14px; margin-bottom:8px; }
.legend-item { display:flex; align-items:center; gap:5px; font-size:10px; color:#5A6B84; }
.legend-swatch { width:10px; height:10px; border-radius:2px; }
.table-card { background:#fff; border:1px solid var(--border); border-radius:10px;
              padding:18px 20px; box-shadow:0 1px 4px rgba(28,53,94,0.06); margin-bottom:16px; }
.table-card h3 { font-size:11px; font-weight:700; color:var(--muted); text-transform:uppercase;
                 letter-spacing:0.6px; border-bottom:2px solid #FF5F00;
                 padding-bottom:7px; margin-bottom:5px; }
table { width:100%; border-collapse:collapse; font-size:11px; }
thead th { background:#1C355E; color:#fff; padding:7px 8px; text-align:left;
           font-size:9px; text-transform:uppercase; letter-spacing:0.4px;
           font-weight:700; white-space:nowrap; cursor:pointer; user-select:none; }
thead th:hover { color:#FFD080; }
thead th.sorted-asc::after  { content:' ▲'; font-size:7px; }
thead th.sorted-desc::after { content:' ▼'; font-size:7px; }
thead th.num { text-align:right; }
tbody td { padding:6px 8px; border-bottom:1px solid var(--border); color:#1C355E; }
tbody td.num { text-align:right; font-variant-numeric:tabular-nums; white-space:nowrap; }
tbody td.name-cell { font-weight:600; }
tbody tr:nth-child(even) { background:#F5F8FC; }
tbody tr:hover { background:#E8F0FA; }
tbody tr:last-child td { border-bottom:none; }
.delta-cell { font-size:10px; font-weight:700; white-space:nowrap; }
.delta-cell.up { color:#2E8B57; }
.delta-cell.down { color:#CC3333; }
.wifi-table { table-layout:fixed; width:100%; }
.wifi-table thead th:nth-child(1) { width:30%; }
.wifi-table thead th { width:11.7%; }
.footer { background:#1C355E; border-top:3px solid #FF5F00; padding:14px 40px;
          text-align:center; color:rgba(255,255,255,0.55); font-size:10px; }
"""

JS = """
const C25='#4A7FC1', C26='#FF5F00';

function fmtN(n){ return n>=1e6?(n/1e6).toFixed(2)+'M':n>=1e3?(n/1e3).toFixed(1)+'K':Math.round(n).toLocaleString(); }

// Trend chart — 2025 full year (12 pts) + 2026 YTD overlay (starts at x=0)
function drawTrendLine(canvasId, arr25, arr26, fmtFn) {
  const canvas = document.getElementById(canvasId); if (!canvas) return;
  const dpr = window.devicePixelRatio || 1;
  const W = canvas.parentElement.clientWidth || 900;
  const H = parseInt(canvas.getAttribute('height')) || 160;
  canvas.width = W*dpr; canvas.height = H*dpr;
  canvas.style.width = W+'px'; canvas.style.height = H+'px';
  const ctx = canvas.getContext('2d'); ctx.scale(dpr, dpr);
  const PAD = {t:20,r:14,b:24,l:52};
  const cW = W-PAD.l-PAD.r, cH = H-PAD.t-PAD.b;
  const n = 12;
  const maxVal = Math.max(...arr25, ...arr26) * 1.18;
  const xp = i => PAD.l + (i/(n-1)) * cW;
  const yp = v => PAD.t + cH * (1 - v/maxVal);
  const moL = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  for (let i=0; i<=4; i++) {
    const v = maxVal*i/4, y = yp(v);
    ctx.strokeStyle='rgba(208,216,232,0.7)'; ctx.lineWidth=0.5;
    ctx.beginPath(); ctx.moveTo(PAD.l,y); ctx.lineTo(PAD.l+cW,y); ctx.stroke();
    ctx.fillStyle='#8ba0bf'; ctx.font='8px Arial';
    ctx.textAlign='right'; ctx.fillText(fmtFn(v), PAD.l-4, y+3);
  }
  function drawLine(arr, color, fillAlpha, nPts) {
    ctx.beginPath();
    for (let i=0;i<nPts;i++) { i===0?ctx.moveTo(xp(i),yp(arr[i])):ctx.lineTo(xp(i),yp(arr[i])); }
    ctx.lineTo(xp(nPts-1),PAD.t+cH); ctx.lineTo(xp(0),PAD.t+cH); ctx.closePath();
    ctx.globalAlpha=fillAlpha; ctx.fillStyle=color; ctx.fill(); ctx.globalAlpha=1;
    ctx.beginPath();
    for (let i=0;i<nPts;i++) { i===0?ctx.moveTo(xp(i),yp(arr[i])):ctx.lineTo(xp(i),yp(arr[i])); }
    ctx.strokeStyle=color; ctx.lineWidth=2; ctx.lineJoin='round'; ctx.stroke();
    for (let i=0;i<nPts;i++) {
      ctx.beginPath(); ctx.arc(xp(i),yp(arr[i]),3,0,Math.PI*2);
      ctx.fillStyle=color; ctx.fill();
    }
  }
  drawLine(arr25, C25, 0.07, 12);
  drawLine(arr26, C26, 0.12, arr26.length);
  ctx.fillStyle='#8ba0bf'; ctx.font='8px Arial'; ctx.textAlign='center';
  for (let i=0;i<n;i++) ctx.fillText(moL[i], xp(i), H-4);
  const p25 = arr25.indexOf(Math.max(...arr25));
  ctx.font='bold 8px Arial'; ctx.fillStyle=C25; ctx.textAlign='center';
  ctx.fillText(fmtFn(arr25[p25]), xp(p25), yp(arr25[p25])-6);
  const last26 = arr26.length-1;
  ctx.fillStyle=C26;
  ctx.fillText(fmtFn(arr26[last26]), xp(last26), yp(arr26[last26])-6);
}

let sortState={};
function sortTbl(id, col, numeric) {
  const tbl=document.getElementById(id);
  const tbody=tbl.querySelector('tbody');
  const rows=Array.from(tbody.querySelectorAll('tr'));
  const key=id+'_'+col, asc=sortState[key]!=='asc';
  sortState[key]=asc?'asc':'desc';
  rows.sort((a,b)=>{
    const rawA=a.cells[col].textContent.trim();
    const rawB=b.cells[col].textContent.trim();
    if (numeric) {
      const pn=s=>{const c=s.replace(/[$+,%\u25b2\u25bc\s]/g,'');const m=c.match(/([-\d.]+)\s*([KMpp]*)/i);if(!m)return 0;let v=parseFloat(m[1])||0;if(m[2]==='K'||m[2]==='k')v*=1e3;if(m[2]==='M'||m[2]==='m')v*=1e6;return v;};
      const va=pn(rawA), vb=pn(rawB);
      return asc?va-vb:vb-va;
    } else {
      return asc?rawA.localeCompare(rawB):rawB.localeCompare(rawA);
    }
  });
  rows.forEach(r=>tbody.appendChild(r));
  tbl.querySelectorAll('thead th').forEach((th,i)=>{
    th.classList.remove('sorted-asc','sorted-desc');
    if(i===col) th.classList.add(asc?'sorted-asc':'sorted-desc');
  });
}

function deltaCell(vA, vB, fmtFn) {
  const d=vB-vA;
  const cls=d>=0?'up':'down';
  return `<td class="num delta-cell ${cls}">${d>=0?'+':''}${fmtFn(Math.abs(d))}</td>`;
}
"""

# ── Build HTML ─────────────────────────────────────────────────────────────
def build_html(wifi_list, mo_wifi_25_full, mo_wifi_26,
               wifi_sp_25, wifi_sp_26, wifi_full_25_total, logo):
    months_26 = sorted(mo_wifi_26.keys())
    SESS25 = [round(mo_wifi_25_full[m]) for m in range(1, 13)]
    SESS26 = [round(mo_wifi_26[m]) for m in months_26]

    n_props = len(wifi_list)
    avg_sp_25 = wifi_sp_25 / n_props if n_props else 0
    avg_sp_26 = wifi_sp_26 / n_props if n_props else 0

    data_js = f"""
const SESS25={json.dumps(SESS25)}, SESS26={json.dumps(SESS26)};
const WIFI_DATA={json.dumps(wifi_list)};
"""

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>My Place Hotels — Property WiFi 2026 YTD vs 2025</title>
<style>
{CSS}
</style>
</head>
<body>
<div class="brand-bar"></div>

<div class="header">
  <div class="header-inner">
    <div class="header-logo">
      <img src="{logo}" alt="My Place Hotels" style="height:42px;display:block;">
      <div>
        <div class="header-eyebrow">My Place Hotels of America</div>
        <h1>Property WiFi Dashboard — 2026 YTD</h1>
        <div class="header-subtitle">In-Hotel Digital Concierge &nbsp;·&nbsp; Jan 1 – Apr 30, 2026 vs same period 2025 &nbsp;·&nbsp; {n_props} Properties</div>
      </div>
    </div>
    <div class="header-meta">
      <div style="display:flex;gap:8px;">
        <a href="wifi-dashboard.html" class="portal-btn">2024 vs 2025 WiFi</a>
        <a href="index.html" class="portal-btn">← Portal</a>
      </div>
      <div class="updated-pill">Jan 1 – Apr 30, 2026</div>
    </div>
  </div>
</div>

<div class="main">

  <div class="info-bar">
    📅 <span><strong>YoY comparison is same-period only</strong> — 2026 YTD (Jan 1–Apr 30, 2026) is compared against the identical date range in 2025 (Jan 1–Apr 30, 2025) for a clean apples-to-apples view. Monthly trend chart shows all 12 months of 2025 for broader context. Property WiFi sessions reflect guest access on the in-hotel digital concierge pages (/wifi/ path).</span>
  </div>

  <!-- KPI Cards -->
  <div class="kpi-row">
    <div class="kpi-card">
      <div class="kpi-label">2026 YTD WiFi Sessions</div>
      <div class="kpi-value">{fmt_n(wifi_sp_26)}</div>
      <div class="kpi-sub">2025 (same period): {fmt_n(wifi_sp_25)}</div>
      <div class="kpi-yoy">{yoy_badge(wifi_sp_26, wifi_sp_25)} vs 2025</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">Full Year 2025 WiFi Sessions</div>
      <div class="kpi-value">{fmt_n(wifi_full_25_total)}</div>
      <div class="kpi-sub">Jan – Dec 2025 total</div>
      <div class="kpi-yoy"><span style="font-size:11px;color:#5A6B84;">Full year baseline</span></div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">Active Properties (YTD)</div>
      <div class="kpi-value">{n_props}</div>
      <div class="kpi-sub">WiFi pages with sessions in either period</div>
      <div class="kpi-yoy"><span style="font-size:11px;color:#5A6B84;">Digital concierge pages</span></div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">Avg Sessions / Property</div>
      <div class="kpi-value">{fmt_n(avg_sp_26)}</div>
      <div class="kpi-sub">2025 same period: {fmt_n(avg_sp_25)}</div>
      <div class="kpi-yoy">{yoy_badge(avg_sp_26, avg_sp_25)} vs 2025</div>
    </div>
  </div>

  <!-- Monthly Trend -->
  <div class="section-hdr">
    <h2>Monthly WiFi Sessions Trend</h2>
    <span class="desc">Full 2025 (blue) + 2026 YTD (orange)</span>
    <div class="section-divider"></div>
  </div>

  <div class="chart-card">
    <h3>Monthly Property WiFi Sessions</h3>
    <p class="chart-desc">All 12 months of 2025 shown for context, with 2026 YTD (Jan–Mar) overlaid. Sessions represent guest interactions on in-hotel digital concierge pages.</p>
    <div class="legend-row">
      <div class="legend-item"><div class="legend-swatch" style="background:var(--C25);"></div>2025 (full year)</div>
      <div class="legend-item"><div class="legend-swatch" style="background:var(--C26);"></div>2026 YTD</div>
    </div>
    <canvas id="cSessions" style="width:100%;display:block;" height="160"></canvas>
  </div>

  <!-- Property Table -->
  <div class="section-hdr">
    <h2>Property WiFi Performance — {n_props} Properties</h2>
    <span class="desc">Same period Jan 1–Apr 30 · sorted by combined views · click any column to re-sort</span>
    <div class="section-divider"></div>
  </div>
  <div class="table-card">
    <h3>Property WiFi Page Views &amp; Active Users — 2025 vs 2026 YTD</h3>
    <p class="chart-desc">Views and unique active users per property WiFi page. Both years reflect Jan 1–Apr 30 only. Properties without activity in either period are excluded.</p>
    <table id="wifiTable" class="wifi-table">
      <thead><tr>
        <th onclick="sortTbl('wifiTable',0,false)">Property</th>
        <th class="num" onclick="sortTbl('wifiTable',1,true)">Views '25</th>
        <th class="num" onclick="sortTbl('wifiTable',2,true)">Views '26</th>
        <th class="num" onclick="sortTbl('wifiTable',3,true)">Δ Views</th>
        <th class="num" onclick="sortTbl('wifiTable',4,true)">Users '25</th>
        <th class="num" onclick="sortTbl('wifiTable',5,true)">Users '26</th>
        <th class="num" onclick="sortTbl('wifiTable',6,true)">Δ Users</th>
      </tr></thead>
      <tbody id="wifiBody"></tbody>
    </table>
  </div>

</div>

<div class="footer">
  My Place Hotels of America &nbsp;·&nbsp; Property WiFi 2026 YTD &nbsp;·&nbsp; Google Analytics 4 &nbsp;·&nbsp; Confidential — Internal Use Only
</div>

<script>
{JS}
{data_js}

function renderWifiTable() {{
  document.getElementById('wifiBody').innerHTML = WIFI_DATA.map(r=>{{
    return `<tr>
      <td class="name-cell">${{r.label}}</td>
      <td class="num">${{r.v25.toLocaleString()}}</td>
      <td class="num">${{r.v26.toLocaleString()}}</td>
      ${{deltaCell(r.v25,r.v26,n=>Math.round(n).toLocaleString())}}
      <td class="num">${{r.u25.toLocaleString()}}</td>
      <td class="num">${{r.u26.toLocaleString()}}</td>
      ${{deltaCell(r.u25,r.u26,n=>Math.round(n).toLocaleString())}}
    </tr>`;
  }}).join('');
}}

window.addEventListener('load', ()=>{{
  drawTrendLine('cSessions', SESS25, SESS26, fmtN);
  renderWifiTable();
}});
window.addEventListener('resize', ()=>{{
  drawTrendLine('cSessions', SESS25, SESS26, fmtN);
}});
</script>
</body>
</html>"""

# ── Main ──────────────────────────────────────────────────────────────────
def main():
    print('Loading hotel names...')
    hotel_names, hotel_slug_index = load_hotel_names()
    print(f'  Loaded {len(hotel_names)} hotel name mappings')

    logo = load_logo()
    print(f'  Logo: {"loaded" if logo else "NOT FOUND"}')

    print('Parsing acquisition CSV...')
    acq_rows = read_multi_csv(ACQ_CSVS)
    mo_wifi_25_full, mo_wifi_26, wifi_sp_25, wifi_sp_26 = parse_acq(acq_rows)
    wifi_full_25_total = sum(mo_wifi_25_full.values())

    print('Parsing landing page CSVs...')
    lp_rows = read_multi_csv(LP_CSVS)
    print(f'  Total LP rows: {len(lp_rows):,}')
    wifi_list = parse_lp(lp_rows, hotel_names, hotel_slug_index)

    print(f'  WiFi properties: {len(wifi_list)}')
    print(f'  2026 YTD WiFi sessions: {fmt_n(wifi_sp_26)}, 2025 same period: {fmt_n(wifi_sp_25)}')

    print(f'Writing {OUT_FILE}...')
    html = build_html(wifi_list, mo_wifi_25_full, mo_wifi_26,
                      wifi_sp_25, wifi_sp_26, wifi_full_25_total, logo)
    with open(OUT_FILE, 'w', encoding='utf-8') as f:
        f.write(html)

    print('Done!')

if __name__ == '__main__':
    main()
