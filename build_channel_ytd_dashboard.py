#!/usr/bin/env python3
"""
Build the Channel Performance YTD dashboard HTML for My Place Hotels.

This script reads the monthly Channel-XXX-YY (Stay Date).xlsx exports
for the current-year YTD period and the same-period prior year, computes
the per-channel/per-month metrics the dashboard needs, and rewrites the
data constants block and period labels in
`channel-performance-dashboard-2026-ytd.html`.

To roll forward at month-end: append the new month abbreviation to
`MONTHS` (e.g. add "May"), drop the new Excel into the workspace folder,
and run:

    python3 build_channel_ytd_dashboard.py

The HTML layout/CSS/chart-drawing code is left untouched — only data
constants and period text are replaced.
"""

from __future__ import annotations

import json
import re
import sys
from datetime import date
from pathlib import Path

try:
    import openpyxl  # type: ignore
except ImportError:
    sys.stderr.write(
        "openpyxl is required: install with `pip3 install openpyxl`\n"
    )
    raise

# ─────────────────────────────────────────────────────────────────────────────
# Configuration — update MONTHS each month-end
# ─────────────────────────────────────────────────────────────────────────────
MONTHS = ["Jan", "Feb", "Mar", "Apr"]  # extend this list each month-end
CY = 2026
PY = 2025

# Resolve BASE: prefer the macOS path; fall back to the script's own folder so
# the script also works when invoked from a Cowork workspace mount.
_MAC_BASE = Path("/Users/jeffthomas/Documents/Claude/Projects/MPH Performance Dashboards")
if _MAC_BASE.exists():
    BASE = _MAC_BASE
else:
    BASE = Path(__file__).resolve().parent
HTML_PATH = BASE / "channel-performance-dashboard-2026-ytd.html"

# Channel display order (matches the existing dashboard chip order)
CHANNELS = [
    "PMS",
    "Expedia",
    "Website",
    "Booking",
    "GDS",
    "Expedia Programs",
    "Other OTA",
    "Priceline/Agoda",
    "HotelBeds",
    "BnBerry",
    "Hotel Tonight",
    "Hopper",
]

MONTH_FULL = {
    "Jan": "January", "Feb": "February", "Mar": "March", "Apr": "April",
    "May": "May", "Jun": "June", "Jul": "July", "Aug": "August",
    "Sep": "September", "Oct": "October", "Nov": "November", "Dec": "December",
}


def file_for(month: str, yy_short: int) -> Path:
    """Return Path to monthly Excel for the given short year (e.g. 26 or 25)."""
    return BASE / f"Channel-{month}-{yy_short} (Stay Date).xlsx"


def read_month(path: Path) -> dict:
    """Read a Channel-XXX-YY xlsx; return {channel: {rn, rev, adr}} for each row.

    Skips the Total row. Net columns are C (Net RN's, index 2),
    G (Net ADR, index 6), K (Net Revenue, index 10).
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Export"]
    out: dict[str, dict[str, float]] = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row is None or row[0] is None:
            continue
        channel = str(row[0]).strip()
        if channel.lower() == "total":
            continue
        net_rn = float(row[2] or 0)
        net_adr = float(row[6] or 0)
        net_rev = float(row[10] or 0)
        out[channel] = {"rn": net_rn, "adr": net_adr, "rev": net_rev}
    return out


def yy(year: int) -> int:
    return year % 100


def build_period_data(year: int) -> dict:
    """Load all MONTHS for a year and return per-channel monthly arrays plus totals."""
    yy_short = yy(year)
    monthly_data: dict[str, dict] = {}  # month -> channel -> metrics
    for m in MONTHS:
        path = file_for(m, yy_short)
        if not path.exists():
            raise FileNotFoundError(f"Missing source file: {path}")
        monthly_data[m] = read_month(path)

    # Per-channel monthly arrays (aligned to MONTHS order)
    rev = {ch: [] for ch in CHANNELS}
    rn = {ch: [] for ch in CHANNELS}
    adr = {ch: [] for ch in CHANNELS}
    for m in MONTHS:
        for ch in CHANNELS:
            mdata = monthly_data[m].get(ch, {"rn": 0, "adr": 0, "rev": 0})
            rev[ch].append(mdata["rev"])
            rn[ch].append(mdata["rn"])
            adr[ch].append(mdata["adr"])

    # Monthly totals across all channels (excluding the Total row)
    mo_rev = [sum(rev[ch][mi] for ch in CHANNELS) for mi, _ in enumerate(MONTHS)]
    mo_rn = [sum(rn[ch][mi] for ch in CHANNELS) for mi, _ in enumerate(MONTHS)]
    mo_adr = [
        (mo_rev[mi] / mo_rn[mi]) if mo_rn[mi] else 0.0
        for mi in range(len(MONTHS))
    ]

    # YTD per-channel totals
    ann_rev = {ch: sum(rev[ch]) for ch in CHANNELS}
    ann_rn = {ch: sum(rn[ch]) for ch in CHANNELS}
    ann_adr = {
        ch: (ann_rev[ch] / ann_rn[ch]) if ann_rn[ch] else 0.0
        for ch in CHANNELS
    }

    grand_rev = sum(ann_rev.values())
    grand_rn = sum(ann_rn.values())

    return {
        "rev": rev, "rn": rn, "adr": adr,
        "mo_rev": mo_rev, "mo_rn": mo_rn, "mo_adr": mo_adr,
        "ann_rev": ann_rev, "ann_rn": ann_rn, "ann_adr": ann_adr,
        "grand_rev": grand_rev, "grand_rn": grand_rn,
    }


# ─────────────────────────────────────────────────────────────────────────────
# Formatting helpers
# ─────────────────────────────────────────────────────────────────────────────
def js_num(x: float) -> str:
    """Render a Python float for embedding into a JS literal (no trailing zeros stripped)."""
    # Use repr to preserve full precision matching the existing constants style
    if isinstance(x, int) or float(x).is_integer():
        return str(int(round(x)))
    return repr(float(x))


def js_dict_arrays(d: dict[str, list[float]]) -> str:
    parts = []
    for ch in CHANNELS:
        vals = d[ch]
        rendered = "[" + ", ".join(js_num(v) for v in vals) + "]"
        parts.append(f'"{ch}": {rendered}')
    return "{" + ", ".join(parts) + "}"


def js_dict_scalar(d: dict[str, float], rounded: int | None = None) -> str:
    parts = []
    for ch in CHANNELS:
        v = d[ch]
        if rounded is not None:
            v = round(v, rounded)
        parts.append(f'"{ch}": {js_num(v)}')
    return "{" + ", ".join(parts) + "}"


def js_array(a: list[float]) -> str:
    return "[" + ", ".join(js_num(v) for v in a) + "]"


def fmt_m(v: float) -> str:
    return f"${v / 1e6:.2f}M"


def fmt_dollar(v: float) -> str:
    return f"${v:.2f}"


def fmt_int(v: float) -> str:
    return f"{int(round(v)):,}"


def pct_delta(curr: float, prev: float) -> tuple[float, str, str]:
    """Return (delta_pct, color_hex, arrow_glyph)."""
    if prev == 0:
        return 0.0, "#5A6B84", ""
    delta = (curr - prev) / prev * 100
    if delta >= 0:
        return delta, "#2E8B57", "▲"  # ▲
    return delta, "#C0392B", "▼"  # ▼


# ─────────────────────────────────────────────────────────────────────────────
# Build the data constants block (one big multi-line block)
# ─────────────────────────────────────────────────────────────────────────────
def build_constants_block(cy_data: dict, py_data: dict) -> str:
    months_js = json.dumps(MONTHS)
    channels_js = (
        '["PMS", "Expedia", "Website", "Booking", "GDS", "Expedia Programs", '
        '"Other OTA", "Priceline/Agoda", "HotelBeds", "BnBerry", '
        '"Hotel Tonight", "Hopper"]'
    )
    period_label = f"{MONTHS[0]} – {MONTHS[-1]} {CY}"

    block = []
    block.append(f"const MONTHS  = {months_js};")
    block.append(f"const CHANNELS = {channels_js};")
    block.append(f'const PERIOD_LABEL = "{period_label}";')
    block.append(
        "const COLORS = {\n"
        "  'PMS':'#1C355E','Website':'#2E8B57','Expedia':'#E03030','Booking':'#FF5F00',\n"
        "  'GDS':'#6B4CA8','Expedia Programs':'#0088CC','Priceline/Agoda':'#D4822A',\n"
        "  'Other OTA':'#6B8CAD','HotelBeds':'#4DAA8C','Hopper':'#9B59B6',\n"
        "  'Hotel Tonight':'#C0392B','BnBerry':'#95A5A6',\n"
        "};"
    )
    block.append("")
    # 2026 (current year) monthly per-channel
    block.append(
        f"const monthlyRev26={js_dict_arrays(cy_data['rev'])}, "
        f"monthlyRn26={js_dict_arrays(cy_data['rn'])}, "
        f"monthlyAdr26={js_dict_arrays(cy_data['adr'])};"
    )
    # 2025 same period
    block.append(
        f"const monthlyRev25={js_dict_arrays(py_data['rev'])}, "
        f"monthlyRn25={js_dict_arrays(py_data['rn'])}, "
        f"monthlyAdr25={js_dict_arrays(py_data['adr'])};"
    )
    # Monthly totals
    block.append(
        f"const moRev26={js_array(cy_data['mo_rev'])}, "
        f"moRn26={js_array(cy_data['mo_rn'])}, "
        f"moAdr26={js_array([round(v, 2) for v in cy_data['mo_adr']])};"
    )
    block.append(
        f"const moRev25={js_array(py_data['mo_rev'])}, "
        f"moRn25={js_array(py_data['mo_rn'])}, "
        f"moAdr25={js_array([round(v, 2) for v in py_data['mo_adr']])};"
    )
    # YTD per-channel totals
    block.append(
        f"const annRev26={js_dict_scalar(cy_data['ann_rev'])}, "
        f"annRn26={js_dict_scalar(cy_data['ann_rn'])}, "
        f"annAdr26={js_dict_scalar(cy_data['ann_adr'], rounded=2)};"
    )
    block.append(
        f"const annRev25={js_dict_scalar(py_data['ann_rev'])}, "
        f"annRn25={js_dict_scalar(py_data['ann_rn'])}, "
        f"annAdr25={js_dict_scalar(py_data['ann_adr'], rounded=2)};"
    )
    # Grand totals
    block.append(
        f"const grandRev26={int(round(cy_data['grand_rev']))}, "
        f"grandRn26={int(round(cy_data['grand_rn']))};"
    )
    block.append(
        f"const grandRev25={int(round(py_data['grand_rev']))}, "
        f"grandRn25={int(round(py_data['grand_rn']))};"
    )
    return "\n".join(block)


# ─────────────────────────────────────────────────────────────────────────────
# Layout fragments (sticky nav, year toggles, merged revenue table styles/JS)
# ─────────────────────────────────────────────────────────────────────────────
STICKY_NAV_CSS = """
  /* Sticky section navigation */
  .section-nav{position:sticky;top:0;z-index:100;background:#FFFFFF;border-bottom:1px solid var(--border);box-shadow:0 1px 4px rgba(28,53,94,0.06);}
  .section-nav-inner{max-width:1600px;margin:0 auto;padding:0 32px;display:flex;gap:4px;overflow-x:auto;}
  .section-nav a{display:inline-block;padding:12px 16px;font-size:12px;font-weight:700;color:var(--navy);text-decoration:none;letter-spacing:0.4px;text-transform:uppercase;border-bottom:3px solid transparent;transition:all 0.15s;white-space:nowrap;}
  .section-nav a:hover{color:var(--orange);}
  .section-nav a.active{color:var(--orange);border-bottom-color:var(--orange);}

  /* Per-section year toggle pills */
  .year-toggle{display:inline-flex;gap:4px;background:#EEF1F6;padding:3px;border-radius:20px;margin-left:auto;}
  .year-toggle button{font-family:Arial,'Helvetica Neue',Helvetica,sans-serif;font-size:11px;font-weight:700;padding:5px 14px;border-radius:16px;border:none;background:transparent;color:var(--text);cursor:pointer;transition:all 0.15s;letter-spacing:0.3px;}
  .year-toggle button:hover{background:rgba(28,53,94,0.08);}
  .year-toggle button.active{background:var(--navy);color:#fff;}
  .year-toggle button.active:hover{background:var(--navy);}
  .chart-card-head{display:flex;align-items:center;gap:14px;flex-wrap:wrap;margin-bottom:8px;border-bottom:2px solid #FF5F00;padding-bottom:8px;}
  .chart-card-head h3{border-bottom:none !important;padding-bottom:0 !important;margin-bottom:0 !important;flex:1;min-width:240px;}

  /* Side-by-side chart grid (used when year toggle = both) */
  .yt-pair{display:grid;grid-template-columns:1fr 1fr;gap:14px;}
  .yt-pair .yt-sub-title{font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:0.6px;margin-bottom:6px;text-align:center;}
  .yt-single canvas{display:block;}
  .yt-pair canvas{display:block;}

  /* ADR section: wider left lane (62/38) */
  .grid-adr{display:grid;grid-template-columns:62fr 38fr;gap:18px;margin-bottom:22px;}

  /* Merged revenue table: visual grouping for month blocks */
  table.merged-rev th.mblk-start,table.merged-rev td.mblk-start{border-left:2px solid #1C355E;}
  table.merged-rev th.mblk-mid,table.merged-rev td.mblk-mid{}
  table.merged-rev th.mblk-end,table.merged-rev td.mblk-end{border-right:2px solid #1C355E;}
  table.merged-rev thead th.month-group{background:#1C355E;color:#fff;text-align:center;font-size:10px;letter-spacing:0.5px;}
  table.merged-rev thead th.sub-hdr{font-size:9px;font-weight:600;background:#2A4470;}
  table.merged-rev tbody td.delta-cell{font-size:10px;font-weight:700;}

  @media(max-width:1100px){
    .grid-adr{grid-template-columns:1fr;}
    .yt-pair{grid-template-columns:1fr;}
  }
"""


def build_sticky_nav() -> str:
    return (
        '<div class="section-nav">'
        '<div class="section-nav-inner">'
        '<a href="#sec-revenue" data-target="sec-revenue">Revenue</a>'
        '<a href="#sec-rn" data-target="sec-rn">Room Nights</a>'
        '<a href="#sec-mix" data-target="sec-mix">Channel Mix</a>'
        '<a href="#sec-adr" data-target="sec-adr">ADR</a>'
        '<a href="#sec-detail" data-target="sec-detail">Detail Tables</a>'
        '</div></div>'
    )


def build_revenue_section(period_str: str) -> str:
    return f"""  <!-- REVENUE SECTION (year toggle: 2026 / 2025 / Side-by-Side) -->
  <div id="sec-revenue" class="chart-card" data-yt-section="rev">
    <div class="chart-card-head">
      <h3 id="revTitle">Monthly Net Revenue by Channel — {CY} YTD ({period_str})</h3>
      <div class="year-toggle" data-yt-target="rev">
        <button data-yt="2026" class="active">{CY}</button>
        <button data-yt="2025">{PY}</button>
        <button data-yt="both">Side-by-Side</button>
      </div>
    </div>
    <p class="chart-desc" id="revDesc">Stacked monthly net revenue by booking channel for {CY} YTD. Use <strong>All Off</strong> + a single chip to isolate and analyze any individual channel.</p>
    <div class="legend-chips" id="revLegend"><span class="solo-label" id="revSoloLabel"></span></div>
    <div class="yt-single" data-yt-mode="single">
      <canvas id="revChartSingle" height="320" style="width:100%;"></canvas>
    </div>
    <div class="yt-pair" data-yt-mode="pair" style="display:none;">
      <div>
        <div class="yt-sub-title">{CY} YTD</div>
        <canvas id="revChartPair26" height="280" style="width:100%;"></canvas>
      </div>
      <div>
        <div class="yt-sub-title">{PY} Same Period</div>
        <canvas id="revChartPair25" height="280" style="width:100%;"></canvas>
      </div>
    </div>
  </div>
"""


def build_rn_section(period_str: str) -> str:
    return f"""  <!-- ROOM NIGHTS SECTION (year toggle) -->
  <div id="sec-rn" class="chart-card" data-yt-section="rn">
    <div class="chart-card-head">
      <h3 id="rnTitle">Monthly Net Room Nights by Channel — {CY} YTD ({period_str})</h3>
      <div class="year-toggle" data-yt-target="rn">
        <button data-yt="2026" class="active">{CY}</button>
        <button data-yt="2025">{PY}</button>
        <button data-yt="both">Side-by-Side</button>
      </div>
    </div>
    <p class="chart-desc" id="rnDesc">Monthly room night production by booking channel for {CY} YTD. Reflects demand volume and channel mix.</p>
    <div class="legend-chips" id="rnLegend"><span class="solo-label" id="rnSoloLabel"></span></div>
    <div class="yt-single" data-yt-mode="single">
      <canvas id="rnChartSingle" height="280" style="width:100%;"></canvas>
    </div>
    <div class="yt-pair" data-yt-mode="pair" style="display:none;">
      <div>
        <div class="yt-sub-title">{CY} YTD</div>
        <canvas id="rnChartPair26" height="240" style="width:100%;"></canvas>
      </div>
      <div>
        <div class="yt-sub-title">{PY} Same Period</div>
        <canvas id="rnChartPair25" height="240" style="width:100%;"></canvas>
      </div>
    </div>
  </div>
"""


def build_pct_section() -> str:
    return f"""  <!-- CONTRIBUTION % SECTION (year toggle) -->
  <div id="sec-mix" class="chart-card" data-yt-section="pct">
    <div class="chart-card-head">
      <h3 id="pctTitle">Monthly Revenue Contribution % by Channel — {CY} YTD</h3>
      <div class="year-toggle" data-yt-target="pct">
        <button data-yt="2026" class="active">{CY}</button>
        <button data-yt="2025">{PY}</button>
        <button data-yt="both">Side-by-Side</button>
      </div>
    </div>
    <p class="chart-desc" id="pctDesc">Each bar totals 100% and shows each channel's proportional share of total monthly revenue for {CY}. Useful for spotting early shifts in channel mix vs the prior year.</p>
    <div class="yt-single" data-yt-mode="single">
      <canvas id="pctChartSingle" height="260" style="width:100%;"></canvas>
    </div>
    <div class="yt-pair" data-yt-mode="pair" style="display:none;">
      <div>
        <div class="yt-sub-title">{CY} YTD</div>
        <canvas id="pctChartPair26" height="220" style="width:100%;"></canvas>
      </div>
      <div>
        <div class="yt-sub-title">{PY} Same Period</div>
        <canvas id="pctChartPair25" height="220" style="width:100%;"></canvas>
      </div>
    </div>
    <div class="legend-chips" id="pctLegend" style="margin-top:14px;margin-bottom:0;"></div>
  </div>
"""


def build_merged_revenue_table_section() -> str:
    return f"""  <!-- MERGED REVENUE DETAIL TABLE -->
  <div id="sec-detail" class="chart-card">
    <h3>Monthly Net Revenue by Channel — {CY} YTD vs {PY}</h3>
    <p class="chart-desc">Month-by-month net revenue per channel with {CY} (current), {PY} (same period), and the YoY % change shown side-by-side. Each three-column block represents one month; the rightmost block is the YTD total.</p>
    <div class="tbl-wrap"><table id="revTableMerged" class="merged-rev"></table></div>
  </div>
"""


# JavaScript helpers/functions for the new behavior. We inject these inside the
# <script> block before the existing `function init()` to extend behavior.
def build_new_js_helpers() -> str:
    return """
// ── Year Toggle State (per section) ───────────────────────────────────────────
const yearState = { rev: '2026', rn: '2026', pct: '2026' };
const SECTION_TITLES = {
  rev: {
    '2026': ['Monthly Net Revenue by Channel — __CY__ YTD (__PERIOD__)', "Stacked monthly net revenue by booking channel for __CY__ YTD. Use <strong>All Off</strong> + a single chip to isolate and analyze any individual channel."],
    '2025': ['Monthly Net Revenue by Channel — __PY__ Same Period (__PERIOD__)', "__PY__ revenue for the same __N__-month window, shown alone for a focused historical view. Channel toggles still apply."],
    'both': ['Monthly Net Revenue by Channel — __CY__ vs __PY__', "Stacked monthly net revenue by channel — __CY__ on the left, __PY__ same period on the right — for direct apples-to-apples comparison."]
  },
  rn: {
    '2026': ['Monthly Net Room Nights by Channel — __CY__ YTD (__PERIOD__)', 'Monthly room night production by booking channel for __CY__ YTD. Reflects demand volume and channel mix.'],
    '2025': ['Monthly Net Room Nights by Channel — __PY__ Same Period (__PERIOD__)', '__PY__ room night volume for the same __N__-month window. Compare against __CY__ to identify which channels are driving growth.'],
    'both': ['Monthly Net Room Nights by Channel — __CY__ vs __PY__', 'Monthly room night volume by channel — __CY__ on the left, __PY__ same period on the right.']
  },
  pct: {
    '2026': ['Monthly Revenue Contribution % by Channel — __CY__ YTD', "Each bar totals 100% and shows each channel's proportional share of total monthly revenue for __CY__."],
    '2025': ['Monthly Revenue Contribution % by Channel — __PY__ Same Period', "__PY__ channel mix for the same __N__-month window. Compare against __CY__ to identify which channels are gaining or losing share."],
    'both': ['Monthly Revenue Contribution % by Channel — __CY__ vs __PY__', 'Channel mix percentages by month — __CY__ on the left, __PY__ same period on the right.']
  }
};

function applySectionTitle(section, mode) {
  const tpl = SECTION_TITLES[section][mode];
  if (!tpl) return;
  const period = (typeof PERIOD_LABEL!=='undefined')?PERIOD_LABEL:'';
  const nMonths = MONTHS.length;
  const titleEl = document.getElementById(section + 'Title');
  const descEl  = document.getElementById(section + 'Desc');
  const t = tpl[0].replaceAll('__CY__', '2026').replaceAll('__PY__', '2025').replaceAll('__PERIOD__', period).replaceAll('__N__', nMonths);
  const d = tpl[1].replaceAll('__CY__', '2026').replaceAll('__PY__', '2025').replaceAll('__PERIOD__', period).replaceAll('__N__', nMonths);
  if (titleEl) titleEl.textContent = t;
  if (descEl)  descEl.innerHTML  = d;
}

function setYear(section, mode) {
  yearState[section] = mode;
  // Update active button
  const wrap = document.querySelector('[data-yt-target="'+section+'"]');
  if (wrap) wrap.querySelectorAll('button').forEach(b => b.classList.toggle('active', b.dataset.yt===mode));
  // Toggle single vs pair containers
  const card = document.querySelector('[data-yt-section="'+section+'"]');
  if (card) {
    const single = card.querySelector('[data-yt-mode="single"]');
    const pair   = card.querySelector('[data-yt-mode="pair"]');
    if (mode === 'both') {
      if (single) single.style.display = 'none';
      if (pair)   pair.style.display = 'grid';
    } else {
      if (single) single.style.display = 'block';
      if (pair)   pair.style.display = 'none';
    }
  }
  applySectionTitle(section, mode);
  renderSection(section);
}

function renderSection(section) {
  const mode = yearState[section];
  if (section === 'rev') {
    if (mode === 'both') {
      drawBar('revChartPair26', monthlyRev26, moRev26, fmtM);
      drawBar('revChartPair25', monthlyRev25, moRev25, fmtM);
    } else if (mode === '2026') {
      drawBar('revChartSingle', monthlyRev26, moRev26, fmtM);
    } else {
      drawBar('revChartSingle', monthlyRev25, moRev25, fmtM);
    }
  } else if (section === 'rn') {
    if (mode === 'both') {
      drawBar('rnChartPair26', monthlyRn26, moRn26, v=>fmtK(v));
      drawBar('rnChartPair25', monthlyRn25, moRn25, v=>fmtK(v));
    } else if (mode === '2026') {
      drawBar('rnChartSingle', monthlyRn26, moRn26, v=>fmtK(v));
    } else {
      drawBar('rnChartSingle', monthlyRn25, moRn25, v=>fmtK(v));
    }
  } else if (section === 'pct') {
    if (mode === 'both') {
      drawPct('pctChartPair26', monthlyRev26, moRev26);
      drawPct('pctChartPair25', monthlyRev25, moRev25);
    } else if (mode === '2026') {
      drawPct('pctChartSingle', monthlyRev26, moRev26);
    } else {
      drawPct('pctChartSingle', monthlyRev25, moRev25);
    }
  }
}

function renderAllToggleSections() { renderSection('rev'); renderSection('rn'); renderSection('pct'); }

function wireYearToggles() {
  document.querySelectorAll('.year-toggle').forEach(tg => {
    const section = tg.dataset.ytTarget;
    tg.querySelectorAll('button').forEach(btn => {
      btn.addEventListener('click', () => setYear(section, btn.dataset.yt));
    });
  });
}

// ── Sticky Section Nav (scroll-spy) ──────────────────────────────────────────
function wireSectionNav() {
  const links = document.querySelectorAll('.section-nav a');
  const sections = Array.from(links).map(a => document.getElementById(a.dataset.target)).filter(Boolean);
  function onScroll() {
    const navH = document.querySelector('.section-nav').offsetHeight + 20;
    let activeIdx = 0;
    for (let i = 0; i < sections.length; i++) {
      const rect = sections[i].getBoundingClientRect();
      if (rect.top - navH < 10) activeIdx = i;
    }
    links.forEach((l, i) => l.classList.toggle('active', i === activeIdx));
  }
  window.addEventListener('scroll', onScroll, { passive: true });
  onScroll();
  // Smooth scroll with offset for sticky nav
  links.forEach(l => l.addEventListener('click', e => {
    e.preventDefault();
    const id = l.dataset.target;
    const el = document.getElementById(id);
    if (!el) return;
    const navH = document.querySelector('.section-nav').offsetHeight;
    const y = el.getBoundingClientRect().top + window.pageYOffset - navH - 12;
    window.scrollTo({ top: y, behavior: 'smooth' });
  }));
}

// ── Merged Revenue Detail Table (Channel × [Mo26, Mo25, MoΔ%] × N months + YTD) ──
function renderMergedRevenueTable() {
  const fmt = v => v >= 1e6 ? '$'+(v/1e6).toFixed(2)+'M' : (v>0 ? '$'+Math.round(v).toLocaleString() : '—');
  const delta = (a, b) => {
    if (b === 0) return ['—', '#5A6B84'];
    const pct = (a - b) / b * 100;
    const color = pct >= 0 ? '#2E8B57' : '#C0392B';
    const arrow = pct >= 0 ? '▲' : '▼';
    return [arrow + Math.abs(pct).toFixed(1) + '%', color];
  };
  // Header: two rows. Top row = month group labels (3-cell colspan each).
  // Bottom row = sub-headers (2026 | 2025 | Δ%).
  let html = '<thead>';
  html += '<tr><th rowspan="2">Channel</th>';
  MONTHS.forEach(m => { html += '<th colspan="3" class="month-group">' + m + '</th>'; });
  html += '<th colspan="3" class="month-group">YTD</th>';
  html += '</tr><tr>';
  for (let i = 0; i < MONTHS.length + 1; i++) {
    html += '<th class="sub-hdr mblk-start">2026</th><th class="sub-hdr mblk-mid">2025</th><th class="sub-hdr mblk-end">Δ%</th>';
  }
  html += '</tr></thead><tbody>';

  CHANNELS.forEach(ch => {
    const c26arr = monthlyRev26[ch] || [], c25arr = monthlyRev25[ch] || [];
    const ytd26 = c26arr.reduce((s,v)=>s+v,0), ytd25 = c25arr.reduce((s,v)=>s+v,0);
    html += '<tr><td style="color:'+COLORS[ch]+'">● '+ch+'</td>';
    MONTHS.forEach((_, mi) => {
      const v26 = c26arr[mi] || 0, v25 = c25arr[mi] || 0;
      const [d, dc] = delta(v26, v25);
      html += '<td class="mblk-start">'+fmt(v26)+'</td>';
      html += '<td class="mblk-mid" style="color:#5A6B84;">'+fmt(v25)+'</td>';
      html += '<td class="mblk-end delta-cell" style="color:'+dc+';">'+d+'</td>';
    });
    const [yd, ydc] = delta(ytd26, ytd25);
    html += '<td class="mblk-start" style="font-weight:700;">'+fmt(ytd26)+'</td>';
    html += '<td class="mblk-mid" style="color:#5A6B84;font-weight:700;">'+fmt(ytd25)+'</td>';
    html += '<td class="mblk-end delta-cell" style="color:'+ydc+';">'+yd+'</td>';
    html += '</tr>';
  });

  // Total row
  const t26 = moRev26.reduce((s,v)=>s+v,0), t25 = moRev25.reduce((s,v)=>s+v,0);
  html += '<tr class="total-row"><td>TOTAL</td>';
  MONTHS.forEach((_, mi) => {
    const v26 = moRev26[mi] || 0, v25 = moRev25[mi] || 0;
    const [d] = delta(v26, v25);
    html += '<td class="mblk-start">'+fmt(v26)+'</td>';
    html += '<td class="mblk-mid">'+fmt(v25)+'</td>';
    html += '<td class="mblk-end">'+d+'</td>';
  });
  const [yd] = delta(t26, t25);
  html += '<td class="mblk-start">'+fmt(t26)+'</td><td class="mblk-mid">'+fmt(t25)+'</td><td class="mblk-end">'+yd+'</td>';
  html += '</tr></tbody>';
  document.getElementById('revTableMerged').innerHTML = html;
}

// Override the legacy refreshAllCharts (called by chip toggles) to drive the
// new toggle-aware containers.
function refreshAllCharts() {
  renderAllToggleSections();
}

// Adjust renderChips so the single shared legend container drives all charts.
const SHARED_LEGENDS = ['revLegend', 'rnLegend', 'pctLegend'];
const SHARED_SOLOS   = ['revSoloLabel', 'rnSoloLabel'];

function updateSoloLabels() {
  const active = CHANNELS.filter(c => activeChannels.has(c));
  const isSolo = active.length === 1;
  const name = isSolo ? active[0] : '', col = isSolo ? COLORS[name] : '#FF5F00';
  SHARED_SOLOS.forEach(id => {
    const el = document.getElementById(id); if (!el) return;
    if (isSolo) {
      el.classList.add('visible'); el.style.color = col;
      el.style.background = col + '15'; el.style.borderColor = col + '50';
      el.innerHTML = `<span style="width:8px;height:8px;border-radius:50%;background:${col};display:inline-block;"></span> ${name} — solo view`;
    } else el.classList.remove('visible');
  });
}

function renderChips(id) {
  const el = document.getElementById(id); if (!el) return;
  const solo = el.querySelector('.solo-label');
  el.innerHTML = `<div class="chip-ctrl">
    <button class="ctrl-btn" onclick="selAll()">All On</button>
    <button class="ctrl-btn" onclick="selNone()">All Off</button>
  </div>` + CHANNELS.map(ch => `
    <div class="chip ${activeChannels.has(ch)?'':'inactive'}" onclick="toggle('${ch}')">
      <div class="chip-dot" style="background:${COLORS[ch]}"></div>${ch}
    </div>`).join('');
  if (solo) el.appendChild(solo);
}
function selAll()  { CHANNELS.forEach(c=>activeChannels.add(c)); SHARED_LEGENDS.forEach(renderChips); updateSoloLabels(); refreshAllCharts(); }
function selNone() { activeChannels.clear(); SHARED_LEGENDS.forEach(renderChips); updateSoloLabels(); refreshAllCharts(); }
function toggle(ch){ activeChannels.has(ch)?activeChannels.delete(ch):activeChannels.add(ch); SHARED_LEGENDS.forEach(renderChips); updateSoloLabels(); refreshAllCharts(); }
"""


def build_new_init_block() -> str:
    """Replacement for the legacy init() function + window listeners."""
    return """function init() {
  SHARED_LEGENDS.forEach(renderChips); updateSoloLabels();
  wireYearToggles();
  wireSectionNav();
  // Render the three toggle-controlled sections in their default 2026 view
  renderAllToggleSections();
  drawAdrLine(); renderDonuts(); renderAdrBars(); renderYoyBars();
  renderMergedRevenueTable();
  renderMergedRnTable(); renderMergedPctTable();
}

window.addEventListener('resize', () => {
  renderAllToggleSections();
  drawAdrLine();
});
window.addEventListener('load', init);"""


LAYOUT_MARKER = "/* MPH-LAYOUT-OVERHAUL-v1 */"


def apply_layout_overhaul(html: str) -> str:
    """Idempotently update the dashboard HTML to the new layout.

    Detects whether the overhaul has been applied (looks for LAYOUT_MARKER in
    the <style> block). If absent, performs the full set of swaps:
      • Inject CSS for sticky nav, year-toggle pills, grid-adr, merged-rev table
      • Insert sticky nav HTML after the header
      • Collapse Revenue / RN / Contribution % paired chart-cards into single
        toggle-controlled sections (with single + side-by-side containers)
      • Replace the two revenue detail tables with one merged table
      • Switch the ADR grid from grid-2 to grid-adr (62/38)
      • Inject new JS helpers and replace init()
    Period-string placeholders in section titles are filled in by the period
    rewrite step that runs after this.
    """
    first_month = MONTHS[0]
    last_month = MONTHS[-1]
    period_str = f"{first_month} – {last_month} {CY}"

    # 1) CSS injection (only once)
    if LAYOUT_MARKER not in html:
        css_block = "\n" + LAYOUT_MARKER + "\n" + STICKY_NAV_CSS + "\n  "
        html = html.replace("@media(max-width:1100px)", css_block + "@media(max-width:1100px)", 1)

    # 2) Sticky section nav — insert right after the header closing `</div>` and
    # before `<div class="dashboard">`. Idempotent via marker check.
    if 'class="section-nav"' not in html:
        nav_html = build_sticky_nav()
        html = html.replace(
            '<div class="dashboard">',
            nav_html + '\n\n<div class="dashboard">',
            1,
        )

    # 3) Replace the two REVENUE chart cards (2026 + 2025) with a single
    # toggle-driven section. Anchor by the comment AFTER the pair so we
    # consume both cards cleanly without ambiguity over which </div> to stop at.
    rev_pair = re.compile(
        r'<!-- REVENUE CHART 2026 -->.*?(?=<!-- RN CHART 2026 -->)',
        re.DOTALL,
    )
    if rev_pair.search(html):
        html = rev_pair.sub(build_revenue_section(period_str), html, count=1)

    # 4) Replace the two RN chart cards
    rn_pair = re.compile(
        r'<!-- RN CHART 2026 -->.*?(?=<!-- CONTRIBUTION % 2026 -->)',
        re.DOTALL,
    )
    if rn_pair.search(html):
        html = rn_pair.sub(build_rn_section(period_str), html, count=1)

    # 5) Replace the two Contribution % chart cards
    pct_pair = re.compile(
        r'<!-- CONTRIBUTION % 2026 -->.*?(?=<!-- DONUT — Revenue share -->)',
        re.DOTALL,
    )
    if pct_pair.search(html):
        html = pct_pair.sub(build_pct_section(), html, count=1)

    # 6) Section IDs for the donut row (channel mix overflow) + ADR + detail.
    # The donut rows directly follow the pct section and don't have an id of
    # their own — the user requested #sec-mix on Channel Mix. We've already
    # set #sec-mix on the pct (contribution %) card, which is the canonical
    # "channel mix" section. Donuts continue as supporting context.

    # 7) ADR grid — swap grid-2 → grid-adr ONLY for the ADR block. Anchor via
    # the surrounding `<!-- ADR -->` comment.
    html = re.sub(
        r'<!-- ADR -->\s*\n\s*<div class="grid-2">',
        '<!-- ADR -->\n  <div id="sec-adr" class="grid-adr">',
        html, count=1,
    )

    # 8) Replace the two revenue detail tables with one merged table. Anchor
    # via `<!-- DETAIL TABLES -->` and consume through to the `<!-- MERGED RN`
    # comment that always follows.
    detail_pair = re.compile(
        r'<!-- DETAIL TABLES -->.*?(?=<!-- MERGED RN TABLE -->)',
        re.DOTALL,
    )
    if detail_pair.search(html) and 'id="revTableMerged"' not in html:
        replacement = '<!-- DETAIL TABLES -->\n' + build_merged_revenue_table_section() + '\n  '
        html = detail_pair.sub(replacement, html, count=1)
    # Clean up stale placeholder div from earlier overhaul iterations.
    html = re.sub(r'\n?\s*<div id="sec-detail"></div>\n', '\n', html)
    # Ensure the merged revenue chart-card has id="sec-detail" so anchor links work.
    html = re.sub(
        r'<!-- MERGED REVENUE DETAIL TABLE -->\s*\n\s*<div class="chart-card">',
        '<!-- MERGED REVENUE DETAIL TABLE -->\n  <div id="sec-detail" class="chart-card">',
        html, count=1,
    )

    # 9) JS: inject new helpers + replace the legacy init() block. We swap out
    # the old `updateSoloLabels`, `renderChips`, `selAll/selNone/toggle`,
    # `refreshAllCharts`, `init`, and the resize/load listeners. Easiest path:
    # remove old functions by anchor and append the new block right before
    # `</script>`.
    js_marker = "// MPH-JS-OVERHAUL-v1"
    if js_marker not in html:
        # Remove the legacy refreshAllCharts/updateSoloLabels/renderChips/selAll/selNone/toggle
        # We'll cut from `// ── Channel Toggles ───` through the end of the
        # `function toggle(ch){...}` line.
        html = re.sub(
            r'// ── Channel Toggles ─+.*?function toggle\(ch\)\{[^}]*\}\n',
            "// (legacy channel-toggle block replaced by MPH-JS-OVERHAUL-v1)\n"
            "const activeChannels = new Set(CHANNELS);\n",
            html, count=1, flags=re.DOTALL,
        )
        # Remove legacy init() and the two window listeners at the bottom
        html = re.sub(
            r'// ── Init ─+.*?window\.addEventListener\(\'load\',init\);',
            "// (legacy init replaced by MPH-JS-OVERHAUL-v1)",
            html, count=1, flags=re.DOTALL,
        )
        new_js = (
            "\n" + js_marker + "\n"
            + build_new_js_helpers()
            + "\n"
            + build_new_init_block()
            + "\n"
        )
        html = html.replace("</script>", new_js + "</script>", 1)

    return html


# ─────────────────────────────────────────────────────────────────────────────
# HTML rewrite
# ─────────────────────────────────────────────────────────────────────────────
def rewrite_html(cy_data: dict, py_data: dict) -> str:
    html = HTML_PATH.read_text(encoding="utf-8")

    # ── 1) Replace the JS data constants block ──────────────────────────────
    # The block spans from `const MONTHS  = ` to `grandRn25=...;` inclusive.
    new_block = build_constants_block(cy_data, py_data)
    pattern = re.compile(
        r"const MONTHS\s*=.*?const grandRev25=.*?grandRn25=\d+\s*;",
        re.DOTALL,
    )
    if not pattern.search(html):
        raise RuntimeError("Could not locate the data constants block to replace.")
    html = pattern.sub(new_block, html, count=1)

    # ── 1b) Inject CSS, sticky nav, new section markup, merged table, and JS ─
    html = apply_layout_overhaul(html)

    # ── 2) Period labels and text ───────────────────────────────────────────
    first_month = MONTHS[0]
    last_month = MONTHS[-1]
    last_full = MONTH_FULL[last_month]
    period_str = f"{first_month} – {last_month} {CY}"  # "Jan – Apr 2026"
    months_count = len(MONTHS)
    today_str = date.today().strftime("%B %-d, %Y") if sys.platform != "win32" else date.today().strftime("%B %#d, %Y")

    # Badge label: anchor to the most recent month included (e.g. "Apr YTD"),
    # except Q1 and Full-Year which have natural labels.
    if months_count == 3:
        q_label = "Q1"
    elif months_count == 12:
        q_label = "Full Year"
    else:
        q_label = f"{last_month} YTD"

    # Subtitle line
    html = re.sub(
        r'<div class="subtitle">[^<]*</div>',
        f'<div class="subtitle">{period_str} · Stay Date · All Properties · '
        f'Excludes Cancelled Reservations · Compared vs same period {PY}</div>',
        html, count=1,
    )

    # Update badge
    html = re.sub(
        r'<div class="update-badge">[^<]*</div>',
        f'<div class="update-badge">Last updated: {today_str} · '
        f'Data through {last_month} {CY} ({months_count} of 12 months)</div>',
        html, count=1,
    )

    # YTD badge
    if q_label.endswith("YTD"):
        # avoid awkward "2026 YTD · Apr YTD" — prefer "2026 YTD · Through Apr"
        badge_text = f"{CY} YTD · Through {last_month}"
    else:
        badge_text = f"{CY} YTD · {q_label}"
    html = re.sub(
        r'<div class="ytd-badge">[^<]*</div>',
        f'<div class="ytd-badge">{badge_text}</div>',
        html, count=1,
    )

    # Chart H3s and chart descriptions that reference old periods
    # We do targeted replacements only for the strings actually present.
    replacements = [
        # Revenue chart 2026
        (
            r"<h3>Monthly Net Revenue by Channel — 2026 YTD \(Jan – [A-Za-z]+ 2026\)</h3>",
            f"<h3>Monthly Net Revenue by Channel — {CY} YTD ({period_str})</h3>",
        ),
        # Revenue chart 2025 same period
        (
            r"<h3>Monthly Net Revenue by Channel — 2025 Same Period \(Jan – [A-Za-z]+ 2026\)</h3>",
            f"<h3>Monthly Net Revenue by Channel — {PY} Same Period ({period_str})</h3>",
        ),
        # RN chart 2026
        (
            r"<h3>Monthly Net Room Nights by Channel — 2026 YTD \(Jan – [A-Za-z]+ 2026\)</h3>",
            f"<h3>Monthly Net Room Nights by Channel — {CY} YTD ({period_str})</h3>",
        ),
        # RN chart 2025 same period
        (
            r"<h3>Monthly Net Room Nights by Channel — 2025 Same Period \(Jan – [A-Za-z]+ 2026\)</h3>",
            f"<h3>Monthly Net Room Nights by Channel — {PY} Same Period ({period_str})</h3>",
        ),
    ]
    for pat, repl in replacements:
        html = re.sub(pat, repl, html)

    # Generic "Jan – {something} 2026" inside chart-desc paragraphs and donut subtitles
    # Replace any "Jan – <Month> 2026" with our current period.
    html = re.sub(
        r"Jan – [A-Za-z]+ 2026",
        period_str,
        html,
    )
    # Replace "through <Month> 2026" with current last month
    html = re.sub(
        r"through (?:Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) 2026",
        f"through {last_month} {CY}",
        html,
    )
    # Replace n-month window phrasing like "the same 2-month window" / "3-month"
    html = re.sub(
        r"the same \d+-month window",
        f"the same {months_count}-month window",
        html,
    )

    # Footer text — period and last updated date
    # Pattern matches "...· Jan – Feb 2026 · Stay-date reports · ... Last updated <date>"
    html = re.sub(
        r"(Channel Performance Dashboard[^<]*?)Last updated [A-Za-z]+ \d+,? \d{4}",
        rf"\1Last updated {today_str}",
        html,
    )
    # And patch the period token in the footer separately (since the date regex handled the tail)
    html = re.sub(
        r"2026 YTD vs 2025 Same Period · [^<·]+ · Stay-date reports",
        f"{CY} YTD vs {PY} Same Period · {period_str} · Stay-date reports",
        html,
    )

    # ── 3) KPI top-row values ───────────────────────────────────────────────
    grand_rev_cy = cy_data["grand_rev"]
    grand_rev_py = py_data["grand_rev"]
    grand_rn_cy = cy_data["grand_rn"]
    grand_rn_py = py_data["grand_rn"]
    blended_adr_cy = grand_rev_cy / grand_rn_cy if grand_rn_cy else 0
    blended_adr_py = grand_rev_py / grand_rn_py if grand_rn_py else 0

    rev_delta, rev_color, rev_arrow = pct_delta(grand_rev_cy, grand_rev_py)
    rn_delta, rn_color, rn_arrow = pct_delta(grand_rn_cy, grand_rn_py)
    adr_delta, adr_color, adr_arrow = pct_delta(blended_adr_cy, blended_adr_py)

    # Top channel = highest CY revenue
    top_channel = max(CHANNELS, key=lambda c: cy_data["ann_rev"][c])
    top_rev_cy = cy_data["ann_rev"][top_channel]
    top_rev_py = py_data["ann_rev"][top_channel]
    top_share = top_rev_cy / grand_rev_cy * 100 if grand_rev_cy else 0
    top_delta, top_color, top_arrow = pct_delta(top_rev_cy, top_rev_py)

    # Best month = highest CY monthly total revenue
    best_mi = max(range(len(MONTHS)), key=lambda i: cy_data["mo_rev"][i])
    best_month = MONTHS[best_mi]
    best_rev = cy_data["mo_rev"][best_mi]
    best_rn = cy_data["mo_rn"][best_mi]
    best_rev_py = py_data["mo_rev"][best_mi]
    best_delta, best_color, best_arrow = pct_delta(best_rev, best_rev_py)

    # Online channels = everything except PMS
    online_rev_cy = sum(cy_data["ann_rev"][c] for c in CHANNELS if c != "PMS")
    online_rev_py = sum(py_data["ann_rev"][c] for c in CHANNELS if c != "PMS")
    online_share = online_rev_cy / grand_rev_cy * 100 if grand_rev_cy else 0
    online_delta, online_color, online_arrow = pct_delta(online_rev_cy, online_rev_py)

    # Rebuild the entire 6-card KPI row deterministically (the block from
    # `<div class="kpi-row">` through its closing `</div>` is replaced).
    new_kpi_row = f"""<div class="kpi-row">
    <div class="kpi-card">
      <div class="kpi-label">{CY} YTD Net Revenue</div>
      <div class="kpi-value">{fmt_m(grand_rev_cy)}</div>
      <div class="kpi-sub">{PY} same period: {fmt_m(grand_rev_py)}</div>
      <div class="kpi-yoy"><span style="color:{rev_color};font-size:11px;font-weight:700;">{rev_arrow} {abs(rev_delta):.1f}% vs {PY}</span></div>
    </div>
    <div class="kpi-card green">
      <div class="kpi-label">{CY} YTD Net Room Nights</div>
      <div class="kpi-value">{fmt_int(grand_rn_cy)}</div>
      <div class="kpi-sub">{PY} same period: {fmt_int(grand_rn_py)}</div>
      <div class="kpi-yoy"><span style="color:{rn_color};font-size:11px;font-weight:700;">{rn_arrow} {abs(rn_delta):.1f}% vs {PY}</span></div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">{CY} YTD Blended ADR</div>
      <div class="kpi-value">{fmt_dollar(blended_adr_cy)}</div>
      <div class="kpi-sub">{PY} same period: {fmt_dollar(blended_adr_py)}</div>
      <div class="kpi-yoy"><span style="color:{adr_color};font-size:11px;font-weight:700;">{adr_arrow} {abs(adr_delta):.1f}% vs {PY}</span></div>
    </div>
    <div class="kpi-card accent">
      <div class="kpi-label">Top Channel — {CY} YTD</div>
      <div class="kpi-value">{fmt_m(top_rev_cy)}</div>
      <div class="kpi-sub">{top_channel} · {top_share:.1f}% of YTD total</div>
      <div class="kpi-yoy"><span style="color:{top_color};font-size:11px;font-weight:700;">{top_arrow} {abs(top_delta):.1f}% vs {PY}</span></div>
    </div>
    <div class="kpi-card green">
      <div class="kpi-label">Best Month — {CY} YTD</div>
      <div class="kpi-value">{best_month}</div>
      <div class="kpi-sub">{fmt_m(best_rev)} · {fmt_int(best_rn)} RNs</div>
      <div class="kpi-yoy"><span style="color:{best_color};font-size:11px;font-weight:700;">{best_arrow} {abs(best_delta):.1f}% vs {PY}</span></div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">{CY} Online Channels YTD</div>
      <div class="kpi-value">{fmt_m(online_rev_cy)}</div>
      <div class="kpi-sub">{online_share:.1f}% of YTD total</div>
      <div class="kpi-yoy"><span style="color:{online_color};font-size:11px;font-weight:700;">{online_arrow} {abs(online_delta):.1f}% vs {PY}</span></div>
    </div>
  </div>"""

    kpi_pattern = re.compile(
        r'<div class="kpi-row">.*?</div>\s*</div>\s*(?=\s*<!-- (?:REVENUE CHART 2026|REVENUE SECTION) )',
        re.DOTALL,
    )
    if not kpi_pattern.search(html):
        raise RuntimeError("Could not locate the KPI row block to replace.")
    html = kpi_pattern.sub(new_kpi_row + "\n\n  ", html, count=1)

    return html


def main() -> None:
    print(f"Building Channel Performance dashboard for {MONTHS[0]}–{MONTHS[-1]} {CY} vs {PY}…")
    print(f"  Months: {MONTHS}")
    print(f"  Source folder: {BASE}")

    cy_data = build_period_data(CY)
    py_data = build_period_data(PY)

    # Sanity print
    print()
    print(f"  {CY} YTD Net Revenue : ${cy_data['grand_rev']:>14,.2f}")
    print(f"  {PY} YTD Net Revenue : ${py_data['grand_rev']:>14,.2f}")
    print(f"  {CY} YTD Net RNs     : {int(round(cy_data['grand_rn'])):>15,}")
    print(f"  {PY} YTD Net RNs     : {int(round(py_data['grand_rn'])):>15,}")
    blended_cy = cy_data['grand_rev'] / cy_data['grand_rn'] if cy_data['grand_rn'] else 0
    blended_py = py_data['grand_rev'] / py_data['grand_rn'] if py_data['grand_rn'] else 0
    print(f"  {CY} YTD Blended ADR : ${blended_cy:.2f}")
    print(f"  {PY} YTD Blended ADR : ${blended_py:.2f}")

    new_html = rewrite_html(cy_data, py_data)
    HTML_PATH.write_text(new_html, encoding="utf-8")
    print()
    print(f"Wrote {HTML_PATH} ({len(new_html):,} bytes)")


if __name__ == "__main__":
    main()
