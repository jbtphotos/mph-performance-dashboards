#!/usr/bin/env python3
"""
build_ga4_ytd_dashboard.py
Generates ga4-ytd-dashboard.html
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
OUT_FILE = os.path.join(SCRIPT_DIR, 'ga4-ytd-dashboard.html')

# ── Constants ─────────────────────────────────────────────────────────────
WIFI_CH   = "Property WiFi"   # matched via startswith to handle varying trailing spaces
CUTOFF_25 = "20250430"           # YTD same-period end date for 2025
CUTOFF_26 = "20260430"           # YTD end date for 2026

REMOVED_HOTELS = {
    '/wifi/north-aurora-il',
    '/wifi/spokane-valley',
    '/wifi/spokane-valley-wa',
    '/wifi/avondale-az',
    '/locations/my-place-hotel-north-aurora-il',
    '/locations/my-place-hotel-spokane-valley',
    '/locations/my-place-hotel-avondale-az',
}

LP_PAGE_MERGES = {
    '/locations/my-place-hotel-nampa-idaho-center': '/locations/my-place-hotel-nampa-ididaho-center',
    '/locations/my-place-hotel-davenport-ialacehotels.com/my-place-hotel-davenport-ia': '/locations/my-place-hotel-davenport-ia',
    '/locations/my-place-hotel-tucson-north-maranacortaro-': '/locations/my-place-hotel-tucson-north-maranacortaro-az',
    '/locations/my-place-hotel-augusta-ga-2': '/locations/my-place-hotel-augusta',
    '/locations/my-place-hotel-augusta-ga': '/locations/my-place-hotel-augusta',
    '/locations/my-place-hotel-yakima-wam/my-place-hotel-yakima-wa': '/locations/my-place-hotel-yakima-wa',
    '/locations/my-place-hotel-idaho-falls-id[https://...CDN junk...': '/locations/my-place-hotel-idaho-falls-id',
    '/locations/randolph-vt': '/locations/my-place-hotel-randolph-vt',
    '/locations/my-place-hotel-Randolph-vt': '/locations/my-place-hotel-randolph-vt',
    '/locations/my-place-hotel-randolph-vtBringing': '/locations/my-place-hotel-randolph-vt',
    '/locations/my-place-hotel-jonesboro': '/locations/my-place-hotel-jonesboro-ar',
    '/locations/my-place-h\u2026"': None,
    '/locations/my-place-hotel-anchorage-a': '/locations/my-place-hotel-anchorage-ak',
    # ── Corrupted hotel property URLs (discovered in quarterly data) ──
    '/locations/my-place-hoel-amarillo-tx': '/locations/my-place-hotel-amarillo-tx',
    '/locations/my-place-hotel-amarillo-t': '/locations/my-place-hotel-amarillo-tx',
    '/locations/my-place-hotel-lubbock-t': '/locations/my-place-hotel-lubbock-tx',
    '/locations/my-place-hotel-raleigh-': '/locations/my-place-hotel-raleigh-garner-nc',
    '/locations/my-place-hotel-bentonville-rogers.com': '/locations/my-place-hotel-bentonville-rogers-ar',
    '/locations/my-place-hotel-council-bluffs-ia.': '/locations/my-place-hotel-council-bluffs-ia',
    '/locations/my-place-hotel-colorado-springs-coF': '/locations/my-place-hotel-colorado-springs-co',
    '/locations/my-place-hotel-ketchikan-ak.': '/locations/my-place-hotel-ketchikan-ak',
    '/locations/my-place-hotel-raleigh-garner-nc</font': '/locations/my-place-hotel-raleigh-garner-nc',
    '/locations/my-place-hotel-tucson-north-maranacortaro-az)': '/locations/my-place-hotel-tucson-north-maranacortaro-az',
    '/locations/my-place-hotel-tucson-south-az-2': '/locations/my-place-hotel-tucson-south-az',
    '/locations/my-place-hotel-vancouver-wa': '/locations/my-place-hotels-vancouver-wa',
    '/locations/my-place-hotel-portland-or': '/locations/my-place-hotel-portland-east-or',
    '/locations/myplace-hotel-hurricane-ut': '/locations/my-place-hotel-hurricane-ut',
    '/locations/redding-ca': '/locations/my-place-hotel-redding-ca',
    '/locations/savannah-airport-poo': '/locations/savannah-airport-pooler-ga',
    '/locations/idaho/nampa-my-place-hotel-nampa-id': '/locations/my-place-hotel-nampa-ididaho-center',
    # ── Short-slug hotel paths (missing my-place-hotel- prefix) ──
    '/locations/st.joseph-mo': '/locations/my-place-hotel-st-joseph-mo',
    '/locations/sioux-falls-sd': '/locations/my-place-hotel-sioux-falls-sd',
    '/locations/green-bay-wi': '/locations/my-place-hotel-green-bay-wi',
    '/locations/my-place-st-george-ut': '/locations/my-place-hotel-st-george-ut',
    '/locations/my-place-idaho-falls': '/locations/my-place-hotel-idaho-falls-id',
    '/locations/my-place-east-molinequad-cities-il': '/locations/my-place-hotel-east-molinequad-cities-il',
    '/locations/my-hotel-jacksonville-nc': '/locations/my-place-hotel-jacksonville-nc',
    '/locations/myplace-hotel-colorado-springs-co': '/locations/my-place-hotel-colorado-springs-co',
    '/locations/spokane-valley': None,  # removed hotel
    '/locations/location-results&scrlybrkr=13e802df': '/locations/location-results',
    # ── Garbage / artifact paths → discard ──
    '/locations/bXktcGxhY2': None,
    '/locations/c2F2YW5uYW': None,
    '/locations/my-': None,
    '/locations/my-pl': None,
    '/locations/my=': None,
    '/locations/m': None,
    '/locations/my-place-hote': None,
    '/locations/null': None,
    '/locations/sava': None,
}

# ── DB fix: Jonesboro URL is wrong in Excel (points to Mesa AZ) ──
HOTEL_NAME_OVERRIDES = {
    '/locations/my-place-hotel-jonesboro-ar': 'My Place Hotel-Jonesboro, AR',
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

def fmt_r(n):
    if n >= 1e6:   return f'${n/1e6:.2f}M'
    elif n >= 1e3: return f'${n/1e3:.1f}K'
    else:          return f'${round(n):,}'

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
    hotel_id_map = {}       # Hotel ID (int) → canonical /locations/ path
    hotel_mgmt_map = {}     # canonical /locations/ path → Management Company (Col H)
    name_to_mgmt = {}       # Hotel name (str) → Management Company — for override fallback
    try:
        wb = openpyxl.load_workbook(DB_XLSX, read_only=True, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or len(row) < 13:
                continue
            hotel_id = row[0]    # Column A — internal Hotel ID
            name     = row[1]    # Column B — Hotel Name
            mgmt_co  = row[7]    # Column H — Management Company
            url      = row[12]   # Column M — Property URL (shifted from L→M in May 2026 layout after Pilot col added at I)
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
            mgmt_val = str(mgmt_co).strip() if mgmt_co is not None and str(mgmt_co).strip() else 'Unknown'
            hotel_mgmt_map[path] = mgmt_val
            name_to_mgmt[str(name).strip()] = mgmt_val
            # Map Hotel ID to canonical path for booking URL attribution
            if hotel_id is not None:
                try:
                    hotel_id_map[int(hotel_id)] = path
                except (ValueError, TypeError):
                    pass
        wb.close()
    except Exception as e:
        print(f'Warning: could not load hotel names: {e}')
    # Apply manual overrides (fixes DB errors like Jonesboro URL pointing to Mesa)
    for path, name in HOTEL_NAME_OVERRIDES.items():
        hotel_names[path] = name
        slug = path.split('/')[-1]
        hotel_slug_index[slug] = path
        # Mgmt company override: look up by the corrected hotel name
        if name in name_to_mgmt and path not in hotel_mgmt_map:
            hotel_mgmt_map[path] = name_to_mgmt[name]
    return hotel_names, hotel_slug_index, hotel_id_map, hotel_mgmt_map

def resolve_hotel_name(path, hotel_names, hotel_slug_index):
    if path in hotel_names:
        return hotel_names[path]
    slug = path.split('/')[-1]
    if slug in hotel_slug_index:
        return hotel_names[hotel_slug_index[slug]]
    parts = slug.replace('my-place-hotel-', '').split('-')
    if len(parts) >= 2:
        state = parts[-1].upper()
        city  = ' '.join(p.capitalize() for p in parts[:-1])
        return f'My Place Hotel-{city}, {state}'
    return slug

# ── Parse acquisition CSV ─────────────────────────────────────────────────
def parse_acq(rows):
    # Full-year 2025 monthly (all 12 months) — for trend chart
    mo_sess_25_full = {m: 0 for m in range(1, 13)}
    mo_ke_25_full   = {m: 0.0 for m in range(1, 13)}
    mo_rev_25_full  = {m: 0.0 for m in range(1, 13)}
    # 2026 monthly (available months)
    mo_sess_26 = defaultdict(float)
    mo_ke_26   = defaultdict(float)
    mo_rev_26  = defaultdict(float)
    # Same-period channel data (both years Jan 1–Apr 30)
    ch_sess_sp = defaultdict(lambda: {'25': 0.0, '26': 0.0})
    ch_ke_sp   = defaultdict(lambda: {'25': 0.0, '26': 0.0})
    ch_rev_sp  = defaultdict(lambda: {'25': 0.0, '26': 0.0})

    for row in rows:
        channel = row.get('Session Updated Channel Group', '').strip()
        date    = str(row.get('Date', '')).strip()
        if len(date) != 8:
            continue
        year  = int(date[:4])
        month = int(date[4:6])
        if year not in (2025, 2026):
            continue
        if channel.startswith(WIFI_CH):
            continue  # exclude WiFi from website dashboard

        try:
            sess = float(row.get('Sessions') or 0)
            ke   = float(row.get('Key events') or 0)
            rev  = float(row.get('Total revenue') or 0)
        except (ValueError, TypeError):
            continue

        if year == 2025 and month in range(1, 13):
            mo_sess_25_full[month] += sess
            mo_ke_25_full[month]   += ke
            mo_rev_25_full[month]  += rev
            # Same-period channel
            if date <= CUTOFF_25:
                ch_sess_sp[channel]['25'] += sess
                ch_ke_sp[channel]['25']   += ke
                ch_rev_sp[channel]['25']  += rev

        elif year == 2026 and date <= CUTOFF_26:
            mo_sess_26[month] += sess
            mo_ke_26[month]   += ke
            mo_rev_26[month]  += rev
            # Same-period channel
            if date <= CUTOFF_26:
                ch_sess_sp[channel]['26'] += sess
                ch_ke_sp[channel]['26']   += ke
                ch_rev_sp[channel]['26']  += rev

    return (mo_sess_25_full, mo_ke_25_full, mo_rev_25_full,
            mo_sess_26, mo_ke_26, mo_rev_26,
            ch_sess_sp, ch_ke_sp, ch_rev_sp)

# ── Parse landing page CSV ────────────────────────────────────────────────
def parse_lp(rows, hotel_names, hotel_slug_index, hotel_id_map):
    # Full-year 2025 page views (for trend chart)
    mo_pv_25_full = {m: 0 for m in range(1, 13)}
    # 2026 page views (available months)
    mo_pv_26 = defaultdict(int)
    # Same-period hotel data + site page data
    hotel = defaultdict(lambda: {'v25': 0, 'v26': 0, 'ke25': 0.0, 'ke26': 0.0, 'r25': 0.0, 'r26': 0.0})
    # Consolidated site pages: home, locations index, and individual non-/locations/ pages
    home_page       = {'v25': 0, 'v26': 0, 'ke25': 0.0, 'ke26': 0.0, 'r25': 0.0, 'r26': 0.0}
    locations_index = {'v25': 0, 'v26': 0, 'ke25': 0.0, 'ke26': 0.0, 'r25': 0.0, 'r26': 0.0}
    site_pages      = defaultdict(lambda: {'v25': 0, 'v26': 0, 'ke25': 0.0, 'ke26': 0.0, 'r25': 0.0, 'r26': 0.0})
    state_city      = defaultdict(lambda: {'v25': 0, 'v26': 0, 'ke25': 0.0, 'ke26': 0.0, 'r25': 0.0, 'r26': 0.0})
    booking_attributed = 0

    for row in rows:
        raw_path = str(row.get('Landing page + query string', '')).strip()
        date  = str(row.get('Date', '')).strip()
        if len(date) != 8:
            continue
        year  = int(date[:4])
        month = int(date[4:6])
        if year not in (2025, 2026):
            continue

        try:
            views = int(float(row.get('Views') or 0))
            ke    = float(row.get('Key events') or 0)
            rev   = float(row.get('Total revenue') or 0)
        except (ValueError, TypeError):
            continue

        ys = '25' if year == 2025 else '26'

        # ── Booking URL handling: attribute &hotel=XXXXX to the hotel ──
        hotel_id_match = re.search(r'[?&]hotel=(\d+)', raw_path)
        if hotel_id_match:
            hid = int(hotel_id_match.group(1))
            canon_path = hotel_id_map.get(hid)
            if canon_path and canon_path not in REMOVED_HOTELS:
                if year == 2025 and month in range(1, 13):
                    mo_pv_25_full[month] += views
                    if date <= CUTOFF_25:
                        hotel[canon_path][f'v{ys}'] += views
                        hotel[canon_path][f'ke{ys}'] += ke
                        hotel[canon_path][f'r{ys}'] += rev
                elif year == 2026 and date <= CUTOFF_26:
                    mo_pv_26[month] += views
                    hotel[canon_path][f'v{ys}'] += views
                    hotel[canon_path][f'ke{ys}'] += ke
                    hotel[canon_path][f'r{ys}'] += rev
                booking_attributed += 1
            continue

        # ── Standard path handling ──
        path = raw_path.split('?')[0]
        if is_artifact(path) or path == '(not set)':
            continue
        if path.startswith('/wifi/'):
            continue

        p = LP_PAGE_MERGES.get(path, path)
        if p is None or p in REMOVED_HOTELS:
            continue

        # Determine if this row falls in the reporting window
        in_window = False
        if year == 2025 and month in range(1, 13):
            mo_pv_25_full[month] += views
            if date <= CUTOFF_25:
                in_window = True
        elif year == 2026 and date <= CUTOFF_26:
            mo_pv_26[month] += views
            in_window = True

        if not in_window:
            continue

        # ── Classify into hotel vs site pages ──
        if p in hotel_names:
            hotel[p][f'v{ys}'] += views
            hotel[p][f'ke{ys}'] += ke
            hotel[p][f'r{ys}'] += rev
        elif p.startswith('/locations/') and _is_hotel_property_path(p):
            slug = p.split('/')[-1]
            if slug in hotel_slug_index:
                canon = hotel_slug_index[slug]
                hotel[canon][f'v{ys}'] += views
                hotel[canon][f'ke{ys}'] += ke
                hotel[canon][f'r{ys}'] += rev
        elif p == '/':
            home_page[f'v{ys}'] += views
            home_page[f'ke{ys}'] += ke
            home_page[f'r{ys}'] += rev
        elif p in ('/locations', '/locations/', '/locations/location-results'):
            locations_index[f'v{ys}'] += views
            locations_index[f'ke{ys}'] += ke
            locations_index[f'r{ys}'] += rev
        elif p.startswith('/locations/'):
            pass  # state/city browse pages — drop (minimal traffic, no attribution value)
        elif p.startswith('/state-city/'):
            state_city[p][f'v{ys}'] += views
            state_city[p][f'ke{ys}'] += ke
            state_city[p][f'r{ys}'] += rev
        elif p.strip():
            # Non-/locations/ site pages (e.g. /signin, /stayrewarded, /special-offers)
            site_pages[p][f'v{ys}'] += views
            site_pages[p][f'ke{ys}'] += ke
            site_pages[p][f'r{ys}'] += rev

    print(f'  Booking URLs attributed to hotels: {booking_attributed}')
    return mo_pv_25_full, mo_pv_26, hotel, home_page, locations_index, site_pages, state_city

def _is_hotel_property_path(p):
    """Check if a /locations/ path is a hotel property page (not a state/city hub)."""
    sub = p.replace('/locations/', '')
    if sub.startswith('my-place-hotel') or sub.startswith('savannah-airport'):
        return True
    return False

# ── Build JS data arrays ───────────────────────────────────────────────────
def build_js_arrays(mo_sess_25_full, mo_ke_25_full, mo_rev_25_full,
                    mo_sess_26, mo_ke_26, mo_rev_26,
                    mo_pv_25_full, mo_pv_26,
                    ch_sess_sp, ch_ke_sp, ch_rev_sp,
                    hotel, home_page, locations_index, site_pages, state_city,
                    hotel_names, hotel_slug_index, hotel_mgmt_map):

    # Full 2025 monthly arrays (12 elements)
    SESS25 = [round(mo_sess_25_full[m]) for m in range(1, 13)]
    PV25   = [round(mo_pv_25_full[m])   for m in range(1, 13)]
    REV25  = [round(mo_rev_25_full[m], 2) for m in range(1, 13)]
    PURCH25= [round(mo_ke_25_full[m])   for m in range(1, 13)]

    # 2026 YTD monthly arrays (available months only, sorted)
    months_26 = sorted(mo_sess_26.keys())
    SESS26  = [round(mo_sess_26[m]) for m in months_26]
    PV26    = [round(mo_pv_26[m])   for m in months_26]
    REV26   = [round(mo_rev_26[m], 2) for m in months_26]
    PURCH26 = [round(mo_ke_26[m])   for m in months_26]

    # Channel data (same-period)
    all_channels = set(ch_sess_sp.keys())
    CH_NAMES = sorted(all_channels, key=lambda c: ch_sess_sp[c]['26'], reverse=True)

    CH_SESS25 = {c: round(ch_sess_sp[c]['25']) for c in CH_NAMES}
    CH_SESS26 = {c: round(ch_sess_sp[c]['26']) for c in CH_NAMES}
    CH_REV25  = {c: round(ch_rev_sp[c]['25'], 2) for c in CH_NAMES}
    CH_REV26  = {c: round(ch_rev_sp[c]['26'], 2) for c in CH_NAMES}
    CH_CONV25 = {c: round(ch_ke_sp[c]['25'] / ch_sess_sp[c]['25'] * 100, 3)
                 if ch_sess_sp[c]['25'] else 0 for c in CH_NAMES}
    CH_CONV26 = {c: round(ch_ke_sp[c]['26'] / ch_sess_sp[c]['26'] * 100, 3)
                 if ch_sess_sp[c]['26'] else 0 for c in CH_NAMES}

    # Hotel data (same-period), sorted by combined revenue
    hotel_list = []
    for path, d in hotel.items():
        label = resolve_hotel_name(path, hotel_names, hotel_slug_index)
        cr25 = round(d['ke25'] / d['v25'] * 100, 2) if d['v25'] else 0
        cr26 = round(d['ke26'] / d['v26'] * 100, 2) if d['v26'] else 0
        # Resolve management company via canonical path; fall back via slug index; default Unknown
        mgmt_co = hotel_mgmt_map.get(path)
        if not mgmt_co:
            slug = path.split('/')[-1]
            canon = hotel_slug_index.get(slug)
            if canon:
                mgmt_co = hotel_mgmt_map.get(canon)
        if not mgmt_co:
            mgmt_co = 'Unknown'
        hotel_list.append({
            'label': label, 'path': path,
            'v25': round(d['v25']), 'v26': round(d['v26']),
            'ke25': round(d['ke25']), 'ke26': round(d['ke26']),
            'r25': round(d['r25'], 2), 'r26': round(d['r26'], 2),
            'cr25': cr25, 'cr26': cr26,
            'mgmt_co': mgmt_co,
        })
    hotel_list.sort(key=lambda x: x['r25'] + x['r26'], reverse=True)

    # Management Company dropdown options — sorted by hotel count desc, "All Companies" first
    mgmt_counts = defaultdict(int)
    for h in hotel_list:
        mgmt_counts[h['mgmt_co']] += 1
    mgmt_companies = [{'name': 'All Companies', 'count': len(hotel_list)}]
    for name, cnt in sorted(mgmt_counts.items(), key=lambda x: (-x[1], x[0])):
        mgmt_companies.append({'name': name, 'count': cnt})

    # Site pages: Home Page (consolidated), Locations Index (consolidated),
    # then individual non-/locations/ pages that have revenue
    hub_list = []
    # 1. Home Page
    hp = home_page
    hub_list.append({
        'label': 'Home Page', 'path': '/',
        'v25': round(hp['v25']), 'v26': round(hp['v26']),
        'ke25': round(hp['ke25']), 'ke26': round(hp['ke26']),
        'r25': round(hp['r25'], 2), 'r26': round(hp['r26'], 2),
    })
    # 2. Locations Index
    li = locations_index
    hub_list.append({
        'label': 'Locations Index', 'path': '/locations',
        'v25': round(li['v25']), 'v26': round(li['v26']),
        'ke25': round(li['ke25']), 'ke26': round(li['ke26']),
        'r25': round(li['r25'], 2), 'r26': round(li['r26'], 2),
    })
    # 3. Individual non-/locations/ pages with revenue, sorted by revenue desc
    for path, d in sorted(site_pages.items(),
                          key=lambda x: x[1]['r25'] + x[1]['r26'], reverse=True):
        if d['r25'] > 0 or d['r26'] > 0:
            hub_list.append({
                'label': path, 'path': path,
                'v25': round(d['v25']), 'v26': round(d['v26']),
                'ke25': round(d['ke25']), 'ke26': round(d['ke26']),
                'r25': round(d['r25'], 2), 'r26': round(d['r26'], 2),
            })

    # State/City landing pages — all pages with any page views, sorted by views desc
    sc_list = []
    for path, d in sorted(state_city.items(),
                          key=lambda x: x[1]['v25'] + x[1]['v26'], reverse=True):
        if d['v25'] > 0 or d['v26'] > 0:
            sc_list.append({
                'label': path, 'path': path,
                'v25': round(d['v25']), 'v26': round(d['v26']),
                'ke25': round(d['ke25']), 'ke26': round(d['ke26']),
                'r25': round(d['r25'], 2), 'r26': round(d['r26'], 2),
            })

    return dict(
        SESS25=SESS25, SESS26=SESS26,
        PV25=PV25, PV26=PV26,
        REV25=REV25, REV26=REV26,
        PURCH25=PURCH25, PURCH26=PURCH26,
        CH_NAMES=CH_NAMES,
        CH_SESS25=CH_SESS25, CH_SESS26=CH_SESS26,
        CH_REV25=CH_REV25, CH_REV26=CH_REV26,
        CH_CONV25=CH_CONV25, CH_CONV26=CH_CONV26,
        HOTEL_DATA=hotel_list,
        HUB_DATA=hub_list,
        SC_DATA=sc_list,
        MGMT_COMPANIES_GA4=mgmt_companies,
    )

# ── KPI calculations (same-period) ────────────────────────────────────────
def calc_kpis(d):
    # Same-period sums for 2025: sum of channel data (already same-period filtered)
    s25_sp = sum(d['CH_SESS25'].values())
    s26_sp = sum(d['CH_SESS26'].values())
    r25_sp = sum(d['CH_REV25'].values())
    r26_sp = sum(d['CH_REV26'].values())
    # Key events same-period: sum from hotel data
    ke25_sp = sum(x['ke25'] for x in d['HOTEL_DATA'])
    ke25_sp += sum(x['ke25'] for x in d['HUB_DATA'])
    ke26_sp = sum(x['ke26'] for x in d['HOTEL_DATA'])
    ke26_sp += sum(x['ke26'] for x in d['HUB_DATA'])
    conv25 = ke25_sp / s25_sp * 100 if s25_sp else 0
    conv26 = ke26_sp / s26_sp * 100 if s26_sp else 0
    rpr25  = r25_sp / ke25_sp if ke25_sp else 0
    rpr26  = r26_sp / ke26_sp if ke26_sp else 0
    return dict(
        s25=s25_sp, s26=s26_sp,
        r25=r25_sp, r26=r26_sp,
        ke25=ke25_sp, ke26=ke26_sp,
        conv25=conv25, conv26=conv26,
        rpr25=rpr25, rpr26=rpr26,
    )

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
.kpi-row { display:grid; grid-template-columns:repeat(5,1fr); gap:14px; margin-bottom:24px; }
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
.grid-2 { display:grid; grid-template-columns:1fr 1fr; gap:16px; margin-bottom:20px; }
.chart-card { background:#fff; border:1px solid var(--border); border-radius:10px;
              padding:18px 20px; box-shadow:0 1px 4px rgba(28,53,94,0.06); margin-bottom:0; }
.chart-card h3 { font-size:11px; font-weight:700; color:var(--muted); text-transform:uppercase;
                 letter-spacing:0.6px; border-bottom:2px solid #FF5F00;
                 padding-bottom:7px; margin-bottom:5px; }
.chart-desc { font-size:10px; color:#8ba0bf; font-style:italic; margin-bottom:10px; line-height:1.5; }
.legend-row { display:flex; gap:14px; margin-bottom:8px; }
.legend-item { display:flex; align-items:center; gap:5px; font-size:10px; color:#5A6B84; }
.legend-swatch { width:10px; height:10px; border-radius:2px; }
.ch-bar-section { margin-bottom:10px; }
.ch-bar-label { font-size:11px; font-weight:600; color:#1C355E; margin-bottom:4px;
                display:flex; align-items:center; gap:6px; }
.bar-row { display:flex; align-items:center; gap:7px; margin-bottom:2px; }
.bar-yr  { font-size:9px; color:#8ba0bf; width:26px; flex-shrink:0; text-align:right; }
.bar-track { flex:1; height:16px; background:#EEF1F6; border-radius:3px; overflow:hidden; }
.bar-fill  { height:100%; border-radius:3px; transition:width 0.7s ease; }
.bar-val { font-size:10px; font-weight:600; color:#1C355E; width:66px; flex-shrink:0; text-align:right; }
.delta-pill { font-size:9px; font-weight:700; padding:1px 6px; border-radius:7px; flex-shrink:0; }
.up   { background:rgba(46,139,87,0.12);  color:#2E8B57; }
.down { background:rgba(204,51,51,0.1);   color:#CC3333; }
.new  { background:rgba(255,95,0,0.1);    color:#FF5F00; }
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
tbody td.num, tfoot td.num { text-align:right; font-variant-numeric:tabular-nums; white-space:nowrap; }
tbody td.name-cell { font-weight:600; }
tbody td.path-cell { font-family:monospace; font-size:10px; color:#5A6B84;
                     max-width:240px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
tbody tr:nth-child(even) { background:#F5F8FC; }
tbody tr:hover { background:#E8F0FA; }
tbody tr:last-child td { border-bottom:none; }
.delta-cell { font-size:10px; font-weight:700; white-space:nowrap; }
.delta-cell.up { color:#2E8B57; }
.delta-cell.down { color:#CC3333; }
.hotel-table { table-layout:fixed; width:100%; }
.hotel-table thead th:nth-child(1) { width:17%; }
.hotel-table thead th:nth-child(2) { width:9%; }
.hotel-table thead th { width:6.16%; }
.mgmt-tag { display:inline-block; background:#EEF1F6; color:#1C355E;
            font-size:10px; font-weight:600; padding:2px 8px; border-radius:10px;
            white-space:nowrap; max-width:100%; overflow:hidden; text-overflow:ellipsis; }
.mgmt-filter-row { display:flex; align-items:center; justify-content:space-between;
                   gap:14px; margin:6px 0 12px; flex-wrap:wrap; }
.mgmt-filter-left { display:flex; flex-direction:column; gap:4px; }
.mgmt-filter-label { display:flex; align-items:center; gap:8px;
                     font-size:11px; color:#5A6B84; font-weight:600; }
.mgmt-filter-label select {
  font-family:inherit; font-size:11px; color:#1C355E; background:#fff;
  border:1px solid #D0D8E8; border-radius:6px; padding:4px 26px 4px 10px;
  appearance:none; -webkit-appearance:none; cursor:pointer; font-weight:600;
  background-image:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='10' height='6' viewBox='0 0 10 6'><path fill='%231C355E' d='M0 0l5 6 5-6z'/></svg>");
  background-repeat:no-repeat; background-position:right 9px center;
  transition:border-color 0.15s, box-shadow 0.15s;
}
.mgmt-filter-label select:focus { outline:none; border-color:#1C355E;
                                  box-shadow:0 0 0 2px rgba(28,53,94,0.15); }
.mgmt-filter-meta { font-size:10px; color:#8ba0bf; font-style:italic;
                    margin-left:54px; min-height:14px; }
.mgmt-filter-helper { font-size:11px; color:#8ba0bf; font-style:italic; }
.footer { background:#1C355E; border-top:3px solid #FF5F00; padding:14px 40px;
          text-align:center; color:rgba(255,255,255,0.55); font-size:10px; }
"""

JS = """
const C25='#4A7FC1', C26='#FF5F00';

function fmtN(n){ return n>=1e6?(n/1e6).toFixed(2)+'M':n>=1e3?(n/1e3).toFixed(1)+'K':Math.round(n).toLocaleString(); }
function fmtR(n){ return n>=1e6?'$'+(n/1e6).toFixed(2)+'M':n>=1e3?'$'+(n/1e3).toFixed(1)+'K':'$'+Math.round(n).toLocaleString(); }

// Trend chart — 2025 full year (12 pts) + 2026 YTD overlay (starts at x=0)
function drawTrendLine(canvasId, arr25, arr26, fmtFn) {
  const canvas = document.getElementById(canvasId); if (!canvas) return;
  const dpr = window.devicePixelRatio || 1;
  const W = canvas.parentElement.clientWidth || 900;
  const H = parseInt(canvas.getAttribute('height')) || 150;
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

function renderHorizBars(id, channels, d25, d26, fmtFn) {
  const el=document.getElementById(id); if(!el) return;
  const maxV=Math.max(...channels.map(c=>Math.max(d25[c]||0,d26[c]||0)));
  el.innerHTML=channels.map(ch=>{
    const v25=d25[ch]||0, v26=d26[ch]||0;
    const w25=maxV?(v25/maxV*100).toFixed(1):0, w26=maxV?(v26/maxV*100).toFixed(1):0;
    const delta=v25>0?((v26-v25)/v25*100):999;
    const dcls=delta>=0?'up':'down';
    const dTxt=v25===0?'<span class="delta-pill new">New</span>'
      :`<span class="delta-pill ${dcls}">${delta>=0?'▲':'▼'} ${Math.abs(delta).toFixed(0)}%</span>`;
    return `<div class="ch-bar-section">
      <div class="ch-bar-label">${ch} ${dTxt}</div>
      <div class="bar-row"><div class="bar-yr">'25</div>
        <div class="bar-track"><div class="bar-fill" style="width:${w25}%;background:${C25};opacity:0.75;"></div></div>
        <div class="bar-val">${fmtFn(v25)}</div></div>
      <div class="bar-row"><div class="bar-yr">'26</div>
        <div class="bar-track"><div class="bar-fill" style="width:${w26}%;background:${C26};"></div></div>
        <div class="bar-val">${fmtFn(v26)}</div></div>
    </div>`;
  }).join('');
}

function renderConvBars(id, channels, c25, c26) {
  const el=document.getElementById(id); if(!el) return;
  const rel=channels.filter(c=>(c25[c]||0)+(c26[c]||0)>0);
  const maxR=Math.max(...rel.map(c=>Math.max(c25[c]||0,c26[c]||0)));
  const half=Math.ceil(rel.length/2), left=rel.slice(0,half), right=rel.slice(half);
  const barH=ch=>ch.map(c=>{
    const v25=c25[c]||0,v26=c26[c]||0;
    const w25=maxR?(v25/maxR*100).toFixed(1):0, w26=maxR?(v26/maxR*100).toFixed(1):0;
    const delta=v25>0?(v26-v25):0, dcls=delta>=0?'up':'down';
    const dTxt=v25>0?`<span class="delta-pill ${dcls}">${delta>=0?'▲':'▼'} ${Math.abs(delta).toFixed(2)}pp</span>`
      :'<span class="delta-pill new">New</span>';
    return `<div class="ch-bar-section">
      <div class="ch-bar-label">${c} ${dTxt}</div>
      <div class="bar-row"><div class="bar-yr">'25</div>
        <div class="bar-track"><div class="bar-fill" style="width:${w25}%;background:${C25};opacity:0.75;"></div></div>
        <div class="bar-val">${v25.toFixed(2)}%</div></div>
      <div class="bar-row"><div class="bar-yr">'26</div>
        <div class="bar-track"><div class="bar-fill" style="width:${w26}%;background:${C26};"></div></div>
        <div class="bar-val">${v26.toFixed(2)}%</div></div>
    </div>`;
  }).join('');
  el.innerHTML=`<div style="display:grid;grid-template-columns:1fr 1fr;gap:20px;">
    <div>${barH(left)}</div><div>${barH(right)}</div></div>`;
}

let sortState={};
let lastSort={};  // per-table-id: {col, numeric, asc} \u2014 to re-apply after re-render
function applySort(id, col, numeric, asc) {
  const tbl=document.getElementById(id);
  const tbody=tbl.querySelector('tbody');
  const rows=Array.from(tbody.querySelectorAll('tr'));
  // Skip if table is showing a single placeholder row (e.g. "No properties match")
  if (rows.length <= 1 || !rows[0].cells || rows[0].cells.length <= col) return;
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
function sortTbl(id, col, numeric) {
  const key=id+'_'+col, asc=sortState[key]!=='asc';
  sortState[key]=asc?'asc':'desc';
  lastSort[id]={col, numeric, asc};
  applySort(id, col, numeric, asc);
}
function reapplySort(id) {
  const s=lastSort[id]; if (!s) return;
  applySort(id, s.col, s.numeric, s.asc);
}

function deltaCell(vA, vB, fmtFn) {
  const d=vB-vA;
  const cls=d>=0?'up':'down';
  return `<td class="num delta-cell ${cls}">${d>=0?'+':''}${fmtFn(Math.abs(d))}</td>`;
}
"""

# ── Build HTML ─────────────────────────────────────────────────────────────
def build_html(d, k, logo):
    n_hotels = len(d['HOTEL_DATA'])

    data_js = f"""
const SESS25={json.dumps(d['SESS25'])}, SESS26={json.dumps(d['SESS26'])};
const PV25={json.dumps(d['PV25'])},  PV26={json.dumps(d['PV26'])};
const REV25={json.dumps(d['REV25'])},   REV26={json.dumps(d['REV26'])};
const PURCH25={json.dumps(d['PURCH25'])},PURCH26={json.dumps(d['PURCH26'])};
const CH_NAMES={json.dumps(d['CH_NAMES'])};
const CH_SESS25={json.dumps(d['CH_SESS25'])}, CH_SESS26={json.dumps(d['CH_SESS26'])};
const CH_REV25={json.dumps(d['CH_REV25'])},   CH_REV26={json.dumps(d['CH_REV26'])};
const CH_CONV25={json.dumps(d['CH_CONV25'])}, CH_CONV26={json.dumps(d['CH_CONV26'])};
const HUB_DATA={json.dumps(d['HUB_DATA'])};
const HOTEL_DATA={json.dumps(d['HOTEL_DATA'])};
const SC_DATA={json.dumps(d['SC_DATA'])};
const MGMT_COMPANIES_GA4={json.dumps(d['MGMT_COMPANIES_GA4'])};
"""

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>My Place Hotels — Website Analytics 2026 YTD vs 2025</title>
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
        <h1>Website Analytics Dashboard — 2026 YTD</h1>
        <div class="header-subtitle">Google Analytics 4 &nbsp;·&nbsp; Jan 1 – Apr 30, 2026 vs same period 2025 &nbsp;·&nbsp; myplacehotels.com &nbsp;·&nbsp; Property WiFi excluded</div>
      </div>
    </div>
    <div class="header-meta">
      <div style="display:flex;gap:8px;">
        <a href="ga4-website-dashboard.html" class="portal-btn">2024 vs 2025 Dashboard</a>
        <a href="index.html" class="portal-btn">← Portal</a>
      </div>
      <div class="updated-pill">Jan 1 – Apr 30, 2026</div>
    </div>
  </div>
</div>

<div class="main">

  <div class="info-bar">
    📅 <span><strong>YoY comparison is same-period only</strong> — 2026 YTD (Jan 1–Apr 30, 2026) is compared against the identical date range in 2025 (Jan 1–Apr 30, 2025) for a clean apples-to-apples view. Monthly trend charts show all 12 months of 2025 for broader context.</span>
  </div>

  <!-- KPI Cards -->
  <div class="kpi-row">
    <div class="kpi-card">
      <div class="kpi-label">2026 YTD Sessions</div>
      <div class="kpi-value">{fmt_n(k['s26'])}</div>
      <div class="kpi-sub">2025 (same period): {fmt_n(k['s25'])}</div>
      <div class="kpi-yoy">{yoy_badge(k['s26'], k['s25'])} vs 2025</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">2026 YTD Revenue</div>
      <div class="kpi-value">{fmt_r(k['r26'])}</div>
      <div class="kpi-sub">2025 (same period): {fmt_r(k['r25'])}</div>
      <div class="kpi-yoy">{yoy_badge(k['r26'], k['r25'])} vs 2025</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">2026 YTD Reservations</div>
      <div class="kpi-value">{fmt_n(k['ke26'])}</div>
      <div class="kpi-sub">2025 (same period): {fmt_n(k['ke25'])}</div>
      <div class="kpi-yoy">{yoy_badge(k['ke26'], k['ke25'])} vs 2025</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">2026 Reservation Rate</div>
      <div class="kpi-value">{k['conv26']:.2f}%</div>
      <div class="kpi-sub">2025 (same period): {k['conv25']:.2f}%</div>
      <div class="kpi-yoy">{yoy_badge(k['conv26'], k['conv25'])} vs 2025</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">Revenue per Reservation</div>
      <div class="kpi-value">{fmt_r(k['rpr26'])}</div>
      <div class="kpi-sub">2025 (same period): {fmt_r(k['rpr25'])}</div>
      <div class="kpi-yoy">{yoy_badge(k['rpr26'], k['rpr25'])} vs 2025</div>
    </div>
  </div>

  <!-- Monthly Trends -->
  <div class="section-hdr">
    <h2>Monthly Trends</h2>
    <span class="desc">Full 2025 (blue) + 2026 YTD (orange) — excludes Property WiFi</span>
    <div class="section-divider"></div>
  </div>

  <div class="grid-2" style="margin-bottom:16px;">
    <div class="chart-card">
      <h3>Monthly Sessions</h3>
      <p class="chart-desc">All 12 months of 2025 shown for context, with 2026 YTD (Jan–Mar) overlaid. The 2026 line begins at January.</p>
      <div class="legend-row">
        <div class="legend-item"><div class="legend-swatch" style="background:var(--C25);"></div>2025 (full year)</div>
        <div class="legend-item"><div class="legend-swatch" style="background:var(--C26);"></div>2026 YTD</div>
      </div>
      <canvas id="cSessions" style="width:100%;display:block;" height="150"></canvas>
    </div>
    <div class="chart-card">
      <h3>Monthly Page Views</h3>
      <p class="chart-desc">Total page views across all website pages per month. Full 2025 shown for context with 2026 YTD overlaid. Excludes Property WiFi (/wifi/) pages.</p>
      <div class="legend-row">
        <div class="legend-item"><div class="legend-swatch" style="background:var(--C25);"></div>2025 (full year)</div>
        <div class="legend-item"><div class="legend-swatch" style="background:var(--C26);"></div>2026 YTD</div>
      </div>
      <canvas id="cPageViews" style="width:100%;display:block;" height="150"></canvas>
    </div>
  </div>

  <div class="grid-2" style="margin-bottom:16px;">
    <div class="chart-card">
      <h3>Monthly Website Revenue</h3>
      <p class="chart-desc">GA4-attributed booking revenue per month. 2026 YTD vs full 2025 trend.</p>
      <div class="legend-row">
        <div class="legend-item"><div class="legend-swatch" style="background:var(--C25);"></div>2025</div>
        <div class="legend-item"><div class="legend-swatch" style="background:var(--C26);"></div>2026 YTD</div>
      </div>
      <canvas id="cRevenue" style="width:100%;display:block;" height="150"></canvas>
    </div>
    <div class="chart-card">
      <h3>Monthly Reservations</h3>
      <p class="chart-desc">Purchase (reservation) events per month — primary booking conversion metric.</p>
      <div class="legend-row">
        <div class="legend-item"><div class="legend-swatch" style="background:var(--C25);"></div>2025</div>
        <div class="legend-item"><div class="legend-swatch" style="background:var(--C26);"></div>2026 YTD</div>
      </div>
      <canvas id="cPurchases" style="width:100%;display:block;" height="150"></canvas>
    </div>
  </div>

  <!-- Channel Performance -->
  <div class="section-hdr">
    <h2>Channel Performance</h2>
    <span class="desc">2025 same period (Jan 1–Apr 30) vs 2026 YTD (Jan 1–Apr 30)</span>
    <div class="section-divider"></div>
  </div>
  <div class="grid-2" style="margin-bottom:16px;">
    <div class="chart-card">
      <h3>Sessions by Channel</h3>
      <p class="chart-desc">Session volume per channel for the same Jan 1–Apr 30 window in both years.</p>
      <div id="chSessionBars"></div>
    </div>
    <div class="chart-card">
      <h3>Website Revenue by Channel</h3>
      <p class="chart-desc">GA4-attributed booking revenue by channel — same period comparison.</p>
      <div id="chRevBars"></div>
    </div>
  </div>

  <div class="section-hdr">
    <h2>Reservation Rate by Channel</h2>
    <span class="desc">Reservations ÷ Sessions — same period Jan 1–Apr 30</span>
    <div class="section-divider"></div>
  </div>
  <div class="chart-card" style="margin-bottom:16px;">
    <h3>Session Reservation Rate by Channel</h3>
    <p class="chart-desc">Percentage of sessions resulting in a reservation. Same-period comparison to isolate true YoY channel intent.</p>
    <div id="chConvBars"></div>
  </div>

  <!-- Site Pages with Revenue -->
  <div class="section-hdr">
    <h2>Site Page Performance</h2>
    <span class="desc">Home page, locations index, and other site pages with attributed revenue — same period Jan 1–Apr 30</span>
    <div class="section-divider"></div>
  </div>
  <div class="table-card">
    <h3>Revenue-Generating Site Pages — 2025 vs 2026 YTD</h3>
    <p class="chart-desc">Top 25 revenue-generating pages. Includes consolidated home page, locations index, and individual site pages. Totals reflect all pages.</p>
    <div style="overflow-x:auto;"><table id="hubTable">
      <thead><tr>
        <th onclick="sortTbl('hubTable',0,false)">Page</th>
        <th class="num" onclick="sortTbl('hubTable',1,true)">Views '25</th>
        <th class="num" onclick="sortTbl('hubTable',2,true)">Views '26</th>
        <th class="num" onclick="sortTbl('hubTable',3,true)">Δ Views</th>
        <th class="num" onclick="sortTbl('hubTable',4,true)">Res '25</th>
        <th class="num" onclick="sortTbl('hubTable',5,true)">Res '26</th>
        <th class="num" onclick="sortTbl('hubTable',6,true)">Revenue '25</th>
        <th class="num" onclick="sortTbl('hubTable',7,true)">Revenue '26</th>
        <th class="num" onclick="sortTbl('hubTable',8,true)">Δ Revenue</th>
      </tr></thead>
      <tbody id="hubBody"></tbody>
      <tfoot id="hubFoot"></tfoot>
    </table></div>
  </div>

  <!-- State/City Landing Pages -->
  <div class="section-hdr">
    <h2>State &amp; City Landing Pages</h2>
    <span class="desc">All /state-city/ landing pages with traffic — same period Jan 1–Apr 30 · click any column to re-sort</span>
    <div class="section-divider"></div>
  </div>
  <div class="table-card">
    <h3>State/City Pages — 2025 vs 2026 YTD</h3>
    <p class="chart-desc">Performance of state and city browse pages. Includes all pages with page views in either year.</p>
    <div style="overflow-x:auto;"><table id="scTable">
      <thead><tr>
        <th onclick="sortTbl('scTable',0,false)">Page</th>
        <th class="num" onclick="sortTbl('scTable',1,true)">Views '25</th>
        <th class="num" onclick="sortTbl('scTable',2,true)">Views '26</th>
        <th class="num" onclick="sortTbl('scTable',3,true)">Δ Views</th>
        <th class="num" onclick="sortTbl('scTable',4,true)">Res '25</th>
        <th class="num" onclick="sortTbl('scTable',5,true)">Res '26</th>
        <th class="num" onclick="sortTbl('scTable',6,true)">Revenue '25</th>
        <th class="num" onclick="sortTbl('scTable',7,true)">Revenue '26</th>
        <th class="num" onclick="sortTbl('scTable',8,true)">Δ Revenue</th>
      </tr></thead>
      <tbody id="scBody"></tbody>
      <tfoot id="scFoot"></tfoot>
    </table></div>
  </div>

  <!-- Hotel Property Pages -->
  <div class="section-hdr">
    <h2>Property Page Performance — {n_hotels} Hotels</h2>
    <span class="desc">Same period Jan 1–Apr 30 · sorted by total revenue</span>
    <div class="section-divider"></div>
  </div>
  <div class="table-card">
    <h3>Hotel Property Pages — 2026 YTD with YoY Variance</h3>
    <p class="chart-desc">Page views, reservations, and revenue per property. Both years reflect Jan 1–Apr 30 only.</p>
    <div class="mgmt-filter-row">
      <div class="mgmt-filter-left">
        <label class="mgmt-filter-label">Filter:
          <select id="mgmtFilter"></select>
        </label>
        <div class="mgmt-filter-meta" id="mgmtFilterMeta"></div>
      </div>
      <span class="mgmt-filter-helper">click any column to re-sort</span>
    </div>
    <table id="hotelTable" class="hotel-table">
      <thead><tr>
        <th onclick="sortTbl('hotelTable',0,false)">Property</th>
        <th onclick="sortTbl('hotelTable',1,false)">Mgmt Co</th>
        <th class="num" onclick="sortTbl('hotelTable',2,true)">Views '25</th>
        <th class="num" onclick="sortTbl('hotelTable',3,true)">Views '26</th>
        <th class="num" onclick="sortTbl('hotelTable',4,true)">Δ Views</th>
        <th class="num" onclick="sortTbl('hotelTable',5,true)">Res '25</th>
        <th class="num" onclick="sortTbl('hotelTable',6,true)">Res '26</th>
        <th class="num" onclick="sortTbl('hotelTable',7,true)">Δ Res</th>
        <th class="num" onclick="sortTbl('hotelTable',8,true)">CR '25</th>
        <th class="num" onclick="sortTbl('hotelTable',9,true)">CR '26</th>
        <th class="num" onclick="sortTbl('hotelTable',10,true)">Δ CR</th>
        <th class="num" onclick="sortTbl('hotelTable',11,true)">Revenue '25</th>
        <th class="num" onclick="sortTbl('hotelTable',12,true)">Revenue '26</th>
        <th class="num" onclick="sortTbl('hotelTable',13,true)">Δ Revenue</th>
      </tr></thead>
      <tbody id="hotelBody"></tbody>
      <tfoot id="hotelFoot"></tfoot>
    </table>
  </div>


</div>

<div class="footer">
  My Place Hotels of America &nbsp;·&nbsp; Website Analytics 2026 YTD &nbsp;·&nbsp; Google Analytics 4 &nbsp;·&nbsp; Confidential — Internal Use Only
</div>

<script>
{JS}
{data_js}

function renderHubTable() {{
  const top25 = HUB_DATA.slice(0, 25);
  document.getElementById('hubBody').innerHTML = top25.map(r=>{{
    return `<tr>
      <td class="path-cell">${{r.label}}</td>
      <td class="num">${{r.v25.toLocaleString()}}</td>
      <td class="num">${{r.v26.toLocaleString()}}</td>
      ${{deltaCell(r.v25,r.v26,n=>Math.round(n).toLocaleString())}}
      <td class="num">${{r.ke25.toLocaleString()}}</td>
      <td class="num">${{r.ke26.toLocaleString()}}</td>
      <td class="num">${{fmtR(r.r25)}}</td>
      <td class="num">${{fmtR(r.r26)}}</td>
      ${{deltaCell(r.r25,r.r26,fmtR)}}
    </tr>`;
  }}).join('');
  // Totals from ALL hub pages (not just top 25)
  const tv25=HUB_DATA.reduce((s,r)=>s+r.v25,0), tv26=HUB_DATA.reduce((s,r)=>s+r.v26,0);
  const tk25=HUB_DATA.reduce((s,r)=>s+r.ke25,0), tk26=HUB_DATA.reduce((s,r)=>s+r.ke26,0);
  const tr25=HUB_DATA.reduce((s,r)=>s+r.r25,0), tr26=HUB_DATA.reduce((s,r)=>s+r.r26,0);
  document.getElementById('hubFoot').innerHTML = `<tr style="font-weight:700;border-top:2px solid #1C355E">
    <td class="path-cell">All Site Pages Total (${{HUB_DATA.length}} pages)</td>
    <td class="num">${{tv25.toLocaleString()}}</td>
    <td class="num">${{tv26.toLocaleString()}}</td>
    ${{deltaCell(tv25,tv26,n=>Math.round(n).toLocaleString())}}
    <td class="num">${{tk25.toLocaleString()}}</td>
    <td class="num">${{tk26.toLocaleString()}}</td>
    <td class="num">${{fmtR(tr25)}}</td>
    <td class="num">${{fmtR(tr26)}}</td>
    ${{deltaCell(tr25,tr26,fmtR)}}
  </tr>`;
}}

function renderStateCityTable() {{
  document.getElementById('scBody').innerHTML = SC_DATA.map(r=>{{
    return `<tr>
      <td class="path-cell">${{r.label}}</td>
      <td class="num">${{r.v25.toLocaleString()}}</td>
      <td class="num">${{r.v26.toLocaleString()}}</td>
      ${{deltaCell(r.v25,r.v26,n=>Math.round(n).toLocaleString())}}
      <td class="num">${{r.ke25.toLocaleString()}}</td>
      <td class="num">${{r.ke26.toLocaleString()}}</td>
      <td class="num">${{fmtR(r.r25)}}</td>
      <td class="num">${{fmtR(r.r26)}}</td>
      ${{deltaCell(r.r25,r.r26,fmtR)}}
    </tr>`;
  }}).join('');
  const tv25=SC_DATA.reduce((s,r)=>s+r.v25,0), tv26=SC_DATA.reduce((s,r)=>s+r.v26,0);
  const tk25=SC_DATA.reduce((s,r)=>s+r.ke25,0), tk26=SC_DATA.reduce((s,r)=>s+r.ke26,0);
  const tr25=SC_DATA.reduce((s,r)=>s+r.r25,0), tr26=SC_DATA.reduce((s,r)=>s+r.r26,0);
  document.getElementById('scFoot').innerHTML = `<tr style="font-weight:700;border-top:2px solid #1C355E">
    <td class="path-cell">All State/City Pages (${{SC_DATA.length}} pages)</td>
    <td class="num">${{tv25.toLocaleString()}}</td>
    <td class="num">${{tv26.toLocaleString()}}</td>
    ${{deltaCell(tv25,tv26,n=>Math.round(n).toLocaleString())}}
    <td class="num">${{tk25.toLocaleString()}}</td>
    <td class="num">${{tk26.toLocaleString()}}</td>
    <td class="num">${{fmtR(tr25)}}</td>
    <td class="num">${{fmtR(tr26)}}</td>
    ${{deltaCell(tr25,tr26,fmtR)}}
  </tr>`;
}}

function fmtCR(v) {{ return v.toFixed(1)+'%'; }}
function deltaCR(a,b) {{
  const diff = b - a;
  if(Math.abs(diff)<0.05) return '<td class="num delta-cell">—</td>';
  const cls = diff>0?'up':'down';
  const arrow = diff>0?'▲':'▼';
  return `<td class="num delta-cell ${{cls}}">${{arrow}} ${{Math.abs(diff).toFixed(1)}}pp</td>`;
}}
// Canonical immutable hotel list; refilter from this every change
const HOTEL_DATA_ALL = HOTEL_DATA.slice();
// Total YTD '26 revenue across all properties — used for % share in subtitle
const TOTAL_R26_ALL = HOTEL_DATA_ALL.reduce((s,r)=>s+r.r26,0);
let currentMgmtFilter = 'All Companies';

function escapeHtml(s) {{
  return String(s).replace(/[&<>"']/g, c=>({{
    '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'
  }}[c]));
}}

function buildMgmtFilterOptions() {{
  const sel = document.getElementById('mgmtFilter');
  if (!sel) return;
  sel.innerHTML = MGMT_COMPANIES_GA4.map(o=>{{
    return `<option value="${{escapeHtml(o.name)}}">${{escapeHtml(o.name)}} (${{o.count}})</option>`;
  }}).join('');
  sel.value = currentMgmtFilter;
  sel.addEventListener('change', (e)=>{{
    // Preserve scroll position when filter changes
    const y = window.scrollY;
    currentMgmtFilter = e.target.value;
    renderHotelTable();
    window.scrollTo({{top:y, left:0, behavior:'instant'}});
  }});
}}

function renderHotelTable() {{
  const filtered = currentMgmtFilter === 'All Companies'
    ? HOTEL_DATA_ALL
    : HOTEL_DATA_ALL.filter(r=>r.mgmt_co === currentMgmtFilter);

  const body = document.getElementById('hotelBody');
  if (filtered.length === 0) {{
    body.innerHTML = `<tr><td colspan="14" style="text-align:center;font-style:italic;color:#8ba0bf;padding:18px;">No properties match this filter</td></tr>`;
  }} else {{
    body.innerHTML = filtered.map(r=>{{
      return `<tr>
        <td class="name-cell">${{r.label}}</td>
        <td><span class="mgmt-tag">${{escapeHtml(r.mgmt_co || 'Unknown')}}</span></td>
        <td class="num">${{r.v25.toLocaleString()}}</td>
        <td class="num">${{r.v26.toLocaleString()}}</td>
        ${{deltaCell(r.v25,r.v26,n=>Math.round(n).toLocaleString())}}
        <td class="num">${{r.ke25.toLocaleString()}}</td>
        <td class="num">${{r.ke26.toLocaleString()}}</td>
        ${{deltaCell(r.ke25,r.ke26,n=>Math.round(n).toLocaleString())}}
        <td class="num">${{fmtCR(r.cr25)}}</td>
        <td class="num">${{fmtCR(r.cr26)}}</td>
        ${{deltaCR(r.cr25,r.cr26)}}
        <td class="num">${{fmtR(r.r25)}}</td>
        <td class="num">${{fmtR(r.r26)}}</td>
        ${{deltaCell(r.r25,r.r26,fmtR)}}
      </tr>`;
    }}).join('');
  }}

  // Totals row — based on filtered set
  const tv25=filtered.reduce((s,r)=>s+r.v25,0), tv26=filtered.reduce((s,r)=>s+r.v26,0);
  const tk25=filtered.reduce((s,r)=>s+r.ke25,0), tk26=filtered.reduce((s,r)=>s+r.ke26,0);
  const tr25=filtered.reduce((s,r)=>s+r.r25,0), tr26=filtered.reduce((s,r)=>s+r.r26,0);
  const tcr25=tv25?tk25/tv25*100:0, tcr26=tv26?tk26/tv26*100:0;
  const totalLabel = currentMgmtFilter === 'All Companies'
    ? 'All Properties Total'
    : `${{currentMgmtFilter}} Total (${{filtered.length}})`;
  document.getElementById('hotelFoot').innerHTML = filtered.length === 0 ? '' : `<tr style="font-weight:700;border-top:2px solid #1C355E">
    <td class="name-cell">${{escapeHtml(totalLabel)}}</td>
    <td></td>
    <td class="num">${{tv25.toLocaleString()}}</td>
    <td class="num">${{tv26.toLocaleString()}}</td>
    ${{deltaCell(tv25,tv26,n=>Math.round(n).toLocaleString())}}
    <td class="num">${{tk25.toLocaleString()}}</td>
    <td class="num">${{tk26.toLocaleString()}}</td>
    ${{deltaCell(tk25,tk26,n=>Math.round(n).toLocaleString())}}
    <td class="num">${{fmtCR(tcr25)}}</td>
    <td class="num">${{fmtCR(tcr26)}}</td>
    ${{deltaCR(tcr25,tcr26)}}
    <td class="num">${{fmtR(tr25)}}</td>
    <td class="num">${{fmtR(tr26)}}</td>
    ${{deltaCell(tr25,tr26,fmtR)}}
  </tr>`;

  // Subtitle below dropdown
  const meta = document.getElementById('mgmtFilterMeta');
  if (meta) {{
    if (currentMgmtFilter === 'All Companies') {{
      meta.textContent = '';
    }} else {{
      const pct = TOTAL_R26_ALL > 0 ? (tr26 / TOTAL_R26_ALL * 100) : 0;
      meta.textContent = `${{filtered.length}} hotel${{filtered.length===1?'':'s'}} · ${{pct.toFixed(1)}}% of YTD revenue`;
    }}
  }}

  // Re-apply any active sort so filter changes preserve sort order
  reapplySort('hotelTable');
}}

window.addEventListener('load', ()=>{{
  drawTrendLine('cSessions',  SESS25,  SESS26, fmtN);
  drawTrendLine('cPageViews', PV25,    PV26,   fmtN);
  drawTrendLine('cRevenue',   REV25,   REV26,  fmtR);
  drawTrendLine('cPurchases', PURCH25, PURCH26,fmtN);
  renderHorizBars('chSessionBars', CH_NAMES, CH_SESS25, CH_SESS26, fmtN);
  renderHorizBars('chRevBars',     CH_NAMES, CH_REV25,  CH_REV26,  fmtR);
  renderConvBars('chConvBars', CH_NAMES, CH_CONV25, CH_CONV26);
  renderHubTable();
  renderStateCityTable();
  buildMgmtFilterOptions();
  renderHotelTable();
}});
window.addEventListener('resize', ()=>{{
  drawTrendLine('cSessions',  SESS25,  SESS26, fmtN);
  drawTrendLine('cPageViews', PV25,    PV26,   fmtN);
  drawTrendLine('cRevenue',   REV25,   REV26,  fmtR);
  drawTrendLine('cPurchases', PURCH25, PURCH26,fmtN);
}});
</script>
</body>
</html>"""

# ── Main ──────────────────────────────────────────────────────────────────
def main():
    print('Loading hotel names...')
    hotel_names, hotel_slug_index, hotel_id_map, hotel_mgmt_map = load_hotel_names()
    print(f'  Loaded {len(hotel_names)} hotel name mappings, {len(hotel_id_map)} hotel ID mappings, {len(hotel_mgmt_map)} mgmt co mappings')

    logo = load_logo()
    print(f'  Logo: {"loaded" if logo else "NOT FOUND"}')

    print('Parsing acquisition CSV...')
    acq_rows = read_multi_csv(ACQ_CSVS)
    (mo_sess_25_full, mo_ke_25_full, mo_rev_25_full,
     mo_sess_26, mo_ke_26, mo_rev_26,
     ch_sess_sp, ch_ke_sp, ch_rev_sp) = parse_acq(acq_rows)

    print('Parsing landing page CSVs...')
    lp_rows = read_multi_csv(LP_CSVS)
    print(f'  Total LP rows: {len(lp_rows):,}')
    mo_pv_25_full, mo_pv_26, hotel, home_page, locations_index, site_pages, state_city = parse_lp(lp_rows, hotel_names, hotel_slug_index, hotel_id_map)

    print('Building data arrays...')
    d = build_js_arrays(mo_sess_25_full, mo_ke_25_full, mo_rev_25_full,
                        mo_sess_26, mo_ke_26, mo_rev_26,
                        mo_pv_25_full, mo_pv_26,
                        ch_sess_sp, ch_ke_sp, ch_rev_sp,
                        hotel, home_page, locations_index, site_pages, state_city,
                        hotel_names, hotel_slug_index, hotel_mgmt_map)

    k = calc_kpis(d)
    print(f'  Hotels: {len(d["HOTEL_DATA"])}, Hub pages: {len(d["HUB_DATA"])}, State/City pages: {len(d["SC_DATA"])}')
    print(f'  2026 YTD sessions: {fmt_n(k["s26"])}, revenue: {fmt_r(k["r26"])}')

    print(f'Writing {OUT_FILE}...')
    html = build_html(d, k, logo)
    with open(OUT_FILE, 'w', encoding='utf-8') as f:
        f.write(html)

    print('Done!')

if __name__ == '__main__':
    main()
