#!/usr/bin/env python3
"""
build_ga4_website_dashboard.py
Generates ga4-website-dashboard.html AND wifi-dashboard.html
from GA4 full-year (2024 vs 2025) CSV exports.
"""

import csv, json, os, re
from collections import defaultdict
import openpyxl

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ── Input files ───────────────────────────────────────────────────────────
ACQ_CSV  = os.path.join(SCRIPT_DIR, 'Traffic_acquisition_Session_Updated_Channel_Group (2).csv')
LP_CSVS  = [os.path.join(SCRIPT_DIR, f) for f in [
    '2024-Q1-Landing-Pages.csv',
    '2024-Q2-Landing-Pages.csv',
    '2024-Q3-Landing-Pages.csv',
    '2024-Q4-Landing-Pages.csv',
    '2025-Q1-Landing-Pages.csv',
    '2025-Q2-Landing-Pages.csv',
    '2025-Q3-Landing-Pages.csv',
    '2025-Q4-Landing-Pages.csv',
]]
DB_XLSX  = os.path.join(SCRIPT_DIR, 'My Place Hotels Database - Master.xlsx')
LOGO_TXT = os.path.join(SCRIPT_DIR, 'logo_uri.txt')

# ── Output files ──────────────────────────────────────────────────────────
WEB_OUT  = os.path.join(SCRIPT_DIR, 'ga4-website-dashboard.html')
WIFI_OUT = os.path.join(SCRIPT_DIR, 'wifi-dashboard.html')

# ── Constants ─────────────────────────────────────────────────────────────
WIFI_CH = "Property WiFi"   # matched via startswith to handle varying trailing spaces

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
    '/locations/my-place-hotel-yakima-wam/my-place-hotel-yakima-wa': '/locations/my-place-hotel-yakima-wa',
    '/locations/my-place-hotel-idaho-falls-id[https://...CDN junk...': '/locations/my-place-hotel-idaho-falls-id',
    '/locations/randolph-vt': '/locations/my-place-hotel-randolph-vt',
    '/locations/my-place-hotel-Randolph-vt': '/locations/my-place-hotel-randolph-vt',
    '/locations/my-place-hotel-randolph-vtBringing': '/locations/my-place-hotel-randolph-vt',
    '/locations/my-place-hotel-jonesboro': '/locations/my-place-hotel-jonesboro-ar',
    '/locations/my-place-h\u2026"': None,
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
    '/locations/my-place-hotel-anchorage-a': '/locations/my-place-hotel-anchorage-ak',
    '/locations/my-place-hotel-augusta-ga': '/locations/my-place-hotel-augusta',
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

CH_COLORS = {
    "Organic Search":  "#1C355E",
    "Direct":          "#2E8B57",
    "Google HPA":      "#8B6914",
    "Paid Search":     "#7B2D8B",
    "Cross-network":   "#C44B2B",
    "Referral":        "#2B7CB0",
    "Unassigned":      "#5B8C5A",
    "Email":           "#B0702B",
    "Display":         "#2B4EA8",
    "Organic Shopping":"#8B3030",
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

def read_multi_csv(filepaths):
    """Read and merge multiple CSV files with the same column structure."""
    all_rows = []
    for fp in filepaths:
        chunk = read_csv(fp)
        print(f'    {os.path.basename(fp)}: {len(chunk):,} rows')
        all_rows.extend(chunk)
    return all_rows

# ── Load hotel names from Excel ───────────────────────────────────────────
def load_hotel_names():
    hotel_names = {}      # path → display name
    hotel_slug_index = {} # slug → path
    hotel_id_map = {}     # Hotel ID (int) → canonical /locations/ path
    try:
        wb = openpyxl.load_workbook(DB_XLSX, read_only=True, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or len(row) < 13:
                continue
            hotel_id = row[0]    # Column A — internal Hotel ID
            name     = row[1]    # Column B — Hotel Name
            url      = row[12]   # Column M — Property URL (May 2026 layout after Mgmt Co + Pilot cols added)
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
    return hotel_names, hotel_slug_index, hotel_id_map

def _is_hotel_property_path(p):
    """Check if a /locations/ path is a hotel property page (not a state/city hub)."""
    sub = p.replace('/locations/', '')
    if sub.startswith('my-place-hotel') or sub.startswith('savannah-airport'):
        return True
    if '/' in sub:
        return False
    if sub == 'location-results':
        return False
    return False

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
    # Monthly totals (non-WiFi)
    mo_sess = {2024:{m:0 for m in range(1,13)}, 2025:{m:0 for m in range(1,13)}}
    mo_ke   = {2024:{m:0 for m in range(1,13)}, 2025:{m:0 for m in range(1,13)}}
    mo_rev  = {2024:{m:0.0 for m in range(1,13)}, 2025:{m:0.0 for m in range(1,13)}}
    # WiFi monthly
    wifi_mo_sess = {2024:{m:0 for m in range(1,13)}, 2025:{m:0 for m in range(1,13)}}
    # Channel annual totals
    ch_sess = defaultdict(lambda:{2024:0, 2025:0})
    ch_ke   = defaultdict(lambda:{2024:0, 2025:0})
    ch_rev  = defaultdict(lambda:{2024:0.0, 2025:0.0})

    for row in rows:
        channel = row.get('Session Updated Channel Group','').strip()
        date    = str(row.get('Date','')).strip()
        if len(date) != 8: continue
        year  = int(date[:4])
        month = int(date[4:6])
        if year not in (2024,2025) or month not in range(1,13): continue
        try:
            sess = float(row.get('Sessions') or 0)
            ke   = float(row.get('Key events') or 0)
            rev  = float(row.get('Total revenue') or 0)
        except (ValueError, TypeError):
            continue

        if channel.startswith(WIFI_CH):
            wifi_mo_sess[year][month] += sess
        else:
            mo_sess[year][month] += sess
            mo_ke[year][month]   += ke
            mo_rev[year][month]  += rev
            ch_sess[channel][year] += sess
            ch_ke[channel][year]   += ke
            ch_rev[channel][year]  += rev

    return mo_sess, mo_ke, mo_rev, wifi_mo_sess, ch_sess, ch_ke, ch_rev

# ── Parse landing page CSV ────────────────────────────────────────────────
def parse_lp(rows, hotel_names, hotel_slug_index, hotel_id_map):
    mo_pv  = {2024:{m:0 for m in range(1,13)}, 2025:{m:0 for m in range(1,13)}}
    hotel  = defaultdict(lambda:{'v24':0,'v25':0,'ke24':0,'ke25':0,'r24':0.0,'r25':0.0})
    wifi   = defaultdict(lambda:{'v24':0,'v25':0,'u24':0,'u25':0})
    home_page       = {'v24':0,'v25':0,'ke24':0,'ke25':0,'r24':0.0,'r25':0.0}
    locations_index = {'v24':0,'v25':0,'ke24':0,'ke25':0,'r24':0.0,'r25':0.0}
    site_pages      = defaultdict(lambda:{'v24':0,'v25':0,'ke24':0,'ke25':0,'r24':0.0,'r25':0.0})
    state_city      = defaultdict(lambda:{'v24':0,'v25':0,'ke24':0,'ke25':0,'r24':0.0,'r25':0.0})
    booking_attributed = 0

    for row in rows:
        raw_path = str(row.get('Landing page + query string','')).strip()
        date  = str(row.get('Date','')).strip()
        if len(date) != 8: continue
        year  = int(date[:4])
        month = int(date[4:6])
        if year not in (2024,2025) or month not in range(1,13): continue
        try:
            views = int(float(row.get('Views') or 0))
            users = int(float(row.get('Active users') or 0))
            ke    = float(row.get('Key events') or 0)
            rev   = float(row.get('Total revenue') or 0)
        except (ValueError, TypeError):
            continue

        ys = '24' if year == 2024 else '25'

        # ── Booking URL handling ──
        hotel_id_match = re.search(r'[?&]hotel=(\d+)', raw_path)
        if hotel_id_match:
            hid = int(hotel_id_match.group(1))
            canon_path = hotel_id_map.get(hid)
            if canon_path and canon_path not in REMOVED_HOTELS:
                mo_pv[year][month] += views
                hotel[canon_path][f'v{ys}']  += views
                hotel[canon_path][f'ke{ys}'] += ke
                hotel[canon_path][f'r{ys}']  += rev
                booking_attributed += 1
            continue

        # ── Standard path handling ──
        path = raw_path.split('?')[0]
        if is_artifact(path) or path == '(not set)': continue

        if path.startswith('/wifi/'):
            p = WIFI_PAGE_MERGES.get(path, path)
            if p is None or p in REMOVED_HOTELS: continue
            wifi[p][f'v{ys}'] += views
            wifi[p][f'u{ys}'] += users
        else:
            p = LP_PAGE_MERGES.get(path, path)
            if p is None or p in REMOVED_HOTELS: continue
            mo_pv[year][month] += views

            # Classify
            if p in hotel_names:
                hotel[p][f'v{ys}']  += views
                hotel[p][f'ke{ys}'] += ke
                hotel[p][f'r{ys}']  += rev
            elif p.startswith('/locations/') and _is_hotel_property_path(p):
                slug = p.split('/')[-1]
                if slug in hotel_slug_index:
                    canon = hotel_slug_index[slug]
                    hotel[canon][f'v{ys}']  += views
                    hotel[canon][f'ke{ys}'] += ke
                    hotel[canon][f'r{ys}']  += rev
            elif p == '/':
                home_page[f'v{ys}'] += views
                home_page[f'ke{ys}'] += ke
                home_page[f'r{ys}'] += rev
            elif p in ('/locations', '/locations/', '/locations/location-results'):
                locations_index[f'v{ys}'] += views
                locations_index[f'ke{ys}'] += ke
                locations_index[f'r{ys}'] += rev
            elif p.startswith('/locations/'):
                pass  # state/city browse pages — drop
            elif p.startswith('/state-city/'):
                state_city[p][f'v{ys}'] += views
                state_city[p][f'ke{ys}'] += ke
                state_city[p][f'r{ys}'] += rev
            elif p.strip():
                site_pages[p][f'v{ys}'] += views
                site_pages[p][f'ke{ys}'] += ke
                site_pages[p][f'r{ys}'] += rev

    print(f'  Booking URLs attributed to hotels: {booking_attributed}')
    return mo_pv, hotel, home_page, locations_index, site_pages, state_city, wifi

# ── Build JS data arrays ───────────────────────────────────────────────────
def build_js_arrays(mo_sess, mo_ke, mo_rev, mo_pv, wifi_mo_sess,
                    ch_sess, ch_ke, ch_rev,
                    hotel, home_page, locations_index, site_pages, state_city, wifi,
                    hotel_names, hotel_slug_index):
    # Monthly arrays (12 elements each)
    SESS24  = [round(mo_sess[2024][m]) for m in range(1,13)]
    SESS25  = [round(mo_sess[2025][m]) for m in range(1,13)]
    PV24    = [round(mo_pv[2024][m])   for m in range(1,13)]
    PV25    = [round(mo_pv[2025][m])   for m in range(1,13)]
    REV24   = [round(mo_rev[2024][m],2) for m in range(1,13)]
    REV25   = [round(mo_rev[2025][m],2) for m in range(1,13)]
    PURCH24 = [round(mo_ke[2024][m])   for m in range(1,13)]
    PURCH25 = [round(mo_ke[2025][m])   for m in range(1,13)]
    WIFI_SESS24 = [round(wifi_mo_sess[2024][m]) for m in range(1,13)]
    WIFI_SESS25 = [round(wifi_mo_sess[2025][m]) for m in range(1,13)]

    # Channel names (sorted by 2025 sessions desc)
    all_channels = set(ch_sess.keys())
    CH_NAMES = sorted(all_channels, key=lambda c: ch_sess[c][2025], reverse=True)

    CH_SESS24 = {c: round(ch_sess[c][2024]) for c in CH_NAMES}
    CH_SESS25 = {c: round(ch_sess[c][2025]) for c in CH_NAMES}
    CH_REV24  = {c: round(ch_rev[c][2024],2) for c in CH_NAMES}
    CH_REV25  = {c: round(ch_rev[c][2025],2) for c in CH_NAMES}
    CH_CONV24 = {c: round(ch_ke[c][2024]/ch_sess[c][2024]*100, 3) if ch_sess[c][2024] else 0 for c in CH_NAMES}
    CH_CONV25 = {c: round(ch_ke[c][2025]/ch_sess[c][2025]*100, 3) if ch_sess[c][2025] else 0 for c in CH_NAMES}
    CH_COLORS_OUT = {c: CH_COLORS.get(c, '#888888') for c in CH_NAMES}

    # Hotel data sorted by total revenue desc
    hotel_list = []
    for path, d in hotel.items():
        label = resolve_hotel_name(path, hotel_names, hotel_slug_index)
        cr24 = round(d['ke24'] / d['v24'] * 100, 2) if d['v24'] else 0
        cr25 = round(d['ke25'] / d['v25'] * 100, 2) if d['v25'] else 0
        hotel_list.append({
            'label': label, 'path': path,
            'v24': round(d['v24']), 'v25': round(d['v25']),
            'ke24': round(d['ke24']), 'ke25': round(d['ke25']),
            'r24': round(d['r24'],2), 'r25': round(d['r25'],2),
            'cr24': cr24, 'cr25': cr25,
        })
    hotel_list.sort(key=lambda x: x['r24']+x['r25'], reverse=True)

    # Site pages: Home Page, Locations Index, + non-/locations/ pages with revenue
    hub_list = []
    hp = home_page
    hub_list.append({
        'label': 'Home Page', 'path': '/',
        'v24': round(hp['v24']), 'v25': round(hp['v25']),
        'ke24': round(hp['ke24']), 'ke25': round(hp['ke25']),
        'r24': round(hp['r24'],2), 'r25': round(hp['r25'],2),
    })
    li = locations_index
    hub_list.append({
        'label': 'Locations Index', 'path': '/locations',
        'v24': round(li['v24']), 'v25': round(li['v25']),
        'ke24': round(li['ke24']), 'ke25': round(li['ke25']),
        'r24': round(li['r24'],2), 'r25': round(li['r25'],2),
    })
    for path, d in sorted(site_pages.items(),
                          key=lambda x: x[1]['r24'] + x[1]['r25'], reverse=True):
        if d['r24'] > 0 or d['r25'] > 0:
            hub_list.append({
                'label': path, 'path': path,
                'v24': round(d['v24']), 'v25': round(d['v25']),
                'ke24': round(d['ke24']), 'ke25': round(d['ke25']),
                'r24': round(d['r24'],2), 'r25': round(d['r25'],2),
            })

    # State/City landing pages — all pages with any page views, sorted by views desc
    sc_list = []
    for path, d in sorted(state_city.items(),
                          key=lambda x: x[1]['v24'] + x[1]['v25'], reverse=True):
        if d['v24'] > 0 or d['v25'] > 0:
            sc_list.append({
                'label': path, 'path': path,
                'v24': round(d['v24']), 'v25': round(d['v25']),
                'ke24': round(d['ke24']), 'ke25': round(d['ke25']),
                'r24': round(d['r24'],2), 'r25': round(d['r25'],2),
            })

    # WiFi data sorted by total views desc
    wifi_list = []
    for path, d in wifi.items():
        label = resolve_wifi_name(path, hotel_names, hotel_slug_index)
        wifi_list.append({
            'label': label, 'path': path,
            'v24': round(d['v24']), 'v25': round(d['v25']),
            'u24': round(d['u24']), 'u25': round(d['u25']),
        })
    wifi_list.sort(key=lambda x: x['v24']+x['v25'], reverse=True)

    return dict(
        SESS24=SESS24, SESS25=SESS25,
        PV24=PV24, PV25=PV25,
        REV24=REV24, REV25=REV25,
        PURCH24=PURCH24, PURCH25=PURCH25,
        WIFI_SESS24=WIFI_SESS24, WIFI_SESS25=WIFI_SESS25,
        CH_NAMES=CH_NAMES,
        CH_SESS24=CH_SESS24, CH_SESS25=CH_SESS25,
        CH_REV24=CH_REV24, CH_REV25=CH_REV25,
        CH_CONV24=CH_CONV24, CH_CONV25=CH_CONV25,
        CH_COLORS=CH_COLORS_OUT,
        HOTEL_DATA=hotel_list,
        HUB_DATA=hub_list,
        SC_DATA=sc_list,
        WIFI_DATA=wifi_list,
    )

# ── KPI calculations ───────────────────────────────────────────────────────
def calc_kpis(d):
    s24 = sum(d['SESS24']); s25 = sum(d['SESS25'])
    r24 = sum(d['REV24']);  r25 = sum(d['REV25'])
    p24 = sum(d['PURCH24']); p25 = sum(d['PURCH25'])
    conv24 = p24/s24*100 if s24 else 0
    conv25 = p25/s25*100 if s25 else 0
    rpr24  = r24/p24 if p24 else 0
    rpr25  = r25/p25 if p25 else 0
    ws24 = sum(d['WIFI_SESS24']); ws25 = sum(d['WIFI_SESS25'])
    # WiFi page views
    wv24 = sum(x['v24'] for x in d['WIFI_DATA'])
    wv25 = sum(x['v25'] for x in d['WIFI_DATA'])
    wu24 = sum(x['u24'] for x in d['WIFI_DATA'])
    wu25 = sum(x['u25'] for x in d['WIFI_DATA'])
    n_wifi = len(d['WIFI_DATA'])
    avg_ws24 = ws24/n_wifi if n_wifi else 0
    avg_ws25 = ws25/n_wifi if n_wifi else 0
    return dict(s24=s24,s25=s25,r24=r24,r25=r25,p24=p24,p25=p25,
                conv24=conv24,conv25=conv25,rpr24=rpr24,rpr25=rpr25,
                ws24=ws24,ws25=ws25,wv24=wv24,wv25=wv25,
                wu24=wu24,wu25=wu25,n_wifi=n_wifi,
                avg_ws24=avg_ws24,avg_ws25=avg_ws25)

# ── HTML: shared JS/CSS ───────────────────────────────────────────────────
COMMON_CSS = """
:root {
  --bg:#EEF1F6; --card:#FFFFFF; --border:#D0D8E8;
  --text:#1C355E; --muted:#5A6B84; --navy:#1C355E;
  --orange:#FF5F00; --green:#2E8B57; --yellow:#FFD080;
  --C24:#4A7FC1; --C25:#FF5F00;
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
            color:#5A6B84; display:flex; align-items:center; gap:10px; }
.info-bar.purple { background:rgba(123,45,139,0.06); border:1px solid rgba(123,45,139,0.25); }
.info-bar.purple strong { color:#7B2D8B; }
.info-bar.blue { background:rgba(74,127,193,0.07); border:1px solid rgba(74,127,193,0.25); }
.info-bar.blue strong { color:#1C355E; }
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
.hotel-table thead th:nth-child(1) { width:19%; }
.hotel-table thead th { width:6.75%; }
.footer { background:#1C355E; border-top:3px solid #FF5F00; padding:14px 40px;
          text-align:center; color:rgba(255,255,255,0.55); font-size:10px; }
"""

COMMON_JS = """
const C24='#4A7FC1', C25='#FF5F00';

function fmtN(n){ return n>=1e6?(n/1e6).toFixed(2)+'M':n>=1e3?(n/1e3).toFixed(1)+'K':Math.round(n).toLocaleString(); }
function fmtR(n){ return n>=1e6?'$'+(n/1e6).toFixed(2)+'M':n>=1e3?'$'+(n/1e3).toFixed(1)+'K':'$'+Math.round(n).toLocaleString(); }

function drawTrendLine(canvasId, arr24, arr25, fmtFn) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  const dpr = window.devicePixelRatio || 1;
  const W = canvas.parentElement.clientWidth || 800;
  const H = parseInt(canvas.getAttribute('height')) || 150;
  canvas.width = W*dpr; canvas.height = H*dpr;
  canvas.style.width = W+'px'; canvas.style.height = H+'px';
  const ctx = canvas.getContext('2d');
  ctx.scale(dpr, dpr);
  const PAD = {t:20, r:14, b:24, l:52};
  const cW = W - PAD.l - PAD.r, cH = H - PAD.t - PAD.b;
  const n = arr24.length;
  const maxVal = Math.max(...arr24, ...arr25) * 1.18;
  const xp = i => PAD.l + (i/(n-1)) * cW;
  const yp = v => PAD.t + cH * (1 - v/maxVal);
  const moL = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  for (let i=0; i<=4; i++) {
    const v = maxVal * i/4, y = yp(v);
    ctx.strokeStyle='rgba(208,216,232,0.7)'; ctx.lineWidth=0.5;
    ctx.beginPath(); ctx.moveTo(PAD.l,y); ctx.lineTo(PAD.l+cW,y); ctx.stroke();
    ctx.fillStyle='#8ba0bf'; ctx.font='8px Arial';
    ctx.textAlign='right'; ctx.fillText(fmtFn(v), PAD.l-4, y+3);
  }
  function drawLine(arr, color, fillAlpha) {
    ctx.beginPath();
    for (let i=0;i<n;i++) { i===0 ? ctx.moveTo(xp(i),yp(arr[i])) : ctx.lineTo(xp(i),yp(arr[i])); }
    ctx.lineTo(xp(n-1), PAD.t+cH); ctx.lineTo(xp(0), PAD.t+cH); ctx.closePath();
    ctx.globalAlpha = fillAlpha; ctx.fillStyle = color; ctx.fill(); ctx.globalAlpha = 1;
    ctx.beginPath();
    for (let i=0;i<n;i++) { i===0 ? ctx.moveTo(xp(i),yp(arr[i])) : ctx.lineTo(xp(i),yp(arr[i])); }
    ctx.strokeStyle=color; ctx.lineWidth=2; ctx.lineJoin='round'; ctx.stroke();
    for (let i=0;i<n;i++) {
      ctx.beginPath(); ctx.arc(xp(i),yp(arr[i]),3,0,Math.PI*2);
      ctx.fillStyle=color; ctx.fill();
    }
  }
  drawLine(arr24, C24, 0.08);
  drawLine(arr25, C25, 0.10);
  ctx.fillStyle='#8ba0bf'; ctx.font='8px Arial'; ctx.textAlign='center';
  for (let i=0;i<n;i++) ctx.fillText(moL[i], xp(i), H-4);
  const maxIdx = arr25.indexOf(Math.max(...arr25));
  ctx.font='bold 8px Arial'; ctx.fillStyle=C25; ctx.textAlign='center';
  ctx.fillText(fmtFn(arr25[maxIdx]), xp(maxIdx), yp(arr25[maxIdx])-6);
}

function renderHorizBars(id, channels, d24, d25, fmtFn) {
  const el=document.getElementById(id); if(!el) return;
  const maxV = Math.max(...channels.map(c=>Math.max(d24[c]||0,d25[c]||0)));
  el.innerHTML = channels.map(ch=>{
    const v24=d24[ch]||0, v25=d25[ch]||0;
    const w24=maxV?(v24/maxV*100).toFixed(1):0, w25=maxV?(v25/maxV*100).toFixed(1):0;
    const delta = v24>0?((v25-v24)/v24*100):999;
    const dcls  = delta>=0?'up':'down';
    const dTxt  = v24===0
      ? '<span class="delta-pill new">New</span>'
      : `<span class="delta-pill ${dcls}">${delta>=0?'▲':'▼'} ${Math.abs(delta).toFixed(0)}%</span>`;
    return `<div class="ch-bar-section">
      <div class="ch-bar-label">${ch} ${dTxt}</div>
      <div class="bar-row"><div class="bar-yr">'24</div>
        <div class="bar-track"><div class="bar-fill" style="width:${w24}%;background:${C24};opacity:0.75;"></div></div>
        <div class="bar-val">${fmtFn(v24)}</div></div>
      <div class="bar-row"><div class="bar-yr">'25</div>
        <div class="bar-track"><div class="bar-fill" style="width:${w25}%;background:${C25};"></div></div>
        <div class="bar-val">${fmtFn(v25)}</div></div>
    </div>`;
  }).join('');
}

function renderConvBars(id, channels, c24, c25) {
  const el=document.getElementById(id); if(!el) return;
  const rel = channels.filter(c=>(c24[c]||0)+(c25[c]||0)>0);
  const maxR = Math.max(...rel.map(c=>Math.max(c24[c]||0,c25[c]||0)));
  const half=Math.ceil(rel.length/2), left=rel.slice(0,half), right=rel.slice(half);
  const barH=ch=>{
    return ch.map(c=>{
      const v24=c24[c]||0,v25=c25[c]||0;
      const w24=maxR?(v24/maxR*100).toFixed(1):0, w25=maxR?(v25/maxR*100).toFixed(1):0;
      const delta=v24>0?(v25-v24):0;
      const dcls=delta>=0?'up':'down';
      const dTxt=v24>0
        ?`<span class="delta-pill ${dcls}">${delta>=0?'▲':'▼'} ${Math.abs(delta).toFixed(2)}pp</span>`
        :'<span class="delta-pill new">New</span>';
      return `<div class="ch-bar-section">
        <div class="ch-bar-label">${c} ${dTxt}</div>
        <div class="bar-row"><div class="bar-yr">'24</div>
          <div class="bar-track"><div class="bar-fill" style="width:${w24}%;background:${C24};opacity:0.75;"></div></div>
          <div class="bar-val">${v24.toFixed(2)}%</div></div>
        <div class="bar-row"><div class="bar-yr">'25</div>
          <div class="bar-track"><div class="bar-fill" style="width:${w25}%;background:${C25};"></div></div>
          <div class="bar-val">${v25.toFixed(2)}%</div></div>
      </div>`;
    }).join('');
  };
  el.innerHTML=`<div style="display:grid;grid-template-columns:1fr 1fr;gap:20px;">
    <div>${barH(left)}</div><div>${barH(right)}</div></div>`;
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

# ── Generate website HTML ─────────────────────────────────────────────────
def build_website_html(d, k, logo):
    n_hotels = len(d['HOTEL_DATA'])
    wifi_total24 = fmt_n(sum(d['WIFI_SESS24']))
    wifi_total25 = fmt_n(sum(d['WIFI_SESS25']))

    data_js = f"""
const SESS24={json.dumps(d['SESS24'])}, SESS25={json.dumps(d['SESS25'])};
const PV24={json.dumps(d['PV24'])},    PV25={json.dumps(d['PV25'])};
const REV24={json.dumps(d['REV24'])},  REV25={json.dumps(d['REV25'])};
const PURCH24={json.dumps(d['PURCH24'])},PURCH25={json.dumps(d['PURCH25'])};
const CH_NAMES={json.dumps(d['CH_NAMES'])};
const CH_SESS24={json.dumps(d['CH_SESS24'])}, CH_SESS25={json.dumps(d['CH_SESS25'])};
const CH_REV24={json.dumps(d['CH_REV24'])},   CH_REV25={json.dumps(d['CH_REV25'])};
const CH_CONV24={json.dumps(d['CH_CONV24'])}, CH_CONV25={json.dumps(d['CH_CONV25'])};
const CH_COLORS={json.dumps(d['CH_COLORS'])};
const HUB_DATA={json.dumps(d['HUB_DATA'])};
const SC_DATA={json.dumps(d['SC_DATA'])};
const HOTEL_DATA={json.dumps(d['HOTEL_DATA'])};
"""

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>My Place Hotels — Website Analytics 2024 vs 2025</title>
<style>
{COMMON_CSS}
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
        <h1>Website Analytics Dashboard</h1>
        <div class="header-subtitle">Google Analytics 4 &nbsp;·&nbsp; Full Year 2024 vs 2025 &nbsp;·&nbsp; myplacehotels.com &nbsp;·&nbsp; Property WiFi excluded</div>
      </div>
    </div>
    <div class="header-meta">
      <div style="display:flex;gap:8px;">
        <a href="wifi-dashboard.html" class="portal-btn">📶 WiFi Dashboard</a>
        <a href="index.html" class="portal-btn">← Portal</a>
      </div>
      <div class="updated-pill">Jan 2024 – Dec 2025</div>
    </div>
  </div>
</div>

<div class="main">

  <div class="info-bar purple">
    📶 <span><strong>Property WiFi sessions are fully excluded from this dashboard</strong> — {wifi_total24} sessions (2024) and {wifi_total25} sessions (2025) are reported in the <a href="wifi-dashboard.html" style="color:#7B2D8B;font-weight:700;">Property WiFi Dashboard</a>. Every metric below reflects website acquisition traffic only.</span>
  </div>

  <div class="kpi-row">
    <div class="kpi-card">
      <div class="kpi-label">2025 Web Sessions</div>
      <div class="kpi-value">{fmt_n(k['s25'])}</div>
      <div class="kpi-sub">2024: {fmt_n(k['s24'])}</div>
      <div class="kpi-yoy">{yoy_badge(k['s25'],k['s24'])} vs 2024</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">2025 Website Revenue</div>
      <div class="kpi-value">{fmt_r(k['r25'])}</div>
      <div class="kpi-sub">2024: {fmt_r(k['r24'])}</div>
      <div class="kpi-yoy">{yoy_badge(k['r25'],k['r24'])} vs 2024</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">2025 Reservations</div>
      <div class="kpi-value">{fmt_n(k['p25'])}</div>
      <div class="kpi-sub">2024: {fmt_n(k['p24'])}</div>
      <div class="kpi-yoy">{yoy_badge(k['p25'],k['p24'])} vs 2024</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">2025 Reservation Rate</div>
      <div class="kpi-value">{k['conv25']:.2f}%</div>
      <div class="kpi-sub">2024: {k['conv24']:.2f}%</div>
      <div class="kpi-yoy">{yoy_badge(k['conv25'],k['conv24'])} vs 2024</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">Revenue per Reservation</div>
      <div class="kpi-value">{fmt_r(k['rpr25'])}</div>
      <div class="kpi-sub">2024: {fmt_r(k['rpr24'])}</div>
      <div class="kpi-yoy">{yoy_badge(k['rpr25'],k['rpr24'])} vs 2024</div>
    </div>
  </div>

  <div class="section-hdr">
    <h2>Monthly Trends</h2>
    <span class="desc">2024 (blue) vs 2025 (orange) — excludes Property WiFi</span>
    <div class="section-divider"></div>
  </div>

  <div class="grid-2" style="margin-bottom:16px;">
    <div class="chart-card">
      <h3>Monthly Sessions</h3>
      <p class="chart-desc">Total website sessions per month, 2024 vs 2025. Excludes Property WiFi channel.</p>
      <div class="legend-row">
        <div class="legend-item"><div class="legend-swatch" style="background:var(--C24);"></div>2024</div>
        <div class="legend-item"><div class="legend-swatch" style="background:var(--C25);"></div>2025</div>
      </div>
      <canvas id="cSessions" style="width:100%;display:block;" height="150"></canvas>
    </div>
    <div class="chart-card">
      <h3>Monthly Page Views</h3>
      <p class="chart-desc">Total page views per month across all website pages, 2024 vs 2025. Excludes Property WiFi (/wifi/) pages.</p>
      <div class="legend-row">
        <div class="legend-item"><div class="legend-swatch" style="background:var(--C24);"></div>2024</div>
        <div class="legend-item"><div class="legend-swatch" style="background:var(--C25);"></div>2025</div>
      </div>
      <canvas id="cPageViews" style="width:100%;display:block;" height="150"></canvas>
    </div>
  </div>

  <div class="grid-2" style="margin-bottom:16px;">
    <div class="chart-card">
      <h3>Monthly Website Revenue</h3>
      <p class="chart-desc">GA4-attributed booking revenue per month. Reflects revenue from sessions originating on the website.</p>
      <div class="legend-row">
        <div class="legend-item"><div class="legend-swatch" style="background:var(--C24);"></div>2024</div>
        <div class="legend-item"><div class="legend-swatch" style="background:var(--C25);"></div>2025</div>
      </div>
      <canvas id="cRevenue" style="width:100%;display:block;" height="150"></canvas>
    </div>
    <div class="chart-card">
      <h3>Monthly Reservations</h3>
      <p class="chart-desc">Reservation key events per month — the primary conversion tracked in GA4.</p>
      <div class="legend-row">
        <div class="legend-item"><div class="legend-swatch" style="background:var(--C24);"></div>2024</div>
        <div class="legend-item"><div class="legend-swatch" style="background:var(--C25);"></div>2025</div>
      </div>
      <canvas id="cPurchases" style="width:100%;display:block;" height="150"></canvas>
    </div>
  </div>

  <div class="section-hdr">
    <h2>Channel Performance</h2>
    <span class="desc">Sessions and revenue by acquisition channel — 2024 vs 2025</span>
    <div class="section-divider"></div>
  </div>
  <div class="grid-2" style="margin-bottom:16px;">
    <div class="chart-card">
      <h3>Sessions by Channel</h3>
      <p class="chart-desc">Annual session volume per channel, showing where traffic originates and how the mix shifted year over year.</p>
      <div id="chSessionBars"></div>
    </div>
    <div class="chart-card">
      <h3>Website Revenue by Channel</h3>
      <p class="chart-desc">GA4-attributed booking revenue by channel — highlights which acquisition sources drive the most transactional value.</p>
      <div id="chRevBars"></div>
    </div>
  </div>

  <div class="section-hdr">
    <h2>Reservation Rate by Channel</h2>
    <span class="desc">Reservations ÷ Sessions per channel — 2024 vs 2025</span>
    <div class="section-divider"></div>
  </div>
  <div class="chart-card" style="margin-bottom:16px;">
    <h3>Session Reservation Rate by Channel</h3>
    <p class="chart-desc">Percentage of sessions resulting in a reservation. Higher rates indicate high-intent visitors already primed to book.</p>
    <div id="chConvBars"></div>
  </div>

  <div class="section-hdr">
    <h2>Site Page Performance</h2>
    <span class="desc">Home page, locations index, and other site pages with attributed revenue — full year 2024 vs 2025</span>
    <div class="section-divider"></div>
  </div>
  <div class="table-card">
    <h3>Revenue-Generating Site Pages — 2024 vs 2025</h3>
    <p class="chart-desc">Top 25 revenue-generating pages. Includes consolidated home page, locations index, and individual site pages. Totals reflect all pages.</p>
    <div style="overflow-x:auto;"><table id="hubTable">
      <thead><tr>
        <th onclick="sortTbl('hubTable',0,false)">Page</th>
        <th class="num" onclick="sortTbl('hubTable',1,true)">Views '24</th>
        <th class="num" onclick="sortTbl('hubTable',2,true)">Views '25</th>
        <th class="num" onclick="sortTbl('hubTable',3,true)">Δ Views</th>
        <th class="num" onclick="sortTbl('hubTable',4,true)">Res '24</th>
        <th class="num" onclick="sortTbl('hubTable',5,true)">Res '25</th>
        <th class="num" onclick="sortTbl('hubTable',6,true)">Revenue '24</th>
        <th class="num" onclick="sortTbl('hubTable',7,true)">Revenue '25</th>
        <th class="num" onclick="sortTbl('hubTable',8,true)">Δ Revenue</th>
      </tr></thead>
      <tbody id="hubBody"></tbody>
      <tfoot id="hubFoot"></tfoot>
    </table></div>
  </div>

  <!-- State/City Landing Pages -->
  <div class="section-hdr">
    <h2>State &amp; City Landing Pages</h2>
    <span class="desc">All /state-city/ landing pages with traffic — full year 2024 vs 2025 · click any column to re-sort</span>
    <div class="section-divider"></div>
  </div>
  <div class="table-card">
    <h3>State/City Pages — 2024 vs 2025</h3>
    <p class="chart-desc">Performance of state and city browse pages. Includes all pages with page views in either year.</p>
    <div style="overflow-x:auto;"><table id="scTable">
      <thead><tr>
        <th onclick="sortTbl('scTable',0,false)">Page</th>
        <th class="num" onclick="sortTbl('scTable',1,true)">Views '24</th>
        <th class="num" onclick="sortTbl('scTable',2,true)">Views '25</th>
        <th class="num" onclick="sortTbl('scTable',3,true)">Δ Views</th>
        <th class="num" onclick="sortTbl('scTable',4,true)">Res '24</th>
        <th class="num" onclick="sortTbl('scTable',5,true)">Res '25</th>
        <th class="num" onclick="sortTbl('scTable',6,true)">Revenue '24</th>
        <th class="num" onclick="sortTbl('scTable',7,true)">Revenue '25</th>
        <th class="num" onclick="sortTbl('scTable',8,true)">Δ Revenue</th>
      </tr></thead>
      <tbody id="scBody"></tbody>
      <tfoot id="scFoot"></tfoot>
    </table></div>
  </div>

  <div class="section-hdr">
    <h2>Property Page Performance — All {n_hotels} Hotels</h2>
    <span class="desc">Individual hotel landing pages — all properties, sorted by total revenue. Click any column to re-sort.</span>
    <div class="section-divider"></div>
  </div>
  <div class="table-card">
    <h3>Hotel Property Pages — 2024 vs 2025</h3>
    <p class="chart-desc">Page views, reservations, and revenue for every property page. Δ columns show year-over-year change.</p>
    <table id="hotelTable" class="hotel-table">
      <thead><tr>
        <th onclick="sortTbl('hotelTable',0,false)">Property</th>
        <th class="num" onclick="sortTbl('hotelTable',1,true)">Views '24</th>
        <th class="num" onclick="sortTbl('hotelTable',2,true)">Views '25</th>
        <th class="num" onclick="sortTbl('hotelTable',3,true)">Δ Views</th>
        <th class="num" onclick="sortTbl('hotelTable',4,true)">Res '24</th>
        <th class="num" onclick="sortTbl('hotelTable',5,true)">Res '25</th>
        <th class="num" onclick="sortTbl('hotelTable',6,true)">Δ Res</th>
        <th class="num" onclick="sortTbl('hotelTable',7,true)">CR '24</th>
        <th class="num" onclick="sortTbl('hotelTable',8,true)">CR '25</th>
        <th class="num" onclick="sortTbl('hotelTable',9,true)">Δ CR</th>
        <th class="num" onclick="sortTbl('hotelTable',10,true)">Revenue '24</th>
        <th class="num" onclick="sortTbl('hotelTable',11,true)">Revenue '25</th>
        <th class="num" onclick="sortTbl('hotelTable',12,true)">Δ Revenue</th>
      </tr></thead>
      <tbody id="hotelBody"></tbody>
      <tfoot id="hotelFoot"></tfoot>
    </table>
  </div>


</div>

<div class="footer">
  My Place Hotels of America &nbsp;·&nbsp; Website Analytics Dashboard &nbsp;·&nbsp; Google Analytics 4 &nbsp;·&nbsp; Confidential — Internal Use Only
</div>

<script>
{COMMON_JS}
{data_js}

function renderHubTable() {{
  const top25 = HUB_DATA.slice(0, 25);
  document.getElementById('hubBody').innerHTML = top25.map(r=>{{
    return `<tr>
      <td class="path-cell">${{r.label}}</td>
      <td class="num">${{r.v24.toLocaleString()}}</td>
      <td class="num">${{r.v25.toLocaleString()}}</td>
      ${{deltaCell(r.v24,r.v25,n=>Math.round(n).toLocaleString())}}
      <td class="num">${{r.ke24.toLocaleString()}}</td>
      <td class="num">${{r.ke25.toLocaleString()}}</td>
      <td class="num">${{fmtR(r.r24)}}</td>
      <td class="num">${{fmtR(r.r25)}}</td>
      ${{deltaCell(r.r24,r.r25,fmtR)}}
    </tr>`;
  }}).join('');
  // Totals from ALL hub pages (not just top 25)
  const tv24=HUB_DATA.reduce((s,r)=>s+r.v24,0), tv25=HUB_DATA.reduce((s,r)=>s+r.v25,0);
  const tk24=HUB_DATA.reduce((s,r)=>s+r.ke24,0), tk25=HUB_DATA.reduce((s,r)=>s+r.ke25,0);
  const tr24=HUB_DATA.reduce((s,r)=>s+r.r24,0), tr25=HUB_DATA.reduce((s,r)=>s+r.r25,0);
  document.getElementById('hubFoot').innerHTML = `<tr style="font-weight:700;border-top:2px solid #1C355E">
    <td class="path-cell">All Site Pages Total (${{HUB_DATA.length}} pages)</td>
    <td class="num">${{tv24.toLocaleString()}}</td>
    <td class="num">${{tv25.toLocaleString()}}</td>
    ${{deltaCell(tv24,tv25,n=>Math.round(n).toLocaleString())}}
    <td class="num">${{tk24.toLocaleString()}}</td>
    <td class="num">${{tk25.toLocaleString()}}</td>
    <td class="num">${{fmtR(tr24)}}</td>
    <td class="num">${{fmtR(tr25)}}</td>
    ${{deltaCell(tr24,tr25,fmtR)}}
  </tr>`;
}}

function renderStateCityTable() {{
  document.getElementById('scBody').innerHTML = SC_DATA.map(r=>{{
    return `<tr>
      <td class="path-cell">${{r.label}}</td>
      <td class="num">${{r.v24.toLocaleString()}}</td>
      <td class="num">${{r.v25.toLocaleString()}}</td>
      ${{deltaCell(r.v24,r.v25,n=>Math.round(n).toLocaleString())}}
      <td class="num">${{r.ke24.toLocaleString()}}</td>
      <td class="num">${{r.ke25.toLocaleString()}}</td>
      <td class="num">${{fmtR(r.r24)}}</td>
      <td class="num">${{fmtR(r.r25)}}</td>
      ${{deltaCell(r.r24,r.r25,fmtR)}}
    </tr>`;
  }}).join('');
  const tv24=SC_DATA.reduce((s,r)=>s+r.v24,0), tv25=SC_DATA.reduce((s,r)=>s+r.v25,0);
  const tk24=SC_DATA.reduce((s,r)=>s+r.ke24,0), tk25=SC_DATA.reduce((s,r)=>s+r.ke25,0);
  const tr24=SC_DATA.reduce((s,r)=>s+r.r24,0), tr25=SC_DATA.reduce((s,r)=>s+r.r25,0);
  document.getElementById('scFoot').innerHTML = `<tr style="font-weight:700;border-top:2px solid #1C355E">
    <td class="path-cell">All State/City Pages (${{SC_DATA.length}} pages)</td>
    <td class="num">${{tv24.toLocaleString()}}</td>
    <td class="num">${{tv25.toLocaleString()}}</td>
    ${{deltaCell(tv24,tv25,n=>Math.round(n).toLocaleString())}}
    <td class="num">${{tk24.toLocaleString()}}</td>
    <td class="num">${{tk25.toLocaleString()}}</td>
    <td class="num">${{fmtR(tr24)}}</td>
    <td class="num">${{fmtR(tr25)}}</td>
    ${{deltaCell(tr24,tr25,fmtR)}}
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
function renderHotelTable() {{
  document.getElementById('hotelBody').innerHTML = HOTEL_DATA.map(r=>{{
    return `<tr>
      <td class="name-cell">${{r.label}}</td>
      <td class="num">${{r.v24.toLocaleString()}}</td>
      <td class="num">${{r.v25.toLocaleString()}}</td>
      ${{deltaCell(r.v24,r.v25,n=>Math.round(n).toLocaleString())}}
      <td class="num">${{r.ke24.toLocaleString()}}</td>
      <td class="num">${{r.ke25.toLocaleString()}}</td>
      ${{deltaCell(r.ke24,r.ke25,n=>Math.round(n).toLocaleString())}}
      <td class="num">${{fmtCR(r.cr24)}}</td>
      <td class="num">${{fmtCR(r.cr25)}}</td>
      ${{deltaCR(r.cr24,r.cr25)}}
      <td class="num">${{fmtR(r.r24)}}</td>
      <td class="num">${{fmtR(r.r25)}}</td>
      ${{deltaCell(r.r24,r.r25,fmtR)}}
    </tr>`;
  }}).join('');
  // Totals row
  const tv24=HOTEL_DATA.reduce((s,r)=>s+r.v24,0), tv25=HOTEL_DATA.reduce((s,r)=>s+r.v25,0);
  const tk24=HOTEL_DATA.reduce((s,r)=>s+r.ke24,0), tk25=HOTEL_DATA.reduce((s,r)=>s+r.ke25,0);
  const tr24=HOTEL_DATA.reduce((s,r)=>s+r.r24,0), tr25=HOTEL_DATA.reduce((s,r)=>s+r.r25,0);
  const tcr24=tv24?tk24/tv24*100:0, tcr25=tv25?tk25/tv25*100:0;
  document.getElementById('hotelFoot').innerHTML = `<tr style="font-weight:700;border-top:2px solid #1C355E">
    <td class="name-cell">All Properties Total</td>
    <td class="num">${{tv24.toLocaleString()}}</td>
    <td class="num">${{tv25.toLocaleString()}}</td>
    ${{deltaCell(tv24,tv25,n=>Math.round(n).toLocaleString())}}
    <td class="num">${{tk24.toLocaleString()}}</td>
    <td class="num">${{tk25.toLocaleString()}}</td>
    ${{deltaCell(tk24,tk25,n=>Math.round(n).toLocaleString())}}
    <td class="num">${{fmtCR(tcr24)}}</td>
    <td class="num">${{fmtCR(tcr25)}}</td>
    ${{deltaCR(tcr24,tcr25)}}
    <td class="num">${{fmtR(tr24)}}</td>
    <td class="num">${{fmtR(tr25)}}</td>
    ${{deltaCell(tr24,tr25,fmtR)}}
  </tr>`;
}}

window.addEventListener('load', ()=>{{
  drawTrendLine('cSessions',  SESS24,  SESS25, fmtN);
  drawTrendLine('cPageViews', PV24,    PV25,   fmtN);
  drawTrendLine('cRevenue',   REV24,   REV25,  fmtR);
  drawTrendLine('cPurchases', PURCH24, PURCH25,fmtN);
  renderHorizBars('chSessionBars', CH_NAMES, CH_SESS24, CH_SESS25, fmtN);
  renderHorizBars('chRevBars',     CH_NAMES, CH_REV24,  CH_REV25,  fmtR);
  renderConvBars('chConvBars', CH_NAMES, CH_CONV24, CH_CONV25);
  renderHubTable();
  renderStateCityTable();
  renderHotelTable();
}});
window.addEventListener('resize', ()=>{{
  drawTrendLine('cSessions',  SESS24,  SESS25, fmtN);
  drawTrendLine('cPageViews', PV24,    PV25,   fmtN);
  drawTrendLine('cRevenue',   REV24,   REV25,  fmtR);
  drawTrendLine('cPurchases', PURCH24, PURCH25,fmtN);
}});
</script>
</body>
</html>"""

# ── Generate WiFi HTML ────────────────────────────────────────────────────
def build_wifi_html(d, k, logo):
    n_wifi = k['n_wifi']
    WIFI_CSS_EXTRA = "\n.kpi-card { border-top-color: #4A7FC1; }"

    wifi_data_js = f"""
const WIFI_SESS24={json.dumps(d['WIFI_SESS24'])};
const WIFI_SESS25={json.dumps(d['WIFI_SESS25'])};
const WIFI_DATA={json.dumps(d['WIFI_DATA'])};
"""

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>My Place Hotels — Property WiFi Analytics 2024 vs 2025</title>
<style>
{COMMON_CSS}{WIFI_CSS_EXTRA}
.hotel-table thead th:nth-child(2) {{ width:12%; }}
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
        <h1>Property WiFi Analytics Dashboard</h1>
        <div class="header-subtitle">Google Analytics 4 &nbsp;·&nbsp; Full Year 2024 vs 2025 &nbsp;·&nbsp; In-Property Digital Concierge Pages</div>
      </div>
    </div>
    <div class="header-meta">
      <div style="display:flex;gap:8px;">
        <a href="ga4-website-dashboard.html" class="portal-btn">🌐 Website Dashboard</a>
        <a href="index.html" class="portal-btn">← Portal</a>
      </div>
      <div class="updated-pill">Jan 2024 – Dec 2025</div>
    </div>
  </div>
</div>

<div class="main">

  <div class="info-bar blue">
    📶 <span><strong>Property WiFi pages are in-hotel digital concierge pages</strong> — these are not acquisition pages. Guests connecting to hotel WiFi are directed to property-specific /wifi/ pages for local info, amenities, and services. Sessions here represent <em>in-stay engagement</em>, not website acquisition traffic.</span>
  </div>

  <div class="kpi-row">
    <div class="kpi-card">
      <div class="kpi-label">2025 WiFi Sessions</div>
      <div class="kpi-value">{fmt_n(k['ws25'])}</div>
      <div class="kpi-sub">2024: {fmt_n(k['ws24'])}</div>
      <div class="kpi-yoy">{yoy_badge(k['ws25'],k['ws24'])} vs 2024</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">2025 WiFi Page Views</div>
      <div class="kpi-value">{fmt_n(k['wv25'])}</div>
      <div class="kpi-sub">2024: {fmt_n(k['wv24'])}</div>
      <div class="kpi-yoy">{yoy_badge(k['wv25'],k['wv24'])} vs 2024</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">2025 Active Users</div>
      <div class="kpi-value">{fmt_n(k['wu25'])}</div>
      <div class="kpi-sub">2024: {fmt_n(k['wu24'])}</div>
      <div class="kpi-yoy">{yoy_badge(k['wu25'],k['wu24'])} vs 2024</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">Properties with WiFi Pages</div>
      <div class="kpi-value">{n_wifi}</div>
      <div class="kpi-sub">Unique /wifi/ property pages</div>
      <div class="kpi-yoy" style="font-size:10px;color:#8ba0bf;">in-hotel digital concierge</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">Avg Sessions / Property</div>
      <div class="kpi-value">{fmt_n(k['avg_ws25'])}</div>
      <div class="kpi-sub">2024: {fmt_n(k['avg_ws24'])}</div>
      <div class="kpi-yoy">{yoy_badge(k['avg_ws25'],k['avg_ws24'])} vs 2024</div>
    </div>
  </div>

  <div class="section-hdr">
    <h2>Monthly WiFi Sessions</h2>
    <span class="desc">In-property WiFi page sessions per month — 2024 (blue) vs 2025 (orange)</span>
    <div class="section-divider"></div>
  </div>
  <div class="chart-card" style="margin-bottom:16px;">
    <h3>Monthly WiFi Sessions</h3>
    <p class="chart-desc">Total sessions across all property WiFi pages per month. Reflects in-stay guest engagement with digital concierge content.</p>
    <div class="legend-row">
      <div class="legend-item"><div class="legend-swatch" style="background:var(--C24);"></div>2024</div>
      <div class="legend-item"><div class="legend-swatch" style="background:var(--C25);"></div>2025</div>
    </div>
    <canvas id="cWifiSessions" style="width:100%;display:block;" height="150"></canvas>
  </div>

  <div class="section-hdr">
    <h2>Property WiFi Page Performance — All {n_wifi} Properties</h2>
    <span class="desc">Individual hotel WiFi pages — sorted by total views. Click any column to re-sort.</span>
    <div class="section-divider"></div>
  </div>
  <div class="table-card">
    <h3>Hotel WiFi Pages — 2024 vs 2025</h3>
    <p class="chart-desc">Page views and active users for each property's WiFi concierge page. Sorted by combined 2024+2025 views.</p>
    <table id="wifiTable" class="hotel-table">
      <thead><tr>
        <th onclick="sortTbl('wifiTable',0,false)">Property</th>
        <th onclick="sortTbl('wifiTable',1,false)" style="font-size:8px;color:#aac;">WiFi URL</th>
        <th class="num" onclick="sortTbl('wifiTable',2,true)">Views '24</th>
        <th class="num" onclick="sortTbl('wifiTable',3,true)">Views '25</th>
        <th class="num" onclick="sortTbl('wifiTable',4,true)">Δ Views</th>
        <th class="num" onclick="sortTbl('wifiTable',5,true)">Active Users '24</th>
        <th class="num" onclick="sortTbl('wifiTable',6,true)">Active Users '25</th>
        <th class="num" onclick="sortTbl('wifiTable',7,true)">Δ Users</th>
      </tr></thead>
      <tbody id="wifiBody"></tbody>
    </table>
  </div>

</div>

<div class="footer">
  My Place Hotels of America &nbsp;·&nbsp; Property WiFi Analytics Dashboard &nbsp;·&nbsp; Google Analytics 4 &nbsp;·&nbsp; Confidential — Internal Use Only
</div>

<script>
{COMMON_JS}
{wifi_data_js}

function renderWifiTable() {{
  document.getElementById('wifiBody').innerHTML = WIFI_DATA.map(r=>{{
    return `<tr>
      <td class="name-cell">${{r.label}}</td>
      <td class="path-cell" style="font-size:9px;">${{r.path}}</td>
      <td class="num">${{r.v24.toLocaleString()}}</td>
      <td class="num">${{r.v25.toLocaleString()}}</td>
      ${{deltaCell(r.v24,r.v25,n=>Math.round(n).toLocaleString())}}
      <td class="num">${{r.u24.toLocaleString()}}</td>
      <td class="num">${{r.u25.toLocaleString()}}</td>
      ${{deltaCell(r.u24,r.u25,n=>Math.round(n).toLocaleString())}}
    </tr>`;
  }}).join('');
}}

window.addEventListener('load', ()=>{{
  drawTrendLine('cWifiSessions', WIFI_SESS24, WIFI_SESS25, fmtN);
  renderWifiTable();
}});
window.addEventListener('resize', ()=>{{
  drawTrendLine('cWifiSessions', WIFI_SESS24, WIFI_SESS25, fmtN);
}});
</script>
</body>
</html>"""

# ── Main ──────────────────────────────────────────────────────────────────
def main():
    print('Loading hotel names...')
    hotel_names, hotel_slug_index, hotel_id_map = load_hotel_names()
    print(f'  Loaded {len(hotel_names)} hotel name mappings, {len(hotel_id_map)} hotel ID mappings')

    logo = load_logo()
    print(f'  Logo: {"loaded" if logo else "NOT FOUND — using empty string"}')

    print('Parsing acquisition CSV...')
    acq_rows = read_csv(ACQ_CSV)
    mo_sess, mo_ke, mo_rev, wifi_mo_sess, ch_sess, ch_ke, ch_rev = parse_acq(acq_rows)

    print('Parsing landing page CSVs...')
    lp_rows = read_multi_csv(LP_CSVS)
    print(f'  Total LP rows: {len(lp_rows):,}')
    mo_pv, hotel, home_page, locations_index, site_pages, state_city, wifi = parse_lp(lp_rows, hotel_names, hotel_slug_index, hotel_id_map)

    print('Building data arrays...')
    d = build_js_arrays(mo_sess, mo_ke, mo_rev, mo_pv, wifi_mo_sess,
                        ch_sess, ch_ke, ch_rev,
                        hotel, home_page, locations_index, site_pages, state_city, wifi,
                        hotel_names, hotel_slug_index)

    k = calc_kpis(d)
    print(f'  Hotels: {len(d["HOTEL_DATA"])}, Hub pages: {len(d["HUB_DATA"])}, State/City pages: {len(d["SC_DATA"])}, WiFi properties: {len(d["WIFI_DATA"])}')
    print(f'  2025 sessions: {fmt_n(k["s25"])}, revenue: {fmt_r(k["r25"])}')

    print(f'Writing {WEB_OUT}...')
    html = build_website_html(d, k, logo)
    with open(WEB_OUT, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f'Writing {WIFI_OUT}...')
    html = build_wifi_html(d, k, logo)
    with open(WIFI_OUT, 'w', encoding='utf-8') as f:
        f.write(html)

    print('Done!')

if __name__ == '__main__':
    main()
