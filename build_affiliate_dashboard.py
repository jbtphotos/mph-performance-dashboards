#!/usr/bin/env python3
"""
Build / refresh the My Place Hotels Affiliate Channel — Impact Platform dashboard.

Reads the 8 monthly Impact CSV exports (Sep 2025 – Apr 2026), computes monthly
time-series and partner-level totals, and rewrites the inline JS data constants
plus header / KPI / footer text inside affiliate-performance-dashboard.html.
"""

import csv
import json
import os
import re
import sys

# Resolve BASE to the directory holding this script so it works whether run
# from the user's Mac or from a mounted workspace.
BASE = os.path.dirname(os.path.abspath(__file__))
HTML_PATH = os.path.join(BASE, "affiliate-performance-dashboard.html")

# (filename, display label, key) — chronological
MONTHS = [
    ("impact-affiliate-Sep-25.csv", "Sep 2025", "Sep25"),
    ("impact-affiliate-Oct-25.csv", "Oct 2025", "Oct25"),
    ("impact-affiliate-Nov-25.csv", "Nov 2025", "Nov25"),
    ("impact-affiliate-Dec-25.csv", "Dec 2025", "Dec25"),
    ("impact-affiliate-Jan-26.csv", "Jan 2026", "Jan26"),
    ("impact-affiliate-Feb-26.csv", "Feb 2026", "Feb26"),
    ("impact-affiliate-Mar-26.csv", "Mar 2026", "Mar26"),
    ("impact-affiliate-Apr-26.csv", "Apr 2026", "Apr26"),
]


def parse_num(s, default=0.0):
    s = (s or "").strip()
    if not s:
        return default
    try:
        return float(s)
    except ValueError:
        return default


def read_month(path):
    """Return list of partner-row dicts for a month CSV."""
    rows = []
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for r in reader:
            partner = r.get("Partner")
            if partner is None:
                continue
            # Preserve the partner string EXACTLY (no strip) so quirks like
            # double spaces and trailing periods aggregate correctly.
            if partner == "":
                continue
            rows.append({
                "partner": partner,
                "clicks": int(parse_num(r.get("Clicks"), 0)),
                "actions": int(parse_num(r.get("Actions"), 0)),
                "revenue": parse_num(r.get("Revenue"), 0.0),
                "action_cost": parse_num(r.get("Action Cost"), 0.0),
                "total_cost": parse_num(r.get("Total Cost"), 0.0),
            })
    return rows


def build_data():
    monthly = []
    # partner_name -> { totals + monthly dict }
    partners = {}

    for fname, label, key in MONTHS:
        path = os.path.join(BASE, fname)
        rows = read_month(path)

        m_clicks = sum(r["clicks"] for r in rows)
        m_actions = sum(r["actions"] for r in rows)
        m_revenue = round(sum(r["revenue"] for r in rows), 2)
        m_cost = round(sum(r["total_cost"] for r in rows), 2)
        active = sum(1 for r in rows if r["actions"] >= 1)
        total_partners = len(rows)

        monthly.append({
            "label": label,
            "key": key,
            "clicks": m_clicks,
            "actions": m_actions,
            "revenue": m_revenue,
            "cost": m_cost,
            "active": active,
            "total_partners": total_partners,
        })

        for r in rows:
            name = r["partner"]
            p = partners.setdefault(name, {
                "name": name,
                "clicks": 0,
                "actions": 0,
                "revenue": 0.0,
                "cost": 0.0,
                "monthly": {k: {"actions": 0, "revenue": 0.0} for _, _, k in MONTHS},
            })
            p["clicks"] += r["clicks"]
            p["actions"] += r["actions"]
            p["revenue"] += r["revenue"]
            p["cost"] += r["total_cost"]
            # If duplicate partner row in a single month (shouldn't happen but
            # safe to handle), accumulate within the month bucket as well.
            mbucket = p["monthly"][key]
            mbucket["actions"] += r["actions"]
            mbucket["revenue"] += r["revenue"]

    # Filter to partners with any clicks or actions across the 8 months.
    partner_list = []
    for p in partners.values():
        if p["clicks"] > 0 or p["actions"] > 0:
            partner_list.append(p)

    # Compute derived per-partner metrics and round.
    for p in partner_list:
        p["revenue"] = round(p["revenue"], 2)
        p["cost"] = round(p["cost"], 2)
        p["conv"] = round((p["actions"] / p["clicks"] * 100), 2) if p["clicks"] > 0 else 0.0
        p["roas"] = round((p["revenue"] / p["cost"]), 1) if p["cost"] > 0 else 0.0
        p["avg_booking"] = round((p["revenue"] / p["actions"]), 2) if p["actions"] > 0 else 0.0
        # Round monthly values
        for k in p["monthly"]:
            p["monthly"][k]["revenue"] = round(p["monthly"][k]["revenue"], 2)

    # Sort by total revenue desc, then actions desc, then clicks desc as a
    # tiebreaker for partners with $0 revenue.
    partner_list.sort(key=lambda p: (-p["revenue"], -p["actions"], -p["clicks"], p["name"]))

    return monthly, partner_list


def fmt_revenue_kpi(total):
    """Match the existing dashboard formatting: $X.XK or $X.XXM."""
    if total >= 1_000_000:
        return f"${total/1_000_000:.2f}M"
    return f"${total/1000:.1f}K"


def fmt_money(n):
    return f"${n:,.2f}"


def fmt_int(n):
    return f"{n:,}"


def main():
    monthly, partners = build_data()

    # ---- 8-month totals for KPI row ----
    tot_clicks = sum(m["clicks"] for m in monthly)
    tot_actions = sum(m["actions"] for m in monthly)
    tot_revenue = round(sum(m["revenue"] for m in monthly), 2)
    tot_cost = round(sum(m["cost"] for m in monthly), 2)
    overall_roas = (tot_revenue / tot_cost) if tot_cost > 0 else 0.0
    overall_conv = (tot_actions / tot_clicks * 100) if tot_clicks > 0 else 0.0
    avg_booking = (tot_revenue / tot_actions) if tot_actions > 0 else 0.0
    avg_cpa = (tot_cost / tot_actions) if tot_actions > 0 else 0.0

    # Active affiliates = partners with at least one action in the full window.
    active_total = sum(1 for p in partners if p["actions"] > 0)
    total_signed_up = len(partners)  # all partners that have any data

    # Top partner (Capital One expected)
    top = partners[0]
    top_share = (top["revenue"] / tot_revenue * 100) if tot_revenue > 0 else 0.0
    top_conv = top["conv"]

    # ---- Read existing HTML ----
    with open(HTML_PATH, "r", encoding="utf-8") as f:
        html = f.read()

    # ---- Replace inline data constants ----
    monthly_json = json.dumps(monthly, separators=(", ", ": "))
    partners_json = json.dumps(partners, separators=(", ", ": "))

    html = re.sub(
        r"const monthly = \[.*?\];",
        f"const monthly = {monthly_json};",
        html,
        count=1,
        flags=re.DOTALL,
    )
    html = re.sub(
        r"const partners = \[.*?\];",
        f"const partners = {partners_json};",
        html,
        count=1,
        flags=re.DOTALL,
    )

    # ---- Header subtitle ----
    html = html.replace(
        "Impact Platform · September 2025 – March 2026",
        "Impact Platform · September 2025 – April 2026",
    )

    # ---- "X months since launch" detail ----
    html = re.sub(
        r"\d+\s+months since launch",
        "8 months since launch",
        html,
    )

    # ---- Footer date range ----
    html = html.replace(
        "Sep 2025 – Mar 2026",
        "Sep 2025 – Apr 2026",
    )

    # ---- KPI Row updates ----
    # Total Revenue KPI
    html = re.sub(
        r'(<div class="kpi accent">\s*<div class="label">Total Revenue</div>\s*<div class="value">)[^<]+(</div>)',
        lambda m: f"{m.group(1)}{fmt_revenue_kpi(tot_revenue)}{m.group(2)}",
        html,
    )

    # Total Bookings KPI + avg booking detail
    html = re.sub(
        r'(<div class="kpi">\s*<div class="label">Total Bookings</div>\s*<div class="value">)[^<]+(</div>\s*<div class="detail">)[^<]+(</div>)',
        lambda m: f"{m.group(1)}{fmt_int(tot_actions)}{m.group(2)}${avg_booking:,.0f} avg booking value{m.group(3)}",
        html,
    )

    # Total Clicks KPI + conv rate detail
    html = re.sub(
        r'(<div class="kpi">\s*<div class="label">Total Clicks</div>\s*<div class="value">)[^<]+(</div>\s*<div class="detail">)[^<]+(</div>)',
        lambda m: f"{m.group(1)}{fmt_int(tot_clicks)}{m.group(2)}{overall_conv:.2f}% conversion rate{m.group(3)}",
        html,
    )

    # ROAS KPI + total cost detail
    html = re.sub(
        r'(<div class="kpi green">\s*<div class="label">ROAS</div>\s*<div class="value">)[^<]+(</div>\s*<div class="detail">)[^<]+(</div>)',
        lambda m: f"{m.group(1)}{overall_roas:.1f}x{m.group(2)}${tot_cost:,.0f} total cost{m.group(3)}",
        html,
    )

    # Avg CPA KPI
    html = re.sub(
        r'(<div class="kpi">\s*<div class="label">Avg CPA</div>\s*<div class="value">)[^<]+(</div>)',
        lambda m: f"{m.group(1)}${avg_cpa:,.2f}{m.group(2)}",
        html,
    )

    # Active Affiliates KPI + signed-up detail
    html = re.sub(
        r'(<div class="kpi">\s*<div class="label">Active Affiliates</div>\s*<div class="value">)[^<]+(</div>\s*<div class="detail">)[^<]+(</div>)',
        lambda m: f"{m.group(1)}{active_total}{m.group(2)}of {total_signed_up} total signed up{m.group(3)}",
        html,
    )

    # ---- Top Partner callout ----
    html = re.sub(
        r"(<h2>Top Partner — )[^<]+(</h2>)",
        lambda m: f"{m.group(1)}{top['name']}{m.group(2)}",
        html,
    )
    html = re.sub(
        r'(<div class="tp-name">)[^<]+(</div>)',
        lambda m: f"{m.group(1)}{top['name']}{m.group(2)}",
        html,
        count=1,
    )
    html = re.sub(
        r'(<div class="tp-stat">)[^<]+(</div>)',
        lambda m: f"{m.group(1)}{fmt_revenue_kpi(top['revenue'])}{m.group(2)}",
        html,
        count=1,
    )
    html = re.sub(
        r'(<div class="tp-sub">)[^<]+(</div>)',
        lambda m: (
            f"{m.group(1)}{fmt_int(top['actions'])} bookings · "
            f"{fmt_int(top['clicks'])} clicks · {top_conv}% conv rate{m.group(2)}"
        ),
        html,
        count=1,
    )
    html = re.sub(
        r'(<span class="growth up">)[^<]+(% of total revenue</span>)',
        lambda m: f"{m.group(1)}{top_share:.1f}{m.group(2)}",
        html,
        count=1,
    )

    with open(HTML_PATH, "w", encoding="utf-8") as f:
        f.write(html)

    # ---- Verification ----
    assert len(monthly) == 8, f"expected 8 months, got {len(monthly)}"
    for p in partners:
        assert set(p["monthly"].keys()) == {k for _, _, k in MONTHS}, \
            f"partner {p['name']!r} missing month keys"

    # Re-parse the inlined arrays to ensure JSON-validity
    m_match = re.search(r"const monthly = (\[.*?\]);", html, re.DOTALL)
    p_match = re.search(r"const partners = (\[.*?\]);", html, re.DOTALL)
    assert m_match and p_match, "could not relocate data constants after write"
    json.loads(m_match.group(1))
    json.loads(p_match.group(1))

    # Tag balance sanity
    open_div = len(re.findall(r"<div\b", html))
    close_div = len(re.findall(r"</div>", html))
    open_script = len(re.findall(r"<script\b", html))
    close_script = len(re.findall(r"</script>", html))

    size = os.path.getsize(HTML_PATH)

    print("=== Affiliate Dashboard Build Report ===")
    print(f"HTML size: {size:,} bytes")
    print(f"<div> open/close: {open_div}/{close_div}")
    print(f"<script> open/close: {open_script}/{close_script}")
    print(f"Months: {len(monthly)}  |  Partners (with activity): {len(partners)}")
    print()
    print("8-MONTH TOTALS")
    print(f"  Clicks:        {fmt_int(tot_clicks)}")
    print(f"  Actions:       {fmt_int(tot_actions)}")
    print(f"  Revenue:       {fmt_money(tot_revenue)}")
    print(f"  Cost:          {fmt_money(tot_cost)}")
    print(f"  ROAS:          {overall_roas:.2f}x")
    print(f"  Conv rate:     {overall_conv:.2f}%")
    print(f"  Avg booking:   {fmt_money(avg_booking)}")
    print(f"  Avg CPA:       {fmt_money(avg_cpa)}")
    print(f"  Active (>=1):  {active_total} of {total_signed_up} signed up")
    print()
    apr = monthly[-1]
    mar = monthly[-2]
    print("APRIL 2026 vs MARCH 2026")
    for label, key in [("Clicks", "clicks"), ("Actions", "actions"),
                       ("Revenue", "revenue"), ("Cost", "cost"),
                       ("Active", "active"), ("Total partners", "total_partners")]:
        a = apr[key]; mvm = mar[key]
        delta = a - mvm
        pct = (delta / mvm * 100) if mvm else 0.0
        arrow = "up" if delta >= 0 else "down"
        print(f"  {label:<14} Apr={a}  Mar={mvm}  Δ={delta:+}  ({pct:+.1f}%, {arrow})")
    print()
    print("TOP 5 PARTNERS BY 8-MONTH REVENUE")
    for i, p in enumerate(partners[:5], 1):
        share = (p["revenue"] / tot_revenue * 100) if tot_revenue else 0
        print(f"  {i}. {p['name']:<45} {fmt_money(p['revenue']):>14}  "
              f"{p['actions']:>4} bk  share={share:5.1f}%")
    print()
    # New partners that first appear in April: every prior month bucket is zero
    # actions AND zero revenue, AND Apr26 has clicks>0 or actions>0
    new_in_april = []
    apr_key = "Apr26"
    prior_keys = [k for _, _, k in MONTHS if k != apr_key]
    for p in partners:
        apr_b = p["monthly"][apr_key]
        if apr_b["actions"] == 0 and apr_b["revenue"] == 0:
            # The partner has the April bucket empty; but maybe they had clicks
            # this month with no revenue. Check the raw clicks in April CSV.
            pass
        prior_active = any(
            p["monthly"][k]["actions"] > 0 or p["monthly"][k]["revenue"] > 0
            for k in prior_keys
        )
        # We also want to know if they had ANY click activity prior.
        # The monthly bucket only stores actions+revenue; we can't tell about
        # clicks from `partners`. Re-derive from raw CSVs.
        # Cheaper path: track via dedicated dict during build_data. We'll
        # recompute by reading clicks per month directly here.
        # (kept simple: a partner is "new in April" if no prior month had
        # any actions or revenue.)
        if not prior_active and (apr_b["actions"] > 0 or apr_b["revenue"] > 0
                                 or p["clicks"] > 0 and all(
                                     p["monthly"][k]["actions"] == 0 and
                                     p["monthly"][k]["revenue"] == 0
                                     for k in prior_keys)):
            # The compound clicks check is noisy; do the stricter:
            pass

    # Cleaner re-derivation of "new in April":
    # A partner is "new in April" if April is the first month any of their
    # rows show up in any of the CSVs. We re-scan the CSV partner sets.
    seen_before_april = set()
    apr_partners = set()
    for fname, _, key in MONTHS:
        path = os.path.join(BASE, fname)
        with open(path, newline="", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            for r in reader:
                name = r.get("Partner")
                if not name:
                    continue
                if key == "Apr26":
                    apr_partners.add(name)
                else:
                    seen_before_april.add(name)
    truly_new = sorted(apr_partners - seen_before_april)
    print(f"NEW PARTNERS IN APRIL 2026 ({len(truly_new)})")
    for name in truly_new:
        # Find the partner record (if they had clicks/actions); some may be
        # zero-click rows we filtered out.
        rec = next((p for p in partners if p["name"] == name), None)
        if rec:
            apr_b = rec["monthly"]["Apr26"]
            print(f"  - {name}  |  Apr clicks={rec['clicks']}  "
                  f"actions={apr_b['actions']}  rev={fmt_money(apr_b['revenue'])}")
        else:
            print(f"  - {name}  (zero clicks & zero actions in April; filtered)")
    print()
    print("Build complete.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
