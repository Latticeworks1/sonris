#!/usr/bin/env python3
"""
SONRIS Data Portal bulk downloader.

Every page on the Louisiana DNR SONRIS data portal is an Oracle APEX 24.2
application whose data grids are Interactive Grid (IG) widgets.  Each IG
exposes a signed two-step download endpoint: a preparatory POST that validates
a per-session checksum and returns a transient file identifier, followed by a
GET that streams the actual payload.  This script automates that flow for any
endpoint in the portal without requiring a browser or proxy.

Usage
-----
  python3 sonris.py list
  python3 sonris.py download well-logs-ig
  python3 sonris.py download well-logs-ig --format XLSX --out wells.xlsx
  python3 sonris.py download-all --format CSV --outdir ./sonris_data/
  python3 sonris.py docs 7921 --type LOG --download --outdir ./docs/
  python3 sonris.py check serials.csv --info has_logs,doc_types
  python3 sonris.py depth --serials 7921,155058 --out depths.csv
  python3 sonris.py depth --csv serials.csv --out depths.csv
  python3 sonris.py map --csv serials.csv --filter-col has_logs --filter-val yes
  python3 sonris.py aor --lat 30.45 --lon -91.18 --radius 2.0
  python3 sonris.py search 7921
  python3 sonris.py search 7921 --aor
  python3 sonris.py search 7921 --aor --logs
  python3 sonris.py search 7921 --aor --logs --radius 5
  python3 sonris.py search --lat 30.45 --lon -91.18 --aor --logs
"""

import re
import sys
import csv
import json
import math
import time
import argparse
import requests
from pathlib import Path

BASE_URL = "https://sonlite.dnr.state.la.us/ords"
PORTAL   = f"{BASE_URL}/r/sonris_pub/sonris_data_portal"
AJAX     = f"{BASE_URL}/wwv_flow.ajax"
UA       = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/143.0.0.0 Safari/537.36"
)

# Complete endpoint manifest: slug -> (clear_param, category, description)
ENDPOINTS = {
    # Wells
    "well-logs-ig":                                      ("9085",  "wells",       "Well logs with associated electric logs and directional surveys"),
    "well-information":                                  ("2700",  "wells",       "Well information details"),
    "well-profile":                                      ("9000",  "wells",       "Well profile (multi-region)"),
    "well-casings":                                      ("2605",  "wells",       "Well casings"),
    "well-history-all-records":                          ("2620",  "wells",       "Well history — all records"),
    "well-history-by-operator":                          ("2905",  "wells",       "Well history by operator"),
    "well-information-details-by-operator-ojc":          ("2904",  "wells",       "Well information details by operator"),
    "wells-for-organization-by-field":                   ("2744",  "wells",       "Wells for organization by field"),
    "wells-and-usdw-by-coordinates":                     ("2216",  "wells",       "Wells and USDW by coordinates"),
    "wells-by-luw-code":                                 ("2366",  "wells",       "Wells by LUW code"),
    "well-completions-and-recompletions":                ("2690",  "wells",       "Well completions and recompletions"),
    "well-test-information":                             ("2670",  "wells",       "Well test information"),
    "work-permits":                                      ("2980",  "wells",       "Work permits"),
    "delinquent-wells-by-organization":                  ("2000",  "wells",       "Delinquent wells by organization"),
    "inactive-well-reports":                             ("2706",  "wells",       "Inactive well reports"),
    "orphan-wellsite-list":                              ("12050", "wells",       "Orphan wellsite list"),
    "active-financial-security-well-list":               ("2691",  "wells",       "Active financial security well list"),
    "scout-report-information":                          ("2655",  "wells",       "Scout report information"),
    "scout-summary-by-district":                         ("2666",  "wells",       "Scout summary by district"),
    "blackbooks-field-order-index":                      ("9012",  "wells",       "Black books field order index"),
    "well-control-incident-query":                       ("2236",  "wells",       "Well control incident query"),
    "severance-tax-relief-report":                       ("2600",  "wells",       "Severance tax relief report"),
    # Organizations
    "organization-search":                               ("9006",  "orgs",        "Organization search"),
    "organization-officers-contacts-search":             ("9015",  "orgs",        "Organization officers and contacts search"),
    "organization-address-listing":                      ("2010",  "orgs",        "Organization address listing"),
    "comprehensive-operator-information":                ("2228",  "orgs",        "Comprehensive operator information"),
    "active-licensed-drilling-companies":                ("2027",  "orgs",        "Active licensed drilling companies"),
    "operator-compliance-details":                       ("2382",  "orgs",        "Operator compliance details"),
    "consumers":                                         ("3012",  "orgs",        "Consumers"),
    # Fields and locations
    "field-listing":                                     ("9010",  "fields",      "Field listing (multi-region)"),
    "location-code-list":                                ("2292",  "fields",      "Location code list"),
    "drip-point":                                        ("2546",  "fields",      "Drip point"),
    # Production
    "luw-search":                                        ("2005",  "production",  "LUW search"),
    "field-production-by-year":                          ("2710",  "production",  "Field production by year"),
    "inception-to-date-field-production-by-operator":    ("11010", "production",  "Inception-to-date field production by operator"),
    "ogp-field-production-for-an-organization-by-month": ("2015",  "production",  "OGP field production for an organization by month"),
    "oil-and-gas-detail-production-by-month":            ("2381",  "production",  "Oil and gas detail production by month"),
    "ogp-production-by-luw":                             ("2739",  "production",  "OGP production by LUW"),
    "operator-oil-gas-production-for-a-year":            ("2550",  "production",  "Operator oil and gas production for a year"),
    "yearly-production-by-parish":                       ("2014",  "production",  "Yearly production by parish"),
    "statewide-monthly-production-for-a-year":           ("2001",  "production",  "Statewide monthly production for a year"),
    "cross-unit-well-production":                        ("2725",  "production",  "Cross-unit well production"),
    "underground-gas-storage":                           ("2715",  "production",  "Underground gas storage"),
    "imports-and-exports-by-company-by-month":           ("2031",  "production",  "Imports and exports by company by month"),
    # Transporters / Reporting
    "active-suspended-r4-authorizations-and-transporters":("2281", "reporting",   "Active/suspended R4 authorizations and transporters"),
    "current-authorized-transporters":                   ("3011",  "reporting",   "Current authorized transporters"),
    "r2-oil-transporter-report-by-date":                 ("2456",  "reporting",   "R2 oil transporter report by date"),
    "r2-oil-transporter-storer-information":             ("12006", "reporting",   "R2 oil transporter/storer information"),
    "r2p2-ogp-reconciliation-erros":                     ("12045", "reporting",   "R2P2 OGP reconciliation errors"),
    "r3-activity-report-product-listing":                ("2466",  "reporting",   "R3 activity report product listing"),
    "r3-oil-refinery":                                   ("12008", "reporting",   "R3 oil refinery"),
    "r3-refinery-summary-report":                        ("3017",  "reporting",   "R3 refinery summary report"),
    "r4-emergency-clearance-authorizations":             ("11035", "reporting",   "R4 emergency clearance authorizations"),
    "r5d-gas-disposition-ledger":                        ("11045", "reporting",   "R5D gas disposition ledger"),
    "r5t-natural-gas-transporters-information":          ("12007", "reporting",   "R5T natural gas transporters information"),
    "r6-gasoline-cycling-plant-information":             ("12035", "reporting",   "R6 gasoline cycling plant information"),
    "transporter-cross-check-listing":                   ("3020",  "reporting",   "Transporter cross-check listing"),
    # Compliance / Environmental
    "compliance-order-and-notice-query":                 ("2226",  "compliance",  "Compliance order and notice query"),
    "public-complaint-query":                            ("2231",  "compliance",  "Public complaint query"),
    "site-clearance":                                    ("2703",  "compliance",  "Site clearance"),
    "production-pit-query":                              ("2241",  "compliance",  "Production pit query"),
    "reserve-pit-query":                                 ("2246",  "compliance",  "Reserve pit query"),
    "administrative-applications":                       ("2704",  "compliance",  "Administrative applications"),
    "hydraulic-fracturing-and-drill-rig-supply-volumes": ("9090",  "compliance",  "Hydraulic fracturing and drill rig supply volumes"),
    # Facilities
    "e-p-commercial-facilities":                         ("9030",  "facilities",  "E&P commercial facilities"),
    "facilities":                                        ("2291",  "facilities",  "Facilities"),
    # UIC / Injection
    "injection-wells-by-operator":                       ("2547",  "uic",         "Injection wells by operator"),
    "injection-well-applications":                       ("9040",  "uic",         "Injection well applications"),
    "e-p-disposal-permits":                              ("9060",  "uic",         "E&P disposal permits"),
    "historic-uic10-annual-disposal":                    ("2735",  "uic",         "Historic UIC10 annual disposal"),
    "uic10-annual-disposal-injection-well-monitoring-report": ("6001","uic",      "UIC10 annual disposal injection well monitoring report"),
    "uic24-class-i-quarterly-reports":                   ("6021",  "uic",         "UIC24 class I quarterly reports"),
    "uic33-34-class-iii-daily-logs":                     ("6036",  "uic",         "UIC33/34 class III daily logs"),
    "class-i-manifests":                                 ("2740",  "uic",         "Class I manifests"),
    "uic-well-test-inspections":                         ("2038",  "uic",         "UIC well test inspections"),
    "salt-dome-cavern-well-sonar-mit-by-serial-number":  ("9055",  "uic",         "Salt dome cavern well sonar MIT by serial number"),
    "salt-sulfur-brine-summary-historic":                ("9050",  "uic",         "Salt/sulfur/brine summary historic"),
    # Water wells
    "water-wells-search":                                ("2086",  "water",       "Water wells search"),
    "water-wells-by-lat-long":                           ("2087",  "water",       "Water wells by lat/long"),
    # Leasing
    "current-lease-owners":                              ("1002",  "leasing",     "Current lease owners"),
    "lease-sale-tract-report":                           ("3001",  "leasing",     "Lease sale tract report"),
    "lease-by-owner":                                    ("9080",  "leasing",     "Lease by owner"),
    "view-payor-allocation-by-payor":                    ("9065",  "leasing",     "View payor allocation by payor"),
    "prospective-leaseholders":                          ("2036",  "leasing",     "Prospective leaseholders"),
    "state-lease-information":                           ("9075",  "leasing",     "State lease information"),
    "view-tract-bids":                                   ("9070",  "leasing",     "View tract bids"),
    "monthly-ldwf-lease-income":                         ("3014",  "leasing",     "Monthly LDWF lease income"),
    "active-financial-security-by-operator":             ("2692",  "leasing",     "Active financial security by operator"),
    # Reference codes
    "group-codes":                                       ("2024",  "codes",       "Group codes"),
    "luw-status-codes":                                  ("2025",  "codes",       "LUW status codes"),
    "luw-type-codes":                                    ("2026",  "codes",       "LUW type codes"),
    "operation-type-codes":                              ("2552",  "codes",       "Operation type codes"),
    "organization-type-codes":                           ("2554",  "codes",       "Organization type codes"),
    "parish-codes":                                      ("2537",  "codes",       "Parish codes"),
    "product-type-codes":                                ("2022",  "codes",       "Product type codes"),
    "well-status-codes":                                 ("2023",  "codes",       "Well status codes"),
    "uic-well-class-types":                              ("2580",  "codes",       "UIC well class types"),
    # Other
    "cross-unit-wells-and-associated-luws":              ("2286",  "other",       "Cross-unit wells and associated LUWs"),
    "legacy-lawsuits":                                   ("5061",  "other",       "Legacy lawsuits"),
    "outstanding-invoices":                              ("",      "other",       "Outstanding invoices (no clear param)"),
}


# ---------------------------------------------------------------------------
# Session
# ---------------------------------------------------------------------------

COOKIE_FILE = Path.home() / ".sonris_cookies.json"


def _load_cookies(s: requests.Session) -> None:
    if COOKIE_FILE.exists():
        try:
            data = json.loads(COOKIE_FILE.read_text())
            for domain, pairs in data.items():
                for name, value in pairs.items():
                    s.cookies.set(name, value, domain=domain)
        except Exception:
            pass


def _save_cookies(s: requests.Session) -> None:
    data: dict = {}
    for c in s.cookies:
        data.setdefault(c.domain or "", {})[c.name] = c.value
    try:
        COOKIE_FILE.write_text(json.dumps(data))
    except Exception:
        pass


def make_session() -> requests.Session:
    s = requests.Session()
    s.verify = True
    s.headers.update({"User-Agent": UA, "Accept-Language": "en-US,en;q=0.5"})
    _load_cookies(s)
    return s


# ---------------------------------------------------------------------------
# CAPTCHA handling
# ---------------------------------------------------------------------------

_BROWSER_CANDIDATES = [
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/Applications/Chromium.app/Contents/MacOS/Chromium",
    "/usr/bin/google-chrome",
    "/usr/bin/chromium-browser",
    "/usr/bin/chromium",
]


def _launch_browser(pw):
    exe = next((p for p in _BROWSER_CANDIDATES if Path(p).is_file()), None)
    kw: dict = {"headless": False}
    if exe:
        kw["executable_path"] = exe
    return pw.chromium.launch(**kw)


def _require_playwright():
    try:
        from playwright.sync_api import sync_playwright
        return sync_playwright
    except ImportError:
        sys.exit(
            "[!] playwright is required for CAPTCHA pages.\n"
            "    pip install playwright && playwright install chromium"
        )


def _port_cookies(ctx, s: requests.Session) -> None:
    try:
        for c in ctx.cookies():
            s.cookies.set(c["name"], c["value"], domain=c.get("domain", ""))
    except Exception:
        pass


def _has_captcha(html: str) -> bool:
    # P0_CALL_CAPTCHA is a page-0 item present on every APEX page with value="".
    # A live challenge sets it to a non-empty value.
    for pat in (r'id="P0_CALL_CAPTCHA"\s+value="([^"]+)"',
                r'value="([^"]+)"\s+id="P0_CALL_CAPTCHA"'):
        if re.search(pat, html):
            return True
    return False


def _is_dnr_captcha(data: bytes) -> bool:
    return b"captchaLoad" in data or b"grecaptcha" in data


def _browser_session(s: requests.Session, url: str, page_items: dict | None = None) -> str:
    """
    Open a visible browser at url for the APEX CAPTCHA flow.  Pre-fills any
    page_items by element ID.  Polls until the page no longer contains the
    CAPTCHA widget, then ports all cookies to s and re-fetches the page with
    requests to obtain fresh APEX state tokens.
    """
    sync_playwright = _require_playwright()
    print("[!] CAPTCHA detected — opening browser window.", file=sys.stderr)
    print("    Solve the CAPTCHA, then close the browser window to continue.", file=sys.stderr)

    with sync_playwright() as pw:
        browser = _launch_browser(pw)
        ctx = browser.new_context(user_agent=UA)
        page = ctx.new_page()
        page.goto(url)

        for item_id, value in (page_items or {}).items():
            try:
                page.fill(f"#{item_id}", value)
            except Exception:
                pass

        last_html = ""
        while True:
            try:
                last_html = page.content()
                page.wait_for_timeout(800)
            except Exception:
                break

        _port_cookies(ctx, s)
        try:
            browser.close()
        except Exception:
            pass

    _save_cookies(s)
    print("[*] Session ported — re-fetching page to get fresh APEX tokens.", file=sys.stderr)
    try:
        r = s.get(url, allow_redirects=True, timeout=30)
        r.raise_for_status()
        if not _has_captcha(r.text):
            return r.text
    except Exception:
        pass
    return last_html


def _solve_cloudflare(s: requests.Session, url: str) -> None:
    """
    Open a real browser at url so Cloudflare's JS challenge runs and sets the
    cf_clearance cookie.  The challenge is solved automatically by the browser
    engine; no user action is required.  Once cf_clearance appears in the
    browser cookie jar, port all cookies back to s.
    """
    sync_playwright = _require_playwright()
    print("[*] Waiting for Cloudflare clearance (automatic) ...", file=sys.stderr)

    with sync_playwright() as pw:
        browser = _launch_browser(pw)
        ctx = browser.new_context(user_agent=UA)
        page = ctx.new_page()
        page.goto(url)

        for _ in range(30):
            try:
                if any(c["name"] == "cf_clearance" for c in ctx.cookies()):
                    break
                page.wait_for_timeout(500)
            except Exception:
                break

        _port_cookies(ctx, s)
        try:
            browser.close()
        except Exception:
            pass

    _save_cookies(s)
    print("[*] Cloudflare clearance obtained.", file=sys.stderr)


def _solve_dnr_captcha(s: requests.Session, url: str) -> None:
    """
    Open a visible browser at the DNR file-service URL so the user can solve
    the reCAPTCHA enterprise challenge.  Polls until the reCAPTCHA widget
    disappears from the page (indicating the challenge was completed and the
    cookie was set), then ports all cookies back to s so subsequent requests
    carry the token without re-challenging.
    """
    sync_playwright = _require_playwright()
    print("[!] DNR file-service reCAPTCHA — opening browser window.", file=sys.stderr)
    print("    Solve the reCAPTCHA; the browser closes automatically when done.", file=sys.stderr)

    with sync_playwright() as pw:
        browser = _launch_browser(pw)
        ctx = browser.new_context(user_agent=UA)
        page = ctx.new_page()
        page.goto(url)

        while True:
            try:
                if "grecaptcha" not in page.content():
                    page.wait_for_timeout(600)
                    break
                page.wait_for_timeout(500)
            except Exception:
                break

        _port_cookies(ctx, s)
        try:
            browser.close()
        except Exception:
            pass

    _save_cookies(s)
    print("[*] DNR CAPTCHA solved — cookies saved.", file=sys.stderr)


# ---------------------------------------------------------------------------
# Shared Playwright helpers
# ---------------------------------------------------------------------------

def _pw_load_cookies(ctx) -> None:
    if COOKIE_FILE.exists():
        try:
            cdata = json.loads(COOKIE_FILE.read_text())
            for domain, pairs in cdata.items():
                for name, value in pairs.items():
                    ctx.add_cookies([{
                        "name": name, "value": value,
                        "domain": "sonlite.dnr.state.la.us", "path": "/",
                    }])
        except Exception:
            pass


def _pw_save_cookies(ctx) -> None:
    cd: dict = {}
    for c in ctx.cookies():
        cd.setdefault(c["domain"].lstrip("."), {})[c["name"]] = c["value"]
    COOKIE_FILE.write_text(json.dumps(cd))


def _pw_ig_csv(bpage, state: dict, region: dict) -> list[dict]:
    """Download one IG region as CSV using an in-browser fetch(), which avoids
    the APEX concurrent-session deadlock that blocks requests-based downloads."""
    reg_payload = {
        "pageItems": {},
        "regions": [{
            "id":             region["regionId"],
            "ajaxIdentifier": region["ajaxIdentifier"],
            "ajaxColumns":    region["ajaxColumns"],
            "view":           "grid",
            "download":       {"downloadCs": region["downloadCs"], "format": "CSV"},
            "reportId":       region["reportId"],
        }],
        "salt": state["pSalt"],
    }
    js = bpage.evaluate("""
        async ([ajaxUrl, flowId, stepId, instance, rp]) => {
            const enc = v => encodeURIComponent(v);
            const body = `p_flow_id=${enc(flowId)}&p_flow_step_id=${enc(stepId)}&p_instance=${enc(instance)}&p_json=${enc(JSON.stringify(rp))}`;
            const hdrs = {'Content-Type':'application/x-www-form-urlencoded; charset=UTF-8','X-Requested-With':'XMLHttpRequest'};
            const r2 = await fetch(ajaxUrl, {method:'POST', headers:hdrs, body});
            const reg = await r2.json();
            if (reg.error) return {error: reg.error};
            const fid = reg.regions[0].download.id;
            const sp = JSON.stringify({regions:[{id:rp.regions[0].id,ajaxIdentifier:rp.regions[0].ajaxIdentifier,downloadFileId:fid,reportId:rp.regions[0].reportId}],salt:rp.salt});
            const params = `p_flow_id=${enc(flowId)}&p_flow_step_id=${enc(stepId)}&p_instance=${enc(instance)}&p_debug=&p_json=${enc(sp)}`;
            const r3 = await fetch(`${ajaxUrl}?${params}`, {headers:{Accept:'*/*'}});
            return {csv: await r3.text()};
        }
    """, [AJAX, state["pFlowId"], state["pFlowStepId"], state["pInstance"], reg_payload])
    if js.get("error"):
        raise RuntimeError(js["error"])
    raw = js.get("csv", "")
    return list(csv.DictReader(raw.splitlines()))


# ---------------------------------------------------------------------------
# Page parsing
# ---------------------------------------------------------------------------

def fetch_page(s: requests.Session, slug: str, clear: str,
               page_items: dict | None = None) -> str:
    url = f"{PORTAL}/{slug}" + (f"?clear={clear}" if clear else "")
    r = s.get(url, allow_redirects=True, timeout=30)
    r.raise_for_status()
    html = r.text
    if _has_captcha(html):
        html = _browser_session(s, url, page_items)
    return html


def parse_apex_state(html: str) -> dict:
    """Extract hidden APEX form tokens from the page."""
    state = {}
    for fid in ("pFlowId", "pFlowStepId", "pInstance", "pSalt"):
        for pat in (rf'id="{fid}"\s+value="([^"]*)"',
                    rf'value="([^"]*)"\s+id="{fid}"'):
            m = re.search(pat, html)
            if m:
                state[fid] = m.group(1)
                break
    return state


def parse_page_items(html: str) -> list[str]:
    """Return all APEX page item IDs present on the page (P<n>_<NAME> pattern)."""
    seen, items = set(), []
    for m in re.finditer(r'<(?:input|select|textarea)[^>]+id="(P\d+_[A-Z_0-9]+)"', html):
        name = m.group(1)
        if name not in seen:
            seen.add(name)
            items.append(name)
    return items


def parse_ig_regions(html: str) -> list[dict]:
    """
    Extract each IG region config from the page JS.

    The config block layout in APEX 24.2 is:
      "ajaxColumns":"<signed>" , "regionId":"<id>", "regionAccTitle":"<title>",
      "regionStaticId":"...", "ajaxIdentifier":"<signed>", ...,
      "download":{"downloadCs":"<signed>", ...}

    We anchor on "ajaxColumns" (the first distinct token) and scan forward
    ~3000 chars to collect the rest.  Only regions that carry a downloadCs
    are returned.
    """
    def unescape(v: str) -> str:
        return v.replace(r"\u002F", "/")

    regions = []
    for m in re.finditer(r'"ajaxColumns"\s*:\s*"([^"]+)"', html):
        cols = unescape(m.group(1))
        chunk = html[m.start(): m.start() + 3000]

        rid_m   = re.search(r'"regionId"\s*:\s*"(\d+)"', chunk)
        title_m = re.search(r'"regionAccTitle"\s*:\s*"([^"]+)"', chunk)
        ajax_m  = re.search(r'"ajaxIdentifier"\s*:\s*"([^"]+)"', chunk)
        dlcs_m  = re.search(r'"downloadCs"\s*:\s*"([^"]+)"', chunk)
        rep_m   = re.search(r'"reportId"\s*:\s*"(\d+)"', chunk)

        if not (rid_m and ajax_m and dlcs_m):
            continue

        regions.append({
            "regionId":       rid_m.group(1),
            "regionTitle":    title_m.group(1) if title_m else rid_m.group(1),
            "ajaxIdentifier": unescape(ajax_m.group(1)),
            "downloadCs":     unescape(dlcs_m.group(1)),
            "ajaxColumns":    cols,
            "reportId":       rep_m.group(1) if rep_m else "",
        })

    # Deduplicate by regionId
    seen, unique = set(), []
    for r in regions:
        if r["regionId"] not in seen:
            seen.add(r["regionId"])
            unique.append(r)
    return unique


# ---------------------------------------------------------------------------
# Download
# ---------------------------------------------------------------------------

def _ajax_headers(slug: str, clear: str) -> dict:
    ref = f"{PORTAL}/{slug}" + (f"?clear={clear}" if clear else "")
    return {
        "Referer":          ref,
        "Origin":           "https://sonlite.dnr.state.la.us",
        "X-Requested-With": "XMLHttpRequest",
        "Content-Type":     "application/x-www-form-urlencoded; charset=UTF-8",
        "Accept":           "application/json, text/javascript, */*; q=0.01",
    }


def _register_download(
    s: requests.Session,
    state: dict,
    region: dict,
    fmt: str,
    slug: str,
    clear: str,
    page_items: dict | None = None,
) -> str:
    """Step 1 — register the download job; returns the transient file ID."""
    region_payload: dict = {
        "id":             region["regionId"],
        "ajaxIdentifier": region["ajaxIdentifier"],
        "ajaxColumns":    region["ajaxColumns"],
        "view":           "grid",
        "download": {
            "downloadCs": region["downloadCs"],
            "format":     fmt,
        },
    }
    if region.get("reportId"):
        region_payload["reportId"] = region["reportId"]
    if page_items:
        items_payload = {"itemsToSubmit": [{"n": k, "v": v} for k, v in page_items.items()]}
    else:
        items_payload = {}
    p_json = json.dumps({
        "pageItems": items_payload,
        "regions":   [region_payload],
        "salt":      state["pSalt"],
    })
    payload = {
        "p_flow_id":      state["pFlowId"],
        "p_flow_step_id": state["pFlowStepId"],
        "p_instance":     state["pInstance"],
        "p_json":         p_json,
    }
    r = s.post(AJAX, data=payload, headers=_ajax_headers(slug, clear), timeout=300)
    r.raise_for_status()
    try:
        body = r.json()
        return body["regions"][0]["download"]["id"]
    except Exception as exc:
        raise RuntimeError(
            f"Step-1 failed for region {region['regionId']}: {r.text[:400]}"
        ) from exc


def _fetch_file(
    s: requests.Session,
    state: dict,
    region: dict,
    file_id: str,
    slug: str,
    clear: str,
) -> bytes:
    """Step 2 — stream the actual file using the file ID from step 1."""
    fetch_region: dict = {
        "id":             region["regionId"],
        "ajaxIdentifier": region["ajaxIdentifier"],
        "downloadFileId": file_id,
    }
    if region.get("reportId"):
        fetch_region["reportId"] = region["reportId"]
    p_json = json.dumps({
        "regions": [fetch_region],
        "salt":    state["pSalt"],
    })
    params = {
        "p_flow_id":      state["pFlowId"],
        "p_flow_step_id": state["pFlowStepId"],
        "p_instance":     state["pInstance"],
        "p_debug":        "",
        "p_json":         p_json,
    }
    ref = f"{PORTAL}/{slug}" + (f"?clear={clear}" if clear else "")
    r = s.get(AJAX, params=params,
              headers={"Referer": ref, "Accept": "*/*", "User-Agent": UA},
              timeout=600)
    r.raise_for_status()
    return r.content


def download_region(
    s: requests.Session,
    state: dict,
    region: dict,
    slug: str,
    clear: str,
    fmt: str = "CSV",
    page_items: dict | None = None,
) -> bytes:
    file_id = _register_download(s, state, region, fmt, slug, clear, page_items)
    return _fetch_file(s, state, region, file_id, slug, clear)


# ---------------------------------------------------------------------------
# CLI helpers
# ---------------------------------------------------------------------------

def cmd_list(args):
    by_cat = {}
    for slug, (clear, cat, desc) in ENDPOINTS.items():
        by_cat.setdefault(cat, []).append((slug, desc))
    for cat in sorted(by_cat):
        print(f"\n{cat.upper()}")
        for slug, desc in sorted(by_cat[cat]):
            print(f"  {slug:55s}  {desc}")


def cmd_fields(args):
    slug = args.slug
    if slug not in ENDPOINTS:
        print(f"Unknown endpoint '{slug}'. Run 'list' to see available endpoints.")
        sys.exit(1)
    clear, _, desc = ENDPOINTS[slug]
    s = make_session()
    html = fetch_page(s, slug, clear)
    items = parse_page_items(html)
    if not items:
        print(f"No filter items found for '{slug}'.")
    else:
        print(f"Filter items for '{slug}':")
        for item in items:
            print(f"  {item}")


def cmd_download(args):
    slug = args.slug
    if slug not in ENDPOINTS:
        print(f"Unknown endpoint '{slug}'. Run 'list' to see available endpoints.")
        sys.exit(1)

    clear, _, desc = ENDPOINTS[slug]
    s = make_session()

    # parse --filter KEY=VALUE pairs
    page_items: dict = {}
    for f in (args.filter or []):
        if "=" not in f:
            print(f"[!] --filter must be KEY=VALUE, got: {f}", file=sys.stderr)
            sys.exit(1)
        k, v = f.split("=", 1)
        page_items[k] = v

    print(f"[*] {desc}", file=sys.stderr)
    print("[*] Fetching page ...", file=sys.stderr)
    html  = fetch_page(s, slug, clear, page_items=page_items)
    state = parse_apex_state(html)
    if not state.get("pInstance"):
        print("[!] Failed to extract APEX session tokens.", file=sys.stderr)
        sys.exit(1)

    regions = parse_ig_regions(html)
    if not regions:
        print("[!] No downloadable IG regions found on this page.", file=sys.stderr)
        sys.exit(1)

    if getattr(args, "debug", False):
        trunc = {k: (v[:70] + "...") if len(v) > 70 else v for k, v in state.items()}
        print(json.dumps({"state": trunc, "regions": regions}, indent=2), file=sys.stderr)

    if not page_items:
        available = parse_page_items(html)
        if available:
            print(f"[~] No filters set. Available filter items: {available}", file=sys.stderr)
            print("[~] Results may be empty. Use --filter KEY=VALUE to query data.", file=sys.stderr)

    print(f"[*] Found {len(regions)} region(s): {[r['regionTitle'] for r in regions]}", file=sys.stderr)

    for region in regions:
        safe_title = re.sub(r'[^\w]+', '_', region['regionTitle']).strip('_').lower()
        if args.out and len(regions) == 1:
            out_path = Path(args.out)
        elif args.out:
            out_path = Path(args.out).with_stem(f"{Path(args.out).stem}_{safe_title}")
        else:
            ext = args.format.lower()
            out_path = Path(f"{slug}__{safe_title}.{ext}")

        print(f"[*] Downloading '{region['regionTitle']}' ...", file=sys.stderr)
        data = download_region(s, state, region, slug, clear, fmt=args.format, page_items=page_items)
        out_path.write_bytes(data)
        print(f"[+] {out_path}  ({len(data):,} bytes)", file=sys.stderr)


def cmd_download_all(args):
    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    fmt = args.format
    ext = fmt.lower()

    failed = []
    for slug, (clear, cat, desc) in ENDPOINTS.items():
        if not clear:
            print(f"[~] Skipping '{slug}' (no clear param)", file=sys.stderr)
            continue

        s = make_session()
        try:
            html   = fetch_page(s, slug, clear)
            state  = parse_apex_state(html)
            if not state.get("pInstance"):
                raise RuntimeError("No APEX session tokens")

            regions = parse_ig_regions(html)
            if not regions:
                print(f"[~] {slug}: no downloadable regions", file=sys.stderr)
                continue

            for region in regions:
                safe_title = re.sub(r'[^\w]+', '_', region['regionTitle']).strip('_').lower()
                out_path   = outdir / f"{slug}__{safe_title}.{ext}"
                data       = download_region(s, state, region, slug, clear, fmt=fmt)
                out_path.write_bytes(data)
                print(f"[+] {out_path}  ({len(data):,} bytes)", file=sys.stderr)

            time.sleep(0.5)   # be polite to the server between pages

        except Exception as e:
            print(f"[!] {slug}: {e}", file=sys.stderr)
            failed.append(slug)

    if failed:
        print(f"\nFailed endpoints: {failed}", file=sys.stderr)


# ---------------------------------------------------------------------------
# Document fetching (dnrservices redirectUrl system)
# ---------------------------------------------------------------------------

DOC_BASE        = "https://sonlite.dnr.state.la.us"
DOC_SEARCH      = f"{DOC_BASE}/ords/r/sonris_pub/document_access/finddocuments"
DOC_ALL_RESULTS = f"{DOC_BASE}/ords/r/sonris_pub/document_access/findalldocumentsresults"
DOC_DL_BASE     = f"{DOC_BASE}/dnrservices/redirectUrl.jsp"

# Known search indexes accepted by the finddocuments page
DOC_INDEXES = {
    "well":       "xwellserialnumber",
    "operator":   "xoperatorcode",
    "field":      "xfieldcode",
    "docname":    "xdocname",
    "doctype":    "xdocumenttype",
}


def _parse_doc_panel(html: str) -> list[dict]:
    """Parse an IR panel from finddocuments or findalldocumentsresults.

    Column layout in findalldocumentsresults (0-indexed):
      3  Content Id (docname)   4  Create Date   5  Document Type
      13 Description
    Columns are extracted by fixed position from the full <tr>, so empty
    cells are preserved and positional mapping stays correct.
    """
    def _clean(td: str) -> str:
        t = re.sub(r'<[^>]+>', '', td)
        t = re.sub(r'&#x2F;', '/', t)
        t = re.sub(r'&#x27;', "'", t)
        return re.sub(r'&#x[0-9a-fA-F]+;', '', t).strip()

    docs, seen = [], set()
    for m in re.finditer(r'dDocname=([^&"]+)&showInline=True', html):
        docname = m.group(1)
        if docname in seen:
            continue
        seen.add(docname)
        tr_start = html.rfind('<tr', 0, m.start())
        tr_end   = html.find('</tr>', m.start())
        if tr_start < 0 or tr_end < 0:
            continue
        cells = [_clean(td) for td in
                 re.findall(r'<td[^>]*>(.*?)</td>', html[tr_start: tr_end + 5], re.DOTALL)]
        docs.append({
            "docname":     docname,
            "create_date": cells[4]  if len(cells) > 4  else "",
            "doc_type":    cells[5]  if len(cells) > 5  else "",
            "description": cells[13] if len(cells) > 13 else "",
        })
    return docs


def fetch_all_doc_list(s: requests.Session, idx: str, val: str) -> list[dict]:
    """
    Navigate finddocuments → findalldocumentsresults to get the comprehensive
    document listing.  finddocuments sets P22_QUERY on the server when it links
    to findalldocumentsresults, so following that link returns results scoped to
    the same idx/val without requiring a cs checksum.
    """
    seed = s.get(DOC_SEARCH, params={"idx": idx, "val": val},
                 allow_redirects=True, timeout=30)
    if seed.status_code == 403:
        print("[!] 403 on document_access — opening browser for Cloudflare clearance.",
              file=sys.stderr)
        _solve_cloudflare(s, DOC_SEARCH)
        seed = s.get(DOC_SEARCH, params={"idx": idx, "val": val},
                     allow_redirects=True, timeout=30)
    seed.raise_for_status()

    pi_m = re.search(r'findalldocumentsresults[/\\]?(\d+)', seed.text)
    if not pi_m:
        return _parse_doc_panel(seed.text)

    r = s.get(f"{DOC_ALL_RESULTS}/{pi_m.group(1)}", allow_redirects=True, timeout=60)
    r.raise_for_status()
    return _parse_doc_panel(r.text)


def download_doc(s: requests.Session, docname: str, native: bool = False) -> tuple[bytes, str]:
    """
    Download a single document.  native=True returns the original TIFF;
    native=False returns the PDF conversion.  Returns (bytes, filename).
    If the DNR file service responds with a reCAPTCHA challenge the browser
    is opened once to solve it, after which the request is retried using the
    ported session cookie.
    """
    params = {"dDocname": docname, "showInline": "True"}
    if native:
        params["nativeFile"] = "True"
    url = f"{DOC_DL_BASE}?{'&'.join(f'{k}={v}' for k, v in params.items())}"

    for attempt in range(2):
        r = s.get(DOC_DL_BASE, params=params, allow_redirects=True, timeout=60)
        r.raise_for_status()
        if _is_dnr_captcha(r.content):
            if attempt == 0:
                _solve_dnr_captcha(s, url)
                continue
            raise RuntimeError(f"DNR reCAPTCHA still present after solve attempt for {docname}")
        break

    cd = r.headers.get("content-disposition", "")
    fname_m = re.search(r'filename="([^"]+)"', cd, re.IGNORECASE)
    fname = fname_m.group(1) if fname_m else f"{docname}.{'tif' if native else 'pdf'}"
    return r.content, fname


def cmd_docs(args):
    native = args.format == "tiff"
    outdir = Path(args.outdir) if args.outdir else None
    if outdir:
        outdir.mkdir(parents=True, exist_ok=True)

    s = make_session()

    if not args.val:
        print("[!] Provide a search value.", file=sys.stderr)
        sys.exit(1)
    idx_key = args.idx
    if idx_key not in DOC_INDEXES:
        print(f"Unknown index '{idx_key}'. Choose from: {list(DOC_INDEXES)}")
        sys.exit(1)
    idx   = DOC_INDEXES[idx_key]
    val   = args.val
    slug  = val
    label = f"{idx_key}={val}"
    print(f"[*] Fetching documents ({label}) via findalldocumentsresults ...", file=sys.stderr)
    docs = fetch_all_doc_list(s, idx, val)

    if not docs:
        print("[~] No documents found.", file=sys.stderr)
        return

    if args.type:
        filt = args.type.upper()
        docs = [d for d in docs if filt in d["doc_type"].upper()]

    print(f"[*] Found {len(docs)} document(s):", file=sys.stderr)
    for d in docs:
        desc = f"  {d['description']}" if d.get("description") else ""
        print(f"    {d['docname']:12s}  {d['doc_type']:30s}  {d['create_date']}{desc}", file=sys.stderr)

    if not args.download:
        return

    failed = []
    for d in docs:
        try:
            data, fname = download_doc(s, d["docname"], native=native)
        except Exception as e:
            print(f"[!] {d['docname']} ({d['doc_type']}): {e}", file=sys.stderr)
            failed.append(d["docname"])
            continue
        safe_type = re.sub(r'[^\w]+', '_', d['doc_type']).strip('_').lower()
        if outdir:
            out_path = outdir / f"{slug}__{safe_type}__{fname}"
        else:
            out_path = Path(f"{slug}__{safe_type}__{fname}")
        out_path.write_bytes(data)
        print(f"[+] {out_path}  ({len(data):,} bytes)", file=sys.stderr)
        time.sleep(0.5)
    if failed:
        print(f"[~] Failed: {failed}", file=sys.stderr)


# ---------------------------------------------------------------------------
# CSV batch check
# ---------------------------------------------------------------------------

INFO_FIELDS: dict = {
    "has_logs":         lambda docs: "yes" if any("WELL LOG" in d["doc_type"].upper() for d in docs) else "no",
    "log_count":        lambda docs: str(sum(1 for d in docs if "WELL LOG" in d["doc_type"].upper())),
    "log_descriptions": lambda docs: "; ".join(
                            d["description"] for d in docs
                            if "WELL LOG" in d["doc_type"].upper() and d.get("description")
                        ),
    "has_docs":         lambda docs: "yes" if docs else "no",
    "doc_count":        lambda docs: str(len(docs)),
    "doc_types":        lambda docs: "; ".join(sorted({d["doc_type"] for d in docs})),
}


def cmd_check(args):
    fields = [f.strip() for f in args.info.split(",")]
    bad = [f for f in fields if f not in INFO_FIELDS]
    if bad:
        print(f"[!] Unknown info field(s): {bad}. Choose from: {list(INFO_FIELDS)}", file=sys.stderr)
        sys.exit(1)

    with open(args.csv_file, newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.reader(fh))

    if not rows:
        print("[!] CSV is empty.", file=sys.stderr)
        sys.exit(1)

    # Detect header: first cell is non-numeric → treat as header row
    has_header = not rows[0][0].strip().lstrip("-").isdigit()
    serial_col = 0
    if args.serial_col.isdigit():
        serial_col = int(args.serial_col)
    elif has_header:
        names = [c.strip().lower() for c in rows[0]]
        if args.serial_col.lower() in names:
            serial_col = names.index(args.serial_col.lower())
        else:
            print(f"[!] Column '{args.serial_col}' not found. Header: {rows[0]}", file=sys.stderr)
            sys.exit(1)

    out: list = []
    if has_header:
        out.append(rows[0] + fields)
        data = rows[1:]
    else:
        data = rows

    s = make_session()
    idx = DOC_INDEXES["well"]
    total = sum(1 for r in data if r and r[serial_col].strip())
    done  = 0

    for row in data:
        if not row or not row[serial_col].strip():
            out.append(row)
            continue
        serial = row[serial_col].strip()
        try:
            docs   = fetch_all_doc_list(s, idx, serial)
            values = [INFO_FIELDS[f](docs) for f in fields]
        except Exception as e:
            print(f"[!] {serial}: {e}", file=sys.stderr)
            values = ["error"] * len(fields)
        out.append(list(row) + values)
        done += 1
        print(f"    [{done}/{total}] {serial}: {dict(zip(fields, values))}", file=sys.stderr)
        time.sleep(0.5)

    out_path = Path(args.out) if args.out else Path(args.csv_file).with_stem(Path(args.csv_file).stem + "_results")
    with open(out_path, "w", newline="", encoding="utf-8") as fh:
        csv.writer(fh).writerows(out)
    print(f"[+] {out_path}", file=sys.stderr)


# ---------------------------------------------------------------------------
# Depth / well-info batch fetch (Playwright-assisted)
# ---------------------------------------------------------------------------

def _extract_well_info_tokens(html: str) -> tuple[dict, dict]:
    """Return (apex_state, region) from a well-information page HTML."""
    state = parse_apex_state(html)

    # pPageItemsProtected (the signed page-item checksum APEX requires in DA calls)
    m = re.search(r'id="pPageItemsProtected"\s+value="([^"]*)"', html)
    if not m:
        m = re.search(r'value="([^"]*)"\s+id="pPageItemsProtected"', html)
    state["pPageItemsProtected"] = m.group(1).replace("&#x2F;", "/") if m else ""

    # DA ajaxIdentifier that builds P2700_QUERY_CRITERIA
    pos_start = html.find("gEventList =")
    pos_end   = html.find("apex.da.initDaEventList()", pos_start)
    block     = html[pos_start: pos_end if pos_end > 0 else pos_start + 30000]
    da_ajax   = ""
    for mm in re.finditer(r'"ajaxIdentifier"\s*:\s*"([^"]+)"', block):
        chunk = block[mm.start(): mm.start() + 800]
        a2 = re.search(r'"attribute02"\s*:\s*"([^"]*)"', chunk)
        if a2 and "P2700_QUERY_CRITERIA" in a2.group(1):
            da_ajax = mm.group(1).replace("\\u002F", "/")
            break
    state["daAjaxQC"] = da_ajax

    regions = parse_ig_regions(html)
    region  = regions[0] if regions else {}
    return state, region


def cmd_depth(args):
    serials: list[str] = []

    if args.serials:
        serials = [x.strip() for x in args.serials.split(",") if x.strip()]
    elif args.csv_file:
        with open(args.csv_file, newline="", encoding="utf-8-sig") as fh:
            rows = list(csv.reader(fh))
        if not rows:
            print("[!] CSV is empty.", file=sys.stderr)
            sys.exit(1)
        has_header = not rows[0][0].strip().lstrip("-").isdigit()
        serial_col = 0
        if args.serial_col.isdigit():
            serial_col = int(args.serial_col)
        elif has_header:
            names = [c.strip().lower() for c in rows[0]]
            if args.serial_col.lower() in names:
                serial_col = names.index(args.serial_col.lower())
            else:
                print(f"[!] Column '{args.serial_col}' not found.", file=sys.stderr)
                sys.exit(1)
        data = rows[1:] if has_header else rows
        if args.filter_col and args.filter_val and has_header:
            names = [c.strip().lower() for c in rows[0]]
            if args.filter_col.lower() in names:
                fc = names.index(args.filter_col.lower())
                data = [r for r in data if r and len(r) > fc and r[fc].strip().lower() == args.filter_val.lower()]
        serials = [r[serial_col].strip() for r in data if r and len(r) > serial_col and r[serial_col].strip()]

    if not serials:
        print("[!] No serial numbers found.", file=sys.stderr)
        sys.exit(1)

    out_path = Path(args.out) if args.out else Path("well_depths.csv")
    print(f"[*] Fetching well info for {len(serials)} serial(s) → {out_path}", file=sys.stderr)

    sync_playwright = _require_playwright()

    out_rows: list[dict] = []
    fieldnames = ["serial", "well_name", "operator", "parish", "status",
                  "measured_depth", "latitude", "longitude",
                  "api_num", "spud_date", "permit_date"]

    # Open output CSV for incremental writes
    with open(out_path, "w", newline="", encoding="utf-8") as out_fh:
        writer = csv.DictWriter(out_fh, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        out_fh.flush()

        with sync_playwright() as pw:
            browser = pw.chromium.launch(headless=True)
            ctx = browser.new_context(user_agent=UA)
            if COOKIE_FILE.exists():
                try:
                    cdata = json.loads(COOKIE_FILE.read_text())
                    for domain, pairs in cdata.items():
                        for name, value in pairs.items():
                            ctx.add_cookies([{
                                "name": name, "value": value,
                                "domain": "sonlite.dnr.state.la.us", "path": "/",
                            }])
                except Exception:
                    pass

            bpage = ctx.new_page()
            bpage.goto(f"{PORTAL}/well-information?clear=2700",
                       wait_until="networkidle", timeout=30000)

            # Extract initial tokens to validate page loaded correctly
            initial_html = bpage.content()
            state, region = _extract_well_info_tokens(initial_html)

            if not state.get("pInstance") or not region:
                print("[!] Could not extract APEX tokens from page.", file=sys.stderr)
                sys.exit(1)
            if not state.get("daAjaxQC"):
                print("[!] Could not find DA ajaxIdentifier for QUERY_CRITERIA.", file=sys.stderr)
                sys.exit(1)

            search_btn = bpage.locator('[id="B275295944573754186"]')
            well_info_url = f"{PORTAL}/well-information"

            for i, serial in enumerate(serials, 1):
                try:
                    # Navigate with ?clear=2700 so the search form is reset to
                    # its default expanded state (the Search DA collapses it and
                    # APEX session stores that collapsed state between loads)
                    bpage.goto(f"{well_info_url}?clear=2700",
                               wait_until="networkidle", timeout=20000)
                    bpage.fill("#P2700_WELL_SERIAL_NUM", serial)
                    search_btn.first.click()
                    bpage.wait_for_load_state("networkidle", timeout=15000)
                    bpage.wait_for_timeout(500)

                    # Capture fresh tokens from live DOM
                    fresh_html    = bpage.content()
                    state, region = _extract_well_info_tokens(fresh_html)

                    # Run both AJAX calls inside the page's JS context so the
                    # browser's session cookies are used automatically and
                    # there is no concurrent-session contention
                    reg_payload = {
                        "pageItems": {},
                        "regions": [{
                            "id":             region["regionId"],
                            "ajaxIdentifier": region["ajaxIdentifier"],
                            "ajaxColumns":    region["ajaxColumns"],
                            "view":           "grid",
                            "download":       {"downloadCs": region["downloadCs"], "format": "CSV"},
                            "reportId":       region["reportId"],
                        }],
                        "salt": state["pSalt"],
                    }
                    js_result = bpage.evaluate("""
                        async ([ajaxUrl, flowId, stepId, instance, regPayload]) => {
                            const enc = v => encodeURIComponent(v);
                            const body = `p_flow_id=${enc(flowId)}&p_flow_step_id=${enc(stepId)}&p_instance=${enc(instance)}&p_json=${enc(JSON.stringify(regPayload))}`;
                            const hdrs = {'Content-Type':'application/x-www-form-urlencoded; charset=UTF-8','X-Requested-With':'XMLHttpRequest','Referer': ajaxUrl.replace('/wwv_flow.ajax','')};
                            const r2 = await fetch(ajaxUrl, {method:'POST', headers:hdrs, body});
                            const reg = await r2.json();
                            if (reg.error) return {error: reg.error};
                            const fileId = reg.regions[0].download.id;
                            const streamPayload = JSON.stringify({regions:[{id:regPayload.regions[0].id,ajaxIdentifier:regPayload.regions[0].ajaxIdentifier,downloadFileId:fileId,reportId:regPayload.regions[0].reportId}],salt:regPayload.salt});
                            const params = `p_flow_id=${enc(flowId)}&p_flow_step_id=${enc(stepId)}&p_instance=${enc(instance)}&p_debug=&p_json=${enc(streamPayload)}`;
                            const r3 = await fetch(`${ajaxUrl}?${params}`, {headers:{Accept:'*/*'}});
                            const text = await r3.text();
                            return {csv: text};
                        }
                    """, [AJAX, state["pFlowId"], state["pFlowStepId"], state["pInstance"],
                          reg_payload])

                    if js_result.get("error"):
                        raise RuntimeError(js_result["error"])
                    raw  = js_result.get("csv", "")
                    rows = list(csv.DictReader(raw.splitlines()))
                    if not rows:
                        print(f"    [{i}/{len(serials)}] {serial}: no data", file=sys.stderr)
                        writer.writerow({"serial": serial})
                        out_fh.flush()
                        continue

                    row = rows[0]
                    out_row = {
                        "serial":        serial,
                        "well_name":     row.get("Well Name", ""),
                        "operator":      row.get("Operator Name", ""),
                        "parish":        row.get("Parish Name", ""),
                        "status":        row.get("Well Status Code Description", ""),
                        "measured_depth":row.get("Measured Depth", ""),
                        "latitude":      row.get("Latitude", ""),
                        "longitude":     row.get("Longitude", ""),
                        "api_num":       row.get("API Num", ""),
                        "spud_date":     row.get("Spud Date", ""),
                        "permit_date":   row.get("Permit Date", ""),
                    }
                    writer.writerow(out_row)
                    out_fh.flush()
                    print(
                        f"    [{i}/{len(serials)}] {serial}  "
                        f"depth={out_row['measured_depth']}  "
                        f"lat={out_row['latitude'][:10] if out_row['latitude'] else ''}  "
                        f"lon={out_row['longitude'][:11] if out_row['longitude'] else ''}",
                        file=sys.stderr,
                    )
                except Exception as exc:
                    print(f"    [{i}/{len(serials)}] {serial}: {exc}", file=sys.stderr)
                    writer.writerow({"serial": serial})
                    out_fh.flush()

            # Save cookies for next run
            cd: dict = {}
            for c in ctx.cookies():
                cd.setdefault(c["domain"].lstrip("."), {})[c["name"]] = c["value"]
            COOKIE_FILE.write_text(json.dumps(cd))
            browser.close()

    print(f"[+] {out_path}", file=sys.stderr)


# ---------------------------------------------------------------------------
# Map generation
# ---------------------------------------------------------------------------

def _generate_map_html(wells: list[dict]) -> str:
    lats = [w["lat"] for w in wells]
    lons = [w["lon"] for w in wells]

    marker_lines = []
    for w in wells:
        def js(v: str) -> str:
            return str(v).replace("\\", "\\\\").replace('"', '\\"').replace("\n", " ")
        serial    = js(w["serial"])
        well_name = js(w.get("well_name", ""))
        operator  = js(w.get("operator", ""))
        parish    = js(w.get("parish", ""))
        status    = js(w.get("status", ""))
        lat, lon  = w["lat"], w["lon"]
        popup     = f"{serial}<br>{well_name}<br>{operator}<br>{parish}<br>{status}"
        marker_lines.append(
            f'L.marker([{lat},{lon}]).addTo(map)'
            f'.bindPopup("{popup}")'
            f'.bindTooltip("{serial}",{{permanent:true,direction:"top",offset:[0,-8]}});'
        )

    markers_js   = "\n".join(marker_lines)
    bounds_arr   = ",".join(f"[{w['lat']},{w['lon']}]" for w in wells)
    tile_url     = "https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png"
    attribution  = "© OpenStreetMap contributors"

    tile_line    = f"L.tileLayer('{tile_url}',{{attribution:'{attribution}'}}).addTo(map);"
    bounds_line  = f"map.fitBounds([{bounds_arr}],{{padding:[40,40]}});"

    return (
        "<!DOCTYPE html>\n<html>\n<head>\n"
        '<meta charset="utf-8">\n'
        "<title>SONRIS Wells</title>\n"
        '<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>\n'
        '<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>\n'
        "<style>html,body,#map{height:100%;margin:0;padding:0}</style>\n"
        "</head>\n<body>\n"
        '<div id="map"></div>\n'
        "<script>\n"
        "var map = L.map('map');\n"
        f"{tile_line}\n"
        f"{markers_js}\n"
        f"{bounds_line}\n"
        "</script>\n</body>\n</html>"
    )


def cmd_map(args):
    serials: list[str] = []

    if args.serials:
        serials = [x.strip() for x in args.serials.split(",") if x.strip()]
    elif args.csv_file:
        with open(args.csv_file, newline="", encoding="utf-8-sig") as fh:
            rows = list(csv.reader(fh))
        if not rows:
            print("[!] CSV is empty.", file=sys.stderr)
            sys.exit(1)
        has_header = not rows[0][0].strip().lstrip("-").isdigit()
        serial_col = 0
        if args.serial_col.isdigit():
            serial_col = int(args.serial_col)
        elif has_header:
            names = [c.strip().lower() for c in rows[0]]
            if args.serial_col.lower() in names:
                serial_col = names.index(args.serial_col.lower())
            else:
                print(f"[!] Column '{args.serial_col}' not found. Header: {rows[0]}", file=sys.stderr)
                sys.exit(1)

        data = rows[1:] if has_header else rows

        if args.filter_col and args.filter_val and has_header:
            names = [c.strip().lower() for c in rows[0]]
            if args.filter_col.lower() in names:
                fc = names.index(args.filter_col.lower())
                data = [r for r in data if r and len(r) > fc and r[fc].strip().lower() == args.filter_val.lower()]

        serials = [r[serial_col].strip() for r in data if r and len(r) > serial_col and r[serial_col].strip()]
    else:
        print("[!] Provide --serials or --csv.", file=sys.stderr)
        sys.exit(1)

    if not serials:
        print("[!] No serial numbers found.", file=sys.stderr)
        sys.exit(1)

    print(f"[*] Fetching coordinates for {len(serials)} well(s) ...", file=sys.stderr)

    slug  = "well-information"
    clear = ENDPOINTS[slug][0]
    s     = make_session()
    html  = fetch_page(s, slug, clear)
    state = parse_apex_state(html)
    if not state.get("pInstance"):
        print("[!] Failed to extract APEX session tokens.", file=sys.stderr)
        sys.exit(1)

    regions = parse_ig_regions(html)
    if not regions:
        print("[!] No downloadable IG regions found for well-information.", file=sys.stderr)
        sys.exit(1)

    region   = regions[0]
    app_id   = state["pFlowId"]
    page_id  = state["pFlowStepId"]
    instance = state["pInstance"]

    # APEX page protection rejects setting filter items via the download AJAX
    # payload.  Instead, use the classic f?p= URL scheme which sets session
    # state server-side before each download.  The download then runs against
    # the already-filtered session state with no pageItems needed.
    wells: list[dict] = []
    for i, serial in enumerate(serials, 1):
        try:
            fp_url = (f"{BASE_URL}/f?p={app_id}:{page_id}:{instance}"
                      f":::{page_id}:P{page_id}_WELL_SERIAL_NUM:{serial}")
            r = s.get(fp_url, allow_redirects=True, timeout=30)
            r.raise_for_status()

            new_state   = parse_apex_state(r.text)
            new_regions = parse_ig_regions(r.text)
            if not new_state.get("pInstance") or not new_regions:
                print(f"    [{i}/{len(serials)}] {serial}: no APEX state in response",
                      file=sys.stderr)
                continue

            # Update instance so the next f?p= call uses the live session
            instance = new_state["pInstance"]

            raw  = download_region(s, new_state, new_regions[0], slug, clear,
                                   fmt="CSV", page_items=None)
            rows = list(csv.DictReader(raw.decode("utf-8-sig").splitlines()))
            if not rows:
                print(f"    [{i}/{len(serials)}] {serial}: no rows returned", file=sys.stderr)
                continue

            row     = rows[0]
            lat_str = row.get("Latitude", "").strip()
            lon_str = row.get("Longitude", "").strip()
            if not lat_str or not lon_str:
                print(f"    [{i}/{len(serials)}] {serial}: coordinates absent", file=sys.stderr)
                continue

            wells.append({
                "serial":    serial,
                "lat":       float(lat_str),
                "lon":       float(lon_str),
                "well_name": row.get("Well Name", ""),
                "operator":  row.get("Operator Name", ""),
                "parish":    row.get("Parish Name", ""),
                "status":    row.get("Well Status Code Description", ""),
            })
            print(f"    [{i}/{len(serials)}] {serial}  {lat_str}, {lon_str}", file=sys.stderr)
        except Exception as exc:
            print(f"    [{i}/{len(serials)}] {serial}: {exc}", file=sys.stderr)
        time.sleep(0.3)

    if not wells:
        print("[!] No coordinates could be retrieved.", file=sys.stderr)
        sys.exit(1)

    print(f"[*] Mapped {len(wells)}/{len(serials)} wells.", file=sys.stderr)
    map_html = _generate_map_html(wells)

    out_path = Path(args.out) if args.out else Path("wells_map.html")
    out_path.write_text(map_html, encoding="utf-8")
    print(f"[+] {out_path}", file=sys.stderr)

    try:
        import subprocess
        subprocess.Popen(["open", str(out_path)])
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Area of Review (AOR) — wells within a radius + optional log check
# ---------------------------------------------------------------------------

def _haversine_miles(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    R = 3958.8
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = (math.sin(dlat / 2) ** 2
         + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dlon / 2) ** 2)
    return 2 * R * math.asin(math.sqrt(a))


def _dd_to_utm15(lat: float, lon: float) -> tuple[float, float]:
    """WGS84 decimal degrees → UTM Zone 15 NAD83 easting/northing (meters)."""
    try:
        from pyproj import Transformer
    except ImportError:
        raise RuntimeError("pyproj is required for coordinate conversion: pip install pyproj")
    t = Transformer.from_crs("EPSG:4326", "EPSG:26915", always_xy=True)
    return t.transform(lon, lat)


def _pw_get_serial_coords(bpage, serial: str) -> tuple[float, float] | None:
    """Navigate to well-information, search for serial, and return (lat, lon).
    Returns None when the well record carries no surface coordinates."""
    bpage.goto(f"{PORTAL}/well-information?clear=2700",
               wait_until="networkidle", timeout=20000)
    bpage.fill("#P2700_WELL_SERIAL_NUM", serial)
    bpage.locator('[id="B275295944573754186"]').first.click()
    bpage.wait_for_load_state("networkidle", timeout=15000)
    bpage.wait_for_timeout(500)

    state, region = _extract_well_info_tokens(bpage.content())
    if not region:
        return None
    try:
        rows = _pw_ig_csv(bpage, state, region)
    except Exception:
        return None
    if not rows:
        return None
    lat_s = rows[0].get("Latitude",  "").strip()
    lon_s = rows[0].get("Longitude", "").strip()
    try:
        return float(lat_s), float(lon_s)
    except (ValueError, TypeError):
        return None


def _pw_aor_wells(bpage, lat: float, lon: float, radius_mi: float) -> list[dict]:
    """Run the SONRIS coordinate search and return all wells within radius_mi
    miles as a list of dicts, each augmented with a distance_miles field.
    Uses UTM Zone 15 NAD83 (covers lon -90° to -96°) with a haversine
    post-filter to correct for UTM scale distortion at the circle boundary."""
    easting, northing = _dd_to_utm15(lat, lon)
    radius_ft = round(radius_mi * 5280)

    bpage.goto(f"{PORTAL}/wells-and-usdw-by-coordinates?clear=2216",
               wait_until="networkidle", timeout=30000)
    bpage.fill("#P2216_X_COORDINATE", f"{easting:.2f}")
    bpage.fill("#P2216_Y_COORDINATE", f"{northing:.2f}")
    bpage.select_option("#P2216_SURFACE_COORD_SYSTEM", "26915")
    bpage.fill("#P2216_RADIUS", str(radius_ft))

    bpage.locator('[id="B397106514738617835"]').first.click()
    bpage.wait_for_load_state("networkidle", timeout=30000)
    bpage.wait_for_timeout(1500)

    html    = bpage.content()
    state   = parse_apex_state(html)
    regions = parse_ig_regions(html)
    if not state.get("pInstance") or not regions:
        return []
    try:
        rows = _pw_ig_csv(bpage, state, regions[0])
    except Exception:
        return []

    wells = []
    for row in rows:
        lat_s = row.get("Latitude",  "").strip()
        lon_s = row.get("Longitude", "").strip()
        if lat_s and lon_s:
            try:
                dist = _haversine_miles(lat, lon, float(lat_s), float(lon_s))
                if dist > radius_mi:
                    continue
                row["distance_miles"] = f"{dist:.4f}"
            except (ValueError, TypeError):
                row["distance_miles"] = ""
        else:
            row["distance_miles"] = ""
        wells.append(row)
    return wells


def cmd_aor(args):
    lat, lon, radius_mi = args.lat, args.lon, args.radius
    easting, northing   = _dd_to_utm15(lat, lon)
    print(f"[*] AOR centre: lat={lat}, lon={lon}, radius={radius_mi} mi ({round(radius_mi*5280)} ft)",
          file=sys.stderr)
    print(f"[*] UTM Zone 15 NAD83: E={easting:.1f} m, N={northing:.1f} m", file=sys.stderr)

    sync_playwright = _require_playwright()
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        ctx     = browser.new_context(user_agent=UA)
        _pw_load_cookies(ctx)
        bpage   = ctx.new_page()
        print("[*] Executing spatial query ...", file=sys.stderr)
        wells   = _pw_aor_wells(bpage, lat, lon, radius_mi)
        _pw_save_cookies(ctx)
        browser.close()

    if not wells:
        print(f"[~] No wells within {radius_mi} mi.", file=sys.stderr)
        sys.exit(0)

    print(f"[*] {len(wells)} well(s) in AOR.", file=sys.stderr)

    if not args.skip_log_check:
        s_req = make_session()
        idx   = DOC_INDEXES["well"]
        print(f"[*] Checking log availability ...", file=sys.stderr)
        for i, well in enumerate(wells, 1):
            serial = well.get("Well Serial Num", "").strip()
            if not serial:
                well["has_logs"] = ""
                continue
            try:
                docs = fetch_all_doc_list(s_req, idx, serial)
                well["has_logs"] = INFO_FIELDS["has_logs"](docs)
            except Exception as exc:
                print(f"    [{i}/{len(wells)}] {serial}: {exc}", file=sys.stderr)
                well["has_logs"] = "error"
            print(f"    [{i}/{len(wells)}] {serial}: {well['has_logs']}", file=sys.stderr)
            time.sleep(0.4)

    out_path   = Path(args.out)
    fieldnames = list(wells[0].keys())
    with open(out_path, "w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(wells)

    log_count = sum(1 for w in wells if w.get("has_logs") == "yes")
    summary   = f"{len(wells)} well(s)"
    if not args.skip_log_check:
        summary += f", {log_count} with logs"
    print(f"[+] {out_path}  ({summary})", file=sys.stderr)


# ---------------------------------------------------------------------------
# search — unified well-record lookup and AOR document query
# ---------------------------------------------------------------------------

def cmd_search(args):
    """
    Three operating modes, selected by flags:

    No --aor: fetch and print every document on file for a single serial.
      The output is a plain table printed to stdout — useful for a quick
      record check before deciding whether to download.

    --aor: spatially query SONRIS for all wells within --radius miles of the
      target, then for each well retrieve its complete document list and append
      doc_count, doc_types, has_logs, and log_count columns.  Output is a CSV
      that contains the full spatial result set annotated with document metadata.

    --aor --logs: same spatial query, but the per-well document step writes only
      has_logs and log_count.  Use this when you specifically want to know which
      wells in an area have well logs on file.

    The target is either a well serial number (positional argument) or an
    explicit latitude/longitude pair supplied via --lat and --lon.  When a
    serial is given with --aor, the well's surface coordinates are looked up
    automatically before the spatial query runs.
    """
    serial = args.serial
    lat    = args.lat
    lon    = args.lon

    if serial is None and (lat is None or lon is None):
        print("[!] Provide a serial number, or both --lat and --lon.", file=sys.stderr)
        sys.exit(1)
    if args.logs and not args.aor:
        print("[!] --logs requires --aor.", file=sys.stderr)
        sys.exit(1)

    # ------------------------------------------------------------------
    # Mode 1 — single serial, no AOR: print document list to stdout
    # ------------------------------------------------------------------
    if not args.aor:
        s    = make_session()
        docs = fetch_all_doc_list(s, DOC_INDEXES["well"], serial)
        if not docs:
            print(f"No documents found for serial {serial}.")
            return
        col_w = max(len(d["doc_type"]) for d in docs)
        print(f"\nSerial {serial} — {len(docs)} document(s)\n")
        for d in docs:
            desc = f"  {d['description']}" if d.get("description") else ""
            print(f"  {d['docname']:>12s}  {d['doc_type']:<{col_w}}  {d['create_date']}{desc}")
        print()
        return

    # ------------------------------------------------------------------
    # Modes 2 & 3 — AOR: spatial query + per-well document check
    # ------------------------------------------------------------------
    sync_playwright = _require_playwright()

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        ctx     = browser.new_context(user_agent=UA)
        _pw_load_cookies(ctx)
        bpage   = ctx.new_page()

        if serial is not None and lat is None:
            print(f"[*] Looking up coordinates for serial {serial} ...", file=sys.stderr)
            coords = _pw_get_serial_coords(bpage, serial)
            if coords is None:
                print(f"[!] No surface coordinates on record for serial {serial}.",
                      file=sys.stderr)
                browser.close()
                sys.exit(1)
            lat, lon = coords
            print(f"[*] Surface location: lat={lat:.6f}, lon={lon:.6f}", file=sys.stderr)

        print(f"[*] Querying wells within {args.radius} mi ...", file=sys.stderr)
        wells = _pw_aor_wells(bpage, lat, lon, args.radius)
        _pw_save_cookies(ctx)
        browser.close()

    if not wells:
        print(f"[~] No wells found within {args.radius} mi.", file=sys.stderr)
        sys.exit(0)

    print(f"[*] {len(wells)} well(s) in AOR. Fetching document records ...",
          file=sys.stderr)

    s_req = make_session()
    idx   = DOC_INDEXES["well"]

    # Cache results by serial so wells with multiple completion rows are only
    # queried once (the coordinate search returns one row per completion interval)
    doc_cache: dict[str, dict] = {}

    def _fetch_doc_info(w_serial: str) -> dict:
        if w_serial in doc_cache:
            return doc_cache[w_serial]
        docs = fetch_all_doc_list(s_req, idx, w_serial)
        result = {
            "doc_count": INFO_FIELDS["doc_count"](docs),
            "doc_types": INFO_FIELDS["doc_types"](docs),
            "has_logs":  INFO_FIELDS["has_logs"](docs),
            "log_count": INFO_FIELDS["log_count"](docs),
        }
        doc_cache[w_serial] = result
        time.sleep(0.4)
        return result

    unique_serials = sorted({w.get("Well Serial Num", "").strip() for w in wells} - {""})
    print(f"[*] Checking {len(unique_serials)} unique serial(s) ...", file=sys.stderr)

    for i, w_serial in enumerate(unique_serials, 1):
        try:
            info = _fetch_doc_info(w_serial)
        except Exception as exc:
            print(f"    [{i}/{len(unique_serials)}] {w_serial}: {exc}", file=sys.stderr)
            doc_cache[w_serial] = {f: "error" for f in ("doc_count", "doc_types", "has_logs", "log_count")}
            continue
        if args.logs:
            print(f"    [{i}/{len(unique_serials)}] {w_serial}: has_logs={info['has_logs']}",
                  file=sys.stderr)
        else:
            print(f"    [{i}/{len(unique_serials)}] {w_serial}: "
                  f"{info['doc_count']} docs  has_logs={info['has_logs']}",
                  file=sys.stderr)

    for well in wells:
        w_serial = well.get("Well Serial Num", "").strip()
        if not w_serial:
            if args.logs:
                well["has_logs"] = well["log_count"] = ""
            else:
                for f in ("doc_count", "doc_types", "has_logs", "log_count"):
                    well[f] = ""
            continue
        info = doc_cache.get(w_serial, {})
        if args.logs:
            well["has_logs"]  = info.get("has_logs",  "")
            well["log_count"] = info.get("log_count", "")
        else:
            well["doc_count"] = info.get("doc_count", "")
            well["doc_types"] = info.get("doc_types", "")
            well["has_logs"]  = info.get("has_logs",  "")
            well["log_count"] = info.get("log_count", "")

    out_path   = Path(args.out)
    fieldnames = list(wells[0].keys())
    with open(out_path, "w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(wells)

    with_logs = sum(1 for w in wells if w.get("has_logs") == "yes")
    label     = "log check" if args.logs else "full doc scan"
    print(f"[+] {out_path}  ({len(wells)} wells, {with_logs} with logs, {label})",
          file=sys.stderr)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(
        description="SONRIS Data Portal bulk downloader",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    sub = ap.add_subparsers(dest="cmd", required=True)

    sub.add_parser("list", help="List all available tabular endpoints")

    fi = sub.add_parser("fields", help="List filter items available for an endpoint")
    fi.add_argument("slug", help="Endpoint slug")

    dl = sub.add_parser("download", help="Download one tabular endpoint")
    dl.add_argument("slug",   help="Endpoint slug (e.g. well-logs-ig)")
    dl.add_argument("--format",  choices=["CSV", "XLSX", "HTML"], default="CSV")
    dl.add_argument("--out",     help="Output file path (default: auto-named)")
    dl.add_argument("--filter",  action="append", metavar="KEY=VALUE",
                    help="APEX page item filter, e.g. --filter P2700_WELL_SERIAL_NUM=123456 (repeatable)")
    dl.add_argument("--debug",   action="store_true",
                    help="Print parsed APEX state tokens and region metadata")

    da = sub.add_parser("download-all", help="Download every tabular endpoint")
    da.add_argument("--format",  choices=["CSV", "XLSX", "HTML"], default="CSV")
    da.add_argument("--outdir",  default="./sonris_data", help="Output directory")

    dc = sub.add_parser("docs", help="List or download documents for a well or operator")
    dc.add_argument("val",  help="Value to search (e.g. a well serial number)")
    dc.add_argument("--idx", choices=list(DOC_INDEXES), default="well",
                    help="Search index: well, operator, field, docname, doctype (default: well)")
    dc.add_argument("--type", metavar="KEYWORD",
                    help="Filter listed/downloaded documents to types containing this string "
                         "(e.g. --type LOG, --type PERMIT)")
    dc.add_argument("--download", action="store_true",
                    help="Download the listed documents (default is list only)")
    dc.add_argument("--format", choices=["pdf", "tiff"], default="pdf",
                    help="File format for --download: pdf (default) or tiff (original scan)")
    dc.add_argument("--outdir", help="Output directory for --download (default: current directory)")

    ck = sub.add_parser("check", help="Batch-check well serials from a CSV and append result columns")
    ck.add_argument("csv_file", help="Input CSV file (serial numbers in one column)")
    ck.add_argument("--serial-col", default="0", metavar="COL",
                    help="Column index (0-based) or header name containing serial numbers (default: 0)")
    ck.add_argument("--info", default="has_logs",
                    metavar="FIELDS",
                    help="Comma-separated result columns to append. "
                         "Options: has_logs, log_count, log_descriptions, has_docs, doc_count, doc_types  (default: has_logs)")
    ck.add_argument("--out", help="Output CSV path (default: <input>_results.csv)")

    dp = sub.add_parser("depth", help="Batch-fetch well info (depth, lat/lon, etc.) for a list of serials")
    dp_src = dp.add_mutually_exclusive_group(required=True)
    dp_src.add_argument("--serials", metavar="S1,S2,...",
                        help="Comma-separated list of well serial numbers")
    dp_src.add_argument("--csv", dest="csv_file", metavar="FILE",
                        help="CSV file containing serial numbers")
    dp.add_argument("--serial-col", default="0", metavar="COL",
                    help="Column index or header name for serial numbers (default: 0)")
    dp.add_argument("--filter-col", metavar="COL",
                    help="Header name of a column to filter by when using --csv")
    dp.add_argument("--filter-val", metavar="VAL",
                    help="Value to match in --filter-col")
    dp.add_argument("--out", default="well_depths.csv",
                    help="Output CSV path (default: well_depths.csv)")

    mp = sub.add_parser("map", help="Fetch coordinates for wells and generate a Leaflet.js HTML map")
    mp_src = mp.add_mutually_exclusive_group(required=True)
    mp_src.add_argument("--serials", metavar="S1,S2,...",
                        help="Comma-separated list of well serial numbers")
    mp_src.add_argument("--csv", dest="csv_file", metavar="FILE",
                        help="CSV file containing serial numbers")
    mp.add_argument("--serial-col", default="0", metavar="COL",
                    help="Column index or header name for serial numbers when using --csv (default: 0)")
    mp.add_argument("--filter-col", metavar="COL",
                    help="Header name of a column to filter by when using --csv (e.g. has_logs)")
    mp.add_argument("--filter-val", metavar="VAL",
                    help="Value to match in --filter-col (e.g. yes)")
    mp.add_argument("--out", default="wells_map.html",
                    help="Output HTML file path (default: wells_map.html)")

    ar = sub.add_parser("aor",
                        help="Find all wells within a radius of a lat/lon and check for logs")
    ar.add_argument("--lat",    type=float, required=True,
                    help="Centre latitude in decimal degrees (WGS84)")
    ar.add_argument("--lon",    type=float, required=True,
                    help="Centre longitude in decimal degrees (WGS84)")
    ar.add_argument("--radius", type=float, default=2.0,
                    help="Search radius in miles (default: 2.0). "
                         "Uses SONRIS UTM Zone 15 NAD83 — best coverage for lon -90° to -96°.")
    ar.add_argument("--no-log-check", dest="skip_log_check", action="store_true",
                    help="Skip per-well log availability check (default: check each well)")
    ar.add_argument("--out", default="aor_wells.csv",
                    help="Output CSV path (default: aor_wells.csv)")

    sr = sub.add_parser(
        "search",
        help="Look up records for a well, or find and annotate all wells in an AOR",
        description=(
            "Without --aor: print every document on file for a single well serial.\n\n"
            "With --aor: find all wells within --radius miles of the target (a serial\n"
            "  or a --lat/--lon pair), fetch the full document list for each well, and\n"
            "  write the annotated results to a CSV (columns: all spatial fields from\n"
            "  SONRIS + distance_miles + doc_count + doc_types + has_logs + log_count).\n\n"
            "With --aor --logs: same spatial query, but the per-well document step is\n"
            "  narrowed to log availability only (has_logs, log_count) — answers the\n"
            "  question \"which wells near here have well logs on file?\"\n\n"
            "Examples:\n"
            "  python3 sonris.py search 7921\n"
            "  python3 sonris.py search 7921 --aor\n"
            "  python3 sonris.py search 7921 --aor --logs\n"
            "  python3 sonris.py search 7921 --aor --logs --radius 5\n"
            "  python3 sonris.py search --lat 30.45 --lon -91.18 --aor --logs\n"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    sr.add_argument("serial", nargs="?", metavar="SERIAL",
                    help="Well serial number (required unless --lat/--lon are given)")
    sr.add_argument("--lat",    type=float, metavar="LAT",
                    help="Centre latitude in decimal degrees — use instead of a serial")
    sr.add_argument("--lon",    type=float, metavar="LON",
                    help="Centre longitude in decimal degrees (required with --lat)")
    sr.add_argument("--aor",    action="store_true",
                    help="Find all wells within --radius miles of the target")
    sr.add_argument("--logs",   action="store_true",
                    help="With --aor: report only log availability (has_logs, log_count)")
    sr.add_argument("--radius", type=float, default=2.0,
                    help="AOR radius in miles (default: 2.0)")
    sr.add_argument("--out",    default="search_aor.csv",
                    help="Output CSV for AOR results (default: search_aor.csv)")

    args = ap.parse_args()
    {
        "list":         cmd_list,
        "fields":       cmd_fields,
        "download":     cmd_download,
        "download-all": cmd_download_all,
        "docs":         cmd_docs,
        "check":        cmd_check,
        "depth":        cmd_depth,
        "map":          cmd_map,
        "aor":          cmd_aor,
        "search":       cmd_search,
    }[args.cmd](args)


if __name__ == "__main__":
    main()
