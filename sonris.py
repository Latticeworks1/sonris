#!/usr/bin/env python3
"""
SONRIS Data Portal — standalone bulk downloader.

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
"""

import re
import sys
import json
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

def make_session() -> requests.Session:
    s = requests.Session()
    s.verify = True
    s.headers.update({"User-Agent": UA, "Accept-Language": "en-US,en;q=0.5"})
    return s


# ---------------------------------------------------------------------------
# Page parsing
# ---------------------------------------------------------------------------

def fetch_page(s: requests.Session, slug: str, clear: str) -> str:
    url = f"{PORTAL}/{slug}" + (f"?clear={clear}" if clear else "")
    r = s.get(url, allow_redirects=True, timeout=30)
    r.raise_for_status()
    return r.text


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

        if not (rid_m and ajax_m and dlcs_m):
            continue

        regions.append({
            "regionId":       rid_m.group(1),
            "regionTitle":    title_m.group(1) if title_m else rid_m.group(1),
            "ajaxIdentifier": unescape(ajax_m.group(1)),
            "downloadCs":     unescape(dlcs_m.group(1)),
            "ajaxColumns":    cols,
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
) -> str:
    """Step 1 — register the download job; returns the transient file ID."""
    p_json = json.dumps({
        "pageItems": {},
        "regions": [{
            "id":             region["regionId"],
            "ajaxIdentifier": region["ajaxIdentifier"],
            "ajaxColumns":    region["ajaxColumns"],
            "view":           "grid",
            "download": {
                "downloadCs": region["downloadCs"],
                "format":     fmt,
            },
        }],
        "salt": state["pSalt"],
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
    p_json = json.dumps({
        "regions": [{
            "id":             region["regionId"],
            "ajaxIdentifier": region["ajaxIdentifier"],
            "downloadFileId": file_id,
        }],
        "salt": state["pSalt"],
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
) -> bytes:
    file_id = _register_download(s, state, region, fmt, slug, clear)
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


def cmd_download(args):
    slug = args.slug
    if slug not in ENDPOINTS:
        print(f"Unknown endpoint '{slug}'. Run 'list' to see available endpoints.")
        sys.exit(1)

    clear, _, desc = ENDPOINTS[slug]
    s = make_session()

    print(f"[*] {desc}", file=sys.stderr)
    print("[*] Fetching page ...", file=sys.stderr)
    html  = fetch_page(s, slug, clear)
    state = parse_apex_state(html)
    if not state.get("pInstance"):
        print("[!] Failed to extract APEX session tokens.", file=sys.stderr)
        sys.exit(1)

    regions = parse_ig_regions(html)
    if not regions:
        print("[!] No downloadable IG regions found on this page.", file=sys.stderr)
        sys.exit(1)

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
        data = download_region(s, state, region, slug, clear, fmt=args.format)
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

DOC_BASE    = "https://sonlite.dnr.state.la.us"
DOC_SEARCH  = f"{DOC_BASE}/ords/r/sonris_pub/document_access/finddocuments"
DOC_DL_BASE = f"{DOC_BASE}/dnrservices/redirectUrl.jsp"

# Known search indexes accepted by the finddocuments page
DOC_INDEXES = {
    "well":       "xwellserialnumber",
    "operator":   "xoperatorcode",
    "field":      "xfieldcode",
    "docname":    "xdocname",
    "doctype":    "xdocumenttype",
}


def fetch_doc_list(s: requests.Session, idx: str, val: str) -> list[dict]:
    """
    Fetch the document listing for a given index/value pair and return a list
    of dicts with keys: docname, create_date, doc_type, and raw_cells.

    The finddocuments page renders its results as an APEX Interactive Report
    whose rows are included directly in the initial HTML response.  Each row
    carries two download link anchors — one for the PDF conversion and one for
    the native file (TIFF) — both of which embed the dDocname identifier.
    """
    r = s.get(DOC_SEARCH, params={"idx": idx, "val": val},
              allow_redirects=True, timeout=30)
    r.raise_for_status()
    html = r.text

    panel_m = re.search(
        r'id="DR_IR_data_panel"[^>]*>(.*?)<div\s+role="dialog"', html, re.DOTALL
    )
    if not panel_m:
        return []

    panel = panel_m.group(1)
    docs, seen = [], set()

    for m in re.finditer(r'dDocname=([^&"]+)&showInline=True">', panel):
        docname = m.group(1)
        if docname in seen:
            continue
        seen.add(docname)

        # Collect adjacent <td> cell text as metadata
        chunk = panel[m.start(): m.start() + 2000]
        tds   = re.findall(r'<td[^>]*>(.*?)</td>', chunk, re.DOTALL)
        cells = []
        for td in tds:
            text = re.sub(r'<[^>]+>', '', td)
            text = re.sub(r'&#x2F;', '/', text)
            text = re.sub(r'&#x27;', "'", text)
            text = re.sub(r'&#x[0-9a-fA-F]+;', '', text).strip()
            if text:
                cells.append(text)

        docs.append({
            "docname":     docname,
            "create_date": cells[1] if len(cells) > 1 else "",
            "doc_type":    cells[2] if len(cells) > 2 else "",
            "raw_cells":   cells,
        })

    return docs


def download_doc(s: requests.Session, docname: str, native: bool = False) -> tuple[bytes, str]:
    """
    Download a single document.  native=True returns the original TIFF;
    native=False returns the PDF conversion.  Returns (bytes, filename).
    """
    params = {"dDocname": docname, "showInline": "True"}
    if native:
        params["nativeFile"] = "True"
    r = s.get(DOC_DL_BASE, params=params, allow_redirects=True, timeout=60)
    r.raise_for_status()
    cd = r.headers.get("content-disposition", "")
    fname_m = re.search(r'fileName="([^"]+)"', cd)
    fname = fname_m.group(1) if fname_m else f"{docname}.{'tif' if native else 'pdf'}"
    return r.content, fname


def cmd_docs(args):
    idx_key = args.idx
    if idx_key not in DOC_INDEXES:
        print(f"Unknown index '{idx_key}'. Choose from: {list(DOC_INDEXES)}")
        sys.exit(1)
    idx = DOC_INDEXES[idx_key]
    val = args.val
    native = args.format == "tiff"

    outdir = Path(args.outdir) if args.outdir else None
    if outdir:
        outdir.mkdir(parents=True, exist_ok=True)

    s = make_session()
    print(f"[*] Fetching document list ({idx_key}={val}) ...", file=sys.stderr)
    docs = fetch_doc_list(s, idx, val)

    if not docs:
        print("[~] No documents found.", file=sys.stderr)
        return

    print(f"[*] Found {len(docs)} document(s):", file=sys.stderr)
    for d in docs:
        print(f"    {d['docname']:12s}  {d['doc_type']:30s}  {d['create_date']}", file=sys.stderr)

    if args.list_only:
        return

    for d in docs:
        data, fname = download_doc(s, d["docname"], native=native)
        safe_type = re.sub(r'[^\w]+', '_', d['doc_type']).strip('_').lower()
        if outdir:
            out_path = outdir / f"{val}__{safe_type}__{fname}"
        else:
            out_path = Path(f"{val}__{safe_type}__{fname}")
        out_path.write_bytes(data)
        print(f"[+] {out_path}  ({len(data):,} bytes)", file=sys.stderr)


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

    dl = sub.add_parser("download", help="Download one tabular endpoint")
    dl.add_argument("slug",   help="Endpoint slug (e.g. well-logs-ig)")
    dl.add_argument("--format",  choices=["CSV", "XLSX", "HTML"], default="CSV")
    dl.add_argument("--out",     help="Output file path (default: auto-named)")

    da = sub.add_parser("download-all", help="Download every tabular endpoint")
    da.add_argument("--format",  choices=["CSV", "XLSX", "HTML"], default="CSV")
    da.add_argument("--outdir",  default="./sonris_data", help="Output directory")

    dc = sub.add_parser("docs", help="Download documents (PDFs/TIFFs) for a well or operator")
    dc.add_argument("val",  help="Value to search (e.g. a well serial number)")
    dc.add_argument("--idx", choices=list(DOC_INDEXES), default="well",
                    help="Search index: well, operator, field, docname, doctype (default: well)")
    dc.add_argument("--format", choices=["pdf", "tiff"], default="pdf",
                    help="File format: pdf (default) or tiff (original scan)")
    dc.add_argument("--outdir", help="Output directory (default: current directory)")
    dc.add_argument("--list-only", action="store_true",
                    help="Print document metadata without downloading")

    args = ap.parse_args()
    {
        "list":         cmd_list,
        "download":     cmd_download,
        "download-all": cmd_download_all,
        "docs":         cmd_docs,
    }[args.cmd](args)


if __name__ == "__main__":
    main()
