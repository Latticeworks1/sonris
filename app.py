#!/usr/bin/env python3
"""
Thin FastAPI wrapper around sonris.py.

Exposes document search, batch serial checking, and AOR spatial queries as
HTTP endpoints, allowing the SONRIS toolchain to be used from a browser
without installing Python or managing session cookies locally.
"""

import io
import csv
import sys
import asyncio
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles

sys.path.insert(0, str(Path(__file__).parent))
import sonris  # noqa: E402

app = FastAPI(title="SONRIS")

STATIC = Path(__file__).parent / "static"

app.mount("/static", StaticFiles(directory=str(STATIC)), name="static")


@app.get("/", response_class=HTMLResponse)
async def index():
    return (STATIC / "index.html").read_text()


# ---------------------------------------------------------------------------
# /api/docs/{serial}
# ---------------------------------------------------------------------------

@app.get("/api/docs/{serial}")
async def get_docs(serial: str, idx: str = "well"):
    if idx not in sonris.DOC_INDEXES:
        raise HTTPException(400, f"Unknown index '{idx}'. Options: {list(sonris.DOC_INDEXES)}")

    def _run():
        s = sonris.make_session()
        return sonris.fetch_all_doc_list(s, sonris.DOC_INDEXES[idx], serial)

    loop = asyncio.get_event_loop()
    docs = await loop.run_in_executor(None, _run)
    return docs


# ---------------------------------------------------------------------------
# /api/check — CSV upload, annotate and return
# ---------------------------------------------------------------------------

@app.post("/api/check")
async def batch_check(
    file: UploadFile = File(...),
    serial_col: str = Form("0"),
    info: str = Form("has_logs,log_count"),
):
    raw = (await file.read()).decode("utf-8", errors="replace")
    reader = csv.DictReader(io.StringIO(raw))
    rows = list(reader)
    fieldnames = list(reader.fieldnames or [])

    if not rows:
        raise HTTPException(400, "CSV is empty.")

    if serial_col.isdigit():
        idx = int(serial_col)
        if idx >= len(fieldnames):
            raise HTTPException(400, f"Column index {idx} out of range.")
        col_name = fieldnames[idx]
    else:
        if serial_col not in fieldnames:
            raise HTTPException(400, f"Column '{serial_col}' not found.")
        col_name = serial_col

    info_fields = [f.strip() for f in info.split(",") if f.strip()]
    unknown = [f for f in info_fields if f not in sonris.INFO_FIELDS]
    if unknown:
        raise HTTPException(400, f"Unknown fields: {unknown}. Options: {list(sonris.INFO_FIELDS)}")

    def _run():
        s = sonris.make_session()
        doc_idx = sonris.DOC_INDEXES["well"]
        for row in rows:
            serial = row.get(col_name, "").strip()
            if not serial:
                for f in info_fields:
                    row[f] = ""
                continue
            try:
                docs = sonris.fetch_all_doc_list(s, doc_idx, serial)
                for f in info_fields:
                    row[f] = sonris.INFO_FIELDS[f](docs)
            except Exception as exc:
                for f in info_fields:
                    row[f] = f"error: {exc}"
        return rows

    loop = asyncio.get_event_loop()
    result_rows = await loop.run_in_executor(None, _run)

    out_fields = fieldnames + [f for f in info_fields if f not in fieldnames]
    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=out_fields, extrasaction="ignore")
    writer.writeheader()
    writer.writerows(result_rows)

    return StreamingResponse(
        io.BytesIO(buf.getvalue().encode()),
        media_type="text/csv",
        headers={"Content-Disposition": "attachment; filename=results.csv"},
    )


# ---------------------------------------------------------------------------
# /api/aor — spatial query (Playwright, runs in thread pool)
# ---------------------------------------------------------------------------

@app.post("/api/aor")
async def aor_search(body: dict):
    lat = body.get("lat")
    lon = body.get("lon")
    radius = float(body.get("radius", 2.0))
    logs_only = bool(body.get("logs", False))

    if lat is None or lon is None:
        raise HTTPException(400, "lat and lon are required.")

    def _run():
        sync_playwright = sonris._require_playwright()
        with sync_playwright() as pw:
            browser = pw.chromium.launch(headless=True)
            ctx = browser.new_context(user_agent=sonris.UA)
            sonris._pw_load_cookies(ctx)
            bpage = ctx.new_page()
            wells = sonris._pw_aor_wells(bpage, float(lat), float(lon), radius)
            sonris._pw_save_cookies(ctx)
            browser.close()

        if not wells:
            return []

        s = sonris.make_session()
        doc_idx = sonris.DOC_INDEXES["well"]
        doc_cache: dict = {}

        unique = sorted({w.get("Well Serial Num", "").strip() for w in wells} - {""})
        for serial in unique:
            try:
                docs = sonris.fetch_all_doc_list(s, doc_idx, serial)
                doc_cache[serial] = {
                    "doc_count": sonris.INFO_FIELDS["doc_count"](docs),
                    "doc_types": sonris.INFO_FIELDS["doc_types"](docs),
                    "has_logs":  sonris.INFO_FIELDS["has_logs"](docs),
                    "log_count": sonris.INFO_FIELDS["log_count"](docs),
                }
            except Exception:
                doc_cache[serial] = {k: "error" for k in
                                     ("doc_count", "doc_types", "has_logs", "log_count")}

        for well in wells:
            serial = well.get("Well Serial Num", "").strip()
            info = doc_cache.get(serial, {})
            if logs_only:
                well["has_logs"]  = info.get("has_logs",  "")
                well["log_count"] = info.get("log_count", "")
            else:
                well.update(info)

        return wells

    loop = asyncio.get_event_loop()
    wells = await loop.run_in_executor(None, _run)
    return wells
