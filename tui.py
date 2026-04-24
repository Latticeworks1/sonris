#!/usr/bin/env python3
"""
Interactive CLI for SONRIS.
Run: python3 tui.py
"""

import sys
import csv
import io
import time
from pathlib import Path

try:
    import questionary
    from questionary import Style
    from rich.console import Console
    from rich.table import Table
    from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn, TaskProgressColumn
    from rich import box as rbox
except ImportError:
    sys.exit("pip install questionary rich")

sys.path.insert(0, str(Path(__file__).parent))
import sonris

console = Console()

STYLE = Style([
    ("qmark",        "fg:#7ab bold"),
    ("question",     "bold"),
    ("answer",       "fg:#7ab bold"),
    ("pointer",      "fg:#7ab bold"),
    ("highlighted",  "fg:#7ab bold"),
    ("selected",     "fg:#aaa"),
    ("separator",    "fg:#444"),
    ("instruction",  "fg:#555"),
])


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

def _doc_table(docs: list[dict], serial: str) -> None:
    if not docs:
        console.print("[dim]  no documents found[/dim]")
        return
    t = Table(box=rbox.SIMPLE, header_style="bold cyan", show_edge=False)
    t.add_column("docname",     style="dim",    no_wrap=True)
    t.add_column("date",                        no_wrap=True)
    t.add_column("type",        style="white",  no_wrap=True)
    t.add_column("description", style="dim")
    for d in docs:
        t.add_row(d["docname"], d["create_date"], d["doc_type"], d.get("description", ""))
    console.print(f"\n  [cyan]{len(docs)} document(s) for serial {serial}[/cyan]")
    console.print(t)


def _well_table(wells: list[dict]) -> None:
    if not wells:
        console.print("[dim]  no wells found[/dim]")
        return
    cols = list(wells[0].keys())
    t = Table(box=rbox.SIMPLE, header_style="bold cyan", show_edge=False)
    for c in cols:
        t.add_column(c, no_wrap=(c != "doc_types"))
    for w in wells:
        row = []
        for c in cols:
            v = str(w.get(c, ""))
            if c == "has_logs":
                v = f"[green]{v}[/green]" if v == "yes" else f"[dim]{v}[/dim]"
            row.append(v)
        t.add_row(*row)
    console.print(t)


# ---------------------------------------------------------------------------
# search
# ---------------------------------------------------------------------------

def run_search() -> None:
    serial = questionary.text("Well serial number:", style=STYLE).ask()
    if not serial or not serial.strip():
        return
    serial = serial.strip()
    with console.status("  fetching documents…", spinner="dots"):
        s    = sonris.make_session()
        docs = sonris.fetch_all_doc_list(s, sonris.DOC_INDEXES["well"], serial)
    _doc_table(docs, serial)


# ---------------------------------------------------------------------------
# batch check
# ---------------------------------------------------------------------------

def run_check() -> None:
    path_str = questionary.path(
        "CSV file:", style=STYLE, validate=lambda p: Path(p).is_file() or "file not found"
    ).ask()
    if not path_str:
        return

    col = questionary.text(
        "Serial column (header name or 0-based index):", default="0", style=STYLE
    ).ask()
    if col is None:
        return

    fields = questionary.checkbox(
        "Fields to append:",
        choices=list(sonris.INFO_FIELDS.keys()),
        default=["has_logs", "log_count"],
        style=STYLE,
    ).ask()
    if not fields:
        return

    out_path = questionary.text(
        "Output CSV path:", default="results.csv", style=STYLE
    ).ask()
    if not out_path:
        return

    raw = Path(path_str).read_text(encoding="utf-8", errors="replace")
    reader = csv.DictReader(io.StringIO(raw))
    rows = list(reader)
    fieldnames = list(reader.fieldnames or [])

    if not rows:
        console.print("[red]  CSV is empty[/red]")
        return

    col_name = fieldnames[int(col)] if col.isdigit() else col
    if col_name not in fieldnames:
        console.print(f"[red]  column '{col_name}' not found[/red]")
        return

    s       = sonris.make_session()
    doc_idx = sonris.DOC_INDEXES["well"]

    with Progress(
        SpinnerColumn(),
        TextColumn("[cyan]{task.description}"),
        BarColumn(bar_width=30),
        TaskProgressColumn(),
        console=console,
        transient=True,
    ) as progress:
        task = progress.add_task("checking", total=len(rows))
        for row in rows:
            serial = row.get(col_name, "").strip()
            progress.update(task, description=f"checking {serial or '—'}")
            if not serial:
                for f in fields:
                    row[f] = ""
                progress.advance(task)
                continue
            try:
                docs = sonris.fetch_all_doc_list(s, doc_idx, serial)
                for f in fields:
                    row[f] = sonris.INFO_FIELDS[f](docs)
            except Exception as exc:
                for f in fields:
                    row[f] = f"error: {exc}"
            progress.advance(task)

    out_fields = fieldnames + [f for f in fields if f not in fieldnames]
    with open(out_path, "w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=out_fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)

    with_logs = sum(1 for r in rows if r.get("has_logs") == "yes")
    console.print(f"\n  [green]saved {out_path}[/green]  ({len(rows)} rows"
                  + (f", {with_logs} with logs" if "has_logs" in fields else "") + ")")


# ---------------------------------------------------------------------------
# aor
# ---------------------------------------------------------------------------

def run_aor() -> None:
    use_serial = questionary.confirm(
        "Look up center point from a well serial?", default=False, style=STYLE
    ).ask()

    lat = lon = None
    if use_serial:
        serial = questionary.text("Well serial number:", style=STYLE).ask()
        if not serial:
            return
        sync_playwright = sonris._require_playwright()
        with console.status("  looking up coordinates…", spinner="dots"):
            with sync_playwright() as pw:
                browser = pw.chromium.launch(headless=True)
                ctx     = browser.new_context(user_agent=sonris.UA)
                sonris._pw_load_cookies(ctx)
                bpage   = ctx.new_page()
                coords  = sonris._pw_get_serial_coords(bpage, serial.strip())
                sonris._pw_save_cookies(ctx)
                browser.close()
        if coords is None:
            console.print("[red]  no surface coordinates on record for that serial[/red]")
            return
        lat, lon = coords
        console.print(f"  [dim]surface location: lat={lat:.6f}, lon={lon:.6f}[/dim]")
    else:
        lat_s = questionary.text("Latitude (decimal degrees):", style=STYLE).ask()
        lon_s = questionary.text("Longitude (decimal degrees):", style=STYLE).ask()
        if not lat_s or not lon_s:
            return
        try:
            lat, lon = float(lat_s), float(lon_s)
        except ValueError:
            console.print("[red]  invalid coordinates[/red]")
            return

    radius_s = questionary.text("Radius (miles):", default="2.0", style=STYLE).ask()
    try:
        radius = float(radius_s)
    except (TypeError, ValueError):
        radius = 2.0

    logs_only = questionary.confirm(
        "Log availability check only (faster)?", default=False, style=STYLE
    ).ask()

    out_path = questionary.text(
        "Output CSV path:", default="aor_results.csv", style=STYLE
    ).ask()
    if not out_path:
        return

    with console.status(f"  querying wells within {radius} mi…", spinner="dots"):
        sync_playwright = sonris._require_playwright()
        with sync_playwright() as pw:
            browser = pw.chromium.launch(headless=True)
            ctx     = browser.new_context(user_agent=sonris.UA)
            sonris._pw_load_cookies(ctx)
            bpage   = ctx.new_page()
            wells   = sonris._pw_aor_wells(bpage, lat, lon, radius)
            sonris._pw_save_cookies(ctx)
            browser.close()

    if not wells:
        console.print(f"  [dim]no wells found within {radius} mi[/dim]")
        return

    console.print(f"  [cyan]{len(wells)} well interval(s)[/cyan] — fetching document records…")

    s       = sonris.make_session()
    doc_idx = sonris.DOC_INDEXES["well"]
    cache: dict = {}

    unique = sorted({w.get("Well Serial Num", "").strip() for w in wells} - {""})
    with Progress(
        SpinnerColumn(),
        TextColumn("[cyan]{task.description}"),
        BarColumn(bar_width=30),
        TaskProgressColumn(),
        console=console,
        transient=True,
    ) as progress:
        task = progress.add_task("fetching docs", total=len(unique))
        for serial in unique:
            progress.update(task, description=f"serial {serial}")
            try:
                docs = sonris.fetch_all_doc_list(s, doc_idx, serial)
                cache[serial] = {
                    "doc_count": sonris.INFO_FIELDS["doc_count"](docs),
                    "doc_types": sonris.INFO_FIELDS["doc_types"](docs),
                    "has_logs":  sonris.INFO_FIELDS["has_logs"](docs),
                    "log_count": sonris.INFO_FIELDS["log_count"](docs),
                }
            except Exception:
                cache[serial] = {k: "error" for k in
                                 ("doc_count", "doc_types", "has_logs", "log_count")}
            progress.advance(task)

    for well in wells:
        serial = well.get("Well Serial Num", "").strip()
        info   = cache.get(serial, {})
        if logs_only:
            well["has_logs"]  = info.get("has_logs",  "")
            well["log_count"] = info.get("log_count", "")
        else:
            well.update(info)

    _well_table(wells)

    fieldnames = list(wells[0].keys())
    with open(out_path, "w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(wells)

    with_logs = sum(1 for w in wells if w.get("has_logs") == "yes")
    console.print(f"\n  [green]saved {out_path}[/green]  ({len(wells)} wells, {with_logs} with logs)")


# ---------------------------------------------------------------------------
# download
# ---------------------------------------------------------------------------

def run_download() -> None:
    by_cat: dict[str, list[str]] = {}
    for slug, (_, cat, _) in sonris.ENDPOINTS.items():
        by_cat.setdefault(cat, []).append(slug)

    cat = questionary.select(
        "Category:", choices=sorted(by_cat.keys()) + ["— back —"], style=STYLE
    ).ask()
    if not cat or cat == "— back —":
        return

    choices = by_cat[cat] + ["— back —"]
    slug = questionary.select("Endpoint:", choices=choices, style=STYLE).ask()
    if not slug or slug == "— back —":
        return

    fmt = questionary.select(
        "Format:", choices=["CSV", "XLSX", "HTML"], default="CSV", style=STYLE
    ).ask()
    if not fmt:
        return

    default_out = f"{slug}.{fmt.lower()}"
    out_path = questionary.text("Output file:", default=default_out, style=STYLE).ask()
    if not out_path:
        return

    filter_str = questionary.text(
        "Filters (KEY=VALUE, comma-separated, or blank):", default="", style=STYLE
    ).ask()

    filters = []
    if filter_str and filter_str.strip():
        for part in filter_str.split(","):
            part = part.strip()
            if "=" in part:
                filters.append(part)

    class _Args:
        pass

    args = _Args()
    args.slug   = slug
    args.format = fmt
    args.out    = out_path
    args.filter = filters or None
    args.debug  = False

    with console.status(f"  downloading {slug}…", spinner="dots"):
        try:
            sonris.cmd_download(args)
        except SystemExit:
            pass

    if Path(out_path).exists():
        console.print(f"\n  [green]saved {out_path}[/green]")
    else:
        console.print(f"\n  [red]download may have failed — check output above[/red]")


# ---------------------------------------------------------------------------
# main loop
# ---------------------------------------------------------------------------

MENU = ["search", "batch check", "aor", "download", "quit"]


def main() -> None:
    console.print("\n  [bold cyan]SONRIS[/bold cyan]  Louisiana DNR data portal\n")
    while True:
        choice = questionary.select("", choices=MENU, style=STYLE).ask()
        if choice is None or choice == "quit":
            break
        elif choice == "search":
            run_search()
        elif choice == "batch check":
            run_check()
        elif choice == "aor":
            run_aor()
        elif choice == "download":
            run_download()
        console.print()


if __name__ == "__main__":
    main()
