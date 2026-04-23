# sonris

A headless command-line tool for bulk data access against the Louisiana DNR [SONRIS data portal](https://sonlite.dnr.state.la.us/ords/r/sonris_pub/sonris_data_portal). Every interactive grid on the portal is an Oracle APEX 24.2 application. This tool automates the two-step signed download flow that those grids expose, as well as the document access and spatial search subsystems, without requiring a browser proxy or manual interaction.

## How it works

Oracle APEX Interactive Grids expose a preparatory POST endpoint that validates a per-session checksum and returns a transient file identifier, followed by a streaming GET that delivers the payload. This tool extracts the APEX state tokens from the page HTML, constructs the signed download request, and streams the response to disk. For pages that require a CAPTCHA, the tool opens a visible browser window, waits for the user to solve it, then ports the resulting session cookies back into the requests session.

Document retrieval operates through a separate APEX Interactive Report. The tool parses the report HTML directly, anchoring on table rows rather than filtering individual cells, which preserves the fixed column layout needed for reliable positional field extraction (content ID, create date, document type, description).

Area-of-record spatial queries use SONRIS page 2216 (wells-and-usdw-by-coordinates). Coordinates are converted from WGS84 decimal degrees to UTM Zone 15 NAD83 (EPSG:26915) using pyproj, because the SONRIS form expects UTM easting/northing in meters and a radius in feet. A Haversine post-filter removes marginal outliers introduced by UTM scale distortion at the edges of the search radius.

## Requirements

Python 3.10 or later. Core dependencies:

```
pip install requests pyproj
```

Playwright is required for spatial queries and document downloads:

```
pip install playwright && playwright install chromium
```

## Commands

### list

Print all available tabular endpoints grouped by category.

```
python3 sonris.py list
```

### download

Download a single tabular endpoint to CSV, XLSX, or HTML.

```
python3 sonris.py download well-logs-ig
python3 sonris.py download well-information --filter P2700_WELL_SERIAL_NUM=7921
python3 sonris.py download well-logs-ig --format XLSX --out wells.xlsx
```

`--filter KEY=VALUE` passes an APEX page item filter and may be repeated. `--debug` prints the parsed APEX state tokens and region metadata.

### download-all

Download every endpoint in the manifest to a directory.

```
python3 sonris.py download-all --outdir ./sonris_data/
```

### docs

List or download the documents attached to a well, operator, field, or document name.

```
python3 sonris.py docs 7921
python3 sonris.py docs 7921 --type LOG
python3 sonris.py docs 7921 --type LOG --download --outdir ./logs/
```

`--idx` selects the search index: `well` (default), `operator`, `field`, `docname`, `doctype`. `--type KEYWORD` filters results to document types containing that string. `--download` fetches the actual files; `--format` selects `pdf` (default) or `tiff`.

### check

Batch-check a list of well serials from a CSV and append result columns.

```
python3 sonris.py check serials.csv --info has_logs,log_count,log_descriptions
python3 sonris.py check serials.csv --serial-col well_serial --info has_logs,doc_types
```

`--serial-col` accepts a zero-based column index or a header name. `--info` is a comma-separated list of fields to append:

- `has_logs` — `yes` or `no`
- `log_count` — number of well log documents
- `log_descriptions` — semicolon-joined description strings for each log document
- `has_docs` — `yes` or `no`
- `doc_count` — total number of documents
- `doc_types` — semicolon-joined set of document type strings

Output is written to `<input>_results.csv` unless `--out` is specified.

### search

Single-serial document lookup, full AOR document scan, or AOR log-focused check.

```
python3 sonris.py search 7921
python3 sonris.py search 7921 --aor
python3 sonris.py search 7921 --aor --logs
python3 sonris.py search 7921 --aor --logs --radius 5
python3 sonris.py search --lat 30.45 --lon -91.18 --aor --logs
```

Without `--aor`, the command prints every document on file for the given serial to stdout, including content ID, create date, document type, and description.

With `--aor`, the command finds all wells within `--radius` miles of the target (resolved from the serial's surface coordinates, or from `--lat`/`--lon` directly), fetches the full document record for each unique well serial, and writes a CSV annotated with `doc_count`, `doc_types`, `has_logs`, and `log_count`.

With `--aor --logs`, the per-well document step is narrowed to log availability only (`has_logs`, `log_count`), which reduces the number of API calls when the only question is whether well logs exist.

The SONRIS spatial query returns one row per completion interval, so a single well serial may appear multiple times. Document API calls are deduplicated by serial — each unique well is queried exactly once regardless of how many completion intervals it has.

### depth

Fetch well information records (total depth, surface location, operator, status) for a list of serials.

```
python3 sonris.py depth --serials 7921,155058 --out depths.csv
python3 sonris.py depth --csv serials.csv --out depths.csv
python3 sonris.py depth --csv results.csv --filter-col has_logs --filter-val yes
```

### map

Fetch surface coordinates for a list of wells and generate a Leaflet.js HTML map.

```
python3 sonris.py map --serials 7921,155058
python3 sonris.py map --csv results.csv --filter-col has_logs --filter-val yes
```

### aor

Find all wells within a radius of a lat/lon point and check each for log availability.

```
python3 sonris.py aor --lat 30.45 --lon -91.18 --radius 2.0
python3 sonris.py aor --lat 30.45 --lon -91.18 --radius 2.0 --no-log-check
```

This command predates `search --aor` and is retained for compatibility. `search --aor --logs` is the preferred interface for new work.

## Coordinate system

The SONRIS spatial search form accepts UTM Zone 15 NAD83 coordinates (EPSG:26915) and a radius in feet. This zone covers longitudes -90° to -96°, which includes most of Louisiana's active production areas. The tool converts WGS84 decimal degrees to UTM Zone 15 using pyproj before submitting the query.

## Session management

Cookies are persisted to `~/.sonris_cookies.json` across runs. This avoids redundant session establishment on repeated calls and reduces the chance of triggering a CAPTCHA challenge on pages that rate-limit new sessions.

## Notes

The SONRIS portal runs Oracle APEX 24.2. The signed download tokens embedded in page HTML are per-session and expire when the session does. The two-step download flow (preparatory POST + streaming GET) is standard across all Interactive Grid endpoints in the portal, but the exact region and state parameters vary by page. The tool extracts these dynamically from the page HTML on each run rather than hardcoding them.

Document downloads require an active session with valid cookies. If a download returns an HTTP 400 or 403, deleting `~/.sonris_cookies.json` and re-running will establish a fresh session.
