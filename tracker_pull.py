"""
tracker_pull.py — APC Dashboard tracker feed (Graph workbook API)

Reads the Residential Sales Tracker and Wholesale Sales Tracker directly via
the Microsoft Graph *workbook* API and writes tracker-metrics.json to the
SharePoint "Dashboard" folder (same folder as apc-qb.json / pro-qb.json).

Why this exists (vs. the M365 connector's flattened text reads):
  1. Range reads return CELL VALUES regardless of row visibility — the
     residential Job List keeps pre-current-month rows HIDDEN (not archived,
     not deleted), and the connector silently skips hidden rows. Graph range
     reads include them, unlocking full-history tracker <-> QB reconciliation.
  2. Structured cells kill the flattened-dump ambiguity: sold date vs install
     date come from their own columns, so the runbook's "one date + blank
     STATUS" heuristic is unnecessary for this feed.
  3. The Shipper tab is read through a bounded range (default cap 5,000 rows)
     so a recurrence of the phantom-used-range bloat can't blow up the read.
     usedRange(valuesOnly=true) is also used, which ignores cells that carry
     formatting but no values — the exact cause of the Jul 20 HTTP 406.

Auth: the same app-only client-credentials flow qbo_pull.py already uses for
its SharePoint upload (SHAREPOINT_TENANT_ID / _CLIENT_ID / _CLIENT_SECRET).
No new consents needed — Files.Read/Write.All application permission covers
workbook reads.

Required environment variables:
    SHAREPOINT_TENANT_ID, SHAREPOINT_CLIENT_ID, SHAREPOINT_CLIENT_SECRET
    SHAREPOINT_SITE_ID (used only for upload fallback, same as qbo_pull)
    TRACKER_DRIVE_ID        - drive (document library) holding both trackers
    TRACKER_RES_ITEM_ID     - item id of Residential Sales Tracker.xlsx
    TRACKER_WS_ITEM_ID      - item id of Wholesale Sales Tracker.xlsx

Optional:
    SHIPPER_ROW_CAP         - max Shipper rows to read (default 5000)
    ROWS_PER_REQUEST        - chunk size for range reads (default 500)
    SKIP_SHAREPOINT_UPLOAD  - "true" to write local JSON only (for testing)
"""

import json
import os
import re
import sys
import time
from datetime import date, datetime, timedelta

import requests

# Reuse the existing Graph token + SharePoint upload from the QBO script so
# there is exactly one implementation of each in this repo.
from qbo_pull import get_graph_token, upload_to_sharepoint

GRAPH = "https://graph.microsoft.com/v1.0"

DRIVE_ID = os.environ.get("TRACKER_DRIVE_ID")
RES_ITEM_ID = os.environ.get("TRACKER_RES_ITEM_ID")
WS_ITEM_ID = os.environ.get("TRACKER_WS_ITEM_ID")
SHIPPER_ROW_CAP = int(os.environ.get("SHIPPER_ROW_CAP", "5000"))
ROWS_PER_REQUEST = int(os.environ.get("ROWS_PER_REQUEST", "500"))

EXCEL_EPOCH = date(1899, 12, 30)  # Excel serial 1 = 1900-01-01 (1900 system)


# ---------------------------------------------------------------------------
# Graph workbook helpers
# ---------------------------------------------------------------------------

def graph_get(token: str, url: str, params: dict | None = None) -> dict:
    """GET with retry on throttling / transient errors (429/503/504)."""
    for attempt in range(5):
        resp = requests.get(
            url,
            headers={"Authorization": f"Bearer {token}"},
            params=params,
            timeout=120,
        )
        if resp.status_code in (429, 503, 504):
            wait = int(resp.headers.get("Retry-After", "0") or 0) or (2 ** attempt * 3)
            print(f"  throttled ({resp.status_code}), retrying in {wait}s ...")
            time.sleep(wait)
            continue
        if resp.status_code >= 400:
            raise RuntimeError(f"Graph GET {url} -> HTTP {resp.status_code}: {resp.text[:500]}")
        return resp.json()
    raise RuntimeError(f"Graph GET {url} still throttled after retries")


def list_worksheets(token: str, item_id: str) -> list:
    url = f"{GRAPH}/drives/{DRIVE_ID}/items/{item_id}/workbook/worksheets"
    return graph_get(token, url, {"$select": "id,name"}).get("value", [])


def find_sheet(sheets: list, needle: str) -> dict:
    """Case-insensitive contains-match on worksheet name."""
    needle_l = needle.lower()
    exact = [s for s in sheets if s["name"].strip().lower() == needle_l]
    if exact:
        return exact[0]
    partial = [s for s in sheets if needle_l in s["name"].strip().lower()]
    if partial:
        return partial[0]
    raise RuntimeError(
        f"No worksheet matching '{needle}'. Available: {[s['name'] for s in sheets]}"
    )


def used_range_meta(token: str, item_id: str, sheet_name: str) -> dict:
    """Address/rowCount/columnCount of the used range, ignoring format-only cells.

    valuesOnly=true is the phantom-used-range guard: cells that were formatted
    (whole-column fills etc.) but hold no values are excluded.
    """
    url = (
        f"{GRAPH}/drives/{DRIVE_ID}/items/{item_id}/workbook/"
        f"worksheets('{sheet_name}')/usedRange(valuesOnly=true)"
    )
    return graph_get(token, url, {"$select": "address,rowCount,columnCount"})


def read_range(token: str, item_id: str, sheet_name: str, address: str) -> list:
    """Values of an explicit A1-style range. Hidden rows/columns ARE included."""
    url = (
        f"{GRAPH}/drives/{DRIVE_ID}/items/{item_id}/workbook/"
        f"worksheets('{sheet_name}')/range(address='{address}')"
    )
    return graph_get(token, url, {"$select": "values"}).get("values", [])


def parse_address(address: str):
    """'Sheet1!B2:AC512' -> (start_col, start_row, end_col, end_row)."""
    a1 = address.split("!")[-1]
    m = re.match(r"^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$", a1)
    if not m:
        m2 = re.match(r"^\$?([A-Z]+)\$?(\d+)$", a1)  # single-cell used range
        if m2:
            return m2.group(1), int(m2.group(2)), m2.group(1), int(m2.group(2))
        raise RuntimeError(f"Unparseable range address: {address}")
    return m.group(1), int(m.group(2)), m.group(3), int(m.group(4))


def read_sheet(token: str, item_id: str, sheet_name: str, row_cap: int | None = None) -> dict:
    """Read a sheet's used range in row chunks. Returns values + truncation info."""
    meta = used_range_meta(token, item_id, sheet_name)
    c1, r1, c2, r2 = parse_address(meta["address"])
    total_rows = r2 - r1 + 1
    truncated = False
    if row_cap and total_rows > row_cap:
        r2 = r1 + row_cap - 1
        truncated = True

    values: list = []
    row = r1
    while row <= r2:
        chunk_end = min(row + ROWS_PER_REQUEST - 1, r2)
        values.extend(read_range(token, item_id, sheet_name, f"{c1}{row}:{c2}{chunk_end}"))
        row = chunk_end + 1
    print(f"  [{sheet_name}] used range {meta['address']} -> read rows {r1}..{r2}"
          f"{' (TRUNCATED at cap)' if truncated else ''}")
    return {
        "values": values,
        "usedRangeAddress": meta["address"],
        "usedRowCount": total_rows,
        "readRowCount": len(values),
        "firstRow": r1,
        "truncated": truncated,
    }


# ---------------------------------------------------------------------------
# Cell coercion
# ---------------------------------------------------------------------------

def to_iso_date(cell):
    """Excel serial number or M/D/YYYY string -> 'YYYY-MM-DD' (else None)."""
    if cell is None or cell == "":
        return None
    if isinstance(cell, (int, float)):
        serial = float(cell)
        if 20000 <= serial <= 80000:  # ~1954..2118, filters out money values
            return (EXCEL_EPOCH + timedelta(days=int(serial))).isoformat()
        return None
    s = str(cell).strip()
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})", s)
    if m:
        try:
            return date(int(m.group(3)), int(m.group(1)), int(m.group(2))).isoformat()
        except ValueError:
            return None
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", s)
    if m:
        return s[:10]
    return None


def to_money(cell):
    if cell is None or cell == "":
        return None
    if isinstance(cell, (int, float)):
        return round(float(cell), 2)
    s = str(cell).replace(",", "").replace("$", "").strip()
    try:
        return round(float(s), 2)
    except ValueError:
        return None


def to_str(cell):
    if cell is None:
        return None
    s = str(cell).strip()
    return s or None


def row_is_empty(row: list) -> bool:
    return all(c is None or str(c).strip() == "" for c in row)


# ---------------------------------------------------------------------------
# Residential Job List normalization
# ---------------------------------------------------------------------------

# header token -> output field. Contains-match, case-insensitive, first hit wins
# per field. "INSTALL DATE" also matches the sheet's actual "INSTALL DATE2".
RES_HEADER_MAP = [
    ("INSTALL DATE", "installDate"),
    ("SOLD DATE", "soldDate"),
    ("COMPANY", "company"),
    ("FIRST NAME", "firstName"),
    ("SURNAME", "surname"),
    ("LOCATION", "location"),
    ("SIZE", "size"),
    ("TYPE", "type"),
    ("FRAME", "frame"),
    ("ACRYLIC", "acrylic"),
    ("JOB VALUE", "jobValue"),
    ("TAX", "tax"),
    ("JOB TOTAL", "jobTotal"),
    ("DEPOSIT", "deposit"),
    ("BALANCE", "balance"),
    ("MATERIAL", "materialJB2"),
    ("DIRECT LABOR", "directLabor"),
    ("REMEASURE", "remeasure"),
    ("STATUS", "status"),
    ("GPMA", "gpma"),   # before GPM so contains-match doesn't shadow it
    ("GPM", "gpm"),
    ("ADDRESS", "address"),
]
DATE_FIELDS = {"installDate", "soldDate"}
MONEY_FIELDS = {"jobValue", "tax", "jobTotal", "deposit", "balance",
                "materialJB2", "directLabor"}


def find_header_row(values: list, max_scan: int = 15):
    """Locate the Job List header row: the row matching >=4 known tokens."""
    tokens = ["INSTALL DATE", "SOLD DATE", "SURNAME", "JOB VALUE", "STATUS", "ADDRESS"]
    for idx, row in enumerate(values[:max_scan]):
        cells = [str(c).upper() for c in row if c not in (None, "")]
        hits = sum(1 for t in tokens if any(t in c for c in cells))
        if hits >= 4:
            return idx
    raise RuntimeError("Could not locate the Job List header row (scanned first 15 rows)")


def build_column_map(header_row: list) -> dict:
    colmap = {}
    claimed = set()
    for token, field in RES_HEADER_MAP:
        for ci, cell in enumerate(header_row):
            if ci in claimed or cell in (None, ""):
                continue
            if token in str(cell).upper():
                colmap[field] = ci
                claimed.add(ci)
                break
    # CODE column carries the ZZ/Z/X/R/I codes; its header is the "JOB LIST A-Z"
    # title cell, so map it as the column immediately left of INSTALL DATE.
    if "installDate" in colmap and colmap["installDate"] > 0:
        colmap["code"] = colmap["installDate"] - 1
    missing = [f for _, f in RES_HEADER_MAP
               if f in ("installDate", "soldDate", "surname", "jobValue", "status")
               and f not in colmap]
    if missing:
        raise RuntimeError(f"Job List header found but key columns missing: {missing}")
    return colmap


def normalize_residential(sheet: dict) -> dict:
    values = sheet["values"]
    h = find_header_row(values)
    colmap = build_column_map(values[h])

    def cell(row, field):
        ci = colmap.get(field)
        if ci is None or ci >= len(row):
            return None
        return row[ci]

    jobs = []
    for row in values[h + 1:]:
        if row_is_empty(row):
            continue
        job = {}
        for _, field in RES_HEADER_MAP + [("", "code")]:
            raw = cell(row, field)
            if field in DATE_FIELDS:
                job[field] = to_iso_date(raw)
            elif field in MONEY_FIELDS:
                job[field] = to_money(raw)
            else:
                job[field] = to_str(raw)
        # skip decoration rows (spacers, totals) with no identity and no value
        if not job.get("surname") and job.get("jobValue") is None \
           and not job.get("installDate") and not job.get("soldDate"):
            continue
        jobs.append(job)

    return {
        "headerRowIndex": h + sheet["firstRow"],  # 1-based sheet row of headers
        "columnMap": colmap,
        "jobs": jobs,
        "summary": residential_summary(jobs),
    }


def month_key(iso: str) -> str:
    return iso[:7]


def is_closed_status(status) -> bool:
    s = (status or "").lower()
    return "finished" in s or "billed" in s


def residential_summary(jobs: list) -> dict:
    """Owner-approved view-model rules (REFRESH-runbook §3), computed over the
    FULL job list (hidden rows included)."""
    today = date.today().isoformat()
    sold, installed = {}, {}
    backlog_jobs, unscheduled_jobs = [], []
    zz_blank_status, past_install_open = [], []

    for j in jobs:
        jv = j.get("jobValue") or 0.0
        if j.get("soldDate"):
            m = sold.setdefault(month_key(j["soldDate"]), {"amount": 0.0, "jobs": 0})
            m["amount"] += jv
            m["jobs"] += 1
        if j.get("installDate") and is_closed_status(j.get("status")):
            m = installed.setdefault(month_key(j["installDate"]), {"amount": 0.0, "jobs": 0})
            m["amount"] += jv
            m["jobs"] += 1

        code = (j.get("code") or "").strip().upper()
        blank_status = not (j.get("status") or "").strip()
        if blank_status and code != "ZZ":
            backlog_jobs.append(j)
            if not j.get("installDate"):
                unscheduled_jobs.append(j)
            if j.get("installDate") and j["installDate"] < today:
                past_days = (date.today() - date.fromisoformat(j["installDate"])).days
                if past_days > 30:
                    past_install_open.append(
                        {"surname": j.get("surname"), "installDate": j["installDate"],
                         "code": code or None, "jobValue": jv}
                    )
        if code == "ZZ" and blank_status:
            zz_blank_status.append({"surname": j.get("surname"), "jobValue": jv})

    def money_bucket(js):
        return {"amount": round(sum(j.get("jobValue") or 0 for j in js), 2), "jobs": len(js)}

    return {
        "soldByMonth": {k: {"amount": round(v["amount"], 2), "jobs": v["jobs"]}
                        for k, v in sorted(sold.items())},
        "installedRevenueByMonth": {k: {"amount": round(v["amount"], 2), "jobs": v["jobs"]}
                                    for k, v in sorted(installed.items())},
        "backlog": money_bucket(backlog_jobs),
        "unscheduled": money_bucket(unscheduled_jobs),
        "hygiene": {
            "zzBlankStatus": zz_blank_status,
            "zzBlankStatusTotal": round(sum(x["jobValue"] or 0 for x in zz_blank_status), 2),
            "pastInstallOpen30d": past_install_open,
        },
    }


# ---------------------------------------------------------------------------
# Wholesale tabs
# ---------------------------------------------------------------------------

def labeled_rows(values: list) -> list:
    """PPT Metrics-style output: label (first non-empty cell) + remaining cells.
    Downstream parsing matches on label TEXT, never position (runbook §3b)."""
    out = []
    for row in values:
        if row_is_empty(row):
            continue
        cells = [c for c in row if c not in (None, "")]
        label = str(cells[0]).strip()
        rest = cells[1:]
        out.append({"label": label, "values": rest})
    return out


def parse_kpi_weekly(values: list) -> list:
    """KPI Dashboard weekly series: header row contains 'Week Ending'."""
    header_idx, cols = None, {}
    wanted = {"WEEK ENDING": "weekEnding", "SALES": "sales", "TRAILING": "trailing4wk",
              "INVOICE": "invoices", "REVENUE": "revenueShipped"}
    for idx, row in enumerate(values):
        uppers = [str(c).upper() if c not in (None, "") else "" for c in row]
        if any("WEEK ENDING" in u for u in uppers):
            header_idx = idx
            for ci, u in enumerate(uppers):
                for token, field in wanted.items():
                    if token in u and field not in cols:
                        cols[field] = ci
            break
    if header_idx is None or "weekEnding" not in cols:
        return []
    series = []
    for row in values[header_idx + 1:]:
        we = to_iso_date(row[cols["weekEnding"]] if cols["weekEnding"] < len(row) else None)
        if not we:
            continue
        rec = {"weekEnding": we}
        for field in ("sales", "trailing4wk", "revenueShipped"):
            ci = cols.get(field)
            rec[field] = to_money(row[ci]) if ci is not None and ci < len(row) else None
        ci = cols.get("invoices")
        raw = row[ci] if ci is not None and ci < len(row) else None
        rec["invoices"] = int(raw) if isinstance(raw, (int, float)) else None
        series.append(rec)
    return series


def shipper_payload(sheet: dict) -> dict:
    """Capped raw pass-through of The Shipper: headers + non-empty rows,
    date-ish columns converted from Excel serials to ISO strings."""
    values = sheet["values"]
    if not values:
        return {"headers": [], "rows": [], "truncated": sheet["truncated"]}
    headers = [to_str(c) for c in values[0]]
    date_cols = {i for i, hcell in enumerate(headers)
                 if hcell and "DATE" in hcell.upper()}
    rows = []
    for row in values[1:]:
        if row_is_empty(row):
            continue
        out = []
        for i, c in enumerate(row):
            if i in date_cols:
                out.append(to_iso_date(c) or to_str(c))
            elif isinstance(c, float):
                out.append(round(c, 4))
            else:
                out.append(to_str(c) if isinstance(c, str) else c)
        # trim trailing Nones to keep the file lean
        while out and out[-1] is None:
            out.pop()
        rows.append(out)
    return {
        "headers": headers,
        "rows": rows,
        "usedRowCount": sheet["usedRowCount"],
        "readRowCount": sheet["readRowCount"],
        "rowCap": SHIPPER_ROW_CAP,
        "truncated": sheet["truncated"],
    }


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------

def main() -> None:
    required = ["SHAREPOINT_TENANT_ID", "SHAREPOINT_CLIENT_ID", "SHAREPOINT_CLIENT_SECRET",
                "TRACKER_DRIVE_ID", "TRACKER_RES_ITEM_ID", "TRACKER_WS_ITEM_ID"]
    missing = [v for v in required if not os.environ.get(v)]
    if missing:
        print(f"ERROR: missing required env vars: {missing}")
        sys.exit(1)

    token = get_graph_token()
    notes = []

    # --- Residential --------------------------------------------------------
    print("--- Residential Sales Tracker ---")
    res_sheets = list_worksheets(token, RES_ITEM_ID)
    job_sheet = find_sheet(res_sheets, "job list")
    res_read = read_sheet(token, RES_ITEM_ID, job_sheet["name"])
    residential = {
        "sourceFile": "Residential Sales Tracker.xlsx",
        "worksheet": job_sheet["name"],
        "hiddenRowsIncluded": True,  # Graph range reads ignore visibility
        **normalize_residential(res_read),
    }
    print(f"  normalized {len(residential['jobs'])} jobs "
          f"(header row {residential['headerRowIndex']})")

    # --- Wholesale ----------------------------------------------------------
    print("--- Wholesale Sales Tracker ---")
    ws_sheets = list_worksheets(token, WS_ITEM_ID)

    ppt_sheet = find_sheet(ws_sheets, "ppt metrics")
    ppt_read = read_sheet(token, WS_ITEM_ID, ppt_sheet["name"])
    ppt = {"worksheet": ppt_sheet["name"], "labeledRows": labeled_rows(ppt_read["values"])}
    print(f"  PPT Metrics: {len(ppt['labeledRows'])} labeled rows")

    kpi_sheet = find_sheet(ws_sheets, "kpi dashboard")
    kpi_read = read_sheet(token, WS_ITEM_ID, kpi_sheet["name"])
    weekly = parse_kpi_weekly(kpi_read["values"])
    kpi = {"worksheet": kpi_sheet["name"], "weeklySeries": weekly,
           "labeledRows": labeled_rows(kpi_read["values"][:20])}  # header stats live up top
    print(f"  KPI Dashboard: {len(weekly)} weekly records")
    if not weekly:
        notes.append("KPI Dashboard weekly series parse returned 0 rows — check header text")

    shp_sheet = find_sheet(ws_sheets, "shipper")
    shp_read = read_sheet(token, WS_ITEM_ID, shp_sheet["name"], row_cap=SHIPPER_ROW_CAP)
    shipper = {"worksheet": shp_sheet["name"], **shipper_payload(shp_read)}
    print(f"  The Shipper: {len(shipper['rows'])} non-empty rows "
          f"(cap {SHIPPER_ROW_CAP}, truncated={shipper['truncated']})")
    if shipper["truncated"]:
        notes.append(f"Shipper read truncated at {SHIPPER_ROW_CAP} rows "
                     f"(used range reports {shipper['usedRowCount']}) — if this is "
                     f"phantom bloat, re-trim per runbook §3b; if real data, raise "
                     f"SHIPPER_ROW_CAP")

    payload = {
        "generatedAt": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "schemaVersion": 1,
        "residential": residential,
        "wholesale": {
            "sourceFile": "Wholesale Sales Tracker.xlsx",
            "pptMetrics": ppt,
            "kpiDashboard": kpi,
            "shipper": shipper,
        },
        "meta": {"shipperRowCap": SHIPPER_ROW_CAP, "notes": notes},
    }

    out = "tracker-metrics.json"
    with open(out, "w") as f:
        json.dump(payload, f, indent=1)
    print(f"Wrote {out}.")

    if os.environ.get("SKIP_SHAREPOINT_UPLOAD", "").lower() != "true":
        upload_to_sharepoint(out, payload)
    else:
        print("SKIP_SHAREPOINT_UPLOAD=true — local file only.")
    print("Tracker pull complete.")


if __name__ == "__main__":
    main()
