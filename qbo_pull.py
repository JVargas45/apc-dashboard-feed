"""
qbo_pull.py — APC Dashboard QBO feed

Pulls company info, invoice-level detail, and monthly income totals from
QuickBooks Online for BOTH company realms (APC, PRO), writes apc-qb.json /
pro-qb.json, and uploads them to the SharePoint "Dashboard" folder via
Microsoft Graph — the same folder that holds APC-Dashboard.html and
REFRESH-runbook.md.

Runs unattended on a GitHub Actions cron. One Intuit dev app (single
Client ID/Secret) is used for both realms; each realm has its own
refresh token, obtained once via a separate authorization consent
(see SETUP-GUIDE.md).

QBO issues a NEW refresh token on every use ("rotating" refresh tokens,
~100 day validity). This script persists the new token back into the
GitHub repo's secrets after every run so the next scheduled run keeps
working unattended. If that persistence step is not configured
(GH_SECRETS_PAT missing), the script still completes the pull but warns
that the *next* run will fail once the current token is invalidated.

Invoice-level output (customer, doc number, date, amount, balance, line
items) is what the brief calls "invoice-level detail for reconciliation" —
it lets the dashboard refresh cross-check QB-billed amounts against the
Shipper/tracker "shipped" basis, and resolve tracker status ambiguity
(brief: "QB reconciliation is the flagship future hygiene check").

Required environment variables (see SETUP-GUIDE.md for where each comes from):
    QBO_CLIENT_ID, QBO_CLIENT_SECRET          - one Intuit dev app
    QBO_ENVIRONMENT                            - "sandbox" or "production"
    QBO_APC_REALM_ID, QBO_APC_REFRESH_TOKEN    - APC company consent
    QBO_PRO_REALM_ID, QBO_PRO_REFRESH_TOKEN    - PRO company consent
    GH_SECRETS_PAT                             - PAT to rewrite rotated refresh tokens
    SHAREPOINT_TENANT_ID, SHAREPOINT_CLIENT_ID,
    SHAREPOINT_CLIENT_SECRET, SHAREPOINT_SITE_ID,
    SHAREPOINT_DRIVE_ID (optional), SHAREPOINT_FOLDER_PATH

Optional:
    QBO_LOOKBACK_MONTHS   - how far back to pull invoices/income (default 15)
    SKIP_SHAREPOINT_UPLOAD - "true" to write local JSON only (for testing)
"""

import base64
import json
import os
import sys
import time
from datetime import date, timedelta

import requests

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

QBO_ENV = os.environ.get("QBO_ENVIRONMENT", "production")
QBO_API_BASE = (
    "https://sandbox-quickbooks.api.intuit.com"
    if QBO_ENV == "sandbox"
    else "https://quickbooks.api.intuit.com"
)
QBO_TOKEN_URL = "https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer"
QBO_MINOR_VERSION = "73"  # bump periodically per Intuit's changelog

LOOKBACK_MONTHS = int(os.environ.get("QBO_LOOKBACK_MONTHS", "15"))

COMPANIES = [
    {
        "name": "APC",
        "realm_id": os.environ.get("QBO_APC_REALM_ID"),
        "refresh_token": os.environ.get("QBO_APC_REFRESH_TOKEN"),
        "refresh_token_secret_name": "QBO_APC_REFRESH_TOKEN",
        "output_file": "apc-qb.json",
    },
    {
        "name": "PRO",
        "realm_id": os.environ.get("QBO_PRO_REALM_ID"),
        "refresh_token": os.environ.get("QBO_PRO_REFRESH_TOKEN"),
        "refresh_token_secret_name": "QBO_PRO_REFRESH_TOKEN",
        "output_file": "pro-qb.json",
    },
]


# ---------------------------------------------------------------------------
# GitHub secret rotation (so the next scheduled run still has a valid token)
# ---------------------------------------------------------------------------

def encrypt_secret_for_github(public_key_b64: str, secret_value: str) -> str:
    """Encrypt a value for the GitHub 'update repo secret' API (libsodium sealed box)."""
    from nacl import encoding, public  # imported here so a missing PyNaCl only
    # breaks rotation, not the whole script, if someone strips requirements.

    pk = public.PublicKey(public_key_b64.encode("utf-8"), encoding.Base64Encoder())
    sealed_box = public.SealedBox(pk)
    encrypted = sealed_box.encrypt(secret_value.encode("utf-8"))
    return base64.b64encode(encrypted).decode("utf-8")


def update_github_secret(name: str, value: str) -> None:
    pat = os.environ.get("GH_SECRETS_PAT")
    repo = os.environ.get("GITHUB_REPOSITORY")  # auto-set by GitHub Actions
    if not pat or not repo:
        print(
            f"WARNING: cannot persist rotated refresh token for {name} "
            f"(missing GH_SECRETS_PAT or GITHUB_REPOSITORY). The NEXT run will "
            f"fail once this token is invalidated. See SETUP-GUIDE.md."
        )
        return
    headers = {"Authorization": f"Bearer {pat}", "Accept": "application/vnd.github+json"}
    pk_resp = requests.get(
        f"https://api.github.com/repos/{repo}/actions/secrets/public-key",
        headers=headers,
        timeout=30,
    )
    pk_resp.raise_for_status()
    pk_json = pk_resp.json()
    encrypted_value = encrypt_secret_for_github(pk_json["key"], value)
    put_resp = requests.put(
        f"https://api.github.com/repos/{repo}/actions/secrets/{name}",
        headers=headers,
        json={"encrypted_value": encrypted_value, "key_id": pk_json["key_id"]},
        timeout=30,
    )
    put_resp.raise_for_status()
    print(f"Rotated refresh token persisted to GitHub secret {name}.")


# ---------------------------------------------------------------------------
# QBO auth + fetch helpers
# ---------------------------------------------------------------------------

def auth_headers(access_token: str) -> dict:
    return {"Authorization": f"Bearer {access_token}", "Accept": "application/json"}


def raise_with_tid(resp: requests.Response, context: str) -> None:
    """
    Raise on HTTP error, including Intuit's intuit_tid response header in the
    message. The intuit_tid is a per-request trace ID; Intuit support uses it
    to locate a specific failed request on their side, so it must appear in
    the job logs whenever an API call fails.
    """
    if resp.status_code >= 400:
        tid = resp.headers.get("intuit_tid", "n/a")
        raise RuntimeError(
            f"{context}: HTTP {resp.status_code} (intuit_tid={tid}) {resp.text[:500]}"
        )


def refresh_access_token(company: dict) -> str:
    client_id = os.environ["QBO_CLIENT_ID"]
    client_secret = os.environ["QBO_CLIENT_SECRET"]
    basic = base64.b64encode(f"{client_id}:{client_secret}".encode()).decode()
    resp = requests.post(
        QBO_TOKEN_URL,
        headers={
            "Authorization": f"Basic {basic}",
            "Accept": "application/json",
            "Content-Type": "application/x-www-form-urlencoded",
        },
        data={"grant_type": "refresh_token", "refresh_token": company["refresh_token"]},
        timeout=30,
    )
    if resp.status_code != 200:
        tid = resp.headers.get("intuit_tid", "n/a")
        raise RuntimeError(
            f"[{company['name']}] token refresh failed: {resp.status_code} "
            f"(intuit_tid={tid}) {resp.text}"
        )
    tokens = resp.json()
    new_refresh_token = tokens["refresh_token"]
    if new_refresh_token != company["refresh_token"]:
        update_github_secret(company["refresh_token_secret_name"], new_refresh_token)
    return tokens["access_token"]


def fetch_company_info(company: dict, access_token: str) -> dict:
    url = f"{QBO_API_BASE}/v3/company/{company['realm_id']}/companyinfo/{company['realm_id']}"
    resp = requests.get(
        url,
        headers=auth_headers(access_token),
        params={"minorversion": QBO_MINOR_VERSION},
        timeout=30,
    )
    raise_with_tid(resp, f"[{company['name']}] company info")
    ci = resp.json().get("CompanyInfo", {})
    return {
        "companyName": ci.get("CompanyName"),
        "legalName": ci.get("LegalName"),
        "country": ci.get("Country"),
        "fiscalYearStartMonth": ci.get("FiscalYearStartMonth"),
    }


def simplify_invoice(inv: dict) -> dict:
    total = inv.get("TotalAmt", 0) or 0
    balance = inv.get("Balance", 0) or 0
    if balance <= 0:
        status = "Paid"
    elif balance < total:
        status = "Partially Paid"
    else:
        status = "Open"
    lines = []
    for line in inv.get("Line", []):
        if line.get("DetailType") == "SalesItemLineDetail":
            detail = line.get("SalesItemLineDetail", {})
            lines.append(
                {
                    "description": line.get("Description"),
                    "amount": line.get("Amount"),
                    "item": detail.get("ItemRef", {}).get("name"),
                }
            )
    customer_ref = inv.get("CustomerRef", {})
    return {
        "id": inv.get("Id"),
        "docNumber": inv.get("DocNumber"),
        "txnDate": inv.get("TxnDate"),
        "dueDate": inv.get("DueDate"),
        "customer": customer_ref.get("name"),
        "customerId": customer_ref.get("value"),
        "totalAmt": total,
        "balance": balance,
        "status": status,
        "lines": lines,
    }


def fetch_invoices(company: dict, access_token: str, since_date: str) -> list:
    invoices = []
    start_position = 1
    page_size = 200
    url = f"{QBO_API_BASE}/v3/company/{company['realm_id']}/query"
    while True:
        query = (
            f"SELECT * FROM Invoice WHERE TxnDate >= '{since_date}' "
            f"ORDERBY TxnDate STARTPOSITION {start_position} MAXRESULTS {page_size}"
        )
        resp = requests.get(
            url,
            headers=auth_headers(access_token),
            params={"query": query, "minorversion": QBO_MINOR_VERSION},
            timeout=60,
        )
        raise_with_tid(resp, f"[{company['name']}] invoice query")
        batch = resp.json().get("QueryResponse", {}).get("Invoice", [])
        if not batch:
            break
        invoices.extend(simplify_invoice(inv) for inv in batch)
        if len(batch) < page_size:
            break
        start_position += page_size
    return invoices


def find_row_by_label(rows: list, label: str):
    """Recursively search a QBO report's Row tree for a row whose header/summary matches label."""
    for row in rows:
        header_cells = row.get("Header", {}).get("ColData", [])
        if header_cells and header_cells[0].get("value") == label:
            return row
        summary_cells = row.get("Summary", {}).get("ColData", [])
        if summary_cells and summary_cells[0].get("value") == label:
            return row
        nested = row.get("Rows", {}).get("Row", [])
        if nested:
            found = find_row_by_label(nested, label)
            if found:
                return found
    return None


def parse_income_by_month(report: dict) -> list:
    """
    Best-effort parse of the ProfitAndLoss report's 'Total Income' row into
    [{month, totalIncome}, ...]. QBO's report JSON nesting can vary slightly
    by minor version — verify this against a live report during setup and
    adjust the label lookup if needed. The raw report is also included in
    the output under 'profitAndLossRaw' so nothing is lost if parsing misses.
    """
    columns = report.get("Columns", {}).get("Column", [])
    month_labels = [c.get("ColTitle", "") for c in columns]
    row = find_row_by_label(report.get("Rows", {}).get("Row", []), "Total Income")
    if not row:
        return []
    col_data = row.get("Summary", {}).get("ColData") or row.get("ColData") or []
    results = []
    for label, cell in zip(month_labels, col_data):
        if not label or label.strip().lower() == "total":
            continue
        try:
            value = float(cell.get("value") or 0)
        except (TypeError, ValueError):
            value = 0.0
        results.append({"month": label, "totalIncome": value})
    return results


def fetch_income_by_month(company: dict, access_token: str, since_date: str, end_date: str):
    url = f"{QBO_API_BASE}/v3/company/{company['realm_id']}/reports/ProfitAndLoss"
    resp = requests.get(
        url,
        headers=auth_headers(access_token),
        params={
            "start_date": since_date,
            "end_date": end_date,
            "summarize_column_by": "Month",
            "minorversion": QBO_MINOR_VERSION,
        },
        timeout=60,
    )
    raise_with_tid(resp, f"[{company['name']}] ProfitAndLoss report")
    report = resp.json()
    return parse_income_by_month(report), report


# ---------------------------------------------------------------------------
# SharePoint upload (Microsoft Graph, app-only client-credentials)
# ---------------------------------------------------------------------------

def get_graph_token() -> str:
    tenant_id = os.environ["SHAREPOINT_TENANT_ID"]
    client_id = os.environ["SHAREPOINT_CLIENT_ID"]
    client_secret = os.environ["SHAREPOINT_CLIENT_SECRET"]
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    resp = requests.post(
        url,
        data={
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials",
        },
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()["access_token"]


def resolve_drive_id(token: str, site_id: str, library_name: str) -> str:
    """Find the drive (document library) on the site whose name matches library_name."""
    resp = requests.get(
        f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives",
        headers={"Authorization": f"Bearer {token}"},
        timeout=30,
    )
    resp.raise_for_status()
    drives = resp.json().get("value", [])
    for drive in drives:
        if drive.get("name", "").strip().lower() == library_name.strip().lower():
            return drive["id"]
    available = [d.get("name") for d in drives]
    raise RuntimeError(
        f"No document library named '{library_name}' on site. Available: {available}"
    )


def upload_to_sharepoint(filename: str, payload: dict) -> None:
    site_id = os.environ["SHAREPOINT_SITE_ID"]
    drive_id = os.environ.get("SHAREPOINT_DRIVE_ID")
    library_name = os.environ.get("SHAREPOINT_LIBRARY_NAME")
    folder_path = os.environ.get("SHAREPOINT_FOLDER_PATH", "Dashboard").strip("/")
    token = get_graph_token()
    body = json.dumps(payload, indent=2).encode("utf-8")

    if not drive_id and library_name:
        drive_id = resolve_drive_id(token, site_id, library_name)

    if drive_id:
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{folder_path}/{filename}:/content"
    else:
        url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{folder_path}/{filename}:/content"

    resp = requests.put(
        url,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        data=body,
        timeout=60,
    )
    resp.raise_for_status()
    print(f"Uploaded {filename} to SharePoint ({len(body)} bytes).")


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------

def process_company(company: dict) -> None:
    print(f"--- {company['name']} ---")
    if not company["realm_id"] or not company["refresh_token"]:
        raise RuntimeError(
            f"Missing realm_id or refresh_token for {company['name']}. Check GitHub secrets."
        )

    access_token = refresh_access_token(company)
    since = (date.today() - timedelta(days=30 * LOOKBACK_MONTHS)).isoformat()
    today = date.today().isoformat()

    company_info = fetch_company_info(company, access_token)
    invoices = fetch_invoices(company, access_token, since)

    try:
        income_by_month, raw_report = fetch_income_by_month(company, access_token, since, today)
    except Exception as exc:  # noqa: BLE001 - want to keep going even if P&L pull fails
        print(f"WARNING: income-by-month pull failed for {company['name']}: {exc}")
        income_by_month, raw_report = [], None

    payload = {
        "realm": company["name"],
        "realmId": company["realm_id"],
        "environment": QBO_ENV,
        "generatedAt": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "companyInfo": company_info,
        "incomeByMonth": income_by_month,
        "invoices": invoices,
        "profitAndLossRaw": raw_report,
        "meta": {
            "invoiceCount": len(invoices),
            "lookbackMonths": LOOKBACK_MONTHS,
            "pulledThrough": today,
        },
    }

    with open(company["output_file"], "w") as f:
        json.dump(payload, f, indent=2)
    print(f"Wrote {company['output_file']} ({len(invoices)} invoices).")

    if os.environ.get("SKIP_SHAREPOINT_UPLOAD", "").lower() != "true":
        upload_to_sharepoint(company["output_file"], payload)
    else:
        print("SKIP_SHAREPOINT_UPLOAD=true — local file only.")


def main() -> None:
    required = ["QBO_CLIENT_ID", "QBO_CLIENT_SECRET"]
    missing = [v for v in required if not os.environ.get(v)]
    if missing:
        print(f"ERROR: missing required env vars: {missing}")
        sys.exit(1)

    failures = []
    for company in COMPANIES:
        try:
            process_company(company)
        except Exception as exc:  # noqa: BLE001 - report and keep processing the other realm
            print(f"ERROR processing {company['name']}: {exc}")
            failures.append(company["name"])

    if failures:
        print(f"Completed with failures: {failures}")
        sys.exit(1)
    print("QBO pull complete for all companies.")


if __name__ == "__main__":
    main()
