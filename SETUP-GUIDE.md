# QBO API Feed — Setup Guide

This is the "QBO API feed (GitHub Actions cron → SharePoint JSON)" item from
the project brief's status board. It's a one-time setup (roughly 45-60
minutes), done by whoever owns the Intuit, Azure/365 admin, and GitHub repo
accounts — likely three different logins, possibly the same person.

## What this builds

```
GitHub Actions (cron, weekday mornings)
        |
        |  refreshes access token, one per realm
        v
QuickBooks Online API  --- two realms, one Intuit dev app ---
   APC company                              PRO company
        |                                        |
        v                                        v
  apc-qb.json                              pro-qb.json
  (company info, invoices, monthly income)
        |                                        |
        +--------------------+-------------------+
                             v
              Microsoft Graph (app-only)
                             v
        SharePoint: "KPIs and Reporting" -> Dashboard/
        (same folder as APC-Dashboard.html and REFRESH-runbook.md)
```

Each JSON file contains company info, a monthly income total (for the KPI
tiles), and full invoice-level detail — customer, date, amount, balance,
line items — going back `QBO_LOOKBACK_MONTHS` (default 15 months). The
invoice list is what lets the dashboard refresh reconcile QB-billed amounts
against the wholesale tracker's "shipped" basis, per the brief's flagship
hygiene check.

**Files in this delivery:**
- `qbo_pull.py` — the script GitHub Actions runs
- `get_refresh_token.py` — one-time local helper for the initial OAuth consent (optional — see Step 3)
- `requirements.txt` — Python dependencies
- `qbo-refresh.yml` — the GitHub Actions workflow (copy to `.github/workflows/qbo-refresh.yml`)

---

## Step 1 — Create the Intuit developer app

1. Go to **https://developer.intuit.com** and sign in (create an Intuit
   developer account if needed — it's free, separate from your regular QBO login).
2. **Dashboard → Create an app → QuickBooks Online and Payments**.
3. Give it a name (e.g. "APC Dashboard Feed"). Select the **Accounting** scope.
4. Once created, open **Keys & OAuth** for the app:
   - You'll see **Development** keys (Client ID/Secret for sandbox testing)
     and a **Production** tab.
   - Under **Production**, add a **Redirect URI**. For the simplest path,
     use Intuit's own OAuth Playground redirect:
     `https://developer.intuit.com/v2/OAuth2Playground/RedirectUrl`
     (This lets you use Intuit's official Playground tool in Step 3 with
     zero code. If you'd rather use your own redirect URI, that's fine too —
     just keep it consistent with what you enter in Step 3.)
   - Requesting **Production keys** may ask you to fill in some basic app
     info (this is Intuit's standard flow for apps that will only ever be
     used by the app owner's own companies, not a public listing — it's
     usually a short form, not a full review). Follow whatever Intuit's UI
     currently asks for.
5. Copy the **Production** Client ID and Client Secret somewhere safe —
   you'll paste these into GitHub secrets in Step 5. This ONE app/ONE
   key pair is shared by both APC and PRO.

*(You can do everything below against the Development/sandbox keys first if
you want to test the pipeline before touching real company data — just set
`QBO_ENVIRONMENT=sandbox` and use sandbox company connections. Switch to
Production keys + `QBO_ENVIRONMENT=production` when you're ready to go live.)*

## Step 2 — Understand what you're about to authorize

QBO's OAuth model: one dev app can connect to many companies, but each
company connection needs its own **authorization consent** — i.e. someone
has to log into QuickBooks *as that company* and click "Connect." That's
why the brief calls for two consents (APC, then PRO) against the one app.

Each consent produces a **realm ID** (QuickBooks' company identifier) and a
**refresh token**. The refresh token is what lets the script get new access
tokens indefinitely without a human logging in again — as long as it's used
at least once every ~100 days, which the weekday cron easily satisfies.

## Step 3 — Authorize APC, then PRO

Do this once per company. Two options — pick whichever's easier:

### Option A — Intuit's OAuth 2.0 Playground (no code, recommended)

1. Go to **https://developer.intuit.com/app/developer/playground**.
2. Select your app, select the **Accounting** scope.
3. Click **Get authorization code**, sign into QuickBooks **as the APC
   company**, approve access.
4. The Playground shows you the **Realm ID** and lets you click **Get
   tokens** to reveal the **refresh token** directly on the page.
5. Copy the Realm ID and refresh token → these become
   `QBO_APC_REALM_ID` and `QBO_APC_REFRESH_TOKEN`.
6. Repeat, this time signing in **as the PRO company**, to get
   `QBO_PRO_REALM_ID` and `QBO_PRO_REFRESH_TOKEN`.

### Option B — `get_refresh_token.py` (scripted, same result)

```bash
pip install requests --break-system-packages
python get_refresh_token.py
```

Follow the prompts (it prints a URL to open in a browser, then asks you to
paste back the redirected URL). Run it twice — once signed in as APC, once
as PRO.

Either way, by the end of this step you have **four values**:
`QBO_APC_REALM_ID`, `QBO_APC_REFRESH_TOKEN`, `QBO_PRO_REALM_ID`,
`QBO_PRO_REFRESH_TOKEN`.

## Step 4 — Set up SharePoint write access (Microsoft Graph)

The script needs its own app registration to write files — separate from
your normal 365 login, and separate from Claude's read-only 365 connector.

1. In **Azure Portal → Microsoft Entra ID → App registrations → New
   registration**. Name it something like "APC Dashboard QBO Writer."
   Single tenant is fine.
2. **API permissions → Add a permission → Microsoft Graph → Application
   permissions** → add `Sites.Selected` (preferred, scoped to just this one
   site) or `Sites.ReadWrite.All` (simpler, broader — fine for an internal
   tool). Click **Grant admin consent**.
   - If you use `Sites.Selected`, you also need to grant this specific app
     `write` access to the one SharePoint site — this requires a separate
     Graph call (`POST /sites/{site-id}/permissions`) that a 365 admin runs
     once. Happy to script that too if you want to go the scoped route
     instead of `Sites.ReadWrite.All`.
3. **Certificates & secrets → New client secret**. Copy the value
   immediately (Azure only shows it once) → `SHAREPOINT_CLIENT_SECRET`.
4. From the app's **Overview** page, copy:
   - **Application (client) ID** → `SHAREPOINT_CLIENT_ID`
   - **Directory (tenant) ID** → `SHAREPOINT_TENANT_ID`
5. Find your **Site ID**: with an admin account, open this in a browser
   (or Graph Explorer at https://developer.microsoft.com/graph/graph-explorer):
   ```
   https://graph.microsoft.com/v1.0/sites/{yourtenant}.sharepoint.com:/sites/{site-name}
   ```
   The response's `id` field is your `SHAREPOINT_SITE_ID`. (`SHAREPOINT_DRIVE_ID`
   is optional — leave it unset and the script will resolve the default
   document library from the site ID.)
6. Confirm the exact folder path under the site's document library that
   matches "KPIs and Reporting → Dashboard" — the workflow file already
   assumes `KPIs and Reporting/Dashboard`, but double-check against what
   REFRESH-runbook.md uses for `APC-Dashboard.html`, and adjust the
   `SHAREPOINT_FOLDER_PATH` value in `qbo-refresh.yml` if it differs.

## Step 5 — Create a GitHub PAT for secret rotation

Because QBO rotates refresh tokens on every use, the workflow needs
permission to update its own secrets after each run.

1. **GitHub → Settings → Developer settings → Fine-grained tokens → Generate new token.**
2. Scope it to **this one repository only**.
3. Under **Repository permissions**, set **Secrets** to **Read and write**.
4. Set an expiration you're comfortable renewing (fine-grained PATs max out
   at 1 year) — put a reminder on your calendar to rotate it.
5. Copy the token → this becomes the `GH_SECRETS_PAT` secret itself (yes,
   a secret that manages other secrets — that's expected here).

## Step 6 — Add everything to the repo

1. Add `qbo_pull.py` and `requirements.txt` to the repo root (or wherever
   makes sense for this codebase).
2. Add `qbo-refresh.yml` at `.github/workflows/qbo-refresh.yml`.
3. **Repo → Settings → Secrets and variables → Actions → New repository
   secret**, and add each of these:

| Secret name | Value from |
|---|---|
| `QBO_CLIENT_ID` | Step 1 |
| `QBO_CLIENT_SECRET` | Step 1 |
| `QBO_APC_REALM_ID` | Step 3 |
| `QBO_APC_REFRESH_TOKEN` | Step 3 |
| `QBO_PRO_REALM_ID` | Step 3 |
| `QBO_PRO_REFRESH_TOKEN` | Step 3 |
| `GH_SECRETS_PAT` | Step 5 |
| `SHAREPOINT_TENANT_ID` | Step 4 |
| `SHAREPOINT_CLIENT_ID` | Step 4 |
| `SHAREPOINT_CLIENT_SECRET` | Step 4 |
| `SHAREPOINT_SITE_ID` | Step 4 |
| `SHAREPOINT_DRIVE_ID` | Step 4 (optional) |

Nothing here gets committed to git history — the script writes the two
JSON files to the runner's temp workspace and pushes them straight to
SharePoint; they never touch the repo itself.

## Step 7 — Test before trusting the cron

1. **Repo → Actions → QBO Refresh → Run workflow** (this works because of
   the `workflow_dispatch` trigger in the yml).
2. Watch the log. First-run things to check:
   - Both companies process without error.
   - The line `Rotated refresh token persisted to GitHub secret ...` appears
     for each company (confirms Step 5 is wired up correctly).
   - `apc-qb.json` / `pro-qb.json` show up in the SharePoint Dashboard
     folder with a recent timestamp.
3. Open one of the JSON files and sanity check: company name matches,
   invoice count looks plausible, a handful of invoices have the right
   customer/amount/date.
4. If `incomeByMonth` comes back empty, the P&L report parsing likely needs
   a small adjustment — see the comment above `parse_income_by_month` in
   `qbo_pull.py`. The raw report is preserved under `profitAndLossRaw` in
   the JSON either way, so nothing is lost while that gets tuned.
5. Once a manual run looks right, leave the cron schedule as-is (weekday
   mornings) — no further action needed.

## After this is live

Per the brief, the next step is updating `REFRESH-runbook.md` so the
dashboard refresh process pulls from `apc-qb.json` / `pro-qb.json` instead
of showing QB data as sample-badged. That's a small edit to the runbook,
not covered here — flag it when you're ready and it can be done as part of
a normal refresh pass.

## Troubleshooting

- **401 on token refresh** — the refresh token was invalidated (used past
  its ~100 day life, or revoked in Intuit's connected-apps list). Re-run
  Step 3 for that company.
- **403 from SharePoint upload** — admin consent wasn't granted in Step 4,
  or (if using `Sites.Selected`) the app wasn't separately granted access
  to this specific site.
- **Workflow succeeds but next scheduled run fails on token refresh** — the
  `GH_SECRETS_PAT` is missing, expired, or scoped to the wrong repo/permission
  — check the WARNING line in the previous run's log, it calls this out
  explicitly.
- **Rate limits** — QBO's API allows 500 requests/minute per app across all
  realms in production; a weekday-morning run for two companies is nowhere
  close to that ceiling.
