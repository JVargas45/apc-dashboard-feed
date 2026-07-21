# QBO API Feed — Setup Guide (v2, updated Jul 21 2026)

This is the "QBO API feed (GitHub Actions cron → SharePoint JSON)" item from
the project brief's status board. v2 reflects lessons learned during actual
setup — Intuit's production-key unlock process, the irreversible-scope trap,
and the GitHub Pages compliance pages, none of which v1 covered.

## Progress checklist (as of Jul 21, 2026)

- [x] GitHub repo `apc-dashboard-feed` created (public, JVargas45)
- [x] Repo files uploaded: script, workflow, requirements, docs
- [x] GitHub Pages live: privacy.html, eula.html, index.html
- [x] Intuit app v2 created — **accounting scope ONLY**
- [x] `qbo_pull.py` updated to log `intuit_tid` on API errors (make sure the
      repo copy is the updated one)
- [~] Intuit production-key unlock: App details + Compliance questionnaire (in progress)
- [ ] Production redirect URI added (unlocks after Compliance is done)
- [ ] OAuth consents: APC, then PRO (Step 3 below)
- [ ] Azure app for SharePoint write (Step 4)
- [ ] GitHub PAT for secret rotation (Step 5)
- [ ] GitHub secrets added (Step 6)
- [ ] First manual test run (Step 7)
- [ ] Optional but recommended: sandbox dry-run before production consents

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

**Repo layout** (github.com/JVargas45/apc-dashboard-feed):
- `qbo_pull.py` — the script GitHub Actions runs (v2: logs intuit_tid on errors)
- `.github/workflows/qbo-refresh.yml` — weekday-morning cron + manual trigger
- `get_refresh_token.py` — one-time local helper for OAuth consent (optional; Playground does the same)
- `requirements.txt` — Python dependencies
- `docs/` — privacy.html, eula.html, index.html, served via GitHub Pages for Intuit compliance
- `SETUP-GUIDE.md` — this file

## Step 1 — Intuit developer app  ✅ done, lessons recorded

What we learned the hard way:

- **Scopes are permanent.** Once a scope is added to an app it can never be
  removed, and scopes apply to both dev and production credentials. The
  first app accidentally included `com.intuit.quickbooks.payment`; the fix
  was creating a fresh app with **only `com.intuit.quickbooks.accounting`**.
  Never add payments to this app.
- **Production keys are gated.** The Production Client ID/Secret and the
  Production Redirect URI section are locked until two unlock tasks are
  complete on the Keys & Credentials page: **App details** (~8 min) and
  **Compliance** (~40 min questionnaire). Development keys and their
  redirect URIs are available immediately with no unlock.
- **Compliance needs public URLs.** The questionnaire requires a publicly
  reachable Privacy Policy URL, EULA URL, host domain, launch URL,
  disconnect URL, and connect URL. Ours are served from GitHub Pages:
  - Privacy: `https://jvargas45.github.io/apc-dashboard-feed/privacy.html`
  - EULA: `https://jvargas45.github.io/apc-dashboard-feed/eula.html`
  - Host domain: `jvargas45.github.io`
  - Launch / Disconnect / Connect URL (all three):
    `https://jvargas45.github.io/apc-dashboard-feed/`
- **Hosting answers**: hosted on GitHub (Actions runners on Microsoft Azure
  + GitHub Pages CDN), United States, no static IP (GitHub Pages resolves
  to 185.199.108.153 and siblings if a literal IP is forced).
- **Questionnaire stances** (all truthful for this app): no webhooks, no
  CDC, no WebSockets; daily API calls (weekday cron); ~4–6 calls per
  company per run; no retry loops on auth failure; tokens refreshed
  programmatically once per run; intuit_tid captured in error logs; logs in
  GitHub Actions (~90-day retention, shareable); no credentials or QBO data
  in logs; no version-specific QBO features used (core objects only).

Once Compliance completes and Production unlocks:

1. **Keys & Credentials → Production tab** → copy the Production
   **Client ID** and **Client Secret** (→ GitHub secrets in Step 6).
2. In the Production **Redirect URIs** section, add:
   `https://developer.intuit.com/v2/OAuth2Playground/RedirectUrl`

One app / one key pair serves both APC and PRO.

## Step 2 — Understand what you're about to authorize

QBO's OAuth model: one dev app can connect to many companies, but each
company connection needs its own **authorization consent** — someone logs
into QuickBooks *as that company* and clicks "Connect." Hence two consents
(APC, then PRO) against the one app.

Each consent produces a **realm ID** (QuickBooks' company identifier) and a
**refresh token**. The refresh token lets the script get new access tokens
indefinitely without a human logging in again — provided it's used at least
once every ~100 days, which the weekday cron easily satisfies. Note QBO
*rotates* refresh tokens: every use returns a new one, which the script
automatically persists back into GitHub secrets (that's what the
`GH_SECRETS_PAT` in Step 5 is for).

## Step 3 — Authorize APC, then PRO

Requires production keys (Step 1 complete). Do this once per company.
Two options:

### Option A — Intuit's OAuth 2.0 Playground (no code, recommended)

1. Go to **https://developer.intuit.com/app/developer/playground**.
2. Select the app, select the **Accounting** scope.
3. Click **Get authorization code**, sign into QuickBooks **as the APC
   company**, approve access.
4. The Playground shows the **Realm ID**; click **Get tokens** to reveal
   the **refresh token**.
5. Copy both → `QBO_APC_REALM_ID` and `QBO_APC_REFRESH_TOKEN`.
6. Repeat signed in **as the PRO company** → `QBO_PRO_REALM_ID` and
   `QBO_PRO_REFRESH_TOKEN`.

**Timing note:** a refresh token obtained here is only guaranteed while
unused for ~100 days, but more importantly the FIRST use by the script
rotates it. Do the consents when you're ready to add secrets and run the
workflow soon after — don't harvest tokens weeks before wiring things up.

### Option B — `get_refresh_token.py` (scripted, same result)

```bash
pip install requests
python get_refresh_token.py
```

Run twice — once signed in as APC, once as PRO.

## Step 4 — SharePoint write access (Microsoft Graph)

The script needs its own app registration to write files — separate from
your normal 365 login, and separate from Claude's read-only 365 connector.

1. **Azure Portal → Microsoft Entra ID → App registrations → New
   registration**. Name: "APC Dashboard QBO Writer." Single tenant.
2. **API permissions → Add a permission → Microsoft Graph → Application
   permissions** → add `Sites.Selected` (preferred, scoped to one site) or
   `Sites.ReadWrite.All` (simpler, broader — acceptable for an internal
   tool). Click **Grant admin consent**.
   - `Sites.Selected` additionally requires granting this app `write` on
     the one SharePoint site via a one-time Graph call
     (`POST /sites/{site-id}/permissions`) by a 365 admin. Ask Claude to
     script it if going this route.
3. **Certificates & secrets → New client secret**. Copy immediately (shown
   once) → `SHAREPOINT_CLIENT_SECRET`. Note its expiry date on a calendar.
4. From **Overview**: Application (client) ID → `SHAREPOINT_CLIENT_ID`;
   Directory (tenant) ID → `SHAREPOINT_TENANT_ID`.
5. **Site ID**: in Graph Explorer
   (https://developer.microsoft.com/graph/graph-explorer) run:
   `https://graph.microsoft.com/v1.0/sites/{yourtenant}.sharepoint.com:/sites/{site-name}`
   — the response `id` → `SHAREPOINT_SITE_ID`. (`SHAREPOINT_DRIVE_ID` is
   optional; unset, the script uses the site's default document library.)
6. Confirm the folder path for "KPIs and Reporting → Dashboard." The
   workflow assumes `KPIs and Reporting/Dashboard` — check against where
   `APC-Dashboard.html` actually lives and adjust `SHAREPOINT_FOLDER_PATH`
   in `qbo-refresh.yml` if needed.

## Step 5 — GitHub PAT for secret rotation

QBO rotates refresh tokens on every use, so the workflow must be able to
update its own secrets after each run.

1. **GitHub → Settings → Developer settings → Fine-grained tokens →
   Generate new token.**
2. Scope: **only the `apc-dashboard-feed` repository**.
3. Repository permissions → **Secrets: Read and write**.
4. Expiration: up to 1 year — calendar a renewal reminder.
5. Copy the token → becomes the `GH_SECRETS_PAT` secret.

## Step 6 — Add GitHub secrets

Repo files are already in place. **Repo → Settings → Secrets and variables
→ Actions → New repository secret** for each:

| Secret name | Value from |
|---|---|
| `QBO_CLIENT_ID` | Step 1 (Production) |
| `QBO_CLIENT_SECRET` | Step 1 (Production) |
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

No QuickBooks data or credentials ever enter the repo or its history — the
JSON goes straight from the runner to SharePoint, and secrets live only in
GitHub's encrypted store.

## Optional Step 6.5 — Sandbox dry-run (recommended)

Validates the whole pipeline before real books are involved, and makes the
"have you tested error handling" compliance answer concretely true:

1. Keys & Credentials → **Development** tab → copy dev Client ID/Secret;
   add the Playground redirect URI under Development Redirect URIs.
2. Find your **sandbox company** (auto-created with the developer account,
   under the API/Sandbox menu).
3. Do one Playground consent against the sandbox company.
4. Temporarily set secrets: dev `QBO_CLIENT_ID`/`QBO_CLIENT_SECRET`,
   sandbox realm/token in the APC slots, and add a repo **variable or
   temporary edit** `QBO_ENVIRONMENT: sandbox` in the workflow env. Set
   `SKIP_SHAREPOINT_UPLOAD: "true"` if Step 4 isn't done yet.
5. Actions → QBO Refresh → **Run workflow**. Expect PRO to fail (no
   sandbox token in its slots) — that's fine, it also proves per-company
   error isolation works.
6. Revert env to production values afterward.

## Step 7 — Test before trusting the cron

1. **Repo → Actions → QBO Refresh → Run workflow.**
2. In the log, confirm:
   - Both companies process without error.
   - `Rotated refresh token persisted to GitHub secret ...` appears for
     each company (proves Step 5 works — without this the NEXT run fails).
   - `apc-qb.json` / `pro-qb.json` appear in the SharePoint Dashboard
     folder with fresh timestamps.
3. Open a JSON: company name right, invoice count plausible, spot-check a
   few invoices.
4. If `incomeByMonth` is empty, the P&L parse needs a tweak — see the
   comment above `parse_income_by_month` in `qbo_pull.py`; the raw report
   is preserved under `profitAndLossRaw` so nothing is lost meanwhile.
5. Good manual run → leave the weekday cron alone; done.

## After this is live

Update `REFRESH-runbook.md` (on SharePoint) so the dashboard refresh reads
`apc-qb.json` / `pro-qb.json` instead of sample-badged QB data, and flip
the QB status dot from 🔴 sample. Small edit, done as part of a normal
refresh pass — flag it when ready.

## Troubleshooting

- **401 on token refresh** — refresh token invalidated (unused past ~100
  days, revoked in Intuit's connected apps, or a rotated token failed to
  persist). Re-run Step 3 for that company.
- **`invalid_grant`** — same causes/remedy as above; also check the right
  realm's token is in the right secret slot.
- **403 from SharePoint upload** — admin consent missing in Step 4, or
  (with `Sites.Selected`) the per-site grant wasn't done.
- **Run succeeds but the NEXT run fails auth** — token rotation didn't
  persist: `GH_SECRETS_PAT` missing/expired/mis-scoped. The prior run's log
  contains an explicit WARNING when this is the case.
- **Errors include `intuit_tid=`** — that's Intuit's per-request trace ID;
  quote it if you ever contact Intuit developer support.
- **Rate limits** — QBO allows 500 req/min per app in production; this
  feed uses ~10/day. Not a concern.
