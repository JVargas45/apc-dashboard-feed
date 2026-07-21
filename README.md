# APC Dashboard Feed

Automated data feed for the American Patio Covers KPI dashboard.

A scheduled GitHub Actions workflow pulls accounting data (company info,
invoices, monthly income) from QuickBooks Online for both company realms
and writes `apc-qb.json` / `pro-qb.json` to the internal SharePoint
dashboard folder.

**No financial data is stored in this repository.** The workflow sends
results directly to SharePoint; all credentials live in encrypted GitHub
Actions secrets.

| File | Purpose |
|---|---|
| `qbo_pull.py` | The pull script the workflow runs |
| `.github/workflows/qbo-refresh.yml` | Cron schedule (weekday mornings) + manual trigger |
| `get_refresh_token.py` | One-time local helper for the initial QuickBooks OAuth consent |
| `SETUP-GUIDE.md` | Full setup instructions (Intuit app, Azure app, secrets) |
| `docs/` | Privacy policy & EULA pages, served via GitHub Pages for Intuit's app compliance requirement |

Internal tool — not intended for external use.
