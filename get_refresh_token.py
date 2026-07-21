"""
get_refresh_token.py — one-time, run-on-your-own-laptop helper

Completes Intuit's OAuth2 authorization-code flow for ONE company realm and
prints the refresh token + realm ID you'll save as GitHub secrets.

You need to run this TWICE with the SAME Client ID/Secret (one Intuit dev
app covers both realms per the brief) but logged into the correct company
each time:
    1st run -> sign in as APC       -> save QBO_APC_REALM_ID / QBO_APC_REFRESH_TOKEN
    2nd run -> sign in as PRO       -> save QBO_PRO_REALM_ID / QBO_PRO_REFRESH_TOKEN

NOTE: Intuit also ships an official point-and-click tool that does the same
thing without running any code — the OAuth 2.0 Playground:
    https://developer.intuit.com/app/developer/playground
Use whichever you find easier; SETUP-GUIDE.md covers both. This script
exists so the whole process is scripted/reproducible if you ever need to
redo a consent.

Usage:
    pip install requests --break-system-packages
    python get_refresh_token.py
"""

import base64
import urllib.parse

import requests

AUTH_BASE = "https://appcenter.intuit.com/connect/oauth2"
TOKEN_URL = "https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer"
DEFAULT_REDIRECT_URI = "https://developer.intuit.com/v2/OAuth2Playground/RedirectUrl"
SCOPE = "com.intuit.quickbooks.accounting"


def main():
    client_id = input("QBO_CLIENT_ID: ").strip()
    client_secret = input("QBO_CLIENT_SECRET: ").strip()
    redirect_uri = (
        input(f"Redirect URI registered on your Intuit app [{DEFAULT_REDIRECT_URI}]: ").strip()
        or DEFAULT_REDIRECT_URI
    )

    params = {
        "client_id": client_id,
        "response_type": "code",
        "scope": SCOPE,
        "redirect_uri": redirect_uri,
        "state": "apc-dashboard-setup",
    }
    auth_url = f"{AUTH_BASE}?{urllib.parse.urlencode(params)}"

    print("\n1. Open this URL in a browser, signed in as the QuickBooks company")
    print("   you want to authorize right now (do APC, then re-run for PRO):\n")
    print(f"   {auth_url}\n")
    print("2. Approve access. Intuit redirects you to the URL above with 'code'")
    print("   and 'realmId' added to the query string — the page itself may show")
    print("   an error, that's fine, you only need the URL from the address bar.\n")

    redirected_url = input("Paste the FULL redirected URL here: ").strip()
    parsed = urllib.parse.urlparse(redirected_url)
    qs = urllib.parse.parse_qs(parsed.query)
    code = qs.get("code", [None])[0]
    realm_id = qs.get("realmId", [None])[0]
    if not code or not realm_id:
        print("\nCould not find 'code' and 'realmId' in that URL. Check and retry.")
        return

    basic = base64.b64encode(f"{client_id}:{client_secret}".encode()).decode()
    resp = requests.post(
        TOKEN_URL,
        headers={
            "Authorization": f"Basic {basic}",
            "Accept": "application/json",
            "Content-Type": "application/x-www-form-urlencoded",
        },
        data={"grant_type": "authorization_code", "code": code, "redirect_uri": redirect_uri},
        timeout=30,
    )
    resp.raise_for_status()
    tokens = resp.json()

    print("\n--- SUCCESS ---")
    print(f"Realm ID:      {realm_id}")
    print(f"Refresh token: {tokens['refresh_token']}")
    print(f"(access token expires in {tokens['expires_in']}s — not needed after setup, discard it)")
    print("\nSave the Realm ID and Refresh token as GitHub secrets now — see SETUP-GUIDE.md, Step 4.")


if __name__ == "__main__":
    main()
