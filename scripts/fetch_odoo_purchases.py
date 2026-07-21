import os
import sys
import json
import requests
from datetime import datetime, date
import gspread
from google.oauth2.service_account import Credentials

ODOO_URL = "https://taps.odoo.com".rstrip("/")
ODOO_DB = os.environ.get("ODOO_DB", "")
ODOO_USERNAME = os.environ.get("ODOO_USERNAME", "")
ODOO_PASSWORD = os.environ.get("ODOO_PASSWORD", "")

SPREADSHEET_ID = "1XDPawgWLd34FbRiSOEsoR8qgpWlUVIu5kM1YLVEXB2o"
SHEET_NAME = "testing"
GOOGLE_CREDENTIALS_JSON = os.environ.get("GOOGLE_CREDENTIALS_JSON", "")

START_DATE_ENV = os.environ.get("START_DATE", "").strip()

HEADERS = {"Content-Type": "application/json"}


def odoo_authenticate():
    url = f"{ODOO_URL}/web/session/authenticate"
    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "db": ODOO_DB,
            "login": ODOO_USERNAME,
            "password": ODOO_PASSWORD,
        },
    }
    resp = requests.post(url, data=json.dumps(payload), headers=HEADERS, timeout=30)
    resp.raise_for_status()
    result = resp.json()
    if result.get("result") and result["result"].get("uid", 0) > 0:
        return resp.cookies
    raise Exception(f"Odoo authentication failed: {result}")


def odoo_search_count(cookies, model, domain):
    url = f"{ODOO_URL}/web/dataset/call_kw"
    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": model,
            "method": "search_count",
            "args": [domain],
            "kwargs": {},
        },
    }
    resp = requests.post(url, data=json.dumps(payload), headers=HEADERS, cookies=cookies, timeout=30)
    resp.raise_for_status()
    result = resp.json()
    if "result" in result:
        return result["result"]
    raise Exception(f"Odoo search_count failed: {result}")


def get_worksheet(sh, name):
    try:
        return sh.worksheet(name)
    except gspread.exceptions.WorksheetNotFound:
        titles = [ws.title for ws in sh.worksheets()]
        for title in titles:
            if title.strip().lower() == name.strip().lower():
                return sh.worksheet(title)
        raise Exception(
            f"Worksheet '{name}' not found. Available worksheets: {titles}"
        )


def get_start_date():
    if START_DATE_ENV:
        return datetime.strptime(START_DATE_ENV, "%Y-%m-%d").date()

    creds_dict = json.loads(GOOGLE_CREDENTIALS_JSON)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID)
    ws = get_worksheet(sh, SHEET_NAME)

    start_date_str = ws.acell("A1").value
    if not start_date_str:
        raise ValueError(f"Cell A1 in sheet '{SHEET_NAME}' is empty. Please set the start date (YYYY-MM-DD).")
    return datetime.strptime(start_date_str.strip(), "%Y-%m-%d").date()


def update_sheet(count, start_date, today):
    creds_dict = json.loads(GOOGLE_CREDENTIALS_JSON)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID)
    ws = get_worksheet(sh, SHEET_NAME)

    ws.update(
        [[start_date.strftime("%Y-%m-%d"), today.strftime("%Y-%m-%d"), count, datetime.now().strftime("%Y-%m-%d %H:%M:%S")]],
        "A2:D2",
    )


def main():
    if not GOOGLE_CREDENTIALS_JSON:
        print("Error: GOOGLE_CREDENTIALS_JSON environment variable is not set.", file=sys.stderr)
        sys.exit(1)

    start_date = get_start_date()
    today = date.today()
    if start_date > today:
        print(f"Start date {start_date} is in the future. Nothing to fetch.", file=sys.stderr)
        sys.exit(0)

    print(f"Authenticating with Odoo...")
    cookies = odoo_authenticate()
    print(f"Authenticated successfully (UID: {cookies.get('session_id', 'N/A')}).")

    domain = [["date_order", ">=", start_date.strftime("%Y-%m-%d")]]
    count = odoo_search_count(cookies, "purchase.order", domain)
    print(f"Local purchased count from {start_date} to {today}: {count}")

    update_sheet(count, start_date, today)
    print(f"Updated Google Sheet '{SHEET_NAME}' successfully.")


if __name__ == "__main__":
    main()
