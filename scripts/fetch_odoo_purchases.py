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

START_DATE_ENV = os.environ.get("START_DATE", "").strip() or "2026-07-01"

HEADERS = {"Content-Type": "application/json"}

FIELDS_SPEC = {
    "priority": {},
    "partner_ref": {},
    "has_alternatives": {},
    "name": {},
    "partner_id": {"fields": {"display_name": {}}},
    "x_studio_pi_no": {},
    "create_date": {},
    "x_studio_order_status": {},
    "company_id": {"fields": {"display_name": {}}},
    "date_planned": {},
    "user_id": {"fields": {}},
    "create_uid": {"fields": {"display_name": {}}},
    "last_approver": {"fields": {"display_name": {}}},
    "next_approver": {"fields": {"display_name": {}}},
    "date_approve": {},
    "activity_ids": {},
    "activity_exception_decoration": {},
    "activity_exception_icon": {},
    "activity_state": {},
    "activity_summary": {},
    "activity_type_icon": {},
    "activity_type_id": {"fields": {"display_name": {}}},
    "origin": {},
    "amount_untaxed": {},
    "amount_total": {},
    "x_studio_currency": {"fields": {"display_name": {}}},
    "x_studio_gate_entry": {},
    "currency_id": {"fields": {}},
    "state": {},
    "invoice_status": {},
}

FLAT_HEADERS = [
    "ID",
    "Priority",
    "Partner Ref",
    "Has Alternatives",
    "Name",
    "Partner",
    "PI No",
    "Create Date",
    "Order Status",
    "Company",
    "Date Planned",
    "User",
    "Created By",
    "Last Approver",
    "Next Approver",
    "Date Approve",
    "Activity IDs",
    "Activity Exception Decoration",
    "Activity Exception Icon",
    "Activity State",
    "Activity Summary",
    "Activity Type Icon",
    "Activity Type",
    "Origin",
    "Amount Untaxed",
    "Amount Total",
    "Currency",
    "Gate Entry",
    "State",
    "Invoice Status",
]


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


def odoo_web_search_read(cookies, model, domain, offset=0, limit=80):
    url = f"{ODOO_URL}/web/dataset/call_kw"
    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": model,
            "method": "web_search_read",
            "args": [],
            "kwargs": {
                "specification": FIELDS_SPEC,
                "offset": offset,
                "order": "date_order desc",
                "limit": limit,
                "context": {
                    "lang": "en_US",
                    "tz": "Asia/Almaty",
                    "uid": 10,
                    "allowed_company_ids": [1],
                    "bin_size": True,
                    "quotation_only": True,
                    "current_company_id": 1,
                },
                "count_limit": 10001,
                "domain": domain,
            },
        },
    }
    resp = requests.post(url, data=json.dumps(payload), headers=HEADERS, cookies=cookies, timeout=60)
    resp.raise_for_status()
    result = resp.json()
    if "result" in result:
        return result["result"]
    raise Exception(f"Odoo web_search_read failed: {result}")


def flatten_record(record):
    row = []
    row.append(record.get("id", ""))
    row.append(record.get("priority", ""))
    row.append(record.get("partner_ref", ""))
    row.append(record.get("has_alternatives", ""))
    row.append(record.get("name", ""))
    partner = record.get("partner_id") or {}
    row.append(partner.get("display_name", "") if partner else "")
    row.append(record.get("x_studio_pi_no", ""))
    row.append(record.get("create_date", ""))
    row.append(record.get("x_studio_order_status", ""))
    company = record.get("company_id") or {}
    row.append(company.get("display_name", "") if company else "")
    row.append(record.get("date_planned", ""))
    user = record.get("user_id") or {}
    row.append(user.get("id", "") if user else "")
    create_uid = record.get("create_uid") or {}
    row.append(create_uid.get("display_name", "") if create_uid else "")
    last_approver = record.get("last_approver") or {}
    row.append(last_approver.get("display_name", "") if last_approver else "")
    next_approver = record.get("next_approver") or {}
    row.append(next_approver.get("display_name", "") if next_approver else "")
    row.append(record.get("date_approve", ""))
    row.append(", ".join(map(str, record.get("activity_ids", []))) if record.get("activity_ids") else "")
    row.append(record.get("activity_exception_decoration", ""))
    row.append(record.get("activity_exception_icon", ""))
    row.append(record.get("activity_state", ""))
    row.append(record.get("activity_summary", ""))
    row.append(record.get("activity_type_icon", ""))
    activity_type = record.get("activity_type_id") or {}
    row.append(activity_type.get("display_name", "") if activity_type else "")
    row.append(record.get("origin", ""))
    row.append(record.get("amount_untaxed", ""))
    row.append(record.get("amount_total", ""))
    currency = record.get("x_studio_currency") or {}
    row.append(currency.get("display_name", "") if currency else "")
    row.append(record.get("x_studio_gate_entry", ""))
    currency_id = record.get("currency_id") or {}
    row.append(currency_id.get("id", "") if currency_id else "")
    row.append(record.get("state", ""))
    row.append(record.get("invoice_status", ""))
    return row


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


def update_sheet(rows, start_date, today):
    creds_dict = json.loads(GOOGLE_CREDENTIALS_JSON)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID)
    ws = get_worksheet(sh, SHEET_NAME)

    ws.clear()
    ws.update([FLAT_HEADERS], "A1")
    if rows:
        ws.update(rows, f"A2:{chr(64 + len(FLAT_HEADERS))}{len(rows) + 1}")


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
    print(f"Fetching purchase orders from {start_date} to {today}...")

    all_records = []
    offset = 0
    limit = 80
    while True:
        result = odoo_web_search_read(cookies, "purchase.order", domain, offset=offset, limit=limit)
        records = result.get("records", [])
        all_records.extend(records)
        print(f"Fetched {len(records)} records at offset {offset}")
        if len(records) < limit:
            break
        offset += limit

    print(f"Total records fetched: {len(all_records)}")

    rows = [flatten_record(r) for r in all_records]
    update_sheet(rows, start_date, today)
    print(f"Updated Google Sheet '{SHEET_NAME}' with {len(rows)} records successfully.")


if __name__ == "__main__":
    main()
