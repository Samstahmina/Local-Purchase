import os
import sys
import json
import time
import requests
from datetime import datetime, date, timedelta
import gspread
from google.oauth2.service_account import Credentials

ODOO_URL = "https://taps.odoo.com".rstrip("/")
ODOO_DB = os.environ.get("ODOO_DB", "")
ODOO_USERNAME = os.environ.get("ODOO_USERNAME", "")
ODOO_PASSWORD = os.environ.get("ODOO_PASSWORD", "")

SPREADSHEET_ID = "1409cItrf5EG1DS94WMfWqpFXC6oMlSYYRCs3IJf3Rk8"
GOOGLE_CREDENTIALS_JSON = os.environ.get("GOOGLE_CREDENTIALS_JSON", "")

HEADERS = {"Content-Type": "application/json"}

FIELDS_SPEC = {
    "company_id": {"fields": {"display_name": {}}},
    "parent_category": {"fields": {"display_name": {}}},
    "product_category": {"fields": {"display_name": {}}},
    "classification_id": {"fields": {"display_name": {}}},
    "product_type": {"fields": {"display_name": {}}},
    "item_category": {"fields": {"display_name": {}}},
    "product_id": {"fields": {"display_name": {}}},
    "pr_code": {},
    "product_uom": {"fields": {"display_name": {}}},
    "lot_id": {"fields": {"display_name": {}}},
    "location": {"fields": {"display_name": {}}},
    "rejected": {},
    "lot_price": {},
    "pur_price": {},
    "landed_cost": {},
    "opening_qty": {},
    "opening_value": {},
    "receive_date": {},
    "receive_qty": {},
    "receive_value": {},
    "issue_qty": {},
    "issue_value": {},
    "cloing_qty": {},
    "cloing_value": {},
    "shipment_mode": {},
    "po_type": {},
    "partner_id": {"fields": {"display_name": {}}},
    "po_number": {},
}

FLAT_HEADERS = [
    "Company",
    "Product",
    "Category",
    "Classification",
    "Product Type",
    "Item Type",
    "Item",
    "Item Code",
    "Unit",
    "Invoice",
    "Location",
    "Rejected",
    "Price",
    "Pur Price",
    "Landed Cost",
    "Opening Quantity",
    "Opening Value",
    "Receive Date",
    "Receive Quantity",
    "Receive Value",
    "Issue Quantity",
    "Issue Value",
    "Closing Quantity",
    "Closing Value",
    "Shipment Mode",
    "PO Type",
    "Vendor",
    "PO",
]


def get_worksheet(sh, name):
    try:
        return sh.worksheet(name)
    except gspread.exceptions.WorksheetNotFound:
        try:
            return sh.add_worksheet(title=name, rows="1000", cols="30")
        except gspread.exceptions.APIError:
            titles = [ws.title for ws in sh.worksheets()]
            for title in titles:
                if title.strip().lower() == name.strip().lower():
                    return sh.worksheet(title)
            return sh.add_worksheet(title=name, rows="1000", cols="30")


def retry_gspread(func, *args, max_retries=5, backoff_factor=2, **kwargs):
    for attempt in range(1, max_retries + 1):
        try:
            return func(*args, **kwargs)
        except gspread.exceptions.APIError as e:
            status = e.response.status_code if hasattr(e, "response") and e.response else None
            if status == 429 or (status and status >= 500):
                wait = backoff_factor ** attempt
                print(f"Retrying gspread call after {status} error (attempt {attempt}/{max_retries}, wait {wait}s)...")
                time.sleep(wait)
            else:
                raise
        except requests.exceptions.ConnectionError:
            wait = backoff_factor ** attempt
            print(f"Retrying gspread call after connection error (attempt {attempt}/{max_retries}, wait {wait}s)...")
            time.sleep(wait)
    return func(*args, **kwargs)


def get_gspread_client():
    creds_dict = json.loads(GOOGLE_CREDENTIALS_JSON)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)


def get_worksheet_cached(gc, name):
    sh = retry_gspread(gc.open_by_key, SPREADSHEET_ID)
    return retry_gspread(get_worksheet, sh, name)


def col_to_letter(col):
    result = ""
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        result = chr(65 + remainder) + result
    return result


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
                "order": "",
                "limit": limit,
                "context": {
                    "lang": "en_US",
                    "tz": "Asia/Almaty",
                    "uid": 10,
                    "allowed_company_ids": [1],
                    "bin_size": True,
                    "active_model": "stock.forecast.report",
                    "active_id": 92437,
                    "active_ids": [92437],
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


def parse_product_string(raw):
    if not raw:
        return "", ""
    raw = str(raw).strip()
    start = raw.find("[")
    end = raw.find("]")
    if start != -1 and end != -1 and end > start:
        code = raw[start + 1:end]
        name = raw[end + 1:].strip()
        return name, code
    return raw, ""


def flatten_record(record):
    row = []
    company_id = record.get("company_id") or {}
    row.append(company_id.get("display_name", "") if company_id else "")
    parent_category = record.get("parent_category") or {}
    row.append(parent_category.get("display_name", "") if parent_category else "")
    product_category = record.get("product_category") or {}
    row.append(product_category.get("display_name", "") if product_category else "")
    classification_id = record.get("classification_id") or {}
    row.append(classification_id.get("display_name", "") if classification_id else "")
    product_type = record.get("product_type") or {}
    row.append(product_type.get("display_name", "") if product_type else "")
    item_category = record.get("item_category") or {}
    row.append(item_category.get("display_name", "") if item_category else "")
    product_id = record.get("product_id") or {}
    raw_name = product_id.get("display_name", "") if product_id else ""
    pname, pcode = parse_product_string(raw_name)
    row.append(pname)
    row.append(pcode)
    product_uom = record.get("product_uom") or {}
    row.append(product_uom.get("display_name", "") if product_uom else "")
    lot_id = record.get("lot_id") or {}
    row.append(lot_id.get("display_name", "") if lot_id else "")
    location = record.get("location") or {}
    row.append(location.get("display_name", "") if location else "")
    row.append(record.get("rejected", ""))
    row.append(record.get("lot_price", ""))
    row.append(record.get("pur_price", ""))
    row.append(record.get("landed_cost", ""))
    row.append(record.get("opening_qty", ""))
    row.append(record.get("opening_value", ""))
    row.append(record.get("receive_date", ""))
    row.append(record.get("receive_qty", ""))
    row.append(record.get("receive_value", ""))
    row.append(record.get("issue_qty", ""))
    row.append(record.get("issue_value", ""))
    row.append(record.get("cloing_qty", ""))
    row.append(record.get("cloing_value", ""))
    row.append(record.get("shipment_mode", ""))
    row.append(record.get("po_type", ""))
    partner_id = record.get("partner_id") or {}
    row.append(partner_id.get("display_name", "") if partner_id else "")
    row.append(record.get("po_number", ""))
    return row


def update_sheet(ws, rows):
    retry_gspread(ws.clear)
    retry_gspread(ws.update, [FLAT_HEADERS], "A1")
    if rows:
        max_cols = max(len(FLAT_HEADERS), max((len(r) for r in rows), default=0))
        end_col = col_to_letter(max_cols)
        chunk_size = 50
        for i in range(0, len(rows), chunk_size):
            chunk = rows[i:i + chunk_size]
            start_row = i + 2
            end_row = i + len(chunk) + 1
            retry_gspread(ws.update, chunk, f"A{start_row}:{end_col}{end_row}")


def main():
    if not GOOGLE_CREDENTIALS_JSON:
        print("Error: GOOGLE_CREDENTIALS_JSON environment variable is not set.", file=sys.stderr)
        sys.exit(1)

    today = date.today()
    month_start = today.replace(day=1)
    if today.month == 12:
        month_end = today.replace(year=today.year + 1, month=1, day=1) - timedelta(days=1)
    else:
        month_end = today.replace(month=today.month + 1, day=1) - timedelta(days=1)

    sheet_name = today.strftime("%B-%y")

    print(f"Authenticating with Odoo...")
    cookies = odoo_authenticate()
    print(f"Authenticated successfully (UID: {cookies.get('session_id', 'N/A')}).")

    domain = [
        ["product_id.categ_id.complete_name", "ilike", "All / RM"],
        ["company_id", "=", 1],
        ["receive_date", ">=", month_start.strftime("%Y-%m-%d")],
        ["receive_date", "<=", month_end.strftime("%Y-%m-%d")],
    ]
    print(f"Fetching stock report from {month_start} to {month_end} for company 1...")

    all_records = []
    offset = 0
    limit = 80
    while True:
        result = odoo_web_search_read(cookies, "stock.opening.closing", domain, offset=offset, limit=limit)
        records = result.get("records", [])
        all_records.extend(records)
        print(f"Fetched {len(records)} records at offset {offset}")
        if len(records) < limit:
            break
        offset += limit

    print(f"Total records fetched for company 1: {len(all_records)}")

    rows = [flatten_record(r) for r in all_records]

    print(f"Connecting to Google Sheets...")
    gc = get_gspread_client()
    ws = get_worksheet_cached(gc, sheet_name)

    update_sheet(ws, rows)
    print(f"Updated Google Sheet '{sheet_name}' with {len(rows)} records successfully.")


if __name__ == "__main__":
    main()
