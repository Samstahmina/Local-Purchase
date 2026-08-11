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

# Odoo stores datetimes in UTC. The company runs on Asia/Dhaka, which is a fixed
# UTC+6 with no DST, so a plain offset is enough to render issue dates locally.
LOCAL_UTC_OFFSET = timedelta(hours=6)

ODOO_CONTEXT = {
    "lang": "en_US",
    "tz": "Asia/Dhaka",
    "allowed_company_ids": [1],
    "current_company_id": 1,
}

# Same domain as the "RM" filter that Odoo applies by default on this report's
# list view (search view filter name="rm_stock"), so the sheet matches what a
# user sees on screen.
RM_DOMAIN = [["product_id.categ_id.complete_name", "ilike", "All / RM"]]

# Keep only rows that were issued during the period. Set to False to export
# every row of the report (including untouched stock) with the issue dates
# still filled in where they exist.
ONLY_ISSUED_ROWS = True

FIELDS_SPEC = {
    "parent_category": {"fields": {"display_name": {}}},
    "product_category": {"fields": {"display_name": {}}},
    "classification_id": {"fields": {"display_name": {}}},
    "product_type": {"fields": {"display_name": {}}},
    "item_category": {"fields": {"display_name": {}}},
    "product_id": {"fields": {"display_name": {}}},
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
    "Product",
    "Category",
    "Classification",
    "Product Type",
    "Item Type",
    "Item",
    "Product ID",
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
    "Issue Date",
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


def odoo_call(cookies, model, method, args, kwargs=None, timeout=300):
    url = f"{ODOO_URL}/web/dataset/call_kw"
    call_kwargs = {"context": ODOO_CONTEXT}
    if kwargs:
        call_kwargs.update(kwargs)
    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": model,
            "method": method,
            "args": args,
            "kwargs": call_kwargs,
        },
    }
    resp = requests.post(url, data=json.dumps(payload), headers=HEADERS, cookies=cookies, timeout=timeout)
    resp.raise_for_status()
    result = resp.json()
    if "error" in result:
        message = result["error"].get("data", {}).get("message") or result["error"].get("message")
        raise Exception(f"Odoo {model}.{method} failed: {message}")
    return result.get("result")


def generate_stock_report(cookies, from_date, to_date):
    """Populate stock.opening.closing for the given window.

    The report table is empty until someone runs the wizard; it is rebuilt on
    every run rather than accumulating rows. Without this step the script reads
    whatever the last user happened to generate, or nothing at all.

    from_date behaves as the opening snapshot, so pass the last day of the
    previous month to cover a full calendar month.
    """
    wizard_id = odoo_call(
        cookies,
        "stock.forecast.report",
        "create",
        [{
            "report_type": "rmstock",
            "report_for": "rm",
            "from_date": from_date.strftime("%Y-%m-%d"),
            "to_date": to_date.strftime("%Y-%m-%d"),
        }],
    )
    odoo_call(cookies, "stock.forecast.report", "print_date_wise_stock_register", [[wizard_id]])
    return wizard_id


def odoo_web_search_read(cookies, model, domain, offset=0, limit=80):
    return odoo_call(
        cookies,
        model,
        "web_search_read",
        [],
        {
            "specification": FIELDS_SPEC,
            "offset": offset,
            "order": "",
            "limit": limit,
            "count_limit": 10001,
            "domain": domain,
        },
    )


def utc_to_local_date(value):
    if not value:
        return ""
    try:
        dt = datetime.strptime(str(value)[:19], "%Y-%m-%d %H:%M:%S")
    except ValueError:
        return str(value)[:10]
    return (dt + LOCAL_UTC_OFFSET).strftime("%Y-%m-%d")


def fetch_issue_dates(cookies, from_date, to_date):
    """Map (product_id, lot_id) -> sorted list of dates the lot was issued.

    stock.opening.closing has no issue date of its own, so the dates come from
    the underlying stock.move.line records: a raw-material issue is a done move
    out of an internal location into a production location.
    """
    domain = [
        ["date", ">=", from_date.strftime("%Y-%m-%d 00:00:00")],
        ["date", "<=", to_date.strftime("%Y-%m-%d 23:59:59")],
        ["state", "=", "done"],
        ["location_id.usage", "=", "internal"],
        ["location_dest_id.usage", "=", "production"],
    ]
    issue_dates = {}
    offset = 0
    limit = 1000
    while True:
        batch = odoo_call(
            cookies,
            "stock.move.line",
            "search_read",
            [domain, ["date", "product_id", "lot_id"]],
            {"offset": offset, "limit": limit, "order": "date"},
        ) or []
        for line in batch:
            product = line.get("product_id") or []
            lot = line.get("lot_id") or []
            if not product:
                continue
            key = (product[0], lot[0] if lot else 0)
            issue_dates.setdefault(key, set()).add(utc_to_local_date(line.get("date")))
        print(f"Fetched {len(batch)} issue move lines at offset {offset}")
        if len(batch) < limit:
            break
        offset += limit
    return {key: sorted(values) for key, values in issue_dates.items()}


def cell(record, field):
    """Odoo returns False for unset scalars; write a blank cell instead."""
    value = record.get(field)
    return "" if value is False or value is None else value


def record_key(record):
    product_id = record.get("product_id") or {}
    lot_id = record.get("lot_id") or {}
    return (product_id.get("id", 0), lot_id.get("id", 0))


def flatten_record(record, issue_dates):
    row = []
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
    row.append(product_id.get("display_name", "") if product_id else "")
    row.append(product_id.get("id", "") if product_id else "")
    product_uom = record.get("product_uom") or {}
    row.append(product_uom.get("display_name", "") if product_uom else "")
    lot_id = record.get("lot_id") or {}
    row.append(lot_id.get("display_name", "") if lot_id else "")
    location = record.get("location") or {}
    row.append(location.get("display_name", "") if location else "")
    row.append(cell(record, "rejected"))
    row.append(cell(record, "lot_price"))
    row.append(cell(record, "pur_price"))
    row.append(cell(record, "landed_cost"))
    row.append(cell(record, "opening_qty"))
    row.append(cell(record, "opening_value"))
    row.append(cell(record, "receive_date"))
    row.append(cell(record, "receive_qty"))
    row.append(cell(record, "receive_value"))
    row.append(", ".join(issue_dates.get(record_key(record), [])))
    row.append(cell(record, "issue_qty"))
    row.append(cell(record, "issue_value"))
    row.append(cell(record, "cloing_qty"))
    row.append(cell(record, "cloing_value"))
    row.append(cell(record, "shipment_mode"))
    row.append(cell(record, "po_type"))
    partner_id = record.get("partner_id") or {}
    row.append(partner_id.get("display_name", "") if partner_id else "")
    row.append(cell(record, "po_number"))
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

    # The wizard treats from_date as the opening snapshot and counts movements
    # after it, so start from the last day of the previous month.
    from_date = month_start - timedelta(days=1)
    to_date = month_end

    sheet_name = today.strftime("%B-%y")

    print("Authenticating with Odoo...")
    cookies = odoo_authenticate()
    print("Authenticated successfully.")

    print(f"Generating stock report for {from_date} to {to_date}...")
    wizard_id = generate_stock_report(cookies, from_date, to_date)
    print(f"Report generated (wizard id {wizard_id}).")

    print("Fetching issue dates from stock move lines...")
    issue_dates = fetch_issue_dates(cookies, from_date, to_date)
    print(f"Issue dates found for {len(issue_dates)} product/lot combinations.")

    all_records = []
    offset = 0
    limit = 80
    while True:
        result = odoo_web_search_read(cookies, "stock.opening.closing", RM_DOMAIN, offset=offset, limit=limit)
        records = result.get("records", [])
        all_records.extend(records)
        print(f"Fetched {len(records)} records at offset {offset}")
        if len(records) < limit:
            break
        offset += limit

    print(f"Total records fetched: {len(all_records)}")

    if not all_records:
        print("Error: the report returned no rows. Check that the wizard ran and that company 1 has RM stock.", file=sys.stderr)
        sys.exit(1)

    if ONLY_ISSUED_ROWS:
        selected = [r for r in all_records if (r.get("issue_qty") or 0) != 0]
        print(f"Rows issued between {from_date} and {to_date}: {len(selected)} of {len(all_records)}")
    else:
        selected = all_records

    rows = [flatten_record(r, issue_dates) for r in selected]

    print("Connecting to Google Sheets...")
    gc = get_gspread_client()
    ws = get_worksheet_cached(gc, sheet_name)

    update_sheet(ws, rows)
    print(f"Updated Google Sheet '{sheet_name}' with {len(rows)} records successfully.")


if __name__ == "__main__":
    main()
