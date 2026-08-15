"""Reproduce the "Procurement 1" purchase-order export into Google Sheets.

The columns, their order and their labels come from the saved Odoo export
template (ir.exports id 572, "Procurement 1") captured in
`Purchase Module History.har`; the record filter comes from the list view the
same capture was taken from:

    date_approve >= <start>  AND  date_approve <= now  AND  state = 'purchase'

Odoo explodes `order_line/*` columns into one row per order line, so an order
with three lines produces three rows with the order-level columns repeated.
Orders with no lines still produce one row with the line columns blank.
"""

import os
import sys
import json
import time
from datetime import datetime, timedelta

import requests
import gspread
from google.oauth2.service_account import Credentials

ODOO_URL = "https://taps.odoo.com".rstrip("/")
ODOO_DB = os.environ.get("ODOO_DB", "")
ODOO_USERNAME = os.environ.get("ODOO_USERNAME", "")
ODOO_PASSWORD = os.environ.get("ODOO_PASSWORD", "")

SPREADSHEET_ID = "1u6JyG3xvYXO4ID83r7f_G8GQIz7wFt6gpLTYZVoZFnE"
SHEET_NAME = "Purchase Module History"
GOOGLE_CREDENTIALS_JSON = os.environ.get("GOOGLE_CREDENTIALS_JSON", "")

# The list view was captured with tz=Asia/Almaty, which is UTC+6 like Dhaka.
# Odoo stores and filters datetimes in UTC but displays them in the user's tz,
# so every datetime crossing the wire gets shifted by this offset.
TZ_OFFSET = timedelta(hours=6)

# Local calendar date the history starts at. The capture used 2025-01-01 local,
# which Odoo sent as "2024-12-31 18:00:24" UTC.
START_DATE = os.environ.get("PURCHASE_START_DATE", "").strip() or "2025-01-01"

HEADERS = {"Content-Type": "application/json"}

ODOO_CONTEXT = {
    "lang": "en_US",
    "tz": "Asia/Almaty",
    "allowed_company_ids": [1],
    "current_company_id": 1,
}

# ---------------------------------------------------------------------------
# Column layout, verbatim from /web/export/namelist for export_id 572.
# (odoo field path, sheet header, source) where source is "order", "line",
# "order_product" or "line_product".
# ---------------------------------------------------------------------------
COLUMNS = [
    ("priority", "Priority", "order"),
    ("name", "Order Reference", "order"),
    ("partner_id", "Vendor", "order"),
    ("x_studio_pi_no", "PI No.", "order"),
    ("create_date", "Created on", "order"),
    ("x_studio_order_status", "Order Status", "order"),
    ("company_id", "Company", "order"),
    ("create_uid", "Created by", "order"),
    ("last_approver", "Last Approver", "order"),
    ("date_approve", "Confirmation Date", "order"),
    ("origin", "Source Document", "order"),
    ("amount_total", "Total", "order"),
    ("x_studio_currency", "Currency.", "order"),
    ("x_studio_gate_entry", "Gate Entry", "order"),
    ("state", "Status", "order"),
    ("payment_term_id", "Payment Terms", "order"),
    ("product_uom_qty", "Order Lines/Total Quantity", "line"),
    ("product_id", "Order Lines/Product", "line"),
    ("po_type", "PO Type", "order"),
    ("qty_received", "Order Lines/Received Qty", "line"),
    ("price_unit", "Order Lines/Unit Price", "line"),
    ("product_uom", "Order Lines/Unit of Measure", "line"),
    ("last_purchase_price", "Order Lines/Last Purchase", "line"),
    ("categ_type", "Product/Category Type/Type of Categories", "order_product"),
    ("incoterm_id", "Incoterm", "order"),
    ("itemtype", "Item Types", "order"),
    ("categ_id", "Order Lines/Product/Product Category", "line_product"),
    ("shipment_mode", "Shipment Mode", "order"),
]

FLAT_HEADERS = [label for _, label, _ in COLUMNS]

# Some field names moved between Odoo versions. The first candidate that
# fields_get reports is the one actually used.
FIELD_ALIASES = {
    "product_uom_qty": ["product_uom_qty", "product_qty"],
    "product_uom": ["product_uom", "product_uom_id"],
}

DATETIME_FIELDS = {"create_date", "date_approve"}


# ---------------------------------------------------------------------------
# Odoo
# ---------------------------------------------------------------------------

def odoo_authenticate():
    url = f"{ODOO_URL}/web/session/authenticate"
    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {"db": ODOO_DB, "login": ODOO_USERNAME, "password": ODOO_PASSWORD},
    }
    resp = requests.post(url, data=json.dumps(payload), headers=HEADERS, timeout=60)
    resp.raise_for_status()
    result = resp.json()
    if result.get("result") and result["result"].get("uid", 0) > 0:
        return resp.cookies, result["result"]["uid"]
    raise Exception(f"Odoo authentication failed: {result}")


def call_kw(cookies, model, method, args, kwargs=None, timeout=180):
    url = f"{ODOO_URL}/web/dataset/call_kw"
    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": model,
            "method": method,
            "args": args,
            "kwargs": kwargs or {},
        },
    }
    last_error = None
    for attempt in range(1, 4):
        try:
            resp = requests.post(
                url, data=json.dumps(payload), headers=HEADERS, cookies=cookies, timeout=timeout
            )
            resp.raise_for_status()
            result = resp.json()
        except (requests.exceptions.RequestException, ValueError) as e:
            last_error = e
            wait = 2 ** attempt
            print(f"  {model}.{method} transport error ({e}); retrying in {wait}s")
            time.sleep(wait)
            continue
        if "error" in result:
            raise Exception(f"Odoo {model}.{method} failed: {json.dumps(result['error'])[:800]}")
        return result.get("result")
    raise Exception(f"Odoo {model}.{method} failed after retries: {last_error}")


def fields_get(cookies, model):
    """Field metadata for `model`, keyed by field name."""
    return call_kw(
        cookies,
        model,
        "fields_get",
        [[], ["string", "type", "selection", "relation"]],
        {"context": dict(ODOO_CONTEXT)},
    ) or {}


def resolve(meta, name):
    """Return the field name `model` actually has for `name`, or None."""
    for candidate in FIELD_ALIASES.get(name, [name]):
        if candidate in meta:
            return candidate
    return None


def search_read(cookies, model, domain, fields, order="id", batch=500):
    """Read every record matching `domain`, paging until the server runs dry."""
    records = []
    offset = 0
    while True:
        page = call_kw(
            cookies,
            model,
            "search_read",
            [domain, fields],
            {"offset": offset, "limit": batch, "order": order, "context": dict(ODOO_CONTEXT)},
        ) or []
        records.extend(page)
        print(f"  {model}: {len(records)} records")
        if len(page) < batch:
            return records
        offset += batch


def read_ids(cookies, model, ids, fields, batch=500):
    """Read specific ids (used for the product lookups)."""
    records = []
    ids = sorted(set(ids))
    for i in range(0, len(ids), batch):
        chunk = ids[i:i + batch]
        records.extend(
            call_kw(cookies, model, "read", [chunk, fields], {"context": dict(ODOO_CONTEXT)}) or []
        )
    return records


# ---------------------------------------------------------------------------
# Value formatting — match what the Odoo export writes, not the raw JSON
# ---------------------------------------------------------------------------

def to_local(value):
    """Odoo hands back naive UTC datetimes; the export shows the user's tz."""
    if not value:
        return ""
    try:
        return (datetime.strptime(value, "%Y-%m-%d %H:%M:%S") + TZ_OFFSET).strftime(
            "%Y-%m-%d %H:%M:%S"
        )
    except ValueError:
        return value


def format_value(value, field_name, meta):
    if value is None or value is False:
        return ""
    if field_name in DATETIME_FIELDS:
        return to_local(value)
    # many2one arrives as [id, display_name]; the export writes the name.
    if isinstance(value, list):
        if len(value) == 2 and isinstance(value[0], int) and isinstance(value[1], str):
            return value[1]
        return ", ".join(str(v) for v in value)
    info = meta.get(field_name) or {}
    if info.get("type") == "selection" and info.get("selection"):
        for raw, label in info["selection"]:
            if raw == value:
                return label
    return value


# ---------------------------------------------------------------------------
# Google Sheets
# ---------------------------------------------------------------------------

def retry_gspread(func, *args, max_retries=5, backoff_factor=2, **kwargs):
    for attempt in range(1, max_retries + 1):
        try:
            return func(*args, **kwargs)
        except gspread.exceptions.APIError as e:
            status = e.response.status_code if getattr(e, "response", None) else None
            if status == 429 or (status and status >= 500):
                wait = backoff_factor ** attempt
                print(f"Retrying after {status} (attempt {attempt}/{max_retries}, wait {wait}s)...")
                time.sleep(wait)
            else:
                raise
        except requests.exceptions.ConnectionError:
            wait = backoff_factor ** attempt
            print(f"Retrying after connection error (attempt {attempt}/{max_retries}, wait {wait}s)...")
            time.sleep(wait)
    return func(*args, **kwargs)


def get_worksheet(sh, name):
    try:
        return sh.worksheet(name)
    except gspread.exceptions.WorksheetNotFound:
        for ws in sh.worksheets():
            if ws.title.strip().lower() == name.strip().lower():
                return ws
        return sh.add_worksheet(title=name, rows=1000, cols=len(FLAT_HEADERS))


def col_to_letter(col):
    result = ""
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        result = chr(65 + remainder) + result
    return result


def update_sheet(ws, rows):
    end_col = col_to_letter(len(FLAT_HEADERS))
    needed_rows = len(rows) + 1
    if ws.row_count < needed_rows or ws.col_count < len(FLAT_HEADERS):
        retry_gspread(
            ws.resize,
            rows=max(needed_rows, ws.row_count),
            cols=max(len(FLAT_HEADERS), ws.col_count),
        )
    retry_gspread(ws.clear)
    retry_gspread(ws.update, [FLAT_HEADERS], "A1")

    chunk_size = 2000
    for i in range(0, len(rows), chunk_size):
        chunk = rows[i:i + chunk_size]
        start_row = i + 2
        end_row = start_row + len(chunk) - 1
        retry_gspread(ws.update, chunk, f"A{start_row}:{end_col}{end_row}")
        print(f"  wrote rows {start_row}-{end_row}")


# ---------------------------------------------------------------------------

def main():
    missing = [
        name
        for name, value in [
            ("ODOO_DB", ODOO_DB),
            ("ODOO_USERNAME", ODOO_USERNAME),
            ("ODOO_PASSWORD", ODOO_PASSWORD),
            ("GOOGLE_CREDENTIALS_JSON", GOOGLE_CREDENTIALS_JSON),
        ]
        if not value
    ]
    if missing:
        print(f"Error: missing environment variables: {', '.join(missing)}", file=sys.stderr)
        sys.exit(1)

    print("Authenticating with Odoo...")
    cookies, uid = odoo_authenticate()
    ODOO_CONTEXT["uid"] = uid
    print(f"Authenticated (uid {uid}).")

    order_meta = fields_get(cookies, "purchase.order")
    line_meta = fields_get(cookies, "purchase.order.line")
    product_meta = fields_get(cookies, "product.product")

    # Split the export template's columns by where the value comes from, and
    # drop anything this database does not actually have rather than failing
    # the whole read.
    order_fields, line_fields = [], []
    resolved = {}
    for name, label, source in COLUMNS:
        meta = {"order": order_meta, "line": line_meta}.get(source)
        if meta is None:
            resolved[(source, name)] = name  # product lookups, checked below
            continue
        actual = resolve(meta, name)
        resolved[(source, name)] = actual
        if actual is None:
            print(f"  note: '{name}' ({label}) is not on this database; column left blank")
        elif source == "order":
            order_fields.append(actual)
        else:
            line_fields.append(actual)

    order_product_field = resolve(order_meta, "product_id")
    if order_product_field:
        order_fields.append(order_product_field)
    line_product_field = resolve(line_meta, "product_id")

    product_fields = [f for f in ("categ_id", "categ_type") if f in product_meta]

    # The list view filter, verbatim from the capture.
    start_utc = (
        datetime.strptime(START_DATE, "%Y-%m-%d") - TZ_OFFSET
    ).strftime("%Y-%m-%d %H:%M:%S")
    end_utc = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
    domain = [
        "&", "&",
        ["date_approve", ">=", start_utc],
        ["date_approve", "<=", end_utc],
        ["state", "=", "purchase"],
    ]
    print(f"Fetching confirmed purchase orders, {start_utc} to {end_utc} UTC...")

    orders = search_read(cookies, "purchase.order", domain, sorted(set(order_fields)))
    print(f"Total orders: {len(orders)}")

    lines = []
    if orders:
        order_ids = [o["id"] for o in orders]
        wanted = sorted(set(line_fields + ["order_id"] + ([line_product_field] if line_product_field else [])))
        for i in range(0, len(order_ids), 500):
            chunk = order_ids[i:i + 500]
            lines.extend(
                search_read(cookies, "purchase.order.line", [["order_id", "in", chunk]], wanted)
            )
    print(f"Total order lines: {len(lines)}")

    lines_by_order = {}
    for line in lines:
        order_id = line.get("order_id")
        oid = order_id[0] if isinstance(order_id, list) and order_id else order_id
        lines_by_order.setdefault(oid, []).append(line)

    # One extra read resolves both product-derived columns.
    products = {}
    if product_fields:
        product_ids = []
        for order in orders:
            value = order.get(order_product_field) if order_product_field else None
            if isinstance(value, list) and value:
                product_ids.append(value[0])
        for line in lines:
            value = line.get(line_product_field) if line_product_field else None
            if isinstance(value, list) and value:
                product_ids.append(value[0])
        if product_ids:
            print(f"Resolving categories for {len(set(product_ids))} products...")
            for record in read_ids(cookies, "product.product", product_ids, product_fields):
                products[record["id"]] = record

    def product_of(record, field):
        value = record.get(field) if field else None
        if isinstance(value, list) and value:
            return products.get(value[0], {})
        return {}

    rows = []
    for order in orders:
        order_product = product_of(order, order_product_field)
        for line in lines_by_order.get(order["id"], [None]):
            line_product = product_of(line, line_product_field) if line else {}
            row = []
            for name, _label, source in COLUMNS:
                actual = resolved[(source, name)]
                if source == "order":
                    row.append(format_value(order.get(actual), actual, order_meta) if actual else "")
                elif source == "line":
                    row.append(
                        format_value(line.get(actual), actual, line_meta)
                        if (line and actual)
                        else ""
                    )
                elif source == "order_product":
                    row.append(format_value(order_product.get(name), name, product_meta))
                else:
                    row.append(format_value(line_product.get(name), name, product_meta))
            rows.append(row)

    print(f"Built {len(rows)} export rows.")

    print("Connecting to Google Sheets...")
    creds = Credentials.from_service_account_info(
        json.loads(GOOGLE_CREDENTIALS_JSON),
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ],
    )
    gc = gspread.authorize(creds)
    sh = retry_gspread(gc.open_by_key, SPREADSHEET_ID)
    ws = retry_gspread(get_worksheet, sh, SHEET_NAME)

    update_sheet(ws, rows)
    print(f"Updated '{SHEET_NAME}' with {len(rows)} rows.")


if __name__ == "__main__":
    main()
