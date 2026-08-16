"""Reproduce the "Procurement 1" purchase-order export into Google Sheets.

The columns, their order and their labels come from the saved Odoo export
template (ir.exports id 572, "Procurement 1") captured in
`Purchase Module History.har`. The record filter is by creation date, not
status -- orders in any state (draft, sent, purchase, done, ...) are
included except cancelled ones:

    create_date >= 2024-12-31 18:00:24 UTC  (2025-01-01 00:00 Dhaka)
    AND create_date <= now
    AND company_id in (1, 2, 3)
    AND state != 'cancel'

Two departures from the template. Odoo's "Order Lines/Last Purchase" packs an
amount, a currency and the originating PO reference into one string, so it is
split into an amount column and a "Currency-Last Purchase" column and the PO
reference is dropped. And seven of the template's columns are not wanted here:
Priority, PI No., Order Status, Source Document, Gate Entry, Order
Lines/Product/Product Category and Total.

On top of the template there is a Responsible column, holding the "Add a
section" heading(s) -- where the requisitioning person and APR number are
written -- that each product line sits under. Several headings can appear
back-to-back before the next product line; only ones mentioning "APR" or
"By" somewhere in the text (not "By" followed straight by a number) are
kept, everything else is ignored, and consecutive kept headings are joined
together. A kept heading (or group of them) covers every line beneath it
until the next kept heading. Lines above the first kept heading stay blank.

Odoo explodes `order_line/*` columns into one row per order line, so an order
with three lines produces three rows with the order-level columns repeated.
Orders with no lines still produce one row with the line columns blank.
"""

import os
import re
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

# Start of the window: 2025-01-01 00:00 Dhaka expressed in UTC. Fixed, not
# configurable — the sheet is meant to hold the same history the manual
# download produces. Filters on create_date, not date_approve.
START_UTC = "2024-12-31 18:00:24"

# Companies whose orders belong in this sheet.
COMPANY_IDS = [1, 2, 3]

HEADERS = {"Content-Type": "application/json"}

ODOO_CONTEXT = {
    "lang": "en_US",
    "tz": "Asia/Almaty",
    "allowed_company_ids": COMPANY_IDS,
    "current_company_id": COMPANY_IDS[0],
}

# ---------------------------------------------------------------------------
# Column layout: the order and labels of /web/export/namelist for export_id
# 572, minus the six unwanted columns and plus the split-out last-purchase
# currency and the Responsible column. (odoo field name, sheet header, source),
# where source says which record the value is read off and how:
#   order / line             - straight off the order or its line
#   order_product            - off the product the order points at
#   line_product             - off the product the line points at
#   last_purchase_*          - the amount or currency parsed out of the line's
#                              packed last_purchase_price string
#   sections                 - the "Add a section" heading the line sits under
# ---------------------------------------------------------------------------
COLUMNS = [
    ("name", "Order Reference", "order"),
    ("partner_id", "Vendor", "order"),
    ("create_date", "Created on", "order"),
    ("company_id", "Company", "order"),
    ("create_uid", "Created by", "order"),
    ("last_approver", "Last Approver", "order"),
    ("date_approve", "Confirmation Date", "order"),
    ("x_studio_currency", "Currency.", "order"),
    ("state", "Status", "order"),
    ("payment_term_id", "Payment Terms", "order"),
    ("product_uom_qty", "Order Lines/Total Quantity", "line"),
    ("product_id", "Order Lines/Product", "line"),
    ("po_type", "PO Type", "order"),
    ("qty_received", "Order Lines/Received Qty", "line"),
    ("price_unit", "Order Lines/Unit Price", "line"),
    ("product_uom", "Order Lines/Unit of Measure", "line"),
    ("last_purchase_price", "Order Lines/Last Purchase Price", "last_purchase_value"),
    ("last_purchase_price", "Currency-Last Purchase", "last_purchase_currency"),
    ("categ_type", "Product/Category Type/Type of Categories", "order_product"),
    ("incoterm_id", "Incoterm", "order"),
    ("itemtype", "Item Types", "order"),
    ("shipment_mode", "Shipment Mode", "order"),
    ("section_names", "Responsible", "sections"),
]

FLAT_HEADERS = [label for _, label, _ in COLUMNS]

# Some field names moved between Odoo versions. The first candidate that
# fields_get reports is the one actually used.
FIELD_ALIASES = {
    "product_uom_qty": ["product_uom_qty", "product_qty"],
    "product_uom": ["product_uom", "product_uom_id"],
}

DATETIME_FIELDS = {"create_date", "date_approve"}

# "Add a section" and "Add a note" rows are stored in order_line alongside the
# real lines and are told apart by display_type; a normal line leaves it unset.
# The heading text an "Add a section" row carries is just its `name`. Several
# headings can appear back-to-back before the next product line, and not
# every one of them is meaningful -- only the ones carrying the APR number or
# the requisitioning person's name are, so only those are kept; anything else
# (e.g. a plain grouping label) is ignored rather than clobbering the ones
# that matter. Consecutive kept headings before the next product line are
# joined together, since e.g. an APR heading and a By heading describe the
# same responsible party.
SECTION_DISPLAY_TYPE = "line_section"
LINE_BOOKKEEPING_FIELDS = ["display_type", "name", "sequence"]

# Kept: the heading mentions "APR" anywhere, or mentions "By" anywhere as its
# own word (e.g. "By Sohel Rana" or "Mr Kashem-By WhatsApp") and isn't "By"
# followed straight by a number (that's some other code, not a name).
SECTION_APR = re.compile(r"apr", re.IGNORECASE)
SECTION_BY = re.compile(r"\bby\b", re.IGNORECASE)
SECTION_BY_NUMERIC = re.compile(r"\bby\s*\d", re.IGNORECASE)


def section_is_relevant(text):
    if not text:
        return False
    if SECTION_APR.search(text):
        return True
    return bool(SECTION_BY.search(text) and not SECTION_BY_NUMERIC.search(text))

# Which model each column's source reads from, for the field-existence check.
SOURCE_MODEL = {
    "order": "order",
    "line": "line",
    "last_purchase_value": "line",
    "last_purchase_currency": "line",
}

# "0.00 BDT (P09183)" -> amount "0.00", currency "BDT". The currency pattern
# needs 2-5 *consecutive* capitals, so a reference like P09183 or PO9183 can
# never match it: the letters there are butted against digits.
LAST_PURCHASE_AMOUNT = re.compile(r"-?\d[\d,]*(?:\.\d+)?")
LAST_PURCHASE_CURRENCY = re.compile(r"\b[A-Z]{2,5}\b")


def parse_last_purchase(raw):
    """Split the packed last-purchase string into (amount, currency).

    The amount keeps whatever decimals Odoo wrote, so "0.00" stays "0.00"
    rather than collapsing to 0; thousands separators are stripped so the
    column stays usable in a formula. The PO reference is discarded.
    """
    if raw is None or raw is False or raw == "":
        return "", ""
    if isinstance(raw, (int, float)) and not isinstance(raw, bool):
        return f"{raw:.2f}", ""
    text = str(raw)
    amount = LAST_PURCHASE_AMOUNT.search(text)
    currency = LAST_PURCHASE_CURRENCY.search(text)
    return (
        amount.group(0).replace(",", "") if amount else "",
        currency.group(0) if currency else "",
    )


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
    needed_cols = len(FLAT_HEADERS)
    # Columns are resized to exactly what's needed, not just grown, so a
    # trimmed-down COLUMNS list doesn't leave stale blank columns sitting
    # past the data from a wider run of the script. Rows only grow, so
    # anything a person added manually below the data is left alone.
    if ws.row_count < needed_rows or ws.col_count != needed_cols:
        retry_gspread(
            ws.resize,
            rows=max(needed_rows, ws.row_count),
            cols=needed_cols,
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
        model = SOURCE_MODEL.get(source)
        if model is None:
            resolved[(source, name)] = name  # product lookups, checked below
            continue
        meta = order_meta if model == "order" else line_meta
        actual = resolve(meta, name)
        resolved[(source, name)] = actual
        if actual is None:
            print(f"  note: '{name}' ({label}) is not on this database; column left blank")
        elif model == "order":
            order_fields.append(actual)
        else:
            line_fields.append(actual)

    # Only reach for a product when a product-sourced column survives; with
    # none, the extra product read is pure waste.
    order_product_names = [n for n, _, source in COLUMNS if source == "order_product"]
    line_product_names = [n for n, _, source in COLUMNS if source == "line_product"]

    order_product_field = resolve(order_meta, "product_id") if order_product_names else None
    if order_product_field:
        order_fields.append(order_product_field)
    line_product_field = resolve(line_meta, "product_id") if line_product_names else None

    product_fields = [
        f for f in dict.fromkeys(order_product_names + line_product_names) if f in product_meta
    ]

    # Filtered by creation date and company; status is not otherwise tracked,
    # except that cancelled orders are excluded. Orders in any other state
    # (draft, sent, purchase, done, ...) are included. The context's
    # allowed_company_ids only steers record rules, so the companies are
    # pinned in the domain rather than relied on.
    end_utc = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
    domain = [
        "&", "&", "&",
        ["create_date", ">=", START_UTC],
        ["create_date", "<=", end_utc],
        ["company_id", "in", COMPANY_IDS],
        ["state", "!=", "cancel"],
    ]
    print(
        f"Fetching purchase orders for companies {COMPANY_IDS}, "
        f"{START_UTC} to {end_utc} UTC (by creation date, excluding cancelled)..."
    )

    orders = search_read(cookies, "purchase.order", domain, sorted(set(order_fields)))
    print(f"Total orders: {len(orders)}")

    lines = []
    if orders:
        order_ids = [o["id"] for o in orders]
        wanted = set(line_fields) | {"order_id"}
        if line_product_field:
            wanted.add(line_product_field)
        wanted |= {f for f in LINE_BOOKKEEPING_FIELDS if f in line_meta}
        if "display_type" not in line_meta:
            print("  warning: order lines have no display_type; Responsible will be blank")
        for i in range(0, len(order_ids), 500):
            chunk = order_ids[i:i + 500]
            lines.extend(
                search_read(
                    cookies,
                    "purchase.order.line",
                    [["order_id", "in", chunk]],
                    sorted(wanted),
                    order="sequence, id",
                )
            )
    print(f"Total order lines: {len(lines)}")

    grouped = {}
    for line in lines:
        order_id = line.get("order_id")
        oid = order_id[0] if isinstance(order_id, list) and order_id else order_id
        grouped.setdefault(oid, []).append(line)

    # A relevant section heading (see section_is_relevant) governs every
    # product line beneath it until the next relevant heading. Consecutive
    # relevant headings before any product line accumulate into one group
    # (joined with ", "); a non-relevant heading in between is skipped rather
    # than breaking that group or clearing it. The heading rows themselves
    # are dropped: they hold no product, so leaving them in would put a blank
    # row in the sheet for every section.
    lines_by_order = {}
    section_count = 0
    orders_with_sections = set()
    for oid, order_lines in grouped.items():
        order_lines.sort(key=lambda l: (l.get("sequence") or 0, l.get("id") or 0))
        section_parts = []
        parts_used = True
        for line in order_lines:
            if line.get("display_type") == SECTION_DISPLAY_TYPE:
                section_count += 1
                orders_with_sections.add(oid)
                heading = (line.get("name") or "").strip()
                if section_is_relevant(heading):
                    if parts_used:
                        section_parts = []
                        parts_used = False
                    section_parts.append(heading)
            elif line.get("display_type"):
                continue  # a note row; carries no product either
            else:
                line["_section"] = ", ".join(section_parts)
                parts_used = True
                lines_by_order.setdefault(oid, []).append(line)

    print(
        f"Of those, {section_count} are section headings across "
        f"{len(orders_with_sections)} orders."
    )

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
                elif source in ("last_purchase_value", "last_purchase_currency"):
                    amount, currency = parse_last_purchase(
                        line.get(actual) if (line and actual) else ""
                    )
                    row.append(amount if source == "last_purchase_value" else currency)
                elif source == "sections":
                    row.append(line.get("_section", "") if line else "")
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
