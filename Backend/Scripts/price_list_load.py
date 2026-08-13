"""
price_list_load.py
--------------------
Loads Price List rows from all .xlsx files in the configured Price List
folder into QAD via the priceListV2s API. Mirrors Customer_load.py's shape:
one TokenManager per run, GET->patch->POST for updates, field-level error
highlighting via resolve_qad_errors(), progress callback for SSE.

CHANGE LOG (this revision)
---------------------------
AUTH REWORK (same fix applied to Customer_load.py / Supplier_load.py):
config.json no longer has a "qad" key with a stored service-account
username/password -- it now has "environments": {"TEST": {...},
"PROD": {...}}, which only holds base_url/client_id/grant_type, NOT
credentials. This script can no longer fetch its own token, so
TokenManager now just holds a token + base_url handed to it by the caller
(main.py's session, or the __main__ CLI login block below) instead of
calling _fetch_token(). Every function that built a URL from
CONFIG['qad']['base_url'] (post_price_list, get_price_list,
search_price_list_by_key, and by extension expire_existing_record) now
takes an explicit base_url parameter instead. run() now takes a
TokenManager instead of building its own.

CONFIRMED business rules this loader assumes (2026-07-23):
  - Amount is QAD-calculated. Never sent in any payload, never patched.
  - Break Category is free text, no lookup/format constraint.
  - One priceListDetailV2 per key -- no quantity-break handling.
  - Attribute Code / Order Code are always "" (system-defaulted); this
    loader hardcodes them and never reads them from Excel.
  - LOCKED_ON_UPDATE fields (Domain, Price List, Customer Code, Item Code,
    Currency, Unit Of Measure, Start Date) are protected in Excel and are
    read directly from the current cell -- they cannot have changed since
    fetch.
  - Currency is QAD-locked. It is part of the record's identity/lookup key
    and is never editable, never patched. There is no hidden "_Fetch
    Currency" column -- Currency is read straight from the visible cell
    for both the GET key and (never) for patching.
  - Amount Type / Quantity Type / Combination Type are stored in Excel as
    friendly text (e.g. "Discount %", "Quantity", "Exclusive") and must be
    converted to QAD's internal numeric codes via mappings_price_list.py
    before being sent, on both CREATE and UPDATE.
  - priceListConfs (business-rule confirmation messages, e.g. "Price List
    will apply to all customers and items") do NOT block the load -- they
    are informational. Logged into the Error/note column as a suffix on
    success, never treated as failure, no second confirming POST issued.

CREATE flow (Data Operation = C):
  1. Build the business key (Domain, Price List, Customer, Item, Currency,
     Unit Of Measure -- dates are NOT part of the key).
  2. Search QAD for any existing record(s) sharing that key, regardless of
     Start Date, using the priceListV2s BROWSE (api/qracore/browses,
     browseId urn:browse:bebrowse:com.qad.erp.sales.priceListV2s) -- NOT
     the plain priceListV2s GET, which is a keyed single-record fetch and
     400s without a Start Date. Multiple conditions are sent as repeated
     filter= params shaped "field,eq,value,literal" (confirmed via network
     capture 2026-07-24). A blank Customer Code / Item Code in Excel means
     "applies to all" and must be searched for using QAD's wildcard
     literal ("qadall--+--+--+--+--+"), since that's what such records
     actually store in those fields -- not blank.
  3. AUTO-VERSIONING (2026-07-24 enhancement): if the incoming Start Date
     falls inside an existing record's validity window (Expire Date
     blank/null = open-ended, not a special case -- just no upper bound),
     that existing record is automatically expired one day before the new
     Start Date (re-fetch via the keyed GET -> patch expireDate -> POST,
     same shape as the UPDATE flow) before the new record is created. This
     only runs for Data Operation = C; Update rows are untouched. If the
     expire step fails, the row is marked as an error and the new record
     is NOT created, so we never end up with an overlap.
  4. Build payload from Excel (fresh record, empty priceListDetailV2s)
  5. POST to priceListV2s
  (No domain-settings follow-up step like Customer -- Price List has none.)

UPDATE flow (Data Operation = U):
  1. GET existing record using LOCKED_ON_UPDATE fields (Currency included,
     read directly from the current cell)
  2. Patch editable fields onto the fetched payload
  3. POST the whole record back

DONE-skip rule:
  - "DONE" + Data Operation "C" rows are skipped forever (already created).
  - "DONE" + Data Operation "U" rows are always reloadable, since a user
    may fetch -> edit -> load -> edit again -> load again.
"""

import os
import sys
import time
import json
import uuid
from datetime import datetime, timedelta

import requests
import openpyxl

SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR    = os.path.abspath(os.path.join(SCRIPTS_DIR, ".."))
sys.path.insert(0, ROOT_DIR)
sys.path.insert(0, SCRIPTS_DIR)
from config import CONFIG
from excel_format_utils import mark_done, mark_error, resolve_qad_errors
from qad_response_utils import normalize_response
from mappings_price_list import (
    get_amount_type_code,
    get_quantity_type_code,
    get_combination_type_code,
)

DETAIL_VIEW_URI = "urn:be:com.qad.sales.pricing.IPriceListV2"

# Browse (list/search) resource -- confirmed via network capture 2026-07-24.
# The plain priceListV2s GET is a keyed single-record fetch (requires the
# full key including Start Date); it 400s if Start Date is omitted. To find
# *every* record sharing a partial key (needed for auto-versioning) QAD
# requires this separate browse endpoint instead.
PRICE_LIST_BROWSE_ID = "urn:browse:bebrowse:com.qad.erp.sales.priceListV2s"

# QAD's placeholder for "applies to all customers/items" on a Price List --
# confirmed via network capture: a record with no specific Customer/Item
# comes back with this literal string in customerCode/itemCode rather than
# a blank value. Excel rows with a blank Customer Code / Item Code (i.e.
# "applies to all") must be searched for using this wildcard, not "".
QAD_ALL_WILDCARD = "qadall--+--+--+--+--+"


def _s(row_data: dict, col: str) -> str:
    """
    Safely read a cell as a stripped string. row_data.get(col, "") does NOT
    protect against None -- an empty Excel cell produces a dict entry that
    EXISTS with value None, so the default is never used and str(None)
    silently becomes the literal string "None". Use this everywhere instead
    of str(row_data.get(...)).
    """
    v = row_data.get(col)
    return str(v).strip() if v is not None else ""

# ── Progress callback (same convention as Customer_load.py) ───────────────
_progress_callback = None


def set_progress_callback(callback):
    global _progress_callback
    _progress_callback = callback


# ── Mandatory columns ──────────────────────────────────────────────────────
MANDATORY_COLUMNS = [
    "Domain",
    "Price List",
    "Currency",
]

# Locked on Update -- read directly from the current cell, never patched,
# never re-derived, because Excel protection prevents them from changing.
# Currency is locked: it is part of QAD's record identity, not editable.
LOCKED_ON_UPDATE = [
    "Domain", "Price List", "Customer Code", "Item Code",
    "Currency", "Unit Of Measure", "Start Date",
]

# ── QAD fieldName -> Excel column, for highlighting the right cell on error ─
QAD_FIELD_TO_COLUMN = {
    "domainCode":           "Domain",
    "priceListCode":        "Price List",
    "priceListDescription": "Description",
    "customerCode":         "Customer Code",
    "itemCode":             "Item Code",
    "currencyCode":         "Currency",
    "unitOfMeasure":        "Unit Of Measure",
    "startDate":            "Start Date",
    "expireDate":           "Expire Date",
    "amountType":           "Amount Type",
    "quantityType":         "Quantity Type",
    "combinationType":      "Combination Type",
    "minimumQuantity":      "Minimum Order",
    "maximumQty":           "Maximum Quantity",
    "maxOrders":            "Maximum orders",
    "breakCategory":        "Break category",
}


# =============================================================================
# 1. AUTHENTICATION
# =============================================================================
#
# CHANGED: this script no longer fetches its own OAuth token. config.json's
# old "qad" key (which held a service-account username/password) is gone,
# replaced by "environments": {"TEST": {...}, "PROD": {...}} -- and those
# blocks only have base_url/client_id/grant_type, no password. The real
# credentials only exist for the moment a user logs in via main.py's
# /api/login, which keeps the resulting access_token server-side in
# SESSIONS[session_id]. TokenManager is now just a holder for that token
# plus the base_url it's valid against, handed in by the caller.

class TokenManager:
    """Holds the QAD access token + base_url for one run.

    token:    OAuth access token already obtained by main.py at login
              (stored server-side in SESSIONS[session_id]), or obtained by
              the __main__ CLI login block below when run standalone.
    base_url: the qracore base_url for the session's environment (same
              host used for the original OAuth call).
    """

    def __init__(self, token: str, base_url: str):
        if not token:
            raise ValueError("TokenManager requires a valid session access_token")
        if not base_url:
            raise ValueError("TokenManager requires a base_url")
        self._token   = token
        self.base_url = base_url.rstrip("/")

    def get(self) -> str:
        return self._token

    def refresh(self) -> str:
        # No stored credentials to re-authenticate with from here -- the
        # session's token came from the user's login and this script never
        # had the password. Surface this clearly instead of silently
        # failing to hit an /oauth/token endpoint with nothing to send.
        raise RuntimeError(
            "QAD session token expired mid-run and cannot be refreshed "
            "automatically -- please log in again and re-run the load."
        )


class _TokenExpired(Exception):
    pass


# =============================================================================
# 2. DATE HELPERS
# =============================================================================

_DATE_STRING_FORMATS = (
    "%d-%m-%Y",
    "%Y-%m-%d",
    "%Y-%m-%d %H:%M:%S",
    "%d/%m/%Y",
    "%m/%d/%Y",
    "%d-%m-%y",
    "%d.%m.%Y",
)


def _excel_to_iso(value, _field_label: str = "date") -> str:
    """
    Excel cell -> 2026-07-22T00:00:00.000Z (QAD).

    Handles three distinct shapes a "date" cell can arrive in from openpyxl:
      1. datetime/date object       -- normal date-formatted Excel cell
      2. int/float (Excel serial)   -- cell's number format wasn't one
                                        openpyxl recognised as a date, so it
                                        handed back the raw day-count serial
                                        (days since 1899-12-30, with Excel's
                                        historical leap-year-1900 quirk)
      3. string                     -- text cell, tried against several
                                        common layouts, not just dd-mm-yyyy

    If none of these match, this is no longer swallowed silently -- it's
    printed to console with the raw repr so the actual offending value is
    visible instead of a mysteriously blank startDate/expireDate in QAD.
    """
    if value is None or value == "":
        return ""

    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%dT00:00:00.000Z")

    if hasattr(value, "strftime"):  # datetime.date (no time component)
        return value.strftime("%Y-%m-%dT00:00:00.000Z")

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            # Excel serial date: day 1 = 1899-12-31, but Excel treats 1900
            # as a leap year (it wasn't), so anchor at 1899-12-30 to land
            # on the correct real-world date for serials >= 60.
            dt = datetime(1899, 12, 30) + timedelta(days=value)
            return dt.strftime("%Y-%m-%dT00:00:00.000Z")
        except (OverflowError, ValueError):
            pass

    v = str(value).strip()
    if not v:
        return ""

    for fmt in _DATE_STRING_FORMATS:
        try:
            dt = datetime.strptime(v, fmt)
            return dt.strftime("%Y-%m-%dT00:00:00.000Z")
        except ValueError:
            continue

    print(f"[price_list_load] WARNING: could not parse {_field_label} value "
          f"{value!r} (type={type(value).__name__}) -- sending blank to QAD")
    return ""


def _uri_safe_date(iso_value: str) -> str:
    """QAD's URI scheme percent-encodes the date's internal '.' as %2E to
    avoid colliding with the URI's own dot-separated key segments."""
    return iso_value.replace(".", "%2E")


def _iso_to_date(iso_value: str):
    """
    QAD iso string (2026-07-22T00:00:00.000Z) -> datetime.date, or None if
    blank/unparseable. Used only for the auto-versioning overlap check --
    never sent back to QAD (we always resend the original iso string).
    """
    if not iso_value:
        return None
    try:
        return datetime.strptime(iso_value, "%Y-%m-%dT00:00:00.000Z").date()
    except ValueError:
        return None


def _shift_iso_date(iso_value: str, days: int) -> str:
    """Shift a QAD iso date string by `days` (may be negative) and return
    it in the same iso format."""
    dt = datetime.strptime(iso_value, "%Y-%m-%dT00:00:00.000Z")
    dt = dt + timedelta(days=days)
    return dt.strftime("%Y-%m-%dT00:00:00.000Z")


# =============================================================================
# 3. CREATE PAYLOAD BUILDER
# =============================================================================

def build_payload(row: dict, domain: str) -> dict:
    def val(col: str) -> str:
        v = row.get(col)
        return str(v).strip() if v is not None else ""

    price_list_code = val("Price List")
    customer_code    = val("Customer Code")
    item_code        = val("Item Code")
    currency_code    = val("Currency")
    uom              = val("Unit Of Measure")
    start_iso        = _excel_to_iso(val("Start Date"), "Start Date")
    expire_iso       = _excel_to_iso(val("Expire Date"), "Expire Date")

    uri = (
        f"urn:be:com.qad.sales.pricing.IPriceListV2:"
        f"{domain}.{price_list_code}.{customer_code}.{item_code}...{currency_code}.{uom}."
        f"{_uri_safe_date(start_iso)}"
    )

    return {
        "supplementaryMessages": [],
        "priceListV2s": [
            {
                "uri":                     uri,
                "domainCode":              domain,
                "priceListCode":           price_list_code,
                "priceListDescription":    val("Description"),
                "customerCode":            customer_code,
                "itemCode":                item_code,
                "attributeCode":           "",   # always system-defaulted -- confirmed
                "orderCode":               "",   # always system-defaulted -- confirmed
                "currencyCode":            currency_code,
                "unitOfMeasure":           uom,
                "startDate":               start_iso,
                "expireDate":              expire_iso,
                "amountType":              get_amount_type_code(val("Amount Type")),
                "quantityType":            get_quantity_type_code(val("Quantity Type")),
                "combinationType":         get_combination_type_code(val("Combination Type")),
                "minNetOrd":               float(val("Minimum Order") or 0),
                "maximumQty":              float(val("Maximum Quantity") or 0),
                "maxOrders":               float(val("Maximum orders") or 0),
                "breakCategory":           val("Break category"),
                #----- list pricing ----#
                "maximumPrice":            float(val("Maximum Price") or 0),
                "minimumPrice":            float(val("Minimum Price") or 0),
                "listPrice":               float(val("List Price") or 0),
                "dataOperation":           "C",
                "concurrencyHash":         "",
                "priceListId":             "",   # QAD-assigned on save, never sent by us
                "priceListGroup":          "",
                "searchType":              2,
                "priceListDetailV2s": [
                    {
                        # Amount deliberately omitted -- QAD calculates it.
                        "dataOperation":   "C",
                    }
                ],
            }
        ],
    }


# =============================================================================
# 4. PRICE LIST API
# =============================================================================
#
# CHANGED: post_price_list, get_price_list, and search_price_list_by_key
# used to build their URLs from CONFIG['qad']['base_url']. That key no
# longer exists in config.json, so each now takes an explicit base_url
# parameter instead -- supplied by the caller (process_file /
# expire_existing_record), which gets it from tm.base_url.

def post_price_list(payload: dict, token: str, base_url: str, key_params: dict | None = None) -> tuple[bool, list, list]:
    """
    POST payload to QAD. Returns (success, errors, confirmation_messages).
    confirmation_messages comes from priceListConfs -- informational warnings
    that do NOT block success (confirmed: loads complete even with them).
    """
    url = f"{base_url}/api/erp/priceListV2s"
    params = {"viewUri": DETAIL_VIEW_URI}

    if key_params:
        params.update(key_params)

    resp = requests.post(
        url,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type":  "application/json",
        },
        params=params,
        json=payload,
        timeout=30,
    )

    if resp.status_code == 401:
        raise _TokenExpired()

        # ── TEMP DEBUG BLOCK ────────────────────────────────────────────────────
    print("\n" + "="*70)
    print(f"DEBUG post_price_list | HTTP {resp.status_code}")
    print("="*70)
    print("RAW response text:")
    print(resp.text)
    print("-"*70)
    try:
        _debug_json = resp.json()
        print("Parsed JSON top-level keys:", list(_debug_json.keys()))
        print("submitResult:", json.dumps(_debug_json.get("submitResult"), indent=2))
        print("priceListConfs:", json.dumps(_debug_json.get("priceListConfs"), indent=2))
        pl = (_debug_json.get("priceListV2s") or [{}])[0]
        print("Returned priceListId:", pl.get("priceListId"))
        print("Returned concurrencyHash:", pl.get("concurrencyHash"))
    except Exception as _e:
        print(f"Could not parse response as JSON: {_e}")
    print("="*70 + "\n")
# ── END TEMP DEBUG BLOCK ─────────────────────────────────────────────────


    try:
        resp_json = resp.json()
    except Exception:
        resp_json = {}

    success, errors = normalize_response(resp_json, resp.status_code)

    conf_messages = [
        c.get("messageText", "")
        for c in (resp_json.get("priceListConfs") or [])
        if c.get("messageText")
    ]

    return success, errors, conf_messages


def get_price_list(key: dict, token: str, base_url: str) -> dict:
    url = f"{base_url}/api/erp/priceListV2s"
    resp = requests.get(
        url,
        headers={"Authorization": f"Bearer {token}"},
        params={
            "domainCode":     key.get("domainCode", ""),
            "priceListCode":  key.get("priceListCode", ""),
            "customerCode":   key.get("customerCode", ""),
            "itemCode":       key.get("itemCode", ""),
            "attributeCode":  "",
            "orderCode":      "",
            "currencyCode":   key.get("currencyCode", ""),
            "unitOfMeasure":  key.get("unitOfMeasure", ""),
            "startDate":      key.get("startDate", ""),
            "viewUri":        DETAIL_VIEW_URI,
        },
        timeout=30,
    )
    if resp.status_code == 401:
        raise _TokenExpired()
    resp.raise_for_status()
    return resp.json()


def _browse_filter_value(raw_value: str) -> str:
    """
    AUTO-VERSIONING support. Excel leaves Customer Code / Item Code blank
    to mean "applies to all customers/items"; QAD's browse stores and
    matches that as the literal wildcard string, not blank -- so a blank
    cell must be searched for using the wildcard, or it will never match.
    """
    v = (raw_value or "").strip()
    return v if v else QAD_ALL_WILDCARD


def search_price_list_by_key(key: dict, token: str, base_url: str) -> list:
    """
    AUTO-VERSIONING support. Search QAD's priceListV2s browse for every
    record sharing the business key (Domain, Price List, Customer, Item,
    Currency, Unit Of Measure) -- deliberately with no Start Date filter,
    since dates are not part of the business key and we need every
    validity window on file to check for an overlap.

    Confirmed via network capture 2026-07-24: this is a *different*
    endpoint (api/qracore/browses) from the keyed single-record GET used
    elsewhere in this file -- the plain priceListV2s GET 400s if Start
    Date is left out, it does not double as a search. Multiple conditions
    are sent as repeated `filter=` query params, each shaped
    "field,eq,value,literal".

    Only called for Data Operation = C rows (see module docstring). Returns
    a list of raw browse-row dicts (priceListV2.<field> keys, may be
    empty). Never raises for "no match", only for real HTTP/auth failures.
    """
    url = f"{base_url}/api/qracore/browses"

    filters = [
        f"priceListV2.domainCode,eq,{key.get('domainCode', '')},literal",
        f"priceListV2.priceListCode,eq,{key.get('priceListCode', '')},literal",
        f"priceListV2.customerCode,eq,{_browse_filter_value(key.get('customerCode', ''))},literal",
        f"priceListV2.itemCode,eq,{_browse_filter_value(key.get('itemCode', ''))},literal",
        f"priceListV2.currencyCode,eq,{key.get('currencyCode', '')},literal",
        f"priceListV2.unitOfMeasure,eq,{key.get('unitOfMeasure', '')},literal",
    ]

    params = [
        ("browseId",   PRICE_LIST_BROWSE_ID),
        ("callId",     str(uuid.uuid4())),
        ("page",       1),
        ("pageSize",   100),
        ("pageAction", "first"),
    ] + [("filter", f) for f in filters]

    resp = requests.get(
        url,
        headers={"Authorization": f"Bearer {token}"},
        params=params,
        timeout=30,
    )
    if resp.status_code == 401:
        raise _TokenExpired()
    resp.raise_for_status()
    body = resp.json()

    rows = body.get("data") or []
    aux  = body.get("auxillaryData") or {}
    if str(aux.get("HasMoreParams", "")).lower() == "true":
        print(f"[price_list_load] WARNING: auto-versioning search for key {key} "
              f"returned a full page ({len(rows)} rows) -- there may be more "
              f"matching records than pageSize=100 could return")

    return rows


def find_overlapping_record(records: list, incoming_start_date) -> dict | None:
    """
    AUTO-VERSIONING support. Given the raw browse rows returned for a
    business key (priceListV2.<field>-prefixed, dates as null or iso
    string), return the one whose validity window contains
    incoming_start_date, or None if there's no overlap.

    Expire Date blank/null is NOT special-cased -- it just means
    open-ended / unlimited validity, per the spec's clarification. The
    overlap test is purely: existing_start <= incoming_start <= (existing
    expire or +inf).
    """
    for rec in records:
        existing_start = _iso_to_date(rec.get("priceListV2.startDate") or "")
        if existing_start is None:
            continue

        expire_raw = rec.get("priceListV2.expireDate") or ""
        existing_expire = _iso_to_date(expire_raw) if expire_raw else None

        if existing_start <= incoming_start_date and (
            existing_expire is None or existing_expire >= incoming_start_date
        ):
            return rec

    return None


def expire_existing_record(rec: dict, new_expire_iso: str, token: str, base_url: str) -> tuple[bool, list]:
    """
    AUTO-VERSIONING support. Shrinks an existing record's validity window
    so it ends the day before an incoming Create row's Start Date. Mirrors
    the UPDATE flow's fetch -> patch -> post shape exactly (re-fetch the
    full record by its own key/Start Date via the keyed GET, patch only
    expireDate, POST the whole thing back) so we never send QAD a partial
    record.

    `rec` comes from the browse search, so its fields are
    priceListV2.<field>-prefixed -- including the wildcard string for
    customerCode/itemCode when the record applies to all customers/items,
    which is exactly what the keyed GET needs to re-fetch it.
    """
    key = {
        "domainCode":    rec.get("priceListV2.domainCode", ""),
        "priceListCode": rec.get("priceListV2.priceListCode", ""),
        "customerCode":  rec.get("priceListV2.customerCode", ""),
        "itemCode":      rec.get("priceListV2.itemCode", ""),
        "currencyCode":  rec.get("priceListV2.currencyCode", ""),
        "unitOfMeasure": rec.get("priceListV2.unitOfMeasure", ""),
        "startDate":     rec.get("priceListV2.startDate", ""),
    }

    existing = get_price_list(key, token, base_url)
    pl_list = existing.get("data", {}).get("priceListV2s") or existing.get("priceListV2s") or []
    if not pl_list:
        return False, [{
            "fieldName": None,
            "message": "Auto-versioning: existing record could not be re-fetched to expire it",
            "fieldValue": None,
            "code": None,
        }]

    payload = existing.get("data", existing)
    pl = payload["priceListV2s"][0]
    pl["expireDate"] = new_expire_iso

    post_key_params = {
        "domainCode":    pl.get("domainCode", ""),
        "priceListCode": pl.get("priceListCode", ""),
        "customerCode":  pl.get("customerCode", ""),
        "itemCode":      pl.get("itemCode", ""),
        "currencyCode":  pl.get("currencyCode", ""),
        "unitOfMeasure": pl.get("unitOfMeasure", ""),
        "startDate":     pl.get("startDate", ""),
        "attributeCode": pl.get("attributeCode", ""),
        "orderCode":     pl.get("orderCode", ""),
    }

    success, errors, _ = post_price_list(payload, token, base_url, post_key_params)
    return success, errors


# =============================================================================
# 5. MANDATORY FIELD CHECK  (belt-and-suspenders; Validate should have
#    already caught these, but Load re-checks in case a file skipped
#    Validate)
# =============================================================================

def _check_mandatory(row_data: dict) -> list[str]:
    missing = []
    for col in MANDATORY_COLUMNS:
        v = row_data.get(col)
        if v is None or str(v).strip() == "" or str(v).strip().lower() == "none":
            missing.append(col)
    return missing


# =============================================================================
# 6. SINGLE-FILE PROCESSOR
# =============================================================================

def process_file(file_path: str, tm: TokenManager) -> tuple[int, int]:
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active

    header_row = [
        str(c.value).strip() if c.value is not None else ""
        for c in ws[2]
    ]

    for col_name in ("Status", "Error"):
        if col_name not in header_row:
            ws.cell(row=2, column=len(header_row) + 1, value=col_name)
            header_row.append(col_name)

    status_col = header_row.index("Status") + 1 if "Status" in header_row else header_row.index("Status ") + 1
    error_col  = header_row.index("Error")  + 1

    ok_count   = 0
    fail_count = 0

    total_rows = ws.max_row - 2

    for row_idx, row in enumerate(ws.iter_rows(min_row=3), start=3):

        if _progress_callback:
            _progress_callback(row_idx - 2, total_rows)

        row_values = [cell.value for cell in row]
        if not any(row_values):
            continue

        row_data = dict(zip(header_row, row_values))

        status         = (_s(row_data, "Status") or _s(row_data, "Status ")).upper()
        data_operation = _s(row_data, "Data Operation").upper()

        # Created rows skip forever once DONE. Update rows must always
        # remain reloadable (fetch -> edit -> load -> edit again -> load).
        if status == "DONE" and data_operation == "C":
            continue

        missing = _check_mandatory(row_data)
        if missing:
            fail_count += 1
            mark_error(
                ws, row_idx, status_col, error_col, header_row,
                missing, f"Missing mandatory fields: {', '.join(missing)}",
            )
            wb.save(file_path)
            continue

        if data_operation not in {"C", "U"}:
            fail_count += 1
            mark_error(
                ws, row_idx, status_col, error_col, header_row,
                ["Data Operation"], "Data Operation must be C or U",
            )
            wb.save(file_path)
            continue

        domain = _s(row_data, "Domain")

        # ── Build payload ─────────────────────────────────────────────────
        if data_operation == "U":
            # Currency is locked on update: part of the record's identity,
            # read straight from the current cell like every other
            # LOCKED_ON_UPDATE field. No hidden fetch column, never patched.
            currency_code = _s(row_data, "Currency")

            key = {
                "domainCode":    domain,
                "priceListCode": _s(row_data, "Price List"),
                "customerCode":  _s(row_data, "Customer Code"),
                "itemCode":      _s(row_data, "Item Code"),
                "currencyCode":  currency_code,
                "unitOfMeasure": _s(row_data, "Unit Of Measure"),
                "startDate":     _excel_to_iso(row_data.get("Start Date", ""), "Start Date"),
            }

            existing = get_price_list(key, tm.get(), tm.base_url)
            pl_list = existing.get("data", {}).get("priceListV2s") or existing.get("priceListV2s") or []

            if not pl_list:
                fail_count += 1
                mark_error(
                    ws, row_idx, status_col, error_col, header_row, [],
                    f"UPDATE failed: Price List '{key['priceListCode']}' not found in QAD "
                    f"(looked up with Currency='{currency_code}')",
                )
                wb.save(file_path)
                continue

            payload  = existing.get("data", existing)
            pl       = payload["priceListV2s"][0]

            def _patch_if_present(target_dict, key_name, raw_value, transform=lambda v: v):
                if raw_value is not None and str(raw_value).strip() != "":
                    transformed = transform(str(raw_value).strip())
                    if transformed is None:
                        print(f"[price_list_load] WARNING: '{raw_value}' did not match any known code for "
                            f"'{key_name}' -- skipping this field rather than sending null to QAD")
                        return
                    target_dict[key_name] = transformed


            _patch_if_present(pl, "priceListDescription", row_data.get("Description"))
            # Currency intentionally NOT patched -- locked on update.
            _patch_if_present(pl, "expireDate",           row_data.get("Expire Date"), _excel_to_iso)
            _patch_if_present(pl, "amountType",           row_data.get("Amount Type"), get_amount_type_code)
            _patch_if_present(pl, "quantityType",         row_data.get("Quantity Type"), get_quantity_type_code)
            _patch_if_present(pl, "combinationType",      row_data.get("Combination Type"), get_combination_type_code)
            _patch_if_present(pl, "minNetOrd",            row_data.get("Minimum Order"), float)
            _patch_if_present(pl, "combinationType",      row_data.get("Combination Type"), get_combination_type_code)
            _patch_if_present(pl, "maximumQty",           row_data.get("Maximum Quantity"), float)
            _patch_if_present(pl, "maxOrders",            row_data.get("Maximum orders"),   float)
            _patch_if_present(pl, "maximumPrice",         row_data.get("Maximum Price"), float)
            _patch_if_present(pl, "minimumPrice",         row_data.get("Minimum Price"), float)
            _patch_if_present(pl, "listPrice",            row_data.get("List Price"), float)
            _patch_if_present(pl, "breakCategory",        row_data.get("Break category"))


            post_key_params = {
                "domainCode":    pl.get("domainCode", ""),
                "priceListCode": pl.get("priceListCode", ""),
                "customerCode":  pl.get("customerCode", ""),
                "itemCode":      pl.get("itemCode", ""),
                "currencyCode":  pl.get("currencyCode", ""),
                "unitOfMeasure": pl.get("unitOfMeasure", ""),
                "startDate":     pl.get("startDate", ""),
                "attributeCode":  pl.get("attributeCode", ""),
                "orderCode":      pl.get("orderCode", ""),
            }

        else:
            # ── CREATE: auto-versioning (2026-07-24 enhancement) ─────────
            # Business key deliberately excludes dates. Only runs for
            # Data Operation = C -- Updates are untouched (see module
            # docstring / spec).
            business_key = {
                "domainCode":    domain,
                "priceListCode": _s(row_data, "Price List"),
                "customerCode":  _s(row_data, "Customer Code"),
                "itemCode":      _s(row_data, "Item Code"),
                "currencyCode":  _s(row_data, "Currency"),
                "unitOfMeasure": _s(row_data, "Unit Of Measure"),
            }
            incoming_start_iso  = _excel_to_iso(row_data.get("Start Date", ""), "Start Date")
            incoming_start_date = _iso_to_date(incoming_start_iso)

            versioning_error = None

            if incoming_start_date is not None:
                try:
                    try:
                        matches = search_price_list_by_key(business_key, tm.get(), tm.base_url)
                    except _TokenExpired:
                        tm.refresh()
                        matches = search_price_list_by_key(business_key, tm.get(), tm.base_url)
                except requests.RequestException as e:
                    matches = []
                    versioning_error = f"Auto-versioning lookup failed: {e}"

                if versioning_error is None:
                    overlap = find_overlapping_record(matches, incoming_start_date)
                    if overlap is not None:
                        # Overlap found -> expire the existing record one
                        # day before the new Start Date, per the spec's
                        # decision rule, before creating the replacement.
                        new_expire_iso = _shift_iso_date(incoming_start_iso, -1)
                        try:
                            try:
                                exp_success, exp_errors = expire_existing_record(overlap, new_expire_iso, tm.get(), tm.base_url)
                            except _TokenExpired:
                                tm.refresh()
                                exp_success, exp_errors = expire_existing_record(overlap, new_expire_iso, tm.get(), tm.base_url)
                        except requests.RequestException as e:
                            exp_success, exp_errors = False, [{
                                "fieldName": None, "message": f"Network error: {e}",
                                "fieldValue": None, "code": None,
                            }]

                        if not exp_success:
                            _, exp_msg = resolve_qad_errors(exp_errors, QAD_FIELD_TO_COLUMN)
                            versioning_error = f"could not auto-expire existing record: {exp_msg}"

            if versioning_error:
                # Never create the new record if we couldn't safely close
                # out the old one first -- that would leave an overlap.
                fail_count += 1
                mark_error(
                    ws, row_idx, status_col, error_col, header_row,
                    ["Start Date"], f"Auto-versioning failed, record not created: {versioning_error}",
                )
                wb.save(file_path)
                continue

            payload = build_payload(row_data, domain)
            post_key_params = None  # no key params for create

        # ── POST (with one token-refresh retry) ────────────────────────────
        success   = False
        errors    = []
        conf_msgs = []

        for attempt in range(2):
            try:
                success, errors, conf_msgs = post_price_list(payload, tm.get(), tm.base_url, post_key_params)
                break
            except _TokenExpired:
                if attempt == 0:
                    try:
                        tm.refresh()
                        continue
                    except RuntimeError as refresh_exc:
                        errors = [{"fieldName": None, "message": str(refresh_exc), "fieldValue": None, "code": None}]
                        break
                errors = [{"fieldName": None, "message": "Token refresh failed -- unauthorised", "fieldValue": None, "code": None}]
                break
            except requests.RequestException as e:
                errors = [{"fieldName": None, "message": f"Network error: {e}", "fieldValue": None, "code": None}]
                break

        if not success:
            fail_count += 1
            bad_cols, error_msg = resolve_qad_errors(errors, QAD_FIELD_TO_COLUMN)
            mark_error(ws, row_idx, status_col, error_col, header_row, bad_cols, error_msg)
            wb.save(file_path)
            time.sleep(0.1)
            continue

        ok_count += 1
        mark_done(ws, row_idx, status_col, error_col)
        if conf_msgs:
            # Informational only -- confirmed these do not block the load.
            # Recorded so users can see QAD's business-rule notes even on
            # a fully successful row (e.g. "applies to all customers/items").
            ws.cell(row=row_idx, column=error_col, value="Note: " + "; ".join(conf_msgs))
        wb.save(file_path)
        time.sleep(0.1)

    return ok_count, fail_count


# =============================================================================
# 7. ORCHESTRATOR
# =============================================================================
#
# CHANGED: run() used to build its own TokenManager() with zero arguments.
# Since TokenManager now requires (token, base_url), run() takes an
# already-constructed TokenManager from the caller instead -- main.py's
# session, or the __main__ CLI login block below.

def run(folder_path: str, tm: TokenManager) -> tuple[int, int]:
    folder = os.path.abspath(folder_path)

    if not os.path.exists(folder):
        raise RuntimeError(f"Folder not found: {folder}")

    xlsx_files = [
        f for f in os.listdir(folder)
        if f.endswith(".xlsx") and not f.startswith("~$")
    ]
    if not xlsx_files:
        raise RuntimeError(f"No .xlsx files found in: {folder}")

    total_ok, total_fail = 0, 0

    for file_name in xlsx_files:
        file_path = os.path.join(folder, file_name)
        ok, fail = process_file(file_path, tm)
        total_ok   += ok
        total_fail += fail
        print(f"{file_name}: ok={ok} fail={fail}")

    return total_ok, total_fail


# =============================================================================
# 8. ENTRY POINT
# =============================================================================
#
# CHANGED: the CLI path has no session to borrow a token from, so it
# performs its own one-off OAuth login here (same flow main.py's
# /api/login uses) and builds a TokenManager from the result. Set
# QAD_ENVIRONMENT / QAD_USERNAME / QAD_PASSWORD to skip the prompts.
#
# Allow either:
#   python price_list_load.py
#   python price_list_load.py D:\CustomFolder
# Same folder-argument behavior as before.

if __name__ == "__main__":
    import getpass

    environment = os.environ.get("QAD_ENVIRONMENT", "TEST").upper()
    env_cfg     = CONFIG["environments"][environment]

    username = os.environ.get("QAD_USERNAME") or input("QAD username: ")
    password = os.environ.get("QAD_PASSWORD") or getpass.getpass("QAD password: ")

    token_resp = requests.post(
        f"{env_cfg['base_url']}/oauth/token",
        data={
            "client_id":  env_cfg["client_id"],
            "username":   username,
            "password":   password,
            "grant_type": env_cfg.get("grant_type", "password"),
        },
        timeout=30,
    )
    token_resp.raise_for_status()
    access_token = token_resp.json().get("access_token")
    if not access_token:
        sys.exit("ERROR: OAuth response did not contain access_token")

    tm = TokenManager(token=access_token, base_url=env_cfg["base_url"])

    if len(sys.argv) > 1:
        folder = os.path.abspath(sys.argv[1])
    else:
        folder = os.path.abspath(os.path.join(ROOT_DIR, CONFIG["folders"]["price_list"]))

    ok, fail = run(folder, tm)
    sys.exit(0 if fail == 0 else 1)