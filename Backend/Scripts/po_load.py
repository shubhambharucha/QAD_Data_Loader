"""
PO_load.py
----------
Loads Purchase Order headers and lines from all .xlsx files in the
configured PurchaseOrder folder into QAD via the purchaseOrderHeaders /
purchaseOrderLines APIs.

Behaviour
---------
- Uses the OAuth token already obtained at login (held server-side in
  main.py's SESSIONS store) — this script no longer fetches its own token
  from config, because config.json no longer contains service-account
  credentials (see TokenManager notes below). base_url is likewise taken
  from the logged-in session's environment (tm.base_url), not a module-
  level constant, so PO now correctly follows TEST vs PROD like every
  other loader instead of always hitting TEST.
- Skips rows with Status = DONE
- Error/response handling now shares the same framework as
  Customer_load.py:
    * qad_response_utils.normalize_response() turns every QAD POST
      response into a uniform (success, errors[]) shape, regardless of
      whether QAD sent a submitResult-wrapped envelope or a bare
      top-level errors[] (gateway/permission failures).
    * excel_format_utils.resolve_qad_errors() maps each errors[]
      fieldName back to the Excel column name via PO_HEADER_FIELD_TO_COLUMN
      / PO_LINE_FIELD_TO_COLUMN below, so the exact bad cell gets
      highlighted instead of just col 1 + Error.
    * excel_format_utils.mark_done() / mark_error() replace the old
      mark_row_success() / mark_row_error() helpers and share the same
      RED_FILL / CLEAR_FILL constants used by every other loader.
- Lightweight mandatory-field check runs before any API call, for both
  Header rows and Line rows.
- Network errors, persistent-401 (token refresh failure), and gateway/
  permission errors are all surfaced as normalized error dicts so they
  flow through the exact same resolve_qad_errors() / mark_error() path
  as ordinary field validation errors.
- Saves workbook after every row so progress survives a crash.
- Renames file to error_<name> if any failures occurred (unchanged).
- Returns (total_success, total_fail)

Sheet layout (UNCHANGED)
------------------------
Workbook must contain two sheets: 'Header' and 'Lines'.
Each sheet must have a header row (row 1) with column names.

Purchase Order business logic / API workflow — UNCHANGED
----------------------------------------------------------
  Header : single POST to purchaseOrderHeaders
  Lines  : GET  init (QAD assigns line number)
           POST fieldChange(siteCode)
           POST fieldChange(itemCode)
           GET  isReceived check (fire-and-forget)
           POST fieldChange(quantityOrdered)
           POST fieldChange(purchaseCost)
           POST sync to purchaseOrderLinesGrid
None of the endpoints, payload fields, sequencing, or line-number
handling have been changed — only how responses are parsed and how
Excel is annotated.
"""

import os
import sys
import time
import requests
import openpyxl
from datetime import datetime
from collections import defaultdict

# ── Progress callback (used by main.py SSE streaming) ──────────────────

_progress_callback = None

def set_progress_callback(callback):
    global _progress_callback
    _progress_callback = callback


# ── Path / config ─────────────────────────────────────────────────────────
ROOT_DIR    = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT_DIR)
sys.path.insert(0, SCRIPTS_DIR)
from config import CONFIG
from excel_format_utils import RED_FILL, CLEAR_FILL, mark_done, mark_error, resolve_qad_errors
from qad_response_utils import normalize_response

# ── API endpoint templates ──────────────────────────────────────────────
#
# CHANGED (see explanation in chat): these used to be built once at import
# time from a module-level _BASE_URL, itself resolved from sys.argv[1] /
# CONFIG["environments"][...]. That meant PO silently always hit TEST
# whenever it was imported as a module (main.py never passes argv), no
# matter which environment the logged-in user actually selected.
#
# Now each is a small function that takes base_url explicitly, supplied by
# the caller (process_file / create_line) from tm.base_url — the same
# base_url the session's token was issued against. This mirrors
# Customer_load.py's pattern exactly.

def _header_url(base_url: str) -> str:
    return f"{base_url}/api/erp/purchaseOrderHeaders?viewUri=urn:be:com.qad.purchasing.purchaseorders.IPurchaseOrderHeader"


def _init_line_url(base_url: str, domain: str, po: str) -> str:
    return f"{base_url}/api/erp/purchaseOrderLinesGrid?initialize=true&domainCode={domain}&purchaseOrderNumber={po}"


def _field_change_url(base_url: str, field_name: str) -> str:
    return f"{base_url}/api/erp/purchaseOrderLines/fieldChangeV2?fieldName={field_name}"


def _is_received_url(base_url: str, domain: str, po: str, line: int) -> str:
    return (
        f"{base_url}/api/erp/purchaseOrderLines/isReceivedPurchaseOrderLineV2"
        f"?domainCode={domain}&purchaseOrderNumber={po}&purchaseOrderLine={line}"
    )


def _sync_line_url(base_url: str) -> str:
    return f"{base_url}/api/erp/purchaseOrderLinesGrid"

# ── Mandatory columns — Header sheet ───────────────────────────────────────
MANDATORY_HEADER_COLUMNS = [
    "PO Number",
    "Domain Code",
    "Supplier Code",
    "Currency",
    "Credit Terms",
    "Daybook Set",
    "Ship To Site",
    "Bill To Code",
]

# ── Mandatory columns — Lines sheet ─────────────────────────────────────────
MANDATORY_LINE_COLUMNS = [
    "PO Number",
    "Site Code",
    "Item Code",
    "Quantity Ordered",
    "Unit Price",
]

# ── QAD API fieldName → Excel column name (Header / purchaseOrderHeaders) ──
# Extend this map whenever a new QAD fieldName is observed in errors[]
# that should highlight a specific Excel cell. Anything NOT in this map
# still shows up in the Error message text — it just won't get its own
# highlighted cell.
PO_HEADER_FIELD_TO_COLUMN = {
    "purchaseOrderNumber": "PO Number",
    "domainCode":          "Domain Code",
    "supplierCode":        "Supplier Code",
    "currencyCode":        "Currency",
    "creditTermsCode":     "Credit Terms",
    "daybookSetCode":      "Daybook Set",
    "shipToSite":          "Ship To Site",
    "billToCode":          "Bill To Code",
    "taxEnvironment":      "Tax Environment",
    "languageCode":        "Language Code",
    "contact":             "Contact",
    "remarks":             "Remarks",
    "orderDate":           "Order Date",
    "dueDate":             "Due Date",
}

# ── QAD API fieldName → Excel column name (Lines / purchaseOrderLines) ─────
PO_LINE_FIELD_TO_COLUMN = {
    "purchaseOrderNumber": "PO Number",
    "siteCode":            "Site Code",
    "itemCode":            "Item Code",
    "quantityOrdered":     "Quantity Ordered",
    "purchaseCost":        "Unit Price",
    "dueDate":             "Due Date",
}


# =============================================================================
# 1. AUTHENTICATION
# =============================================================================
#
# CHANGED (see explanation in chat): this script used to fetch its own OAuth
# token from CONFIG['qad']['auth'] (a service-account username/password kept
# in config.json). config.json no longer has a "qad" key — it now has
# "environments": {"TEST": {...}, "PROD": {...}}, and those blocks only hold
# base_url/client_id/grant_type, NOT a username/password. The real
# username/password now only ever exists for the moment the user logs in via
# main.py's /api/login, which performs the OAuth exchange itself and keeps
# the resulting access_token server-side in SESSIONS[session_id].
#
# So TokenManager no longer knows how to mint its own token — it's handed
# the token (and the base_url that token is valid against) that main.py
# already obtained for this session, and just holds onto it for the run.
#
# If the token expires mid-run (401), this script CANNOT silently get a new
# one — it never had the password. TokenManager.refresh() raises instead of
# retrying, so process_file() surfaces a clear "please log in again" error
# on that row rather than a crash.

class TokenManager:
    """Holds the QAD access token + base_url for one run.

    token:    OAuth access token already obtained by main.py at login
              (stored server-side in SESSIONS[session_id]).
    base_url: the qracore base_url for the session's environment
              (same host used for the original OAuth call), e.g.
              CONFIG['environments']['TEST']['base_url'].
    """

    def __init__(self, token: str, base_url: str):
        if not token:
            raise ValueError("TokenManager requires a valid session access_token")
        if not base_url:
            raise ValueError("TokenManager requires a base_url")
        self._token    = token
        self.base_url  = base_url.rstrip("/")

    def get(self) -> str:
        return self._token

    def refresh(self) -> str:
        # No stored credentials to re-authenticate with from here — the
        # session's token came from the user's login and this script never
        # had the password. Surface this clearly instead of trying (and
        # failing) to hit an /oauth/token endpoint with nothing to send.
        raise RuntimeError(
            "QAD session token expired mid-run and cannot be refreshed "
            "automatically — please log in again and re-run the load."
        )

    def headers(self) -> dict:
        return {
            "Content-Type":  "application/json",
            "Authorization": f"Bearer {self.get()}",
        }


# =============================================================================
# 2. HTTP HELPERS
# =============================================================================
#
# CHANGED: on a 401, these used to call tm.refresh() and silently retry —
# that worked when refresh() could mint a brand-new token from stored
# credentials. Now refresh() raises RuntimeError (see above), so that's
# caught here: the retry is abandoned and the original 401 response is
# returned as-is, which safe_get()/safe_post() below already translate into
# a normalized "token expired" error for the row instead of an uncaught
# exception killing the whole run.

def api_get(url: str, tm: TokenManager) -> tuple[requests.Response, dict]:
    """GET with one refresh attempt on 401 (see note above)."""
    resp = None
    for attempt in range(2):
        resp = requests.get(url, headers=tm.headers(), timeout=30)
        if resp.status_code == 401 and attempt == 0:
            try:
                tm.refresh()
            except RuntimeError:
                break
            continue
        try:
            return resp, resp.json()
        except Exception:
            return resp, {}
    try:
        return resp, resp.json()
    except Exception:
        return resp, {}


def api_post(url: str, payload: dict, tm: TokenManager) -> tuple[requests.Response, dict]:
    """POST with one refresh attempt on 401 (see note above)."""
    resp = None
    for attempt in range(2):
        resp = requests.post(url, json=payload, headers=tm.headers(), timeout=30)
        if resp.status_code == 401 and attempt == 0:
            try:
                tm.refresh()
            except RuntimeError:
                break
            continue
        try:
            return resp, resp.json()
        except Exception:
            return resp, {}
    try:
        return resp, resp.json()
    except Exception:
        return resp, {}


# ── Normalized-error wrappers ────────────────────────────────────────────
# These wrap api_get()/api_post() so every caller (header creation, line
# creation) gets back the same (resp, resp_json, errors) shape used
# throughout Customer_load.py — errors is either None (success) or a
# normalized list of {"fieldName", "message", "fieldValue", "code"} dicts,
# ready to be handed straight to resolve_qad_errors().

def _network_error(exc: Exception) -> dict:
    return {"fieldName": None, "message": f"Network error: {exc}", "fieldValue": None, "code": None}


def _token_error() -> dict:
    return {
        "fieldName": None,
        "message": "QAD session token expired or invalid — please log in again and re-run the load.",
        "fieldValue": None,
        "code": None,
    }


def _http_error(status_code: int) -> dict:
    return {"fieldName": None, "message": f"Unexpected response — HTTP {status_code}", "fieldValue": None, "code": None}


def safe_get(url: str, tm: TokenManager) -> tuple[requests.Response | None, dict | None, list | None]:
    """GET wrapped with network-error trapping. Does NOT run normalize_response
    (GET responses here are plain data payloads, not submit envelopes) —
    mirrors get_customer()/get_mfg_customer() in Customer_load.py, which are
    also not normalized."""
    try:
        resp, resp_json = api_get(url, tm)
    except requests.RequestException as e:
        return None, None, [_network_error(e)]

    if resp.status_code == 401:
        return resp, resp_json, [_token_error()]
    if resp.status_code != 200:
        return resp, resp_json, [_http_error(resp.status_code)]
    return resp, resp_json, None


def safe_post(url: str, payload: dict, tm: TokenManager) -> tuple[requests.Response | None, dict | None, list | None]:
    """POST wrapped with network-error trapping and normalize_response(),
    exactly like post_customer()/post_mfg_customer() in Customer_load.py."""
    try:
        resp, resp_json = api_post(url, payload, tm)
    except requests.RequestException as e:
        return None, None, [_network_error(e)]

    if resp.status_code == 401:
        return resp, resp_json, [_token_error()]

    success, errors = normalize_response(resp_json, resp.status_code)
    if not success:
        if not errors:
            errors = [_http_error(resp.status_code)]
        return resp, resp_json, errors
    return resp, resp_json, None


# =============================================================================
# 3. VALUE EXTRACTORS  (unchanged)
# =============================================================================

def sv(row_data: dict, key: str, default: str = "") -> str:
    val = row_data.get(key, default)
    return str(val).strip() if val is not None else default


def fv(row_data: dict, key: str, default: float = 0.0) -> float:
    try:
        return float(row_data.get(key, default))
    except (TypeError, ValueError):
        return default


def to_iso_date(val) -> str:
    """Convert a cell value or datetime to ISO 8601 UTC string."""
    if isinstance(val, datetime):
        return val.strftime("%Y-%m-%dT%H:%M:%S.000Z")
    if val and isinstance(val, str):
        try:
            return datetime.strptime(val, "%Y-%m-%d %H:%M:%S").strftime("%Y-%m-%dT%H:%M:%S.000Z")
        except ValueError:
            return val
    return val or ""


# =============================================================================
# 4. MANDATORY FIELD CHECK
# =============================================================================

def _check_mandatory(row_data: dict, columns: list[str]) -> list[str]:
    missing = []
    for col in columns:
        v = row_data.get(col)
        if v is None or str(v).strip() == "" or str(v).strip().lower() == "none":
            missing.append(col)
    return missing


# =============================================================================
# 5. WORKBOOK HELPERS  (sheet/header layout UNCHANGED — row 1 = headers)
# =============================================================================

def ensure_columns(ws, *col_names: str) -> list:
    """Add any missing column names to row 1 and return the full header list.

    Stops at the first blank cell — real headers are always contiguous
    from column 1, so this naturally excludes blank_template.py's hidden
    dropdown-option columns further right (column 100+) from being read as
    if they were data columns. Without this, a dropdown's option list
    sharing row numbers with early data rows would make an otherwise-blank
    row look "non-empty" to the any(row_values) checks below.
    """
    header_row = []
    for cell in ws[1]:
        if cell.value is None:
            break
        header_row.append(cell.value.strip() if isinstance(cell.value, str) else cell.value)
    for name in col_names:
        if name not in header_row:
            ws.cell(row=1, column=len(header_row) + 1, value=name)
            header_row.append(name)
    return header_row


# ── File rename helpers (UNCHANGED) ─────────────────────────────────────────

def rename_error(file_path: str) -> str:
    folder, name = os.path.dirname(file_path), os.path.basename(file_path)
    if not name.startswith("error_"):
        new_path = os.path.join(folder, "error_" + name)
        os.rename(file_path, new_path)
        return new_path
    return file_path


def rename_restore(file_path: str) -> str:
    folder, name = os.path.dirname(file_path), os.path.basename(file_path)
    if name.startswith("error_"):
        new_path = os.path.join(folder, name[len("error_"):])
        os.rename(file_path, new_path)
        return new_path
    return file_path


# =============================================================================
# 6. PAYLOAD BUILDER  (UNCHANGED)
# =============================================================================

def build_header_payload(row_data: dict) -> dict:
    order_dt = to_iso_date(row_data.get("Order Date")) or to_iso_date(datetime.now())
    due_dt   = to_iso_date(row_data.get("Due Date"))   or to_iso_date(datetime.now())

    return {
        "purchaseOrderHeaders": [{
            "purchaseOrderNumber":    sv(row_data, "PO Number"),
            "domainCode":             sv(row_data, "Domain Code"),
            "supplierCode":           sv(row_data, "Supplier Code"),
            "currencyCode":           sv(row_data, "Currency"),
            "exchangeRate":           1,
            "exchangeRate2":          1,
            "exchangeRateType":       "ACCOUNTING",
            "creditTermsCode":        sv(row_data, "Credit Terms"),
            "daybookSetCode":         sv(row_data, "Daybook Set"),
            "shipToSite":             sv(row_data, "Ship To Site"),
            "billToCode":             sv(row_data, "Bill To Code"),
            "taxEnvironment":         sv(row_data, "Tax Environment", "USA"),
            "languageCode":           sv(row_data, "Language Code", "us"),
            "contact":                sv(row_data, "Contact"),
            "remarks":                sv(row_data, "Remarks"),
            "orderDate":              order_dt,
            "dueDate":                due_dt,
            "pricingDate":            order_dt,
            "lastPriceDate":          order_dt,
            "startEffective":         order_dt,
            "endEffective":           order_dt,
            "orderRevisionDate":      order_dt,
            "taxDateSummary":         order_dt,
            "orderStatus":            "O",
            "isConfirmed":            True,
            "isFixedPrice":           True,
            "isPrintPurchaseOrder":   True,
            "isTaxable":              True,
            "isDisplayTaxAmounts":    True,
            "isUsingConsignmentInventory": True,
            "isIntrastatUsed":        True,
            "isPredefaulted":         True,
            "maximumAgingDays":       90,
            "transmitFlag":           "1",
            "ERSOption":              "1",
            "ERSPriceListOption":     "0",
            "dataOperation":          "C",
            "concurrencyHash":        "",
            "disallowedActions":      "",
            "disallowedActionsMessage": "",
            "supplementaryMessages":  [],
            "POIntrastats":           [],
        }]
    }


# =============================================================================
# 7. LINE CREATION FLOW  (API sequencing UNCHANGED — only response/error
#    handling now goes through safe_get()/safe_post() + normalize_response())
# =============================================================================

def create_line(
    domain:        str,
    po_num:        str,
    line_row_data: dict,
    tm:            TokenManager,
) -> tuple[bool, list, int]:
    """
    Full multi-step line creation:
      init → fieldChange(siteCode) → fieldChange(itemCode) →
      isReceived check → fieldChange(quantityOrdered) →
      fieldChange(purchaseCost) → sync/commit

    Line number is assigned by QAD in the init response — we never override it.
    Returns (success, errors, line_number), where errors is a normalized
    list of {"fieldName", "message", "fieldValue", "code"} dicts (empty on
    success) ready for resolve_qad_errors().
    """
    site_code = sv(line_row_data, "Site Code")
    item_code = sv(line_row_data, "Item Code")
    qty       = fv(line_row_data, "Quantity Ordered")
    price     = fv(line_row_data, "Unit Price")
    due_dt    = to_iso_date(line_row_data.get("Due Date")) or to_iso_date(datetime.now())

    # 1. Initialise blank line — QAD returns the next available line number
    print("    ↳ Init line (QAD will assign number)...")
    resp, resp_json, errors = safe_get(_init_line_url(tm.base_url, domain, po_num), tm)
    if errors:
        return False, errors, 0

    lines = (resp_json or {}).get("data", {}).get("purchaseOrderLines", [])
    if not lines:
        return False, [{"fieldName": None, "message": "Init returned no line object", "fieldValue": None, "code": None}], 0

    line        = lines[0]
    line_number = line["purchaseOrderLine"]   # QAD-assigned — do not override
    line["dueDate"] = due_dt

    # 2. fieldChange: siteCode
    line["siteCode"] = site_code
    resp, resp_json, errors = safe_post(
        _field_change_url(tm.base_url, "siteCode"),
        {"purchaseOrderLines": [line]}, tm,
    )
    if errors:
        return False, errors, line_number
    lines = (resp_json or {}).get("data", {}).get("purchaseOrderLines", [])
    if not lines:
        return False, [{"fieldName": "siteCode", "message": "fieldChange(siteCode) returned no line", "fieldValue": site_code, "code": None}], line_number
    line = lines[0]

    # 3. fieldChange: itemCode
    line["itemCode"] = item_code
    resp, resp_json, errors = safe_post(
        _field_change_url(tm.base_url, "itemCode"),
        {"purchaseOrderLines": [line]}, tm,
    )
    if errors:
        return False, errors, line_number
    lines = (resp_json or {}).get("data", {}).get("purchaseOrderLines", [])
    if not lines:
        return False, [{"fieldName": "itemCode", "message": "fieldChange(itemCode) returned no line", "fieldValue": item_code, "code": None}], line_number
    line = lines[0]

    time.sleep(0.1)

    # 4. isReceived check (fire-and-forget) — use QAD-assigned line number
    safe_get(_is_received_url(tm.base_url, domain, po_num, line_number), tm)

    time.sleep(0.1)

    # 5. fieldChange: quantityOrdered
    line["quantityOrdered"] = qty
    resp, resp_json, errors = safe_post(
        _field_change_url(tm.base_url, "quantityOrdered"),
        {"purchaseOrderLines": [line]}, tm,
    )
    if errors:
        return False, errors, line_number
    lines = (resp_json or {}).get("data", {}).get("purchaseOrderLines", [])
    if not lines:
        return False, [{"fieldName": "quantityOrdered", "message": "fieldChange(quantityOrdered) returned no line", "fieldValue": qty, "code": None}], line_number
    line = lines[0]

    # 6. fieldChange: purchaseCost
    line["purchaseCost"] = price
    resp, resp_json, errors = safe_post(
        _field_change_url(tm.base_url, "purchaseCost"),
        {"purchaseOrderLines": [line]}, tm,
    )
    if errors:
        return False, errors, line_number
    lines = (resp_json or {}).get("data", {}).get("purchaseOrderLines", [])
    if not lines:
        return False, [{"fieldName": "purchaseCost", "message": "fieldChange(purchaseCost) returned no line", "fieldValue": price, "code": None}], line_number
    line = lines[0]

    # 7. Sync / commit line
    resp, resp_json, errors = safe_post(_sync_line_url(tm.base_url), {"purchaseOrderLines": [line]}, tm)
    if errors:
        return False, errors, line_number

    return True, [], line_number


# =============================================================================
# 8. SINGLE-FILE PROCESSOR
# =============================================================================

def process_file(file_path: str, tm: TokenManager) -> tuple[int, int]:
    print(f"\n{'─'*55}")
    print(f"  Loading: {os.path.basename(file_path)}")
    print(f"{'─'*55}")

    wb = openpyxl.load_workbook(file_path)

    if "Header" not in wb.sheetnames or "Lines" not in wb.sheetnames:
        print("ERROR: Workbook must have sheets named 'Header' and 'Lines'")
        raise RuntimeError("Workbook must have sheets named 'Header' and 'Lines'")

    ws_h = wb["Header"]
    ws_l = wb["Lines"]

    h_headers = ensure_columns(ws_h, "Status", "Error")
    l_headers = ensure_columns(ws_l, "Status", "Error")

    h_status_col = h_headers.index("Status") + 1
    h_error_col  = h_headers.index("Error")  + 1
    l_status_col = l_headers.index("Status") + 1
    l_error_col  = l_headers.index("Error")  + 1

    # ── Index header rows by PO number ────────────────────────────────────
    header_rows: dict[str, tuple[int, dict]] = {}
    for row_idx, row in enumerate(ws_h.iter_rows(min_row=2), start=2):
        row_values = [cell.value for cell in row]
        if not any(row_values[:len(h_headers)]):
            continue
        row_data = dict(zip(h_headers, row_values))
        po_num   = sv(row_data, "PO Number")
        if po_num:
            header_rows[po_num] = (row_idx, row_data)

    # ── Index line rows by PO number ──────────────────────────────────────
    line_rows: dict[str, list[tuple[int, dict]]] = defaultdict(list)
    for row_idx, row in enumerate(ws_l.iter_rows(min_row=2), start=2):
        row_values = [cell.value for cell in row]
        if not any(row_values[:len(l_headers)]):
            continue
        row_data = dict(zip(l_headers, row_values))
        po_num   = sv(row_data, "PO Number")
        if po_num:
            line_rows[po_num].append((row_idx, row_data))

    h_success = h_skip = h_fail = 0
    l_success = l_skip = l_fail = 0

    total_pos = len(header_rows)

    for idx, (po_num, (h_row_idx, h_row_data)) in enumerate(header_rows.items(), start=1):
        if _progress_callback:
            _progress_callback(idx, total_pos)

        print(f"\n  PO: {po_num}")

        # ── Step 1: Create header ─────────────────────────────────────────
        if str(h_row_data.get("Status", "")).strip().upper() == "DONE":
            print("  Header — already DONE, skipping")
            h_skip += 1
        else:
            # Mandatory check before any API call
            missing = _check_mandatory(h_row_data, MANDATORY_HEADER_COLUMNS)
            if missing:
                h_fail += 1
                mark_error(
                    ws_h, h_row_idx, h_status_col, h_error_col, h_headers,
                    missing,
                    f"Missing mandatory fields: {', '.join(missing)}",
                )
                wb.save(file_path)
                continue

            payload = build_header_payload(h_row_data)
            resp, resp_json, errors = safe_post(_header_url(tm.base_url), payload, tm)

            if errors:
                h_fail += 1
                bad_cols, error_msg = resolve_qad_errors(errors, PO_HEADER_FIELD_TO_COLUMN)
                mark_error(ws_h, h_row_idx, h_status_col, h_error_col, h_headers, bad_cols, error_msg)
                print(f"  Header FAILED: {po_num} — {error_msg}")
                wb.save(file_path)
                continue

            mark_done(ws_h, h_row_idx, h_status_col, h_error_col)
            h_success += 1
            print(f"  Header OK: {po_num}")
            wb.save(file_path)
            time.sleep(0.3)

        # ── Step 2: Create lines ──────────────────────────────────────────
        domain   = sv(h_row_data, "Domain Code")
        po_lines = line_rows.get(po_num, [])

        if not po_lines:
            print(f"  WARNING: No lines found for PO: {po_num}")
            continue

        for seq, (l_row_idx, l_row_data) in enumerate(po_lines, start=1):
            if str(l_row_data.get("Status", "")).strip().upper() == "DONE":
                print(f"    Line {seq} — already DONE, skipping")
                l_skip += 1
                continue

            missing = _check_mandatory(l_row_data, MANDATORY_LINE_COLUMNS)
            if missing:
                l_fail += 1
                mark_error(
                    ws_l, l_row_idx, l_status_col, l_error_col, l_headers,
                    missing,
                    f"Missing mandatory fields: {', '.join(missing)}",
                )
                wb.save(file_path)
                continue

            item  = sv(l_row_data, "Item Code")
            qty   = fv(l_row_data, "Quantity Ordered")
            price = fv(l_row_data, "Unit Price")
            print(f"    → {item} | qty={qty} | price={price}")

            success, errors, line_no = create_line(domain, po_num, l_row_data, tm)

            if success:
                mark_done(ws_l, l_row_idx, l_status_col, l_error_col)
                l_success += 1
                print(f"    Line {line_no} OK: {item}")
            else:
                l_fail += 1
                bad_cols, error_msg = resolve_qad_errors(errors, PO_LINE_FIELD_TO_COLUMN)
                mark_error(ws_l, l_row_idx, l_status_col, l_error_col, l_headers, bad_cols, error_msg)
                print(f"    Line {line_no} FAILED: {item} — {error_msg}")

            wb.save(file_path)
            time.sleep(0.2)

    total_success = h_success + l_success
    total_fail    = h_fail    + l_fail

    print(f"\n  Headers — OK {h_success} | Skipped {h_skip} | Failed {h_fail}")
    print(f"  Lines   — OK {l_success} | Skipped {l_skip} | Failed {l_fail}")

    return total_success, total_fail


# =============================================================================
# 9. ORCHESTRATOR
# =============================================================================
#
# CHANGED: run() used to build its own TokenManager() with zero arguments
# (self-authenticating from config). It now takes an already-built
# TokenManager (token + base_url) as a parameter, same as Customer_load.py —
# constructed once by the caller (main.py per-session, or the __main__
# block below for standalone CLI use) and reused across every file in the
# folder.

def run(folder_path: str, tm: TokenManager) -> tuple[int, int]:
    folder = os.path.abspath(folder_path)

    if not os.path.exists(folder):
        print(f"ERROR: Folder not found: {folder}")
        raise RuntimeError(f"Folder not found: {folder}")

    xlsx_files = [
        f for f in os.listdir(folder)
        if f.endswith(".xlsx") and not f.startswith("~$")
    ]

    if not xlsx_files:
        print(f"ERROR: No .xlsx files found in: {folder}")
        raise RuntimeError(f"No .xlsx files found in: {folder}")

    total_success = 0
    total_fail    = 0

    for file_name in xlsx_files:
        file_path = os.path.join(folder, file_name)

        s, fail = process_file(file_path, tm)
        total_success += s
        total_fail    += fail

        if fail > 0:
            file_path = rename_error(file_path)
            print(f"\n  Errors found — renamed to: {os.path.basename(file_path)}")
        elif s > 0:
            file_path = rename_restore(file_path)

        print(f"\n  Load Summary {'-'*30}")
        print(f"    Rows loaded : {s}")
        print(f"    Rows failed : {fail}")
        if fail == 0:
            print("    Result      : ALL ROWS LOADED SUCCESSFULLY ✓")
        else:
            print("    Result      : COMPLETED WITH ERRORS — fix red rows and re-run")

    print(f"\n{'='*55}")
    print(f"  TOTAL — Success: {total_success} | Failed: {total_fail}")
    print(f"{'='*55}")

    return total_success, total_fail


# =============================================================================
# 10. ENTRY POINT
# =============================================================================
#
# CHANGED: standalone CLI use now performs the same interactive OAuth login
# Customer_load.py does (config.json has no stored service-account creds
# to fetch a token with anymore). Set QAD_ENVIRONMENT / QAD_USERNAME /
# QAD_PASSWORD to skip the prompts.

if __name__ == "__main__":
    import getpass
    import requests as _requests

    environment = os.environ.get("QAD_ENVIRONMENT", "TEST").upper()
    env_cfg     = CONFIG["environments"][environment]

    username = os.environ.get("QAD_USERNAME") or input("QAD username: ")
    password = os.environ.get("QAD_PASSWORD") or getpass.getpass("QAD password: ")

    token_resp = _requests.post(
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

    folder = os.path.abspath(
        os.path.join(ROOT_DIR, CONFIG["folders"]["purchase_order"])
    )
    ok, fail = run(folder, tm)
    sys.exit(0 if fail == 0 else 1)