"""
PO_load.py
----------
Loads Purchase Order headers and lines from all .xlsx files in the
configured PurchaseOrder folder into QAD via the purchaseOrders API.

Behaviour
---------
- Reads base_url and auth from config.json (via config.py)
- Fetches one OAuth token per run; refreshes automatically on 401
- Skips rows with Status = DONE
- On success  → clears all fills, sets Status = DONE
- On failure  → red-fills col 1 + Error cell, sets Status = ERROR
- Saves workbook after every row so progress survives a crash
- Renames file to error_<name> if any failures occurred
- Returns (total_success, total_fail)

Sheet layout
------------
Workbook must contain two sheets: 'Header' and 'Lines'.
Each sheet must have a header row (row 1) with column names.
"""

import os
import sys
import time
import requests
import openpyxl
from openpyxl.styles import PatternFill
from datetime import datetime
from collections import defaultdict

# ── Config ────────────────────────────────────────────────────────────────
ROOT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, ROOT_DIR)
from config import CONFIG

_BASE_URL = CONFIG["qad"]["base_url"]

# ── API endpoint templates ─────────────────────────────────────────────────
HEADER_URL       = f"{_BASE_URL}/api/erp/purchaseOrderHeaders?viewUri=urn:be:com.qad.purchasing.purchaseorders.IPurchaseOrderHeader"
INIT_LINE_URL    = f"{_BASE_URL}/api/erp/purchaseOrderLinesGrid?initialize=true&domainCode={{domain}}&purchaseOrderNumber={{po}}"
FIELD_CHANGE_URL = f"{_BASE_URL}/api/erp/purchaseOrderLines/fieldChangeV2?fieldName={{fieldName}}"
IS_RECEIVED_URL  = f"{_BASE_URL}/api/erp/purchaseOrderLines/isReceivedPurchaseOrderLineV2?domainCode={{domain}}&purchaseOrderNumber={{po}}&purchaseOrderLine={{line}}"
SYNC_LINE_URL    = f"{_BASE_URL}/api/erp/purchaseOrderLinesGrid"

# ── Fill constants ────────────────────────────────────────────────────────
RED_FILL   = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
CLEAR_FILL = PatternFill(fill_type=None)

# ── Date formats accepted in the spreadsheet ──────────────────────────────
DATE_FORMATS = ["%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y"]


# =============================================================================
# 1. AUTHENTICATION
# =============================================================================

class _TokenExpired(Exception):
    pass


def _fetch_token() -> str:
    token_url = f"{_BASE_URL}/oauth/token"
    resp = requests.post(token_url, params=CONFIG["qad"]["auth"], timeout=30)
    resp.raise_for_status()
    token = resp.json().get("access_token")
    if not token:
        raise RuntimeError("OAuth response did not contain access_token")
    return token


class TokenManager:
    """Holds one token for the run; refreshes on demand (401)."""

    def __init__(self):
        self._token: str | None = None

    def get(self) -> str:
        if self._token is None:
            self._token = _fetch_token()
        return self._token

    def refresh(self) -> str:
        self._token = _fetch_token()
        return self._token

    def headers(self) -> dict:
        return {
            "Content-Type":  "application/json",
            "Authorization": f"Bearer {self.get()}",
        }


# =============================================================================
# 2. HTTP HELPERS  (auto-refresh on 401, one retry)
# =============================================================================

def api_get(url: str, tm: TokenManager) -> tuple[requests.Response, dict]:
    """GET with one automatic token-refresh retry on 401."""
    for attempt in range(2):
        resp = requests.get(url, headers=tm.headers(), timeout=30)
        if resp.status_code == 401 and attempt == 0:
            tm.refresh()
            continue
        try:
            return resp, resp.json()
        except Exception:
            return resp, {}
    return resp, {}


def api_post(url: str, payload: dict, tm: TokenManager) -> tuple[requests.Response, dict]:
    """POST with one automatic token-refresh retry on 401."""
    for attempt in range(2):
        resp = requests.post(url, json=payload, headers=tm.headers(), timeout=30)
        if resp.status_code == 401 and attempt == 0:
            tm.refresh()
            continue
        try:
            return resp, resp.json()
        except Exception:
            return resp, {}
    return resp, {}


# =============================================================================
# 3. VALUE EXTRACTORS
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
# 4. WORKBOOK HELPERS
# =============================================================================

def ensure_columns(ws, *col_names: str) -> list:
    """Add any missing column names to row 1 and return the full header list."""
    header_row = [
        cell.value.strip() if isinstance(cell.value, str) else cell.value
        for cell in ws[1]
    ]
    for name in col_names:
        if name not in header_row:
            ws.cell(row=1, column=len(header_row) + 1, value=name)
            header_row.append(name)
    return header_row


def mark_row_success(ws, row_idx: int, status_col: int, error_col: int):
    for cell in ws[row_idx]:
        cell.fill = CLEAR_FILL
    ws.cell(row=row_idx, column=status_col, value="DONE")
    ws.cell(row=row_idx, column=error_col,  value="")


def mark_row_error(ws, row_idx: int, status_col: int, error_col: int, error_msg: str):
    """Red-fill col 1 and the Error cell; clear everything else."""
    for cell in ws[row_idx]:
        cell.fill = CLEAR_FILL
    ws.cell(row=row_idx, column=1).fill = RED_FILL
    ws.cell(row=row_idx, column=status_col, value="ERROR")
    ws.cell(row=row_idx, column=error_col,  value=error_msg)
    ws.cell(row=row_idx, column=error_col).fill = RED_FILL


# =============================================================================
# 5. FILE RENAME HELPERS
# =============================================================================

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
# 6. API ERROR PARSER
# =============================================================================

def parse_api_errors(resp_json: dict, entity_name: str = "") -> str:
    errors = resp_json.get("submitResult", {}).get("errors", [])
    msgs = []
    for e in errors:
        msg = e.get("message", "")
        if "already exists" in msg.lower():
            msgs.append(f"Duplicate — '{entity_name}' already exists")
        elif e.get("fieldName"):
            msgs.append(f"Field '{e['fieldName']}': {msg}")
        else:
            msgs.append(msg)
    return "; ".join(msgs) or resp_json.get("message", "Unknown error")


# =============================================================================
# 7. PAYLOAD BUILDER
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
# 8. LINE CREATION FLOW  (unchanged logic)
# =============================================================================

def create_line(
    domain:        str,
    po_num:        str,
    line_row_data: dict,
    tm:            TokenManager,
) -> tuple[bool, str, int]:
    """
    Full multi-step line creation:
      init → fieldChange(siteCode) → fieldChange(itemCode) →
      isReceived check → fieldChange(quantityOrdered) →
      fieldChange(purchaseCost) → sync/commit

    Line number is assigned by QAD in the init response — we never override it.
    Returns (success, error_msg, line_number).
    """
    site_code = sv(line_row_data, "Site Code")
    item_code = sv(line_row_data, "Item Code")
    qty       = fv(line_row_data, "Quantity Ordered")
    price     = fv(line_row_data, "Unit Price")
    due_dt    = to_iso_date(line_row_data.get("Due Date")) or to_iso_date(datetime.now())

    # 1. Initialise blank line — QAD returns the next available line number
    print(f"    ↳ Init line (QAD will assign number)...")
    resp, resp_json = api_get(
        INIT_LINE_URL.format(domain=domain, po=po_num), tm
    )
    if resp.status_code != 200:
        return False, f"Init failed: HTTP {resp.status_code}", 0

    lines = resp_json.get("data", {}).get("purchaseOrderLines", [])
    if not lines:
        return False, "Init returned no line object", 0

    line        = lines[0]
    line_number = line["purchaseOrderLine"]   # QAD-assigned — do not override
    line["dueDate"] = due_dt

    # 2. fieldChange: siteCode
    line["siteCode"] = site_code
    resp, resp_json = api_post(
        FIELD_CHANGE_URL.format(fieldName="siteCode"),
        {"purchaseOrderLines": [line]}, tm,
    )
    if resp.status_code != 200:
        return False, f"fieldChange(siteCode) failed: HTTP {resp.status_code}", line_number
    lines = resp_json.get("data", {}).get("purchaseOrderLines", [])
    if not lines:
        return False, "fieldChange(siteCode) returned no line", line_number
    line = lines[0]

    # 3. fieldChange: itemCode
    line["itemCode"] = item_code
    resp, resp_json = api_post(
        FIELD_CHANGE_URL.format(fieldName="itemCode"),
        {"purchaseOrderLines": [line]}, tm,
    )
    if resp.status_code != 200:
        return False, f"fieldChange(itemCode) failed: HTTP {resp.status_code}", line_number
    lines = resp_json.get("data", {}).get("purchaseOrderLines", [])
    if not lines:
        return False, "fieldChange(itemCode) returned no line", line_number
    line = lines[0]

    time.sleep(0.1)

    # 4. isReceived check (fire-and-forget) — use QAD-assigned line number
    api_get(IS_RECEIVED_URL.format(domain=domain, po=po_num, line=line_number), tm)

    time.sleep(0.1)

    # 5. fieldChange: quantityOrdered
    line["quantityOrdered"] = qty
    resp, resp_json = api_post(
        FIELD_CHANGE_URL.format(fieldName="quantityOrdered"),
        {"purchaseOrderLines": [line]}, tm,
    )
    if resp.status_code != 200:
        return False, f"fieldChange(quantityOrdered) failed: HTTP {resp.status_code}", line_number
    lines = resp_json.get("data", {}).get("purchaseOrderLines", [])
    if not lines:
        return False, "fieldChange(quantityOrdered) returned no line", line_number
    line = lines[0]

    # 6. fieldChange: purchaseCost
    line["purchaseCost"] = price
    resp, resp_json = api_post(
        FIELD_CHANGE_URL.format(fieldName="purchaseCost"),
        {"purchaseOrderLines": [line]}, tm,
    )
    if resp.status_code != 200:
        return False, f"fieldChange(purchaseCost) failed: HTTP {resp.status_code}", line_number
    lines = resp_json.get("data", {}).get("purchaseOrderLines", [])
    if not lines:
        return False, "fieldChange(purchaseCost) returned no line", line_number
    line = lines[0]

    # 7. Sync / commit line
    resp, resp_json = api_post(SYNC_LINE_URL, {"purchaseOrderLines": [line]}, tm)
    if resp.status_code == 200 and resp_json.get("submitResult", {}).get("success"):
        return True, "", line_number

    error_msg = parse_api_errors(resp_json, f"PO {po_num} Line {line_number}")
    return False, f"Sync failed: {error_msg}", line_number


# =============================================================================
# 9. SINGLE-FILE PROCESSOR
# =============================================================================

def process_file(file_path: str, tm: TokenManager) -> tuple[int, int]:
    print(f"\n{'─'*55}")
    print(f"  📄 {os.path.basename(file_path)}")
    print(f"{'─'*55}")

    wb = openpyxl.load_workbook(file_path)

    if "Header" not in wb.sheetnames or "Lines" not in wb.sheetnames:
        print("❌ Workbook must have sheets named 'Header' and 'Lines'")
        return 0, 1

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
        if not any(row_values):
            continue
        row_data = dict(zip(h_headers, row_values))
        po_num   = sv(row_data, "PO Number")
        if po_num:
            header_rows[po_num] = (row_idx, row_data)

    # ── Index line rows by PO number ──────────────────────────────────────
    line_rows: dict[str, list[tuple[int, dict]]] = defaultdict(list)
    for row_idx, row in enumerate(ws_l.iter_rows(min_row=2), start=2):
        row_values = [cell.value for cell in row]
        if not any(row_values):
            continue
        row_data = dict(zip(l_headers, row_values))
        po_num   = sv(row_data, "PO Number")
        if po_num:
            line_rows[po_num].append((row_idx, row_data))

    h_success = h_skip = h_fail = 0
    l_success = l_skip = l_fail = 0

    for po_num, (h_row_idx, h_row_data) in header_rows.items():
        print(f"\n  📦 PO: {po_num}")

        # ── Step 1: Create header ─────────────────────────────────────────
        if str(h_row_data.get("Status", "")).strip().upper() == "DONE":
            print("  ⏭  Header — already DONE")
            h_skip += 1
        else:
            try:
                payload = build_header_payload(h_row_data)
                resp, resp_json = api_post(HEADER_URL, payload, tm)

                if resp.status_code == 200 and resp_json.get("submitResult", {}).get("success"):
                    mark_row_success(ws_h, h_row_idx, h_status_col, h_error_col)
                    h_success += 1
                    print(f"  ✔  Header: {po_num}")
                else:
                    error_msg = parse_api_errors(resp_json, po_num)
                    mark_row_error(ws_h, h_row_idx, h_status_col, h_error_col, error_msg)
                    h_fail += 1
                    print(f"  ✘  Header: {po_num} — {error_msg}")
                    wb.save(file_path)
                    continue

            except Exception as e:
                mark_row_error(ws_h, h_row_idx, h_status_col, h_error_col, str(e))
                h_fail += 1
                print(f"  ✘  Header: {po_num} — {e}")
                wb.save(file_path)
                continue

            time.sleep(0.3)

        # ── Step 2: Create lines ──────────────────────────────────────────
        domain   = sv(h_row_data, "Domain Code")
        po_lines = line_rows.get(po_num, [])

        if not po_lines:
            print(f"  ⚠  No lines found for PO: {po_num}")
            continue

        for seq, (l_row_idx, l_row_data) in enumerate(po_lines, start=1):
            if str(l_row_data.get("Status", "")).strip().upper() == "DONE":
                print(f"    ⏭  Line {seq} — already DONE")
                l_skip += 1
                continue

            item  = sv(l_row_data, "Item Code")
            qty   = fv(l_row_data, "Quantity Ordered")
            price = fv(l_row_data, "Unit Price")
            print(f"    → {item} | qty={qty} | price={price}")

            try:
                success, error_msg, line_no = create_line(domain, po_num, l_row_data, tm)

                if success:
                    mark_row_success(ws_l, l_row_idx, l_status_col, l_error_col)
                    l_success += 1
                    print(f"    ✔  Line {line_no}: {item}")
                else:
                    mark_row_error(ws_l, l_row_idx, l_status_col, l_error_col, error_msg)
                    l_fail += 1
                    print(f"    ✘  Line {line_no}: {item} — {error_msg}")

            except Exception as e:
                mark_row_error(ws_l, l_row_idx, l_status_col, l_error_col, str(e))
                l_fail += 1
                print(f"    ✘  {item} — {e}")

            wb.save(file_path)
            time.sleep(0.2)

    total_success = h_success + l_success
    total_fail    = h_fail    + l_fail

    print(f"\n  📊 Headers — ✔ {h_success} | ⏭ {h_skip} | ✘ {h_fail}")
    print(f"  📊 Lines   — ✔ {l_success} | ⏭ {l_skip} | ✘ {l_fail}")

    return total_success, total_fail


# =============================================================================
# 10. ORCHESTRATOR
# =============================================================================

def run(folder_path: str) -> tuple[int, int]:
    folder = os.path.abspath(folder_path)

    if not os.path.exists(folder):
        print(f"❌ Folder not found: {folder}")
        raise RuntimeError(f"Folder not found: {folder}")

    xlsx_files = [
        f for f in os.listdir(folder)
        if f.endswith(".xlsx") and not f.startswith("~$")
    ]

    if not xlsx_files:
        print(f"⚠  No .xlsx files found in: {folder}")
        raise RuntimeError(f"No .xlsx files found in: {folder}")

    tm = TokenManager()

    total_success = 0
    total_fail    = 0

    for file_name in xlsx_files:
        file_path = os.path.join(folder, file_name)
        s, fail = process_file(file_path, tm)
        total_success += s
        total_fail    += fail

        if fail > 0:
            file_path = rename_error(file_path)
            print(f"\n  ⚠  Errors found — renamed to: {os.path.basename(file_path)}")
        elif s > 0:
            file_path = rename_restore(file_path)

        print(f"\n  Load Summary {'─'*30}")
        print(f"    Rows loaded : {s}")
        print(f"    Rows failed : {fail}")
        print("    Result      :", "ALL ROWS LOADED SUCCESSFULLY ✓" if fail == 0 else "COMPLETED WITH ERRORS — fix red rows and re-run")

    print(f"\n{'═'*55}")
    print(f"  TOTAL — Success: {total_success} | Failed: {total_fail}")
    print(f"{'═'*55}")

    return total_success, total_fail


# =============================================================================
# 11. ENTRY POINT
# =============================================================================

if __name__ == "__main__":
    folder = os.path.abspath(
        os.path.join(ROOT_DIR, CONFIG["folders"]["purchase_order"])
    )
    ok, fail = run(folder)
    sys.exit(0 if fail == 0 else 1)