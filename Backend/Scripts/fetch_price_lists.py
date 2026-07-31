"""
fetch_price_lists.py
--------------------
Fetches QAD Price List (priceListV2) records into an Excel workbook whose
layout is IDENTICAL to the one price_list_load.py consumes, so the workflow
is: Fetch -> Edit -> Validate -> Load, with no restructuring in between.

SINGLE-STAGE FETCH (Browse only -- Detail GET removed 2026-07-24)
-------------------------------------------------------------------
Earlier versions of this script assumed Browse only returned "header"
fields and that pricing/detail fields (Amount Type, Quantity Type,
Combination Type, Minimum Order, Maximum Quantity, Maximum orders, Break
Category, dates, description) required a second per-row Detail GET call.

A browser capture against the real Browse endpoint
(urn:browse:bebrowse:com.qad.erp.sales.priceListV2s) showed all of those
fields are already present directly on each Browse row. So:

  - Detail GET is no longer called at all. One paginated Browse pass is
    the entire fetch.
  - "Minimum Order" is now sourced from Browse's `minNetOrd` field
    (CONFIRMED 2026-07-24 as equivalent to the old detail-only
    `minimumQuantity` field).
  - `priceListId` is intentionally NOT fetched/written anywhere. It is
    QAD-internal and price_list_load.py's UPDATE flow re-fetches it fresh
    via its own GET at load time -- fetch has no use for it.

EXCEL LAYOUT (must match price_list_load.py's loader exactly)
---------------------------------------------------------------
Row 1 (section labels): Main (A:H) | Pricing (I:O)
Row 2 (headers):
  A Domain                 B Price List          C Description
  D Customer Code          E Item Code            F Currency
  G Unit Of Measure        H Start Date           I Expire Data
  J Amount Type            K Quantity Type        L Combination Type
  M Minimum Order          N Maximum Quantity     O Maximum orders
  P Break category
  Q Data Operation         R Status               S Error

FETCHED-ROW HIGHLIGHTING (replaces the old cell-locking approach)
---------------------------------------------------------------
Earlier versions of this script locked the LOCKED_ON_UPDATE columns
(Domain, Price List, Customer Code, Item Code, Currency, Unit Of
Measure, Start Date) via cell-level Protection + sheet protection, so
users couldn't accidentally edit key fields on existing records.

That approach doesn't work now that new rows are typed directly into
the same fetched workbook -- sheet protection gets in the way of adding
rows. So locking has been removed entirely. Instead, every row written
by the fetch is filled light purple (FILL_FETCHED) so it's visually
obvious which rows came from Browse vs. which rows the user typed in
themselves (those stay unfilled/white). All data-cell text is plain
black, and every cell (header + data) gets a thin all-around border.

LOCKED_ON_UPDATE is kept only as a reference of which fields
price_list_load.py itself treats as identity fields -- it is no longer
used to lock anything in this script.
"""

import os
import sys
import uuid
from datetime import datetime
from typing import Any, Iterator

import requests
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

ROOT_DIR    = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT_DIR)
sys.path.insert(0, SCRIPTS_DIR)
from config import CONFIG  # noqa: E402

BROWSE_ID = "urn:browse:bebrowse:com.qad.erp.sales.priceListV2s"

# ── Progress callback (same convention as Customer_load.py) ───────────────
_progress_callback = None


def set_progress_callback(callback):
    global _progress_callback
    _progress_callback = callback


# ── Visible header layout (Row 2) ──────────────────────────────────────────
MAIN_HEADERS = [
    "Domain", "Price List", "Description", "Customer Code", "Item Code",
    "Currency", "Unit Of Measure", "Start Date", "Expire Date",
]
PRICING_HEADERS = [
    "Amount Type", "Quantity Type", "Combination Type",
    "Minimum Order", "Maximum Quantity", "Maximum orders", "Break category",
]
LIST_PRICE = [
    "Maximum Price", "Minimum Price", "List Price",
]   
LOADER_HEADERS = ["Data Operation", "Status ", "Error"]

ALL_HEADERS = MAIN_HEADERS + PRICING_HEADERS + LIST_PRICE + LOADER_HEADERS

# Reference only -- price_list_load.py's own identity/locked fields.
# No longer used to lock cells in this script (see module docstring).

# =============================================================================
# VALUE MAPS -- business label (shown in UI) <-> QAD numeric code (on the wire)
# =============================================================================
AMOUNT_TYPE_MAP = {
    "List Price":    "1",
    "Discount %":    "2",
    "Discount Amt":  "9",
    "Freight Terms": "6",
    "Freight List":  "7",
    "Credit Terms":  "8",
}
AMOUNT_TYPE_REVERSE_MAP = {v: k for k, v in AMOUNT_TYPE_MAP.items()}

QUANTITY_TYPE_MAP = {
    "Quantity": "1",
    "Amount":   "2",
}
QUANTITY_TYPE_REVERSE_MAP = {v: k for k, v in QUANTITY_TYPE_MAP.items()}

COMBINATION_TYPE_MAP = {
    "Combinable":      "2",
    "Base-Combinable": "3",
    "Exclusive":       "4",
}
COMBINATION_TYPE_REVERSE_MAP = {v: k for k, v in COMBINATION_TYPE_MAP.items()}


# =============================================================================
# 1. AUTH  (identical pattern to Customer_load.TokenManager)
# =============================================================================

def _fetch_token() -> str:
    url = f"{CONFIG['qad']['base_url']}/oauth/token"
    resp = requests.post(url, data=CONFIG["qad"]["auth"], timeout=30)
    resp.raise_for_status()
    token = resp.json().get("access_token")
    if not token:
        raise RuntimeError("OAuth response did not contain access_token")
    return token


class TokenManager:
    def __init__(self):
        self._token: str | None = None

    def get(self) -> str:
        if self._token is None:
            self._token = _fetch_token()
        return self._token

    def refresh(self) -> str:
        self._token = _fetch_token()
        return self._token


class _TokenExpired(Exception):
    pass


# =============================================================================
# 2. FILTER MODEL  (maps 1:1 onto the popup UI: Field | Operator | Value1 | Value2)
# =============================================================================

# CONFIRMED 2026-07-24 via browser capture:
#   eq = equals              cn = contains
#   sw = starts with         ew = ends with
#   gt = greater than        ge = greater than or equal
#   lt = less than           le = less than or equal
#   rg = range (NATIVE -- one fragment, two values: field,rg,from,literal,to,literal)
# Earlier versions of this file assumed "rg" wasn't real and expanded Range
# into two ge/le fragments. That assumption was wrong -- rg is a real,
# single-fragment QAD condition. If only one bound of a range is supplied
# (from-only or to-only), we fall back to ge/le since rg with a missing
# bound was never captured/confirmed.
VALID_CONDITIONS = {"eq", "sw", "ew", "cn", "gt", "ge", "lt", "le", "rg"}


class FilterCriteria:
    def __init__(self, field: str, condition: str, value: str, value_to: str | None = None):
        self.field = field
        self.condition = condition.lower()
        self.value = value
        self.value_to = value_to
        if self.condition not in VALID_CONDITIONS:
            raise ValueError(f"Unsupported filter condition: {self.condition}")

    def to_query_fragment(self) -> str:
        if self.condition == "rg":
            return f"{self.field},rg,{self.value},literal,{self.value_to},literal"
        return f"{self.field},{self.condition},{self.value},literal"


def build_range_criteria(field: str, value_from: str, value_to: str) -> list["FilterCriteria"]:
    """
    Both bounds present -> single native `rg` fragment (CONFIRMED format).
    Only one bound present -> fall back to a single ge/le fragment, since
    a partial rg was never captured/confirmed against the real API.
    """
    if value_from and value_to:
        return [FilterCriteria(field, "rg", value_from, value_to)]
    if value_from:
        return [FilterCriteria(field, "ge", value_from)]
    if value_to:
        return [FilterCriteria(field, "le", value_to)]
    return []


# =============================================================================
# 3. BROWSE  (search + pagination)
# =============================================================================
# CONFIRMED 2026-07-24 via browser capture:
#   - Pagination uses page / pageSize / pageAction (NOT offset/limit).
#     First page: pageAction="first". Subsequent-page value was NOT
#     captured -- "next" is an ASSUMPTION based on common QAD browse
#     convention. FLAG FOR VERIFICATION with a real page-2 capture before
#     relying on multi-page fetches in production.
#   - Multiple filters are sent as REPEATED `filter` query params, not
#     ";"-joined into one param.
#   - A callId (uuid) is sent and appears to identify one browse session
#     across pages -- generated once per fetch and reused.
#   - Response shape is {"data": [...], "auxillaryData": {...}}, where
#     each row uses FLAT DOTTED KEYS like "priceListV2.priceListCode",
#     not nested objects.
#   - auxillaryData.HasMoreParams ("true"/"false" as a STRING) indicates
#     whether another page should be requested.

def call_browse_api(
    filters: list[FilterCriteria],
    token: str,
    page: int,
    page_size: int,
    page_action: str,
    call_id: str,
) -> dict:
    url = f"{CONFIG['qad']['base_url']}/api/qracore/browses"
    params: list[tuple[str, str]] = [
        ("browseId", BROWSE_ID),
        ("callId", call_id),
        ("page", str(page)),
        ("pageSize", str(page_size)),
        ("pageAction", page_action),
    ]
    for f in filters:
        params.append(("filter", f.to_query_fragment()))

    resp = requests.get(
        url,
        headers={"Authorization": f"Bearer {token}"},
        params=params,
        timeout=30,
    )
    if resp.status_code == 401:
        raise _TokenExpired()
    resp.raise_for_status()
    return resp.json()


def iter_browse_rows(
    filters: list[FilterCriteria],
    tm: TokenManager,
    page_size: int = 25,
) -> Iterator[dict]:
    """Yields one flat Browse row dict at a time, handling pagination + 401 retry."""
    call_id = str(uuid.uuid4())
    page = 1
    page_action = "first"

    while True:
        for attempt in range(2):
            try:
                resp_json = call_browse_api(filters, tm.get(), page, page_size, page_action, call_id)
                break
            except _TokenExpired:
                if attempt == 0:
                    tm.refresh()
                    continue
                raise

        rows = resp_json.get("data") or []
        if not rows:
            return
        for row in rows:
            yield row

        aux = resp_json.get("auxillaryData") or {}
        has_more = str(aux.get("HasMoreParams", "false")).strip().lower() == "true"
        if not has_more:
            return

        page += 1
        page_action = "next"  # ASSUMPTION -- see note above, not yet confirmed

# =============================================================================
# 3b. SCOPED PRICING DETAIL GET
# =============================================================================
# Maximum Price / Minimum Price / List Price are CONFIRMED absent from the
# Browse row (network capture 2026-07-27 -- full key list checked, none of
# the three present). They only exist on the keyed Detail GET. Only called
# when Amount Type == "1" (List Price), since that's the only pricing type
# these fields are relevant for. Soft-fail by design: any error here just
# leaves the three price columns blank on that row -- never fails the row
# or the overall fetch.

DETAIL_VIEW_URI = "urn:be:com.qad.sales.pricing.IPriceListV2"


def get_price_list_pricing_fields(key: dict, tm: "TokenManager") -> dict | None:
    url = f"{CONFIG['qad']['base_url']}/api/erp/priceListV2s"
    params = {
        "domainCode":    key.get("domainCode", ""),
        "priceListCode": key.get("priceListCode", ""),
        "customerCode":  key.get("customerCode", ""),
        "itemCode":      key.get("itemCode", ""),
        "attributeCode": key.get("attributeCode", ""),
        "orderCode":     key.get("orderCode", ""),
        "currencyCode":  key.get("currencyCode", ""),
        "unitOfMeasure": key.get("unitOfMeasure", ""),
        "startDate":     key.get("startDate", ""),
        "viewUri":       DETAIL_VIEW_URI,
    }

    for attempt in range(2):
        try:
            resp = requests.get(
                url,
                headers={"Authorization": f"Bearer {tm.get()}"},
                params=params,
                timeout=30,
            )
            if resp.status_code == 401:
                if attempt == 0:
                    tm.refresh()
                    continue
                return None
            resp.raise_for_status()
            body = resp.json()
            pl_list = body.get("data", {}).get("priceListV2s") or body.get("priceListV2s") or []
            if not pl_list:
                return None
            pl = pl_list[0]
            return {
                "Maximum Price": pl.get("maximumPrice", ""),
                "Minimum Price": pl.get("minimumPrice", ""),
                "List Price":    pl.get("listPrice", ""),
            }
        except requests.RequestException as e:
            print(f"[fetch_price_lists] WARNING: pricing detail GET failed for "
                  f"{key.get('priceListCode')}/{key.get('itemCode')}: {e} -- "
                  f"leaving price columns blank")
            return None

    return None

# =============================================================================
# 4. BROWSE ROW -> EXCEL ROW  (flat dotted keys -> friendly labels)
# =============================================================================

def flatten_browse_row(row: dict) -> dict:
    def g(field: str) -> str:
        v = row.get(f"priceListV2.{field}")
        return "" if v is None else v

    amount_type_code = str(g("amountType"))
    quantity_type_code = str(g("quantityType"))
    combination_type_code = str(g("combinationType"))

    return {
        "Domain":            g("domainCode"),
        "Price List":        g("priceListCode"),
        "Description":       g("priceListDescription"),
        "Customer Code":     g("customerCode"),
        "Item Code":         g("itemCode"),
        "Currency":          g("currencyCode"),
        "Unit Of Measure":   g("unitOfMeasure"),
        "Start Date":        _to_excel_date(g("startDate")),
        "Expire Date":       _to_excel_date(g("expireDate")),
        "Amount Type":       AMOUNT_TYPE_REVERSE_MAP.get(amount_type_code, amount_type_code),
        "Quantity Type":     QUANTITY_TYPE_REVERSE_MAP.get(quantity_type_code, quantity_type_code),
        "Combination Type":  COMBINATION_TYPE_REVERSE_MAP.get(combination_type_code, combination_type_code),
        # Minimum Order sourced from Browse's minNetOrd -- CONFIRMED 2026-07-24
        # equivalent to the old detail-only minimumQuantity field.
        "Minimum Order":     g("minNetOrd"),
        "Maximum Quantity":  g("maximumQty"),
        "Maximum orders":    g("maxOrders"),
        #--- pricing ---#
        "Maximum Price":      g("maximumPrice"),
        "Minimum Price":      g("minimumPrice"),
        "List Price":         g("listPrice"),
        "Break category":    g("breakCategory"),
        "Data Operation":    "U",
        "Status ":           "",
        "Error":             "",
    }


def _to_excel_date(iso_value) -> str:
    if not iso_value:
        return ""
    try:
        dt = datetime.strptime(str(iso_value).replace("Z", ""), "%Y-%m-%dT%H:%M:%S.%f")
        return dt.strftime("%d-%m-%Y")  # match Customer_load.py's date_val format
    except Exception:
        return str(iso_value)


def _to_qad_date(value: str) -> str:
    """UI date pickers submit 'YYYY-MM-DD'. QAD filters expect ISO8601 with
    a time component. Assumption: midnight UTC, since the UI only collects
    a date, not a time."""
    if not value:
        return value
    try:
        dt = datetime.strptime(value, "%Y-%m-%d")
        return dt.strftime("%Y-%m-%dT00:00:00.000Z")
    except ValueError:
        return value


# =============================================================================
# 5. UI FIELD -> QAD FIELD / OPERATOR MAPPING
# =============================================================================
UI_FIELD_TO_QAD = {
    "domain":      "priceListV2.domainCode",
    "price_list":  "priceListV2.priceListCode",
    "customer":    "priceListV2.customerCode",
    "item":        "priceListV2.itemCode",
    "currency":    "priceListV2.currencyCode",
    "uom":         "priceListV2.unitOfMeasure",
    "start_date":  "priceListV2.startDate",
    "expire_date": "priceListV2.expireDate",
    "amount_type": "priceListV2.amountType",
}

DATE_FIELDS = {"start_date", "expire_date"}

UI_OPERATOR_TO_CONDITION = {
    "equals":       "eq",
    "starts with":  "sw",
    "ends with":    "ew",
    "contains":     "cn",
    "greater than": "gt",
    "less than":    "lt",
}


def _resolve_value(field_key: str, raw_value: str) -> str:
    if not raw_value:
        return raw_value
    if field_key == "amount_type":
        return AMOUNT_TYPE_MAP.get(raw_value, raw_value)
    if field_key in DATE_FIELDS:
        return _to_qad_date(raw_value)
    return raw_value


def build_filter_criteria(ui_filters: list[dict]) -> list[FilterCriteria]:
    criteria: list[FilterCriteria] = []

    for f in ui_filters:
        field_key = (f.get("field") or "").strip()
        operator  = (f.get("operator") or "").strip().lower()
        qad_field = UI_FIELD_TO_QAD.get(field_key)

        if not qad_field or not operator:
            continue

        value_from = _resolve_value(field_key, (f.get("value_from") or "").strip())
        value_to   = _resolve_value(field_key, (f.get("value_to") or "").strip())

        if operator == "range":
            criteria.extend(build_range_criteria(qad_field, value_from, value_to))
            continue

        condition = UI_OPERATOR_TO_CONDITION.get(operator)
        if not condition or not value_from:
            continue

        criteria.append(FilterCriteria(qad_field, condition, value_from))

    return criteria


# =============================================================================
# 6. EXCEL WRITER  (must match price_list_load.py's expected layout exactly)
# =============================================================================

FILL_MAIN    = PatternFill(start_color="00B0F0", end_color="00B0F0", fill_type="solid")  # cyan
FILL_PRICING = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")  # green
FILL_LOADER  = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")  # navy
FILL_HEADER  = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")  # yellow

# Light purple fill for every row written by the fetch, so fetched records
# are visually distinct from new rows the user types in afterwards (which
# are left with no fill / white).
FILL_FETCHED = PatternFill(start_color="D9D2E9", end_color="D9D2E9", fill_type="solid")  # light purple

FONT_GROUP  = Font(bold=True, color="FFFFFF")   # white bold text on group bands
FONT_HEADER = Font(bold=True, color="000000")   # black, bold (row 2 headers)
FONT_DATA   = Font(bold=False, color="000000")  # black, not bold (fetched data rows)

CENTER = Alignment(horizontal="center", vertical="center")

# Thin border on all four sides -- applied to every cell (header + data)
# so the sheet reads like Excel's "All Borders" button was used.
_THIN_SIDE = Side(style="thin", color="000000")
BORDER_ALL = Border(left=_THIN_SIDE, right=_THIN_SIDE, top=_THIN_SIDE, bottom=_THIN_SIDE)

def _write_headers(ws: Worksheet) -> None:
    ws.cell(row=1, column=1, value="Main")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(MAIN_HEADERS))

    pricing_start = len(MAIN_HEADERS) + 1
    pricing_end = pricing_start + len(PRICING_HEADERS) - 1
    ws.cell(row=1, column=pricing_start, value="Pricing")
    ws.merge_cells(start_row=1, start_column=pricing_start, end_row=1, end_column=pricing_end)

    loader_start = pricing_end + 1
    loader_end = loader_start + len(LOADER_HEADERS) - 1
    ws.merge_cells(start_row=1, start_column=loader_start, end_row=1, end_column=loader_end)
    # Loader band intentionally left unlabeled -- colour only, matches reference sheet.

    # ── Row 1 group band coloring ───────────────────────────────────────
    for col in range(1, len(MAIN_HEADERS) + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = FILL_MAIN
        cell.font = FONT_GROUP
        cell.alignment = CENTER
        cell.border = BORDER_ALL

    for col in range(pricing_start, pricing_end + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = FILL_PRICING
        cell.font = FONT_GROUP
        cell.alignment = CENTER
        cell.border = BORDER_ALL

    for col in range(loader_start, loader_end + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = FILL_LOADER
        cell.font = FONT_GROUP
        cell.alignment = CENTER
        cell.border = BORDER_ALL

    # ── Row 2 column headers ────────────────────────────────────────────
    for i, header in enumerate(ALL_HEADERS, start=1):
        cell = ws.cell(row=2, column=i, value=header)
        cell.fill = FILL_HEADER
        cell.font = FONT_HEADER
        cell.alignment = CENTER
        cell.border = BORDER_ALL

def fetch_to_excel(
    filters: list[FilterCriteria],
    output_path: str,
    domain_code: str | None = None,
) -> tuple[int, int]:
    """
    Single-stage: Browse (paginated) -> write output_path.
    Returns (rows_written, rows_skipped). skipped is always 0 now --
    there's no per-row API call left that can fail independently.

    Every fetched row is filled light purple (FILL_FETCHED) with plain
    black text and a thin all-around border, so it's visually distinct
    from any new rows a user later types into the same sheet. There is
    no cell/sheet protection -- users are free to edit or add rows.
    """
    tm = TokenManager()

    wb = openpyxl.Workbook()
    ws = wb.active
    _write_headers(ws)

    row_idx = 3
    written = 0

    browse_rows = list(iter_browse_rows(filters, tm))  # materialize for a stable total
    total = len(browse_rows)

    for i, browse_row in enumerate(browse_rows, start=1):
        if _progress_callback:
            _progress_callback(i, total)

        flat = flatten_browse_row(browse_row)

        amount_type_code = str(browse_row.get("priceListV2.amountType", ""))
        if amount_type_code == "1":
            detail_key = {
                "domainCode":    browse_row.get("priceListV2.domainCode", ""),
                "priceListCode": browse_row.get("priceListV2.priceListCode", ""),
                "customerCode":  browse_row.get("priceListV2.customerCode", ""),
                "itemCode":      browse_row.get("priceListV2.itemCode", ""),
                "attributeCode": browse_row.get("priceListV2.attributeCode", ""),
                "orderCode":     browse_row.get("priceListV2.orderCode", ""),
                "currencyCode":  browse_row.get("priceListV2.currencyCode", ""),
                "unitOfMeasure": browse_row.get("priceListV2.unitOfMeasure", ""),
                "startDate":     browse_row.get("priceListV2.startDate", ""),
            }
            pricing = get_price_list_pricing_fields(detail_key, tm)
            if pricing:
                flat.update(pricing)

        for col_idx, header in enumerate(ALL_HEADERS, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=flat.get(header, ""))
            cell.fill = FILL_FETCHED
            cell.font = FONT_DATA
            cell.border = BORDER_ALL
        row_idx += 1
        written += 1

    wb.save(output_path)
    return written, 0


# =============================================================================
# 7. ENTRY POINT  (called from main.py's /api/fetch-price-list route)
# =============================================================================
# Contract expected by main.py:
#   def fetch(filters: list[dict], output_path: str, tm=None) -> dict
#   returns {"ok": bool, "message": str, "rows": int, "skipped": int}

def fetch(filters: list[dict], output_path: str, tm: TokenManager | None = None) -> dict:
    try:
        criteria = build_filter_criteria(filters)
    except Exception as exc:
        return {"ok": False, "message": f"Invalid filter: {exc}", "rows": 0, "skipped": 0}

    try:
        written, skipped = fetch_to_excel(criteria, output_path)
    except Exception as exc:
        return {"ok": False, "message": str(exc), "rows": 0, "skipped": 0}

    return {
        "ok": True,
        "message": "Extraction successful",
        "rows": written,
        "skipped": skipped,
    }


if __name__ == "__main__":
    demo_filters = [
        {"field": "price_list", "operator": "range", "value_from": "10AUTO1", "value_to": "CLP1"},
        {"field": "amount_type", "operator": "equals", "value_from": "Discount %", "value_to": None},
    ]
    result = fetch(demo_filters, "price_list_fetch_output.xlsx")
    print(result)