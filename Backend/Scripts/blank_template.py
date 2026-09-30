"""
Scripts/blank_template.py
==========================
Generates blank (headers-only) QAD import templates, fully in memory.

Contract expected by main.py (mirrors the other Scripts/*.py modules):

    TEMPLATE_SPECS: dict[str, dict]
        entity_id -> spec. Presence of a key here is what makes an entity
        show up as "supported" for the Blank Template button.

    supported_entities() -> list[str]
        Ordered list of entity ids main.py/the frontend can offer.

    build(entity_id: str) -> io.BytesIO
        Builds the workbook for one entity and returns it as an in-memory
        buffer (never touches disk) — main.py streams this straight back
        to the browser as a file download.

Why this lives in Scripts/, not main.py:
    Every other entity (Supplier, Customer, PriceList, ...) already has its
    validate_<entity>.py / <entity>_load.py living here, loaded dynamically
    by main.py via load_module(). Blank Template follows the same shape, so
    adding/adjusting a template later never touches main.py — just this file.

TWO SPEC SHAPES
----------------------------------------------------------------------------
This file supports two different TEMPLATE_SPECS shapes, because the
loaders themselves use two different sheet layouts:

1. BANNER shape (Supplier / Customer / PriceList) — row 1 = colored
   section banners, row 2 = the actual column headers those loaders read
   (Customer_load.py/Supplier_load.py/price_list_load.py all read row 2:
   "Row 1 = section labels, Row 2 = column headers"), row 3+ = data.
   Detected by the spec having a top-level "groups" key. UNCHANGED from
   before — nothing about this path was touched.

   To add a new BANNER-style entity:
     a. Add an entry with "groups" (row-1 section spans) and "columns"
        (row-2 headers), copied verbatim from a real export/import file.
     b. Each group is (title, span) or (title, span, color_hex). Omit
        color_hex to fall back to DEFAULT_SECTION_FILL.
     c. Nothing else changes — main.py and index.html pick it up
        automatically.

2. FLAT shape (Site / Location / PurchaseOrder / SalesOrder) — row 1 IS
   the column header row (no banner), row 2+ = data. This matches how
   site_load.py / location_load.py / po_load.py / so_load.py themselves
   read their files (ws[1] = headers, not ws[2]) — a BANNER-shaped
   template would be actively wrong for these, since the loader would
   try to read the banner text as field names.

   Detected by the spec having a top-level "sheets" key: a list of one
   sheet-spec per worksheet (one entry for Site/Location's single sheet,
   two for PurchaseOrder/SalesOrder's Header + Lines). Each sheet-spec:
     - sheet_name       worksheet title ("Sheet1", "Header", "Lines", ...)
     - columns           row-1 headers, in order, EXACTLY matching what
                          the loader's own sv()/row_data.get() calls read
                          (not necessarily what an old sample file has —
                          see the SalesOrder "Currency Code" note below)
     - header_fill        optional 6-char hex for the header row (falls
                           back to DEFAULT_HEADER_FILL if omitted)
     - no_fill_columns     optional set/list of column names that should
                           NOT get header_fill — e.g. {"Data operation",
                           "Status", "Error"} — still bordered, just
                           left white so they read as "loader-only", not
                           something the user fills in
     - dropdowns          optional {column_name: {"options": [...],
                           "error_title": "...", "error_msg": "..."}} —
                           each becomes an Excel list-validation dropdown
                           restricted to exactly those options, backed by
                           a hidden column on that same sheet (one hidden
                           column per dropdown, auto-allocated so multiple
                           dropdowns on one sheet never collide)

   To add a new FLAT-style entity: add a "sheets" list following the
   shape above. Column names MUST match what the load script actually
   reads (grep its sv(row_data, "...") / row_data.get("...") calls) —
   copying a sample .xlsx's headers isn't safe here the way it is for
   Supplier, because at least one existing sample (SalesOrder's HP.xlsx)
   has a column name ("Currency") that doesn't match what so_load.py
   actually reads ("Currency Code") — that's a pre-existing mismatch in
   that sample file, not something to replicate.

Row layout — BANNER shape:
    Row 1 — merged section banner cells (grouping related columns)
    Row 2 — the column headers the loader reads
    Row 3+ — blank / no-fill, freeze_panes = "A3"

Row layout — FLAT shape:
    Row 1 — the column headers the loader reads (yellow by default, with
             no_fill_columns left white)
    Row 2+ — blank / no-fill, freeze_panes = "A2"

Both shapes: header row(s) locked (sheet-protected), data rows explicitly
unlocked so the user can still type into a protected sheet.
"""

import io

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side, Protection
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# ═════════════════════════════════════════════════════════════════════════════
# TEMPLATE SPECS
# ═════════════════════════════════════════════════════════════════════════════
#
# groups: list of (section_title, column_count) or (section_title,
#         column_count, color_hex), consumed left-to-right — the spans must
#         sum to the same total width as `columns`. A blank "" title still
#         reserves/merges that span of columns under row 1, it just renders
#         with no banner text. color_hex is an optional 6-char RGB hex
#         string (no leading "FF"/"#") — omit it to use the default violet.
# columns: the row-2 headers, in order.
# header_fill: optional 6-char RGB hex for the row-2 header band for this
#              entity. Omit to use the default light-gray band.
# file_prefix: used to name the downloaded file, e.g. "Supplier_Blank_Template.xlsx"

TEMPLATE_SPECS: dict[str, dict] = {
    # Supplier's groups/columns/header color below are copied verbatim from
    # the current real Supplier import/export file (Supplier_test6.xlsx),
    # replacing the old/stale column set. The four banner colors (green /
    # orange / blue / orange) and the solid-yellow header row are the same
    # colors used in that source file, so this template now visually matches
    # what people already see when they open a real exported Supplier sheet.
    # Customer and PriceList below reuse this same green/orange/blue +
    # yellow-header color theme (cycled across their groups) for visual
    # consistency, even though there's no real source file to copy exact
    # colors from for them — their columns/spans are unchanged.
    "Supplier": {
        "sheet_name": "Sheet1",
        "file_prefix": "Supplier",
        "header_fill": "FFFF00",  # yellow — matches source file's row 2
        "groups": [
            ("Main & Domain Settings", 10, "92D050"),   # green
            ("Business Relation", 16, "FFC000"),        # orange
            ("Payment & Accounting", 6, "92D050"),      # green
            ("Tax Panel", 11, "00B0F0"),                # blue
            ("Develop Column (ignore) (leave by default C)", 3, "FFC000"),  # orange
        ],
        "columns": [
            # Main & Domain Settings (10)
            "Supplier", "Active", "Supplier Type", "Purchase Type", "Shared Set",
            "Site Code", "Daybook Set", "Domain", "Language", "Currency",
            # Business Relation (16)
            "Business Relation", "Active", "Name", "Search Name", "Second Name ",
            "Address Type", "Address 1", "Address 2", "Address 3", "City", "State",
            "Postal Code", "Country", "Telephone", "Email", "Tax Zone",
            # Payment & Accounting (6)
            "Credit Terms", "Invoice Status", "Invoice Control GL Profile",
            "Credit Note Control GL Profile", "Prepayment Control GL Profile",
            "Purchase Account GL Profile",
            # Tax Panel (11)
            "Taxable (Yes/No)", "Tax Included(Yes/No)", "Tax in City(Yes/No)",
            "Tax Report", "Federal tax", "State Tax", "Miscellaneous Tax 1",
            "Miscellaneous Tax 2", "Miscellaneous Tax 3", "Tax Class", "Tax Usuage ",
            # Develop Column (ignore) (3)
            "Data Operation", "Status", "Error",
        ],
    },
    "Customer": {
        "sheet_name": "Sheet1",
        "file_prefix": "Customer",
        "header_fill": "FFFF00",  # yellow — same theme as Supplier
        "groups": [
            ("Main & Domain Settings", 9, "92D050"),         # green
            ("Business Relation", 14, "FFC000"),             # orange
            ("Credit Limits Panel", 10, "00B0F0"),           # blue
            ("Credit Check Panel", 9, "92D050"),             # green
            ("Payment & Accouting Profile", 6, "FFC000"),    # orange
            ("Tax Panel", 11, "00B0F0"),                     # blue
            ("", 3),                                          # blank banner for loader
        ],
        "columns": [
            "Customer", "Active", "Customer Type", "Shared Set", "Site Code",
            "Daybook Set", "Domain", "Language", "Currency",
            "Business Relation", "Active", "Name", "Search Name", "Address Type",
            "Address 1", "Address 2", "City", "State", "Postal Code", "Country",
            "Telephone", "Email", "Tax Zone", "Fixed Credit Limit",
            "Apply Fixed Ceiling", "Apply % of Turnover", "Percentage of Turnover",
            "Maximum Days Overdue", "Apply Maximum Days Overdue", "Credit Hold (Yes/ No)",
            "Warning Ceiling %", "Credit Rating", "Credit Agency Ref",
            "Overrule Allowed SO(Yes/No)", "Calculate before Order Entry(Yes/No)",
            "Calculate after Order Entry(Yes/No)", "Calculate before Invoice(Yes/No)",
            "Calculate after Invoice Entry(Yes/No)", "Overrule Allowed Invoice(Yes/No)",
            "Include Drafts(Yes/No)", "Include Open Items(Yes/No)", "Include Sales Orders(Yes/No)",
            "Credit Terms", "Invoice Status", "Invoice Control GL Profile",
            "Credit Note Control GL Profile", "Prepayment Control GL Profile",
            "Sales Account GL Profile", "Taxable (Yes/No)", "Tax Included(Yes/No)",
            "Tax in City(Yes/No)", "Tax Report", "Federal Tax", "State Tax",
            "Miscellaneous Tax 1", "Miscellaneous Tax 2", "Miscellaneous Tax 3",
            "Tax Class", "Tax Usuage ", "Data Operation", "Status", "Error",
        ],
    },
    "PriceList": {
        "sheet_name": "Sheet",
        "file_prefix": "PriceList",
        "header_fill": "FFFF00",  # yellow — same theme as Supplier
        "groups": [
            ("Main", 9, "92D050"),     # green
            ("Pricing", 10, "FFC000"),  # orange
            ("", 3),                    # blank banner for loader
        ],
        "columns": [
            "Domain", "Price List", "Description", "Customer Code", "Item Code",
            "Currency", "Unit Of Measure", "Start Date", "Expire Date",
            "Amount Type", "Quantity Type", "Combination Type", "Minimum Order",
            "Maximum Quantity", "Maximum orders", "Maximum Price", "Minimum Price",
            "List Price", "Break category", "Data Operation", "Status ", "Error",
        ],
    },

    # ── FLAT-shape entities below (row 1 = headers, no banner row) ────────
    # See the module docstring's "TWO SPEC SHAPES" section. Column names
    # here are verified against each loader's own sv(row_data, "...") /
    # row_data.get("...") calls, not copied blind from a sample file.

    "Site": {
        "file_prefix": "Site",
        "sheets": [
            {
                "sheet_name": "Sheet1",
                "header_fill": "FFFF00",  # yellow — same theme as every other entity
                "no_fill_columns": {"Data operation", "Status", "Error"},
                "columns": [
                    "Domain", "Entity", "Site", "Description",
                    "Default inventory status", "Data operation", "Status", "Error",
                ],
                "dropdowns": {
                    "Data operation": {
                        "options": ["C", "U"],
                        "error_title": "Invalid Data operation",
                        "error_msg": "Please enter C (Create) or U (Update) only",
                    },
                    "Default inventory status": {
                        "options": [
                            "Expired", "N-N-N", "N-N-Y", "N-Y-N", "N-Y-Y",
                            "Y-N-N", "Y-N-Y", "Y-Y-N", "Y-Y-Y",
                        ],
                        "error_title": "Invalid inventory status",
                        "error_msg": "Please pick one of the listed inventory status combinations",
                    },
                },
            },
        ],
    },

    "Location": {
        "file_prefix": "Location",
        "sheets": [
            {
                "sheet_name": "Sheet1",
                "header_fill": "FFFF00",
                "no_fill_columns": {"Data Operation", "Status", "Error"},
                "columns": [
                    "Site", "Location", "Description", "Inventory Status",
                    "Location Type", "Data Operation", "Status", "Error",
                ],
                "dropdowns": {
                    "Data Operation": {
                        "options": ["C", "U"],
                        "error_title": "Invalid Data Operation",
                        "error_msg": "Please enter C (Create) or U (Update) only",
                    },
                    # "default" (see location_load.py) means: don't send
                    # inventoryStatusCode at all — QAD inherits the site's
                    # own default, same as leaving it blank on the manual
                    # Location Maintenance screen.
                    "Inventory Status": {
                        "options": [
                            "Expired", "N-N-N", "N-N-Y", "N-Y-N", "N-Y-Y",
                            "Y-N-N", "Y-N-Y", "Y-Y-N", "Y-Y-Y", "default",
                        ],
                        "error_title": "Invalid inventory status",
                        "error_msg": "Please pick one of the listed inventory status combinations, or 'default' to inherit the site's own default",
                    },
                },
            },
        ],
    },

    "PurchaseOrder": {
        "file_prefix": "PurchaseOrder",
        "sheets": [
            {
                # Columns verified against po_load.py's build_header_payload()
                # sv(row_data, "...") calls, matching the order and names
                # already used in the real PO_testing.xlsx / error_Kangaroo.xlsx
                # sample files (both already in Data/PurchaseOrder /
                # Data/FILES_COPIES) — those match po_load.py exactly, unlike
                # SalesOrder's sample (see below). "Ship Via" is present in
                # those real files but is NOT currently read by po_load.py —
                # kept here for parity with the real files; harmless no-op
                # until/unless po_load.py is extended to use it.
                "sheet_name": "Header",
                "header_fill": "FFFF00",
                "no_fill_columns": {"Data Operation", "Status", "Error"},
                "columns": [
                    "PO Number", "Domain Code", "Supplier Code", "Currency",
                    "Credit Terms", "Daybook Set", "Ship To Site", "Bill To Code",
                    "Ship Via", "Order Date", "Due Date",
                    "Data Operation", "Status", "Error",
                ],
                "dropdowns": {
                    "Data Operation": {
                        "options": ["C", "U"],
                        "error_title": "Invalid Data Operation",
                        "error_msg": "Please enter C (Create) or U (Update) only",
                    },
                },
            },
            {
                # Verified against po_load.py's create_line() sv()/fv() calls.
                # "Domain Code" is present in the real sample files' Lines
                # sheet too, but po_load.py takes domain from the HEADER row
                # only (create_line(domain, po_num, ...)) — the Lines-sheet
                # column, if present, is never read. Kept for parity with
                # the real files; harmless.
                "sheet_name": "Lines",
                "header_fill": "FFFF00",
                "no_fill_columns": {"Data Operation", "Status", "Error"},
                "columns": [
                    "PO Number", "Domain Code", "Item Code", "Site Code",
                    "Quantity Ordered", "Unit Price",
                    "Data Operation", "Status", "Error",
                ],
                "dropdowns": {
                    "Data Operation": {
                        "options": ["C", "U"],
                        "error_title": "Invalid Data Operation",
                        "error_msg": "Please enter C (Create) or U (Update) only",
                    },
                },
            },
        ],
    },

    "SalesOrder": {
        "file_prefix": "SalesOrder",
        "sheets": [
            {
                # IMPORTANT: "Currency Code" here, not "Currency". so_load.py
                # reads sv(row_data, "Currency Code") — the real HP.xlsx
                # sample file actually has a column named "Currency" (no
                # "Code"), which is a pre-existing mismatch in that sample:
                # every row in HP.xlsx currently fails so_load.py's own
                # mandatory-field check on this field. This template uses
                # the name the loader actually reads, not the sample's.
                "sheet_name": "Header",
                "header_fill": "FFFF00",
                "no_fill_columns": {"Data Operation", "Status", "Error"},
                "columns": [
                    "Domain Code", "SO Number", "Sold To Customer Code",
                    "Bill To Customer Code", "Ship To Customer Code", "Site Code",
                    "Currency Code", "Daybook Set", "Credit Terms",
                    "Order Date", "Due Date", "Ship Via", "Freight List", "Freight Terms",
                    "Data Operation", "Status", "Error",
                ],
                "dropdowns": {
                    "Data Operation": {
                        "options": ["C", "U"],
                        "error_title": "Invalid Data Operation",
                        "error_msg": "Please enter C (Create) or U (Update) only",
                    },
                },
            },
            {
                # Verified against so_load.py's create_line() sv()/fv() calls.
                # "Line Number" is supported (an explicit override) but
                # optional — omitted here to match the real HP.xlsx sample;
                # so_load.py falls back to sequential position when absent.
                "sheet_name": "Lines",
                "header_fill": "FFFF00",
                "no_fill_columns": {"Data Operation", "Status", "Error"},
                "columns": [
                    "SO Number", "Item Code", "Site Code", "Quantity Ordered",
                    "List Price", "Discount", "Net Price", "Due Date",
                    "Data Operation", "Status", "Error",
                ],
                "dropdowns": {
                    "Data Operation": {
                        "options": ["C", "U"],
                        "error_title": "Invalid Data Operation",
                        "error_msg": "Please enter C (Create) or U (Update) only",
                    },
                },
            },
        ],
    },

    "ReceiptsUnplanned": {
        # Flat single-sheet, one row = one complete receipt — confirmed
        # with the requester that unplannedReceives has no real multi-line
        # flow despite the API's header/detail grid shape (see
        # receipts-unplanned_load.py's module docstring). Columns verified
        # against that loader's own sv()/fv() reads, not a sample file —
        # there wasn't one for this entity, only live network captures.
        "file_prefix": "ReceiptsUnplanned",
        "sheets": [
            {
                "sheet_name": "Sheet1",
                "header_fill": "FFFF00",
                "no_fill_columns": {"Data Operation", "Status", "Error"},
                "columns": [
                    "Domain", "Site", "Item Code", "Quantity",
                    "Data Operation", "Status", "Error",
                ],
                "dropdowns": {
                    "Data Operation": {
                        "options": ["C", "U"],
                        "error_title": "Invalid Data Operation",
                        "error_msg": "Please enter C (Create) or U (Update) only",
                    },
                },
            },
        ],
    },
}

# ═════════════════════════════════════════════════════════════════════════════
# STYLES
# ═════════════════════════════════════════════════════════════════════════════
# Row 1 = section banner. Default is a light violet fill / violet text,
#         unless a group supplies its own color_hex (see Supplier above).
# Row 2 = column headers. Default is plain light-gray / black text, unless
#         the entity's spec supplies "header_fill" (see Supplier above).
# Row 3+ = explicitly forced to NO FILL so nothing ever bleeds into the rows
#          where the user actually types data.

DEFAULT_SECTION_FILL = PatternFill("solid", fgColor="EDE7F6")
DEFAULT_HEADER_FILL  = PatternFill("solid", fgColor="F5F5F5")  # light gray
_NO_FILL             = PatternFill(fill_type=None)              # forced white/no-fill for data rows

_ALL_BORDER   = Border(*(Side(style="thin", color="000000") for _ in range(4)))  # solid black, all 4 sides

_SECTION_FONT = Font(name="Arial", size=10, bold=True, color="000000")  # black
_HEADER_FONT  = Font(name="Arial", size=10, bold=True, color="000000")
_CENTER       = Alignment(horizontal="center", vertical="center", wrap_text=True)
_LEFT         = Alignment(horizontal="left", vertical="center", wrap_text=True)

# How many blank data rows below the headers get an explicit no-fill stamp.
# Doesn't limit how many rows a user can actually use — Excel cells beyond
# this are already unstyled/white by default — this just belt-and-braces
# the first chunk of rows people are most likely to click into or drag-fill.
_FORCE_WHITE_ROWS = 500


def _section_fill(color_hex: str | None) -> PatternFill:
    """Per-group banner fill: a group's own color if given, else the default."""
    return PatternFill("solid", fgColor=color_hex) if color_hex else DEFAULT_SECTION_FILL


def _header_fill(color_hex: str | None) -> PatternFill:
    """Per-entity header-row fill: the spec's own color if given, else the default."""
    return PatternFill("solid", fgColor=color_hex) if color_hex else DEFAULT_HEADER_FILL


def _add_dropdown(
    ws, col_idx: int, options: list[str], error_title: str, error_msg: str,
    hidden_col: int, data_start_row: int, data_end_row: int,
) -> None:
    """
    Restrict one column (data_start_row..data_end_row) to a fixed list of
    options via Excel data validation. The option list itself is written
    into a hidden column on the same sheet (openpyxl/Excel list validation
    needs a real cell range, not an inline literal list, once you're past
    a handful of short options) — hidden_col must be unique per dropdown
    on a given sheet; callers allocate sequentially (100, 101, 102, ...)
    so multiple dropdowns on one sheet never collide.
    """
    for i, opt in enumerate(options, start=1):
        ws.cell(row=i, column=hidden_col, value=opt)
    ws.column_dimensions[get_column_letter(hidden_col)].hidden = True

    hidden_range = f"{get_column_letter(hidden_col)}$1:${get_column_letter(hidden_col)}${len(options)}"
    dv = DataValidation(
        type="list",
        formula1=hidden_range,
        allow_blank=True,
        showDropDown=False,
        showErrorMessage=True,
        errorTitle=error_title,
        error=error_msg,
    )
    ws.add_data_validation(dv)
    col_letter = get_column_letter(col_idx)
    dv.add(f"{col_letter}{data_start_row}:{col_letter}{data_end_row}")


# ═════════════════════════════════════════════════════════════════════════════
# PUBLIC API  —  called from main.py
# ═════════════════════════════════════════════════════════════════════════════

def supported_entities() -> list[str]:
    """Entity ids this script currently knows how to build a template for."""
    return list(TEMPLATE_SPECS.keys())


def _build_banner_sheet(wb: Workbook, spec: dict) -> None:
    """
    ORIGINAL layout, UNCHANGED: row 1 = merged section banners, row 2 =
    column headers, row 3+ = forced no-fill data-entry area. Used by
    Supplier / Customer / PriceList (specs with a "groups" key).
    """
    columns = spec["columns"]
    groups  = spec["groups"]

    if sum(span for _, span, *_ in groups) != len(columns):
        # Defensive check — keeps a future typo in TEMPLATE_SPECS from
        # silently producing a misaligned banner row instead of failing loudly.
        raise ValueError("groups/columns width mismatch")

    ws = wb.active
    ws.title = spec["sheet_name"]

    header_fill = _header_fill(spec.get("header_fill"))

    # ── Row 1: merged section banners ──
    col_cursor = 1
    for group in groups:
        title, span, *rest = group
        color_hex = rest[0] if rest else None
        fill = _section_fill(color_hex)

        start_col = col_cursor
        end_col   = col_cursor + span - 1
        ws.merge_cells(start_row=1, start_column=start_col, end_row=1, end_column=end_col)
        cell = ws.cell(row=1, column=start_col, value=title or None)
        cell.font = _SECTION_FONT
        cell.fill = fill
        cell.alignment = _CENTER
        cell.protection = Protection(locked=True)
        for c in range(start_col, end_col + 1):
            ws.cell(row=1, column=c).border = _ALL_BORDER
            ws.cell(row=1, column=c).fill = fill
            ws.cell(row=1, column=c).protection = Protection(locked=True)
        col_cursor = end_col + 1

    # ── Row 2: column headers ──
    data_operation_col = None  # Track which column has "Data Operation"
    for idx, name in enumerate(columns, start=1):
        cell = ws.cell(row=2, column=idx, value=name)
        cell.font = _HEADER_FONT
        cell.fill = header_fill
        cell.alignment = _LEFT
        cell.border = _ALL_BORDER
        cell.protection = Protection(locked=True)
        width = max(12, min(34, len(str(name)) + 4))
        ws.column_dimensions[get_column_letter(idx)].width = width

        # Track Data Operation column for dropdown
        if name == "Data Operation":
            data_operation_col = idx

    # ── Row 3+: force no-fill and UNLOCK for editing ──
    total_cols = len(columns)
    for r in range(3, 3 + _FORCE_WHITE_ROWS):
        for c in range(1, total_cols + 1):
            cell = ws.cell(row=r, column=c)
            cell.fill = _NO_FILL
            cell.protection = Protection(locked=False)

    # ── ADD DATA VALIDATION DROPDOWN for "Data Operation" ──
    if data_operation_col:
        _add_dropdown(
            ws, data_operation_col, ["C", "U"],
            "Invalid Data Operation",
            "Please enter C (Create) or U (Update) only",
            hidden_col=100,
            data_start_row=3, data_end_row=3 + _FORCE_WHITE_ROWS - 1,
        )

    ws.protection.sheet = True
    ws.protection.enable()

    ws.row_dimensions[1].height = 20
    ws.row_dimensions[2].height = 30
    ws.freeze_panes = "A3"   # keep both header rows visible while scrolling data


def _build_flat_sheet(wb: Workbook, sheet_spec: dict, is_first: bool) -> None:
    """
    NEW layout: row 1 IS the column header row (no banner), row 2+ = forced
    no-fill data-entry area. Used by Site / Location / PurchaseOrder /
    SalesOrder (specs with a "sheets" key) — matches how those loaders
    themselves read row 1 as the header row.
    """
    ws = wb.active if is_first else wb.create_sheet()
    ws.title = sheet_spec["sheet_name"]

    columns         = sheet_spec["columns"]
    header_fill     = _header_fill(sheet_spec.get("header_fill"))
    no_fill_columns = set(sheet_spec.get("no_fill_columns", []))

    # ── Row 1: column headers (no_fill_columns left white, still bordered) ──
    for idx, name in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=idx, value=name)
        cell.font = _HEADER_FONT
        cell.fill = _NO_FILL if name in no_fill_columns else header_fill
        cell.alignment = _LEFT
        cell.border = _ALL_BORDER
        cell.protection = Protection(locked=True)
        width = max(12, min(34, len(str(name)) + 4))
        ws.column_dimensions[get_column_letter(idx)].width = width

    # ── Row 2+: force no-fill and UNLOCK for editing ──
    total_cols = len(columns)
    for r in range(2, 2 + _FORCE_WHITE_ROWS):
        for c in range(1, total_cols + 1):
            cell = ws.cell(row=r, column=c)
            cell.fill = _NO_FILL
            cell.protection = Protection(locked=False)

    # ── Dropdowns — one hidden helper column each, auto-allocated ──
    hidden_col = 100
    for col_name, dd in sheet_spec.get("dropdowns", {}).items():
        if col_name not in columns:
            continue
        col_idx = columns.index(col_name) + 1
        options = dd["options"]
        _add_dropdown(
            ws, col_idx, options,
            dd.get("error_title", f"Invalid {col_name}"),
            dd.get("error_msg", f"Choose one of: {', '.join(options)}"),
            hidden_col=hidden_col,
            data_start_row=2, data_end_row=2 + _FORCE_WHITE_ROWS - 1,
        )
        hidden_col += 1

    ws.protection.sheet = True
    ws.protection.enable()

    ws.row_dimensions[1].height = 28
    ws.freeze_panes = "A2"   # keep the single header row visible while scrolling data


def build(entity_id: str) -> io.BytesIO:
    """
    Build a blank (headers-only) workbook for `entity_id`, entirely in
    memory. Dispatches to the BANNER layout (Supplier/Customer/PriceList —
    spec has "groups") or the FLAT layout (Site/Location/PurchaseOrder/
    SalesOrder — spec has "sheets") — see the module docstring.

    Raises ValueError if entity_id isn't in TEMPLATE_SPECS.
    """
    if entity_id not in TEMPLATE_SPECS:
        raise ValueError(f"No blank template available for entity '{entity_id}'")

    spec = TEMPLATE_SPECS[entity_id]
    wb   = Workbook()

    if "sheets" in spec:
        for i, sheet_spec in enumerate(spec["sheets"]):
            _build_flat_sheet(wb, sheet_spec, is_first=(i == 0))
    else:
        _build_banner_sheet(wb, spec)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf