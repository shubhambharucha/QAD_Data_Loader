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

To add a new entity's template:
    1. Add an entry to TEMPLATE_SPECS below with its "groups" (row-1 section
       banners) and "columns" (row-2 headers), copied verbatim from a real
       export/import file for that entity.
    2. Each group is a (title, span) or (title, span, color_hex) tuple. If
       you omit color_hex the group falls back to DEFAULT_SECTION_FILL (the
       plain violet banner) — that's what Customer/PriceList still use.
       Give a group its own color_hex to make it stand out, the way Supplier
       now mirrors its real source file's green/orange/blue section banners.
    3. Nothing else changes — main.py and index.html pick it up automatically
       (index.html asks main.py for the supported list at load time).

Row layout produced for every entity:
    Row 1 — merged "section" banner cells (grouping related columns),
             each one colorable independently via the group's color_hex
    Row 2 — the actual column headers that validate_<entity>.py / the load
             script read (fill color configurable per-entity via
             spec["header_fill"], defaults to DEFAULT_HEADER_FILL)
    Row 3+ — left completely blank / no-fill for the user to type data into
             (explicitly forced to no-fill so nothing inherits the header
             colors when a user types or drags fill-down in Excel)
"""

import io

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

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
            ("Tax Panel", 14, "00B0F0"),                     # blue — includes trailing Data Operation / Status / Error, no separate banner for those
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
        ],
        "columns": [
            "Domain", "Price List", "Description", "Customer Code", "Item Code",
            "Currency", "Unit Of Measure", "Start Date", "Expire Date",
            "Amount Type", "Quantity Type", "Combination Type", "Minimum Order",
            "Maximum Quantity", "Maximum orders", "Maximum Price", "Minimum Price",
            "List Price", "Break category", "Data Operation", "Status ", "Error",
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


# ═════════════════════════════════════════════════════════════════════════════
# PUBLIC API  —  called from main.py
# ═════════════════════════════════════════════════════════════════════════════

def supported_entities() -> list[str]:
    """Entity ids this script currently knows how to build a template for."""
    return list(TEMPLATE_SPECS.keys())


def build(entity_id: str) -> io.BytesIO:
    """
    Build a blank (headers-only) workbook for `entity_id`, entirely in memory.
    Row 1 = merged section banners (each group optionally colored via its own
    color_hex), Row 2 = column headers (fill optionally overridden per-entity
    via spec["header_fill"]), Row 3+ = forced no-fill data-entry area.
    Raises ValueError if entity_id isn't in TEMPLATE_SPECS.
    """
    if entity_id not in TEMPLATE_SPECS:
        raise ValueError(f"No blank template available for entity '{entity_id}'")

    spec    = TEMPLATE_SPECS[entity_id]
    columns = spec["columns"]
    groups  = spec["groups"]

    if sum(span for _, span, *_ in groups) != len(columns):
        # Defensive check — keeps a future typo in TEMPLATE_SPECS from
        # silently producing a misaligned banner row instead of failing loudly.
        raise ValueError(f"TEMPLATE_SPECS['{entity_id}'] groups/columns width mismatch")

    wb = Workbook()
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
        for c in range(start_col, end_col + 1):
            ws.cell(row=1, column=c).border = _ALL_BORDER
            ws.cell(row=1, column=c).fill = fill
        col_cursor = end_col + 1

    # ── Row 2: column headers ──
    for idx, name in enumerate(columns, start=1):
        cell = ws.cell(row=2, column=idx, value=name)
        cell.font = _HEADER_FONT
        cell.fill = header_fill
        cell.alignment = _LEFT
        cell.border = _ALL_BORDER
        width = max(12, min(34, len(str(name)) + 4))
        ws.column_dimensions[get_column_letter(idx)].width = width

    # ── Row 3+: force no-fill so the data-entry area never inherits a header color ──
    total_cols = len(columns)
    for r in range(3, 3 + _FORCE_WHITE_ROWS):
        for c in range(1, total_cols + 1):
            ws.cell(row=r, column=c).fill = _NO_FILL

    ws.row_dimensions[1].height = 20
    ws.row_dimensions[2].height = 30
    ws.freeze_panes = "A3"   # keep both header rows visible while scrolling data

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf