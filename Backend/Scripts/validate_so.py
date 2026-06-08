"""
SO Validator
------------
Validates SO.xlsx (Header + Lines sheets) against constraints
derived from SO_Constraints.xlsx — only for fields actually present
in the sample file.
"""

import re
import openpyxl
from openpyxl.styles import PatternFill
from datetime import datetime

RED_FILL   = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
CLEAR_FILL = PatternFill(fill_type=None)

# ── Field rules derived from SO_Constraints.xlsx for columns present in SO.xlsx ──
HEADER_RULES = {
    "Domain Code":            {"type": "character", "max_len": 8},   # DomainCode x(8)
    "SO Number":              {"type": "character", "max_len": 8},   # SalesOrderNumber x(8)
    "Sold To Customer Code":  {"type": "character", "max_len": 8},   # SoldToCustomerCode x(8)
    "Bill To Customer Code":  {"type": "character", "max_len": 8},   # BillToCustomerCode x(8)
    "Ship To Customer Code":  {"type": "character", "max_len": 8},   # ShipToCustomerCode x(8)
    "Site Code":              {"type": "character", "max_len": 8},   # SiteCode x(8)
    "Currency":               {"type": "character", "max_len": 3},   # CurrencyCode x(3)
    "Daybook Set":            {"type": "character", "max_len": 8},   # DaybookSetCode x(8)
    "Credit Terms":           {"type": "character", "max_len": 8},   # CreditTermsCode x(8)
    "Order Date":             {"type": "date",      "max_len": None}, # OrderDate date
    "Due Date":               {"type": "date",      "max_len": None}, # DueDate date
    "Ship Via":               {"type": "character", "max_len": 20},  # ShipVia x(20)
}

# Lines: SO Number, Line Number, Item Code, Site Code, Quantity Ordered, List Price, Discount, Net Price, Due Date
# Only SiteCode and DueDate found in SO constraints; rest validated as type-only.
LINES_RULES = {
    "SO Number":        {"type": "character", "max_len": 8},   # FK ref
    "Line Number":      {"type": "integer",   "max_len": None},
    "Item Code":        {"type": "character", "max_len": None}, # not in constraints
    "Site Code":        {"type": "character", "max_len": 8},   # SiteCode x(8)
    "Quantity Ordered": {"type": "decimal",   "max_len": None},
    "List Price ":      {"type": "decimal",   "max_len": None}, # note trailing space in header
    "Discount":         {"type": "decimal",   "max_len": None}, # optional — skip empty
    "Net Price":        {"type": "decimal",   "max_len": None}, # optional — skip empty
    "Due Date":         {"type": "date",      "max_len": None}, # DueDate date
}

# Fields allowed to be empty (optional)
OPTIONAL_LINES = {"Discount", "Net Price"}

HEADER_ENTITY_COL = "SO Number"
LINES_ENTITY_COL  = "SO Number"


def _str_val(value):
    if value is None:
        return ""
    if isinstance(value, datetime):
        return str(value)
    return str(value).strip()


def _is_empty(value):
    sv = _str_val(value)
    return sv == "" or sv.lower() == "none"


def _is_numeric(value):
    try:
        float(value)
        return True
    except (TypeError, ValueError):
        return False


def _is_date(value):
    return isinstance(value, datetime)


def _validate_cell(col_name, value, rule, optional=False):
    """Returns error string or None."""
    if _is_empty(value):
        if optional:
            return None
        return f"{col_name.strip()}: empty"

    dtype = rule["type"]
    max_len = rule.get("max_len")

    if dtype == "date":
        if not _is_date(value):
            sv = _str_val(value)
            try:
                datetime.strptime(sv, "%Y-%m-%d")
            except ValueError:
                try:
                    datetime.strptime(sv, "%m/%d/%Y")
                except ValueError:
                    return f"{col_name.strip()}: invalid date (got '{sv}')"
        return None

    if dtype in ("decimal", "integer"):
        if not _is_numeric(value):
            return f"{col_name.strip()}: must be numeric (got '{_str_val(value)}')"
        if dtype == "integer":
            try:
                int(float(value))
            except (TypeError, ValueError):
                return f"{col_name.strip()}: must be integer"
        return None

    # character
    sv = _str_val(value)
    if max_len and len(sv) > max_len:
        return f"{col_name.strip()}: max {max_len} chars (got {len(sv)})"
    return None


def _validate_sheet(ws, rules, entity_col, optional_cols=None):
    optional_cols = optional_cols or set()
    header_row = [cell.value for cell in ws[1]]

    # Ensure Status / Error columns exist
    if "Status" not in header_row:
        ws.cell(row=1, column=len(header_row) + 1, value="Status")
        header_row.append("Status")
    if "Error" not in header_row:
        ws.cell(row=1, column=len(header_row) + 1, value="Error")
        header_row.append("Error")

    status_col_idx = header_row.index("Status") + 1
    error_col_idx  = header_row.index("Error")  + 1
    entity_col_idx = header_row.index(entity_col) + 1 if entity_col in header_row else None

    error_count = 0

    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        row_values = [cell.value for cell in row]
        if not any(row_values):
            continue

        row_data = dict(zip(header_row, row_values))
        current_status = str(row_data.get("Status", "")).strip().upper()
        if current_status in ("DONE", "READY"):
            continue

        row_errors      = []
        error_cell_idxs = []

        for col_name, rule in rules.items():
            if col_name not in header_row:
                continue
            col_idx  = header_row.index(col_name) + 1
            value    = row_data.get(col_name)
            optional = col_name.strip() in optional_cols
            err      = _validate_cell(col_name, value, rule, optional=optional)
            if err:
                row_errors.append(err)
                error_cell_idxs.append(col_idx)

        # Clear old highlights
        for cell in ws[row_idx]:
            cell.fill = CLEAR_FILL

        if row_errors:
            error_count += 1
            error_msg = "; ".join(row_errors)
            ws.cell(row=row_idx, column=status_col_idx, value="")
            if entity_col_idx:
                ws.cell(row=row_idx, column=entity_col_idx).fill = RED_FILL
            for cidx in error_cell_idxs:
                ws.cell(row=row_idx, column=cidx).fill = RED_FILL
            ws.cell(row=row_idx, column=error_col_idx, value=error_msg).fill = RED_FILL
        else:
            ws.cell(row=row_idx, column=status_col_idx, value="READY")
            ws.cell(row=row_idx, column=error_col_idx,  value="")

    return error_count


def validate(file_path):
    wb = openpyxl.load_workbook(file_path)

    total_errors = 0

    if "Header" in wb.sheetnames:
        total_errors += _validate_sheet(wb["Header"], HEADER_RULES, HEADER_ENTITY_COL)

    if "Lines" in wb.sheetnames:
        total_errors += _validate_sheet(
            wb["Lines"], LINES_RULES, LINES_ENTITY_COL,
            optional_cols=OPTIONAL_LINES
        )

    wb.save(file_path)
    has_errors = total_errors > 0
    return has_errors, total_errors


if __name__ == "__main__":
    import sys
    path = sys.argv[1] if len(sys.argv) > 1 else "SO.xlsx"
    has_errors, count = validate(path)
    print(f"Validation complete. Errors found: {has_errors} ({count} rows with errors)")