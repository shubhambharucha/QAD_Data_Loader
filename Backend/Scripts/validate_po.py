"""
PO Validator
------------
Validates PO_testing.xlsx (Header + Lines sheets) against constraints
derived from PO_Constraints.xlsx — only for fields actually present
in the sample file.
"""

import re
import openpyxl
from openpyxl.styles import PatternFill
from datetime import datetime

RED_FILL   = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
CLEAR_FILL = PatternFill(fill_type=None)

# ── Field rules derived from PO_Constraints.xlsx for columns present in PO_testing.xlsx ──
# Header sheet mapping:  display column name → (data_type, max_len)
#   data_type: "character" | "date" | "decimal" | "integer"
#   max_len:   int or None (no length check for dates/decimals)
HEADER_RULES = {
    "PO Number":     {"type": "character", "max_len": 8},   # PurchaseOrderNumber x(8)
    "Domain Code":   {"type": "character", "max_len": 8},   # DomainCode x(8)
    "Supplier Code": {"type": "character", "max_len": 8},   # SupplierCode x(8)
    "Currency":      {"type": "character", "max_len": 3},   # CurrencyCode x(3)
    "Credit Terms":  {"type": "character", "max_len": 8},   # CreditTermsCode x(8)
    "Daybook Set":   {"type": "character", "max_len": 8},   # DaybookSetCode x(8)
    "Ship To Site":  {"type": "character", "max_len": 8},   # ShipToSite x(8)
    "Bill To Code":  {"type": "character", "max_len": 8},   # BillToCode x(8)
    "Order Date":    {"type": "date",      "max_len": None}, # OrderDate date
    "Due Date":      {"type": "date",      "max_len": None}, # DueDate date
    "Exchange Rate": {"type": "decimal",   "max_len": None}, # ExchangeRate decimal
}

# Lines sheet: Item Code, Site Code, Quantity Ordered, Unit Price not in constraints file.
# We apply type-only rules (no length) for numeric fields, and mark empties.
LINES_RULES = {
    "PO Number":         {"type": "character", "max_len": 8},   # FK ref
    "Domain Code":       {"type": "character", "max_len": 8},
    "Line Number":       {"type": "integer",   "max_len": None},
    "Item Code":         {"type": "character", "max_len": None}, # not in constraints
    "Site Code":         {"type": "character", "max_len": 8},   # SiteCode x(8)
    "Quantity Ordered":  {"type": "decimal",   "max_len": None},
    "Unit Price":        {"type": "decimal",   "max_len": None},
}

HEADER_ENTITY_COL = "PO Number"
LINES_ENTITY_COL  = "PO Number"
VALID_DATE_TYPES  = (datetime,)


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


def _validate_cell(col_name, value, rule):
    """Returns error string or None."""
    if _is_empty(value):
        return f"{col_name}: empty"

    dtype = rule["type"]
    max_len = rule.get("max_len")

    if dtype == "date":
        if not _is_date(value):
            # Try parsing string
            sv = _str_val(value)
            try:
                datetime.strptime(sv, "%Y-%m-%d")
            except ValueError:
                try:
                    datetime.strptime(sv, "%m/%d/%Y")
                except ValueError:
                    return f"{col_name}: invalid date (got '{sv}')"
        return None

    if dtype in ("decimal", "integer"):
        if not _is_numeric(value):
            return f"{col_name}: must be numeric (got '{_str_val(value)}')"
        if dtype == "integer":
            try:
                int(float(value))
            except (TypeError, ValueError):
                return f"{col_name}: must be integer"
        return None

    # character
    sv = _str_val(value)
    if max_len and len(sv) > max_len:
        return f"{col_name}: max {max_len} chars (got {len(sv)})"
    return None


def _validate_sheet(ws, rules, entity_col):
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
            col_idx = header_row.index(col_name) + 1
            value   = row_data.get(col_name)
            err     = _validate_cell(col_name, value, rule)
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
        total_errors += _validate_sheet(wb["Lines"], LINES_RULES, LINES_ENTITY_COL)

    wb.save(file_path)
    has_errors = total_errors > 0
    return has_errors, total_errors


if __name__ == "__main__":
    import sys
    path = sys.argv[1] if len(sys.argv) > 1 else "PO_testing.xlsx"
    has_errors, count = validate(path)
    print(f"Validation complete. Errors found: {has_errors} ({count} rows with errors)")