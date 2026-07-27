"""
validate_price_list.py
-----------------------
Validates Price List workbooks in-place (same file, red-highlighted errors),
following the same contract as validate_customer.py / validate_supplier.py:

    validate(file_path: str) -> dict
        {
          "rows_passed":  int,
          "rows_failed":  int,
          "rows_skipped": int,
          "has_errors":   bool,
        }

Called by main.py's validate_stream(), which also handles the file
rename-to-error_/rename-back-to-clean logic based on has_errors.

Sheet layout (row 1 = section labels, row 2 = headers, row 3+ = data):
  Domain | Price List | Description | Customer Code | Item Code | Currency |
  Unit Of Measure | Start Date | Expire Data | Amount Type | Quantity Type |
  Combination Type | Minimum Order | Maximum Quantity | Maximum orders |
  Break category | Data Operation | Status | Error | _Fetch Currency (hidden)

CONFIRMED business rules (2026-07-23):
  - Amount is QAD-calculated -- not in this sheet, nothing to validate.
  - Break Category is free text with no format/lookup constraint -- QAD
    accepts arbitrary values, so we only check it isn't absurdly long.
  - No quantity-break handling needed (one detail row per key).
  - Attribute Code / Order Code don't exist in this sheet at all.
  - Start Date / Customer Code / Item Code / Unit Of Measure / Domain /
    Price List are LOCKED on Update rows (protected in Excel) -- validation
    does not need to re-derive or double-check their immutability, that's
    a worksheet-protection concern, not a data-quality one.
"""

import os
import sys
from datetime import datetime

import openpyxl

SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPTS_DIR)

from excel_format_utils import mark_error, RED_FILL, CLEAR_FILL  # noqa: E402  (shared util, same as Customer)

from mappings_price_list import (
    VALID_AMOUNT_TYPES,
    VALID_QUANTITY_TYPES,
    VALID_COMBINATION_TYPES,
)


# ── Mandatory columns ──────────────────────────────────────────────────────
MANDATORY_COLUMNS = [
    "Domain",
    "Price List",
    "Description",
    "Currency",
    "Unit Of Measure",
    "Start Date",
]

# ── Columns that must be numeric if present (all optional numerically) ────
NUMERIC_COLUMNS = [
    "Minimum Order",
    "Maximum Quantity",
    "Maximum orders",
]

# ── Length limits ───────────────────────────────────────────────────────────
# NOTE: these are reasonable QAD field-size defaults, not yet confirmed
# against your schema -- adjust MAX_LENGTHS if your instance differs.
MAX_LENGTHS = {
    "Domain":            8,
    "Price List":        20,
    "Description":       100,
    "Customer Code":     15,
    "Item Code":         30,
    "Currency":          3,
    "Unit Of Measure":   4,
    "Break category":    30,
}

DATE_FORMAT = "%d-%m-%Y"  # matches price_list_fetch.py's _to_excel_date output


def _parse_date(value) -> datetime | None:
    if value is None or str(value).strip() == "":
        return None
    if isinstance(value, datetime):
        return value
    try:
        return datetime.strptime(str(value).strip(), DATE_FORMAT)
    except ValueError:
        return None


def _is_numeric(value) -> bool:
    if value is None or str(value).strip() == "":
        return True  # optional numeric fields -- blank is fine
    try:
        float(value)
        return True
    except (TypeError, ValueError):
        return False


def _mark_valid(ws, row_idx: int, status_col: int, error_col: int, header_row: list[str]) -> None:
    """
    Clears any red fill/error text and marks the row as validated -- but
    deliberately NOT "DONE". DONE is reserved for "successfully loaded to
    QAD" (see price_list_load.py's skip-if-DONE logic); marking a merely-
    validated row DONE here would make Load silently skip it without ever
    calling the API.
    """
    for col_idx in range(1, len(header_row) + 1):
        ws.cell(row=row_idx, column=col_idx).fill = CLEAR_FILL
    ws.cell(row=row_idx, column=status_col, value="READY")
    ws.cell(row=row_idx, column=error_col, value="")


def _check_mandatory(row_data: dict) -> list[str]:
    missing = []
    for col in MANDATORY_COLUMNS:
        v = row_data.get(col)
        if v is None or str(v).strip() == "" or str(v).strip().lower() == "none":
            missing.append(col)
    return missing


def _check_lengths(row_data: dict) -> list[str]:
    bad = []
    for col, max_len in MAX_LENGTHS.items():
        v = row_data.get(col)
        if v is not None and len(str(v).strip()) > max_len:
            bad.append(col)
    return bad


def _check_numeric(row_data: dict) -> list[str]:
    bad = []
    for col in NUMERIC_COLUMNS:
        if not _is_numeric(row_data.get(col)):
            bad.append(col)
    return bad


def validate(file_path: str) -> dict:
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
    error_col = header_row.index("Error") + 1

    rows_passed = 0
    rows_failed = 0
    rows_skipped = 0

    for row_idx, row in enumerate(ws.iter_rows(min_row=3), start=3):
        row_values = [cell.value for cell in row]

        if not any(row_values):
            continue

        row_data = dict(zip(header_row, row_values))

        status = str(row_data.get("Status", row_data.get("Status ", ""))).strip().upper()
        data_operation = str(row_data.get("Data Operation", "")).strip().upper()

        # Same convention as Customer_load.py: a previously-DONE Create row

        if status == "DONE" and data_operation == "C":
            rows_skipped += 1
            continue

        errors: list[str] = []
        bad_columns: list[str] = []

        # ── Data Operation validity ────────────────────────────────────────
        if data_operation not in {"C", "U"}:
            bad_columns.append("Data Operation")
            errors.append("Data Operation must be C or U")

        # ── Mandatory fields ────────────────────────────────────────────────
        missing = _check_mandatory(row_data)
        if missing:
            bad_columns.extend(missing)
            errors.append(f"Missing mandatory fields: {', '.join(missing)}")

        # ── Length constraints ──────────────────────────────────────────────
        too_long = _check_lengths(row_data)
        if too_long:
            bad_columns.extend(too_long)
            errors.append(f"Exceeds max length: {', '.join(too_long)}")

        # ── Numeric constraints ─────────────────────────────────────────────
        non_numeric = _check_numeric(row_data)
        if non_numeric:
            bad_columns.extend(non_numeric)
            errors.append(f"Must be numeric: {', '.join(non_numeric)}")

        # -- drop down validators----
        amount_type = (str(row_data.get("Amount Type")).strip() if row_data.get("Amount Type") is not None else "")

        if amount_type:
            if amount_type.lower() not in {
                x.lower() for x in VALID_AMOUNT_TYPES
            }:
                bad_columns.append("Amount Type")
                errors.append(
                    f"Invalid Amount Type: {amount_type}"
                )
        
        quantity_type = str(row_data.get("Quantity Type", "")).strip()

        if quantity_type:
            if quantity_type.lower() not in {
                x.lower() for x in VALID_QUANTITY_TYPES
            }:
                bad_columns.append("Quantity Type")
                errors.append(
                    f"Invalid Quantity Type: {quantity_type}"
                )

        combination_type = str(row_data.get("Combination Type", "")).strip()

        if combination_type:
            if combination_type.lower() not in {
                x.lower() for x in VALID_COMBINATION_TYPES
            }:
                bad_columns.append("Combination Type")
                errors.append(
                    f"Invalid Combination Type: {combination_type}"
                )

        # ── Date constraints ─────────────────────────────────────────────────
        start_date = _parse_date(row_data.get("Start Date"))
        expire_date = _parse_date(row_data.get("Expire Date"))

        if row_data.get("Start Date") and start_date is None:
            bad_columns.append("Start Date")
            errors.append(f"Start Date is not a valid date (expected {DATE_FORMAT})")

        if row_data.get("Expire Date") and expire_date is None:
            bad_columns.append("Expire Date")
            errors.append(f"Expire Date is not a valid date (expected {DATE_FORMAT})")

        # Start Date cannot be before today -- CREATE rows only. Update rows
        # have Start Date locked (it's part of the existing record's key),
        # so re-validating an already-past Start Date on a fetched record
        # would fail every re-run for no reason.
        if data_operation == "C" and start_date is not None:
            if start_date.date() < datetime.now().date():
                bad_columns.append("Start Date")
                errors.append("Start Date cannot be before today")

        # Expire Date cannot be before Start Date -- applies to both C and U,
        # since Expire Date is editable on Update (that's how you expire an
        # existing price list).
        if start_date is not None and expire_date is not None:
            if expire_date.date() < start_date.date():
                bad_columns.append("Expire Date")
                errors.append("Expire Date cannot be before Start Date")

        if errors:
            rows_failed += 1
            mark_error(
                ws, row_idx, status_col, error_col, header_row,
                bad_columns, "; ".join(errors),
            )
        else:
            rows_passed += 1
            _mark_valid(ws, row_idx, status_col, error_col, header_row)

    wb.save(file_path)

    return {
        "rows_passed":  rows_passed,
        "rows_failed":  rows_failed,
        "rows_skipped": rows_skipped,
        "has_errors":   rows_failed > 0,
    }

def run_standalone():

    import sys
    import os

    ROOT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    sys.path.append(ROOT_DIR)

    from config import CONFIG

    folder = os.path.abspath(
        os.path.join(ROOT_DIR, CONFIG["folders"]["price_list"])
    )

    if not os.path.exists(folder):
        print(f"ERROR: Folder not found: {folder}")
        sys.exit(1)

    xlsx_files = [
        f for f in os.listdir(folder)
        if f.endswith(".xlsx") and not f.startswith("~$")
    ]

    if not xlsx_files:
        print(f"ERROR: No .xlsx file found in folder: {folder}")
        sys.exit(1)

    total_failed = 0

    for file_name in xlsx_files:

        file_path = os.path.join(folder, file_name)

        print(f"Validating: {file_path}")

        result = validate(file_path)

        print(result)

        total_failed += result["rows_failed"]

    sys.exit(0 if total_failed == 0 else 1)


if __name__ == "__main__":

    if len(sys.argv) > 1:
        result = validate(sys.argv[1])
        print(result)
    else:
        run_standalone()