import openpyxl
from openpyxl.styles import PatternFill

RED_FILL   = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
CLEAR_FILL = PatternFill(fill_type=None)

FIELD_RULES = {
    "Business Relation": {"max_len": 20,   "type": "character"},
    "Active":            {"max_len": None, "type": "logical"},
    "Name":              {"max_len": 36,   "type": "character"},
    "Search Name":       {"max_len": 28,   "type": "character"},
    "Address Type":      {"max_len": 20,   "type": "character"},
    "Address 1":         {"max_len": 36,   "type": "character"},
    "City":              {"max_len": 20,   "type": "character"},
    "State":             {"max_len": 4,    "type": "character"},
    "Postal Code":       {"max_len": 10,   "type": "character"},
    "Country":           {"max_len": 3,    "type": "character"},
    "Language":          {"max_len": 2,    "type": "character"},
    "Tax Zone":          {"max_len": 16,   "type": "character"},
}

ENTITY_COL    = "Business Relation"
VALID_LOGICAL = {"yes", "no"}


def validate(file_path):
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active

    header_row = [cell.value for cell in ws[1]]

    if "Status" not in header_row:
        ws.cell(row=1, column=len(header_row) + 1, value="Status")
        header_row.append("Status")
    if "Error" not in header_row:
        ws.cell(row=1, column=len(header_row) + 1, value="Error")
        header_row.append("Error")

    status_col_idx = header_row.index("Status") + 1
    error_col_idx  = header_row.index("Error")  + 1
    entity_col_idx = header_row.index(ENTITY_COL) + 1

    has_errors  = False
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

        for col_name, rule in FIELD_RULES.items():
            if col_name not in header_row:
                continue

            col_idx  = header_row.index(col_name) + 1
            value    = row_data.get(col_name)
            str_val  = str(value).strip() if value is not None else ""
            is_empty = str_val == "" or str_val.lower() == "none"

            if is_empty:
                row_errors.append(f"{col_name}: empty")
                error_cell_idxs.append(col_idx)
                continue

            if rule["type"] == "logical":
                if str_val.lower() not in VALID_LOGICAL:
                    row_errors.append(f"{col_name}: must be Yes/No (got '{str_val}')")
                    error_cell_idxs.append(col_idx)
                continue

            max_len = rule.get("max_len")
            if max_len and len(str_val) > max_len:
                row_errors.append(f"{col_name}: max {max_len} chars (got {len(str_val)})")
                error_cell_idxs.append(col_idx)

        for cell in ws[row_idx]:
            cell.fill = CLEAR_FILL

        if row_errors:
            has_errors  = True
            error_count += 1
            error_msg   = "; ".join(row_errors)

            ws.cell(row=row_idx, column=status_col_idx, value="")
            ws.cell(row=row_idx, column=entity_col_idx).fill = RED_FILL
            for cidx in error_cell_idxs:
                ws.cell(row=row_idx, column=cidx).fill = RED_FILL
            ws.cell(row=row_idx, column=error_col_idx, value=error_msg).fill = RED_FILL

        else:
            ws.cell(row=row_idx, column=status_col_idx, value="READY")
            ws.cell(row=row_idx, column=error_col_idx,  value="")

    wb.save(file_path)
    return has_errors, error_count

if __name__ == "__main__":
    import sys
    import os

    if len(sys.argv) < 2:
        # No argument given — auto-find .xlsx files in current directory
        xlsx_files = [f for f in os.listdir(".") if f.endswith(".xlsx")]
        if not xlsx_files:
            print("ERROR: No .xlsx file found. Usage: python validate_br.py <file.xlsx>")
            sys.exit(1)
        file_path = xlsx_files[0]
        print(f"No file specified — using: {file_path}")
    else:
        file_path = sys.argv[1]

    if not os.path.exists(file_path):
        print(f"ERROR: File not found: {file_path}")
        sys.exit(1)

    print(f"Validating: {file_path} ...")
    has_errors, error_count = validate(file_path)

    if has_errors:
        print(f"VALIDATION FAILED — {error_count} row(s) have errors. Check the file for red highlights.")
        sys.exit(1)
    else:
        print(f"VALIDATION PASSED — all rows are valid and marked READY.")
        sys.exit(0)