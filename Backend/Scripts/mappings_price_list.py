"""
price_list_mappings.py

Shared mapping layer for:
- price_list_fetch.py
- validate_price_list.py
- price_list_load.py

Excel always uses business-friendly values.
QAD payloads use internal numeric codes.
"""

# =============================================================================
# AMOUNT TYPE
# =============================================================================

AMOUNT_TYPE_MAP = {
    "List Price":    "1",
    "Discount %":    "2",
    "Freight Terms": "6",
    "Freight List":  "7",
    "Credit Terms":  "8",
    "Discount Amt":  "9",
}

AMOUNT_TYPE_REVERSE_MAP = {
    v: k for k, v in AMOUNT_TYPE_MAP.items()
}


# =============================================================================
# QUANTITY TYPE
# =============================================================================

QUANTITY_TYPE_MAP = {
    "Quantity": "1",
    "Amount":   "2",
}

QUANTITY_TYPE_REVERSE_MAP = {
    v: k for k, v in QUANTITY_TYPE_MAP.items()
}


# =============================================================================
# COMBINATION TYPE
# =============================================================================

COMBINATION_TYPE_MAP = {
    "Combinable":      "2",
    "Base-Combinable": "3",
    "Exclusive":       "4",
}

COMBINATION_TYPE_REVERSE_MAP = {
    v: k for k, v in COMBINATION_TYPE_MAP.items()
}

# ------- Look up helpers, case - insensitive the values 

def get_amount_type_code(value):
    if value is None:
        return None
    
    value = str(value).strip().lower()

    lookup = {
        k.lower(): v 
        for k, v in AMOUNT_TYPE_MAP.items()
    }

    return lookup.get(value)

def get_quantity_type_code(value):
    if value is None:
        return None
    
    value = str(value).strip().lower()

    lookup = {
        k.lower(): v 
        for k, v in QUANTITY_TYPE_MAP.items()
    }

    return lookup.get(value)

def get_combination_type_code(value):
    if value is None:
        return None
    
    value = str(value).strip().lower()

    lookup = {
        k.lower(): v 
        for k, v in COMBINATION_TYPE_MAP.items()
    }

    return lookup.get(value)


# =============================================================================
# VALIDATION LISTS
# =============================================================================

VALID_AMOUNT_TYPES = set(AMOUNT_TYPE_MAP.keys())

VALID_QUANTITY_TYPES = set(QUANTITY_TYPE_MAP.keys())

VALID_COMBINATION_TYPES = set(COMBINATION_TYPE_MAP.keys())