"""
qad_response_utils.py
----------------------
Generic helpers for normalizing QAD API response shapes.

QAD has been observed returning at least two different error envelopes:

1. submitResult-wrapped (normal validation failures):
   {"submitResult": {"success": false, "errors": [{"fieldName": ..., "message": ...}]}}

2. Bare top-level errors (gateway / permission failures, e.g. 403):
   {"errors": [{"severity": 1, "code": "403", "message": "Access is denied", "fieldName": ""}]}

normalize_response() flattens both into one (success, errors) shape so
callers never need to know which envelope QAD decided to send back.

Entity-agnostic — belongs alongside excel_format_utils.py as a shared
util imported by every <Entity>_load.py script.
"""


def normalize_response(resp_json: dict, http_status: int) -> tuple[bool, list]:
    """
    Returns (success, errors). errors is always a list of dicts carrying
    at least 'fieldName' and 'message' keys, regardless of which envelope
    QAD used for this particular call.
    """
    submit = resp_json.get("submitResult")
    if submit is not None:
        if submit.get("success") is True:
            return True, []
        errors = submit.get("errors", [])
        if not errors:
            errors = [{
                "fieldName": None,
                "message":   f"HTTP {http_status} — submitResult.success was not True",
                "value":     None,
                "code":      None,
            }]
        return False, errors

    # No submitResult wrapper — bare top-level errors (e.g. 403 gateway errors)
    bare_errors = resp_json.get("errors")
    if bare_errors:
        return False, bare_errors

    return False, [{
        "fieldName": None,
        "message":   f"HTTP {http_status} — unrecognized response shape (no submitResult, no errors)",
        "value":     None,
        "code":      None,
    }]