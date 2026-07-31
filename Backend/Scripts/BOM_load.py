"""
BOM_load.py
-----------
Loads Product Structure / Formula (BOM) data into QAD via the
boms / bomComponentTreeNodes / bomComponents-fieldChange APIs.

Header "Update" support: a BOM_Header row's Data Operation can be "C"
(Create — default, unchanged behaviour) or "U" (Update). "Update" does NOT
rewrite header fields — Product Structure/BOM has no bulk header-overwrite
API, matching QAD's own UI, where a BOM is built up component-by-component
rather than overwritten wholesale. "Update" means: this BOM already exists
in QAD, skip POST /boms entirely, and just GET /boms to confirm it's
really there before treating it as ready for new components. This is what
makes "add more components to a BOM that already exists" possible: point
BOM_Header at it with Data Operation=U, then list only the NEW rows you
want to add in BOM_Components.

Component rows themselves remain CREATE-only — there's still no way to
edit an existing component's fields, only add new ones.

REFACTOR (this version) — scalable two-sheet architecture
-------------------------------------------------------------------------------
Earlier versions of this loader used one worksheet per hierarchy depth
(Level1, Level2, Level3, Level4), each with a ParentID column the user had
to fill in by hand. That capped the loader at 4 levels deep and forced a
user adding a 5th level to edit the code, not just the workbook.

This version replaces that with:

    BOM_Header       one row per root BOM/item (unchanged).
    BOM_Components   ONE sheet, one row per component, at ANY depth.

The old Level1..Level4 sheets and the ParentID column are GONE. In their
place:

  * Level (user-entered)  — a plain integer >= 1. 1 = attaches directly
    to the root BOM. 2 = attaches under a Level 1 component. N = attaches
    under a Level (N-1) component. There is no ceiling on N other than
    whatever QAD itself supports — 5, 50, or 999 levels all work with zero
    code changes, because nothing here is keyed off a fixed set of sheets
    or a fixed set of "known" levels.
  * ParentID (derived, INTERNAL ONLY — never a worksheet column) — QAD
    still wants a parentID on every initialize/save call, and that value
    is still just a fixed constant determined purely by level:
        Level 1  -> ParentID = -1
        Level N  -> ParentID = N - 1   (N > 1)
    The user never sees or enters this. It's computed on the fly from the
    row's own Level value at the moment we build the QAD payload.
  * ParentComponentCode (user-entered, unchanged in meaning) — the actual
    link between rows. For a Level 1 row this MUST equal RootBOMCode (the
    row attaches straight to the root item). For a Level N row (N > 1) it
    is the ComponentCode of the specific Level (N-1) row this one attaches
    under. It is used directly, as typed, as the "bomCode" parameter QAD
    expects on the initialize call for this row — no lookup or fetch
    needed to resolve it.
  * We never need a parent row's REAL, server-assigned treeListUniqueID for
    anything. The parentID VALUE we send to QAD is always just the fixed
    level constant (-1, 1, 2, 3, ... N-1), derived from this row's own
    Level — never fetched, never looked up.
  * The only thing genuinely unpredictable — and the only thing actually
    fetched from QAD per row — is the uniqueID for the NEW component being
    added (this row itself), via next_unique_id(). This value cannot be
    derived or guessed: existing items being attached can carry their own
    pre-existing sub-BOM in with them when saved, so the tree's occupied ID
    space does not grow in any way we control or can predict. We simply ask
    QAD (via a live tree fetch) for a value to propose on the initialize
    call, and never rely on it being "the same" ID afterward — nothing
    downstream needs to know the real ID QAD ultimately assigns.
  * Per-root bookkeeping is a simple set of (level, ComponentCode) pairs
    confirmed saved so far, used ONLY as a pre-flight sanity check — to
    catch a ParentComponentCode that doesn't correspond to any row we've
    actually processed at the level directly above (typo, wrong order,
    upstream failure, wrong RootBOMCode) — never to build any payload
    value.
  * That local set only knows about rows THIS run has saved (or rows
    already marked DONE in this sheet). It says nothing about components
    that already exist in QAD from a previous session — e.g. adding a new
    Level 2 row under a Level 1 component that was loaded weeks ago and
    isn't listed anywhere in today's BOM_Components sheet. To support that
    append case, if a parent isn't found in the local set, the loader
    falls back to ONE live GET of the actual QAD tree for that root
    (cached per root, not refetched per row) and checks the real tree's
    (level, componentCode) pairs before giving up and flagging an error.
    This makes "add more components to an existing BOM" work whether or
    not every one of that BOM's existing components happens to be listed
    in the sheet.

QAD is still the single source of truth for the tree's own internal node
identifiers (treeListUniqueID / treeListParentID) in the sense that we
never invent or persist one ourselves for later reuse; we just don't need
to read one back for parent linkage, because parent linkage is fully
described by (Level, ParentComponentCode) already sitting in the sheet.

Expected workbook shape
------------------------------------------------------------------------
    BOM_Header:
        DomainCode | BOMCode (=ItemCode) | Description | BatchSize |
        IsFormula (TRUE/FALSE) | Data Operation | Status | Error | Notes
        One row per root item/formula. CREATE only.

    BOM_Components:
        RootBOMCode | DomainCode (optional) | Level | ParentComponentCode |
        ComponentCode | QuantityPer | UM | StartEffectiveDate |
        EndEffectiveDate | IssuePolicy | PickPolicy | AllocationPolicy |
        ScrapPercent | Reference | Remarks | Status | Error | Notes
        One row per component, at ANY depth, for ANY root BOM. Rows are
        processed in ascending Level order within each RootBOMCode
        (parents before children) regardless of what order they were
        typed in the sheet.

        DomainCode here is OPTIONAL and only used as a fallback: if
        BOM_Header has no row for a given RootBOMCode, the loader looks
        for a DomainCode on that root's BOM_Components rows instead (all
        of them must agree). This is what lets a pure "add components to
        an existing BOM" workbook skip BOM_Header entirely — QAD needs a
        domain for every call it makes, and RootBOMCode alone doesn't
        supply one (the same BOM code can exist under more than one
        domain), so SOME row, somewhere, has to carry it.

    NOTE: header names are canonicalized (case/space-insensitive) when the
    sheet is read — see _canonicalize_header() — matching prior behaviour,
    so minor header typos/casing differences still round-trip correctly.
    Keep the canonical map in sync with any future header rename.

Business hierarchy fields — what the user actually fills in
-------------------------------------------------------------------------
    RootBOMCode          which BOM/root item this component belongs to.
    Level                plain integer >= 1. 1 = attaches directly to the
                          root item. N = attaches under a Level (N-1)
                          component. ParentID is DERIVED from this value
                          internally (Level 1 -> -1, Level N -> N-1); the
                          user never enters ParentID anywhere.
    ParentComponentCode  the ComponentCode of the row it attaches under.
                          For Level 1 rows this MUST equal RootBOMCode
                          (the root's own code — Level 1 rows attach
                          straight to the root item, there is no other
                          valid parent for them). For Level > 1 rows it is
                          the ComponentCode of the specific Level (N-1)
                          row this one attaches under. Used directly, as
                          typed, as the "bomCode" parameter QAD expects on
                          the initialize call for this row.
    ComponentCode        the new component being added.

Worked example (RootBOMCode = Dummy10), same tree as before, now expressed
as three rows in ONE sheet instead of three rows spread across three
sheets:
    Level=1, ParentComponentCode=Dummy10, ComponentCode=01012
        -> attaches directly to the root item. ParentID=-1 sent to QAD
           (derived from Level=1), bomCode=Dummy10 sent to QAD (the
           root's own code, and ParentComponentCode correctly equals it).
    Level=2, ParentComponentCode=01012,   ComponentCode=1013
        -> attaches under the Level=1 row whose ComponentCode is 01012.
           ParentID=1 sent to QAD (derived from Level=2, i.e. Level-1;
           nothing to do with 01012's own identity). bomCode=01012 sent
           to QAD, taken directly from ParentComponentCode.
    Level=3, ParentComponentCode=1013,    ComponentCode=70002
        -> attaches under the Level=2 row whose ComponentCode is 1013.
           ParentID=2 sent to QAD (derived from Level=3). bomCode=1013.
    Level=4, ParentComponentCode=70002,   ComponentCode=...
        -> and so on, indefinitely. Nothing about level 4 vs level 40 is
           different to this loader — it's still just "ParentID = Level-1,
           bomCode = ParentComponentCode".

Note: whatever pre-existing sub-BOM an already-existing item (like 01012 or
1013 above) may already have in the system is irrelevant here and is never
inspected, tracked, or reasoned about by this loader — the sheet only
describes the NEW components being added, each attaching to a named parent
by ParentComponentCode, regardless of what else that parent already has.

A parent row must be processed before any row that references it. Rows are
no longer guaranteed parent-before-child by sheet order (they all live in
one sheet now, in whatever order the user typed them), so build_tree_plan()
explicitly sorts every root's rows by Level ascending (stable — preserves
original sheet order within a level) before processing.

Per-BOM sequence (replicates the exact UI flow captured via network trace)
----------------------------------------------------------------------------
    1. POST /boms                                  -> create BOM header
    2. GET  /boms                                  -> verify + refresh concurrencyHash
    For every component row, walked top-down (Level ascending, parents before children):
    3. GET  /bomComponentTreeNodes?initialize=true...   -> blank template row
    4. POST /bomComponents/fieldChange?fieldName=componentCode
    5. POST /bomComponents/fieldChange?fieldName=<remaining fields, batched>
    6. POST /bomComponentTreeNodes                      -> save row (operation=10)

ASSUMPTIONS — validate against a live sandbox before trusting this in production:
  * Only two fieldChange calls were observed in the captured trace: one isolated
    (componentCode alone), one batched (description + dates + qty + UM together).
    This loader reproduces that exact 2-call shape. If a real save 400s / silently
    drops fields, the likely fix is switching to one fieldChange call per field.
  * "operation": 10 on the final save request = insert. This is inferred from a
    single sample row in the trace, not confirmed against QAD documentation.
  * initialize's `uniqueID` query param is NOT something QAD generates and
    hands back in a way we can rely on later — `initialize` is a GET, and the
    response simply echoes the uniqueID sent in as `treeListUniqueID`. It is
    computed here via next_unique_id(), a purely LOCAL incrementing counter
    seeded from the current time — NOT fetched from QAD's live tree (an
    earlier version of this file tried that via a plain GET of
    bomComponentTreeNodes, which 400'd on every live attempt and was
    removed; see the REFACTOR note ahead of the COMPONENT TREE section). It
    is NEVER read from the worksheet, and the worksheet has no column for
    it. This value is fundamentally a proposal, not a prediction: attaching
    an existing item can pull its own pre-existing sub-BOM into the tree as
    a side effect of the save, so the tree's real ID space is NOT something
    we can track or predict client-side regardless of how the value is
    sourced. Nothing downstream of a save depends on this value having
    "worked out" — we never look it up again.
  * `componentType=ITEM` appears on every captured initialize call for
    level >= 2, but not on the Level=1 (root-attach) call. This loader sends
    componentType="ITEM" whenever level >= 2. ASSUMPTION: not yet confirmed
    whether componentType is ever something other than "ITEM".
  * `newlyAddedComponent` — traces are inconsistent (sometimes "null",
    once seen equal to that same request's bomCode value) and there isn't
    enough data yet to derive the real rule. Left as "null" for now.
  * BOM_Header sheet in the actual template has no UM column (UM was
    previously assumed mandatory at header level). Header create payload
    now sends an empty UM string when the column is absent; ASSUMPTION —
    needs confirming against a live sandbox — is that QAD defaults the
    BOM's UM from the item master when the item already exists there. If
    header creates start failing on UM, this needs a UM column reinstated
    (or the loader needs to look the item's UM up from elsewhere first).
  * Validation (validate_bom.py) is intentionally NOT implemented yet — this
    file assumes rows without a BOM_Header row for their RootBOMCode, or with
    obviously missing mandatory fields, are the only pre-flight checks needed
    for now.
  * There is no pre-flight "does ParentComponentCode already exist in QAD"
    check anymore. A prior version of this file tried to add one — falling
    back to a live GET of bomComponentTreeNodes when a parent wasn't found
    in this run's local bookkeeping — to support appending under a
    component that already existed in QAD from an earlier session. That
    GET 400'd on every live attempt (unconfirmed request shape) and was
    removed; a bad ParentComponentCode now simply surfaces as a normal QAD
    save error on that row instead of being caught pre-flight.

Same conventions as every other <Entity>_load.py in this tool:
  * TokenManager        -> one token per run, refreshed only on 401
  * excel_format_utils  -> mark_done / mark_error for Status + Error columns
  * qad_response_utils  -> normalize_response for QAD's two error envelope shapes
"""

import os
import sys
import time
import requests
import openpyxl
from datetime import datetime, timezone

# ── Progress callback (used by main.py SSE streaming) ──────────────────────

_progress_callback = None

def set_progress_callback(callback):
    global _progress_callback
    _progress_callback = callback


# ── Path / config ───────────────────────────────────────────────────────────
ROOT_DIR    = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT_DIR)
sys.path.insert(0, SCRIPTS_DIR)
from config import CONFIG
from excel_format_utils import mark_done, mark_error, resolve_qad_errors
from qad_response_utils import normalize_response

# ── Sheet names ──────────────────────────────────────────────────────────────
HEADER_SHEET     = "BOM_Header"
COMPONENTS_SHEET = "BOM_Components"

# ── Mandatory columns ────────────────────────────────────────────────────────
# NOTE: the actual template's BOM_Header sheet has no UM column — see the
# UM ASSUMPTION in the module docstring. Do not add "UM" back here unless
# the sheet also gets the column back, or every header row will be
# incorrectly flagged as missing a field the user has no way to fill in.
MANDATORY_HEADER_COLUMNS = ["DomainCode", "BOMCode (=ItemCode)", "Description"]

# Level (user-entered hierarchy depth) + ParentComponentCode (the real
# link) are both mandatory business-hierarchy fields. ParentID is NOT in
# this list because it no longer exists as a worksheet column at all — it
# is derived internally from Level, never entered by the user.
MANDATORY_COMPONENT_COLUMNS = ["RootBOMCode", "Level", "ParentComponentCode", "ComponentCode", "QuantityPer", "UM"]

# ── QAD API fieldName -> Excel column name (component rows) ────────────────
# Extend whenever a new fieldName is observed in errors[]. Anything not in
# this map still shows up in the Error message text, it just won't get its
# own highlighted cell.
QAD_FIELD_TO_COLUMN = {
    "componentCode":     "ComponentCode",
    "description":       "Description",
    "quantityPer":       "QuantityPer",
    "unitOfMeasure":     "UM",
    "startEffectiveDate": "StartEffectiveDate",
    "endEffectiveDate":   "EndEffectiveDate",
    "issuePolicy":        "IssuePolicy",
    "pickPolicy":          "PickPolicy",
    "allocationPolicy":    "AllocationPolicy",
    "scrapPercent":        "ScrapPercent",
    "reference":           "Reference",
    "remarks":             "Remarks",
}

QAD_HEADER_FIELD_TO_COLUMN = {
    "bomCode":                "BOMCode (=ItemCode)",
    "description":            "Description",
    "batchSize":              "BatchSize",
    "itemStatus":             "ItemStatus",
    "isFormula":              "IsFormula (TRUE/FALSE)",
    "quantityCompleteMethod": "QuantityCompleteMethod",
}

# ── Header canonicalization for the BOM_Components sheet ───────────────────
# Maps a lowercased/space-stripped header cell to the canonical field name
# the rest of this file uses. Tolerates header typos/casing without needing
# the sheet fixed by hand. ParentID is intentionally absent — it is not a
# worksheet column in this architecture.
_CANONICAL_COMPONENT_HEADERS = {
    "rootbomcode":          "RootBOMCode",
    "domaincode":           "DomainCode",
    "level":                "Level",
    "parentcomponentcode":  "ParentComponentCode",
    "componentcode":        "ComponentCode",
    "quantityper":          "QuantityPer",
    "um":                   "UM",
    "starteffectivedate":   "StartEffectiveDate",
    "endeffectivedate":     "EndEffectiveDate",
    "issuepolicy":          "IssuePolicy",
    "pickpolicy":           "PickPolicy",
    "allocationpolicy":     "AllocationPolicy",
    "scrappercent":         "ScrapPercent",
    "reference":            "Reference",
    "remarks":              "Remarks",
    "status":               "Status",
    "error":                "Error",
    "notes":                "Notes",
}


def _canonicalize_header(raw_name: str) -> str:
    """Map a raw header cell to its canonical field name, tolerating case
    and spacing differences. Falls back to the stripped original if it
    isn't a recognized header, so unexpected columns (like 'Notes') still
    round-trip untouched."""
    key = str(raw_name).strip().lower().replace(" ", "") if raw_name is not None else ""
    return _CANONICAL_COMPONENT_HEADERS.get(key, str(raw_name).strip() if raw_name is not None else "")


# =============================================================================
# 1. AUTHENTICATION  (identical pattern to Customer_load.py)
# =============================================================================

def _fetch_token() -> str:
    url  = f"{CONFIG['qad']['base_url']}/oauth/token"
    resp = requests.post(url, data=CONFIG["qad"]["auth"], timeout=30)
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


class _TokenExpired(Exception):
    pass


def _call(method: str, url: str, tm: TokenManager, **kwargs) -> requests.Response:
    """
    Single request with the same 'refresh only on 401, retry once' policy
    used everywhere else in this tool. Raises _TokenExpired if a refreshed
    token still gets a 401 (caller decides what that means for the row).
    """
    headers = kwargs.pop("headers", {})
    headers["Authorization"] = f"Bearer {tm.get()}"

    resp = requests.request(method, url, headers=headers, timeout=30, **kwargs)

    if resp.status_code == 401:
        tm.refresh()
        headers["Authorization"] = f"Bearer {tm.get()}"
        resp = requests.request(method, url, headers=headers, timeout=30, **kwargs)
        if resp.status_code == 401:
            raise _TokenExpired()

    return resp


def _now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.%f")[:-3] + "Z"


def _excel_date_to_iso(v) -> str | None:
    """Excel gives back real datetime objects for date cells; QAD wants
    full ISO-8601 with a midnight-UTC time component, matching the trace
    (e.g. '2026-07-28T00:00:00.000Z')."""
    if v in (None, ""):
        return None
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%dT00:00:00.000Z")
    # fall back: string already in Excel cell, try a couple of common formats
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d-%b-%Y"):
        try:
            return datetime.strptime(str(v).strip(), fmt).strftime("%Y-%m-%dT00:00:00.000Z")
        except Exception:
            continue
    return None


# =============================================================================
# 2. BOM HEADER  (POST /boms, GET /boms)
# =============================================================================

def build_header_payload(row: dict) -> dict:
    def val(col):
        v = row.get(col)
        return str(v).strip() if v is not None else ""

    def int_val(col, default=1):
        v = val(col)
        try:
            return int(float(v)) if v else default
        except Exception:
            return default

    def bool_val(col):
        return val(col).strip().lower() in ("true", "yes", "1")

    domain = val("DomainCode")
    code   = val("BOMCode (=ItemCode)")
    uri    = f"urn:be:com.qad.engineering.productstructure.IBom:{domain}.{code}"

    return {
        "asOfDate": _now_iso(),
        "bomCode": "",
        "boms": [{
            "uri":                    uri,
            "domainCode":             domain,
            "bomCode":                code,
            "itemCode":               code,
            "description":            val("Description"),
            "itemDescription":        val("Description"),
            "UM":                     val("UM"),  # usually blank — see UM ASSUMPTION in module docstring
            "batchSize":              int_val("BatchSize", 1),
            "itemStatus":             val("ItemStatus") or "ACTIVE",
            "isFormula":              bool_val("IsFormula (TRUE/FALSE)"),
            "quantityCompleteMethod": val("QuantityCompleteMethod") or "1",
            "dataOperation":          "CREATE",
        }],
    }


def post_bom_header(row: dict, tm: TokenManager) -> tuple[bool, list, dict]:
    """POST /boms. Returns (success, errors, created_bom_dict)."""
    payload = build_header_payload(row)
    url = (
        f"{CONFIG['qad']['base_url']}/api/erp/boms"
        f"?viewUri=urn:be:com.qad.engineering.productstructure.IBom"
    )
    resp = _call("POST", url, tm, json=payload)

    try:
        resp_json = resp.json()
    except Exception:
        resp_json = {}

    success, errors = normalize_response(resp_json, resp.status_code)
    created = {}
    if success:
        boms = resp_json.get("boms") or resp_json.get("data", {}).get("boms") or []
        if boms:
            created = boms[0]
    return success, errors, created


def get_bom_header(domain: str, bom_code: str, tm: TokenManager) -> dict:
    """GET /boms — used to verify/refresh concurrencyHash after create."""
    url = f"{CONFIG['qad']['base_url']}/api/erp/boms"
    resp = _call("GET", url, tm, params={"domainCode": domain, "bomCode": bom_code})
    try:
        resp_json = resp.json()
    except Exception:
        return {}
    boms = resp_json.get("data", {}).get("boms") or resp_json.get("boms") or []
    return boms[0] if boms else {}


# =============================================================================
# 3. COMPONENT TREE  (initialize GET, fieldChange POST, save POST)
# =============================================================================
# NOTE: there is deliberately no "fetch the live tree first" step anywhere
# in this section anymore. An earlier version of this file tried to add a
# pre-flight "does this parent really exist in QAD" check and a "read the
# tree to compute the true next uniqueID" step, both built on a plain GET
# of bomComponentTreeNodes. That endpoint call 400'd on every live attempt
# (wrong request shape, unconfirmed against any real trace) and neither
# addition was ever part of the originally proven-working flow: adding a
# component to an existing BOM only ever needed Level (-> derived
# ParentID), ParentComponentCode (-> bomCode), RootBOMCode, DomainCode, and
# a proposed uniqueID — see next_unique_id() below and the module
# docstring. If a ParentComponentCode genuinely doesn't exist in QAD, the
# save call itself returns a real QAD error, which flows through the
# normal mark_error() path just like any other row-level failure — no
# separate pre-check is needed to catch that case.

def next_unique_id() -> int:
    """
    Propose a uniqueID for the next initialize call, for the row currently
    being created. No network call — this is intentionally a LOCAL,
    monotonically increasing counter, not something read back from QAD.

    This is a proposal, not a prediction: QAD's real, server-assigned
    treeListUniqueID for a new row is fundamentally unpredictable from the
    client side — e.g. attaching an existing item can pull that item's own
    pre-existing sub-BOM into the tree as a side effect of the save. Per
    the module docstring, nothing downstream of a save ever reads this
    value back or depends on it having "worked out" as expected — parent
    linkage for subsequent rows is fully determined by ParentComponentCode
    + the Level-derived ParentID constant already in the sheet, not by any
    ID assigned here. A simple incrementing counter, seeded from the
    current time to make collisions with any pre-existing tree IDs
    unlikely, is sufficient and needs no bomComponentTreeNodes GET at all.
    """
    global _unique_id_counter
    _unique_id_counter += 1
    return _unique_id_counter


# Seeded once per process, not per row/root — a plain incrementing counter
# is all that's needed since the value is a proposal only (see docstring
# above), and seeding from the current time keeps it clear of low integers
# QAD's own tree is likely to already be using.
_unique_id_counter = int(time.time()) % 900000 + 100000


def initialize_tree_node(
    domain: str, root_bom_code: str, bom_code: str,
    parent_id: int, unique_id: int, level: int, tm: TokenManager,
) -> dict:
    """
    GET /bomComponentTreeNodes?initialize=true... -> blank template row.

    IMPORTANT: bom_code here means "which item's structure am I currently
    editing" — i.e. the PARENT item's code (root item's code for a Level 1
    row, or the specific parent component's ComponentCode for anything
    deeper). It is NOT the new component's own code, and it is NOT always
    root_bom_code. Confirmed from the captured trace: a Level 3 add used
    bomCode=<Level 2 parent's code>, not the tree's root code. This value
    comes straight from the sheet's ParentComponentCode column (or the
    root's own code, for a Level 1 row) — no lookup or fetch required to
    determine it.

    parent_id is the constant DERIVED from this row's Level (-1 for
    Level 1, N-1 for Level N) — never a real, server-assigned
    treeListUniqueID, and never looked up, and never read from the
    worksheet (there is no ParentID column). unique_id is a proposed fresh
    value for THIS new row — see next_unique_id().

    componentType=ITEM is sent for level >= 2, matching every captured
    trace at that depth (the root-attach / Level 1 call omits it). Not yet
    confirmed whether componentType is ever anything other than "ITEM".
    """
    url = f"{CONFIG['qad']['base_url']}/api/erp/bomComponentTreeNodes"
    params = {
        "initialize":  "true",
        "domainCode":  domain,
        "rootBomCode": root_bom_code,
        "asOfDate":    _now_iso(),
        "levels":      "0",
        "newlyAddedComponent": "null",
        "bomCode":     bom_code,
        "parentID":    parent_id,
        "uniqueID":    unique_id,
        "level":       level,
    }
    if level >= 2:
        params["componentType"] = "ITEM"

    resp = _call("GET", url, tm, params=params)
    try:
        resp_json = resp.json()
    except Exception:
        return {}
    nodes = resp_json.get("data", {}).get("bomComponentTreeNodes") or resp_json.get("bomComponentTreeNodes") or []
    return nodes[0] if nodes else {}


def field_change(row_state: dict, field_name: str, tm: TokenManager) -> dict:
    """POST /bomComponents/fieldChange?fieldName=X -> updated row state."""
    url = f"{CONFIG['qad']['base_url']}/api/erp/bomComponents/fieldChange"
    payload = {"bomComponents": [row_state]}
    resp = _call("POST", url, tm, params={"fieldName": field_name}, json=payload)
    try:
        resp_json = resp.json()
    except Exception:
        return row_state
    comps = resp_json.get("bomComponents") or []
    return comps[0] if comps else row_state


def save_tree_node(row_state: dict, tm: TokenManager) -> tuple[bool, list]:
    """POST /bomComponentTreeNodes -> persist the component row."""
    url = f"{CONFIG['qad']['base_url']}/api/erp/bomComponentTreeNodes"
    row_state = dict(row_state)
    row_state["asOfDate"] = _now_iso()
    row_state["operation"] = 10  # insert — see ASSUMPTIONS in module docstring
    payload = {"bomComponentTreeNodes": [row_state]}
    resp = _call("POST", url, tm, json=payload)

    try:
        resp_json = resp.json()
    except Exception:
        resp_json = {}

    return normalize_response(resp_json, resp.status_code)


def apply_row_fields(template: dict, node: dict) -> dict:
    """Overlay the Excel row's values onto the blank template returned by
    initialize_tree_node(), matching field names to QAD's schema."""
    row = dict(template)
    row["componentCode"]  = node["ComponentCode"]
    row["description"]    = node.get("Description") or ""
    row["quantityPer"]    = node.get("QuantityPer") or 0
    row["unitOfMeasure"]  = node.get("UM") or ""

    start = _excel_date_to_iso(node.get("StartEffectiveDate"))
    end   = _excel_date_to_iso(node.get("EndEffectiveDate"))
    if start:
        row["startEffectiveDate"] = start
    if end:
        row["endEffectiveDate"] = end

    if node.get("IssuePolicy") not in (None, ""):
        row["issuePolicy"] = str(node["IssuePolicy"]).strip()
    if node.get("PickPolicy") not in (None, ""):
        row["pickPolicy"] = str(node["PickPolicy"]).strip()
    if node.get("AllocationPolicy") not in (None, ""):
        row["allocationPolicy"] = str(node["AllocationPolicy"]).strip()
    if node.get("ScrapPercent") not in (None, ""):
        row["scrapPercent"] = node["ScrapPercent"]
    if node.get("Reference") not in (None, ""):
        row["reference"] = str(node["Reference"]).strip()
    if node.get("Remarks") not in (None, ""):
        row["remarks"] = str(node["Remarks"]).strip()

    return row


def load_one_component(
    domain: str, root_bom_code: str, bom_code: str,
    parent_id: int, unique_id: int, level: int,
    node: dict, tm: TokenManager,
) -> tuple[bool, list, dict]:
    """
    Full per-row sequence: initialize -> fieldChange(componentCode) ->
    fieldChange(rest, batched) -> save. Returns (success, errors, saved_row).

    parent_id here is always the constant derived from Level (-1 for
    Level 1, N-1 for Level N), and bom_code is always the parent's
    ComponentCode taken directly from the sheet (or the root's own code)
    — see initialize_tree_node()'s docstring.
    """
    template = initialize_tree_node(domain, root_bom_code, bom_code, parent_id, unique_id, level, tm)
    if not template:
        return False, [{"fieldName": None, "message": "initialize returned no template row", "fieldValue": None, "code": None}], {}

    # Step 1: isolated fieldChange for componentCode (matches trace)
    row_state = dict(template)
    row_state["componentCode"] = node["ComponentCode"]
    row_state = field_change(row_state, "componentCode", tm)

    # Step 2: batched fieldChange for the remaining fields (matches trace)
    row_state = apply_row_fields(row_state, node)
    row_state = field_change(row_state, "description", tm)

    # Step 3: save
    success, errors = save_tree_node(row_state, tm)
    return success, errors, row_state


# =============================================================================
# 4. TREE PLAN  — turn BOM_Components rows into a top-down walk order
# =============================================================================
# NOTE: no per-row ordinal/position numbering, and no ParentID column at
# all. A row's identity within a root's tree is (Level, ComponentCode) —
# same level never repeats a ComponentCode under the same root. ParentID
# sent to QAD is always derived on the fly as (Level - 1), or -1 for
# Level 1 — never read from the worksheet, never used as a lookup key.
# ParentComponentCode is the real link and is used directly.

def _read_components_sheet(ws) -> list[str]:
    """Return the canonical header row for BOM_Components, adding
    Status/Error columns if the sheet doesn't already have them."""
    raw_header_row = [c.value for c in ws[1]]
    header_row = [_canonicalize_header(v) for v in raw_header_row]

    for col_name in ("Status", "Error"):
        if col_name not in header_row:
            ws.cell(row=1, column=len(header_row) + 1, value=col_name)
            header_row.append(col_name)

    return header_row


def _sort_key_level(entry: dict):
    """Sort key for build_tree_plan(): parse Level permissively so a bad
    value doesn't crash the sort — rows with an unparsable Level simply
    sort last and get caught by _parse_level()'s real validation later,
    per-row, with a proper error message."""
    try:
        return int(float(str(entry.get("Level")).strip()))
    except (TypeError, ValueError):
        return 10**9


def build_tree_plan(wb) -> dict[str, list[dict]]:
    """
    Reads BOM_Components and groups rows by RootBOMCode, each as a flat
    list ordered parent-before-child (Level ascending, then original sheet
    order within a level — stable sort). This ordering guarantees a parent
    row is always processed before any row whose ParentComponentCode
    references it, regardless of what order the user typed rows in the
    sheet, and regardless of how many distinct Level values are in use —
    unlike the old per-sheet-per-level model, there is no hardcoded ceiling
    here.

    Each entry: {sheet, row_idx, RootBOMCode, Level, ParentComponentCode,
    ComponentCode, ...fields}
    """
    plan: dict[str, list[dict]] = {}

    if COMPONENTS_SHEET not in wb.sheetnames:
        return plan

    ws = wb[COMPONENTS_SHEET]
    header_row = _read_components_sheet(ws)

    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        row_values = [cell.value for cell in row]
        if not any(row_values):
            continue
        row_data = dict(zip(header_row, row_values))
        root = str(row_data.get("RootBOMCode", "")).strip()
        if not root:
            continue
        entry = dict(row_data)
        entry["sheet"] = COMPONENTS_SHEET
        entry["row_idx"] = row_idx
        plan.setdefault(root, []).append(entry)

    for root in plan:
        plan[root].sort(key=_sort_key_level)  # stable — preserves sheet order within a level

    return plan


# =============================================================================
# 5. MANDATORY FIELD / LEVEL CHECKS
# =============================================================================

def _check_mandatory(row_data: dict, columns: list[str]) -> list[str]:
    missing = []
    for col in columns:
        v = row_data.get(col)
        if v is None or str(v).strip() == "" or str(v).strip().lower() == "none":
            missing.append(col)
    return missing


def _parse_level(entry: dict) -> tuple[int | None, str | None]:
    """Parse the Level cell as an integer >= 1. Level is the ONLY
    hierarchy-depth input the user provides; ParentID is derived from it
    internally and is never a worksheet column. Returns
    (value, error_message); value is None if the cell couldn't be parsed
    as a valid level at all."""
    raw = entry.get("Level")
    try:
        level = int(float(str(raw).strip()))
    except (TypeError, ValueError):
        return None, f"Level '{raw}' is not a valid integer"
    if level < 1:
        return None, f"Level '{raw}' must be >= 1"
    return level, None


# =============================================================================
# 6. SINGLE-FILE PROCESSOR
# =============================================================================

def process_file(file_path: str, tm: TokenManager) -> tuple[int, int]:
    """Open one workbook, process the header sheet then every component
    row in BOM_Components, return (ok_count, fail_count)."""

    wb = openpyxl.load_workbook(file_path)

    ok_count   = 0
    fail_count = 0

    # ── BOM_Header sheet: create every root BOM first ───────────────────────
    header_ws = wb[HEADER_SHEET] if HEADER_SHEET in wb.sheetnames else None
    header_by_root: dict[str, dict] = {}   # RootBOMCode -> created header info
    failed_roots: set[str] = set()

    if header_ws is not None:
        h_header_row = [str(c.value).strip() if c.value is not None else "" for c in header_ws[1]]
        for col_name in ("Status", "Error"):
            if col_name not in h_header_row:
                header_ws.cell(row=1, column=len(h_header_row) + 1, value=col_name)
                h_header_row.append(col_name)
        h_status_col = h_header_row.index("Status") + 1
        h_error_col  = h_header_row.index("Error") + 1

        for row_idx, row in enumerate(header_ws.iter_rows(min_row=2), start=2):
            row_values = [cell.value for cell in row]
            if not any(row_values):
                continue
            row_data = dict(zip(h_header_row, row_values))

            status = str(row_data.get("Status", "")).strip().upper()
            root = str(row_data.get("BOMCode (=ItemCode)", "")).strip()

            if status == "DONE":
                header_by_root[root] = row_data
                continue

            data_op_raw = str(row_data.get("Data Operation", "") or "").strip().upper()
            if data_op_raw in ("", "C", "CREATE"):
                is_update = False
            elif data_op_raw in ("U", "UPDATE"):
                is_update = True
            else:
                fail_count += 1
                failed_roots.add(root)
                mark_error(
                    header_ws, row_idx, h_status_col, h_error_col, h_header_row, ["Data Operation"],
                    f"Data Operation '{row_data.get('Data Operation')}' not recognized — use 'C' "
                    f"(Create) or 'U' (Update: BOM already exists in QAD, just attach new components)",
                )
                wb.save(file_path)
                continue

            missing_cols = list(MANDATORY_HEADER_COLUMNS)
            if is_update:
                # Update doesn't create anything, so Description isn't needed
                # to build a create payload — only enough to look the BOM up.
                missing_cols = ["DomainCode", "BOMCode (=ItemCode)"]
            missing = _check_mandatory(row_data, missing_cols)
            if missing:
                fail_count += 1
                failed_roots.add(root)
                mark_error(
                    header_ws, row_idx, h_status_col, h_error_col, h_header_row,
                    missing, f"Missing mandatory fields: {', '.join(missing)}",
                )
                wb.save(file_path)
                continue

            domain = str(row_data.get("DomainCode", "")).strip()

            if is_update:
                # "Update" = this BOM already exists in QAD. There is no
                # bulk header-overwrite API for Product Structure/BOM (see
                # module docstring), so we never POST here — we just GET to
                # confirm the header is really there, then let it act as
                # the anchor for any new BOM_Components rows against it.
                existing = {}
                errors: list = []
                for attempt in range(2):
                    try:
                        existing = get_bom_header(domain, root, tm)
                        break
                    except _TokenExpired:
                        if attempt == 0:
                            continue
                        errors = [{"fieldName": None, "message": "Token refresh failed — unauthorised", "fieldValue": None, "code": None}]
                        break
                    except requests.RequestException as e:
                        errors = [{"fieldName": None, "message": f"Network error: {e}", "fieldValue": None, "code": None}]
                        break

                if not existing:
                    fail_count += 1
                    failed_roots.add(root)
                    reason = errors[0]["message"] if errors else (
                        f"No BOM found in QAD for domain '{domain}', bomCode '{root}' — "
                        f"Update requires the BOM to already exist"
                    )
                    mark_error(header_ws, row_idx, h_status_col, h_error_col, h_header_row, [], reason)
                    wb.save(file_path)
                    time.sleep(0.1)
                    continue

                ok_count += 1
                header_by_root[root] = row_data
                mark_done(header_ws, row_idx, h_status_col, h_error_col)
                wb.save(file_path)
                time.sleep(0.1)
                continue

            # ── Create (default / unchanged behaviour) ──────────────────────
            success = False
            errors = []
            created: dict = {}

            for attempt in range(2):
                try:
                    success, errors, created = post_bom_header(row_data, tm)
                    break
                except _TokenExpired:
                    if attempt == 0:
                        continue
                    errors = [{"fieldName": None, "message": "Token refresh failed — unauthorised", "fieldValue": None, "code": None}]
                    break
                except requests.RequestException as e:
                    errors = [{"fieldName": None, "message": f"Network error: {e}", "fieldValue": None, "code": None}]
                    break

            if not success:
                fail_count += 1
                failed_roots.add(root)
                bad_cols, error_msg = resolve_qad_errors(errors, QAD_HEADER_FIELD_TO_COLUMN)
                mark_error(header_ws, row_idx, h_status_col, h_error_col, h_header_row, bad_cols, error_msg)
                wb.save(file_path)
                time.sleep(0.1)
                continue

            # verify / refresh concurrencyHash
            _ = get_bom_header(domain, root, tm)

            ok_count += 1
            header_by_root[root] = row_data
            mark_done(header_ws, row_idx, h_status_col, h_error_col)
            wb.save(file_path)
            time.sleep(0.1)

    # ── BOM_Components sheet: walk each root's tree top-down by Level ───────
    plan = build_tree_plan(wb)

    for root, entries in plan.items():

        if root in failed_roots:
            # header failed to create — skip every component under it and
            # mark why, rather than silently dropping them from the count.
            for entry in entries:
                ws = wb[entry["sheet"]]
                header_row = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
                status_col = header_row.index("Status") + 1
                error_col  = header_row.index("Error") + 1
                fail_count += 1
                mark_error(
                    ws, entry["row_idx"], status_col, error_col, header_row, [],
                    f"Skipped: BOM header for '{root}' failed to create",
                )
            wb.save(file_path)
            continue

        domain = str(header_by_root.get(root, {}).get("DomainCode", "")).strip()

        if not domain:
            # BOM_Header didn't have this root — fall back to a DomainCode
            # column on the component rows themselves. This is what lets an
            # append-to-an-existing-BOM workflow skip the header sheet
            # entirely: RootBOMCode alone isn't enough to call QAD (the
            # same BOM code can exist under different domains), but if
            # every row for this root agrees on a DomainCode, that's a
            # perfectly good substitute for a header row.
            component_domains = {str(e.get("DomainCode", "") or "").strip() for e in entries}
            component_domains.discard("")
            if len(component_domains) == 1:
                domain = next(iter(component_domains))
            elif len(component_domains) > 1:
                for entry in entries:
                    ws = wb[entry["sheet"]]
                    header_row = _read_components_sheet(ws)
                    status_col = header_row.index("Status") + 1
                    error_col  = header_row.index("Error") + 1
                    fail_count += 1
                    mark_error(
                        ws, entry["row_idx"], status_col, error_col, header_row, ["DomainCode"],
                        f"Conflicting DomainCode values for RootBOMCode '{root}' across "
                        f"BOM_Components rows ({', '.join(sorted(component_domains))}) — every "
                        f"row for the same root must agree on domain",
                    )
                wb.save(file_path)
                continue

        if not domain:
            # Neither a BOM_Header row nor any BOM_Components row for this
            # root supplied a DomainCode — there's genuinely nothing to
            # infer it from (QAD partitions BOMs by domain, so the same
            # RootBOMCode string isn't guaranteed unique without one).
            for entry in entries:
                ws = wb[entry["sheet"]]
                header_row = _read_components_sheet(ws)
                status_col = header_row.index("Status") + 1
                error_col  = header_row.index("Error") + 1
                fail_count += 1
                mark_error(
                    ws, entry["row_idx"], status_col, error_col, header_row, [],
                    f"No DomainCode found for RootBOMCode '{root}' — add a BOM_Header row for "
                    f"it, or fill in the DomainCode column on its BOM_Components rows",
                )
            wb.save(file_path)
            continue

        # (level, ComponentCode) of every row confirmed saved so far (this
        # run or a prior one). This is purely local bookkeeping — used only
        # so a DONE row's Level is on record for its own children within
        # THIS run. It is NOT used to gate/block a row from being sent to
        # QAD: an earlier version of this file required a parent to be
        # found here (falling back to a live tree GET if not) before
        # allowing a save, which broke the append-to-an-existing-BOM case
        # entirely whenever the parent wasn't listed in today's sheet — and
        # depended on a bomComponentTreeNodes GET that 400'd on every live
        # attempt. That gate has been removed. If ParentComponentCode
        # genuinely doesn't correspond to anything in QAD, the save call
        # itself returns a real QAD error, which is reported through the
        # normal mark_error() path below exactly like any other failure.
        saved: set[tuple[int, str]] = {(0, root)}

        total_rows = len(entries)
        for i, entry in enumerate(entries, start=1):
            if _progress_callback:
                _progress_callback(i, total_rows)

            ws = wb[entry["sheet"]]
            header_row = _read_components_sheet(ws)
            status_col = header_row.index("Status") + 1
            error_col  = header_row.index("Error") + 1

            status = str(entry.get("Status", "")).strip().upper()
            component_code = str(entry.get("ComponentCode", "")).strip()
            parent_component_code = str(entry.get("ParentComponentCode", "") or "").strip()

            level, parse_err = _parse_level(entry)

            if status == "DONE":
                # Row already saved on a prior run — just record it so any
                # child (processed in THIS run) can validate against it.
                # We still need a valid Level to know WHERE to record it;
                # if that can't be parsed for a DONE row, skip recording
                # rather than guessing (a stale/edited DONE row).
                if level is not None:
                    saved.add((level, component_code))
                continue

            missing = _check_mandatory(entry, MANDATORY_COMPONENT_COLUMNS)
            if missing:
                fail_count += 1
                mark_error(
                    ws, entry["row_idx"], status_col, error_col, header_row,
                    missing, f"Missing mandatory fields: {', '.join(missing)}",
                )
                wb.save(file_path)
                continue

            if parse_err:
                fail_count += 1
                mark_error(
                    ws, entry["row_idx"], status_col, error_col, header_row, ["Level"], parse_err,
                )
                wb.save(file_path)
                continue

            # Level 1 rows attach directly to the root item — the only
            # valid ParentComponentCode for a Level 1 row is the root's own
            # code. This is an explicit business rule, not just a
            # side-effect of the lookup below.
            if level == 1 and parent_component_code != root:
                fail_count += 1
                mark_error(
                    ws, entry["row_idx"], status_col, error_col, header_row, ["ParentComponentCode"],
                    f"Level 1 rows must have ParentComponentCode equal to RootBOMCode "
                    f"('{root}'); got '{parent_component_code}'",
                )
                wb.save(file_path)
                continue

            # ParentID sent to QAD: derived purely from Level, never from
            # the worksheet (there is no ParentID column anymore).
            #   Level 1  -> -1
            #   Level N  -> N - 1   (N > 1)
            derived_parent_id = -1 if level == 1 else level - 1

            # bomCode = the parent's own item code, taken directly from the
            # sheet (or the root's code, for a Level 1 row). Neither this
            # nor parentID needs any lookup or fetch. If parent_component_code
            # doesn't actually correspond to a real component in QAD, the
            # save call below returns a real QAD error for this row.
            parent_bom_code = root if level == 1 else parent_component_code

            # uniqueID: a local proposal only, no network call — see
            # next_unique_id()'s docstring for why this is unpredictable
            # anyway and why nothing downstream depends on it.
            new_unique_id = next_unique_id()

            success = False
            errors: list = []

            for attempt in range(2):
                try:
                    success, errors, _ = load_one_component(
                        domain, root, parent_bom_code,
                        derived_parent_id, new_unique_id, level,
                        entry, tm,
                    )
                    break
                except _TokenExpired:
                    if attempt == 0:
                        continue
                    errors = [{"fieldName": None, "message": "Token refresh failed — unauthorised", "fieldValue": None, "code": None}]
                    break
                except requests.RequestException as e:
                    errors = [{"fieldName": None, "message": f"Network error: {e}", "fieldValue": None, "code": None}]
                    break

            if not success:
                fail_count += 1
                bad_cols, error_msg = resolve_qad_errors(errors, QAD_FIELD_TO_COLUMN)
                mark_error(ws, entry["row_idx"], status_col, error_col, header_row, bad_cols, error_msg)
                wb.save(file_path)
                time.sleep(0.1)
                continue

            ok_count += 1
            saved.add((level, component_code))
            mark_done(ws, entry["row_idx"], status_col, error_col)
            wb.save(file_path)
            time.sleep(0.1)

    return ok_count, fail_count


# =============================================================================
# 7. ORCHESTRATOR
# =============================================================================

def run(folder_path: str) -> tuple[int, int]:
    folder = os.path.abspath(folder_path)

    if not os.path.exists(folder):
        raise RuntimeError(f"Folder not found: {folder}")

    xlsx_files = [
        f for f in os.listdir(folder)
        if f.endswith(".xlsx") and not f.startswith("~$")
    ]

    if not xlsx_files:
        raise RuntimeError(f"No .xlsx files found in: {folder}")

    tm = TokenManager()

    total_ok   = 0
    total_fail = 0

    for file_name in xlsx_files:
        file_path = os.path.join(folder, file_name)
        print(f"Loading : {file_path}")

        ok, fail   = process_file(file_path, tm)
        total_ok   += ok
        total_fail += fail

        print("\nLoad Summary")
        print("-" * 40)
        print(f"  Rows loaded successfully : {ok}")
        print(f"  Rows failed              : {fail}")
        print("  Result :", "ALL ROWS LOADED SUCCESSFULLY ✓" if fail == 0 else "COMPLETED WITH ERRORS — fix red rows and re-run")
        print()

    print("=" * 40)
    print(f"  Total loaded : {total_ok}")
    print(f"  Total failed : {total_fail}")
    print("=" * 40)

    return total_ok, total_fail


# =============================================================================
# 8. ENTRY POINT
# =============================================================================

if __name__ == "__main__":
    folder = os.path.abspath(
        os.path.join(ROOT_DIR, CONFIG["folders"]["product_structure"])
    )
    ok, fail = run(folder)
    sys.exit(0 if fail == 0 else 1)