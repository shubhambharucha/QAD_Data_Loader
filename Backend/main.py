"""
main.py  —  QAD Data Loader FastAPI Backend
============================================
Endpoints
---------
POST /api/validate         { "entities": ["Supplier", "Customer", ...] }
POST /api/load              { "entities": ["Supplier", "Customer", ...] }
POST /api/fetch-price-list  { "filters": [...], "filename": "..." }
POST /api/blank-template     { "entities": ["Supplier", "Customer", "PriceList"] }
GET  /api/blank-template/entities
GET  /api/config
POST /api/save-config
POST /api/test-connection
POST /api/login              { "username": "...", "password": "...", "environment": "TEST"|"PROD" }
GET  /api/permissions?session_id=...
POST /api/log-activity       { "username", "environment", "action", "status", "details" }
GET  /health


File-rename logic (validate & load)
------------------------------------
  - Any file with ≥ 1 error row  → rename to  error_<original_name>
  - File previously named error_* but now 0 errors → rename back to <original_name>

Load only (on 0 errors):
  - Move file to  Data/Archive/<EntityFolder>/<original_name>_<timestamp>.xlsx

SSE event types emitted
-----------------------
  entity_start  { entity }
  file_start    { entity, file }
  progress      { entity, file, row, total }
  file_result   { entity, file, ok, fail, skipped, status, note }
  entity_result { entity, passed, failed, skipped }
  done          { message }
  error         { message }


Activity log
------------
Every login attempt and every validate/load/fetch/template action fired from
the frontend gets appended as one JSON object per line to
Backend/activity_log.jsonl — e.g.:

  {"timestamp": "...", "username": "jdoe", "environment": "TEST", "action": "login", "status": "success"}
  {"timestamp": "...", "username": "jdoe", "environment": "TEST", "action": "validate", "status": "started", "details": {"entities": ["Customer"]}}

This is intentionally append-only and best-effort — a logging failure never
blocks the actual operation.


Sessions & QAD access tokens
-----------------------------
Login no longer writes the QAD OAuth token to disk (there used to be a
token_response.json dropped in Backend/ for debugging — that's gone; delete
any existing copy of that file, and make sure it's git-ignored). Instead,
a successful login creates an in-memory session keyed by a random session_id,
which is handed back to the frontend and stored in sessionStorage alongside
the existing qad_authenticated/qad_username/qad_environment flags. The
access token itself never leaves the server. Sessions expire after
SESSION_TTL_SECONDS and are held only in process memory, so they don't
survive a server restart — that's intentional for a short-lived OAuth token;
if you need them to survive restarts/scale across workers, move SESSIONS to
something like Redis with the same TTL semantics rather than writing tokens
to disk.

The /api/permissions endpoint uses the session's stored token to call QAD's
qracore permissionState API once per entity (see ENTITY_URI/PERMISSIONS
config notes near _entity_permission_config below) and reports back a
simple {entity_id: bool} map that the frontend uses to lock/unlock module
cards.
"""

import asyncio
import importlib
import importlib.util
import io
import json
import os
import re
import shutil
import sys
import threading
import time
import traceback
import uuid
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any, AsyncGenerator

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel


# ── main.py lives inside Backend/ ────────────────────────────────────────────
BACKEND_DIR = os.path.abspath(os.path.dirname(__file__))         # .../QAD_data_loader/Backend
ROOT_DIR    = os.path.abspath(os.path.join(BACKEND_DIR, ".."))  # .../QAD_data_loader
SCRIPTS_DIR = os.path.join(BACKEND_DIR, "Scripts")
DATA_DIR    = os.path.join(ROOT_DIR, "Data")

sys.path.insert(0, ROOT_DIR)
sys.path.insert(0, BACKEND_DIR)
sys.path.insert(0, SCRIPTS_DIR)

from config import CONFIG  # noqa: E402  (must come after sys.path setup)

# ═════════════════════════════════════════════════════════════════════════════
# ENTITY REGISTRY
# Maps UI entity id → { validate_script, load_script, folder_key, archive_folder }
# ═════════════════════════════════════════════════════════════════════════════

ENTITY_MAP: dict[str, dict] = {
    "Supplier": {
        "validate_script": "validate_supplier",
        "load_script":     "Supplier_load",
        "folder_key":      "supplier",
        "archive_folder":  "Supplier",
    },
    "Customer": {
        "validate_script": "validate_customer",
        "load_script":     "Customer_load",
        "folder_key":      "customer",
        "archive_folder":  "Customer",
    },
    "GCM": {
        "validate_script": "validate_gcm",
        "load_script":     "GCM_load",
        "folder_key":      "GCM",
        "archive_folder":  "GCM",
    },
    "BR": {
        "validate_script": "validate_br",
        "load_script":     "BR_load",
        "folder_key":      "BR",
        "archive_folder":  "BR",
    },
    "Customer_Item": {
        "validate_script": "validate_cust_item",
        "load_script":     "cust_item",
        "folder_key":      "customer_item",
        "archive_folder":  "Customer_Item",
    },
    "ProductionOrder": {
        "validate_script": "validate_production_order",
        "load_script":     "production_order",
        "folder_key":      "production_order",
        "archive_folder":  "ProductionOrder",
    },
    "PurchaseOrder": {
        "validate_script": "validate_po",
        "load_script":     "po_load",
        "folder_key":      "purchase_order",
        "archive_folder":  "PurchaseOrder",
    },
    "SalesOrder": {
        "validate_script": "validate_so",
        "load_script":     "so_load",
        "folder_key":      "sales_order",
        "archive_folder":  "SalesOrder",
    },
    "Supplier_Item": {
        "validate_script": "validate_Supp_item",
        "load_script":     "Supp_item",
        "folder_key":      "supplier_item",
        "archive_folder":  "Supplier_Item",
    },
    "SupplierPriceList": {
        "validate_script": "validate_supp_price_lists",
        "load_script":     "supp_price_lists",
        "folder_key":      "supplier_price_list",
        "archive_folder":  "SupplierPriceList",
    },
    "PriceList": {
        "validate_script": "validate_price_list",
        "load_script":     "price_list_load",
        "folder_key":      "price_list",
        "archive_folder":  "Price_List",
    },
    "BOM": {
        "validate_script": "validate_bom",
        "load_script":     "bom_load",
        "folder_key":      "bom",
        "archive_folder":  "BOM",
    },
}

# ═════════════════════════════════════════════════════════════════════════════
# FASTAPI APP
# ═════════════════════════════════════════════════════════════════════════════

app = FastAPI(title="QAD Data Loader")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["Content-Disposition"],
)

# ═════════════════════════════════════════════════════════════════════════════
# REQUEST SCHEMAS
# ═════════════════════════════════════════════════════════════════════════════

class OperationRequest(BaseModel):
    entities: list[str]
    session_id: str | None = None   # used to re-check permissions server-side before running


class FetchFilterItem(BaseModel):
    field:       str            # e.g. "price_list", "start_date", "amount_type"
    operator:    str            # "equals" | "starts with" | "ends with" | "contains" |
                                 # "range" | "greater than" | "less than"
    value_from:  str | None = None
    value_to:    str | None = None   # only meaningful when operator == "range"


class FetchPriceListRequest(BaseModel):
    filters:  list[FetchFilterItem] = []
    filename: str


class BlankTemplateRequest(BaseModel):
    entities: list[str]


class LoginRequest(BaseModel):
    username:    str
    password:    str
    environment: str   # "TEST" | "PROD"


class ActivityLogRequest(BaseModel):
    username:    str | None = None
    environment: str | None = None
    action:      str
    status:      str | None = None
    details:     dict | None = None

# ═════════════════════════════════════════════════════════════════════════════
# HELPERS  —  SSE
# ═════════════════════════════════════════════════════════════════════════════

def sse(event: dict) -> str:
    """Encode a dict as a properly formatted SSE data line."""
    return f"data: {json.dumps(event)}\r\n\r\n"


def sse_comment() -> str:
    """SSE keepalive comment — prevents proxies from closing idle connections."""
    return ": keepalive\r\n\r\n"

# ═════════════════════════════════════════════════════════════════════════════
# HELPERS  —  ACTIVITY LOG
# ═════════════════════════════════════════════════════════════════════════════

ACTIVITY_LOG_PATH = os.path.join(BACKEND_DIR, "activity_log.jsonl")


def log_activity(entry: dict) -> None:
    """
    Append one JSON line to activity_log.jsonl. Best-effort — a failure here
    must never break the actual login/validate/load/fetch/template flow.
    """
    try:
        record = {"timestamp": datetime.now().isoformat(timespec="seconds")}
        record.update(entry)
        with open(ACTIVITY_LOG_PATH, "a", encoding="utf-8") as f:
            f.write(json.dumps(record) + "\n")
    except Exception as e:
        print(f"[main.py] Failed to write activity log: {e}")

# ═════════════════════════════════════════════════════════════════════════════
# HELPERS  —  SESSIONS (in-memory, holds the QAD OAuth token server-side only)
# ═════════════════════════════════════════════════════════════════════════════
#
# Nothing here ever touches disk. A session disappears on server restart or
# after SESSION_TTL_SECONDS of age, whichever comes first. That's a deliberate
# trade-off: the QAD access token is short-lived credential material, so the
# safest place for it is "nowhere persistent." If you need sessions to survive
# a restart or to be shared across multiple backend workers, swap this dict
# for something like Redis (same key -> {username, environment, access_token,
# created} shape, same TTL) rather than reintroducing a file on disk.

SESSIONS: dict[str, dict] = {}
SESSIONS_LOCK = threading.Lock()
SESSION_TTL_SECONDS = 8 * 60 * 60  # 8 hours


def _create_session(username: str, environment: str, access_token: str) -> str:
    session_id = uuid.uuid4().hex
    with SESSIONS_LOCK:
        SESSIONS[session_id] = {
            "username":     username,
            "environment":  environment,
            "access_token": access_token,
            "created":      time.time(),
        }
    return session_id


def _get_session(session_id: str) -> dict | None:
    if not session_id:
        return None
    with SESSIONS_LOCK:
        s = SESSIONS.get(session_id)
        if not s:
            return None
        if time.time() - s["created"] > SESSION_TTL_SECONDS:
            del SESSIONS[session_id]
            return None
        return s

# ═════════════════════════════════════════════════════════════════════════════
# HELPERS  —  PATHS & FILES
# ═════════════════════════════════════════════════════════════════════════════

def resolve_folder(entity_id: str) -> str:
    """Return absolute path to entity's data folder."""
    cfg      = ENTITY_MAP[entity_id]
    rel_path = CONFIG["folders"].get(cfg["folder_key"], "")
    if not rel_path:
        raise ValueError(f"No folder configured for entity '{entity_id}' (key: {cfg['folder_key']})")
    return os.path.abspath(os.path.join(BACKEND_DIR, rel_path))


def resolve_archive(entity_id: str) -> str:
    """Return absolute path to entity's Archive folder; creates it if needed."""
    archive_entity = ENTITY_MAP[entity_id]["archive_folder"]
    path = os.path.join(DATA_DIR, "Archive", archive_entity)
    os.makedirs(path, exist_ok=True)
    return path


def resolve_downloads_folder() -> str:
    """
    Return the current user's Downloads folder, creating it if missing.
    Works cross-platform (Windows/macOS/Linux) since it's just ~/Downloads
    on all three.
    """
    downloads = Path.home() / "Downloads"
    downloads.mkdir(parents=True, exist_ok=True)
    return str(downloads)


def list_xlsx(folder: str) -> list[str]:
    """Return sorted list of .xlsx filenames (skip temp files)."""
    if not os.path.isdir(folder):
        return []
    return sorted(
        f for f in os.listdir(folder)
        if f.endswith(".xlsx") and not f.startswith("~$")
    )


def load_module(script_name: str):
    """Dynamically import a script from Scripts/ by name (no .py extension)."""
    path = os.path.join(SCRIPTS_DIR, f"{script_name}.py")
    if not os.path.exists(path):
        raise FileNotFoundError(f"Script not found: {path}")
    spec   = importlib.util.spec_from_file_location(script_name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _safe_filename(folder: str, original: str) -> str:
    """
    Returns a collision-safe filename.
    If 'report.xlsx' exists → 'report (1).xlsx' → 'report (2).xlsx' etc.
    """
    stem, ext = os.path.splitext(original)
    candidate = original
    counter   = 1
    while os.path.exists(os.path.join(folder, candidate)):
        candidate = f"{stem} ({counter}){ext}"
        counter  += 1
    return candidate


def _generate_filename(prefix: str, custom: str | None) -> str:
    """
    Use caller-supplied name if provided (ensure .xlsx extension).
    Otherwise auto-generate: PriceList_20260616_143022.xlsx
    """
    if custom:
        return custom if custom.endswith(".xlsx") else f"{custom}.xlsx"
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{prefix}_{ts}.xlsx"

# ═════════════════════════════════════════════════════════════════════════════
# HELPERS  —  FILE RENAME / ARCHIVE
# ═════════════════════════════════════════════════════════════════════════════

def _error_name(filename: str) -> str:
    return f"error_{filename}" if not filename.startswith("error_") else filename


def _clean_name(filename: str) -> str:
    return re.sub(r"^error_", "", filename)


def rename_file_error(folder, filename):
    old_path = os.path.join(folder, filename)

    if not os.path.exists(old_path):
        return filename  # ✅ skip if file missing

    new_name = _error_name(filename)
    if new_name != filename:
        new_path = os.path.join(folder, new_name)
        try:
            os.rename(old_path, new_path)
        except FileNotFoundError:
            return filename  # ✅ ignore safely
        return new_name

    return filename



def rename_file_clean(folder: str, filename: str) -> str:
    old_path = os.path.join(folder, filename)

    if not os.path.exists(old_path):
        return filename  # ✅ skip safely

    new_name = _clean_name(filename)

    if new_name != filename:
        new_path = os.path.join(folder, new_name)
        try:
            os.rename(old_path, new_path)
        except FileNotFoundError:
            return filename  # ✅ prevent crash
        return new_name

    return filename



def archive_file(folder: str, filename: str, archive_dir: str) -> str:
    clean     = _clean_name(filename)
    stem, ext = os.path.splitext(clean)
    ts        = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest_name = f"{stem}_{ts}{ext}"
    dest_path = os.path.join(archive_dir, dest_name)
    shutil.move(os.path.join(folder, filename), dest_path)
    return dest_path

# ═════════════════════════════════════════════════════════════════════════════
# BLANK TEMPLATE GENERATION  —  Supplier / Customer / PriceList
# ═════════════════════════════════════════════════════════════════════════════
#
# All of the actual template-building logic (section spans, column headers,
# styling) lives in Scripts/blank_template.py, loaded dynamically below the
# same way every validate_<entity>.py / <entity>_load.py script is. Adding a
# new entity's template later only ever means editing that one file.
#
# Scripts/blank_template.py contract:
#   TEMPLATE_SPECS               -> dict of entity_id -> spec (source of truth)
#   supported_entities() -> list[str]
#   build(entity_id: str) -> io.BytesIO   (in-memory workbook, no disk writes)

# ═════════════════════════════════════════════════════════════════════════════
# VALIDATE STREAM GENERATOR
# ═════════════════════════════════════════════════════════════════════════════

async def validate_stream(entities: list[str], session_id: str | None = None) -> AsyncGenerator[str, None]:
    loop = asyncio.get_event_loop()
    perm_ctx = _resolve_permission_context(session_id)

    for entity_id in entities:

        if entity_id not in ENTITY_MAP:
            yield sse({"type": "error", "message": f"Unknown entity: {entity_id}"})
            continue

        cfg = ENTITY_MAP[entity_id]

        try:
            folder = resolve_folder(entity_id)
        except Exception as e:
            yield sse({"type": "error", "message": str(e)})
            continue

        yield sse({"type": "entity_start", "entity": entity_id})
        await asyncio.sleep(0)

        if not _entity_allowed(perm_ctx, entity_id):
            yield sse({"type": "error", "message": f"Access restricted: no permission for {entity_id}"})
            yield sse({"type": "entity_result", "entity": entity_id, "passed": 0, "failed": 0, "skipped": 0})
            continue

        files = list_xlsx(folder)

        if not files:
            yield sse({"type": "error", "message": f"No .xlsx files found in {folder}"})
            yield sse({"type": "entity_result", "entity": entity_id, "passed": 0, "failed": 0, "skipped": 0})
            continue

        try:
            mod = load_module(cfg["validate_script"])
        except Exception as e:
            yield sse({"type": "error", "message": f"Cannot load {cfg['validate_script']}: {e}"})
            continue

        entity_passed  = 0
        entity_failed  = 0
        entity_skipped = 0

        for filename in files:

            file_path = os.path.join(folder, filename)

            yield sse({"type": "file_start", "entity": entity_id, "file": filename})
            await asyncio.sleep(0)

            try:
                result: dict = await loop.run_in_executor(None, mod.validate, file_path)
            except Exception as exc:
                yield sse({
                    "type": "file_result", "entity": entity_id, "file": filename,
                    "ok": 0, "fail": 1, "skipped": 0, "status": "error", "note": str(exc),
                })
                entity_failed += 1
                rename_file_error(folder, filename)
                continue

            rows_passed  = result.get("rows_passed",  0)
            rows_failed  = result.get("rows_failed",  0)
            rows_skipped = result.get("rows_skipped", 0)
            has_errors   = result.get("has_errors",   False)

            if has_errors:
                entity_failed += 1
                rename_file_error(folder, filename)
                status = "failed"
            else:
                entity_passed += 1
                rename_file_clean(folder, filename)
                status = "passed"

            yield sse({
                "type": "file_result", "entity": entity_id, "file": filename,
                "ok": rows_passed, "fail": rows_failed, "skipped": rows_skipped,
                "status": status, "note": f"{rows_skipped} rows skipped" if rows_skipped else "",
            })
            await asyncio.sleep(0)

        yield sse({
            "type": "entity_result", "entity": entity_id,
            "passed": entity_passed, "failed": entity_failed, "skipped": entity_skipped,
        })
        await asyncio.sleep(0)

    yield sse({"type": "done", "message": "Validation complete"})

# ═════════════════════════════════════════════════════════════════════════════
# LOAD STREAM GENERATOR
# ═════════════════════════════════════════════════════════════════════════════

async def load_stream(entities: list[str], session_id: str | None = None) -> AsyncGenerator[str, None]:
    loop = asyncio.get_event_loop()
    perm_ctx = _resolve_permission_context(session_id)

    for entity_id in entities:

        if entity_id not in ENTITY_MAP:
            yield sse({"type": "error", "message": f"Unknown entity: {entity_id}"})
            continue

        cfg = ENTITY_MAP[entity_id]

        try:
            folder      = resolve_folder(entity_id)
            archive_dir = resolve_archive(entity_id)
        except Exception as e:
            yield sse({"type": "error", "message": str(e)})
            continue

        yield sse({"type": "entity_start", "entity": entity_id})
        await asyncio.sleep(0)

        if not _entity_allowed(perm_ctx, entity_id):
            yield sse({"type": "error", "message": f"Access restricted: no permission for {entity_id}"})
            yield sse({"type": "entity_result", "entity": entity_id, "passed": 0, "failed": 0, "skipped": 0})
            continue

        files = list_xlsx(folder)

        if not files:
            yield sse({"type": "error", "message": f"No .xlsx files found in {folder}"})
            yield sse({"type": "entity_result", "entity": entity_id, "passed": 0, "failed": 0, "skipped": 0})
            continue

        try:
            mod = load_module(cfg["load_script"])
        except Exception as e:
            yield sse({"type": "error", "message": f"Cannot load {cfg['load_script']}: {e}"})
            continue

        progress_queue: asyncio.Queue = asyncio.Queue()

        entity_passed  = 0
        entity_failed  = 0
        entity_skipped = 0

        for filename in files:

            file_path = os.path.join(folder, filename)

            yield sse({"type": "file_start", "entity": entity_id, "file": filename})
            await asyncio.sleep(0)

            result_holder: list = []

            def _run_process_file():
                try:
                    if hasattr(mod, "_tm"):
                        tm = mod._tm
                    else:
                        tm = mod.TokenManager()
                        mod._tm = tm

                    ok, fail = mod.process_file(file_path, tm)
                    print(f"[main.py] process_file done for {filename}: ok={ok} fail={fail}")

                    # ── Banking pass (Supplier only, for now) ───────────────
                    # MUST run before this file gets archived below, otherwise
                    # a workbook that passes the Suppliers sheet gets archived
                    # immediately and its Banking sheet is never touched.
                    if hasattr(mod, "process_banking_file"):
                        print(f"[main.py] running process_banking_file for {filename}...")
                        bank_ok, bank_fail = mod.process_banking_file(file_path, tm)
                        print(f"[main.py] process_banking_file done for {filename}: "
                              f"ok={bank_ok} fail={bank_fail}")
                        ok   += bank_ok
                        fail += bank_fail

                    result_holder.extend([ok, fail])
                except Exception as exc:
                    print(f"[main.py] EXCEPTION while processing {filename}: {exc}")
                    traceback.print_exc()
                    result_holder.extend([0, -1])
                    result_holder.append(str(exc))

            future = loop.run_in_executor(None, _run_process_file)

            row_counter = [0]

            if hasattr(mod, "set_progress_callback"):
                def _on_progress(row: int, total: int):
                    row_counter[0] = row
                    loop.call_soon_threadsafe(
                        progress_queue.put_nowait,
                        {"row": row, "total": total},
                    )
                mod.set_progress_callback(_on_progress)

            while not future.done():
                await asyncio.sleep(0.4)
                drained = False
                while not progress_queue.empty():
                    pev = progress_queue.get_nowait()
                    yield sse({
                        "type": "progress", "entity": entity_id, "file": filename,
                        "row": pev["row"], "total": pev["total"],
                    })
                    drained = True
                if not drained and row_counter[0] > 0:
                    yield sse({
                        "type": "progress", "entity": entity_id, "file": filename,
                        "row": row_counter[0], "total": row_counter[0],
                    })

            try:
                await future
            except Exception:
                pass

            if hasattr(mod, "set_progress_callback"):
                mod.set_progress_callback(None)

            if len(result_holder) >= 3 and result_holder[1] == -1:
                exc_msg = result_holder[2]
                yield sse({
                    "type": "file_result", "entity": entity_id, "file": filename,
                    "ok": 0, "fail": 1, "skipped": 0, "status": "error", "note": exc_msg,
                })
                entity_failed += 1
                rename_file_error(folder, filename)
                continue

            ok_count   = result_holder[0] if result_holder else 0
            fail_count = result_holder[1] if len(result_holder) > 1 else 1

            if fail_count > 0:
                entity_failed += 1
                rename_file_error(folder, filename)
                status = "failed"
                note   = f"{fail_count} rows failed"
            else:
                entity_passed += 1
                clean_filename = rename_file_clean(folder, filename)
                archive_path   = archive_file(folder, clean_filename, archive_dir)
                status = "passed"
                note   = f"Archived → {os.path.basename(archive_path)}"

            yield sse({
                "type": "file_result", "entity": entity_id, "file": filename,
                "ok": ok_count, "fail": fail_count, "skipped": 0,
                "status": status, "note": note,
            })
            await asyncio.sleep(0)

        yield sse({
            "type": "entity_result", "entity": entity_id,
            "passed": entity_passed, "failed": entity_failed, "skipped": entity_skipped,
        })
        await asyncio.sleep(0)

    yield sse({"type": "done", "message": "Load complete"})

# ═════════════════════════════════════════════════════════════════════════════
# FETCH  —  Price List extraction from QAD (Scripts/fetch_price_lists.py)
# ═════════════════════════════════════════════════════════════════════════════

@app.post("/api/fetch-price-list")
async def api_fetch_price_list(req: FetchPriceListRequest):
    filename = _generate_filename("PriceList", req.filename.strip())

    out_dir = resolve_downloads_folder()
    output_path = os.path.join(out_dir, filename)

    filters_payload = [f.model_dump() for f in req.filters]

    loop = asyncio.get_event_loop()

    def _run_fetch() -> dict:
        try:
            mod = load_module("fetch_price_lists")
        except Exception as e:
            return {"ok": False, "message": f"Cannot load fetch_price_lists: {e}"}

        try:
            if hasattr(mod, "TokenManager"):
                if hasattr(mod, "_tm"):
                    tm = mod._tm
                else:
                    tm = mod.TokenManager()
                    mod._tm = tm
                result = mod.fetch(filters_payload, output_path, tm)
            else:
                result = mod.fetch(filters_payload, output_path)
            return result
        except Exception as exc:
            traceback.print_exc()
            return {"ok": False, "message": str(exc)}

    result = await loop.run_in_executor(None, _run_fetch)
    ok     = bool(result.get("ok"))

    return {
        "ok":       ok,
        "message":  result.get("message") or ("Extraction successful" if ok else "Extraction failed"),
        "rows":     result.get("rows", 0),
        "filename": filename,
        "path":     output_path,
    }

# ═════════════════════════════════════════════════════════════════════════════
# BLANK TEMPLATE  —  headers-only workbook(s), streamed straight to the browser
# ═════════════════════════════════════════════════════════════════════════════

@app.get("/api/blank-template/entities")
def api_blank_template_entities():
    try:
        mod = load_module("blank_template")
        return {"ok": True, "entities": mod.supported_entities()}
    except Exception as e:
        return {"ok": False, "entities": [], "error": str(e)}


@app.post("/api/blank-template")
async def api_blank_template(req: BlankTemplateRequest):
    entities = list(dict.fromkeys(req.entities))  # de-dupe, keep order

    if not entities:
        raise HTTPException(status_code=400, detail="No entities selected.")

    try:
        mod = load_module("blank_template")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Cannot load blank_template: {e}")

    supported   = set(mod.supported_entities())
    unsupported = [e for e in entities if e not in supported]
    if unsupported:
        raise HTTPException(
            status_code=400,
            detail=(
                f"No blank template available for: {', '.join(unsupported)}. "
                f"Currently supported: {', '.join(sorted(supported))}."
            ),
        )

    loop = asyncio.get_event_loop()

    try:
        buffers = {
            entity_id: await loop.run_in_executor(None, mod.build, entity_id)
            for entity_id in entities
        }
    except Exception as exc:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Template generation failed: {exc}")

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")

    def _prefix(entity_id: str) -> str:
        return mod.TEMPLATE_SPECS.get(entity_id, {}).get("file_prefix", entity_id)

    if len(entities) == 1:
        entity_id = entities[0]
        filename  = f"{_prefix(entity_id)}_Blank_Template_{ts}.xlsx"
        buf       = buffers[entity_id]

        return StreamingResponse(
            buf,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for entity_id in entities:
            zf.writestr(f"{_prefix(entity_id)}_Blank_Template.xlsx", buffers[entity_id].getvalue())
    zip_buf.seek(0)

    zip_filename = f"Blank_Templates_{ts}.zip"
    return StreamingResponse(
        zip_buf,
        media_type="application/zip",
        headers={"Content-Disposition": f'attachment; filename="{zip_filename}"'},
    )

# ═════════════════════════════════════════════════════════════════════════════
# CONFIG  —  read / write / reload
# ═════════════════════════════════════════════════════════════════════════════

CONFIG_PATH = os.path.join(BACKEND_DIR, "config.json")


def _read_config_file() -> dict:
    with open(CONFIG_PATH, "r") as f:
        return json.load(f)


def _write_config_file(data: dict):
    with open(CONFIG_PATH, "w") as f:
        json.dump(data, f, indent=2)


def _reload_config():
    """Hot-reload CONFIG in place so running streams pick up new values."""
    global CONFIG
    import config as cfg_mod
    importlib.reload(cfg_mod)
    CONFIG = cfg_mod.CONFIG


class SaveConfigRequest(BaseModel):
    base_url:   str
    client_id:  str
    username:   str
    password:   str
    grant_type: str
    folders:    dict   # { key: path }


@app.get("/api/config")
def api_get_config():
    try:
        data = _read_config_file()
        return {"ok": True, "config": data}
    except Exception as e:
        return {"ok": False, "error": str(e)}


@app.post("/api/save-config")
def api_save_config(req: SaveConfigRequest):
    try:
        data = _read_config_file()

        data["qad"]["base_url"]           = req.base_url.strip()
        data["qad"]["auth"]["client_id"]  = req.client_id.strip()
        data["qad"]["auth"]["username"]   = req.username.strip()
        data["qad"]["auth"]["password"]   = req.password.strip()
        data["qad"]["auth"]["grant_type"] = req.grant_type.strip()

        for key, path in req.folders.items():
            if path.strip():
                data["folders"][key] = path.strip()

        _write_config_file(data)
        _reload_config()
        return {"ok": True, "message": "Config saved and reloaded."}
    except Exception as e:
        return {"ok": False, "error": str(e)}


@app.post("/api/test-connection")
def api_test_connection():
    try:
        import requests as req_lib
        cfg   = _read_config_file()
        url   = f"{cfg['qad']['base_url']}/oauth/token"
        resp  = req_lib.post(url, data=cfg["qad"]["auth"], timeout=10)
        resp.raise_for_status()
        token = resp.json().get("access_token")
        if not token:
            return {"ok": False, "error": "No access_token in response"}
        return {"ok": True, "message": "Connection successful ✓"}
    except Exception as e:
        return {"ok": False, "error": str(e)}

# ═════════════════════════════════════════════════════════════════════════════
# LOGIN  —  live OAuth check against the selected environment
# ═════════════════════════════════════════════════════════════════════════════
#
# Takes username/password straight from the login form and performs a live
# OAuth token request against config.json's environments[<environment>]
# (base_url + client_id). The resulting access token is kept ONLY in the
# in-memory SESSIONS store (see above) — nothing is written to disk anymore.
# The frontend gets back a session_id, which it stores in sessionStorage and
# sends along on later calls that need the token server-side (currently just
# /api/permissions). Access gating on the frontend is still the lightweight
# sessionStorage flag set only after this endpoint returns ok: true.

@app.post("/api/login")
def api_login(req: LoginRequest):
    username    = req.username.strip()
    environment = req.environment.strip().upper()

    try:
        cfg          = _read_config_file()
        environments = cfg.get("environments", {})
        env_cfg      = environments.get(environment)

        if not env_cfg:
            log_activity({
                "username": username, "environment": environment,
                "action": "login", "status": "failed",
                "details": {"error": f"Unknown environment '{environment}'"},
            })
            return {"ok": False, "error": f"Unknown environment '{environment}'"}

        import requests as req_lib
        url = f"{env_cfg['base_url']}/oauth/token"
        payload = {
            "client_id":  env_cfg["client_id"],
            "username":   username,
            "password":   req.password,
            "grant_type": env_cfg.get("grant_type", "password"),
        }

        resp = req_lib.post(url, data=payload, timeout=10)
        resp.raise_for_status()

        oauth_response = resp.json()
        token = oauth_response.get("access_token")

        if not token:
            log_activity({
                "username": username, "environment": environment,
                "action": "login", "status": "failed",
                "details": {"error": "No access_token in response"},
            })
            return {"ok": False, "error": "No access_token in response"}

        session_id = _create_session(username, environment, token)

        log_activity({
            "username": username, "environment": environment,
            "action": "login", "status": "success",
        })
        return {
            "ok": True,
            "message": "Login successful",
            "username": username,
            "environment": environment,
            "session_id": session_id,
        }

    except Exception as e:
        log_activity({
            "username": username, "environment": environment,
            "action": "login", "status": "failed",
            "details": {"error": str(e)},
        })
        return {"ok": False, "error": str(e)}

# ═════════════════════════════════════════════════════════════════════════════
# PERMISSIONS  —  QAD access-based module control
# ═════════════════════════════════════════════════════════════════════════════
#
# config.json is expected to grow a "permissions" section shaped like:
#
#   "permissions": {
#     "entity_uri_map": {
#       "Customer": "urn:view:hybridbrowse:com.qad.erp.base.customerV2s"
#       // ... one urn per ENTITY_MAP key ...
#     },
#     "required_actions": ["WRITE", "CREATE", "IMPORT"],
#     "require_all": true
#   }
#
# Per your call: the qracore host is NOT a separate config value — it's
# always environments[<environment>].base_url, the same base_url already
# used for the OAuth token request at login, so it moves automatically if
# TEST/PROD base_url ever changes in config.json. No extra field needed.
#
# require_all=true means a module only unlocks if EVERY action in
# required_actions comes back hasPermission:true (confirmed).
#
# Enforcement happens in TWO places now (re-check confirmed as the safer
# option):
#   1. GET /api/permissions — called on page load AND again before every
#      Validate/Load click, purely to drive the UI lock state.
#   2. Inside validate_stream/load_stream themselves — re-checked per
#      entity right before that entity is processed, so a stale/tampered
#      frontend can't bypass the lock by just not calling /api/permissions
#      or by re-enabling a disabled button via devtools.
#
# Behavior when config (or an entity's URI) is missing: we do NOT lock the
# module — permission checks are opt-in per entity, so partially-rolled-out
# config never blocks modules that haven't been wired up yet. If the
# permissionState call itself fails for an entity that DOES have a URI
# configured, we fail closed (treat as no permission).

def _entity_permission_config() -> tuple[dict, list[str], bool]:
    cfg = _read_config_file()
    perm_cfg = cfg.get("permissions", {})
    uri_map = perm_cfg.get("entity_uri_map", {})
    actions = perm_cfg.get("required_actions", ["WRITE", "CREATE", "IMPORT"])
    require_all = perm_cfg.get("require_all", True)
    return uri_map, actions, require_all


def _qracore_base_url(environment: str) -> str | None:
    """Same host QAD OAuth uses for this environment — read fresh every
    call so a config.json edit takes effect on the next check, no restart
    needed."""
    cfg = _read_config_file()
    base_url = cfg.get("environments", {}).get(environment, {}).get("base_url")
    return base_url.rstrip("/") if base_url else None


def _check_entity_permission(
    base_url: str, token: str, uri: str, actions: list[str], require_all: bool
) -> bool:
    import requests as req_lib
    results = []
    for action in actions:
        try:
            resp = req_lib.get(
                f"{base_url}/api/qracore/permissionState",
                params={"uri": uri, "action": action},
                headers={"Authorization": f"Bearer {token}"},
                timeout=10,
            )
            resp.raise_for_status()
            has_perm = bool(resp.json().get("data", {}).get("hasPermission"))
        except Exception as e:
            print(f"[main.py] permissionState check failed (uri={uri}, action={action}): {e}")
            has_perm = False
        results.append(has_perm)
    return all(results) if require_all else any(results)


class _PermissionContext:
    """Bundles what's needed to check a single entity's permission, resolved
    once per request/stream rather than re-reading config.json per entity."""
    def __init__(self, active: bool, base_url: str | None, token: str | None,
                 uri_map: dict, actions: list[str], require_all: bool):
        self.active      = active
        self.base_url    = base_url
        self.token       = token
        self.uri_map     = uri_map
        self.actions     = actions
        self.require_all = require_all


def _resolve_permission_context(session_id: str | None) -> _PermissionContext:
    uri_map, actions, require_all = _entity_permission_config()
    if not uri_map:
        return _PermissionContext(False, None, None, uri_map, actions, require_all)

    session = _get_session(session_id) if session_id else None
    if not session:
        # No/expired session — can't call QAD on the user's behalf. Since
        # permissions ARE configured, fail closed rather than silently
        # letting an unauthenticated stream run.
        return _PermissionContext(True, None, None, uri_map, actions, require_all)

    base_url = _qracore_base_url(session["environment"])
    return _PermissionContext(True, base_url, session["access_token"], uri_map, actions, require_all)


def _entity_allowed(ctx: _PermissionContext, entity_id: str) -> bool:
    if not ctx.active:
        return True  # permissions not configured at all -> never lock
    uri = ctx.uri_map.get(entity_id)
    if not uri:
        return True  # this entity has no URI configured -> don't lock it
    if not ctx.base_url or not ctx.token:
        return False  # configured but we have no way to check -> fail closed
    return _check_entity_permission(ctx.base_url, ctx.token, uri, ctx.actions, ctx.require_all)


@app.get("/api/permissions")
def api_permissions(session_id: str):
    session = _get_session(session_id)
    if not session:
        raise HTTPException(status_code=401, detail="Session expired or invalid — please log in again.")

    ctx = _resolve_permission_context(session_id)
    if not ctx.active:
        return {"ok": True, "configured": False, "permissions": {eid: True for eid in ENTITY_MAP}}

    permissions = {eid: _entity_allowed(ctx, eid) for eid in ENTITY_MAP}
    return {"ok": True, "configured": True, "permissions": permissions}

# ═════════════════════════════════════════════════════════════════════════════
# ACTIVITY LOG  —  generic logging endpoint for workspace actions
# ═════════════════════════════════════════════════════════════════════════════

@app.post("/api/log-activity")
def api_log_activity(req: ActivityLogRequest):
    log_activity({
        "username":    req.username,
        "environment": req.environment,
        "action":      req.action,
        "status":      req.status,
        "details":     req.details,
    })
    return {"ok": True}

# ═════════════════════════════════════════════════════════════════════════════
# ROUTES  —  validate / load
# ═════════════════════════════════════════════════════════════════════════════

SSE_HEADERS = {
    "Cache-Control":     "no-cache, no-transform",
    "X-Accel-Buffering": "no",
    "Connection":        "keep-alive",
}


@app.post("/api/validate")
async def api_validate(req: OperationRequest):
    return StreamingResponse(
        validate_stream(req.entities, req.session_id),
        media_type="text/event-stream",
        headers=SSE_HEADERS,
    )


@app.post("/api/load")
async def api_load(req: OperationRequest):
    return StreamingResponse(
        load_stream(req.entities, req.session_id),
        media_type="text/event-stream",
        headers=SSE_HEADERS,
    )


@app.get("/health")
def health():
    return {"status": "ok"}

# ═════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ═════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)