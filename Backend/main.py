"""
main.py  —  QAD Data Loader FastAPI Backend
============================================
Endpoints
---------
POST /api/validate   { "entities": ["Supplier", "Customer", ...] }
POST /api/load       { "entities": ["Supplier", "Customer", ...] }

Both stream SSE events to the frontend terminal.

File-rename logic
-----------------
Validate & Load:
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
"""

import asyncio
import importlib.util
import json
import os
import re
import shutil
import sys
import traceback
from datetime import datetime
from typing import AsyncGenerator

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel

# ── Locate project root (Backend/ sits next to main.py) ──────────────────────
BACKEND_DIR = os.path.abspath(os.path.dirname(__file__))        # .../QAD_data_loader/Backend
ROOT_DIR    = os.path.abspath(os.path.join(BACKEND_DIR, "..")) # .../QAD_data_loader
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
)

# ═════════════════════════════════════════════════════════════════════════════
# REQUEST SCHEMA
# ═════════════════════════════════════════════════════════════════════════════

class OperationRequest(BaseModel):
    entities: list[str]

# ═════════════════════════════════════════════════════════════════════════════
# HELPERS
# ═════════════════════════════════════════════════════════════════════════════

def sse(event: dict) -> str:
    """Encode a dict as an SSE data line."""
    return f"data: {json.dumps(event)}\n\n"


def resolve_folder(entity_id: str) -> str:
    """Return absolute path to entity's data folder."""
    cfg      = ENTITY_MAP[entity_id]
    rel_path = CONFIG["folders"].get(cfg["folder_key"], "")
    if not rel_path:
        raise ValueError(f"No folder configured for entity '{entity_id}' (key: {cfg['folder_key']})")
    # CONFIG paths are relative to Backend/ (e.g. "../Data/Supplier")
    return os.path.abspath(os.path.join(BACKEND_DIR, rel_path))


def resolve_archive(entity_id: str) -> str:
    """Return absolute path to entity's Archive folder; creates it if needed."""
    archive_entity = ENTITY_MAP[entity_id]["archive_folder"]
    path = os.path.join(DATA_DIR, "Archive", archive_entity)
    os.makedirs(path, exist_ok=True)
    return path


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


# ── File rename helpers ───────────────────────────────────────────────────────

def _error_name(filename: str) -> str:
    """supplier_001.xlsx  →  error_supplier_001.xlsx"""
    return f"error_{filename}" if not filename.startswith("error_") else filename


def _clean_name(filename: str) -> str:
    """error_supplier_001.xlsx  →  supplier_001.xlsx"""
    return re.sub(r"^error_", "", filename)


def rename_file_error(folder: str, filename: str) -> str:
    """Rename file to error_<name> if not already. Returns new filename."""
    new_name = _error_name(filename)
    if new_name != filename:
        os.rename(
            os.path.join(folder, filename),
            os.path.join(folder, new_name),
        )
    return new_name


def rename_file_clean(folder: str, filename: str) -> str:
    """Remove error_ prefix from file. Returns new filename."""
    new_name = _clean_name(filename)
    if new_name != filename:
        os.rename(
            os.path.join(folder, filename),
            os.path.join(folder, new_name),
        )
    return new_name


def archive_file(folder: str, filename: str, archive_dir: str) -> str:
    """
    Move file to archive_dir with timestamp appended.
    supplier_001.xlsx  →  Archive/Supplier/supplier_001_20250611_143022.xlsx
    Returns the archive path.
    """
    clean      = _clean_name(filename)                 # strip error_ just in case
    stem, ext  = os.path.splitext(clean)
    ts         = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest_name  = f"{stem}_{ts}{ext}"
    dest_path  = os.path.join(archive_dir, dest_name)
    shutil.move(os.path.join(folder, filename), dest_path)
    return dest_path


# ═════════════════════════════════════════════════════════════════════════════
# VALIDATE STREAM GENERATOR
# ═════════════════════════════════════════════════════════════════════════════

async def validate_stream(entities: list[str]) -> AsyncGenerator[str, None]:
    """
    For each entity:
      1. Emit entity_start
      2. For each .xlsx file in its folder:
         a. Emit file_start
         b. Run validate() in a thread (so the event loop isn't blocked)
            — validate() saves the workbook and returns a summary dict
         c. Emit file_result
         d. Rename file based on has_errors
      3. Emit entity_result
    Finally emit done.
    """

    loop = asyncio.get_event_loop()

    for entity_id in entities:

        if entity_id not in ENTITY_MAP:
            yield sse({"type": "error", "message": f"Unknown entity: {entity_id}"})
            continue

        cfg = ENTITY_MAP[entity_id]

        # ── Resolve folder ────────────────────────────────────────────────
        try:
            folder = resolve_folder(entity_id)
        except Exception as e:
            yield sse({"type": "error", "message": str(e)})
            continue

        yield sse({"type": "entity_start", "entity": entity_id})
        await asyncio.sleep(0)  # flush

        files = list_xlsx(folder)

        if not files:
            yield sse({
                "type":    "error",
                "message": f"No .xlsx files found in {folder}",
            })
            yield sse({
                "type": "entity_result",
                "entity":  entity_id,
                "passed":  0,
                "failed":  0,
                "skipped": 0,
            })
            continue

        # ── Load validate module once per entity ──────────────────────────
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

            # ── Run validate in thread ────────────────────────────────────
            try:
                result: dict = await loop.run_in_executor(
                    None, mod.validate, file_path
                )
            except Exception as exc:
                tb = traceback.format_exc()
                yield sse({
                    "type":   "file_result",
                    "entity": entity_id,
                    "file":   filename,
                    "ok":     0,
                    "fail":   1,
                    "skipped": 0,
                    "status": "error",
                    "note":   str(exc),
                })
                entity_failed += 1
                rename_file_error(folder, filename)
                continue

            rows_passed  = result.get("rows_passed",  0)
            rows_failed  = result.get("rows_failed",  0)
            rows_skipped = result.get("rows_skipped", 0)
            has_errors   = result.get("has_errors",   False)

            # ── Rename file based on error status ─────────────────────────
            if has_errors:
                entity_failed += 1
                rename_file_error(folder, filename)
                status = "failed"
            else:
                entity_passed += 1
                rename_file_clean(folder, filename)   # remove error_ if present
                status = "passed"

            yield sse({
                "type":    "file_result",
                "entity":  entity_id,
                "file":    filename,
                "ok":      rows_passed,
                "fail":    rows_failed,
                "skipped": rows_skipped,
                "status":  status,
                "note":    f"{rows_skipped} rows skipped" if rows_skipped else "",
            })
            await asyncio.sleep(0)

        yield sse({
            "type":    "entity_result",
            "entity":  entity_id,
            "passed":  entity_passed,
            "failed":  entity_failed,
            "skipped": entity_skipped,
        })
        await asyncio.sleep(0)

    yield sse({"type": "done", "message": "Validation complete"})


# ═════════════════════════════════════════════════════════════════════════════
# LOAD STREAM GENERATOR
# ═════════════════════════════════════════════════════════════════════════════

async def load_stream(entities: list[str]) -> AsyncGenerator[str, None]:
    """
    For each entity:
      1. Emit entity_start
      2. For each .xlsx file in its folder:
         a. Emit file_start
         b. Run process_file() in a thread — returns (ok_count, fail_count)
         c. Emit file_result
         d. Rename/archive based on fail_count
      3. Emit entity_result
    Finally emit done.

    Note: progress events (row N of M) are emitted via a queue that
          process_file writes into. We monkey-patch a progress callback
          onto the module for this session.
    """

    loop = asyncio.get_event_loop()

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

        files = list_xlsx(folder)

        if not files:
            yield sse({
                "type":    "error",
                "message": f"No .xlsx files found in {folder}",
            })
            yield sse({
                "type": "entity_result",
                "entity":  entity_id,
                "passed":  0,
                "failed":  0,
                "skipped": 0,
            })
            continue

        # ── Load load-module once per entity ──────────────────────────────
        try:
            mod = load_module(cfg["load_script"])
        except Exception as e:
            yield sse({"type": "error", "message": f"Cannot load {cfg['load_script']}: {e}"})
            continue

        # ── Shared progress queue for this entity ─────────────────────────
        progress_queue: asyncio.Queue = asyncio.Queue()

        entity_passed  = 0
        entity_failed  = 0
        entity_skipped = 0

        for filename in files:

            file_path = os.path.join(folder, filename)

            yield sse({"type": "file_start", "entity": entity_id, "file": filename})
            await asyncio.sleep(0)

            # ── Run process_file in thread; collect progress via queue ─────
            result_holder: list = []   # [ok, fail] written by thread

            def _run_process_file():
                """Called in executor thread. Writes (ok, fail) to result_holder."""
                try:
                    # If the load module exposes TokenManager at module level
                    # we reuse it; otherwise we create a fresh one.
                    if hasattr(mod, "_tm"):
                        tm = mod._tm
                    else:
                        tm = mod.TokenManager()
                        mod._tm = tm   # cache for subsequent files

                    ok, fail = mod.process_file(file_path, tm)
                    result_holder.extend([ok, fail])
                except Exception as exc:
                    result_holder.extend([0, -1])   # -1 signals exception
                    result_holder.append(str(exc))

            future = loop.run_in_executor(None, _run_process_file)

            # ── Poll progress while thread runs ───────────────────────────
            # The load scripts save the workbook row-by-row, so we can
            # approximate progress by watching the workbook's last-modified
            # time — OR simply emit a heartbeat while waiting.
            # A clean solution: periodically yield a progress pulse.
            #
            # For true row-level progress, the load scripts would need a
            # callback hook. We implement an optional hook here: if the
            # module defines `set_progress_callback`, we use it.
            #
            # Until then, emit a "working…" pulse every 0.5 s.

            row_counter = [0]

            if hasattr(mod, "set_progress_callback"):
                def _on_progress(row: int, total: int):
                    row_counter[0] = row
                    # thread-safe put_nowait into the queue
                    loop.call_soon_threadsafe(
                        progress_queue.put_nowait,
                        {"row": row, "total": total},
                    )
                mod.set_progress_callback(_on_progress)

            while not future.done():
                await asyncio.sleep(0.4)
                # Drain any progress events the thread enqueued
                drained = False
                while not progress_queue.empty():
                    pev = progress_queue.get_nowait()
                    yield sse({
                        "type":   "progress",
                        "entity": entity_id,
                        "file":   filename,
                        "row":    pev["row"],
                        "total":  pev["total"],
                    })
                    drained = True
                if not drained and row_counter[0] > 0:
                    # emit last known row as a keepalive
                    yield sse({
                        "type":   "progress",
                        "entity": entity_id,
                        "file":   filename,
                        "row":    row_counter[0],
                        "total":  row_counter[0],
                    })

            # Await to surface any executor exceptions
            try:
                await future
            except Exception:
                pass

            # Clear callback
            if hasattr(mod, "set_progress_callback"):
                mod.set_progress_callback(None)

            # ── Parse result ──────────────────────────────────────────────
            if len(result_holder) >= 3 and result_holder[1] == -1:
                # Exception in thread
                exc_msg = result_holder[2]
                yield sse({
                    "type":    "file_result",
                    "entity":  entity_id,
                    "file":    filename,
                    "ok":      0,
                    "fail":    1,
                    "skipped": 0,
                    "status":  "error",
                    "note":    exc_msg,
                })
                entity_failed += 1
                rename_file_error(folder, filename)
                continue

            ok_count   = result_holder[0] if result_holder else 0
            fail_count = result_holder[1] if len(result_holder) > 1 else 1

            # ── Rename / archive ──────────────────────────────────────────
            if fail_count > 0:
                entity_failed += 1
                new_filename = rename_file_error(folder, filename)
                status = "failed"
                note   = f"{fail_count} rows failed"
            else:
                entity_passed += 1
                # Remove error_ prefix first (if present)
                clean_filename = rename_file_clean(folder, filename)
                # Move to Archive
                archive_path = archive_file(folder, clean_filename, archive_dir)
                status = "passed"
                note   = f"Archived → {os.path.basename(archive_path)}"

            yield sse({
                "type":    "file_result",
                "entity":  entity_id,
                "file":    filename,
                "ok":      ok_count,
                "fail":    fail_count,
                "skipped": 0,
                "status":  status,
                "note":    note,
            })
            await asyncio.sleep(0)

        yield sse({
            "type":    "entity_result",
            "entity":  entity_id,
            "passed":  entity_passed,
            "failed":  entity_failed,
            "skipped": entity_skipped,
        })
        await asyncio.sleep(0)

    yield sse({"type": "done", "message": "Load complete"})


# ═════════════════════════════════════════════════════════════════════════════
# ROUTES
# ═════════════════════════════════════════════════════════════════════════════

@app.post("/api/validate")
async def api_validate(req: OperationRequest):
    return StreamingResponse(
        validate_stream(req.entities),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
        },
    )


@app.post("/api/load")
async def api_load(req: OperationRequest):
    return StreamingResponse(
        load_stream(req.entities),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
        },
    )


@app.get("/health")
def health():
    return {"status": "ok"}


# ═════════════════════════════════════════════════════════════════════════════
# ENTRY POINT  (python main.py  or  uvicorn main:app)
# ═════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)