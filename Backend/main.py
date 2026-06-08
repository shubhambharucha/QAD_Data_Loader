"""
QAD Data Loader — FastAPI Backend
Serves the React frontend and exposes SSE-based validate/load endpoints.
"""

import os, sys, shutil, json, asyncio
from datetime import datetime
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from typing import List

# ── Path setup ────────────────────────────────────────────────────────────
BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
SCRIPTS_DIR = os.path.join(BASE_DIR, "Scripts")
sys.path.insert(0, SCRIPTS_DIR)

app = FastAPI(title="QAD Data Loader API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "http://localhost:3000"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── Entity config ─────────────────────────────────────────────────────────
ENTITIES = [
    "Supplier", "Customer", "GCM", "BR",
    "Customer_Item", "ProductionOrder", "PurchaseOrder",
    "SalesOrder", "Supplier_Item", "SupplierPriceList",
]

VALIDATE_MAP = {
    "Supplier":          "validate_supplier",
    "Customer":          "validate_customer",
    "GCM":               "validate_gcm",
    "BR":                "validate_br",
    "Customer_Item":     "validate_cust_item",
    "ProductionOrder":   "validate_production_order",
    "PurchaseOrder":     "validate_po",
    "SalesOrder":        "validate_so",
    "Supplier_Item":     "validate_Supp_item",
    "SupplierPriceList": "validate_supp_price_lists",
}

LOAD_MAP = {
    "Supplier":          "Supplier_load",
    "Customer":          "Customer_load",
    "GCM":               "GCM_load",
    "BR":                "BR_load",
    "Customer_Item":     "cust_item",
    "ProductionOrder":   "production_order",
    "PurchaseOrder":     "po_load",
    "SalesOrder":        "so_load",
    "Supplier_Item":     "Supp_item",
    "SupplierPriceList": "supp_price_lists",
}

# ── Helpers ───────────────────────────────────────────────────────────────
def ensure_folders():
    for e in ENTITIES:
        os.makedirs(os.path.join(BASE_DIR, e), exist_ok=True)
        os.makedirs(os.path.join(BASE_DIR, "Archive", e), exist_ok=True)

def get_xlsx(entity: str) -> List[str]:
    folder = os.path.join(BASE_DIR, entity)
    if not os.path.exists(folder):
        return []
    return [f for f in os.listdir(folder) if f.endswith(".xlsx") and not f.startswith("~$")]

def archive_file(entity: str, fp: str):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = os.path.join(
        BASE_DIR, "Archive", entity,
        f"{os.path.splitext(os.path.basename(fp))[0]}_{ts}.xlsx"
    )
    shutil.move(fp, dest)

def apply_error_prefix(entity: str, filename: str) -> str:
    folder = os.path.join(BASE_DIR, entity)
    if not filename.startswith("error_"):
        src = os.path.join(folder, filename)
        dst = os.path.join(folder, "error_" + filename)
        if os.path.exists(src):
            os.rename(src, dst)
        return "error_" + filename
    return filename

def remove_error_prefix(entity: str, filename: str) -> str:
    folder = os.path.join(BASE_DIR, entity)
    if filename.startswith("error_"):
        src = os.path.join(folder, filename)
        dst = os.path.join(folder, filename[len("error_"):])
        if os.path.exists(src):
            os.rename(src, dst)
        return filename[len("error_"):]
    return filename

def import_fresh(mod: str):
    sys.modules.pop(mod, None)
    return __import__(mod)

def sse_event(data: dict) -> str:
    return f"data: {json.dumps(data)}\n\n"

ensure_folders()

# ── Models ────────────────────────────────────────────────────────────────
class EntityRequest(BaseModel):
    entities: List[str]

# ── Routes ────────────────────────────────────────────────────────────────
@app.get("/api/status")
def get_status():
    """Return file counts for all entities."""
    return {
        e: {
            "files": get_xlsx(e),
            "count": len(get_xlsx(e)),
        }
        for e in ENTITIES
    }

@app.post("/api/validate")
def validate_stream(req: EntityRequest):
    """SSE stream: validates selected entities, yields progress events."""

    async def generator():
        for entity in req.entities:
            files = get_xlsx(entity)

            if not files:
                yield sse_event({
                    "type": "entity_result",
                    "entity": entity,
                    "ok": False,
                    "reason": "no_files",
                    "error_files": [],
                })
                continue

            script = VALIDATE_MAP.get(entity)
            if not script:
                yield sse_event({
                    "type": "entity_result",
                    "entity": entity,
                    "ok": False,
                    "reason": "no_script",
                    "error_files": [],
                })
                continue

            try:
                mod = import_fresh(script)
            except ImportError as e:
                yield sse_event({
                    "type": "entity_result",
                    "entity": entity,
                    "ok": False,
                    "reason": f"import_error: {e}",
                    "error_files": [],
                })
                continue

            error_files = []
            for f in files:
                fp = os.path.join(BASE_DIR, entity, f)
                yield sse_event({"type": "progress", "entity": entity, "file": f})
                await asyncio.sleep(0)  # yield control

                try:
                    has_err, _ = mod.validate(fp)
                    if has_err:
                        error_files.append(f)
                        apply_error_prefix(entity, f)
                    else:
                        remove_error_prefix(entity, f)
                except Exception as e:
                    error_files.append(f)
                    yield sse_event({"type": "error", "entity": entity, "file": f, "msg": str(e)})

            yield sse_event({
                "type": "entity_result",
                "entity": entity,
                "ok": len(error_files) == 0,
                "error_files": error_files,
            })

        yield sse_event({"type": "done"})

    return StreamingResponse(generator(), media_type="text/event-stream")


@app.post("/api/load")
def load_stream(req: EntityRequest):
    """SSE stream: loads selected entities, yields progress + summary events."""

    async def generator():
        for entity in req.entities:
            files = get_xlsx(entity)
            script = LOAD_MAP.get(entity)

            if not script:
                yield sse_event({
                    "type": "file_result",
                    "entity": entity,
                    "file": "—",
                    "ok": 0,
                    "fail": "?",
                    "status": "error",
                    "note": "Script not found",
                })
                continue

            try:
                mod = import_fresh(script)
            except ImportError:
                yield sse_event({
                    "type": "file_result",
                    "entity": entity,
                    "file": "—",
                    "ok": 0,
                    "fail": "?",
                    "status": "error",
                    "note": "Script not found",
                })
                continue

            for f in files:
                fp = os.path.join(BASE_DIR, entity, f)
                yield sse_event({"type": "progress", "entity": entity, "file": f})
                await asyncio.sleep(0)

                try:
                    ok_count, fail_count = mod.run(fp)
                    if fail_count == 0:
                        restored = remove_error_prefix(entity, f)
                        archive_file(entity, os.path.join(BASE_DIR, entity, restored))
                        status, note = "archived", "Archived"
                    else:
                        apply_error_prefix(entity, f)
                        status, note = "partial", "Fix red rows & re-run"

                    yield sse_event({
                        "type": "file_result",
                        "entity": entity,
                        "file": f,
                        "ok": ok_count,
                        "fail": fail_count,
                        "status": status,
                        "note": note,
                    })
                except Exception as e:
                    yield sse_event({
                        "type": "file_result",
                        "entity": entity,
                        "file": f,
                        "ok": 0,
                        "fail": "?",
                        "status": "error",
                        "note": str(e)[:80],
                    })

        yield sse_event({"type": "done"})

    return StreamingResponse(generator(), media_type="text/event-stream")


# ── Run ───────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
