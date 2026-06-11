"""
debug.py  —  Run this from Backend/ to diagnose main.py issues
Usage:  python debug.py
"""

import importlib.util
import json
import os
import sys

# ── Same path logic as main.py ────────────────────────────────────────────────
BACKEND_DIR = os.path.abspath(os.path.dirname(__file__))
ROOT_DIR    = os.path.abspath(os.path.join(BACKEND_DIR, ".."))
SCRIPTS_DIR = os.path.join(BACKEND_DIR, "Scripts")
DATA_DIR    = os.path.join(ROOT_DIR, "Data")

sys.path.insert(0, ROOT_DIR)
sys.path.insert(0, BACKEND_DIR)
sys.path.insert(0, SCRIPTS_DIR)

# ── Colours ───────────────────────────────────────────────────────────────────
G = "\033[92m"   # green
R = "\033[91m"   # red
Y = "\033[93m"   # yellow
W = "\033[0m"    # reset
OK  = f"{G}  [OK]{W}"
ERR = f"{R} [ERR]{W}"
WRN = f"{Y}[WARN]{W}"

def ok(msg):  print(f"{OK}  {msg}")
def err(msg): print(f"{ERR}  {msg}")
def warn(msg):print(f"{WRN}  {msg}")
def hdr(msg): print(f"\n{'─'*55}\n  {msg}\n{'─'*55}")

# ═══════════════════════════════════════════════════════
# 1. DIRECTORY STRUCTURE
# ═══════════════════════════════════════════════════════
hdr("1. Directory paths")
print(f"  BACKEND_DIR : {BACKEND_DIR}")
print(f"  ROOT_DIR    : {ROOT_DIR}")
print(f"  SCRIPTS_DIR : {SCRIPTS_DIR}")
print(f"  DATA_DIR    : {DATA_DIR}")

for label, path in [
    ("ROOT_DIR",    ROOT_DIR),
    ("BACKEND_DIR", BACKEND_DIR),
    ("SCRIPTS_DIR", SCRIPTS_DIR),
    ("DATA_DIR",    DATA_DIR),
]:
    if os.path.isdir(path): ok(f"{label} exists")
    else:                    err(f"{label} NOT FOUND: {path}")

# ═══════════════════════════════════════════════════════
# 2. CONFIG
# ═══════════════════════════════════════════════════════
hdr("2. Config")
try:
    from config import CONFIG
    ok("config.py imported OK")
    print(f"  folders keys: {list(CONFIG.get('folders', {}).keys())}")
except Exception as e:
    err(f"config.py import failed: {e}")
    CONFIG = {"folders": {}}

# ═══════════════════════════════════════════════════════
# 3. DATA FOLDERS  (resolved paths + xlsx count)
# ═══════════════════════════════════════════════════════
hdr("3. Data folders & xlsx files")

ENTITY_MAP = {
    "Supplier":         {"folder_key": "supplier",          "validate": "validate_supplier",        "load": "Supplier_load"},
    "Customer":         {"folder_key": "customer",          "validate": "validate_customer",        "load": "Customer_load"},
    "GCM":              {"folder_key": "GCM",               "validate": "validate_gcm",             "load": "GCM_load"},
    "BR":               {"folder_key": "BR",                "validate": "validate_br",              "load": "BR_load"},
    "Customer_Item":    {"folder_key": "customer_item",     "validate": "validate_cust_item",       "load": "cust_item"},
    "ProductionOrder":  {"folder_key": "production_order",  "validate": "validate_production_order","load": "production_order"},
    "PurchaseOrder":    {"folder_key": "purchase_order",    "validate": "validate_po",              "load": "po_load"},
    "SalesOrder":       {"folder_key": "sales_order",       "validate": "validate_so",              "load": "so_load"},
    "Supplier_Item":    {"folder_key": "supplier_item",     "validate": "validate_Supp_item",       "load": "Supp_item"},
    "SupplierPriceList":{"folder_key": "supplier_price_list","validate":"validate_supp_price_lists","load": "supp_price_lists"},
}

folders = CONFIG.get("folders", {})
for entity, cfg in ENTITY_MAP.items():
    key = cfg["folder_key"]
    rel = folders.get(key)
    if not rel:
        warn(f"{entity:20s} → key '{key}' missing from config.json folders")
        continue
    abs_path = os.path.abspath(os.path.join(BACKEND_DIR, rel))
    if not os.path.isdir(abs_path):
        err(f"{entity:20s} → folder NOT FOUND: {abs_path}")
        continue
    xlsx = [f for f in os.listdir(abs_path) if f.endswith(".xlsx") and not f.startswith("~$")]
    if xlsx:
        ok(f"{entity:20s} → {len(xlsx)} xlsx file(s): {xlsx}")
    else:
        warn(f"{entity:20s} → folder exists but NO xlsx files: {abs_path}")

# ═══════════════════════════════════════════════════════
# 4. SCRIPT IMPORTS
# ═══════════════════════════════════════════════════════
hdr("4. Script imports")

all_scripts = set()
for cfg in ENTITY_MAP.values():
    all_scripts.add(cfg["validate"])
    all_scripts.add(cfg["load"])

for script in sorted(all_scripts):
    path = os.path.join(SCRIPTS_DIR, f"{script}.py")
    if not os.path.exists(path):
        err(f"{script}.py  NOT FOUND")
        continue
    try:
        spec   = importlib.util.spec_from_file_location(script, path)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        has_validate = hasattr(module, "validate")
        has_process  = hasattr(module, "process_file")
        has_run      = hasattr(module, "run")
        flags = []
        if has_validate: flags.append("validate()")
        if has_process:  flags.append("process_file()")
        if has_run:      flags.append("run()")
        ok(f"{script}.py  [{', '.join(flags) if flags else 'no expected functions!'}]")
    except Exception as e:
        err(f"{script}.py  import FAILED: {e}")

# ═══════════════════════════════════════════════════════
# 5. SERVER REACHABILITY
# ═══════════════════════════════════════════════════════
hdr("5. Server reachability (is main.py running?)")
try:
    import urllib.request
    with urllib.request.urlopen("http://localhost:8000/health", timeout=3) as r:
        body = r.read().decode()
        ok(f"http://localhost:8000/health → {body}")
except Exception as e:
    err(f"Cannot reach http://localhost:8000  ({e})")
    warn("Make sure 'python main.py' is running in another terminal")

print(f"\n{'─'*55}")
print("  Debug complete. Fix any [ERR] lines above.")
print(f"{'─'*55}\n")