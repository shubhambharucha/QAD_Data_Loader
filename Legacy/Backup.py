import streamlit as st
import os
import sys
import shutil
from datetime import datetime

# ── Paths ──────────────────────────────────────────────────────────────────
BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
SCRIPTS_DIR = os.path.join(BASE_DIR, "Scripts")
sys.path.insert(0, SCRIPTS_DIR)

ENTITIES = ["Supplier", "Customer", "GCM", "BR"]

VALIDATE_MAP = {
    "Supplier": "validate_supplier",
    "Customer": "validate_customer",
    "GCM":      "validate_gcm",
    "BR":       "validate_br",
}

LOAD_MAP = {
    "Supplier": "Supplier_load",
    "Customer": "Customer_load",
    "GCM":      "GCM_load",
    "BR":       "BR_load",
}

# ── Helpers ────────────────────────────────────────────────────────────────
def ensure_folders():
    for entity in ENTITIES:
        os.makedirs(os.path.join(BASE_DIR, entity), exist_ok=True)
        os.makedirs(os.path.join(BASE_DIR, "Archive", entity), exist_ok=True)

def get_xlsx_files(entity):
    folder = os.path.join(BASE_DIR, entity)
    if not os.path.exists(folder):
        return []
    return [f for f in os.listdir(folder) if f.endswith(".xlsx")]

def archive_file(entity, file_path):
    timestamp   = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_name   = os.path.splitext(os.path.basename(file_path))[0]
    new_name    = f"{base_name}_{timestamp}.xlsx"
    archive_dir = os.path.join(BASE_DIR, "Archive", entity)
    shutil.move(file_path, os.path.join(archive_dir, new_name))

def import_fresh(module_name):
    if module_name in sys.modules:
        del sys.modules[module_name]
    return __import__(module_name)

# ── Page config ────────────────────────────────────────────────────────────
st.set_page_config(page_title="QAD Data Loader", page_icon="📦", layout="centered")
ensure_folders()

# ── Session state init ─────────────────────────────────────────────────────
if "val_results" not in st.session_state:
    st.session_state.val_results = {}   # { entity: {"ok": bool, "error_files": [...]} }
if "validated" not in st.session_state:
    st.session_state.validated = False

# ── Custom CSS ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* Load button disabled state */
    div[data-testid="stButton"] button:disabled {
        background-color: #555 !important;
        color: #999 !important;
        cursor: not-allowed !important;
        border: none !important;
    }
    .entity-row {
        display: flex;
        align-items: center;
        gap: 8px;
        margin-bottom: 2px;
    }
    .status-icon {
        font-size: 1.1rem;
    }
    .error-tree {
        margin-left: 28px;
        margin-top: 4px;
        margin-bottom: 8px;
    }
    .error-file {
        color: #FF4B4B;
        font-size: 0.85rem;
        margin: 2px 0;
    }
    .success-msg {
        color: #00B050;
        font-size: 0.85rem;
        margin-left: 28px;
        margin-top: 2px;
        margin-bottom: 8px;
    }
    .no-files-msg {
        color: #888;
        font-size: 0.85rem;
        margin-left: 28px;
        margin-top: 2px;
        margin-bottom: 8px;
    }
</style>
""", unsafe_allow_html=True)

# ── Header ─────────────────────────────────────────────────────────────────
st.title("📦 QAD Data Loader")
st.caption("Select entities, validate your files, then load when all checks pass.")
st.divider()

# ── Entity Selection ───────────────────────────────────────────────────────
st.subheader("Select Entities")

checks = {}
for entity in ENTITIES:
    files     = get_xlsx_files(entity)
    file_count = len(files)
    label     = f"**{entity}** — {file_count} file(s) ready"

    col_check, col_status = st.columns([6, 1])

    with col_check:
        checks[entity] = st.checkbox(label, key=f"chk_{entity}")

    # Inline status icon (only shown after validation)
    with col_status:
        if st.session_state.validated and checks[entity]:
            result = st.session_state.val_results.get(entity)
            if result is None:
                pass
            elif file_count == 0:
                st.markdown("⚠️")
            elif result["ok"]:
                st.markdown("✅")
            else:
                st.markdown("❌")

    # Sub-messages below entity row
    if checks[entity]:
        if file_count == 0:
            st.markdown('<p class="no-files-msg">No files added</p>', unsafe_allow_html=True)
        elif st.session_state.validated:
            result = st.session_state.val_results.get(entity)
            if result is not None:
                if result["ok"]:
                    st.markdown('<p class="success-msg">All files are ready</p>', unsafe_allow_html=True)
                else:
                    tree_html = '<div class="error-tree">'
                    for fname in result["error_files"]:
                        tree_html += f'<p class="error-file">└─ {fname}</p>'
                    tree_html += '</div>'
                    st.markdown(tree_html, unsafe_allow_html=True)

st.divider()

# ── Validate Button ────────────────────────────────────────────────────────
selected = [e for e in ENTITIES if checks[e]]

if st.button("🔍 Validate", use_container_width=True):
    if not selected:
        st.warning("Please select at least one entity.")
    else:
        has_any_files = any(len(get_xlsx_files(e)) > 0 for e in selected)
        if not has_any_files:
            st.warning("No files found in any selected entity folder.")
        else:
            new_results = {}
            status_placeholder = st.empty()

            with st.spinner("Validating files..."):
                for entity in selected:
                    files = get_xlsx_files(entity)

                    if not files:
                        new_results[entity] = {"ok": False, "error_files": []}
                        continue

                    module_name = VALIDATE_MAP[entity]
                    try:
                        module      = import_fresh(module_name)
                        error_files = []

                        for fname in files:
                            file_path = os.path.join(BASE_DIR, entity, fname)
                            status_placeholder.markdown(f"⚙️ Validating **{entity}/{fname}**...")
                            has_errors, _ = module.validate(file_path)
                            if has_errors:
                                error_files.append(fname)

                        new_results[entity] = {
                            "ok":          len(error_files) == 0,
                            "error_files": error_files,
                        }

                    except ImportError as e:
                        st.error(f"❌ Validation script not found for {entity}: {e}")
                        new_results[entity] = {"ok": False, "error_files": ["Script missing"]}

            status_placeholder.empty()
            st.session_state.val_results = new_results
            st.session_state.validated   = True
            st.rerun()

# ── Load Button ────────────────────────────────────────────────────────────
st.markdown("<div style='margin-top: 12px;'>", unsafe_allow_html=True)

# Determine if load should be active
all_green = (
    st.session_state.validated
    and len(selected) > 0
    and all(
        st.session_state.val_results.get(e, {}).get("ok", False)
        for e in selected
    )
    and all(len(get_xlsx_files(e)) > 0 for e in selected)
)

if all_green:
    load_clicked = st.button("🚀 Load", use_container_width=True, type="primary")
else:
    st.markdown("""
    <div style="
        background-color: #3a3a3a;
        color: #777;
        text-align: center;
        padding: 12px;
        border-radius: 8px;
        font-size: 1rem;
        cursor: not-allowed;
        user-select: none;
        margin-bottom: 8px;
    ">
        🔒 Load — validate all entities first
    </div>
    """, unsafe_allow_html=True)
    load_clicked = False

st.markdown("</div>", unsafe_allow_html=True)

# ── Load Logic ─────────────────────────────────────────────────────────────
if load_clicked:
    all_results = {}
    status_placeholder = st.empty()

    with st.spinner("Loading data into QAD..."):
        for entity in selected:
            files = get_xlsx_files(entity)
            entity_results = []

            module_name = LOAD_MAP[entity]
            try:
                module = import_fresh(module_name)

                for fname in files:
                    file_path = os.path.join(BASE_DIR, entity, fname)
                    status_placeholder.markdown(f"⚙️ Loading **{entity}/{fname}**...")
                    try:
                        success, fail = module.run(file_path)
                        if fail == 0:
                            archive_file(entity, file_path)
                            note = "✅ Archived"
                        else:
                            note = "⚠️ Kept — fix red rows & re-run"
                        entity_results.append({
                            "file":    fname,
                            "success": success,
                            "fail":    fail,
                            "note":    note,
                        })
                    except Exception as e:
                        entity_results.append({
                            "file":    fname,
                            "success": 0,
                            "fail":    "?",
                            "note":    f"❌ Error: {e}",
                        })

            except ImportError as e:
                entity_results.append({
                    "file":    "—",
                    "success": 0,
                    "fail":    "?",
                    "note":    f"❌ Script not found: {e}",
                })

            all_results[entity] = entity_results

    status_placeholder.empty()

    # Reset validation state after load
    st.session_state.val_results = {}
    st.session_state.validated   = False

    # Determine overall outcome
    total_fail = sum(r["fail"] for res in all_results.values() for r in res if str(r["fail"]).isdigit())
    if total_fail == 0:
        st.success(" Load complete — all records loaded successfully!")
    else:
        st.warning(f"⚠️ Load finished with failures — {total_fail} row(s) could not be loaded. Check the summary and fix red rows.")
    st.divider()
    st.subheader("Summary")

    for entity, results in all_results.items():
        st.markdown(f"** {entity}**")
        h = st.columns([3, 1, 1, 3])
        h[0].markdown("**File**")
        h[1].markdown("** Added**")
        h[2].markdown("** Failed**")
        h[3].markdown("**Status**")

        for r in results:
            cols = st.columns([3, 1, 1, 3])
            cols[0].write(r["file"])
            cols[1].write(str(r["success"]))
            cols[2].write(str(r["fail"]))
            cols[3].write(r["note"])

        st.markdown("")