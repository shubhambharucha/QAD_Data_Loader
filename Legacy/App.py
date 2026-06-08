import streamlit as st
import os, sys, shutil
from datetime import datetime

BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
SCRIPTS_DIR = os.path.join(BASE_DIR, "Scripts")
sys.path.insert(0, SCRIPTS_DIR)

ENTITIES = [
    "Supplier", "Customer", "GCM",
    "BR", "Customer_Item", "ProductionOrder",
    "PurchaseOrder", "SalesOrder", "Supplier_Item",
    "SupplierPriceList"
]

VALIDATE_MAP = {
    "Supplier":         "validate_supplier",
    "Customer":         "validate_customer",
    "GCM":              "validate_gcm",
    "BR":               "validate_br",
    "Customer_Item":    "validate_cust_item",
    "ProductionOrder":  "validate_production_order",
    "PurchaseOrder":    "validate_po",
    "SalesOrder":       "validate_so",
    "Supplier_Item":    "validate_Supp_item",
    "SupplierPriceList":"validate_supp_price_lists",
}

LOAD_MAP = {
    "Supplier":         "Supplier_load",
    "Customer":         "Customer_load",
    "GCM":              "GCM_load",
    "BR":               "BR_load",
    "Customer_Item":    "cust_item",
    "ProductionOrder":  "production_order",
    "PurchaseOrder":    "po_load",
    "SalesOrder":       "so_load",
    "Supplier_Item":    "Supp_item",
    "SupplierPriceList":"supp_price_lists",
}

ENTITY_DESC = {
    "Supplier":         "Vendors, contacts, sites",
    "Customer":         "Accounts, addresses",
    "GCM":              "Global cost masters",
    "BR":               "Business relations",
    "Customer_Item":    "Customer item cross-refs",
    "ProductionOrder":  "Production order headers",
    "PurchaseOrder":    "PO headers & lines",
    "SalesOrder":       "SO headers & lines",
    "Supplier_Item":    "Supplier item cross-refs",
    "SupplierPriceList":"Supplier price lists",
}

ENTITY_ICON = {
    "Supplier":         "ti-user",
    "Customer":         "ti-users",
    "GCM":              "ti-file-description",
    "BR":               "ti-file-text",
    "Customer_Item":    "ti-box",
    "ProductionOrder":  "ti-settings",
    "PurchaseOrder":    "ti-shopping-cart",
    "SalesOrder":       "ti-receipt",
    "Supplier_Item":    "ti-package",
    "SupplierPriceList":"ti-tag",
}


def ensure_folders():
    for e in ENTITIES:
        os.makedirs(os.path.join(BASE_DIR, e), exist_ok=True)
        os.makedirs(os.path.join(BASE_DIR, "Archive", e), exist_ok=True)


def get_xlsx(entity):
    folder = os.path.join(BASE_DIR, entity)
    return [f for f in os.listdir(folder) if f.endswith(".xlsx")] if os.path.exists(folder) else []


def archive_file(entity, fp):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    shutil.move(fp, os.path.join(BASE_DIR, "Archive", entity,
        f"{os.path.splitext(os.path.basename(fp))[0]}_{ts}.xlsx"))


def apply_error_prefix(entity, filename):
    """Rename file to error_filename if not already prefixed."""
    folder = os.path.join(BASE_DIR, entity)
    if not filename.startswith("error_"):
        src = os.path.join(folder, filename)
        dst = os.path.join(folder, "error_" + filename)
        if os.path.exists(src):
            os.rename(src, dst)
        return "error_" + filename
    return filename


def remove_error_prefix(entity, filename):
    """Strip error_ prefix when file is clean."""
    folder = os.path.join(BASE_DIR, entity)
    if filename.startswith("error_"):
        src = os.path.join(folder, filename)
        dst = os.path.join(folder, filename[len("error_"):])
        if os.path.exists(src):
            os.rename(src, dst)
        return filename[len("error_"):]
    return filename


def import_fresh(mod):
    sys.modules.pop(mod, None)
    return __import__(mod)


st.set_page_config(page_title="QAD Data Loader", layout="centered", initial_sidebar_state="collapsed")
ensure_folders()

for k, v in [("val_results", {}), ("validated", False), ("sel", set())]:
    if k not in st.session_state:
        st.session_state[k] = v

# ── Tabler icons ──────────────────────────────────────────────────────────
st.markdown(
    '<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@3.19.0/dist/tabler-icons.min.css">',
    unsafe_allow_html=True
)

# ── Global styles ─────────────────────────────────────────────────────────
st.markdown("""
<style>
html, body,
[data-testid="stApp"],
[data-testid="stAppViewContainer"],
[data-testid="stAppViewContainer"] > .main {
    background-color: #0b1929 !important;
}
[data-testid="stHeader"],
[data-testid="stToolbar"],
footer, #MainMenu { display: none !important; }
.block-container {
    padding-top: 2.5rem !important;
    padding-bottom: 2rem !important;
    max-width: 860px !important;
}

/* ── Title ── */
.qad-title-wrap { text-align: center; margin-bottom: 28px; }
.qad-title {
    font-family: 'Segoe UI', sans-serif;
    font-size: 1.9rem;
    font-weight: 800;
    background: linear-gradient(90deg, #ffffff 0%, #1db890 60%, #00d4aa 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
    filter: drop-shadow(0 0 16px rgba(29,184,144,0.4));
    letter-spacing: 0.5px;
}

/* ── Panel ── */
.panel {
    background: linear-gradient(160deg, #0f2035 0%, #0d1c2e 100%);
    border-radius: 18px;
    border: 1px solid #1a2f45;
    padding: 28px 36px 28px 36px;
}
.panel-title {
    color: #ffffff;
    font-family: 'Segoe UI', sans-serif;
    font-size: 1.3rem;
    font-weight: 700;
    text-align: center;
    margin: 0 0 5px 0;
}
.panel-subtitle {
    color: #5a7a95;
    font-size: 0.78rem;
    text-align: center;
    margin: 0 0 22px 0;
}

/* ── Card visual: disable pointer events so clicks reach the button below ── */
div[data-testid="stMarkdownContainer"]:has(.entity-card-visual) {
    pointer-events: none !important;
}

/* ── Transparent overlay button ── */
div.entity-btn > div[data-testid="stButton"] > button,
div.entity-btn-sel > div[data-testid="stButton"] > button {
    background: transparent !important;
    border: none !important;
    box-shadow: none !important;
    color: transparent !important;
    font-size: 0 !important;
    line-height: 0 !important;
    padding: 0 !important;
    margin-top: -72px !important;
    height: 72px !important;
    width: 100% !important;
    cursor: pointer !important;
    position: relative !important;
    z-index: 10 !important;
    display: block !important;
    min-height: unset !important;
}
div.entity-btn > div[data-testid="stButton"] > button:focus,
div.entity-btn-sel > div[data-testid="stButton"] > button:focus,
div.entity-btn > div[data-testid="stButton"] > button:hover,
div.entity-btn-sel > div[data-testid="stButton"] > button:hover {
    background: transparent !important;
    border: none !important;
    outline: none !important;
    box-shadow: none !important;
}

/* ── Action buttons ── */
div.action-btn > div[data-testid="stButton"] > button {
    background-color: #0f1e2e !important;   
    color: #b8d4e8 !important;
    border: 1.5px solid #1e3348 !important;
    border-radius: 10px !important;
    font-weight: 300 !important;
    font: 
    font-size: 0.83rem !important;
    padding: 11px 0 !important;
    box-shadow: none !important;
    width: 100% !important;
    letter-spacing: 0.3px !important;
}
div.action-btn > div[data-testid="stButton"] > button:hover {
    background-color: #162534 !important;
    border-color: #2a4560 !important;
    color: #fff !important;
}

/* ── Green glow on both action buttons ── */
div.action-btn-green > div[data-testid="stButton"] > button {
    background: linear-gradient(135deg, #0d2e22 0%, #0a2419 100%) !important;
    color: #1db890 !important;
    border: 1.5px solid #1db890 !important;
    border-radius: 10px !important;
    font-weight: 700 !important;
    font-size: 0.82rem !important;
    padding: 11px 0 !important;
    box-shadow: 0 0 18px rgba(29,184,144,0.35), 0 0 6px rgba(29,184,144,0.2) !important;
    width: 100% !important;
    letter-spacing: 0.3px !important;
    transition: box-shadow 0.2s ease !important;
}
div.action-btn-green > div[data-testid="stButton"] > button:hover {
    background: linear-gradient(135deg, #0f3828 0%, #0c2b1e 100%) !important;
    box-shadow: 0 0 28px rgba(29,184,144,0.55), 0 0 10px rgba(29,184,144,0.3) !important;
    color: #2dffc0 !important;
}

.btn-disabled {
    background: #090f18;
    color: #253545;
    border: 1.5px solid #111d2a;
    border-radius: 10px;
    font-weight: 600;
    font-size: 0.82rem;
    padding: 12px 0;
    text-align: center;
    cursor: not-allowed;
}
.ok-msg {
    color: #1db890;
    font-size: 0.78rem;
    text-align: center;
    margin-top: 10px;
    letter-spacing: 0.2px;
}
.toast-error {
    position: fixed;
    bottom: 32px;
    left: 50%;
    transform: translateX(-50%);
    background: #1a0d0d;
    border: 1px solid #a33;
    color: #f08080;
    font-size: 0.8rem;
    font-weight: 500;
    padding: 12px 28px;
    border-radius: 10px;
    box-shadow: 0 4px 24px rgba(0,0,0,0.5);
    z-index: 9999;
    white-space: nowrap;
}
</style>
""", unsafe_allow_html=True)


# ── Card visual HTML ──────────────────────────────────────────────────────
def card_visual(entity):
    sel    = entity in st.session_state.sel
    files  = get_xlsx(entity)
    fcount = len(files)

    border = "#1db890" if sel else "#1a2f45"
    bg     = "#0c2219" if sel else "#0d1e30"
    glow   = "box-shadow:0 0 14px rgba(29,184,144,0.18);" if sel else ""

    if sel and st.session_state.validated:
        r = st.session_state.val_results.get(entity)
        if r is None:
            indicator = ""
        elif not files:
            indicator = (
                '<span style="position:absolute;top:10px;right:10px;'
                'width:8px;height:8px;border-radius:50%;background:#f0a500;'
                'box-shadow:0 0 6px #f0a500;display:inline-block"></span>'
            )
        elif r["ok"]:
            indicator = (
                '<span style="position:absolute;top:10px;right:10px;'
                'width:8px;height:8px;border-radius:50%;background:#1db890;'
                'box-shadow:0 0 6px #1db890;display:inline-block"></span>'
            )
        else:
            indicator = (
                '<span style="position:absolute;top:10px;right:10px;'
                'width:8px;height:8px;border-radius:50%;background:#e05252;'
                'box-shadow:0 0 6px #e05252;display:inline-block"></span>'
            )
    else:
        if fcount == 0:
            badge_bg    = "#1a1a2a"
            badge_color = "#3e5870"
            badge_text  = "No files"
        else:
            badge_bg    = "#0e2e22"
            badge_color = "#1db890"
            badge_text  = f"{fcount} file{'s' if fcount != 1 else ''}"

        indicator = (
            f'<span style="position:absolute;top:50%;right:10px;transform:translateY(-50%);'
            f'background:{badge_bg};color:{badge_color};font-size:0.62rem;font-weight:700;'
            f'padding:2px 7px;border-radius:20px;border:1px solid {badge_color}33;'
            f'white-space:nowrap;letter-spacing:0.3px">{badge_text}</span>'
        )

    icon_cls = ENTITY_ICON[entity]

    return (
        f'<div class="entity-card-visual" '
        f'style="background:{bg};border:1.5px solid {border};border-radius:12px;'
        f'padding:10px 12px;display:flex;align-items:center;gap:10px;'
        f'position:relative;min-height:62px;{glow}">'
        f'<div style="width:34px;height:34px;min-width:34px;background:#0e2e22;'
        f'border-radius:8px;display:flex;align-items:center;justify-content:center">'
        f'<i class="ti {icon_cls}" style="font-size:16px;color:#1db890;'
        f'filter:drop-shadow(0 0 5px rgba(29,184,144,0.6))"></i>'
        f'</div>'
        f'<div style="padding-right:60px">'
        f'<p style="color:#dceaf5;font-size:0.83rem;font-weight:700;margin:0 0 2px 0;'
        f'font-family:Segoe UI,sans-serif">{entity}</p>'
        f'<p style="color:#3e5870;font-size:0.69rem;margin:0;line-height:1.3">{ENTITY_DESC[entity]}</p>'
        f'</div>'
        f'{indicator}'
        f'</div>'
    )


# ── Render one entity card ────────────────────────────────────────────────
def render_card(entity, col):
    sel     = entity in st.session_state.sel
    btn_cls = "entity-btn-sel" if sel else "entity-btn"

    with col:
        st.markdown(card_visual(entity), unsafe_allow_html=True)
        st.markdown(f'<div class="{btn_cls}">', unsafe_allow_html=True)
        clicked = st.button(" Select ", key=f"c_{entity}", use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

        if clicked:
            if entity in st.session_state.sel:
                st.session_state.sel.discard(entity)
            else:
                st.session_state.sel.add(entity)
            st.session_state.validated   = False
            st.session_state.val_results = {}
            st.rerun()


# ── Title ─────────────────────────────────────────────────────────────────
st.markdown(
    '<div class="qad-title-wrap"><span class="qad-title">QAD Data Loader</span></div>',
    unsafe_allow_html=True
)

# ── MAIN PANEL ────────────────────────────────────────────────────────────
#st.markdown('<div class="panel">', unsafe_allow_html=True)

st.markdown(
    '<div class="panel-title">Select entities to load</div>'
    '<div class="panel-subtitle">Choose one or more, validate your files, then load when all checks pass.</div>',
    unsafe_allow_html=True
)

# Row 1 — Supplier, Customer, GCM
c1, c2, c3 = st.columns(3)
render_card("Supplier",  c1)
render_card("Customer",  c2)
render_card("GCM",       c3)

# Row 2 — BR, Customer_Item, ProductionOrder
c4, c5, c6 = st.columns(3)
render_card("BR",              c4)
render_card("Customer_Item",   c5)
render_card("ProductionOrder", c6)

# Row 3 — PurchaseOrder, SalesOrder, Supplier_Item
c7, c8, c9 = st.columns(3)
render_card("PurchaseOrder", c7)
render_card("SalesOrder",    c8)
render_card("Supplier_Item", c9)

# Row 4 — SupplierPriceList centered
_, c10, _ = st.columns([1, 1, 1])
render_card("SupplierPriceList", c10)

st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)

# ── Action buttons ────────────────────────────────────────────────────────
selected  = list(st.session_state.sel)
all_green = (
    st.session_state.validated
    and len(selected) > 0
    and all(st.session_state.val_results.get(e, {}).get("ok", False) for e in selected)
    and all(len(get_xlsx(e)) > 0 for e in selected)
)

b1, b2 = st.columns(2)

# Validate — always green glow
with b1:
    st.markdown('<div class="action-btn-green">', unsafe_allow_html=True)
    validate_clicked = st.button("Validate", key="vbtn", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

# Load — green glow only when ready, otherwise disabled
with b2:
    if all_green:
        st.markdown('<div class="action-btn-green">', unsafe_allow_html=True)
        load_clicked = st.button("Load", key="lbtn", use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="btn-disabled">Load</div>', unsafe_allow_html=True)
        load_clicked = False

if st.session_state.validated and all_green:
    st.markdown('<p class="ok-msg">✓ All entities validated — ready to load.</p>', unsafe_allow_html=True)

if st.session_state.validated and not all_green and selected:
    failed = [e for e in selected if not st.session_state.val_results.get(e, {}).get("ok", False)]
    if failed:
        st.markdown(
            f'<div class="toast-error">Validation failed for: {", ".join(failed)}. Fix errors and re-validate.</div>',
            unsafe_allow_html=True
        )

st.markdown('</div>', unsafe_allow_html=True)  # close .panel


# ── VALIDATE ──────────────────────────────────────────────────────────────
if validate_clicked:
    if not selected:
        st.warning("Select at least one entity to validate.")
    elif not any(get_xlsx(e) for e in selected):
        st.warning("No .xlsx files found in any selected entity folder.")
    else:
        results = {}
        ph = st.empty()
        with st.spinner("Validating…"):
            for entity in selected:
                files = get_xlsx(entity)
                if not files:
                    results[entity] = {"ok": False, "error_files": []}
                    continue
                try:
                    mod  = import_fresh(VALIDATE_MAP[entity])
                    errs = []
                    for f in files:
                        fp = os.path.join(BASE_DIR, entity, f)
                        ph.info(f"Validating {entity} › {f}")
                        has_err, _ = mod.validate(fp)
                        if has_err:
                            errs.append(f)
                            apply_error_prefix(entity, f)
                        else:
                            remove_error_prefix(entity, f)
                    results[entity] = {"ok": len(errs) == 0, "error_files": errs}
                except ImportError as e:
                    st.error(f"Script missing for {entity}: {e}")
                    results[entity] = {"ok": False, "error_files": ["Script missing"]}
        ph.empty()
        st.session_state.val_results = results
        st.session_state.validated   = True
        st.rerun()


# ── LOAD ──────────────────────────────────────────────────────────────────
if load_clicked:
    all_res = {}
    ph = st.empty()
    with st.spinner("Loading into QAD…"):
        for entity in selected:
            rows = []
            try:
                mod = import_fresh(LOAD_MAP[entity])
                for f in get_xlsx(entity):
                    fp = os.path.join(BASE_DIR, entity, f)
                    ph.info(f"Loading {entity} › {f}")
                    try:
                        ok, fail = mod.run(fp)
                        if fail == 0:
                            # clean — restore name if prefixed, then archive
                            restored = remove_error_prefix(entity, f)
                            archive_file(entity, os.path.join(BASE_DIR, entity, restored))
                            note = ("a", "Archived")
                        else:
                            apply_error_prefix(entity, f)
                            note = ("w", "Fix red rows & re-run")
                        rows.append({"file": f, "ok": ok, "fail": fail, "note": note})
                    except Exception as ex:
                        rows.append({"file": f, "ok": 0, "fail": "?", "note": ("e", str(ex)[:60])})
            except ImportError:
                rows.append({"file": "—", "ok": 0, "fail": "?", "note": ("e", "Script not found")})
            all_res[entity] = rows
    ph.empty()
    st.session_state.val_results = {}
    st.session_state.validated   = False
    st.session_state.sel         = set()

    total_fail = sum(
        r["fail"] for res in all_res.values()
        for r in res if isinstance(r["fail"], int)
    )
    if total_fail == 0:
        st.success("Load complete — all records loaded successfully.")
    else:
        st.warning(f"Load finished with {total_fail} failure(s).")

    st.divider()
    st.markdown("**Load Summary**")
    st.markdown("""
<style>
table.st{width:100%;border-collapse:collapse;font-size:.8rem;background:#0d1e30;border:1px solid #1a2f45;border-radius:10px;overflow:hidden;margin-bottom:14px}
table.st th{padding:8px 12px;background:#0a1828;color:#4a6278;font-weight:700;font-size:.7rem;letter-spacing:.5px;text-transform:uppercase;border-bottom:1px solid #1a2f45}
table.st td{padding:8px 12px;border-bottom:1px solid #121e2d;color:#c8dcea}
table.st tr:last-child td{border-bottom:none}
.na{color:#1db890;font-weight:600}.nw{color:#f0a500;font-weight:600}.ne{color:#e05252;font-weight:600}
.sum-title{font-size:.8rem;font-weight:700;color:#1db890;text-transform:uppercase;letter-spacing:1px;margin:18px 0 8px}
</style>
""", unsafe_allow_html=True)

    for entity, rows in all_res.items():
        st.markdown(f'<p class="sum-title">{entity}</p>', unsafe_allow_html=True)
        trs = ""
        for r in rows:
            nc = {"a": "na", "w": "nw", "e": "ne"}.get(r["note"][0], "")
            trs += (
                f"<tr><td><code>{r['file']}</code></td>"
                f"<td style='text-align:center'>{r['ok']}</td>"
                f"<td style='text-align:center'>{r['fail']}</td>"
                f"<td class='{nc}'>{r['note'][1]}</td></tr>"
            )
        st.markdown(
            f'<table class="st"><thead><tr>'
            f'<th>File</th><th style="text-align:center">Added</th>'
            f'<th style="text-align:center">Failed</th><th>Status</th>'
            f'</tr></thead><tbody>{trs}</tbody></table>',
            unsafe_allow_html=True
        )