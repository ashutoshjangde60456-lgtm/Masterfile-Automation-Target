# app_masterfile.py  — Auto-map version (no JSON required, with review UI)

import io
import json
import re
from difflib import SequenceMatcher
from textwrap import dedent

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(page_title="Masterfile Automation", page_icon="📦", layout="wide")

# =========================
# Helpers
# =========================
def norm(s: str) -> str:
    """Normalize header strings for robust matching."""
    if s is None:
        return ""
    x = str(s).strip().lower()
    x = re.sub(r"\s*-\s*en\s*[-_ ]\s*us\s*$", "", x)
    x = x.replace("–", "-").replace("—", "-").replace("−", "-")
    x = re.sub(r"[._/\\-]+", " ", x)
    x = re.sub(r"[^0-9a-z\s]+", " ", x)
    return re.sub(r"\s+", " ", x).strip()

def nonempty_rows(df: pd.DataFrame) -> int:
    if df.empty:
        return 0
    return df.replace("", pd.NA).dropna(how="all").shape[0]

def worksheet_used_cols(ws, header_rows=(1,), hard_cap=512, empty_streak_stop=8):
    max_try = min(ws.max_column, hard_cap)
    last_nonempty, streak = 0, 0
    for c in range(1, max_try + 1):
        any_val = any((ws.cell(row=r, column=c).value not in (None, "")) for r in header_rows)
        if any_val:
            last_nonempty, streak = c, 0
        else:
            streak += 1
            if streak >= empty_streak_stop:
                break
    return max(last_nonempty, 1)

def pick_best_onboarding_sheet(uploaded_file):
    uploaded_file.seek(0)
    xl = pd.ExcelFile(uploaded_file)
    best = None
    best_score = -1
    best_info = ""
    for sheet in xl.sheet_names:
        try:
            df = xl.parse(sheet_name=sheet, header=0, dtype=str).fillna("")
            df.columns = [str(c).strip() for c in df.columns]
        except Exception:
            continue
        rows = nonempty_rows(df)
        score = rows  # prefer non-empty
        if score > best_score:
            best, best_score = (df, sheet), score
            best_info = f"non-empty rows: {rows}"
    if best is None:
        raise ValueError("No readable sheet found in onboarding workbook.")
    return best[0], best[1], best_info

# ---------- Heuristic Auto-Mapping ----------
# Small alias bank (extend as needed). Keys are canonical master display headers.
ALIAS_BANK = {
    "partner sku": ["partner sku", "seller sku", "target sku", "walmart sku", "item sku", "sku", "item_sku", "productcode", "product code"],
    "item sku": ["item sku", "sku", "productcode", "product code", "seller sku"],
    "barcode": ["barcode", "upc", "upc/ean", "ean", "gtin", "external product id", "product id", "barcode value", "gtin13", "gtin-13", "gtin14", "gtin-14"],
    "brand": ["brand", "brand name", "manufacturer brand"],
    "product title": ["product title", "title", "item name", "product name", "item title", "name"],
    "description": ["description", "long description", "product description", "item description", "desc"],
    "package depth": ["package depth", "depth", "length", "package length"],
    "package height": ["package height", "height"],
    "package width": ["package width", "width"],
    "package weight": ["package weight", "weight", "shipping weight"],
    "country of origin": ["country of origin", "made in", "madein", "origin country", "countryoforigin"],
    "price": ["price", "selling price", "sellingprice", "mrp", "msrp"],
    "mrp": ["mrp", "msrp", "maximum retail price"],
    "hsn": ["hsn", "hs code", "hsn code", "hs code number"],
    "tax code": ["tax code", "taxcode"],
    "primary color": ["primary color", "color"],
    "size": ["size", "age size", "apparel size"],
    "keywords": ["keywords", "search keywords", "search terms", "generic keywords"],
    "bullet feature 1": ["bullet feature 1", "bullet point1", "bullet_point1", "key feature 1"],
    "bullet feature 2": ["bullet feature 2", "bullet point2", "bullet_point2", "key feature 2"],
    "bullet feature 3": ["bullet feature 3", "bullet point3", "bullet_point3", "key feature 3"],
    "bullet feature 4": ["bullet feature 4", "bullet point4", "bullet_point4", "key feature 4"],
    "bullet feature 5": ["bullet feature 5", "bullet point5", "bullet_point5", "key feature 5"],
    "image1": ["image1", "image url 1", "imageurl1", "main image", "primary image"],
    "image2": ["image2", "image url 2", "imageurl2"],
    "image3": ["image3", "image url 3", "imageurl3"],
    "image4": ["image4", "image url 4", "imageurl4"],
    "image5": ["image5", "image url 5", "imageurl5"],
}

GENERIC_TOKENS = {"value", "name", "code", "id", "info", "data", "attribute", "field"}

def token_set(s: str):
    return {t for t in norm(s).split() if t and t not in GENERIC_TOKENS}

def fuzzy_ratio(a: str, b: str) -> float:
    return SequenceMatcher(None, norm(a), norm(b)).ratio()

def score_pair(master_disp: str, candidate_header: str) -> float:
    """0..1 confidence scoring combining alias hits, token overlap, and fuzzy."""
    m = norm(master_disp)
    c = norm(candidate_header)
    if not m or not c:
        return 0.0

    # 1) Exact/equivalent
    if m == c:
        return 1.00

    # 2) Alias bank boost
    aliases = ALIAS_BANK.get(m, [])
    if c in [norm(x) for x in aliases]:
        return 0.95

    # 3) Token overlap
    mt, ct = token_set(m), token_set(c)
    if mt and ct:
        j = len(mt & ct) / len(mt | ct)
    else:
        j = 0.0

    # 4) Startswith / Endswith small boosts
    starts = 0.1 if c.startswith(m) or m.startswith(c) else 0.0
    ends   = 0.05 if c.endswith(m) or m.endswith(c) else 0.0

    # 5) Fuzzy ratio
    fz = fuzzy_ratio(m, c)

    # Weighted sum (tuned conservatively)
    score = 0.45 * fz + 0.40 * j + starts + ends

    # 6) Penalize ultra-short collisions (e.g., "id", "name")
    if len(mt) == 1 and len(next(iter(mt))) <= 3:
        score -= 0.1

    return max(0.0, min(1.0, score))

def suggest_mapping(master_headers, onboarding_headers, extra_mapping_json=None):
    """
    Return dict: master_display -> (best_candidate_or_None, confidence, ranked_list)
    If extra_mapping_json is provided, treat as additional aliases with high priority.
    """
    results = {}
    on_norm_index = {norm(h): h for h in onboarding_headers}

    # Build dynamic alias additions from user JSON (optional)
    dynamic_alias = {}
    if isinstance(extra_mapping_json, dict):
        for k, v in extra_mapping_json.items():
            key = norm(k)
            vals = v[:] if isinstance(v, list) else [v]
            dynamic_alias[key] = [norm(x) for x in vals]

    for m in master_headers:
        m_norm = norm(m)
        ranked = []

        # If user JSON explicitly maps this header and a candidate exists, boost it strongly
        if m_norm in dynamic_alias:
            for alias_n in dynamic_alias[m_norm]:
                if alias_n in on_norm_index:
                    ranked.append((on_norm_index[alias_n], 0.99))

        # Score all onboarding headers
        for cand in onboarding_headers:
            ranked.append((cand, score_pair(m, cand)))

        # Sort unique candidates by score desc
        seen = set()
        ranked = [(c, sc) for c, sc in sorted(ranked, key=lambda x: x[1], reverse=True) if not (c in seen or seen.add(c))]

        best_cand, best_sc = (ranked[0] if ranked else (None, 0.0))
        results[m] = (best_cand, best_sc, ranked[:10])  # keep top 10 for UI
    return results

# =========================
# UI
# =========================
st.title("📦 Masterfile Automation")
st.caption("Auto-map onboarding headers to master template with review & confirmation.")

with st.expander("ℹ️ Instructions", expanded=False):
    st.markdown(dedent("""
    - **Masterfile template (.xlsx)**  
      Row 1 = display labels, Row 2 = internal keys (if present). Data is written from Row 3.
    - **Onboarding sheet (.xlsx)**  
      Row 1 = headers, Row 2+ = data.
    - You can run **Auto-map (no JSON)** or optionally add **Mapping JSON** as hints; you will always be able to review and adjust mappings before writing.
    """))

st.divider()

colA, colB = st.columns([1, 1])
with colA:
    masterfile_file = st.file_uploader("📄 Upload Masterfile Template (.xlsx)", type=["xlsx"])
with colB:
    onboarding_file = st.file_uploader("🧾 Upload Onboarding Sheet (.xlsx)", type=["xlsx"])

st.markdown("#### 🔗 Mapping Options")
left, right = st.columns([1, 1])
with left:
    use_automap = st.checkbox("Auto-map headers (no JSON needed)", value=True)
with right:
    mapping_json_text = st.text_area(
        "Optional: Mapping JSON (used as strong hints)",
        height=160,
        placeholder='{\n  "Partner SKU": ["Seller SKU", "item_sku"]\n}',
    )

st.divider()
go = st.button("🔎 Build Mapping Suggestions", type="primary")

# Store user choices across steps
if "proposed_map" not in st.session_state:
    st.session_state.proposed_map = None
if "on_headers" not in st.session_state:
    st.session_state.on_headers = []
if "master_headers" not in st.session_state:
    st.session_state.master_headers = []
if "master_ws_meta" not in st.session_state:
    st.session_state.master_ws_meta = None
if "on_df" not in st.session_state:
    st.session_state.on_df = None

# =========================
# Step 1: Propose Mapping
# =========================
if go:
    if not masterfile_file or not onboarding_file:
        st.error("Please upload both **Masterfile Template** and **Onboarding Sheet**.")
        st.stop()

    # Load master (styles preserved later)
    try:
        master_wb = load_workbook(masterfile_file, keep_links=False)
        master_ws = master_wb.active
    except Exception as e:
        st.error(f"Could not read **Masterfile**: {e}")
        st.stop()

    # Pick best onboarding sheet
    try:
        best_df, best_sheet, info = pick_best_onboarding_sheet(onboarding_file)
        st.success(f"Using onboarding sheet: **{best_sheet}** ({info})")
    except Exception as e:
        st.error(f"Could not read **Onboarding**: {e}")
        st.stop()

    on_df = best_df.fillna("")
    on_df.columns = [str(c).strip() for c in on_df.columns]
    on_headers = list(on_df.columns)

    used_cols = worksheet_used_cols(master_ws, header_rows=(1, 2))
    master_displays = [master_ws.cell(row=1, column=c).value or "" for c in range(1, used_cols + 1)]

    # Optional JSON hints
    user_json = None
    if mapping_json_text.strip():
        try:
            user_json = json.loads(mapping_json_text)
        except Exception as e:
            st.error(f"Mapping JSON could not be parsed. Error: {e}")
            st.stop()

    # Build suggestions
    mapping_suggestions = suggest_mapping(master_displays, on_headers, extra_mapping_json=(user_json if use_automap else None))

    # Save for step 2
    st.session_state.master_headers = master_displays
    st.session_state.on_headers = on_headers
    st.session_state.proposed_map = mapping_suggestions
    st.session_state.master_ws_meta = (master_wb, master_ws)
    st.session_state.on_df = on_df

# =========================
# Step 2: Review & Confirm
# =========================
if st.session_state.proposed_map is not None:
    st.markdown("### ✅ Review & Confirm Column Mapping")
    on_headers = st.session_state.on_headers
    proposed = st.session_state.proposed_map

    # Build UI selects
    chosen_map = {}
    low_conflicts = []

    for m in st.session_state.master_headers:
        best_cand, conf, ranked = proposed.get(m, (None, 0.0, []))
        # default to best candidate if any
        default_idx = 0
        options = ["(leave blank)"] + on_headers
        if best_cand and best_cand in on_headers:
            default_idx = options.index(best_cand) if best_cand in options else 0

        col1, col2 = st.columns([1.2, 2.0])
        with col1:
            st.markdown(f"**{m or '(empty header)'}**")
            choice = st.selectbox(
                "Select onboarding column",
                options=options,
                index=default_idx,
                key=f"sel_{m}_{default_idx}"
            )
        with col2:
            # show top candidates with confidence
            tips = ", ".join([f"{c} ({int(s*100)}%)" for c, s in ranked[:5]])
            st.caption(f"Suggested: {best_cand or '—'} • Confidence: {int(conf*100)}%")
            if tips:
                st.caption(f"Top matches: {tips}")

        if choice != "(leave blank)":
            chosen_map[m] = choice
        if conf < 0.65 and choice != "(leave blank)":
            low_conflicts.append(m)

    if low_conflicts:
        st.warning("Low confidence mappings detected for: " + ", ".join([f"`{m}`" for m in low_conflicts]))

    st.divider()
    confirm = st.button("🚀 Generate Final Masterfile", type="primary")

    if confirm:
        # Write with preserved styles
        master_wb, master_ws = st.session_state.master_ws_meta
        on_df = st.session_state.on_df
        used_cols = worksheet_used_cols(master_ws, header_rows=(1, 2))
        master_displays = [master_ws.cell(row=1, column=c).value or "" for c in range(1, used_cols + 1)]

        # Build mapping col -> Series
        series_by_alias = {h: on_df[h] for h in on_df.columns}
        out_row_start = 3
        num_rows = len(on_df)

        for c, m_disp in enumerate(master_displays, start=1):
            chosen = chosen_map.get(m_disp, None)
            if not chosen:
                continue
            src = series_by_alias.get(chosen)
            if src is None:
                continue
            for i in range(num_rows):
                val = "" if i >= len(src) else str(src.iloc[i])
                master_ws.cell(row=out_row_start + i, column=c, value=val)

        bio = io.BytesIO()
        master_wb.save(bio)
        bio.seek(0)

        st.success("✅ Final masterfile is ready!")
        st.download_button(
            "⬇️ Download Final Masterfile",
            data=bio.getvalue(),
            file_name="final_masterfile.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
