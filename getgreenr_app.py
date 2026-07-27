"""
getgreenr_app.py — Streamlit app: apply the stock rules to GetGreenr Seller Stock.

Run:
    pip install -r requirements.txt
    streamlit run getgreenr_app.py

Upload the four files, review the auto-detected Seller Stock column, click Run,
then download the updated GetGreenr file (ready to re-upload) + the report.

Rules (keyed by GetGreenr SKU ID):
  Locked Matches   -> SUM of live Used Available Qty across linked Masterlist IDs
  Accessories      -> 10
  Match Review     -> "Set to 0"/"Skip / Delist" => 0, otherwise unchanged
  Anything else    -> unchanged
"""
from __future__ import annotations

import io
import pandas as pd
import streamlit as st

import getgreenr_core as C
import registry_report as RR

st.set_page_config(page_title="Mister Mobile · GetGreenr Stock", page_icon="📦", layout="wide")

# Mister Mobile brand: Yellow #FFEB00 (hero) · Black · Gray #6D6962 · Cream #F9F4E1
MM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@600;700;800&family=Open+Sans:wght@400;600&display=swap');

html, body, [class*="css"], .stMarkdown, p, label, div { font-family: 'Open Sans', sans-serif; }
h1, h2, h3, h4 { font-family: 'Montserrat', sans-serif !important; font-weight: 800 !important; color:#000; letter-spacing:-0.5px; }

/* MM header bar */
.mm-header { background:#FFEB00; border-radius:14px; padding:20px 26px; margin:0 0 22px 0;
  display:flex; align-items:center; gap:16px; }
.mm-badge { background:#000; color:#FFEB00; font-family:'Montserrat'; font-weight:800; font-size:22px;
  width:52px; height:52px; border-radius:12px; display:flex; align-items:center; justify-content:center; }
.mm-brand { font-family:'Montserrat'; font-weight:800; font-size:26px; color:#000; line-height:1.05; }
.mm-sub { font-family:'Open Sans'; font-weight:600; font-size:13px; color:#3d3a35; }

/* Primary button: black on yellow */
.stButton > button[kind="primary"] { background:#FFEB00 !important; color:#000 !important;
  border:2px solid #000 !important; font-family:'Montserrat'; font-weight:700; border-radius:10px; }
.stButton > button[kind="primary"]:hover { background:#000 !important; color:#FFEB00 !important; }

/* Download buttons */
.stDownloadButton > button { border:2px solid #000 !important; border-radius:10px; font-weight:700; }

/* Metric accents */
[data-testid="stMetricValue"] { font-family:'Montserrat'; font-weight:800; }
[data-testid="stMetric"] { border-left:5px solid #FFEB00; padding-left:12px; background:#F9F4E1; border-radius:8px; }

/* Numbered section headers get a yellow underline */
h2 { border-bottom:3px solid #FFEB00; padding-bottom:6px; }
</style>
"""


def brand_header():
    st.markdown(MM_CSS, unsafe_allow_html=True)
    st.markdown(
        '<div class="mm-header">'
        '<div class="mm-badge">MM</div>'
        '<div><div class="mm-brand">GetGreenr Stock Sync</div>'
        '<div class="mm-sub">Mister Mobile · GetGreenr Inventory Bulk Update</div></div>'
        '</div>', unsafe_allow_html=True)


def check_password():
    """Simple shared-password gate for the public URL.

    Set APP_PASSWORD in Streamlit Cloud -> Settings -> Secrets. If no password is
    configured (e.g. running locally), the app stays open.
    """
    try:
        expected = st.secrets["APP_PASSWORD"]
    except Exception:
        expected = None
    if not expected:
        return  # no password set -> open access
    if st.session_state.get("auth_ok"):
        return
    st.markdown(MM_CSS, unsafe_allow_html=True)
    st.markdown('<div class="mm-header"><div class="mm-badge">MM</div>'
                '<div><div class="mm-brand">Mister Mobile</div>'
                '<div class="mm-sub">GetGreenr Stock Tool · Team access</div></div></div>',
                unsafe_allow_html=True)
    pw = st.text_input("Team password", type="password")
    if pw:
        if pw == expected:
            st.session_state["auth_ok"] = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    st.stop()


class NamedBytes(io.BytesIO):
    """BytesIO that reliably carries a .name (plain BytesIO can't hold attributes)."""
    def __init__(self, data, name):
        super().__init__(data)
        self.name = name


def buf(uploaded):
    """Fresh, rewound in-memory copy of an uploaded file (keeps the filename)."""
    return NamedBytes(uploaded.getvalue(), uploaded.name)


check_password()
brand_header()

# ---------------------------------------------------------------- How to use #
st.subheader("How to use")
st.markdown(
    "1. **Upload** the Masterlist and the GetGreenr export below.\n"
    "2. **Download the Match Review registry**, open the **New Masterlist SKUs** tab, and set a "
    "**Reviewer Decision** for *every* row (Linked + GetGreenr SKU ID, or Not on GetGreenr Yet / "
    "Not Selling in GetGreenr).\n"
    "3. **Re-upload** the completed registry (box 3). The ready-to-upload GetGreenr file stays "
    "locked until **all New Masterlist SKUs are reviewed**.\n"
    "4. **Download** the GetGreenr file and upload it in GetGreenr → Inventory → Bulk Update."
)
with st.expander("How matching works (GetGreenr rules)"):
    st.markdown(
        "- GetGreenr grade **Excellent** = Masterlist **Used** condition — only Used stock is synced.\n"
        "- **Locked Matches:** Seller Stock = SUM of live Used Available Qty across the linked Masterlist IDs.\n"
        "- **Match Review:** reviewer decision — *Set to 0 / Skip · Delist* → 0, otherwise unchanged.\n"
        "- **New Masterlist SKUs → Linked:** the linked Masterlist ID's Used qty flows to the chosen GetGreenr SKU.\n"
        "- Accessories → 10. Anything unmatched is left unchanged."
    )

# ---------------------------------------------------------------- Uploads #
u1, u2 = st.columns(2)
with u1:
    masterlist_file = st.file_uploader("① Masterlist (stock_report*.xlsx)", type=["xlsx", "xls"])
with u2:
    getgreenr_file = st.file_uploader("② GetGreenr export (inventory*.csv / .xlsx)", type=["csv", "xlsx", "xls"])
registry_file = st.file_uploader(
    "③ Reviewed registry — re-upload your completed GetGreenr_Match_Review.xlsx to carry your "
    "confirmed links & decisions forward", type=["xlsx", "xls"])

if not (masterlist_file and getgreenr_file):
    st.info("Upload the Masterlist and the GetGreenr export to begin.")
    st.stop()

# ---------------------------------------------------------------- Phase A: registry #
with st.spinner("Matching against the Masterlist…"):
    reg_bytes, counts = RR.build_registry_bytes(buf(getgreenr_file), buf(masterlist_file),
                                                gg_name=getgreenr_file.name)

st.divider()
st.subheader("Step 2 · Review the registry")
m = st.columns(3)
m[0].metric("Locked Matches", counts["locked"])
m[1].metric("Match Review", counts["review"])
m[2].metric("New Masterlist SKUs", counts["new_ml"])
st.download_button("⬇️ Download Match Review registry", data=reg_bytes,
                   file_name="GetGreenr_Match_Review.xlsx", type="primary",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
st.caption("Open the **New Masterlist SKUs** tab and set a Reviewer Decision for every row, then re-upload it in box ③.")

# ---------------------------------------------------------------- Phase B: gated apply #
st.divider()
st.subheader("Step 3 · Ready-to-upload GetGreenr file")

if registry_file is None:
    st.warning("🔒 Locked — re-upload your completed registry in box ③ to generate the upload file.")
    st.stop()

total, reviewed, missing = C.new_masterlist_status(buf(registry_file))
if missing > 0:
    st.error(f"🔒 Locked — {missing} of {total} New Masterlist SKUs still need a Reviewer Decision. "
             "Fill every row in the New Masterlist SKUs tab and re-upload.")
    st.progress(reviewed / total if total else 0.0, text=f"{reviewed}/{total} reviewed")
    st.stop()

st.success(f"✓ All {total} New Masterlist SKUs reviewed — file unlocked.")
with st.spinner("Building the upload file…"):
    res = C.run(buf(getgreenr_file), buf(masterlist_file), buf(registry_file),
                filename=getgreenr_file.name)

s = res["summary"]
g = st.columns(4)
g[0].metric("Listings", s["GetGreenr listings"])
g[1].metric("Locked (synced)", s["Locked (synced to Used qty)"])
g[2].metric("Review → 0", s["Match Review set to 0"])
g[3].metric("Errors", s["Errors"])
st.caption(f"Updated column: **{s['Stock column updated']}** · matched on: **{s['SKU column used']}**")

pv = res["preview"]
if pv is not None and not pv.empty:
    with st.expander(f"Changed listings ({int(pv['Changed'].sum())})"):
        st.dataframe(pv[pv["Changed"]], use_container_width=True, hide_index=True)

if res["out_bytes"] is not None:
    mime = ("text/csv" if res["out_name"].endswith(".csv")
            else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.download_button("⬇️ Download GetGreenr upload file", data=res["out_bytes"],
                       file_name=res["out_name"], mime=mime, type="primary")
    st.caption("Upload this in GetGreenr → Inventory → Bulk Update.")
