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

st.set_page_config(page_title="GetGreenr Bulk Stock Update", page_icon="📦", layout="wide")


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
    st.title("🔒 Mister Mobile · GetGreenr Stock Tool")
    pw = st.text_input("Team password", type="password")
    if pw:
        if pw == expected:
            st.session_state["auth_ok"] = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    st.stop()


check_password()

st.title("📦 GetGreenr Bulk Stock Update")
st.caption("Locked → summed Used qty · Accessories → 10 · Match Review → reviewer "
           "decision (Set to 0 / else unchanged) · everything else unchanged. "
           "Original file format is preserved for re-upload.")

st.header("1 · Upload files")
c1, c2 = st.columns(2)
with c1:
    masterlist_file = st.file_uploader("1. Masterlist Excel (live Used stock)", type=["xlsx", "xls"])
    registry_file = st.file_uploader("3. SKU Registry workbook (Locked + Match Review)", type=["xlsx", "xls"])
with c2:
    getgreenr_file = st.file_uploader("2. GetGreenr bulk stock (file to update)", type=["csv", "xlsx", "xls"])
    new_skus_file = st.file_uploader("4. New Masterlist SKUs (optional)", type=["xlsx", "xls"])

ready = all([masterlist_file, registry_file, getgreenr_file])

# ---- Stock column override ---- #
stock_override = None
sku_override = None
if getgreenr_file is not None:
    try:
        getgreenr_file.seek(0)
        name = getgreenr_file.name.lower()
        peek = (pd.read_csv(getgreenr_file, nrows=5, dtype=object) if name.endswith(".csv")
                else pd.read_excel(getgreenr_file, nrows=5, dtype=object))
        getgreenr_file.seek(0)
        cols = list(peek.columns)
        auto_stock = C.find_col(cols, C.STOCK_CANDS)
        auto_sku = C.find_col(cols, C.SKU_CANDS)
        with st.sidebar:
            st.header("⚙️ Column mapping")
            stock_override = st.selectbox(
                "Seller Stock column (to update)", cols,
                index=cols.index(auto_stock) if auto_stock in cols else 0,
                help=f"Auto-detected: {auto_stock}")
            sku_override = st.selectbox(
                "SKU ID column (match key)", cols,
                index=cols.index(auto_sku) if auto_sku in cols else 0,
                help=f"Auto-detected: {auto_sku}")
    except Exception as e:  # noqa: BLE001
        st.sidebar.warning(f"Could not preview columns: {e}")

st.header("2 · Run update")
if not ready:
    st.info("Upload the Masterlist, GetGreenr bulk stock and SKU Registry to enable the update.")
else:
    if st.button("🚀 Apply rules to Seller Stock", type="primary"):
        with st.spinner("Applying rules…"):
            for f in (masterlist_file, registry_file, getgreenr_file, new_skus_file):
                if f is not None:
                    f.seek(0)
            res = C.run(getgreenr_file, masterlist_file, registry_file, new_skus_file,
                        stock_col=stock_override, sku_col=sku_override,
                        filename=getgreenr_file.name)

        s = res["summary"]
        st.subheader("✅ Validation summary")
        m = st.columns(4)
        m[0].metric("Listings", s["GetGreenr listings"])
        m[1].metric("Locked (synced)", s["Locked (synced to Used qty)"])
        m[2].metric("Accessories → 10", s["Accessories set to 10"])
        m[3].metric("Review → 0", s["Match Review set to 0"])
        m2 = st.columns(4)
        m2[0].metric("Review unchanged", s["Match Review left unchanged"])
        m2[1].metric("Unmatched unchanged", s["Unmatched left unchanged"])
        m2[2].metric("Warnings", s["Warnings"])
        m2[3].metric("Errors", s["Errors"])
        st.caption(f"Updated column: **{s['Stock column updated']}**  ·  matched on: **{s['SKU column used']}**")

        st.dataframe(pd.DataFrame(list(s.items()), columns=["Metric", "Value"]),
                     use_container_width=True, hide_index=True)

        # Errors
        st.subheader("⚠️ Error / warning report")
        if res["errors"]:
            edf = pd.DataFrame(res["errors"])
            st.dataframe(edf, use_container_width=True, hide_index=True)
        else:
            st.success("No issues.")

        # Changed rows
        pv = res["preview"]
        if pv is not None and not pv.empty:
            with st.expander(f"Changed listings ({int(pv['Changed'].sum())})", expanded=True):
                st.dataframe(pv[pv["Changed"]], use_container_width=True, hide_index=True)

        # Downloads
        st.subheader("3 · Download")
        d1, d2 = st.columns(2)
        if res["out_bytes"] is not None:
            mime = ("text/csv" if res["out_name"].endswith(".csv")
                    else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            d1.download_button("⬇️ Updated GetGreenr file (ready to upload)",
                               data=res["out_bytes"], file_name=res["out_name"],
                               mime=mime, type="primary")
        # report
        rbuf = io.BytesIO()
        with pd.ExcelWriter(rbuf, engine="openpyxl") as xl:
            pd.DataFrame(list(s.items()), columns=["Metric", "Value"]).to_excel(xl, "Summary", index=False)
            (pd.DataFrame(res["errors"]) if res["errors"] else pd.DataFrame(columns=["level"])).to_excel(xl, "Errors", index=False)
            pv.to_excel(xl, "Per-Listing Changes", index=False)
        d2.download_button("⬇️ Update report", data=rbuf.getvalue(),
                           file_name="getgreenr_update_report.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
