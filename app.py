"""
KPay Monthly Reconciliation — Streamlit Web App
================================================
Run with:  streamlit run app.py
"""

import os
import sys
import tempfile
from datetime import datetime

import streamlit as st
from openpyxl import Workbook

# ── Import reconciliation engine ──────────────────────────────────────────────
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "tools"))
from reconcile_kpay import (
    read_pos, read_kpay, read_dbs,
    build_sheet, validate_output, determine_methods,
)

# ── Helper ────────────────────────────────────────────────────────────────────
def _write_upload(tmpdir: str, uploaded_file, filename: str) -> str:
    path = os.path.join(tmpdir, filename)
    with open(path, "wb") as f:
        f.write(uploaded_file.getvalue())
    return path


# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="KPay Reconciliation",
    page_icon="📊",
    layout="centered",
)

# ── Header ────────────────────────────────────────────────────────────────────
st.title("📊 KPay Monthly Reconciliation")
st.caption(
    "Upload your three source files, set the month, and click **Generate Report** "
    "to produce the reconciliation Excel file."
)

st.divider()

# ── Step 1: File uploads ──────────────────────────────────────────────────────
st.subheader("Step 1 — Upload Source Files")

pos_file  = st.file_uploader("POS Sales File  (.xlsx)",          type=["xlsx", "xls"],  key="pos")
kpay_file = st.file_uploader("KPay Transaction Report  (.xlsx)", type=["xlsx", "xls"],  key="kpay")
dbs_file  = st.file_uploader("DBS Bank Statement  (.xls)",       type=["xls", "xlsx"],  key="dbs")

uploaded = [f for f in [pos_file, kpay_file, dbs_file] if f is not None]
if uploaded:
    st.caption(f"{len(uploaded)}/3 files uploaded")

st.divider()

# ── Step 2: Report parameters ─────────────────────────────────────────────────
st.subheader("Step 2 — Report Parameters")

# Default month: previous month
_now = datetime.now()
_default_month = (_now.month - 2) % 12 + 1     # 1-indexed
_default_year  = _now.year if _now.month > 1 else _now.year - 1

MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]

col_m, col_y, col_s, col_mid = st.columns([2, 1.5, 1.5, 1.5])
with col_m:
    month_name = st.selectbox("Month", MONTH_NAMES, index=_default_month - 1)
with col_y:
    year = st.selectbox("Year", list(range(2024, 2028)), index=list(range(2024, 2028)).index(_default_year))
with col_s:
    store_code = st.text_input("Store Code", value="KYR")
with col_mid:
    merchant_id = st.text_input("Merchant ID", value="146100005")

month = MONTH_NAMES.index(month_name) + 1
month_str = f"{year}{month:02d}"
output_name = f"KPay_{month_name[:3]}{year}_Reconciliation.xlsx"

st.divider()

# ── Step 3: Generate ──────────────────────────────────────────────────────────
st.subheader("Step 3 — Generate")

all_ready = all([pos_file, kpay_file, dbs_file, store_code, merchant_id])
if not all_ready:
    missing = []
    if not pos_file:   missing.append("POS Sales File")
    if not kpay_file:  missing.append("KPay Transaction Report")
    if not dbs_file:   missing.append("DBS Bank Statement")
    st.info(f"Waiting for: {', '.join(missing)}" if missing else "Fill in all parameters above.")

if st.button("Generate Reconciliation Report", type="primary", disabled=not all_ready):

    output_bytes = None
    passed       = False
    val_results  = []

    with tempfile.TemporaryDirectory() as tmpdir:
        # Save uploaded files to temp disk paths (required by openpyxl / xlrd)
        pos_path  = _write_upload(tmpdir, pos_file,  pos_file.name)
        kpay_path = _write_upload(tmpdir, kpay_file, kpay_file.name)
        dbs_path  = _write_upload(tmpdir, dbs_file,  dbs_file.name)
        out_path  = os.path.join(tmpdir, output_name)

        try:
            with st.status("Processing…", expanded=True) as status:

                # 1. POS
                st.write(f"Reading POS data for **{store_code}**…")
                pos_daily, _ = read_pos(pos_path, store_code)
                st.write(f"  → {len(pos_daily)} trading days found")

                # 2. KPay
                st.write("Reading KPay transactions…")
                kpay_data, settle_dates, methods_found = read_kpay(kpay_path)
                methods, rare_methods = determine_methods(methods_found)
                total_gross = sum(
                    v["gross"]
                    for d in kpay_data.values()
                    for v in d.values()
                )
                st.write(f"  → {len(settle_dates)} settle dates, "
                         f"{len(methods)} payment methods, "
                         f"gross total = HKD {total_gross:,.2f}")
                if rare_methods:
                    st.write(f"  → Rare methods (shown separately): {', '.join(rare_methods)}")

                # 3. DBS
                st.write("Reading DBS bank statement…")
                dbs_by_date, dbs_batches = read_dbs(dbs_path, merchant_id, year, month)
                dbs_total = sum(dbs_by_date.values())
                st.write(f"  → {len(dbs_by_date)} credit dates, "
                         f"total = HKD {dbs_total:,.2f}")

                # 4. Build workbook
                st.write("Building Excel workbook…")
                wb = Workbook()
                wb.remove(wb.active)
                build_sheet(wb, pos_daily, kpay_data, settle_dates,
                            dbs_by_date, store_code, year, month,
                            methods, rare_methods)
                wb.save(out_path)

                # 5. Validate
                st.write("Validating output…")
                passed, val_results = validate_output(out_path, year, month)

                # Read output bytes while file still exists
                with open(out_path, "rb") as f:
                    output_bytes = f.read()

                if passed:
                    status.update(label="Reconciliation complete ✓", state="complete")
                else:
                    status.update(label="Completed — some checks failed", state="error")

        except Exception as e:
            st.error("An error occurred during processing.")
            st.exception(e)
            st.stop()

    # ── Validation results ────────────────────────────────────────────────────
    st.subheader("Validation Results")
    for name, ok, detail in val_results:
        msg = f"**{name}**" + (f" — {detail}" if detail else "")
        if ok:
            st.success(msg)
        else:
            st.error(msg)

    st.divider()

    # ── Download ──────────────────────────────────────────────────────────────
    if output_bytes:
        if passed:
            st.success(f"Report ready: **{output_name}**")
        else:
            st.warning(f"Report generated with warnings: **{output_name}**")

        st.download_button(
            label=f"📥  Download  {output_name}",
            data=output_bytes,
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
