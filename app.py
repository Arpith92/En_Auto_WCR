# app.py – Automated WCR Generator (Excel → Word Only, Streamlit)

from __future__ import annotations
import os, io, zipfile
from datetime import datetime
import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate

# ==============================
# Streamlit Page Config
# ==============================
st.set_page_config(page_title="Automated WCR Generator", layout="wide")
st.title("📝 Automated WCR Generator (Word Only)")

# ==============================
# Helpers
# ==============================
def _safe(x):
    """Convert NaN/datetime/None into a clean string."""
    if pd.isna(x):
        return ""
    if isinstance(x, (datetime, pd.Timestamp)):
        return x.strftime("%d-%m-%Y")  # format only date
    return str(x).strip()

# ==============================
# File Upload (Excel only)
# ==============================
uploaded_excel = st.file_uploader("📂 Upload Input Excel (.xlsx)", type=["xlsx"])

# Path to template in repo (always used)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "sample.docx")

# ==============================
# Generate Word Files
# ==============================
def generate_word_zip(df: pd.DataFrame) -> io.BytesIO:
    if not os.path.exists(TEMPLATE_PATH):
        st.error("❌ Word template file (sample.docx) not found in repo.")
        st.stop()

    with open(TEMPLATE_PATH, "rb") as f:
        template_bytes = f.read()

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zipf:
        for i, row in df.iterrows():
            context = {
                "wo_no":          _safe(row.get("wo_no", "")),
                "wo_date":        _safe(row.get("wo_date", "")),
                "wo_des":         _safe(row.get("wo_des", "")),
                "Location_code":  _safe(row.get("Location_code", "")),
                "customername_code": _safe(row.get("customername_code", "")),
                "Capacity_code":  _safe(row.get("Capacity_code", "")),
                "PB_qty":         _safe(row.get("PB_qty", row.get("Previous Bill Qty", ""))),
                "TB_Qty":         _safe(row.get("TB_Qty", row.get("THIS BILL QTY ( Final Bill Qty )", ""))),
                "cu_qty":         _safe(row.get("cu_qty", row.get("CUMULATIVE QTY", ""))),
                "B_qty":          _safe(row.get("B_qty", row.get("BALANCE QTY", ""))),
                "site_incharge":  _safe(row.get("site_incharge", "")),
                "Scada_incharge": _safe(row.get("Scada_incharge", "")),
                "Re_date":        _safe(row.get("Re_date", "")),
            }

            # Render using repo template
            tmp_doc = io.BytesIO()
            doc = DocxTemplate(io.BytesIO(template_bytes))
            doc.render(context)
            doc.save(tmp_doc)

            wo = context["wo_no"] or f"Row{i+1}"
            file_name = f"WCR_{wo}.docx"
            zipf.writestr(file_name, tmp_doc.getvalue())

    zip_buffer.seek(0)
    return zip_buffer

# ==============================
# Main Workflow
# ==============================
if uploaded_excel:
    df = pd.read_excel(uploaded_excel)
    df.columns = df.columns.str.strip()

    st.success(f"✅ Loaded {len(df)} rows from Excel.")
    st.dataframe(df.head(), use_container_width=True)

    if st.button("⬇️ Generate Word Files"):
        zip_buffer = generate_word_zip(df)
        st.download_button(
            "📥 Download All Word Files (ZIP)",
            data=zip_buffer,
            file_name="WCR_Word_Files.zip",
            mime="application/zip",
            use_container_width=True,
        )
