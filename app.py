# app.py – Automated WCR Generator (Excel → Word Only)

from __future__ import annotations
import os, io, zipfile
import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate

# ==============================
# Streamlit Page Config
# ==============================
st.set_page_config(page_title="Automated WCR Generator", layout="wide")
st.title("📝 Automated WCR Generator (Word Only)")

# ==============================
# File Upload (Excel only)
# ==============================
uploaded_excel = st.file_uploader("📂 Upload Input Excel (.xlsx)", type=["xlsx"])

# Path to template inside repo
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "sample.docx")

# ==============================
# Core File Generator (Word only)
# ==============================
def generate_word_files(df: pd.DataFrame):
    if not os.path.exists(TEMPLATE_PATH):
        st.error("❌ Word template file (sample.docx) not found in repo.")
        st.stop()

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zipf:
        for i, row in df.iterrows():
            context = row.to_dict()

            # Format date columns if present
            for fld in ["po_date", "wo_date", "Re_date"]:
                if fld in context and pd.notna(context[fld]):
                    context[fld] = pd.to_datetime(context[fld]).strftime("%Y-%m-%d")

            file_base = context.get("wo_no", f"WCR_{i+1}")
            file_name = f"{file_base}.docx"

            doc = DocxTemplate(TEMPLATE_PATH)
            doc.render(context)
            tmp = io.BytesIO()
            doc.save(tmp)
            zipf.writestr(file_name, tmp.getvalue())

    zip_buffer.seek(0)
    return zip_buffer

# ==============================
# Main Workflow
# ==============================
if uploaded_excel:
    df = pd.read_excel(uploaded_excel)
    st.success(f"✅ Loaded {len(df)} rows from Excel.")
    st.dataframe(df.head(), use_container_width=True)

    if st.button("⬇️ Generate Word Files"):
        zip_buffer = generate_word_files(df)
        st.download_button(
            "📥 Download All Word Files (ZIP)",
            data=zip_buffer,
            file_name="WCR_Word_Files.zip",
            mime="application/zip",
            use_container_width=True,
        )
