# app.py – Automated WCR Generator (Excel → Word/PDF)

from __future__ import annotations
import os, io, zipfile
from datetime import datetime
import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate
from fpdf import FPDF  # fpdf2

# ==============================
# Streamlit Page Config
# ==============================
st.set_page_config(page_title="Automated WCR Generator", layout="wide")
st.title("📝 Automated WCR Generator")

# ==============================
# File Upload (Excel only)
# ==============================
uploaded_excel = st.file_uploader("📂 Upload Input Excel (.xlsx)", type=["xlsx"])

# Path to template inside repo
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "sample.docx")

# ==============================
# PDF Helper (fallback simple)
# ==============================
def pdf_from_context(context: dict) -> bytes:
    pdf = FPDF(format="A4")
    pdf.add_page()
    pdf.set_font("Helvetica", size=12)

    pdf.cell(0, 10, "Work Completion Report", ln=True, align="C")
    pdf.ln(5)

    for k, v in context.items():
        pdf.set_font("Helvetica", "B", 11)
        pdf.cell(60, 8, f"{k}:", border=0)
        pdf.set_font("Helvetica", "", 11)
        pdf.multi_cell(0, 8, str(v))
        pdf.ln(1)

    out = pdf.output(dest="S")
    return out if isinstance(out, (bytes, bytearray)) else str(out).encode("latin-1", errors="ignore")

# ==============================
# Core File Generator
# ==============================
def generate_files(df: pd.DataFrame, as_pdf: bool = False):
    if not os.path.exists(TEMPLATE_PATH):
        st.error("❌ Word template file (sample.docx) not found in repo.")
        st.stop()

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zipf:
        for i, row in df.iterrows():
            context = row.to_dict()

            # Format dates if present
            for fld in ["po_date", "wo_date", "Re_date"]:
                if fld in context and pd.notna(context[fld]):
                    context[fld] = pd.to_datetime(context[fld]).strftime("%Y-%m-%d")

            file_base = context.get("wo_no", f"WCR_{i+1}")
            file_name = f"{file_base}.{'pdf' if as_pdf else 'docx'}"

            if as_pdf:
                # Fallback: create a simplified PDF
                pdf_bytes = pdf_from_context(context)
                zipf.writestr(file_name, pdf_bytes)
            else:
                # Full Word file from template
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

    col1, col2 = st.columns(2)
    with col1:
        if st.button("⬇️ Generate Word Files"):
            zip_buffer = generate_files(df, as_pdf=False)
            st.download_button(
                "📥 Download All Word Files (ZIP)",
                data=zip_buffer,
                file_name="WCR_Word_Files.zip",
                mime="application/zip",
                use_container_width=True,
            )
    with col2:
        if st.button("⬇️ Generate PDF Files"):
            zip_buffer = generate_files(df, as_pdf=True)
            st.download_button(
                "📥 Download All PDF Files (ZIP)",
                data=zip_buffer,
                file_name="WCR_PDF_Files.zip",
                mime="application/zip",
                use_container_width=True,
            )
