# app.py – Automated WCR Generator (Word + PDF export)

from __future__ import annotations
import os, io, zipfile, base64
from datetime import datetime
import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate
from fpdf import FPDF  # fpdf2 for PDF generation

# ==============================
# Streamlit Page Config
# ==============================
st.set_page_config(page_title="Automated WCR Generator", layout="wide")
st.title("📝 Automated WCR Generator")

# ==============================
# File Uploads
# ==============================
uploaded_excel = st.file_uploader("📂 Upload Input Excel (.xlsx)", type=["xlsx"])
uploaded_docx  = st.file_uploader("📂 Upload Word Template (.docx)", type=["docx"])

# ==============================
# PDF Helper (Safe Export)
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

    # Return as bytes (robust for Streamlit Cloud/Linux)
    out = pdf.output(dest="S")
    if isinstance(out, (bytes, bytearray)):
        return out
    return str(out).encode("latin-1", errors="ignore")

# ==============================
# Core File Generator
# ==============================
def generate_files(df: pd.DataFrame, template_bytes: bytes, as_pdf: bool = False):
    template_path = "template.docx"
    with open(template_path, "wb") as f:
        f.write(template_bytes)

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zipf:
        for _, row in df.iterrows():
            context = row.to_dict()

            # Format WO/PO dates (remove time)
            if "po_date" in context and pd.notna(context["po_date"]):
                context["po_date"] = pd.to_datetime(context["po_date"]).strftime("%Y-%m-%d")
            if "wo_date" in context and pd.notna(context["wo_date"]):
                context["wo_date"] = pd.to_datetime(context["wo_date"]).strftime("%Y-%m-%d")

            file_base = context.get("WO_No", f"WCR_{_+1}")
            file_name = f"{file_base}.{'pdf' if as_pdf else 'docx'}"

            if as_pdf:
                pdf_bytes = pdf_from_context(context)
                zipf.writestr(file_name, pdf_bytes)
            else:
                doc = DocxTemplate(template_path)
                doc.render(context)
                tmp = io.BytesIO()
                doc.save(tmp)
                zipf.writestr(file_name, tmp.getvalue())

    zip_buffer.seek(0)
    return zip_buffer

# ==============================
# Main Workflow
# ==============================
if uploaded_excel and uploaded_docx:
    df = pd.read_excel(uploaded_excel)

    st.success(f"✅ Loaded {len(df)} rows from Excel.")
    st.dataframe(df.head(), use_container_width=True)

    col1, col2 = st.columns(2)
    with col1:
        if st.button("⬇️ Generate Word Files"):
            zip_buffer = generate_files(df, uploaded_docx.read(), as_pdf=False)
            st.download_button(
                "📥 Download All Word Files (ZIP)",
                data=zip_buffer,
                file_name="WCR_Word_Files.zip",
                mime="application/zip",
                use_container_width=True,
            )

    with col2:
        if st.button("⬇️ Generate PDF Files"):
            zip_buffer = generate_files(df, uploaded_docx.read(), as_pdf=True)
            st.download_button(
                "📥 Download All PDF Files (ZIP)",
                data=zip_buffer,
                file_name="WCR_PDF_Files.zip",
                mime="application/zip",
                use_container_width=True,
            )
