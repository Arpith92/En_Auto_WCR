import os
import io
import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate
from fpdf import FPDF
from zipfile import ZipFile
from datetime import datetime

# --- Safe conversion function ---
def _safe(x):
    if pd.isna(x):
        return ""
    if isinstance(x, (datetime, pd.Timestamp)):
        return x.strftime("%d-%m-%Y")
    return str(x).strip()

# --- Generate files (Word or PDF) ---
def generate_files(df, as_pdf=False):
    buffer = io.BytesIO()
    with ZipFile(buffer, "w") as zipf:
        for i, row in df.iterrows():
            context = {
                "wo_no": _safe(row.get("wo_no", "")),
                "wo_date": _safe(row.get("wo_date", "")),
                "wo_des": _safe(row.get("wo_des", "")),
                "Location_code": _safe(row.get("Location_code", "")),
                "customername_code": _safe(row.get("customername_code", "")),
                "Capacity_code": _safe(row.get("Capacity_code", "")),
                "PB_qty": _safe(row.get("PB_qty", row.get("Previous Bill Qty", ""))),
                "TB_Qty": _safe(row.get("TB_Qty", row.get("THIS BILL QTY ( Final Bill Qty )", ""))),
                "cu_qty": _safe(row.get("cu_qty", row.get("CUMULATIVE QTY", ""))),
                "B_qty": _safe(row.get("B_qty", row.get("BALANCE QTY", ""))),
                "site_incharge": _safe(row.get("site_incharge", "")),
                "Scada_incharge": _safe(row.get("Scada_incharge", "")),
                "Re_date": _safe(row.get("Re_date", "")),
            }

            wo = context["wo_no"] or f"Row{i+1}"
            filename = f"WCR_{wo}.{'pdf' if as_pdf else 'docx'}"

            if as_pdf:
                # --- Generate PDF ---
                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("Arial", size=12)
                for k, v in context.items():
                    pdf.cell(0, 10, f"{k}: {v}", ln=1)
                pdf_bytes = pdf.output(dest="S")  # ✅ fpdf2 returns bytes
                zipf.writestr(filename, pdf_bytes)
            else:
                # --- Generate Word ---
                template_path = "sample.docx"  # must be in repo
                doc = DocxTemplate(template_path)
                doc.render(context)
                docx_io = io.BytesIO()
                doc.save(docx_io)
                zipf.writestr(filename, docx_io.getvalue())

    buffer.seek(0)
    return buffer

# --- Streamlit UI ---
st.title("📄 Automated WCR Generator")

uploaded_file = st.file_uploader("Upload Input Excel", type=["xlsx"])
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip()

    # Normalize Excel headers
    rename_map = {
        "wo no": "wo_no", "wo_no": "wo_no",
        "wo date": "wo_date", "wo_date": "wo_date",
        "wo des": "wo_des", "wo_des": "wo_des",
        "location_code": "Location_code", "Location_code": "Location_code",
        "customername_code": "customername_code",
        "capacity_code": "Capacity_code", "Capacity_code": "Capacity_code",
        "Previous Bill Qty": "PB_qty",
        "THIS BILL QTY ( Final Bill Qty )": "TB_Qty",
        "CUMULATIVE QTY": "cu_qty",
        "BALANCE QTY": "B_qty",
        "site_incharge": "site_incharge",
        "Scada_incharge": "Scada_incharge",
        "Re_date": "Re_date",
    }
    df = df.rename(columns={c: rename_map.get(c, c) for c in df.columns})

    st.success(f"✅ Loaded {len(df)} records from Excel.")

    col1, col2 = st.columns(2)
    with col1:
        if st.button("📥 Generate Word ZIP"):
            zip_buffer = generate_files(df, as_pdf=False)
            st.download_button(
                label="⬇️ Download Word Files (ZIP)",
                data=zip_buffer,
                file_name="WCR_Word_Files.zip",
                mime="application/zip"
            )
    with col2:
        if st.button("📥 Generate PDF ZIP"):
            zip_buffer = generate_files(df, as_pdf=True)
            st.download_button(
                label="⬇️ Download PDF Files (ZIP)",
                data=zip_buffer,
                file_name="WCR_PDF_Files.zip",
                mime="application/zip"
            )
