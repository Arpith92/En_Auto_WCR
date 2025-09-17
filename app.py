import os
import pandas as pd
from io import BytesIO
from datetime import datetime
from zipfile import ZipFile
import streamlit as st
from docxtpl import DocxTemplate
from fpdf import FPDF

# ======================
# Utility Functions
# ======================

def _safe(x):
    """Convert NaN/datetime/None into clean string."""
    if pd.isna(x):
        return ""
    if isinstance(x, (datetime, pd.Timestamp)):
        return x.strftime("%d-%m-%Y")
    return str(x).strip()

def generate_files(df, template_path, as_pdf=False):
    """Generate Word or PDF files from Excel rows and return as a zip BytesIO."""
    zip_buffer = BytesIO()
    with ZipFile(zip_buffer, "w") as zipf:
        for i, row in df.iterrows():
            # Build context for template
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

            wo = context["wo_no"] or f"Row{i+1}"
            filename = f"WCR_{wo}.{'pdf' if as_pdf else 'docx'}"

            if as_pdf:
                # --- PDF Generation with fpdf2 ---
                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("Arial", size=12)
                pdf.cell(0, 10, "Work Completion Report", ln=1, align="C")
                pdf.ln(10)

                for k, v in context.items():
                    pdf.cell(0, 10, f"{k}: {v}", ln=1)

                pdf_out = pdf.output(dest="S")
                if isinstance(pdf_out, str):   # old fpdf
                    pdf_bytes = pdf_out.encode("latin1")
                else:                          # fpdf2 returns bytes
                    pdf_bytes = pdf_out

                zipf.writestr(filename, pdf_bytes)

            else:
                # --- Word File Generation ---
                doc = DocxTemplate(template_path)
                doc.render(context)
                tmp_buffer = BytesIO()
                doc.save(tmp_buffer)
                zipf.writestr(filename, tmp_buffer.getvalue())

    zip_buffer.seek(0)
    return zip_buffer

# ======================
# Streamlit App
# ======================

st.title("⚡ Automated Work Completion Certificates")

st.write("Upload your Excel input file and generate **Word** or **PDF** WCRs in bulk.")

uploaded_excel = st.file_uploader("Upload Input Excel (.xlsx)", type=["xlsx"])
template_file = st.file_uploader("Upload Word Template (.docx)", type=["docx"])

if uploaded_excel and template_file:
    df = pd.read_excel(uploaded_excel)
    df.columns = df.columns.str.strip()

    # Normalize headers
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

    # Word Button
    if st.button("📄 Generate Word Files"):
        with st.spinner("Generating Word WCRs..."):
            zip_buffer = generate_files(df, template_file, as_pdf=False)
        st.success("✅ Word files generated!")
        st.download_button(
            label="⬇️ Download Word WCRs (ZIP)",
            data=zip_buffer,
            file_name="WCR_Word_Files.zip",
            mime="application/zip",
        )

    # PDF Button
    if st.button("📑 Generate PDF Files"):
        with st.spinner("Generating PDF WCRs..."):
            zip_buffer = generate_files(df, template_file, as_pdf=True)
        st.success("✅ PDF files generated!")
        st.download_button(
            label="⬇️ Download PDF WCRs (ZIP)",
            data=zip_buffer,
            file_name="WCR_PDF_Files.zip",
            mime="application/zip",
        )
