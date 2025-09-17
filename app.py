import os
import pandas as pd
from io import BytesIO
from datetime import datetime
from zipfile import ZipFile
import streamlit as st
from docxtpl import DocxTemplate
from fpdf import FPDF

# --------------------------
# Helpers
# --------------------------

def _safe(x):
    """Convert NaN/datetime/None into clean string."""
    if pd.isna(x):
        return ""
    if isinstance(x, (datetime, pd.Timestamp)):
        return x.strftime("%d-%m-%Y")
    return str(x).strip()

def normalize_headers(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize possible Excel headers to the template token names."""
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
    return df.rename(columns={c: rename_map.get(c, c) for c in df.columns})

def pdf_bytes_from_context(context: dict) -> bytes:
    """
    Build a simple PDF using fpdf2 that lists the context key/value pairs.
    IMPORTANT: fpdf2's output(dest="S") returns BYTES. Do NOT .encode().
    """
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    pdf.cell(0, 10, "Work Completion Report", ln=1, align="C")
    pdf.ln(4)

    for k, v in context.items():
        # wrap long lines if needed
        line = f"{k}: {v}"
        pdf.multi_cell(0, 8, line)

    out = pdf.output(dest="S")
    # fpdf2 returns bytes; older fpdf might return str
    if isinstance(out, bytes):
        return out
    return out.encode("latin1")

def generate_files(df: pd.DataFrame, template_bytes: bytes, as_pdf: bool) -> BytesIO:
    """
    Generate either Word or PDF files for each row in df and return a ZIP in memory.
    - For Word: use docxtpl with the uploaded template bytes.
    - For PDF: use fpdf2 (no external converters).
    """
    zip_buffer = BytesIO()
    with ZipFile(zip_buffer, "w") as zipf:
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

            wo = context["wo_no"] or f"Row{i+1}"

            if as_pdf:
                # ---------- PDF (fpdf2) ----------
                pdf_bytes = pdf_bytes_from_context(context)
                zipf.writestr(f"WCR_{wo}.pdf", pdf_bytes)
            else:
                # ---------- Word (docxtpl) ----------
                # Rebuild a fresh BytesIO for docxtpl each time
                tmpl_io = BytesIO(template_bytes)
                doc = DocxTemplate(tmpl_io)
                doc.render(context)
                out_docx = BytesIO()
                doc.save(out_docx)
                zipf.writestr(f"WCR_{wo}.docx", out_docx.getvalue())

    zip_buffer.seek(0)
    return zip_buffer

# --------------------------
# Streamlit UI
# --------------------------

st.set_page_config(page_title="Automated WCR Generator", page_icon="📑", layout="centered")
st.title("📑 Automated WCR Generator")

uploaded_excel = st.file_uploader("Upload Input Excel (.xlsx)", type=["xlsx"])
uploaded_template = st.file_uploader("Upload Word Template (.docx)", type=["docx"])

if uploaded_excel and uploaded_template:
    df = pd.read_excel(uploaded_excel)
    df.columns = df.columns.str.strip()
    df = normalize_headers(df)

    # Read the template into bytes once; reuse for each row
    template_bytes = uploaded_template.read()

    st.success(f"✅ Loaded {len(df)} rows from Excel.")

    col1, col2 = st.columns(2)

    with col1:
        if st.button("📄 Generate Word (ZIP)"):
            with st.spinner("Generating Word files..."):
                zip_buf = generate_files(df, template_bytes, as_pdf=False)
            st.success("✅ Word files ready.")
            st.download_button(
                "⬇️ Download Word ZIP",
                data=zip_buf,
                file_name="WCR_Word_Files.zip",
                mime="application/zip",
            )

    with col2:
        if st.button("📑 Generate PDF (ZIP)"):
            with st.spinner("Generating PDF files..."):
                zip_buf = generate_files(df, template_bytes, as_pdf=True)
            st.success("✅ PDF files ready.")
            st.download_button(
                "⬇️ Download PDF ZIP",
                data=zip_buf,
                file_name="WCR_PDF_Files.zip",
                mime="application/zip",
            )
