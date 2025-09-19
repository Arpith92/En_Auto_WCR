import os
import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime
import streamlit as st
import zipfile
import io
import pypandoc

# ---- Paths ----
TEMPLATE_DOC = "sample.docx"   # keep this file in your repo
OUT_DIR = "Result"
os.makedirs(OUT_DIR, exist_ok=True)

def _safe(x):
    """Convert values into clean strings with 0.00 format for numbers."""
    if pd.isna(x) or x == "":
        return ""
    if isinstance(x, (datetime, pd.Timestamp)):
        return x.strftime("%d-%m-%Y")
    try:
        num = float(x)
        return f"{num:.2f}"
    except (ValueError, TypeError):
        return str(x).strip()

# ---- Streamlit UI ----
st.title("📑 Automated WCR Generator")

uploaded_file = st.file_uploader("Upload Input Excel File", type=["xlsx"])

if uploaded_file is not None:
    # ---- Load Excel ----
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip()

    # Rename headers for consistency
    rename_map = {
        "wo no": "wo_no", "wo_no": "wo_no",
        "wo date": "wo_date", "wo_date": "wo_date",
        "wo des": "wo_des", "wo_des": "wo_des",
        "location_code": "Location_code", "Location_code": "Location_code",
        "customername_code": "customername_code",
        "capacity_code": "Capacity_code", "Capacity_code": "Capacity_code",
        "site_incharge": "site_incharge",
        "Scada_incharge": "Scada_incharge",
        "Re_date": "Re_date",
        "Site_Name": "Site_Name",
        "Line_1_Workstatus":"Line_1_Workstatus",
        "Line_2_Workstatus":"Line_2_Workstatus",
        "Payment Terms": "Payment_Terms"
    }
    df = df.rename(columns={c: rename_map.get(c, c) for c in df.columns})

    generated_word = []
    generated_pdf = []

    for i, row in df.iterrows():
        context = {col: _safe(row[col]) for col in df.columns}

        # --- Auto-generate Sr. No. ---
        for n in [1, 2, 3]:
            fields = [
                context.get(f"Line_{n}", ""),
                context.get(f"Line_{n}_WO_qty", ""),
                context.get(f"Line_{n}_UOM", ""),
                context.get(f"Line_{n}_PB_qty", ""),
                context.get(f"Line_{n}_TB_Qty", ""),
                context.get(f"Line_{n}_cu_qty", ""),
                context.get(f"Line_{n}_B_qty", "")
            ]
            if any(f for f in fields):
                context[f"item_sr_no_{n}"] = str(n)
            else:
                context[f"item_sr_no_{n}"] = ""

        # Render Word with docxtpl
        doc = DocxTemplate(TEMPLATE_DOC)
        doc.render(context)

        # Save DOCX
        wo = context.get("wo_no", "") or f"Row{i+1}"
        word_path = os.path.join(OUT_DIR, f"WCR_{wo}.docx")
        doc.save(word_path)
        generated_word.append(word_path)

        # Convert DOCX → PDF using pypandoc
        pdf_path = os.path.join(OUT_DIR, f"WCR_{wo}.pdf")
        try:
            pypandoc.convert_file(word_path, "pdf", outputfile=pdf_path, extra_args=['--standalone'])
            generated_pdf.append(pdf_path)
        except Exception as e:
            st.error(f"PDF conversion failed for {word_path}: {e}")

    st.success(f"✅ Generated {len(generated_word)} Word files and {len(generated_pdf)} PDF files")

    # ---- ZIP Word ----
    zip_word = io.BytesIO()
    with zipfile.ZipFile(zip_word, "w") as zipf:
        for file in generated_word:
            zipf.write(file, arcname=os.path.basename(file))
    zip_word.seek(0)

    st.download_button(
        label="⬇️ Download All WCR Files (Word ZIP)",
        data=zip_word,
        file_name="WCR_Word_Files.zip",
        mime="application/zip"
    )

    # ---- ZIP PDF ----
    if generated_pdf:
        zip_pdf = io.BytesIO()
        with zipfile.ZipFile(zip_pdf, "w") as zipf:
            for file in generated_pdf:
                zipf.write(file, arcname=os.path.basename(file))
        zip_pdf.seek(0)

        st.download_button(
            label="⬇️ Download All WCR Files (PDF ZIP)",
            data=zip_pdf,
            file_name="WCR_PDF_Files.zip",
            mime="application/zip"
        )
