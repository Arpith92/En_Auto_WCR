import os
import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime
import streamlit as st
import zipfile
import io

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

    generated_files = []

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

        # Render with docxtpl
        doc = DocxTemplate(TEMPLATE_DOC)
        doc.render(context)

        # Save output file
        wo = context.get("wo_no", "") or f"Row{i+1}"
        out_path = os.path.join(OUT_DIR, f"WCR_{wo}.docx")
        doc.save(out_path)
        generated_files.append(out_path)

    st.success(f"✅ Generated {len(generated_files)} Word files")

    # ---- Zip all generated files ----
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zipf:
        for file in generated_files:
            zipf.write(file, arcname=os.path.basename(file))
    zip_buffer.seek(0)

    # ---- Download button ----
    st.download_button(
        label="⬇️ Download All WCR Files (ZIP)",
        data=zip_buffer,
        file_name="WCR_Files.zip",
        mime="application/zip"
    )
