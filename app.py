import os
import io
import pandas as pd
from datetime import datetime
from docxtpl import DocxTemplate
from zipfile import ZipFile
import streamlit as st

# Try PDF conversion (only works on Windows with MS Word)
try:
    from docx2pdf import convert
    HAS_DOCX2PDF = True
except ImportError:
    HAS_DOCX2PDF = False

st.set_page_config(page_title="Automated WCR Generator", page_icon="📑", layout="centered")
st.title("📑 Automated WCR Generator")

# ---- Upload Excel ----
uploaded_excel = st.file_uploader("Upload Input Excel", type=["xlsx"])
generate_pdf = st.checkbox("Also generate PDFs (⚠️ Works only on Windows with MS Word)", value=False)

# ---- Helper Functions ----
def _safe(x):
    """Convert NaN/datetime/None into a clean string."""
    if pd.isna(x):
        return ""
    if isinstance(x, (datetime, pd.Timestamp)):
        return x.strftime("%d-%m-%Y")
    return str(x).strip()

def normalize_headers(df):
    rename_map = {
        "wo no": "wo_no",
        "wo_no": "wo_no",
        "wo date": "wo_date",
        "wo_date": "wo_date",
        "wo des": "wo_des",
        "wo_des": "wo_des",
        "location_code": "Location_code",
        "Location_code": "Location_code",
        "customername_code": "customername_code",
        "capacity_code": "Capacity_code",
        "Capacity_code": "Capacity_code",
        "Previous Bill Qty": "PB_qty",
        "THIS BILL QTY ( Final Bill Qty )": "TB_Qty",
        "CUMULATIVE QTY": "cu_qty",
        "BALANCE QTY": "B_qty",
        "site_incharge": "site_incharge",
        "Scada_incharge": "Scada_incharge",
        "Re_date": "Re_date",
    }
    return df.rename(columns={c: rename_map.get(c, c) for c in df.columns})

# ---- Process Uploaded File ----
if uploaded_excel:
    df = pd.read_excel(uploaded_excel)
    df.columns = df.columns.str.strip()
    df = normalize_headers(df)

    memory_zip = io.BytesIO()
    with ZipFile(memory_zip, "w") as zf:
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

            # Load template from repo root
            template_doc = "sample.docx"
            if not os.path.exists(template_doc):
                st.error("❌ Template file 'sample.docx' not found in app directory!")
                st.stop()

            doc = DocxTemplate(template_doc)
            doc.render(context)

            wo = context["wo_no"] or f"Row{i+1}"
            word_filename = f"WCR_{wo}.docx"

            # Save Word into memory
            temp_word = io.BytesIO()
            doc.save(temp_word)
            temp_word.seek(0)
            zf.writestr(word_filename, temp_word.read())

            # PDF (only if Windows + Word)
            if generate_pdf:
                if HAS_DOCX2PDF and os.name == "nt":
                    try:
                        import tempfile
                        with tempfile.TemporaryDirectory() as tmpdir:
                            word_path = os.path.join(tmpdir, word_filename)
                            pdf_path = word_path.replace(".docx", ".pdf")
                            doc.save(word_path)
                            convert(word_path, pdf_path)
                            with open(pdf_path, "rb") as fpdf:
                                zf.writestr(f"WCR_{wo}.pdf", fpdf.read())
                    except Exception as e:
                        st.warning(f"⚠️ PDF conversion failed for {wo}: {e}")
                else:
                    st.warning("⚠️ PDF conversion not available on this server (Linux). Download Word files instead.")

    memory_zip.seek(0)
    st.success("✅ All WCR files generated successfully!")

    st.download_button(
        "⬇️ Download All (ZIP)",
        data=memory_zip,
        file_name="WCR_Files.zip",
        mime="application/zip"
    )
