import os
import io
import pandas as pd
from datetime import datetime
from docxtpl import DocxTemplate
from zipfile import ZipFile
import streamlit as st

# Try PDF conversion backends
try:
    from docx2pdf import convert as win_convert
    HAS_DOCX2PDF = True
except ImportError:
    HAS_DOCX2PDF = False

try:
    import pypandoc
    HAS_PYPANDOC = True
except ImportError:
    HAS_PYPANDOC = False

st.set_page_config(page_title="Automated WCR Generator", page_icon="📑", layout="centered")
st.title("📑 Automated WCR Generator")

uploaded_excel = st.file_uploader("Upload Input Excel", type=["xlsx"])

# ---- Helpers ----
def _safe(x):
    if pd.isna(x):
        return ""
    if isinstance(x, (datetime, pd.Timestamp)):
        return x.strftime("%d-%m-%Y")
    return str(x).strip()

def normalize_headers(df):
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

def generate_files(df, as_pdf=False):
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

            template_doc = "sample.docx"
            if not os.path.exists(template_doc):
                st.error("❌ Template file 'sample.docx' not found in repo root!")
                st.stop()

            doc = DocxTemplate(template_doc)
            doc.render(context)

            wo = context["wo_no"] or f"Row{i+1}"
            word_filename = f"WCR_{wo}.docx"

            # Save Word
            temp_word = io.BytesIO()
            doc.save(temp_word)
            temp_word.seek(0)
            zf.writestr(word_filename, temp_word.read())

            # Save PDF if requested
            if as_pdf:
                pdf_filename = f"WCR_{wo}.pdf"
                if HAS_DOCX2PDF and os.name == "nt":  # Windows
                    try:
                        import tempfile
                        with tempfile.TemporaryDirectory() as tmpdir:
                            word_path = os.path.join(tmpdir, word_filename)
                            pdf_path = word_path.replace(".docx", ".pdf")
                            doc.save(word_path)
                            win_convert(word_path, pdf_path)
                            with open(pdf_path, "rb") as fpdf:
                                zf.writestr(pdf_filename, fpdf.read())
                    except Exception as e:
                        st.warning(f"⚠️ Windows PDF conversion failed for {wo}: {e}")
                elif HAS_PYPANDOC:  # Linux/Cloud
                    try:
                        with open(word_filename, "wb") as f:
                            f.write(temp_word.getvalue())
                        pdf_bytes = pypandoc.convert_file(
                            word_filename, "pdf", outputfile=None,
                            extra_args=["--pdf-engine=xelatex"]
                        )
                        zf.writestr(pdf_filename, pdf_bytes)
                        os.remove(word_filename)
                    except Exception as e:
                        st.warning(f"⚠️ Linux PDF conversion failed for {wo}: {e}")
                else:
                    st.warning("⚠️ No PDF conversion backend available.")

    memory_zip.seek(0)
    return memory_zip

# ---- UI ----
if uploaded_excel:
    df = pd.read_excel(uploaded_excel)
    df.columns = df.columns.str.strip()
    df = normalize_headers(df)

    col1, col2 = st.columns(2)

    with col1:
        if st.button("⬇️ Generate Word (ZIP)"):
            zip_buffer = generate_files(df, as_pdf=False)
            st.download_button(
                "Download Word Files",
                data=zip_buffer,
                file_name="WCR_Word_Files.zip",
                mime="application/zip"
            )

    with col2:
        if st.button("⬇️ Generate PDF (ZIP)"):
            zip_buffer = generate_files(df, as_pdf=True)
            st.download_button(
                "Download PDF Files",
                data=zip_buffer,
                file_name="WCR_PDF_Files.zip",
                mime="application/zip"
            )
