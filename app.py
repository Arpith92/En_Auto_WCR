import os
import io
import pandas as pd
from datetime import datetime
from docxtpl import DocxTemplate
from zipfile import ZipFile
import streamlit as st

try:
    from docx2pdf import convert
except ImportError:
    convert = None

# ---- Streamlit UI ----
st.title("📑 Automated WCR Generator")

uploaded_excel = st.file_uploader("Upload Input Excel", type=["xlsx"])
generate_pdf = st.checkbox("Also generate PDFs", value=True)

if uploaded_excel:
    # Load Excel
    df = pd.read_excel(uploaded_excel)
    df.columns = df.columns.str.strip()

    # Normalize headers
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
    df = df.rename(columns={c: rename_map.get(c, c) for c in df.columns})

    def _safe(x):
        if pd.isna(x):
            return ""
        if isinstance(x, (datetime, pd.Timestamp)):
            return x.strftime("%d-%m-%Y")
        return str(x).strip()

    # Prepare in-memory ZIP
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

            # Load template.docx from repo
            template_doc = "sample.docx"   # keep this in your repo
            doc = DocxTemplate(template_doc)
            doc.render(context)

            wo = context["wo_no"] or f"Row{i+1}"
            word_filename = f"WCR_{wo}.docx"

            # Save Word into memory
            temp_word = io.BytesIO()
            doc.save(temp_word)
            temp_word.seek(0)
            zf.writestr(word_filename, temp_word.read())

            # Optionally generate PDF
            if generate_pdf:
                try:
                    from docx2pdf import convert
                    import tempfile

                    with tempfile.TemporaryDirectory() as tmpdir:
                        word_path = os.path.join(tmpdir, word_filename)
                        pdf_path = word_path.replace(".docx", ".pdf")
                        doc.save(word_path)
                        convert(word_path, pdf_path)
                        with open(pdf_path, "rb") as fpdf:
                            zf.writestr(f"WCR_{wo}.pdf", fpdf.read())
                except Exception as e:
                    st.warning(f"PDF conversion failed for {wo}: {e}")

    memory_zip.seek(0)
    st.success("✅ All WCR files generated!")

    st.download_button(
        "⬇️ Download All (ZIP)",
        data=memory_zip,
        file_name="WCR_Files.zip",
        mime="application/zip"
    )
