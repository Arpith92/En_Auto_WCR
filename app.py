import streamlit as st
import pandas as pd
import io, zipfile, os
from docxtpl import DocxTemplate

# Optional PDF conversion
try:
    from docx2pdf import convert  # Only works on Windows with MS Word installed
    HAS_DOCX2PDF = True
except Exception:
    HAS_DOCX2PDF = False

# =========================
# Streamlit Page Config
# =========================
st.set_page_config(page_title="Automated WCR Generator", layout="centered")
st.title("📝 Automated WCR Generator")

st.markdown("Upload your Excel file to auto-generate Word/PDF files based on template.")

# =========================
# File Uploads
# =========================
uploaded_excel = st.file_uploader("📂 Upload Input Excel (.xlsx)", type=["xlsx"])
generate_pdf = st.checkbox("Also generate PDFs (only works on Windows)", value=False)

TEMPLATE_PATH = "sample.docx"  # keep your template in repo

# =========================
# File Generation Function
# =========================
def generate_files(df: pd.DataFrame, as_pdf: bool = False):
    if not os.path.exists(TEMPLATE_PATH):
        st.error("❌ Word template file not found in repo. Add sample.docx")
        st.stop()

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zipf:
        for i, row in df.iterrows():
            context = row.to_dict()
            # Format dates
            for fld in ["po_date", "wo_date", "Re_date"]:
                if fld in context and pd.notna(context[fld]):
                    context[fld] = pd.to_datetime(context[fld]).strftime("%Y-%m-%d")

            file_base = context.get("WO_No", f"WCR_{i+1}")

            # Render Word file
            doc = DocxTemplate(TEMPLATE_PATH)
            doc.render(context)

            tmp_docx = f"{file_base}.docx"
            doc.save(tmp_docx)

            # Add Word file to ZIP
            with open(tmp_docx, "rb") as f:
                zipf.writestr(tmp_docx, f.read())

            # Convert to PDF (Windows only)
            if as_pdf and HAS_DOCX2PDF:
                try:
                    tmp_pdf = f"{file_base}.pdf"
                    convert(tmp_docx, tmp_pdf)
                    with open(tmp_pdf, "rb") as f:
                        zipf.writestr(tmp_pdf, f.read())
                    os.remove(tmp_pdf)
                except Exception as e:
                    st.warning(f"⚠️ PDF conversion failed for {file_base}: {e}")

            # Clean up local docx
            os.remove(tmp_docx)

    zip_buffer.seek(0)
    return zip_buffer

# =========================
# Main Logic
# =========================
if uploaded_excel is not None:
    try:
        df = pd.read_excel(uploaded_excel)
        st.success(f"✅ Loaded {len(df)} records from Excel.")

        if st.button("🚀 Generate Files"):
            with st.spinner("Generating files..."):
                zip_buffer = generate_files(df, as_pdf=generate_pdf)
            st.success("🎉 Files generated successfully!")

            st.download_button(
                label="⬇️ Download ZIP",
                data=zip_buffer,
                file_name="WCR_Files.zip",
                mime="application/zip",
                use_container_width=True
            )
    except Exception as e:
        st.error(f"❌ Failed to process Excel: {e}")
