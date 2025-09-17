import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from zipfile import ZipFile
import io, os

# PDF fallback
import pypandoc

st.set_page_config(page_title="Automated WCR Generator", layout="wide")
st.title("📝 Automated WCR Generator")

# -------------------------
# File Uploads
# -------------------------
excel_file = st.file_uploader("📂 Upload Input Excel (.xlsx)", type=["xlsx"])
template_file = st.file_uploader("📂 Upload Word Template (.docx)", type=["docx"])
generate_pdf = st.checkbox("Also generate PDFs", value=False)

# -------------------------
# Generate Files
# -------------------------
def generate_files(df: pd.DataFrame, template_bytes: bytes, as_pdf: bool = False):
    zip_buffer = io.BytesIO()

    # Save uploaded template temporarily
    template_path = "uploaded_template.docx"
    with open(template_path, "wb") as f:
        f.write(template_bytes)

    with ZipFile(zip_buffer, "w") as zipf:
        for _, row in df.iterrows():
            context = row.to_dict()
            wo = str(context.get("wo_no", context.get("WO_No", "Unknown")))
            file_base = f"WCR_{wo}"

            # --- Generate Word file ---
            tmp_docx = f"{file_base}.docx"
            doc = DocxTemplate(template_path)
            doc.render(context)
            doc.save(tmp_docx)

            with open(tmp_docx, "rb") as f:
                zipf.writestr(tmp_docx, f.read())
            os.remove(tmp_docx)

            # --- Generate PDF if selected ---
            if as_pdf:
                tmp_pdf = f"{file_base}.pdf"
                try:
                    pypandoc.convert_file(template_path, "pdf", outputfile=tmp_pdf)
                    if os.path.exists(tmp_pdf):
                        with open(tmp_pdf, "rb") as f:
                            zipf.writestr(tmp_pdf, f.read())
                        os.remove(tmp_pdf)
                except Exception as e:
                    st.warning(f"⚠️ PDF conversion failed for {file_base}: {e}")

    zip_buffer.seek(0)
    return zip_buffer

# -------------------------
# Main App
# -------------------------
if excel_file and template_file:
    df = pd.read_excel(excel_file)
    st.dataframe(df.head(), use_container_width=True)

    if st.button("🚀 Generate WCR Files"):
        zip_buffer = generate_files(df, template_file.read(), as_pdf=generate_pdf)
        st.success("Files generated successfully!")

        st.download_button(
            "⬇️ Download All WCR Files (ZIP)",
            data=zip_buffer,
            file_name="WCR_Output.zip",
            mime="application/zip",
            use_container_width=True
        )
