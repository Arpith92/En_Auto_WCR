# app.py – Automated WCR Generator (Word Only, with dynamic rows)
import os, io, zipfile
import pandas as pd
import streamlit as st
from datetime import datetime
from docxtpl import DocxTemplate

# ==============================
# Streamlit Page Config
# ==============================
st.set_page_config(page_title="Automated WCR Generator (Word Only)", layout="wide")
st.title("📝 Automated WCR Generator (Word Only)")

# ==============================
# File Upload
# ==============================
uploaded_excel = st.file_uploader("📂 Upload Input Excel (.xlsx)", type=["xlsx"])

# Path to Word template (keep in repo)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "sample.docx")

# ==============================
# Helpers
# ==============================
def _safe(x):
    """Convert NaN/datetime/None into a clean string."""
    if pd.isna(x):
        return ""
    if isinstance(x, (datetime, pd.Timestamp)):
        return x.strftime("%d-%m-%Y")
    return str(x).strip()

def generate_files(df: pd.DataFrame):
    if not os.path.exists(TEMPLATE_PATH):
        st.error("❌ Word template file (sample.docx) not found in repo.")
        st.stop()

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zipf:
        for i, row in df.iterrows():
            context = {
                "wo_no": _safe(row.get("wo_no")),
                "wo_date": _safe(row.get("wo_date")),
                "wo_des": _safe(row.get("wo_des")),
                "Location_code": _safe(row.get("Location_code")),
                "customername_code": _safe(row.get("customername_code")),
                "Capacity_code": _safe(row.get("Capacity_code")),
                "site_incharge": _safe(row.get("site_incharge")),
                "Scada_incharge": _safe(row.get("Scada_incharge")),
                "Re_date": _safe(row.get("Re_date")),
                "Site_Name": _safe(row.get("Site_Name")),
                "Payment_Terms": _safe(row.get("Payment Terms")),
            }

            # Build dynamic line items
            line_items = []
            for n in [1, 2, 3]:
                desc = _safe(row.get(f"Line_{n}", ""))
                if desc:  # only add row if description is present
                    line_items.append({
                        "sr_no": len(line_items) + 1,
                        "description": desc,
                        "WO_qty": _safe(row.get(f"Line_{n}_WO_qty")),
                        "PB_qty": _safe(row.get(f"Line_{n}_PB_qty")),
                        "TB_Qty": _safe(row.get(f"Line_{n}_TB_Qty")),
                        "cu_qty": _safe(row.get(f"Line_{n}_cu_qty")),
                        "B_qty": _safe(row.get(f"Line_{n}_B_qty")),
                    })

            context["line_items"] = line_items

            # Render Word
            doc = DocxTemplate(TEMPLATE_PATH)
            doc.render(context)

            wo = context["wo_no"] or f"Row{i+1}"
            file_name = f"WCR_{wo}.docx"
            tmp = io.BytesIO()
            doc.save(tmp)
            zipf.writestr(file_name, tmp.getvalue())

    zip_buffer.seek(0)
    return zip_buffer

# ==============================
# Main Workflow
# ==============================
if uploaded_excel:
    df = pd.read_excel(uploaded_excel)
    st.success(f"✅ Loaded {len(df)} rows from Excel.")
    st.dataframe(df.head(), use_container_width=True)

    if st.button("⬇️ Generate Word Files"):
        zip_buffer = generate_files(df)
        st.download_button(
            "📥 Download All Word Files (ZIP)",
            data=zip_buffer,
            file_name="WCR_Word_Files.zip",
            mime="application/zip",
            use_container_width=True,
        )
