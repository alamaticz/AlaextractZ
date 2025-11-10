import streamlit as st
import pandas as pd
from PIL import Image
import io
import tempfile
import zipfile
import os
from extract_shipping_bills import process_shipping_bills   # <-- your existing logic

st.set_page_config(page_title="ExtractZ - Shipping Bill Extractor", layout="centered")
logo = Image.open("logo.png") 
st.image(logo, use_container_width=True)

st.markdown("""
## **ExtractZ**
Upload PDF or ZIP files to extract structured shipping data.
""")

uploaded_file = st.file_uploader("Upload File", type=["pdf", "zip"])

if uploaded_file:
    with st.spinner("Processing files... Please wait ⏳"):
        temp_dir = tempfile.mkdtemp()
        pdf_dir = os.path.join(temp_dir, "pdfs")
        os.makedirs(pdf_dir, exist_ok=True)

        file_path = os.path.join(temp_dir, uploaded_file.name)

        with open(file_path, "wb") as f:
            f.write(uploaded_file.read())

        # If ZIP → extract PDFs
        if uploaded_file.name.lower().endswith(".zip"):
            with zipfile.ZipFile(file_path, "r") as zip_ref:
                for zip_info in zip_ref.filelist:
                    if zip_info.filename.lower().endswith(".pdf"):
                        zip_info.filename = os.path.basename(zip_info.filename)
                        zip_ref.extract(zip_info, pdf_dir)
        else:
            os.rename(file_path, os.path.join(pdf_dir, uploaded_file.name))

        df = process_shipping_bills(pdf_dir)

    st.success(f"✅ Extracted {len(df)} rows successfully!")
    st.write("### Preview Data:")
    st.dataframe(df)

    # ✅ Prepare CSV download
    csv_data = df.to_csv(index=False).encode("utf-8")

    # ✅ Prepare Excel download using BytesIO
    excel_buffer = io.BytesIO()
    df.to_excel(excel_buffer, index=False, engine="openpyxl")
    excel_buffer.seek(0)

    # ✅ Download buttons
    st.download_button(
        label="📥 Download CSV",
        data=csv_data,
        file_name="shipping_bill_output.csv",
        mime="text/csv"
    )

    st.download_button(
        label="📥 Download Excel",
        data=excel_buffer,
        file_name="shipping_bill_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("📄 Please upload a PDF or ZIP file to start.")
