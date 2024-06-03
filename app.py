import streamlit as st
from docx2pdf import convert
from pdf2docx import Converter
import tempfile
from PIL import Image
import io
import random
import pandas as pd

st.title("Document Converter")

conversion_type = st.radio("Choose conversion type:", ("DOC/DOCX to PDF", "PDF to DOC","CSV to XLSX","XLSX to CSV"))

uploaded_file = st.file_uploader(
    "Upload a file",
    type=["doc", "docx", "pdf","csv","xlsx"],
)

if uploaded_file is not None:
    with tempfile.NamedTemporaryFile(delete=False) as temp_file:
        temp_file.write(uploaded_file.read())
        file_path = temp_file.name
    try:
        if uploaded_file.type == "text/csv" and conversion_type == "CSV to XLSX":
            df = pd.read_csv(uploaded_file)
            xlsx_bytes = io.BytesIO()
            with pd.ExcelWriter(xlsx_bytes) as writer:
                df.to_excel(writer, index=False)

            st.download_button(
                label="Download XLSX",
                data=xlsx_bytes.getvalue(),
                file_name=uploaded_file.name.replace(".csv", ".xlsx"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.success("Conversion successful!")
        
        elif uploaded_file.type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" and conversion_type == "XLSX to CSV":
            df = pd.read_excel(uploaded_file)
            csv_string = df.to_csv(index=False)

            st.download_button(
                label="Download CSV",
                data=csv_string,
                file_name=uploaded_file.name.replace(".xlsx", ".csv"),
                mime="text/csv",
            )
            st.success("Conversion successful!")

        elif conversion_type in ["DOC/DOCX to PDF", "PDF to DOC"]:

            if conversion_type == "DOC/DOCX to PDF":
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
                    pdf_path = temp_pdf.name

                try:
                    convert(file_path, pdf_path)
                    st.download_button(
                        label="Download PDF",
                        data=open(pdf_path, "rb").read(),
                        file_name=uploaded_file.name.replace(".docx", ".pdf").replace(".doc", ".pdf"),
                        mime="application/pdf",
                    )
                    st.success("Conversion successful!")
                except Exception as e:
                    st.error(f"Conversion failed: {e}")

            elif conversion_type == "PDF to DOC":
                with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_doc:
                    doc_path = temp_doc.name

                try:
                    cv = Converter(file_path)
                    cv.convert(doc_path, start=0, end=None)
                    cv.close()
                    st.download_button(
                        label="Download DOC",
                        data=open(doc_path, "rb").read(),
                        file_name=uploaded_file.name.replace(".pdf", ".docx"),
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    )
                    st.success("Conversion successful!")
                except Exception as e:
                    st.error(f"Conversion failed: {e}")
    except Exception as e:
        st.error(f"Conversion failed: {e}")

#Iamage
st.title("Image Converter")
conversion_type = st.radio("Choose conversion type:", ("JPG to PNG", "PNG to JPG"))

uploaded_file = st.file_uploader("Upload an image", type=["jpg", "jpeg", "png"])

if uploaded_file is not None:
    image_bytes = uploaded_file.read()

    if conversion_type == "JPG to PNG":
        try:
            img = Image.open(io.BytesIO(image_bytes))
            png_bytes = io.BytesIO()
            img.save(png_bytes, format="PNG")

            st.download_button(
                label="Download PNG",
                data=png_bytes.getvalue(),
                file_name=uploaded_file.name.replace(".jpg", ".png").replace(".jpeg", ".png"),
                mime="image/png",
            )
            st.success("Conversion successful!")
        except Exception as e:
            st.error(f"Conversion failed: {e}")
    elif conversion_type == "PNG to JPG":
        try:
            img = Image.open(io.BytesIO(image_bytes))
            if img.mode in ('RGBA', 'LA'):
                img = img.convert('RGB')
            jpg_bytes = io.BytesIO()
            img.save(jpg_bytes, format="JPEG")

            st.download_button(
                label="Download JPG",
                data=jpg_bytes.getvalue(),
                file_name=uploaded_file.name.replace(".png", ".jpg"),
                mime="image/jpeg",
            )
            st.success("Conversion successful!")
        except Exception as e:
            st.error(f"Conversion failed: {e}")
