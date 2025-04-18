import asyncio
try:
    asyncio.get_running_loop()
except RuntimeError:
    asyncio.set_event_loop(asyncio.new_event_loop())

import os
os.environ["STREAMLIT_SERVER_FILE_WATCHER_TYPE"] = "none"

import streamlit as st
from PyPDF2 import PdfReader
from docx import Document
from pdf2image import convert_from_bytes
from PIL import Image
import io
import easyocr 
import numpy as np
import re

# 🧼 Nettoyage du texte pour qu’il soit compatible XML
def clean_text(text):
    return re.sub(r'[^\x09\x0A\x0D\x20-\x7E\u00A0-\uFFFF]', '', text)

# 📘 Extraction du texte natif
def extract_native_text(pdf_file):
    reader = PdfReader(pdf_file)
    full_text = ""
    for page in reader.pages:
        text = page.extract_text()
        if text:
            full_text += text + "\n"
    return full_text.strip()

# 🧠 OCR sur les images du PDF
def extract_text_with_ocr_images(pdf_bytes):
    images = convert_from_bytes(pdf_bytes)
    reader = easyocr.Reader(['fr'], gpu=False)  
    full_text = ""
    for img in images:
        img_rgb = img.convert("RGB")
        result = reader.readtext(np.array(img_rgb), detail=0, paragraph=True)
        if result:
            full_text += "\n".join(result) + "\n"
    return full_text.strip()

# 🎬 Fonction principale de l'app Streamlit
def main():
    st.set_page_config(page_title="PDF vers Word avec OCR", layout="centered")
    st.title("📄 Convertisseur PDF -> Word avec OCR")

    uploaded_file = st.file_uploader("📎 Téléversez un fichier PDF", type="pdf")

    if uploaded_file:
        if st.button("🔍 Extraire et convertir"):
            with st.spinner("⏳ Extraction du texte natif..."):
                native_text = extract_native_text(uploaded_file)

            uploaded_file.seek(0)  # Remise à zéro du fichier pour l’OCR
            with st.spinner("🔍 Analyse OCR des images..."):
                ocr_text = extract_text_with_ocr_images(uploaded_file.read())

            combined_text = native_text + "\n\n" + ocr_text if ocr_text else native_text

            if not combined_text.strip():
                st.warning("⚠️ Aucun texte trouvé dans le document.")
            else:
                doc = Document()
                for para in combined_text.split("\n\n"):
                    clean = clean_text(para.strip().replace("\n", " "))
                    if clean:
                        doc.add_paragraph(clean)

                buffer = io.BytesIO()
                doc.save(buffer)
                buffer.seek(0)

                st.success("✅ Fichier Word prêt au téléchargement !")
                st.download_button(
                    label="⬇️ Télécharger le fichier Word",
                    data=buffer,
                    file_name="document_converti.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )


