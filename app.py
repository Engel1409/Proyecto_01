import os
import re
import io
import zipfile
import base64
import warnings
from collections import Counter
from datetime import datetime
import pdfplumber
import pandas as pd
import streamlit as st
import openpyxl

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ---------------------------------------------------------
# CONFIGURACIÓN DE PÁGINA
# ---------------------------------------------------------
st.set_page_config(page_title="POLIDATA", layout="wide")

st.markdown("""
    <style>
    #MainMenu, footer, header { visibility: hidden; }
    .stApp { background-color: #f6f8fa; font-family: 'Segoe UI', Arial, sans-serif; }
    h1 { color: #0a3d62; border-left: 5px solid #DA291C; padding-left: 10px; }
    div.stButton > button { background-color: #DA291C !important; color: white !important; font-weight: bold !important; border-radius: 8px !important; }
    </style>
""", unsafe_allow_html=True)

st.title("📄 POLIDATA")
st.caption("Extracción automática de datos desde PDF")

# --- Funciones auxiliares ---
def construir_txt(nro_poliza, polizas_grupo):
    lineas = [nro_poliza, "MAPFRE DOLAR: 0", "CARGO EN CUENTA: NO", "ENDOSADO: NO"]
    lineas.extend(polizas_grupo)
    return "\n".join(lineas)

def parsear_pg_txt(contenido):
    mapa = {}
    for linea in contenido.splitlines():
        linea = linea.strip()
        if not linea or "," not in linea: continue
        grupo, poliza = linea.split(",", 1)
        mapa.setdefault(grupo.strip(), []).append(poliza.strip())
    return mapa

def sanear_nombre(nombre):
    import unicodedata
    nombre = unicodedata.normalize("NFKD", nombre).encode("ascii", "ignore").decode("ascii")
    nombre = re.sub(r'[\\/:*?"<>|]', "_", nombre)
    return re.sub(r"\s+", "_", nombre).strip("_") or "SIN_NOMBRE"

# --- Lógica principal ---
if "reset_id" not in st.session_state: st.session_state.reset_id = 0

col_up1, col_up2, col_up3 = st.columns([2, 2, 1])
with col_up1:
    uploaded_files = st.file_uploader("Sube tus archivos PDF", type="pdf", accept_multiple_files=True, key=f"pdf_{st.session_state.reset_id}")
with col_up2:
    pg_file = st.file_uploader("Sube pg.txt (opcional)", type="txt", key=f"pg_{st.session_state.reset_id}")
with col_up3:
    if st.button("🔄 Reiniciar", use_container_width=True):
        st.session_state.reset_id += 1
        st.rerun()

if uploaded_files:
    all_rows, carpetas = [], {}
    mapa_pg = parsear_pg_txt(pg_file.read().decode("utf-8", errors="ignore")) if pg_file else {}

    for uploaded_file in uploaded_files:
        pdf_bytes = uploaded_file.read()
        nombre_pdf = os.path.splitext(uploaded_file.name)[0]
        
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            text = "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())
        
        poliza = re.search(r"(?:P\s*[ÓO]?\s*L\s*I\s*Z\s*A|P[ÓO]LIZA)\s*[:\-]?\s*(\d{4,})", text, re.IGNORECASE)
        nro_poliza = poliza.group(1) if poliza else "SIN_POLIZA"
        
        # [Aquí iría tu lógica de extracción de tablas que ya tenías]
        
        carpetas[nombre_pdf] = {
            "txt": construir_txt(nro_poliza, mapa_pg.get(nro_poliza, [])),
            "pdf_bytes": pdf_bytes,
            "pdf_filename": uploaded_file.name,
            "nro_poliza": nro_poliza
        }

    # --- CORRECCIÓN ZIP ---
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for nombre_pdf, info in carpetas.items():
            zf.writestr(f"{sanear_nombre(nombre_pdf)}/{sanear_nombre(info['pdf_filename'])}", info["pdf_bytes"])
            zf.writestr(f"{sanear_nombre(nombre_pdf)}/{sanear_nombre(info['nro_poliza'])}.txt", info["txt"])
    
    zip_data = zip_buffer.getvalue()
    b64_zip = base64.b64encode(zip_data).decode()
    
    st.divider()
    sello = datetime.now().strftime("%Y%m%d")
    st.markdown(f'''
        <a href="data:application/zip;base64,{b64_zip}" download="Carpetas_POLIDATA_{sello}.zip">
            <button style="width:100%; background:#1e293b; color:white; border:none; padding:10px; border-radius:8px; font-weight:bold;">
                📁 Descargar carpetas (ZIP)
            </button>
        </a>
    ''', unsafe_allow_html=True)
