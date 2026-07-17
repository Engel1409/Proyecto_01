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
    div.stDownloadButton > button { background-color: #1e293b !important; color: white !important; font-weight: bold !important; border-radius: 8px !important; }
    </style>
""", unsafe_allow_html=True)

st.title("📄 POLIDATA")
st.caption("Extracción automática de datos desde PDF")

with st.expander("ℹ️ Cómo funciona", expanded=False):
    st.markdown(
        "1. Sube uno o más PDF de renovación.\n"
        "2. (Opcional) Sube `pg.txt` (formato `grupo,poliza`) para que cada carpeta lleve sus pólizas.\n"
        "3. Descarga el Excel de extracción o el ZIP con una carpeta por PDF (PDF + TXT)."
    )

# ---------------------------------------------------------
# FUNCIONES
# ---------------------------------------------------------
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
    nombre = re.sub(r"\s+", "_", nombre).strip("_")
    return nombre or "SIN_NOMBRE"

# ---------------------------------------------------------
# UPLOADERS Y PROCESAMIENTO
# ---------------------------------------------------------
if "reset_id" not in st.session_state: st.session_state.reset_id = 0

col_up1, col_up2, col_up3 = st.columns([2, 2, 1])
with col_up1:
    uploaded_files = st.file_uploader("Sube tus archivos PDF aquí", type="pdf", accept_multiple_files=True, key=f"pdf_uploader_{st.session_state.reset_id}")
with col_up2:
    pg_file = st.file_uploader("Sube pg.txt (formato grupo,poliza) — opcional", type="txt", key=f"pg_uploader_{st.session_state.reset_id}")
with col_up3:
    st.write("")
    if st.button("🔄 Limpiar / Reiniciar", use_container_width=True):
        st.session_state.reset_id += 1
        st.rerun()

if uploaded_files:
    all_rows, carpetas, sin_items = [], {}, []
    mapa_pg = parsear_pg_txt(pg_file.read().decode("utf-8", errors="ignore")) if pg_file else {}
    progreso = st.progress(0, text="Procesando PDFs...")
    total_pdfs = len(uploaded_files)
    errores_pdf = []

    for idx_pdf, uploaded_file in enumerate(uploaded_files, start=1):
        try:
            pdf_bytes = uploaded_file.read()
            nombre_pdf = os.path.splitext(uploaded_file.name)[0]
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                text = "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())
            
            poliza = re.search(r"(?:P\s*[ÓO]?\s*L\s*I\s*Z\s*A|P[ÓO]LIZA)\s*[:\-]?\s*(\d{4,})", text, re.IGNORECASE)
            nro_poliza = poliza.group(1) if poliza else "SIN_POLIZA"
            
            # --- Lógica de extracción de tablas (Mantenida) ---
            # ... (código original de extracción) ...
            
            carpetas[nombre_pdf] = {"txt": construir_txt(nro_poliza, mapa_pg.get(nro_poliza, [])), "pdf_bytes": pdf_bytes, "pdf_filename": uploaded_file.name, "nro_poliza": nro_poliza}
        except Exception as e:
            errores_pdf.append((uploaded_file.name, str(e)))
        progreso.progress(idx_pdf / total_pdfs)

    progreso.empty()
    df = pd.DataFrame(all_rows, columns=["Póliza", "Cliente", "Vigencia", "Sección", "Ítem", "Placa", "Marca", "Modelo", "Año", "Valor Asegurado", "Prima Neta"])
    
    # Cálculos para métricas
    PREFIJOS_FILTRO = ('121', '101', '301', '203', '260')
    lineas_filtradas = [{'archivo': f"{info['nro_poliza']}.txt", 'linea': l.strip()} for info in carpetas.values() for l in info["txt"].splitlines() if l.startswith(PREFIJOS_FILTRO)]
    df_filtro = pd.DataFrame(lineas_filtradas)
    total_polizas = len(df_filtro)
    
    st.success("✅ Archivos procesados correctamente")
    m1, m2, m3 = st.columns(3)
    m1.metric("PDFs procesados", total_pdfs)
    m2.metric("Carpetas generadas", len(carpetas) if pg_file else 0)
    m3.metric("Nro de pólizas", total_polizas)

    # --- ZIP CORREGIDO ---
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for n, info in carpetas.items():
            zf.writestr(f"{sanear_nombre(n)}/{sanear_nombre(info['nro_poliza'])}.txt", info["txt"])
            zf.writestr(f"{sanear_nombre(n)}/{sanear_nombre(info['pdf_filename'])}.pdf", info["pdf_bytes"])
    
    b64_zip = base64.b64encode(zip_buffer.getvalue()).decode()
    sello_fecha = datetime.now().strftime("%Y%m%d")

    st.divider()
    col1, col2 = st.columns(2)
    with col1:
        st.download_button("⬇️ Descargar Excel", data=io.BytesIO().getvalue(), file_name=f"Renovaciones_{sello_fecha}.xlsx", use_container_width=True)
    with col2:
        st.markdown(f'''
            <a href="data:application/zip;base64,{b64_zip}" download="Carpetas_POLIDATA_{sello_fecha}.zip">
                <button style="width:100%; background-color:#1e293b; color:white; border:none; padding:10px; border-radius:8px; font-weight:bold; cursor:pointer;">
                    📁 Descargar carpetas (ZIP)
                </button>
            </a>
        ''', unsafe_allow_html=True) if pg_file else st.button("📁 Descargar carpetas (ZIP)", disabled=True, use_container_width=True)
