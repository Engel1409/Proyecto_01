import os
import re
import io
import zipfile
import warnings
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

# ---------------------------------------------------------
# PLANTILLA FIJA PARA EL TXT
# ---------------------------------------------------------
def construir_txt(nro_poliza, polizas_grupo):
    """
    nro_poliza: número extraído del PDF (encabezado del txt)
    polizas_grupo: lista de pólizas (de pg.txt) que pertenecen a ese grupo/PDF
    """
    lineas = [
        nro_poliza,
        "MAPFRE DOLAR: 0",
        "CARGO EN CUENTA: NO",
        "ENDOSADO: NO",
    ]
    lineas.extend(polizas_grupo)
    return "\n".join(lineas)


def parsear_pg_txt(contenido):
    """
    Formato esperado por línea: grupo,poliza
    Devuelve un dict: {grupo: [poliza1, poliza2, ...]}
    """
    mapa = {}
    for linea in contenido.splitlines():
        linea = linea.strip()
        if not linea or "," not in linea:
            continue
        grupo, poliza = linea.split(",", 1)
        grupo = grupo.strip()
        poliza = poliza.strip()
        if not grupo or not poliza:
            continue
        mapa.setdefault(grupo, []).append(poliza)
    return mapa


# ---------------------------------------------------------
# UPLOADERS
# ---------------------------------------------------------
uploaded_files = st.file_uploader(
    "Sube tus archivos PDF aquí", type="pdf", accept_multiple_files=True, key="pdf_uploader"
)

pg_file = st.file_uploader(
    "Sube pg.txt (formato grupo,poliza) — opcional, solo si quieres generar carpetas+TXT",
    type="txt",
    key="pg_uploader",
)

if uploaded_files:
    all_rows = []
    carpetas = {}  # nombre_pdf (sin extensión) -> {"txt": contenido, "pdf_bytes": bytes}

    mapa_pg = {}
    if pg_file:
        contenido_pg = pg_file.read().decode("utf-8", errors="ignore")
        mapa_pg = parsear_pg_txt(contenido_pg)

    for uploaded_file in uploaded_files:
        pdf_bytes = uploaded_file.read()
        nombre_pdf = os.path.splitext(uploaded_file.name)[0]

        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            text = "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())

        poliza = re.search(r"(?:P\s*[ÓO]?\s*L\s*I\s*Z\s*A|P[ÓO]LIZA)\s*[:\-]?\s*(\d{4,})", text, re.IGNORECASE)
        cliente = re.search(r"Cliente\s+([A-Z ,]+)", text)
        vigencia = re.search(r"Vigencia\s+(\d{2}/\d{2}/\d{4} - \d{2}/\d{2}/\d{4})", text)

        nro_poliza = poliza.group(1) if poliza else "SIN_POLIZA"
        nombre_cliente = cliente.group(1).strip() if cliente else "SIN_CLIENTE"
        rango_vigencia = vigencia.group(1) if vigencia else "SIN_VIGENCIA"

        seccion_pattern = re.compile(r"(SECCION: \d{3} [A-ZÑÁÉÍÓÚ ]+)")
        seccion_indices = [(m.start(), m.group()) for m in seccion_pattern.finditer(text)]
        seccion_indices.append((len(text), None))

        for i in range(len(seccion_indices) - 1):
            sec = seccion_indices[i][1]
            content = text[seccion_indices[i][0]:seccion_indices[i + 1][0]]
            lineas = content.split("\n")
            for idx, line in enumerate(lineas):
                match = re.match(r"^(.*?)(\d{1,3}(?:,\d{3})*\.\d{2})\s+(\d{1,3}(?:,\d{3})*\.\d{2})$", line.strip())
                if match:
                    item_texto = match.group(1).strip()

                    placa_match = re.search(r"PLACA:\s*([A-Z0-9\-]+)", item_texto, re.IGNORECASE)
                    placa = placa_match.group(1).strip() if placa_match else ""

                    marca = modelo = anio = ""
                    if placa and idx + 1 < len(lineas):
                        siguiente = lineas[idx + 1]
                        marca_match = re.search(r"MARCA:\s*([^,]+)", siguiente, re.IGNORECASE)
                        modelo_match = re.search(r"MODELO:\s*([^,]+)", siguiente, re.IGNORECASE)
                        anio_match = re.search(r"A[ÑN]O:\s*(\d{4})", siguiente, re.IGNORECASE)
                        marca = marca_match.group(1).strip() if marca_match else ""
                        modelo = modelo_match.group(1).strip() if modelo_match else ""
                        anio = anio_match.group(1).strip() if anio_match else ""

                    all_rows.append([nro_poliza, nombre_cliente, rango_vigencia, sec, item_texto, placa, marca, modelo, anio, match.group(2), match.group(3)])

        # Generar el TXT de este PDF (aunque no haya match en pg.txt, se crea con el encabezado base)
        polizas_grupo = mapa_pg.get(nro_poliza, [])
        carpetas[nombre_pdf] = {
            "txt": construir_txt(nro_poliza, polizas_grupo),
            "pdf_bytes": pdf_bytes,
            "pdf_filename": uploaded_file.name,
            "nro_poliza": nro_poliza,
        }

    df = pd.DataFrame(all_rows, columns=["Póliza", "Cliente", "Vigencia", "Sección", "Ítem", "Placa", "Marca", "Modelo", "Año", "Valor Asegurado", "Prima Neta"])
    st.success("✅ Archivos procesados correctamente")
    st.dataframe(df, use_container_width=True)

    if pg_file:
        sin_match = [nombre for nombre, info in carpetas.items() if info["nro_poliza"] not in mapa_pg]
        if sin_match:
            st.warning(f"⚠️ No se encontraron pólizas en pg.txt para: {', '.join(sin_match)}")

    # Excel de extracción (va dentro del ZIP final, en la raíz)
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)

    # ZIP: una carpeta por PDF con su PDF y su TXT (sin el Excel adentro)
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for nombre_pdf, info in carpetas.items():
            zf.writestr(f"{nombre_pdf}/{info['pdf_filename']}", info["pdf_bytes"])
            zf.writestr(f"{nombre_pdf}/{info['nro_poliza']}.txt", info["txt"])

    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            "⬇️ Descargar Excel",
            data=excel_buffer.getvalue(),
            file_name="Renovaciones.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="polidata_download",
        )
    with col2:
        st.download_button(
            "📁 Descargar carpetas (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="Carpetas_POLIDATA.zip",
            mime="application/zip",
            key="resultado_download",
        )
