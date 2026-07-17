import os
import re
import io
import zipfile
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
if "reset_id" not in st.session_state:
    st.session_state.reset_id = 0

col_up1, col_up2, col_up3 = st.columns([2, 2, 1])
with col_up1:
    uploaded_files = st.file_uploader(
        "Sube tus archivos PDF aquí", type="pdf", accept_multiple_files=True,
        key=f"pdf_uploader_{st.session_state.reset_id}"
    )
with col_up2:
    pg_file = st.file_uploader(
        "Sube pg.txt (formato grupo,poliza) — opcional",
        type="txt",
        key=f"pg_uploader_{st.session_state.reset_id}",
    )
with col_up3:
    st.write("")
    st.write("")
    if st.button("🔄 Limpiar / Reiniciar", use_container_width=True):
        st.session_state.reset_id += 1
        st.rerun()

if not uploaded_files:
    st.info("⬆️ Sube al menos un PDF para comenzar.")

if uploaded_files:
    nombres = [f.name for f in uploaded_files]
    duplicados = sorted({n for n, c in Counter(nombres).items() if c > 1})
    if duplicados:
        st.error(f"❌ Hay archivos PDF duplicados, quítalos antes de continuar: {', '.join(duplicados)}")
        st.stop()

    all_rows = []
    carpetas = {}  # nombre_pdf (sin extensión) -> {"txt": contenido, "pdf_bytes": bytes, "pdf_filename": ..., "nro_poliza": ...}
    sin_items = []  # nombres de PDF que no generaron ninguna fila en la extracción

    mapa_pg = {}
    if pg_file:
        contenido_pg = pg_file.read().decode("utf-8", errors="ignore")
        mapa_pg = parsear_pg_txt(contenido_pg)

    progreso = st.progress(0, text="Procesando PDFs...")
    total_pdfs = len(uploaded_files)
    errores_pdf = []  # [(nombre_archivo, mensaje_error)]

    for idx_pdf, uploaded_file in enumerate(uploaded_files, start=1):
        try:
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

            items_antes = len(all_rows)

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

            if len(all_rows) == items_antes:
                sin_items.append(uploaded_file.name)

            # Generar el TXT de este PDF (aunque no haya match en pg.txt, se crea con el encabezado base)
            polizas_grupo = mapa_pg.get(nro_poliza, [])
            carpetas[nombre_pdf] = {
                "txt": construir_txt(nro_poliza, polizas_grupo),
                "pdf_bytes": pdf_bytes,
                "pdf_filename": uploaded_file.name,
                "nro_poliza": nro_poliza,
            }
        except Exception as e:
            errores_pdf.append((uploaded_file.name, str(e)))

        progreso.progress(idx_pdf / total_pdfs, text=f"Procesando PDFs... ({idx_pdf}/{total_pdfs})")

    progreso.empty()

    df = pd.DataFrame(all_rows, columns=["Póliza", "Cliente", "Vigencia", "Sección", "Ítem", "Placa", "Marca", "Modelo", "Año", "Valor Asegurado", "Prima Neta"])
    sin_poliza = [nombre for nombre, info in carpetas.items() if info["nro_poliza"] == "SIN_POLIZA"]

    # Filtrar líneas por prefijo directamente sobre los TXT ya generados (sin subir nada aparte)
    PREFIJOS_FILTRO = ('121', '101', '301', '203', '260')
    lineas_filtradas = []
    for nombre_pdf, info in carpetas.items():
        nombre_txt = f"{info['nro_poliza']}.txt"
        for linea in info["txt"].splitlines():
            if linea.startswith(PREFIJOS_FILTRO):
                lineas_filtradas.append({'archivo': nombre_txt, 'linea': linea.strip()})
    df_filtro = pd.DataFrame(lineas_filtradas)

    cuenta_archivos = pd.DataFrame(columns=['archivo', 'cantidad'])
    total_polizas = 0
    if not df_filtro.empty:
        cuenta_archivos = df_filtro['archivo'].value_counts().reset_index()
        cuenta_archivos.columns = ['archivo', 'cantidad']
        total_polizas = int(cuenta_archivos['cantidad'].sum())

    st.success("✅ Archivos procesados correctamente")

    m1, m2, m3 = st.columns(3)
    m1.metric("PDFs procesados", total_pdfs)
    m2.metric("Carpetas generadas", len(carpetas) if pg_file else 0)
    m3.metric("Nro de pólizas", total_polizas)

    if sin_poliza:
        st.warning(f"⚠️ No se pudo extraer el número de póliza de: {', '.join(sin_poliza)}")

    if errores_pdf:
        with st.expander(f"❌ {len(errores_pdf)} PDF con error al procesar", expanded=True):
            for nombre, msg in errores_pdf:
                st.error(f"{nombre}: {msg}")

    if sin_items:
        st.warning(f"⚠️ Estos PDF no generaron ningún ítem (revisar formato): {', '.join(sin_items)}")

    with st.expander(f"📊 Vista previa de extracción (mostrando 5 de {len(df)})", expanded=False):
        st.dataframe(df.head(5), use_container_width=True)

    if pg_file:
        sin_match = [nombre for nombre, info in carpetas.items() if info["nro_poliza"] not in mapa_pg]
        if sin_match:
            st.warning(f"⚠️ No se encontraron pólizas en pg.txt para: {', '.join(sin_match)}")

        polizas_extraidas = {info["nro_poliza"] for info in carpetas.values()}
        grupos_no_usados = sorted(g for g in mapa_pg if g not in polizas_extraidas)
        if grupos_no_usados:
            st.warning(f"⚠️ pg.txt tiene grupos que no aparecen en ningún PDF subido: {', '.join(grupos_no_usados)}")

    if not df_filtro.empty:
        with st.expander(f"🔍 Líneas filtradas por prefijo (mostrando 5 de {len(df_filtro)})", expanded=False):
            st.dataframe(df_filtro.head(5), use_container_width=True)
            st.caption("Cuenta por archivo:")
            st.dataframe(cuenta_archivos.head(5), use_container_width=True)

    # Excel de extracción (se descarga aparte, no va dentro del ZIP)
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Extraccion", index=False)
        if not df_filtro.empty:
            df_filtro.to_excel(writer, sheet_name="Lineas Filtradas", index=False)
            cuenta_archivos.to_excel(writer, sheet_name="Cuenta por Archivo", index=False)

    # ZIP: una carpeta por PDF con su PDF y su TXT
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for nombre_pdf, info in carpetas.items():
            zf.writestr(f"{nombre_pdf}/{info['pdf_filename']}", info["pdf_bytes"])
            zf.writestr(f"{nombre_pdf}/{info['nro_poliza']}.txt", info["txt"])

    sello_fecha = datetime.now().strftime("%Y%m%d")

    st.divider()
    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            "⬇️ Descargar Excel",
            data=excel_buffer.getvalue(),
            file_name=f"Renovaciones_{sello_fecha}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="polidata_download",
            use_container_width=True,
        )
    with col2:
        st.download_button(
            "📁 Descargar carpetas (ZIP)",
            data=zip_buffer.getvalue(),
            file_name=f"Carpetas_POLIDATA_{sello_fecha}.zip",
            mime="application/zip",
            key="carpetas_zip_download",
            use_container_width=True,
            disabled=not pg_file,
            help=None if pg_file else "Sube pg.txt para habilitar esta descarga",
        )
