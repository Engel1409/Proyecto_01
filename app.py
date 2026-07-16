import os
import re
import io
import warnings
import pdfplumber
import pandas as pd
import streamlit as st
import openpyxl
import datetime
from datetime import datetime as dt
import unicodedata
from io import BytesIO
from docx import Document
from openpyxl.styles import PatternFill, Font

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ---------------------------------------------------------
# CONFIGURACIÓN DE PÁGINA
# ---------------------------------------------------------
st.set_page_config(page_title="Suite Operativa", layout="wide")

st.markdown("""
    <style>
    #MainMenu, footer, header { visibility: hidden; }
    .stApp { background-color: #f6f8fa; font-family: 'Segoe UI', Arial, sans-serif; }
    h1 { color: #0a3d62; border-left: 5px solid #DA291C; padding-left: 10px; }
    div.stButton > button { background-color: #DA291C !important; color: white !important; font-weight: bold !important; border-radius: 8px !important; }
    div.stDownloadButton > button { background-color: #1e293b !important; color: white !important; font-weight: bold !important; border-radius: 8px !important; }
    </style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# TABS (orden: POLIDATA, Cálculo de Primas, Normalizador, Filtrador TXT)
# ---------------------------------------------------------
tab1, tab2, tab3, tab4 = st.tabs([
    "📄 POLIDATA (PDF)",
    "📊 CÁLCULO DE PRIMAS",
    "📝 NORMALIZADOR",
    "🔍 FILTRADOR (TXT)"
])

# ==========================================================
# TAB 1: POLIDATA
# ==========================================================
with tab1:
    st.title("📄 POLIDATA")
    st.caption("Extracción automática de datos desde PDF")

    uploaded_files = st.file_uploader("Sube tus archivos PDF aquí", type="pdf", accept_multiple_files=True, key="pdf_uploader")

    if uploaded_files:
        all_rows = []
        for uploaded_file in uploaded_files:
            with pdfplumber.open(uploaded_file) as pdf:
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
                content = text[seccion_indices[i][0]:seccion_indices[i+1][0]]
                lineas = content.split("\n")
                for idx, line in enumerate(lineas):
                    match = re.match(r"^(.*?)(\d{1,3}(?:,\d{3})*\.\d{2})\s+(\d{1,3}(?:,\d{3})*\.\d{2})$", line.strip())
                    if match:
                        item_texto = match.group(1).strip()

                        # Extraer placa si existe en la línea (SECCION: 006 VEHICULOS)
                        placa_match = re.search(r"PLACA:\s*([A-Z0-9\-]+)", item_texto, re.IGNORECASE)
                        placa = placa_match.group(1).strip() if placa_match else ""

                        # Si hay placa, la marca/modelo/año suelen estar en la línea siguiente
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

        df = pd.DataFrame(all_rows, columns=["Póliza", "Cliente", "Vigencia", "Sección", "Ítem", "Placa", "Marca", "Modelo", "Año", "Valor Asegurado", "Prima Neta"])
        st.success("✅ Archivos procesados correctamente")
        st.dataframe(df, use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
        st.download_button("⬇️ Descargar Excel", data=output.getvalue(), file_name="Renovaciones.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="polidata_download")

# ==========================================================
# TAB 2: CÁLCULO DE PRIMAS
# ==========================================================
with tab2:
    st.title("📊 Validación y Cálculo de Primas - Seguros 📊")

    zona = st.selectbox("Selecciona la zona", options=["Sur", "Norte"], index=0, key="primas_zona")
    NETA = 0.00038 if zona == "Sur" else 0.00036
    V_D_E = 0.03
    V_IGV = 0.18

    usuarios = ["Sofi B", "Engel B", "User_01", "User_02"]
    usuario_seleccionado = st.selectbox("Selecciona tu usuario:", usuarios, key="primas_usuario")

    fecha_reporte = dt.now().strftime("%d/%m/%Y %H:%M:%S")
    st.write(f"📅 **Fecha del reporte:** {fecha_reporte}")

    archivos = st.file_uploader("Sube tus archivos Excel", type=["xlsx"], accept_multiple_files=True, key="primas_uploader")

    if st.button("Procesar archivos", key="primas_procesar") and archivos:

        no_validos = []
        resumen = []

        def validar_documento(row):
            tipo = str(row.get("Tipo de Documento", "")).strip().upper()
            num = str(row.get("Número de Documento", "")).strip()
            if tipo == "DNI":
                return "DNI válido" if num.isdigit() and len(num) == 8 else "DNI inválido"
            return "No es DNI"

        for archivo in archivos:
            nombre_archivo = archivo.name

            df_p = pd.read_excel(archivo, dtype={"Número de Documento": str})
            df_p.columns = df_p.columns.str.strip()
            df_p = df_p.dropna(how="all")
            df_p["fila_en_excel"] = df_p.index + 2

            if df_p.empty:
                resumen.append({"Archivo": nombre_archivo, "Poliza": "no declara"})
                continue

            for col in ["Tipo de Documento", "Número de Documento", "Capital Asegurado", "Prima"]:
                if col not in df_p.columns:
                    df_p[col] = pd.NA
            if "Nombre Completo" not in df_p.columns:
                df_p["Nombre Completo"] = pd.NA

            df_p["validación documento"] = df_p.apply(validar_documento, axis=1)

            df_no_validos = df_p[df_p["validación documento"] == "No es DNI"].copy()
            df_no_validos["archivo_origen"] = nombre_archivo

            columnas_finales = [
                "Tipo de Documento", "Número de Documento", "Nombre Completo",
                "validación documento", "archivo_origen", "fila_en_excel"
            ]
            for col in columnas_finales:
                if col not in df_no_validos.columns:
                    df_no_validos[col] = pd.NA
            df_no_validos = df_no_validos[columnas_finales]

            df_no_validos = df_no_validos[
                df_no_validos["Número de Documento"].notna() &
                df_no_validos["Número de Documento"].astype(str).str.strip().ne("") &
                df_no_validos["Nombre Completo"].notna() &
                df_no_validos["Nombre Completo"].astype(str).str.strip().ne("")
            ]

            if not df_no_validos.empty:
                no_validos.append(df_no_validos)

            ultima_es_subtotal = df_p.iloc[-1].astype(str).str.contains("TOTAL", case=False, na=False).any()

            if ultima_es_subtotal and len(df_p) > 1:
                ultima_fila = df_p.iloc[-1]
                df_sin_ultima = df_p.iloc[:-1].copy()
                sub_capital = ultima_fila.get("Capital Asegurado", "no declara")
                sub_prima = ultima_fila.get("Prima", "no declara")
            else:
                df_sin_ultima = df_p.copy()
                sub_capital = "no declara"
                sub_prima = "no declara"

            total_capital_num = df_sin_ultima["Capital Asegurado"].sum(min_count=1)

            s = (df_sin_ultima["Prima"].astype(str)
                 .str.replace('\u00A0', '', regex=False)
                 .str.replace('\u202F', '', regex=False)
                 .str.replace(' ', '', regex=False)
                 .str.replace('S/', '', regex=False)
                 .str.replace('s/', '', regex=False)
                 .str.replace('.', '', regex=False)
                 .str.replace(',', '.', regex=False))

            total_prima_num = pd.to_numeric(s, errors="coerce").sum(min_count=1)

            capital_num = pd.to_numeric(df_sin_ultima["Capital Asegurado"], errors="coerce")

            prima_neta_reg = capital_num * NETA
            d_e_reg = prima_neta_reg * V_D_E
            igv_reg = (prima_neta_reg + d_e_reg) * V_IGV
            total_reg = prima_neta_reg + d_e_reg + igv_reg

            def red2(x):
                return float(round(x, 2)) if pd.notna(x) else "no declara"

            suma_prima_neta = red2(prima_neta_reg.sum(min_count=1))
            suma_d_e = red2(d_e_reg.sum(min_count=1))
            suma_igv = red2(igv_reg.sum(min_count=1))
            suma_total = red2(total_reg.sum(min_count=1))

            match = re.search(r'\d{10,}', nombre_archivo)
            poliza = match.group(0) if match else "no declara"

            resumen.append({
                "Archivo": nombre_archivo,
                "Poliza": poliza,
                "Usuario": usuario_seleccionado,
                "Zona": zona,
                "Fecha_reporte": fecha_reporte,
                "Cantidad_registros": len(df_sin_ultima),
                "Total_capital": total_capital_num,
                "Total_origen_col_H": sub_capital,
                "Total_origen_col_J": sub_prima,
                "prima_neta": suma_prima_neta,
                "D_E": suma_d_e,
                "IGV": suma_igv,
                "TOTAL": suma_total
            })

        df_no_validos_final = pd.concat(no_validos, ignore_index=True) if no_validos else pd.DataFrame()
        df_resumen = pd.DataFrame(resumen)

        orden_cols = [
            "Archivo", "Poliza", "Usuario", "Zona", "Fecha_reporte",
            "Cantidad_registros", "Total_capital",
            "Total_origen_col_H", "Total_origen_col_J",
            "prima_neta", "D_E", "IGV", "TOTAL"
        ]
        df_resumen = df_resumen[orden_cols]

        st.subheader("Vista previa de datos")
        st.write("**Totales por archivo:**")
        st.dataframe(df_resumen)
        st.write("**No válidos:**")
        st.dataframe(df_no_validos_final)

        output_primas = io.BytesIO()
        with pd.ExcelWriter(output_primas, engine="openpyxl") as writer:
            df_resumen.to_excel(writer, sheet_name="Totales por archivo", index=False)
            df_no_validos_final.to_excel(writer, sheet_name="No válidos", index=False)

            wb_primas = writer.book
            fill = PatternFill(start_color="D53032", end_color="D53032", fill_type="solid")
            font_white = Font(color="FFFFFF", bold=True)

            hojas = ["Totales por archivo", "No válidos"]
            for hoja in hojas:
                ws = wb_primas[hoja]
                for cell in ws[1]:
                    cell.fill = fill
                    cell.font = font_white

        st.success("✅ Proceso completado.")
        st.download_button(
            label="📥 Descargar reporte final",
            data=output_primas.getvalue(),
            file_name="Resumen_Validacion.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="primas_download"
        )

# ==========================================================
# TAB 3: NORMALIZADOR WORD Y EXCEL
# ==========================================================

COLUMNAS_EXCLUIDAS = [
    "poliza", "nro_documento", "documento", "id", "ruc", "dni",
    "nro_asegurados", "asegurados", "vigencia", "plazo", "ano", "periodo",
    "nro", "ciiu_giro_del_negocio", "vigencia_inicio", "vigencia_termino",
    "plazo_asegurar", "pisos", "sotanos", "recibo", "aseg", "fecha"
]

def normalizar(texto):
    if pd.isna(texto):
        return ""
    texto = str(texto).strip().lower()
    texto = unicodedata.normalize('NFKD', texto)
    texto = texto.encode('ascii', 'ignore').decode('utf-8')
    texto = re.sub(r"[ .\-\/]+", "_", texto)
    texto = re.sub(r"[^a-z0-9_]", "", texto)
    texto = re.sub(r"_+", "_", texto)
    return texto.strip("_")


def formatear_por_columna(val, nombre_columna):
    if pd.isna(val) or str(val).strip() == "":
        return ""
    if isinstance(val, (datetime.datetime, datetime.date)):
        return val.strftime("%d/%m/%Y")
    val_str = str(val).strip()
    if any(clave in nombre_columna for clave in COLUMNAS_EXCLUIDAS):
        if val_str.endswith('.0'):
            val_str = val_str[:-2]
        if "-" in val_str and len(val_str) >= 10 and val_str[:4].isdigit():
            try:
                fecha_corta = val_str.split(" ")[0]
                partes = fecha_corta.split("-")
                if len(partes) == 3:
                    return f"{partes[2]}/{partes[1]}/{partes[0]}"
            except Exception:
                pass
        return val_str.upper()
    try:
        num = float(val)
        return f"{num:,.2f}"
    except ValueError:
        return val_str.upper()


def procesar_parrafo(paragraph):
    full_text = "".join(run.text for run in paragraph.runs)
    if "{{" not in full_text:
        return
    variables = re.findall(r"{{(.*?)}}", full_text)
    for var in variables:
        nueva = normalizar(var)
        full_text = full_text.replace("{{" + var + "}}", "{{" + nueva + "}}")
    index = 0
    for run in paragraph.runs:
        length = len(run.text)
        run.text = full_text[index:index + length]
        index += length


def normalizar_word(doc):
    for para in doc.paragraphs:
        procesar_parrafo(para)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    procesar_parrafo(para)


def extraer_tags_word(doc):
    tags = set()
    patron = re.compile(r"{{(.*?)}}")
    for para in doc.paragraphs:
        for var in patron.findall(para.text):
            tags.add(normalizar(var))
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for var in patron.findall(para.text):
                        tags.add(normalizar(var))
    return tags


with tab3:
    st.title("📄 Normalizador Word y Excel 📄")
    st.write("Recuerda que en el Word los tags {{}} deben coincidir con las cabeceras del excel.")

    word_file = st.file_uploader("📄 Subir Word", type=["docx"], key="word_uploader")
    excel_file = st.file_uploader("📊 Subir Excel", type=["xlsx"], key="excel_uploader")

    if word_file and excel_file:
        errores = []

        if not word_file.name.endswith(".docx"):
            st.error("❌ El archivo Word debe ser formato .docx")
            st.stop()
        if not excel_file.name.endswith(".xlsx"):
            st.error("❌ El archivo Excel debe ser formato .xlsx")
            st.stop()

        with st.spinner("Procesando archivos..."):

            # WORD
            try:
                doc = Document(word_file)
            except Exception as e:
                errores.append(f"❌ No se pudo leer el archivo Word: {e}")
                doc = None

            tags_word = set()
            word_buffer = None

            if doc:
                try:
                    tags_word = extraer_tags_word(doc)
                    normalizar_word(doc)
                except Exception as e:
                    errores.append(f"❌ Error al normalizar el Word: {e}")
                try:
                    word_buffer = BytesIO()
                    doc.save(word_buffer)
                except Exception as e:
                    errores.append(f"❌ Error al guardar el Word normalizado: {e}")
                    word_buffer = None

            # EXCEL
            try:
                wb = openpyxl.load_workbook(excel_file)
                sheet = wb.active
            except Exception as e:
                errores.append(f"❌ No se pudo leer el archivo Excel: {e}")
                wb = None
                sheet = None

            excel_buffer = None
            columnas_normalizadas = []
            indices_filtrados = []

            if sheet:
                try:
                    for col_idx in range(1, sheet.max_column + 1):
                        celda = sheet.cell(row=1, column=col_idx)
                        nombre = normalizar(celda.value) if celda.value else f"col_{col_idx}"
                        celda.value = nombre
                        columnas_normalizadas.append(nombre)

                    idx_nro = columnas_normalizadas.index("nro") + 1 if "nro" in columnas_normalizadas else None
                    filas_a_borrar = []

                    for row_idx in range(2, sheet.max_row + 1):
                        if idx_nro:
                            valor = str(sheet.cell(row=row_idx, column=idx_nro).value or "").strip()
                            if valor == "" or valor == "0":
                                filas_a_borrar.append(row_idx)
                                continue
                        for col_idx in range(1, sheet.max_column + 1):
                            celda = sheet.cell(row=row_idx, column=col_idx)
                            nombre_col = columnas_normalizadas[col_idx - 1]
                            celda.value = formatear_por_columna(celda.value, nombre_col)

                    for row_idx in reversed(filas_a_borrar):
                        sheet.delete_rows(row_idx)

                except Exception as e:
                    errores.append(f"❌ Error al procesar los datos del Excel: {e}")

                try:
                    cabeceras = [sheet.cell(row=1, column=c).value for c in range(1, sheet.max_column + 1)]
                    indices_filtrados = [i for i, cab in enumerate(cabeceras) if cab in tags_word]

                    if not indices_filtrados:
                        errores.append("⚠️ Advertencia: Ninguna columna del Excel coincide con los tags del Word.")

                    wb_filtrado = openpyxl.Workbook()
                    ws_filtrado = wb_filtrado.active

                    for row_idx in range(1, sheet.max_row + 1):
                        nueva_fila = [sheet.cell(row=row_idx, column=i + 1).value for i in indices_filtrados]
                        ws_filtrado.append(nueva_fila)

                    excel_buffer = BytesIO()
                    wb_filtrado.save(excel_buffer)

                except Exception as e:
                    errores.append(f"❌ Error al generar el Excel filtrado: {e}")
                    excel_buffer = None

        # ERRORES
        if errores:
            st.markdown("---")
            st.subheader("⚠️ Errores y advertencias")
            for err in errores:
                if err.startswith("❌"):
                    st.error(err)
                else:
                    st.warning(err)
            st.markdown("---")

        errores_criticos = [e for e in errores if e.startswith("❌")]

        if not errores_criticos:
            st.subheader("✅ Previsualización")

            try:
                df_vista = pd.DataFrame(sheet.values)
                if not df_vista.empty:
                    df_vista.columns = df_vista.iloc[0]
                    df_vista = df_vista[1:]
                    df_vista = df_vista.loc[:, ~df_vista.columns.duplicated()]

                    cabeceras_ordenadas = [sheet.cell(row=1, column=i + 1).value for i in indices_filtrados]
                    cols_mostrar = [c for c in cabeceras_ordenadas if c in df_vista.columns]

                    if cols_mostrar:
                        st.info(f"🔍 Mostrando {len(cols_mostrar)} columna(s) usadas en el Word: `{'`, `'.join(sorted(cols_mostrar))}`")
                        st.dataframe(df_vista[cols_mostrar].head(5), use_container_width=True)
                    else:
                        st.warning("⚠️ No se encontraron columnas que coincidan con los tags del Word.")
            except Exception as e:
                st.error(f"❌ Error al generar la previsualización: {e}")

            if tags_word:
                with st.expander("🏷️ Tags detectados en el Word"):
                    st.write(sorted(tags_word))

            st.success("✅ Archivos listos para descargar")

            col1, col2 = st.columns(2)
            with col1:
                if word_buffer:
                    st.download_button("📄 Descargar Word", word_buffer.getvalue(), "word_normalizado.docx", key="norm_download_word")
                else:
                    st.error("❌ Word no disponible para descarga.")
            with col2:
                if excel_buffer:
                    st.download_button("📊 Descargar Excel", excel_buffer.getvalue(), "excel_limpio.xlsx", key="norm_download_excel")
                else:
                    st.error("❌ Excel no disponible para descarga.")

# ==========================================================
# TAB 4: FILTRADOR TXT
# ==========================================================
with tab4:
    st.title("📄 Filtrar líneas (TXT)")
    txt_archivos = st.file_uploader("Sube tus archivos .txt", type=["txt"], accept_multiple_files=True, key="txt_uploader")
    prefijos = ('121', '101', '301', '203', '260')

    if st.button("Procesar TXT", key="txt_procesar") and txt_archivos:
        lineas_filtradas = []
        for archivo in txt_archivos:
            contenido = archivo.read().decode('utf-8', errors='ignore')
            for linea in contenido.splitlines():
                if linea.startswith(prefijos):
                    lineas_filtradas.append({'archivo': archivo.name, 'linea': linea.strip()})

        df_txt = pd.DataFrame(lineas_filtradas)
        if not df_txt.empty:
            st.dataframe(df_txt, use_container_width=True)
            st.download_button("📥 Descargar CSV", data=df_txt.to_csv(index=False), file_name="filtrado.csv", mime="text/csv", key="txt_download")
        else:
            st.warning("No se encontraron líneas con los prefijos seleccionados.")
