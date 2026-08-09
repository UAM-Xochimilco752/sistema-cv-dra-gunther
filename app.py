import io
import json
import os
import pickle
import pandas as pd
import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

# ----------------------------------------------------
# CONFIGURACIÓN GENERAL Y CONSTANTES
# ----------------------------------------------------
st.set_page_config(
    page_title="Control de CV y Probatorios - Dra. Günther",
    page_icon="📄",
    layout="wide",
)

CATEGORIAS = [
    "Coordinación de Libros",
    "Capítulos de Libros / Artículos",
    "Ponencias y Conferencias",
    "Presentaciones de Libros",
    "Comisiones y Arbitrajes",
    "Cursos e Impartición de Clases",
    "Premios y Reconocimientos",
]

ESTILOS_DISENOS = {
    "Azul Ejecutivo": {
        "fuente": "Calibri",
        "color_titulo": RGBColor(0, 51, 102),     # Azul Marino
        "color_sub": RGBColor(100, 100, 100),    # Gris Medio
        "color_cat": RGBColor(0, 51, 102),
    },
    "Académico Clásico": {
        "fuente": "Times New Roman",
        "color_titulo": RGBColor(102, 0, 0),     # Tinto / Guinda Académico
        "color_sub": RGBColor(80, 80, 80),
        "color_cat": RGBColor(102, 0, 0),
    },
    "Minimalista": {
        "fuente": "Arial",
        "color_titulo": RGBColor(40, 40, 40),     # Gris Oscuro
        "color_sub": RGBColor(120, 120, 120),
        "color_cat": RGBColor(40, 40, 40),
    }
}

# ----------------------------------------------------
# FUNCIONES DE CONEXIÓN CON GOOGLE DRIVE API
# ----------------------------------------------------
@st.cache_resource
def obtener_servicio_drive():
    """Autentica al usuario usando el token guardado en la nube."""
    creds = None
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)

    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
        except Exception as e:
            st.error(f"Error al refrescar credenciales de Google: {e}")
            return None

    if not creds:
        st.error("⚠️ No se encontró el archivo 'token.pickle' en el proyecto.")
        return None

    return build("drive", "v3", credentials=creds)


def buscar_excel_en_drive(service):
    """Busca el archivo Excel de la base de datos en Google Drive."""
    try:
        results = service.files().list(
            q="trashed = false",
            fields="files(id, name)"
        ).execute()
        files = results.get('files', [])
        
        for f in files:
            if "Base_de_Datos_Probatorios_y_CV" in f['name']:
                return f['id'], f['name']
                
        return None, None
    except Exception as e:
        st.error(f"Error al consultar Google Drive: {e}")
        return None, None


def cargar_datos_drive(service, file_id):
    """Lee el Excel de Google Drive directamente a un DataFrame de pandas."""
    request = service.files().get_media(fileId=file_id)
    file_buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(file_buffer, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    file_buffer.seek(0)
    return pd.read_excel(file_buffer)


def subir_a_google_drive(service, nombre_archivo, bytes_archivo):
    """Sube el archivo PDF a Google Drive y devuelve su enlace de visualización."""
    try:
        file_metadata = {"name": nombre_archivo}
        media = MediaIoBaseUpload(
            io.BytesIO(bytes_archivo),
            mimetype="application/pdf",
            resumable=True
        )

        file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields="id, webViewLink"
        ).execute()

        file_id = file.get("id")

        # Hacer el archivo accesible mediante enlace
        service.permissions().create(
            fileId=file_id,
            body={'type': 'anyone', 'role': 'reader'}
        ).execute()

        return file.get("webViewLink")
    except Exception as e:
        st.error(f"Error al subir archivo a Google Drive: {e}")
        return None


def actualizar_excel_drive(service, file_id, df):
    """Sobreescribe la base de datos en Google Drive con los nuevos datos."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    
    media = MediaIoBaseUpload(
        output, 
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    service.files().update(fileId=file_id, media_body=media).execute()


# ----------------------------------------------------
# FUNCIONES AUXILIARES
# ----------------------------------------------------
def limpiar_texto(val):
    """Limpia textos vacíos, descarte de palabras basura o valores no deseados."""
    if pd.isna(val):
        return ""
    texto = str(val).strip()
    descartes = ["ninguno", "ninguna", "n/a", "na", "sin_pdf", "sin_enlace", "nan", "none", "-"]
    if texto.lower() in descartes:
        return ""
    return texto


def formatear_fecha_cv(row):
    """Obtiene una cadena de fecha limpia sin horas ni timestamps."""
    fecha_val = limpiar_texto(row.get("Fecha"))
    anio_val = row.get("Año")

    if fecha_val:
        if "00:00:00" in fecha_val:
            fecha_val = fecha_val.split(" ")[0].strip()
        try:
            dt = pd.to_datetime(fecha_val)
            return dt.strftime("%d/%m/%Y")
        except Exception:
            return fecha_val
    elif pd.notna(anio_val) and str(anio_val).replace(".", "").isdigit():
        return str(int(float(anio_val)))
    return ""


# ----------------------------------------------------
# PROCESAMIENTO E INTELIGENCIA ARTIFICIAL (GEMINI API)
# ----------------------------------------------------
def procesar_cv_con_gemini(df_cv, objetivo, api_key):
    """Envía los registros académicos a Gemini para curación y redacción fina."""
    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel("gemini-2.5-flash")

        # Convertir datos relevantes a lista de diccionarios
        registros = []
        for _, row in df_cv.iterrows():
            registros.append({
                "Año": str(row.get("Año", "")),
                "Categoría": limpiar_texto(row.get("Categoría_CV") or row.get("Categoría")),
                "Título": limpiar_texto(row.get("Título_Actividad_o_Publicación") or row.get("Título")),
                "Rol": limpiar_texto(row.get("Rol_Participación")),
                "Institución": limpiar_texto(row.get("Institución_Organización") or row.get("Institución")),
                "Lugar": limpiar_texto(row.get("Lugar_Sede")),
            })

        prompt = f"""
        Actúa como un experto consultor editorial y curricular académico universitario de alto nivel.
        Estás elaborando el Currículum Vitae oficial de la Dra. María Griselda Günther.

        OBJETIVO DEL CV: "{objetivo}"
        DATOS DE ORIGEN:
        {json.dumps(registros, ensure_ascii=False, indent=2)}

        INSTRUCCIONES:
        1. Selecciona y curadoramente filtra las actividades de la Dra. que mejor se alineen al objetivo "{objetivo}".
        2. Mejora la redacción académica de cada título, rol e institución (corrige faltas de ortografía, capitalización y redacción formal).
        3. Escribe una breve **Síntesis / Perfil Profesional Ejecutivo** de 1 párrafo (máximo 5 líneas) introductorio adaptado al objetivo seleccionado.
        4. Agrupa los ítems en categorías claras.

        Responde ÚNICAMENTE en formato JSON válido con la siguiente estructura:
        {{
            "subtitulo_cv": "SÍNTESIS EJECUTIVA / EVALUACIÓN ACADÉMICA",
            "perfil_ejecutivo": "Texto del párrafo de perfil introductorio...",
            "secciones": [
                {{
                    "categoria": "Nombre de la Categoría",
                    "items": [
                        {{
                            "titulo": "Título estilizado y corregido",
                            "detalles": "Rol, Institución, Sede, Año/Fecha formateados de forma impecable."
                        }}
                    ]
                }}
            ]
        }}
        """

        response = model.generate_content(
            prompt,
            generation_config={"response_mime_type": "application/json"}
        )
        return json.loads(response.text)

    except Exception as e:
        st.error(f"Error al procesar con la IA: {e}")
        return None


# ----------------------------------------------------
# GENERADOR DE WORD (.DOCX) CON PLANTILLA
# ----------------------------------------------------
def crear_cv_word_desde_json(datos_ia, nombre_estilo="Azul Ejecutivo"):
    """Genera el Word maquetado profesionalmente a partir del JSON estructurado por Gemini."""
    doc = Document()
    estilo_cfg = ESTILOS_DISENOS.get(nombre_estilo, ESTILOS_DISENOS["Azul Ejecutivo"])

    # Márgenes de página profesionales (2.5 cm)
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # Estilo base
    style_normal = doc.styles["Normal"]
    style_normal.font.name = estilo_cfg["fuente"]
    style_normal.font.size = Pt(11)

    # Encabezado: Nombre
    p_titulo = doc.add_paragraph()
    p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_titulo.paragraph_format.space_after = Pt(2)

    run_nombre = p_titulo.add_run("DRA. MARÍA GRISELDA GÜNTHER")
    run_nombre.bold = True
    run_nombre.font.size = Pt(16)
    run_nombre.font.color.rgb = estilo_cfg["color_titulo"]

    # Subtítulo
    subtitulo_texto = datos_ia.get("subtitulo_cv", "CURRÍCULUM VITAE").upper()
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(14)

    run_sub = p_sub.add_run(f"— {subtitulo_texto} —")
    run_sub.font.size = Pt(10.5)
    run_sub.font.italic = True
    run_sub.font.color.rgb = estilo_cfg["color_sub"]

    # Perfil Ejecutivo
    perfil = datos_ia.get("perfil_ejecutivo", "")
    if perfil:
        p_perfil = doc.add_paragraph()
        p_perfil.paragraph_format.space_after = Pt(14)
        p_perfil.paragraph_format.line_spacing = 1.15
        run_p = p_perfil.add_run(perfil)
        run_p.font.italic = True

    # Secciones
    for sec in datos_ia.get("secciones", []):
        cat_nombre = sec.get("categoria", "")
        items = sec.get("items", [])

        if not items:
            continue

        p_cat = doc.add_paragraph()
        p_cat.paragraph_format.space_before = Pt(14)
        p_cat.paragraph_format.space_after = Pt(6)
        p_cat.paragraph_format.keep_with_next = True

        run_cat = p_cat.add_run(cat_nombre)
        run_cat.bold = True
        run_cat.font.size = Pt(12.5)
        run_cat.font.color.rgb = estilo_cfg["color_cat"]

        for item in items:
            titulo_item = item.get("titulo", "")
            detalles_item = item.get("detalles", "")

            p_item = doc.add_paragraph(style="List Bullet")
            p_item.paragraph_format.space_after = Pt(4)
            p_item.paragraph_format.line_spacing = 1.15

            if titulo_item:
                r_t = p_item.add_run(titulo_item)
                r_t.bold = True
            if detalles_item:
                if titulo_item:
                    p_item.add_run(". ")
                p_item.add_run(detalles_item)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


def crear_cv_word_tradicional(df, nombre_estilo="Azul Ejecutivo"):
    """Genera el CV en Word de forma directa (Sin procesamiento de IA)."""
    doc = Document()
    estilo_cfg = ESTILOS_DISENOS.get(nombre_estilo, ESTILOS_DISENOS["Azul Ejecutivo"])

    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    style_normal = doc.styles["Normal"]
    style_normal.font.name = estilo_cfg["fuente"]
    style_normal.font.size = Pt(11)

    p_titulo = doc.add_paragraph()
    p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_nombre = p_titulo.add_run("DRA. MARÍA GRISELDA GÜNTHER")
    run_nombre.bold = True
    run_nombre.font.size = Pt(16)
    run_nombre.font.color.rgb = estilo_cfg["color_titulo"]

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(16)
    run_sub = p_sub.add_run("CURRÍCULUM VITAE — SÍNTESIS GENERAL")
    run_sub.font.size = Pt(10.5)
    run_sub.font.italic = True
    run_sub.font.color.rgb = estilo_cfg["color_sub"]

    col_incluir = "Incluir_en_CV" if "Incluir_en_CV" in df.columns else df.columns[0]
    df_cv = df[df[col_incluir].astype(str).str.strip().str.lower() == "sí"].copy()
    
    if "Año" in df_cv.columns:
        df_cv["Año_num"] = pd.to_numeric(df_cv["Año"], errors="coerce")
        df_cv = df_cv.sort_values(by=["Año_num"], ascending=False)

    cat_col = "Categoría_CV" if "Categoría_CV" in df_cv.columns else "Categoría"
    categorias_presentes = df_cv[cat_col].unique() if cat_col in df_cv.columns else []

    for cat in CATEGORIAS:
        if cat in categorias_presentes:
            sub_df = df_cv[df_cv[cat_col] == cat]
            if sub_df.empty:
                continue

            p_cat = doc.add_paragraph()
            p_cat.paragraph_format.space_before = Pt(14)
            p_cat.paragraph_format.space_after = Pt(6)
            run_cat = p_cat.add_run(cat)
            run_cat.bold = True
            run_cat.font.size = Pt(12.5)
            run_cat.font.color.rgb = estilo_cfg["color_cat"]

            for _, row in sub_df.iterrows():
                titulo = limpiar_texto(row.get("Título_Actividad_o_Publicación") or row.get("Título"))
                rol = limpiar_texto(row.get("Rol_Participación"))
                inst = limpiar_texto(row.get("Institución_Organización") or row.get("Institución"))
                lugar = limpiar_texto(row.get("Lugar_Sede"))
                fecha_str = formatear_fecha_cv(row)

                if not titulo and not rol and not inst:
                    continue

                p_item = doc.add_paragraph(style="List Bullet")
                p_item.paragraph_format.space_after = Pt(4)
                if titulo:
                    run_t = p_item.add_run(titulo)
                    run_t.bold = True

                detalles = []
                if rol:
                    detalles.append(f"Rol: {rol}" if not rol.lower().startswith("rol") else rol)
                if inst:
                    detalles.append(inst)
                if lugar:
                    detalles.append(lugar)
                if fecha_str:
                    detalles.append(fecha_str)

                if detalles:
                    if titulo:
                        p_item.add_run(". ")
                    p_item.add_run(", ".join(detalles) + ".")

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# ----------------------------------------------------
# INTERFAZ PRINCIPAL (STREAMLIT)
# ----------------------------------------------------
st.title("📄 Sistema de Gestión de CV - Dra. María Griselda Günther")

service = obtener_servicio_drive()

if service:
    excel_id, found_name = buscar_excel_en_drive(service)

    if not excel_id:
        st.warning("⚠️ No se encontró la base de datos en Google Drive.")
        st.info("Sube tu archivo Excel para vincularlo por primera vez:")
        archivo_excel_nuevo = st.file_uploader("Selecciona 'Base_de_Datos_Probatorios_y_CV.xlsx':", type=["xlsx"])
        
        if archivo_excel_nuevo is not None:
            with st.spinner("Subiendo Excel a Google Drive..."):
                file_metadata = {'name': 'Base_de_Datos_Probatorios_y_CV.xlsx'}
                media = MediaIoBaseUpload(io.BytesIO(archivo_excel_nuevo.getvalue()), mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                service.files().create(body=file_metadata, media_body=media, fields='id').execute()
                st.success("¡Base de datos vinculada con éxito!")
                st.rerun()
    else:
        # Cargar DataFrame en vivo desde Drive
        df = cargar_datos_drive(service, excel_id)

        # 3 PESTAÑAS PRINCIPALES
        tab_consulta, tab_registro, tab_cv_word = st.tabs(
            [
                "🔍 Buscar y Consultar Probatorios",
                "➕ Registrar Nueva Actividad",
                "📄 Generar CV Personalizado con IA",
            ]
        )

        # ----------------------------------------------------
        # PESTAÑA 1: BUSCADOR Y CONSULTA
        # ----------------------------------------------------
        with tab_consulta:
            st.subheader("Buscador de Actividades y Documentos")
            st.markdown("##### 🎯 Filtros de Búsqueda")
            col_f1, col_f2, col_f3 = st.columns(3)

            cat_col = "Categoría_CV" if "Categoría_CV" in df.columns else "Categoría"

            with col_f1:
                anios_disponibles = ["Todos"] + sorted(list(df['Año'].dropna().astype(int).unique()), reverse=True) if 'Año' in df.columns else ["Todos"]
                filtro_anio = st.selectbox("Filtrar por Año", anios_disponibles)
            with col_f2:
                filtro_categoria = st.selectbox("Filtrar por Categoría", ["Todas"] + CATEGORIAS)
            with col_f3:
                filtro_texto = st.text_input("🔍 Buscar por palabra clave", placeholder="Ej. Congreso, Comisión, Libro...")

            df_filtrado = df.copy()

            if filtro_anio != "Todos":
                df_filtrado = df_filtrado[df_filtrado["Año"] == int(filtro_anio)]
            if filtro_categoria != "Todas" and cat_col in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado[cat_col] == filtro_categoria]
            if filtro_texto:
                mask = df_filtrado.apply(lambda row: row.astype(str).str.contains(filtro_texto, case=False).any(), axis=1)
                df_filtrado = df_filtrado[mask]

            st.markdown("---")
            st.write(f"Se encontraron **{len(df_filtrado)}** registros coincidentes:")

            col_config = {}
            if "Enlace_Drive_Probatorio" in df_filtrado.columns:
                col_config["Enlace_Drive_Probatorio"] = st.column_config.LinkColumn("Enlace Drive Probatorio")
            elif "Enlace_Probatorio" in df_filtrado.columns:
                col_config["Enlace_Probatorio"] = st.column_config.LinkColumn("Enlace Drive Probatorio")

            st.dataframe(df_filtrado, use_container_width=True, column_config=col_config)

        # ----------------------------------------------------
        # PESTAÑA 2: CAPTURA DE NUEVAS ACTIVIDADES
        # ----------------------------------------------------
        with tab_registro:
            st.subheader("Formulario de Captura de Actividades y Constancias")

            with st.form("form_nueva_actividad", clear_on_submit=True):
                col1, col2 = st.columns(2)

                with col1:
                    nuevo_id = st.text_input("ID de Registro", value=f"ACT-00{len(df) + 1}")
                    anio = st.selectbox("Año", [2026, 2025, 2024, 2023, 2022, 2021, 2020])
                    fecha = st.date_input("Fecha")
                    categoria = st.selectbox("Categoría del CV", CATEGORIAS)
                    rol = st.text_input("Rol / Participación (ej. Ponente, Autora, Coordinadora)")

                with col2:
                    titulo = st.text_input("Título de la Actividad o Publicación *")
                    institucion = st.text_input("Institución u Organización")
                    lugar = st.text_input("Lugar / Sede")
                    estado = st.selectbox("Estado del Probatorio", ["Verificado / En Drive", "Pendiente de Escanear", "En Trámite"])
                    incluir = st.radio("¿Incluir en el CV?", ["Sí", "No"], horizontal=True)

                st.markdown("---")
                st.subheader("📎 Documento Probatorio (PDF)")
                archivo_pdf = st.file_uploader("Sube el PDF escaneado de la constancia", type=["pdf", "png", "jpg", "jpeg"])
                notas = st.text_area("Notas / Observaciones de control interno")

                boton_guardar = st.form_submit_button("💾 Guardar Registro y Subir PDF")

            if boton_guardar:
                if not titulo:
                    st.error("⚠️ El campo 'Título de la Actividad' es obligatorio.")
                else:
                    with st.spinner("Subiendo PDF a Google Drive y actualizando Excel..."):
                        nombre_pdf_guardado = "Sin_PDF"
                        enlace_drive = "Sin_Enlace"

                        if archivo_pdf is not None:
                            bytes_pdf = archivo_pdf.getvalue()
                            titulo_limpio = "".join(x for x in titulo if x.isalnum() or x in " _-")[:20]
                            nombre_pdf_guardado = f"{anio}_{categoria.replace(' ', '_')}_{titulo_limpio}.pdf"

                            # Subir directamente a Google Drive
                            url_drive = subir_a_google_drive(service, nombre_pdf_guardado, bytes_pdf)
                            if url_drive:
                                enlace_drive = url_drive

                        nueva_fila = {
                            "ID": nuevo_id,
                            "Año": anio,
                            "Fecha": str(fecha),
                            "Categoría_CV": categoria,
                            "Rol_Participación": rol,
                            "Título_Actividad_o_Publicación": titulo,
                            "Institución_Organización": institucion,
                            "Lugar_Sede": lugar,
                            "Nombre_Archivo_PDF": nombre_pdf_guardado,
                            "Enlace_Drive_Probatorio": enlace_drive,
                            "Estado_Probatorio": estado,
                            "Incluir_en_CV": incluir,
                            "Notas_Observaciones": notas,
                        }

                        # Mantener compatibilidad con columnas existentes
                        for col in df.columns:
                            if col not in nueva_fila:
                                nueva_fila[col] = ""

                        df_actualizado = pd.concat([df, pd.DataFrame([nueva_fila])], ignore_index=True)
                        actualizar_excel_drive(service, excel_id, df_actualizado)

                        st.success(f"¡Excelente! Registro '{titulo}' guardado exitosamente en Google Drive.")
                        st.balloons()
                        st.rerun()

        # ----------------------------------------------------
        # PESTAÑA 3: GENERADOR DE CV INTELIGENTE EN WORD
        # ----------------------------------------------------
        with tab_cv_word:
            st.subheader("📄 Generador de CV Adaptativo con IA")
            st.write("Genera un documento en Word (.docx) perfectamente redactado y diseñado a la medida según el trámite o evaluación académica.")

            col_inc = "Incluir_en_CV" if "Incluir_en_CV" in df.columns else df.columns[0]
            df_cv_aprobados = df[df[col_inc].astype(str).str.strip().str.lower() == "sí"]

            col_c1, col_c2 = st.columns(2)

            with col_c1:
                objetivo_cv = st.selectbox(
                    "🎯 Objetivo del CV (Perfil)",
                    [
                        "Síntesis Ejecutiva (Resumen de 2 Páginas)",
                        "Evaluación SNI / CONAHCYT (Enfoque Investigación / Publicaciones)",
                        "Perfil Institucional UAM (Docencia, Comisiones y Gestión)",
                        "Semblanza Curricular Narrativa (Para Congresos o Presentaciones)"
                    ]
                )

                estilo_cv = st.selectbox(
                    "🎨 Estilo Visual y Paleta",
                    list(ESTILOS_DISENOS.keys())
                )

            with col_c2:
                usar_ia = st.toggle("🤖 Optimizar y redactar con IA (Gemini)", value=True)

                api_key_input = ""
                if usar_ia:
                    # Intentar obtener la API key desde Secrets de Streamlit
                    api_key_env = st.secrets.get("GEMINI_API_KEY", "") if "GEMINI_API_KEY" in st.secrets else os.getenv("GEMINI_API_KEY", "")
                    
                    if api_key_env:
                        api_key_input = api_key_env
                        st.success("🔑 Llave de Gemini API detectada correctamente.")
                    else:
                        api_key_input = st.text_input("Ingresa tu Gemini API Key:", type="password", help="Obtén tu clave gratis en Google AI Studio")

            st.markdown("---")

            if st.button("🚀 Generar Currículum en Word", type="primary"):
                if df_cv_aprobados.empty:
                    st.warning("⚠️ No hay actividades marcadas con 'Incluir en CV = Sí' en la base de datos.")
                else:
                    buffer_word = None

                    if usar_ia:
                        if not api_key_input:
                            st.error("⚠️ Por favor ingresa una API Key válida de Gemini para utilizar la redacción inteligente.")
                        else:
                            with st.spinner("🤖 La IA está analizando los méritos, mejorando la redacción y maquetando el documento..."):
                                json_cv = procesar_cv_con_gemini(df_cv_aprobados, objetivo_cv, api_key_input)
                                if json_cv:
                                    buffer_word = crear_cv_word_desde_json(json_cv, estilo_cv)
                                    st.success("✨ ¡CV procesado y formateado con éxito por la IA!")
                    else:
                        with st.spinner("Generando documento Word..."):
                            buffer_word = crear_cv_word_tradicional(df, estilo_cv)
                            st.success("✨ Documento generado correctamente (Modo Tradicional).")

                    if buffer_word:
                        st.download_button(
                            label="📥 Descargar CV (.docx)",
                            data=buffer_word,
                            file_name=f"CV_Dra_Gunther_{objetivo_cv.split()[0]}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        )

            st.markdown("---")
            st.markdown("##### 👁️ Actividades que se considerarán:")
            columnas_preview = [c for c in ["Año", "Categoría_CV", "Categoría", "Título_Actividad_o_Publicación", "Título", "Rol_Participación", "Institución_Organización", "Institución"] if c in df_cv_aprobados.columns]
            st.dataframe(df_cv_aprobados[columnas_preview], use_container_width=True)
