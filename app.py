import io
import os
import pickle
import mimetypes

import pandas as pd
import streamlit as st
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
    "Asesorías",
]

# Nombre de la carpeta principal que se creará en Google Drive.
NOMBRE_CARPETA_RAIZ = "CV — Sistema de Gestión"

# Estructura general del repositorio.
ESTRUCTURA_CARPETAS = [
    "00 — Administración",
    "01 — Datos personales y CV",
    "02 — Probatorios",
    "03 — Formación académica",
    "04 — Docencia",
    "05 — Investigación",
    "06 — Ponencias y congresos",
    "07 — Publicaciones",
    "08 — Gestión y cargos",
    "09 — Reconocimientos",
    "10 — CV generados",
]

ANIOS_PROBATORIOS = [2023, 2024, 2025, 2026]


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


# ----------------------------------------------------
# FUNCIONES PARA CREAR Y ADMINISTRAR CARPETAS DE DRIVE
# ----------------------------------------------------
def buscar_carpeta(service, nombre, parent_id=None):
    """
    Busca una carpeta por nombre dentro de una carpeta padre.
    Si parent_id es None, busca en la raíz de Mi Drive.
    """
    try:
        nombre_escapado = nombre.replace("'", "\\'")

        if parent_id:
            query = (
                f"name = '{nombre_escapado}' "
                f"and mimeType = 'application/vnd.google-apps.folder' "
                f"and trashed = false "
                f"and '{parent_id}' in parents"
            )
        else:
            query = (
                f"name = '{nombre_escapado}' "
                f"and mimeType = 'application/vnd.google-apps.folder' "
                f"and trashed = false "
                f"and 'root' in parents"
            )

        resultado = service.files().list(
            q=query,
            spaces="drive",
            fields="files(id, name, webViewLink)",
            pageSize=10,
        ).execute()

        carpetas = resultado.get("files", [])
        return carpetas[0] if carpetas else None

    except Exception as e:
        st.error(f"Error al buscar la carpeta '{nombre}': {e}")
        return None


def crear_carpeta(service, nombre, parent_id=None):
    """Crea una carpeta en Google Drive y devuelve sus datos."""
    try:
        metadata = {
            "name": nombre,
            "mimeType": "application/vnd.google-apps.folder",
        }

        if parent_id:
            metadata["parents"] = [parent_id]

        carpeta = service.files().create(
            body=metadata,
            fields="id, name, webViewLink",
        ).execute()

        return carpeta

    except Exception as e:
        st.error(f"Error al crear la carpeta '{nombre}': {e}")
        return None


def obtener_o_crear_carpeta(service, nombre, parent_id=None):
    """
    Busca una carpeta existente y, si no existe, la crea.
    Esto evita duplicados al ejecutar nuevamente la aplicación.
    """
    carpeta = buscar_carpeta(service, nombre, parent_id)

    if carpeta:
        return carpeta

    return crear_carpeta(service, nombre, parent_id)


@st.cache_resource
def inicializar_estructura_drive(_service):
    """
    Crea la estructura principal del repositorio.
    Devuelve un diccionario con los IDs de las carpetas.
    """
    estructura = {}

    # Carpeta raíz
    raiz = obtener_o_crear_carpeta(_service, NOMBRE_CARPETA_RAIZ)

    if not raiz:
        return {}

    estructura["raiz"] = raiz

    # Carpetas principales
    for nombre in ESTRUCTURA_CARPETAS:
        carpeta = obtener_o_crear_carpeta(
            _service,
            nombre,
            raiz["id"],
        )

        if carpeta:
            estructura[nombre] = carpeta

    # Carpeta de probatorios
    carpeta_probatorios = estructura.get("02 — Probatorios")

    if carpeta_probatorios:
        estructura["probatorios"] = carpeta_probatorios

        # Carpetas por año
        for anio in ANIOS_PROBATORIOS:
            carpeta_anio = obtener_o_crear_carpeta(
                _service,
                str(anio),
                carpeta_probatorios["id"],
            )

            if carpeta_anio:
                estructura[f"probatorios_{anio}"] = carpeta_anio

                # Dentro de cada año se crean las categorías.
                for categoria in CATEGORIAS:
                    carpeta_categoria = obtener_o_crear_carpeta(
                        _service,
                        categoria,
                        carpeta_anio["id"],
                    )

                    if carpeta_categoria:
                        estructura[
                            f"probatorios_{anio}_{categoria}"
                        ] = carpeta_categoria

    return estructura


def obtener_carpeta_probatorio(service, estructura, anio, categoria):
    """
    Devuelve la carpeta específica donde debe guardarse un probatorio:
    02 — Probatorios / Año / Categoría.
    """
    clave = f"probatorios_{anio}_{categoria}"

    if clave in estructura:
        return estructura[clave]

    # Si aparece un año nuevo o una categoría nueva,
    # se crean dinámicamente.
    carpeta_probatorios = estructura.get("probatorios")

    if not carpeta_probatorios:
        return None

    carpeta_anio = obtener_o_crear_carpeta(
        service,
        str(anio),
        carpeta_probatorios["id"],
    )

    if not carpeta_anio:
        return None

    carpeta_categoria = obtener_o_crear_carpeta(
        service,
        categoria,
        carpeta_anio["id"],
    )

    return carpeta_categoria


def obtener_mimetype(nombre_archivo):
    """Obtiene el MIME type correcto para PDF, PNG o JPG."""
    mimetype, _ = mimetypes.guess_type(nombre_archivo)

    if mimetype:
        return mimetype

    return "application/octet-stream"


# ----------------------------------------------------
# FUNCIONES PARA EL EXCEL EN GOOGLE DRIVE
# ----------------------------------------------------
def buscar_excel_en_drive(service):
    """Busca el archivo Excel de la base de datos en Google Drive."""
    try:
        results = service.files().list(
            q="trashed = false",
            fields="files(id, name)",
            pageSize=100,
        ).execute()

        files = results.get("files", [])

        for f in files:
            if "Base_de_Datos_Probatorios_y_CV" in f["name"]:
                return f["id"], f["name"]

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


def actualizar_excel_drive(service, file_id, df):
    """Sobreescribe la base de datos en Google Drive con los nuevos datos."""
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)

    output.seek(0)

    media = MediaIoBaseUpload(
        output,
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        resumable=True,
    )

    service.files().update(
        fileId=file_id,
        media_body=media,
    ).execute()


# ----------------------------------------------------
# FUNCIÓN PARA SUBIR PROBATORIOS A LA CARPETA CORRECTA
# ----------------------------------------------------
def subir_a_google_drive(
    service,
    nombre_archivo,
    bytes_archivo,
    carpeta_destino=None,
):
    """
    Sube un archivo a Google Drive.
    Si carpeta_destino existe, el archivo se guarda dentro de ella.
    """
    try:
        file_metadata = {
            "name": nombre_archivo,
        }

        if carpeta_destino:
            file_metadata["parents"] = [carpeta_destino["id"]]

        media = MediaIoBaseUpload(
            io.BytesIO(bytes_archivo),
            mimetype=obtener_mimetype(nombre_archivo),
            resumable=True,
        )

        file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields="id, name, webViewLink, parents",
        ).execute()

        # NO hacemos público el archivo.
        # Conserva los permisos heredados de la carpeta de Drive.

        return file.get("webViewLink"), file.get("id")

    except Exception as e:
        st.error(f"Error al subir archivo a Google Drive: {e}")
        return None, None


# ----------------------------------------------------
# FUNCIONES AUXILIARES Y GENERACIÓN DE WORD
# ----------------------------------------------------
def limpiar_texto(val):
    """Limpia textos vacíos, descarte de palabras basura o valores no deseados."""
    if pd.isna(val):
        return ""

    texto = str(val).strip()

    descartes = [
        "ninguno",
        "ninguna",
        "n/a",
        "na",
        "sin_pdf",
        "sin_enlace",
        "nan",
        "none",
        "-",
    ]

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


def crear_cv_word(df):
    """Genera el documento Word del CV con formato académico limpio, ejecutivo y organizado."""
    doc = Document()

    # Márgenes de página profesionales (2.5 cm)
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # Estilo base
    style_normal = doc.styles["Normal"]
    style_normal.font.name = "Calibri"
    style_normal.font.size = Pt(11)

    # Encabezado: Nombre de la Dra.
    p_titulo = doc.add_paragraph()
    p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_titulo.paragraph_format.space_after = Pt(2)
    p_titulo.paragraph_format.space_before = Pt(0)

    run_nombre = p_titulo.add_run("DRA. MARÍA GRISELDA GÜNTHER")
    run_nombre.bold = True
    run_nombre.font.size = Pt(16)
    run_nombre.font.color.rgb = RGBColor(0, 51, 102)

    # Subtítulo
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(16)

    run_sub = p_sub.add_run("CURRÍCULUM VITAE — SÍNTESIS EJECUTIVA")
    run_sub.font.size = Pt(10.5)
    run_sub.font.italic = True
    run_sub.font.color.rgb = RGBColor(100, 100, 100)

    # Filtrar solo actividades aprobadas para el CV
    col_incluir = (
        "Incluir_en_CV"
        if "Incluir_en_CV" in df.columns
        else df.columns[0]
    )

    df_cv = df[
        df[col_incluir].astype(str).str.strip().str.lower() == "sí"
    ].copy()

    if "Año" in df_cv.columns:
        df_cv["Año_num"] = pd.to_numeric(
            df_cv["Año"],
            errors="coerce",
        )

        df_cv = df_cv.sort_values(
            by=["Año_num"],
            ascending=False,
        )

    cat_col = (
        "Categoría_CV"
        if "Categoría_CV" in df_cv.columns
        else "Categoría"
    )

    categorias_presentes = (
        df_cv[cat_col].unique()
        if cat_col in df_cv.columns
        else []
    )

    for cat in CATEGORIAS:
        if cat in categorias_presentes:
            sub_df = df_cv[df_cv[cat_col] == cat]

            if sub_df.empty:
                continue

            # Encabezado de la Categoría
            p_cat = doc.add_paragraph()
            p_cat.paragraph_format.space_before = Pt(14)
            p_cat.paragraph_format.space_after = Pt(6)
            p_cat.paragraph_format.keep_with_next = True

            run_cat = p_cat.add_run(cat)
            run_cat.bold = True
            run_cat.font.size = Pt(12.5)
            run_cat.font.color.rgb = RGBColor(0, 51, 102)

            for _, row in sub_df.iterrows():
                titulo = limpiar_texto(
                    row.get("Título_Actividad_o_Publicación")
                    or row.get("Título")
                )

                rol = limpiar_texto(row.get("Rol_Participación"))

                inst = limpiar_texto(
                    row.get("Institución_Organización")
                    or row.get("Institución")
                )

                lugar = limpiar_texto(row.get("Lugar_Sede"))
                fecha_str = formatear_fecha_cv(row)

                if not titulo and not rol and not inst:
                    continue

                p_item = doc.add_paragraph(style="List Bullet")
                p_item.paragraph_format.space_after = Pt(4)
                p_item.paragraph_format.space_before = Pt(0)
                p_item.paragraph_format.line_spacing = 1.15

                if titulo:
                    run_t = p_item.add_run(titulo)
                    run_t.bold = True

                detalles = []

                if rol:
                    detalles.append(
                        rol
                        if rol.lower().startswith(
                            ("rol", "participación")
                        )
                        else f"Rol: {rol}"
                    )

                if inst:
                    detalles.append(inst)

                if lugar:
                    detalles.append(lugar)

                if fecha_str:
                    detalles.append(fecha_str)

                if detalles:
                    if titulo:
                        p_item.add_run(". ")

                    p_item.add_run(
                        ", ".join(detalles) + "."
                    )

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

    # ------------------------------------------------
    # CREAR / VERIFICAR ESTRUCTURA DE CARPETAS
    # ------------------------------------------------
    with st.spinner("🔧 Verificando estructura de carpetas en Google Drive..."):
        estructura_drive = inicializar_estructura_drive(service)

    if estructura_drive:
        st.success(
            "☁️ Estructura de carpetas de Google Drive lista."
        )

    else:
        st.error(
            "No fue posible crear o localizar la estructura de carpetas "
            "en Google Drive."
        )

    # ------------------------------------------------
    # BUSCAR EXCEL
    # ------------------------------------------------
    excel_id, found_name = buscar_excel_en_drive(service)

    if not excel_id:
        st.warning("⚠️ No se encontró la base de datos en Google Drive.")
        st.info(
            "Sube tu archivo Excel para vincularlo por primera vez:"
        )

        archivo_excel_nuevo = st.file_uploader(
            "Selecciona 'Base_de_Datos_Probatorios_y_CV.xlsx':",
            type=["xlsx"],
        )

        if archivo_excel_nuevo is not None:
            with st.spinner(
                "Subiendo Excel a Google Drive..."
            ):
                file_metadata = {
                    "name": "Base_de_Datos_Probatorios_y_CV.xlsx"
                }

                media = MediaIoBaseUpload(
                    io.BytesIO(
                        archivo_excel_nuevo.getvalue()
                    ),
                    mimetype=(
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    ),
                )

                service.files().create(
                    body=file_metadata,
                    media_body=media,
                    fields="id",
                ).execute()

                st.success(
                    "¡Base de datos vinculada con éxito!"
                )

                st.rerun()

    else:
        # Cargar DataFrame en vivo desde Drive
        df = cargar_datos_drive(service, excel_id)

        # 3 PESTAÑAS PRINCIPALES
        tab_consulta, tab_registro, tab_cv_word = st.tabs(
            [
                "🔍 Buscar y Consultar Probatorios",
                "➕ Registrar Nueva Actividad",
                "📄 Generar CV en Word",
            ]
        )

        # ------------------------------------------------
        # PESTAÑA 1: BUSCADOR Y CONSULTA
        # ------------------------------------------------
        with tab_consulta:
            st.subheader(
                "Buscador de Actividades y Documentos"
            )

            st.markdown("##### 🎯 Filtros de Búsqueda")

            col_f1, col_f2, col_f3 = st.columns(3)

            cat_col = (
                "Categoría_CV"
                if "Categoría_CV" in df.columns
                else "Categoría"
            )

            with col_f1:
                if "Año" in df.columns:
                    anios_disponibles = (
                        ["Todos"]
                        + sorted(
                            list(
                                df["Año"]
                                .dropna()
                                .astype(int)
                                .unique()
                            ),
                            reverse=True,
                        )
                    )
                else:
                    anios_disponibles = ["Todos"]

                filtro_anio = st.selectbox(
                    "Filtrar por Año",
                    anios_disponibles,
                )

            with col_f2:
                filtro_categoria = st.selectbox(
                    "Filtrar por Categoría",
                    ["Todas"] + CATEGORIAS,
                )

            with col_f3:
                filtro_texto = st.text_input(
                    "🔍 Buscar por palabra clave",
                    placeholder="Ej. Congreso, Comisión, Libro...",
                )

            df_filtrado = df.copy()

            if filtro_anio != "Todos":
                df_filtrado = df_filtrado[
                    df_filtrado["Año"] == int(filtro_anio)
                ]

            if (
                filtro_categoria != "Todas"
                and cat_col in df_filtrado.columns
            ):
                df_filtrado = df_filtrado[
                    df_filtrado[cat_col] == filtro_categoria
                ]

            if filtro_texto:
                mask = df_filtrado.apply(
                    lambda row: row.astype(str)
                    .str.contains(
                        filtro_texto,
                        case=False,
                        na=False,
                    )
                    .any(),
                    axis=1,
                )

                df_filtrado = df_filtrado[mask]

            st.markdown("---")

            st.write(
                f"Se encontraron **{len(df_filtrado)}** "
                "registros coincidentes:"
            )

            col_config = {}

            if "Enlace_Drive_Probatorio" in df_filtrado.columns:
                col_config[
                    "Enlace_Drive_Probatorio"
                ] = st.column_config.LinkColumn(
                    "Enlace Drive Probatorio"
                )

            elif "Enlace_Probatorio" in df_filtrado.columns:
                col_config[
                    "Enlace_Probatorio"
                ] = st.column_config.LinkColumn(
                    "Enlace Drive Probatorio"
                )

            st.dataframe(
                df_filtrado,
                use_container_width=True,
                column_config=col_config,
            )

        # ------------------------------------------------
        # PESTAÑA 2: CAPTURA DE NUEVAS ACTIVIDADES
        # ------------------------------------------------
        with tab_registro:
            st.subheader(
                "Formulario de Captura de Actividades y Constancias"
            )

            with st.form(
                "form_nueva_actividad",
                clear_on_submit=True,
            ):
                col1, col2 = st.columns(2)

                with col1:
                    nuevo_id = st.text_input(
                        "ID de Registro",
                        value=f"ACT-{len(df) + 1:03d}",
                    )

                    anio = st.selectbox(
                        "Año",
                        [
                            2026,
                            2025,
                            2024,
                            2023,
                            2022,
                            2021,
                            2020,
                        ],
                    )

                    fecha = st.date_input("Fecha")

                    categoria = st.selectbox(
                        "Categoría del CV",
                        CATEGORIAS,
                    )

                    rol = st.text_input(
                        "Rol / Participación "
                        "(ej. Ponente, Autora, Coordinadora)"
                    )

                with col2:
                    titulo = st.text_input(
                        "Título de la Actividad o Publicación *"
                    )

                    institucion = st.text_input(
                        "Institución u Organización"
                    )

                    lugar = st.text_input(
                        "Lugar / Sede"
                    )

                    estado = st.selectbox(
                        "Estado del Probatorio",
                        [
                            "Verificado / En Drive",
                            "Pendiente de Escanear",
                            "En Trámite",
                        ],
                    )

                    incluir = st.radio(
                        "¿Incluir en el CV?",
                        ["Sí", "No"],
                        horizontal=True,
                    )

                st.markdown("---")

                st.subheader(
                    "📎 Documento Probatorio"
                )

                archivo_pdf = st.file_uploader(
                    "Sube el documento probatorio",
                    type=[
                        "pdf",
                        "png",
                        "jpg",
                        "jpeg",
                    ],
                )

                notas = st.text_area(
                    "Notas / Observaciones de control interno"
                )

                boton_guardar = st.form_submit_button(
                    "💾 Guardar Registro y Subir Documento"
                )

            if boton_guardar:
                if not titulo:
                    st.error(
                        "⚠️ El campo "
                        "'Título de la Actividad' "
                        "es obligatorio."
                    )

                else:
                    with st.spinner(
                        "📁 Guardando probatorio en su carpeta "
                        "correspondiente y actualizando Excel..."
                    ):
                        nombre_archivo_guardado = "Sin_PDF"
                        enlace_drive = "Sin_Enlace"
                        archivo_drive_id = ""

                        if archivo_pdf is not None:
                            bytes_archivo = (
                                archivo_pdf.getvalue()
                            )

                            extension = os.path.splitext(
                                archivo_pdf.name
                            )[1].lower()

                            titulo_limpio = "".join(
                                x
                                for x in titulo
                                if x.isalnum()
                                or x in " _-"
                            ).strip()

                            titulo_limpio = (
                                titulo_limpio[:80]
                                or "Sin_Titulo"
                            )

                            nombre_archivo_guardado = (
                                f"{anio}_"
                                f"{categoria.replace(' ', '_')}_"
                                f"{titulo_limpio}"
                                f"{extension}"
                            )

                            carpeta_destino = (
                                obtener_carpeta_probatorio(
                                    service,
                                    estructura_drive,
                                    anio,
                                    categoria,
                                )
                            )

                            if not carpeta_destino:
                                st.error(
                                    "No se pudo determinar "
                                    "la carpeta de destino."
                                )
                                st.stop()

                            resultado = subir_a_google_drive(
                                service,
                                nombre_archivo_guardado,
                                bytes_archivo,
                                carpeta_destino,
                            )

                            if resultado[0]:
                                enlace_drive = resultado[0]
                                archivo_drive_id = (
                                    resultado[1] or ""
                                )

                        nueva_fila = {
                            "ID": nuevo_id,
                            "Año": anio,
                            "Fecha": str(fecha),
                            "Categoría_CV": categoria,
                            "Rol_Participación": rol,
                            "Título_Actividad_o_Publicación": titulo,
                            "Institución_Organización": institucion,
                            "Lugar_Sede": lugar,
                            "Nombre_Archivo_PDF": (
                                nombre_archivo_guardado
                            ),
                            "Enlace_Drive_Probatorio": enlace_drive,
                            "ID_Drive_Probatorio": archivo_drive_id,
                            "Estado_Probatorio": estado,
                            "Incluir_en_CV": incluir,
                            "Notas_Observaciones": notas,
                        }

                        # Mantener compatibilidad con columnas existentes
                        for col in df.columns:
                            if col not in nueva_fila:
                                nueva_fila[col] = ""

                        df_actualizado = pd.concat(
                            [
                                df,
                                pd.DataFrame([nueva_fila]),
                            ],
                            ignore_index=True,
                        )

                        actualizar_excel_drive(
                            service,
                            excel_id,
                            df_actualizado,
                        )

                        st.success(
                            f"¡Excelente! El registro "
                            f"'{titulo}' fue guardado."
                        )

                        if archivo_pdf is not None:
                            st.info(
                                f"📁 Ubicación: "
                                f"02 — Probatorios / "
                                f"{anio} / {categoria}"
                            )

                        st.balloons()
                        st.rerun()

        # ------------------------------------------------
        # PESTAÑA 3: GENERADOR DE CV EN WORD
        # ------------------------------------------------
        with tab_cv_word:
            st.subheader(
                "📄 Generador Automático de CV impreso (Word)"
            )

            st.write(
                "Este módulo compila automáticamente todas las "
                "actividades marcadas con **'¿Incluir en el CV? = Sí'**, "
                "formateadas en estilo académico profesional."
            )

            col_inc = (
                "Incluir_en_CV"
                if "Incluir_en_CV" in df.columns
                else df.columns[0]
            )

            df_cv_aprobados = df[
                df[col_inc]
                .astype(str)
                .str.strip()
                .str.lower()
                == "sí"
            ]

            col_w1, col_w2 = st.columns([2, 1])

            with col_w1:
                st.info(
                    f"📌 Actualmente hay **"
                    f"{len(df_cv_aprobados)} actividades** "
                    "listas para incluirse en el currículum."
                )

            with col_w2:
                if not df_cv_aprobados.empty:
                    archivo_word_bytes = crear_cv_word(df)

                    st.download_button(
                        label="📥 Descargar CV Actualizado (.docx)",
                        data=archivo_word_bytes,
                        file_name=(
                            "CV_Dra_Maria_Griselda_Gunther.docx"
                        ),
                        mime=(
                            "application/vnd.openxmlformats-officedocument."
                            "wordprocessingml.document"
                        ),
                    )

                else:
                    st.warning(
                        "No hay registros marcados con "
                        "'Incluir en CV = Sí'."
                    )

            st.markdown("---")

            st.markdown(
                "##### 👁️ Vista previa de los datos que se incluirán:"
            )

            columnas_preview = [
                c
                for c in [
                    "Año",
                    "Categoría_CV",
                    "Categoría",
                    "Título_Actividad_o_Publicación",
                    "Título",
                    "Rol_Participación",
                    "Institución_Organización",
                    "Institución",
                ]
                if c in df_cv_aprobados.columns
            ]

            st.dataframe(
                df_cv_aprobados[columnas_preview],
                use_container_width=True,
            )
