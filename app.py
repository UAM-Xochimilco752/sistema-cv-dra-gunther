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


# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================

st.set_page_config(
    page_title="Expediente Académico SNII - Dra. Günther",
    page_icon="🎓",
    layout="wide",
)

NOMBRE_CARPETA_RAIZ = "CV — Sistema de Gestión"

ANIOS_PROBATORIOS = [2023, 2024, 2025, 2026]

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

COMPONENTES_SNII = [
    "Componente 1 — Producción de investigación",
    "Componente 2 — Fortalecimiento y consolidación de la comunidad",
    "Componente 3 — Divulgación",
]

TIPOS_PRODUCTO = [
    "Artículo",
    "Libro",
    "Capítulo",
    "Reporte",
    "Informe",
    "Dossier o número temático",
    "Antología",
    "Traducción",
    "Prólogo o estudio introductorio",
    "Curaduría",
    "Datos primarios",
    "Software",
    "Producto tecnológico",
    "Proceso o estrategia digital",
    "Patente",
    "Derecho de autor",
    "Diseño industrial",
    "Modelo de utilidad",
    "Esquema de trazado",
    "Marca",
    "Licenciamiento",
    "Consultoría o asesoría técnica especializada",
    "Creación de empresa",
    "Asociación estratégica",
    "Informe técnico",
    "Curso",
    "Diplomado",
    "Capacitación",
    "Taller",
    "Seminario",
    "Tutoría",
    "Tesis",
    "Tesina",
    "Portafolio",
    "Proyecto de investigación",
    "Plan de estudios",
    "Acuerdo o convenio",
    "Coordinación de programa o centro",
    "Jurado",
    "Evaluación de programa o proyecto SECIHTI",
    "Dictaminación de publicación",
    "Dictaminación especializada",
    "Medios escritos",
    "Medios audiovisuales/radiofónicos/digitales",
    "Museografía / educación no formal",
    "Evento o comunicación",
    "Otro",
]

SUBTIPOS = [
    "Investigación científica/humanística",
    "Desarrollo tecnológico",
    "Innovación",
    "Docencia",
    "Trabajo de titulación",
    "Desarrollo institucional",
    "Evaluación académica",
    "Divulgación",
    "Otro",
]

MODALIDADES = [
    "No aplica",
    "Nacional",
    "Internacional",
    "Presencial",
    "Virtual",
    "Híbrida",
]

CARACTERISTICAS_SNII = [
    "Relevante",
    "Pertinente",
    "Sostenida",
    "Diversa",
    "Comprometida con la formación",
    "Colaborativa",
    "Constante",
    "Inclusiva",
    "Gratuita",
    "Comprensible",
]

COLUMNAS_NUEVAS = [
    "Componente_SNII",
    "Tipo_Producto_SNII",
    "Subtipo_SNII",
    "Categoria_CV",
    "Rol_Participacion",
    "Titulo_Actividad_o_Publicacion",
    "Evento_Revista_Libro",
    "Institucion_Organizacion",
    "Lugar_Sede",
    "Modalidad",
    "Autores",
    "Coautores",
    "Nivel_Formacion",
    "Estudiantes_Beneficiados",
    "Proyecto_Linea_Investigacion",
    "Descripcion_Aportacion",
    "Relevancia_Pertinencia",
    "Impacto_Beneficio_Social",
    "Caracteristicas_SNII",
    "Arbitrado",
    "Publicado",
    "Revista_Editorial",
    "Volumen_Numero",
    "Paginas",
    "ISBN_ISSN",
    "DOI_URL",
    "ID_Drive_Probatorio",
    "Incluir_en_CV_SNII",
    "Redaccion_CV",
]


# ============================================================
# GOOGLE DRIVE
# ============================================================

@st.cache_resource
def obtener_servicio_drive():

    creds = None

    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)

    if creds and creds.expired and creds.refresh_token:

        try:
            creds.refresh(Request())

        except Exception as e:
            st.error(
                f"Error al refrescar credenciales de Google: {e}"
            )
            return None

    if not creds:
        st.error(
            "⚠️ No se encontró el archivo 'token.pickle'."
        )
        return None

    return build(
        "drive",
        "v3",
        credentials=creds
    )


def buscar_carpeta(
    service,
    nombre,
    parent_id=None
):

    nombre_escapado = nombre.replace(
        "'",
        "\\'"
    )

    query = (
        f"name = '{nombre_escapado}' "
        "and mimeType = "
        "'application/vnd.google-apps.folder' "
        "and trashed = false"
    )

    if parent_id:
        query += f" and '{parent_id}' in parents"
    else:
        query += " and 'root' in parents"

    try:

        resultado = service.files().list(
            q=query,
            spaces="drive",
            fields="files(id,name,webViewLink)",
            pageSize=10
        ).execute()

        archivos = resultado.get(
            "files",
            []
        )

        return archivos[0] if archivos else None

    except Exception as e:

        st.error(
            f"Error buscando carpeta '{nombre}': {e}"
        )

        return None


def crear_carpeta(
    service,
    nombre,
    parent_id=None
):

    metadata = {
        "name": nombre,
        "mimeType":
            "application/vnd.google-apps.folder"
    }

    if parent_id:
        metadata["parents"] = [parent_id]

    try:

        return service.files().create(
            body=metadata,
            fields="id,name,webViewLink"
        ).execute()

    except Exception as e:

        st.error(
            f"Error creando carpeta '{nombre}': {e}"
        )

        return None


def obtener_o_crear_carpeta(
    service,
    nombre,
    parent_id=None
):

    carpeta = buscar_carpeta(
        service,
        nombre,
        parent_id
    )

    if carpeta:
        return carpeta

    return crear_carpeta(
        service,
        nombre,
        parent_id
    )


@st.cache_resource
def inicializar_estructura_drive(
    _service
):

    estructura = {}

    raiz = obtener_o_crear_carpeta(
        _service,
        NOMBRE_CARPETA_RAIZ
    )

    if not raiz:
        return {}

    estructura["raiz"] = raiz

    carpetas_principales = [
        "00 — Administración",
        "01 — Datos personales y CV",
        "02 — Probatorios",
        "10 — CV generados",
    ]

    for nombre in carpetas_principales:

        carpeta = obtener_o_crear_carpeta(
            _service,
            nombre,
            raiz["id"]
        )

        if carpeta:
            estructura[nombre] = carpeta

    probatorios = estructura.get(
        "02 — Probatorios"
    )

    if probatorios:

        for anio in ANIOS_PROBATORIOS:

            carpeta_anio = obtener_o_crear_carpeta(
                _service,
                str(anio),
                probatorios["id"]
            )

            if not carpeta_anio:
                continue

            estructura[
                f"probatorios_{anio}"
            ] = carpeta_anio

            for categoria in CATEGORIAS:

                carpeta_categoria = (
                    obtener_o_crear_carpeta(
                        _service,
                        categoria,
                        carpeta_anio["id"]
                    )
                )

                if carpeta_categoria:

                    estructura[
                        f"probatorios_{anio}_{categoria}"
                    ] = carpeta_categoria

    return estructura


def obtener_carpeta_probatorio(
    service,
    estructura,
    anio,
    categoria
):

    clave = (
        f"probatorios_{anio}_{categoria}"
    )

    if clave in estructura:
        return estructura[clave]

    probatorios = estructura.get(
        "02 — Probatorios"
    )

    if not probatorios:
        return None

    carpeta_anio = obtener_o_crear_carpeta(
        service,
        str(anio),
        probatorios["id"]
    )

    if not carpeta_anio:
        return None

    return obtener_o_crear_carpeta(
        service,
        categoria,
        carpeta_anio["id"]
    )


def obtener_mimetype(nombre):

    tipo, _ = mimetypes.guess_type(
        nombre
    )

    return tipo or "application/octet-stream"


def subir_a_google_drive(
    service,
    nombre_archivo,
    bytes_archivo,
    carpeta_destino
):

    metadata = {
        "name": nombre_archivo,
        "parents": [
            carpeta_destino["id"]
        ]
    }

    media = MediaIoBaseUpload(
        io.BytesIO(bytes_archivo),
        mimetype=obtener_mimetype(
            nombre_archivo
        ),
        resumable=True
    )

    try:

        archivo = service.files().create(
            body=metadata,
            media_body=media,
            fields=(
                "id,name,webViewLink,parents"
            )
        ).execute()

        # Importante:
        # NO hacemos público el archivo.

        return (
            archivo.get("webViewLink"),
            archivo.get("id")
        )

    except Exception as e:

        st.error(
            f"Error al subir archivo: {e}"
        )

        return None, None


# ============================================================
# EXCEL
# ============================================================

def buscar_excel_en_drive(service):

    try:

        resultado = service.files().list(
            q="trashed = false",
            fields="files(id,name)",
            pageSize=100
        ).execute()

        for archivo in resultado.get(
            "files",
            []
        ):

            if (
                "Base_de_Datos_Probatorios_y_CV"
                in archivo["name"]
            ):
                return (
                    archivo["id"],
                    archivo["name"]
                )

        return None, None

    except Exception as e:

        st.error(
            f"Error consultando Google Drive: {e}"
        )

        return None, None


def cargar_datos_drive(
    service,
    file_id
):

    request = service.files().get_media(
        fileId=file_id
    )

    buffer = io.BytesIO()

    downloader = MediaIoBaseDownload(
        buffer,
        request
    )

    done = False

    while not done:

        _, done = (
            downloader.next_chunk()
        )

    buffer.seek(0)

    df = pd.read_excel(
        buffer
    )

    # Las columnas antiguas se conservan.
    # Las nuevas se agregan automáticamente.

    for columna in COLUMNAS_NUEVAS:

        if columna not in df.columns:
            df[columna] = ""

    return df


def actualizar_excel_drive(
    service,
    file_id,
    df
):

    buffer = io.BytesIO()

    with pd.ExcelWriter(
        buffer,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False
        )

    buffer.seek(0)

    media = MediaIoBaseUpload(
        buffer,
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        resumable=True
    )

    service.files().update(
        fileId=file_id,
        media_body=media
    ).execute()


# ============================================================
# FUNCIONES DE TEXTO
# ============================================================

def limpiar_texto(valor):

    if pd.isna(valor):
        return ""

    texto = str(valor).strip()

    if texto.lower() in [
        "",
        "nan",
        "none",
        "ninguno",
        "ninguna",
        "n/a",
        "na",
        "-"
    ]:
        return ""

    return texto


def obtener_valor(
    row,
    columna
):

    if columna not in row.index:
        return ""

    return limpiar_texto(
        row.get(columna)
    )


def formatear_fecha(row):

    fecha = obtener_valor(
        row,
        "Fecha"
    )

    anio = obtener_valor(
        row,
        "Año"
    )

    if fecha:

        try:

            return pd.to_datetime(
                fecha
            ).strftime(
                "%d/%m/%Y"
            )

        except Exception:

            return fecha

    return anio


# ============================================================
# REDACCIÓN ACADÉMICA
# ============================================================

def redactar_registro(row):

    """
    Redacción estructurada.
    NO inventa información.
    """

    tipo = obtener_valor(
        row,
        "Tipo_Producto_SNII"
    )

    titulo = obtener_valor(
        row,
        "Titulo_Actividad_o_Publicacion"
    )

    autores = obtener_valor(
        row,
        "Autores"
    )

    rol = obtener_valor(
        row,
        "Rol_Participacion"
    )

    evento = obtener_valor(
        row,
        "Evento_Revista_Libro"
    )

    institucion = obtener_valor(
        row,
        "Institucion_Organizacion"
    )

    lugar = obtener_valor(
        row,
        "Lugar_Sede"
    )

    fecha = formatear_fecha(
        row
    )

    revista = obtener_valor(
        row,
        "Revista_Editorial"
    )

    volumen = obtener_valor(
        row,
        "Volumen_Numero"
    )

    paginas = obtener_valor(
        row,
        "Paginas"
    )

    isbn = obtener_valor(
        row,
        "ISBN_ISSN"
    )

    doi = obtener_valor(
        row,
        "DOI_URL"
    )

    arbitrado = obtener_valor(
        row,
        "Arbitrado"
    )

    if not titulo:
        return ""

    # ---------------------------------------------
    # ARTÍCULOS
    # ---------------------------------------------

    if tipo == "Artículo":

        texto = ""

        if autores:
            texto += (
                autores
                + ". "
            )

        texto += (
            f"({fecha}). "
            f"{titulo}."
        )

        if revista:
            texto += (
                f" {revista}."
            )

        if volumen:
            texto += (
                f" {volumen}."
            )

        if paginas:
            texto += (
                f" pp. {paginas}."
            )

        if arbitrado == "Sí":
            texto += (
                " Publicación arbitrada."
            )

        if isbn:
            texto += (
                f" ISSN/ISBN: {isbn}."
            )

        if doi:
            texto += (
                f" DOI/URL: {doi}."
            )

        return texto.strip()

    # ---------------------------------------------
    # LIBROS Y CAPÍTULOS
    # ---------------------------------------------

    if tipo in [
        "Libro",
        "Capítulo",
        "Antología",
        "Traducción",
        "Prólogo o estudio introductorio"
    ]:

        texto = ""

        if autores:
            texto += (
                autores
                + ". "
            )

        texto += (
            f"({fecha}). "
            f"{titulo}."
        )

        if revista:
            texto += (
                f" {revista}."
            )

        if paginas:
            texto += (
                f" pp. {paginas}."
            )

        if isbn:
            texto += (
                f" ISBN/ISSN: {isbn}."
            )

        if doi:
            texto += (
                f" DOI/URL: {doi}."
            )

        return texto.strip()

    # ---------------------------------------------
    # RESTO DE ACTIVIDADES
    # ---------------------------------------------

    texto = ""

    if autores:
        texto += (
            autores
            + ". "
        )

    texto += (
        titulo
        + "."
    )

    detalles = []

    if rol:
        detalles.append(
            f"Rol: {rol}"
        )

    if tipo:
        detalles.append(
            tipo
        )

    if evento:
        detalles.append(
            evento
        )

    if institucion:
        detalles.append(
            institucion
        )

    if lugar:
        detalles.append(
            lugar
        )

    if fecha:
        detalles.append(
            fecha
        )

    if detalles:

        texto += (
            " "
            + "; ".join(
                detalles
            )
            + "."
        )

    if isbn:
        texto += (
            f" ISBN/ISSN: {isbn}."
        )

    if doi:
        texto += (
            f" DOI/URL: {doi}."
        )

    return texto.strip()


# ============================================================
# GENERACIÓN DEL CV
# ============================================================

def crear_cv_word(df):

    doc = Document()

    # Márgenes
    for section in doc.sections:

        section.top_margin = Inches(
            0.8
        )

        section.bottom_margin = Inches(
            0.8
        )

        section.left_margin = Inches(
            0.9
        )

        section.right_margin = Inches(
            0.9
        )

    # Fuente
    normal = doc.styles[
        "Normal"
    ]

    normal.font.name = (
        "Calibri"
    )

    normal.font.size = Pt(
        10.5
    )

    # ---------------------------------------------
    # ENCABEZADO
    # ---------------------------------------------

    p = doc.add_paragraph()

    p.alignment = (
        WD_ALIGN_PARAGRAPH.CENTER
    )

    r = p.add_run(
        "DRA. MARÍA GRISELDA GÜNTHER"
    )

    r.bold = True
    r.font.size = Pt(17)

    r.font.color.rgb = (
        RGBColor(
            0,
            51,
            102
        )
    )

    p = doc.add_paragraph()

    p.alignment = (
        WD_ALIGN_PARAGRAPH.CENTER
    )

    r = p.add_run(
        "CURRÍCULUM VITAE ACADÉMICO"
    )

    r.bold = True
    r.font.size = Pt(12)

    p = doc.add_paragraph()

    p.alignment = (
        WD_ALIGN_PARAGRAPH.CENTER
    )

    r = p.add_run(
        "Documento de trabajo para evaluación y renovación SNII"
    )

    r.italic = True
    r.font.size = Pt(9.5)

    # ---------------------------------------------
    # COPIA DE TRABAJO
    # ---------------------------------------------

    trabajo = df.copy()

    trabajo["_anio_num"] = pd.to_numeric(
        trabajo["Año"],
        errors="coerce"
    )

    trabajo = trabajo.sort_values(
        [
            "_anio_num",
            "Fecha"
        ],
        ascending=[
            False,
            False
        ]
    )

    # Solo lo marcado para CV SNII.

    trabajo = trabajo[
        trabajo[
            "Incluir_en_CV_SNII"
        ]
        .astype(str)
        .str.strip()
        .str.lower()
        == "sí"
    ]

    # ---------------------------------------------
    # INTRODUCCIÓN
    # ---------------------------------------------

    p = doc.add_paragraph()

    r = p.add_run(
        "TRAYECTORIA ACADÉMICA Y CIENTÍFICA"
    )

    r.bold = True
    r.font.size = Pt(13)

    r.font.color.rgb = (
        RGBColor(
            0,
            51,
            102
        )
    )

    p = doc.add_paragraph()

    p.add_run(
        "El presente documento organiza la trayectoria "
        "académica registrada en la base de datos del "
        "expediente, agrupando las actividades de acuerdo "
        "con los componentes de evaluación del SNII."
    )

    # ---------------------------------------------
    # COMPONENTES
    # ---------------------------------------------

    componentes = [

        (
            "Componente 1 — Producción de investigación",
            "I. PRODUCCIÓN DE INVESTIGACIÓN CIENTÍFICA, "
            "HUMANÍSTICA Y TECNOLÓGICA"
        ),

        (
            "Componente 2 — Fortalecimiento y consolidación de la comunidad",
            "II. FORTALECIMIENTO Y CONSOLIDACIÓN "
            "DE LA COMUNIDAD"
        ),

        (
            "Componente 3 — Divulgación",
            "III. DIVULGACIÓN Y ACCESO UNIVERSAL "
            "AL CONOCIMIENTO"
        ),
    ]

    for componente, titulo_componente in componentes:

        registros = trabajo[
            trabajo[
                "Componente_SNII"
            ]
            == componente
        ]

        if registros.empty:
            continue

        p = doc.add_paragraph()

        p.paragraph_format.space_before = Pt(
            16
        )

        r = p.add_run(
            titulo_componente
        )

        r.bold = True
        r.font.size = Pt(13)

        r.font.color.rgb = (
            RGBColor(
                0,
                51,
                102
            )
        )

        # Agrupación por tipo

        tipos_presentes = (
            registros[
                "Tipo_Producto_SNII"
            ]
            .dropna()
            .unique()
        )

        for tipo in TIPOS_PRODUCTO:

            if tipo not in tipos_presentes:
                continue

            sub = registros[
                registros[
                    "Tipo_Producto_SNII"
                ]
                == tipo
            ]

            if sub.empty:
                continue

            p = doc.add_paragraph()

            p.paragraph_format.space_before = Pt(
                8
            )

            r = p.add_run(
                tipo.upper()
            )

            r.bold = True
            r.font.size = Pt(11)

            for _, row in sub.iterrows():

                texto = obtener_valor(
                    row,
                    "Redaccion_CV"
                )

                if not texto:
                    texto = redactar_registro(
                        row
                    )

                if not texto:
                    continue

                p = doc.add_paragraph(
                    style="List Bullet"
                )

                p.paragraph_format.space_after = Pt(
                    4
                )

                p.paragraph_format.line_spacing = (
                    1.08
                )

                p.add_run(
                    texto
                )

    # ---------------------------------------------
    # ANEXO DE CONTROL
    # ---------------------------------------------

    doc.add_page_break()

    p = doc.add_paragraph()

    r = p.add_run(
        "ANEXO — CONTROL DOCUMENTAL"
    )

    r.bold = True
    r.font.size = Pt(13)

    r.font.color.rgb = (
        RGBColor(
            0,
            51,
            102
        )
    )

    columnas = [
        "Año",
        "Componente_SNII",
        "Tipo_Producto_SNII",
        "Titulo_Actividad_o_Publicacion",
        "Estado_Probatorio"
    ]

    tabla = doc.add_table(
        rows=1,
        cols=len(columnas)
    )

    tabla.style = (
        "Table Grid"
    )

    for i, columna in enumerate(
        columnas
    ):

        tabla.rows[
            0
        ].cells[i].text = columna

    for _, row in trabajo.iterrows():

        cells = tabla.add_row().cells

        for i, columna in enumerate(
            columnas
        ):

            cells[i].text = obtener_valor(
                row,
                columna
            )

    buffer = io.BytesIO()

    doc.save(
        buffer
    )

    buffer.seek(0)

    return buffer


# ============================================================
# TABLERO SNII
# ============================================================

def mostrar_tablero_snii(df):

    st.subheader(
        "📊 Estado del expediente SNII"
    )

    total = len(df)

    if total == 0:

        st.info(
            "Todavía no existen actividades registradas."
        )

        return

    verificados = (
        df[
            "Estado_Probatorio"
        ]
        .astype(str)
        .str.lower()
        .str.contains(
            "verificado"
        )
        .sum()
    )

    pendientes = (
        total
        - verificados
    )

    produccion = (
        df[
            "Componente_SNII"
        ]
        == COMPONENTES_SNII[0]
    ).sum()

    comunidad = (
        df[
            "Componente_SNII"
        ]
        == COMPONENTES_SNII[1]
    ).sum()

    divulgacion = (
        df[
            "Componente_SNII"
        ]
        == COMPONENTES_SNII[2]
    ).sum()

    c1, c2, c3, c4 = st.columns(
        4
    )

    c1.metric(
        "Actividades",
        total
    )

    c2.metric(
        "Probatorios verificados",
        verificados
    )

    c3.metric(
        "Pendientes",
        pendientes
    )

    c4.metric(
        "Producción",
        produccion
    )

    st.markdown(
        "### Distribución por componente"
    )

    resumen = pd.DataFrame(
        {
            "Componente": [
                "Producción",
                "Comunidad",
                "Divulgación"
            ],
            "Registros": [
                produccion,
                comunidad,
                divulgacion
            ]
        }
    )

    st.bar_chart(
        resumen.set_index(
            "Componente"
        )
    )


# ============================================================
# APLICACIÓN PRINCIPAL
# ============================================================

st.title(
    "🎓 Sistema de Gestión de CV y Expediente SNII"
)

st.caption(
    "Dra. María Griselda Günther — "
    "repositorio documental, control académico y generación de CV"
)

service = obtener_servicio_drive()

if service:

    with st.spinner(
        "☁️ Verificando estructura de Google Drive..."
    ):

        estructura_drive = (
            inicializar_estructura_drive(
                service
            )
        )

    if not estructura_drive:

        st.error(
            "No fue posible inicializar "
            "la estructura de Google Drive."
        )

        st.stop()

    excel_id, found_name = (
        buscar_excel_en_drive(
            service
        )
    )

    if not excel_id:

        st.warning(
            "⚠️ No se encontró la base "
            "de datos en Google Drive."
        )

        archivo_excel_nuevo = (
            st.file_uploader(
                "Sube Base_de_Datos_Probatorios_y_CV.xlsx",
                type=["xlsx"]
            )
        )

        if archivo_excel_nuevo:

            with st.spinner(
                "Subiendo base de datos..."
            ):

                metadata = {
                    "name":
                    "Base_de_Datos_Probatorios_y_CV.xlsx"
                }

                media = MediaIoBaseUpload(
                    io.BytesIO(
                        archivo_excel_nuevo.getvalue()
                    ),
                    mimetype=(
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    )
                )

                service.files().create(
                    body=metadata,
                    media_body=media,
                    fields="id"
                ).execute()

                st.success(
                    "Base de datos vinculada correctamente."
                )

                st.rerun()

    else:

        df = cargar_datos_drive(
            service,
            excel_id
        )

        tab_dashboard, tab_consulta, tab_registro, tab_cv = st.tabs(
            [
                "📊 Expediente SNII",
                "🔍 Buscar",
                "➕ Nueva actividad",
                "📄 Generar CV SNII"
            ]
        )

        # ====================================================
        # DASHBOARD
        # ====================================================

        with tab_dashboard:

            mostrar_tablero_snii(
                df
            )

            st.markdown("---")

            st.info(
                "Este tablero es una herramienta de control "
                "interno. No sustituye la evaluación de las "
                "Comisiones del SNII ni calcula automáticamente "
                "la categoría o nivel."
            )

        # ====================================================
        # BUSCADOR
        # ====================================================

        with tab_consulta:

            st.subheader(
                "🔍 Buscador de actividades y probatorios"
            )

            c1, c2, c3 = st.columns(
                3
            )

            with c1:

                anios = ["Todos"]

                valores = pd.to_numeric(
                    df["Año"],
                    errors="coerce"
                ).dropna().astype(int).unique()

                anios += sorted(
                    valores.tolist(),
                    reverse=True
                )

                filtro_anio = st.selectbox(
                    "Año",
                    anios
                )

            with c2:

                filtro_componente = (
                    st.selectbox(
                        "Componente SNII",
                        ["Todos"]
                        + COMPONENTES_SNII
                    )
                )

            with c3:

                filtro_texto = st.text_input(
                    "Buscar",
                    placeholder=(
                        "Título, autor, revista, institución..."
                    )
                )

            resultado = df.copy()

            if filtro_anio != "Todos":

                resultado = resultado[
                    pd.to_numeric(
                        resultado["Año"],
                        errors="coerce"
                    )
                    == int(filtro_anio)
                ]

            if filtro_componente != "Todos":

                resultado = resultado[
                    resultado[
                        "Componente_SNII"
                    ]
                    == filtro_componente
                ]

            if filtro_texto:

                mascara = resultado.apply(
                    lambda row:
                    row.astype(str)
                    .str.contains(
                        filtro_texto,
                        case=False,
                        na=False
                    )
                    .any(),
                    axis=1
                )

                resultado = resultado[
                    mascara
                ]

            st.write(
                f"**{len(resultado)} registros encontrados.**"
            )

            columnas = [
                "ID",
                "Año",
                "Componente_SNII",
                "Tipo_Producto_SNII",
                "Titulo_Actividad_o_Publicacion",
                "Institucion_Organizacion",
                "Estado_Probatorio",
                "Enlace_Drive_Probatorio"
            ]

            columnas = [
                c for c in columnas
                if c in resultado.columns
            ]

            config = {}

            if (
                "Enlace_Drive_Probatorio"
                in resultado.columns
            ):

                config[
                    "Enlace_Drive_Probatorio"
                ] = st.column_config.LinkColumn(
                    "Probatorio en Drive"
                )

            st.dataframe(
                resultado[columnas],
                use_container_width=True,
                column_config=config
            )

        # ====================================================
        # NUEVA ACTIVIDAD
        # ====================================================

        with tab_registro:

            st.subheader(
                "➕ Registrar nueva actividad académica"
            )

            st.caption(
                "Aquí ya no registramos únicamente una constancia: "
                "registramos el producto académico, su función "
                "dentro de la trayectoria y su relación con el SNII."
            )

            with st.form(
                "form_nueva_actividad",
                clear_on_submit=True
            ):

                st.markdown(
                    "### 1. Clasificación SNII"
                )

                c1, c2, c3 = st.columns(
                    3
                )

                with c1:

                    nuevo_id = st.text_input(
                        "ID",
                        value=(
                            f"ACT-{len(df)+1:04d}"
                        )
                    )

                    anio = st.selectbox(
                        "Año",
                        list(
                            range(
                                2026,
                                2019,
                                -1
                            )
                        )
                    )

                    componente = st.selectbox(
                        "Componente SNII",
                        COMPONENTES_SNII
                    )

                with c2:

                    tipo_producto = st.selectbox(
                        "Tipo de producto / actividad",
                        TIPOS_PRODUCTO
                    )

                    subtipo = st.selectbox(
                        "Subtipo",
                        SUBTIPOS
                    )

                    modalidad = st.selectbox(
                        "Modalidad",
                        MODALIDADES
                    )

                with c3:

                    categoria = st.selectbox(
                        "Categoría documental",
                        CATEGORIAS
                    )

                    rol = st.text_input(
                        "Rol / participación"
                    )

                    fecha = st.date_input(
                        "Fecha"
                    )

                st.markdown(
                    "### 2. Información académica"
                )

                titulo = st.text_input(
                    "Título de la actividad o publicación *"
                )

                evento = st.text_input(
                    "Evento / Revista / Libro"
                )

                institucion = st.text_input(
                    "Institución / organización"
                )

                lugar = st.text_input(
                    "Lugar / sede"
                )

                autores = st.text_input(
                    "Autores"
                )

                coautores = st.text_input(
                    "Coautores"
                )

                nivel_formacion = st.text_input(
                    "Nivel de formación relacionado"
                )

                estudiantes = st.text_input(
                    "Estudiantes beneficiados / dirigidos"
                )

                proyecto = st.text_input(
                    "Proyecto / línea de investigación"
                )

                st.markdown(
                    "### 3. Caracterización cualitativa"
                )

                aportacion = st.text_area(
                    "Descripción de la aportación"
                )

                relevancia = st.text_area(
                    "Relevancia / pertinencia"
                )

                impacto = st.text_area(
                    "Impacto / beneficio social"
                )

                caracteristicas = st.multiselect(
                    "Características SNII",
                    CARACTERISTICAS_SNII
                )

                st.markdown(
                    "### 4. Información bibliográfica"
                )

                c1, c2, c3 = st.columns(
                    3
                )

                with c1:

                    arbitrado = st.selectbox(
                        "¿Arbitrado / pares?",
                        [
                            "No aplica",
                            "Sí",
                            "No",
                            "No especificado"
                        ]
                    )

                    publicado = st.selectbox(
                        "Estado",
                        [
                            "No aplica",
                            "Publicado",
                            "Aceptado",
                            "En prensa",
                            "No publicado"
                        ]
                    )

                with c2:

                    revista = st.text_input(
                        "Revista / editorial"
                    )

                    volumen = st.text_input(
                        "Volumen / número"
                    )

                    paginas = st.text_input(
                        "Páginas"
                    )

                with c3:

                    isbn = st.text_input(
                        "ISBN / ISSN"
                    )

                    doi = st.text_input(
                        "DOI / URL"
                    )

                st.markdown(
                    "### 5. Evidencia documental"
                )

                archivo = st.file_uploader(
                    "Sube el probatorio",
                    type=[
                        "pdf",
                        "png",
                        "jpg",
                        "jpeg"
                    ]
                )

                estado = st.selectbox(
                    "Estado del probatorio",
                    [
                        "Verificado / En Drive",
                        "Pendiente de verificar",
                        "Pendiente de escanear",
                        "En trámite"
                    ]
                )

                c1, c2 = st.columns(
                    2
                )

                with c1:

                    incluir_cv = st.radio(
                        "¿Incluir en CV general?",
                        ["Sí", "No"],
                        horizontal=True
                    )

                with c2:

                    incluir_snii = st.radio(
                        "¿Incluir en CV SNII?",
                        ["Sí", "No"],
                        horizontal=True
                    )

                notas = st.text_area(
                    "Notas / observaciones"
                )

                guardar = st.form_submit_button(
                    "💾 Guardar actividad y probatorio"
                )

            if guardar:

                if not titulo.strip():

                    st.error(
                        "⚠️ El título es obligatorio."
                    )

                    st.stop()

                with st.spinner(
                    "📚 Registrando actividad y organizando expediente..."
                ):

                    nombre_archivo = "Sin_PDF"
                    enlace = "Sin_Enlace"
                    drive_id = ""

                    if archivo:

                        extension = (
                            os.path.splitext(
                                archivo.name
                            )[1].lower()
                        )

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

                        nombre_archivo = (
                            f"{anio}_"
                            f"{categoria.replace(' ', '_')}_"
                            f"{titulo_limpio}"
                            f"{extension}"
                        )

                        carpeta = (
                            obtener_carpeta_probatorio(
                                service,
                                estructura_drive,
                                anio,
                                categoria
                            )
                        )

                        if not carpeta:

                            st.error(
                                "No se encontró la carpeta "
                                "de destino."
                            )

                            st.stop()

                        enlace, drive_id = (
                            subir_a_google_drive(
                                service,
                                nombre_archivo,
                                archivo.getvalue(),
                                carpeta
                            )
                        )

                        if not enlace:
                            enlace = "Sin_Enlace"

                    nueva_fila = {

                        "ID":
                            nuevo_id,

                        "Año":
                            anio,

                        "Fecha":
                            str(fecha),

                        "Componente_SNII":
                            componente,

                        "Tipo_Producto_SNII":
                            tipo_producto,

                        "Subtipo_SNII":
                            subtipo,

                        "Categoria_CV":
                            categoria,

                        "Rol_Participacion":
                            rol,

                        "Titulo_Actividad_o_Publicacion":
                            titulo,

                        "Evento_Revista_Libro":
                            evento,

                        "Institucion_Organizacion":
                            institucion,

                        "Lugar_Sede":
                            lugar,

                        "Modalidad":
                            modalidad,

                        "Autores":
                            autores,

                        "Coautores":
                            coautores,

                        "Nivel_Formacion":
                            nivel_formacion,

                        "Estudiantes_Beneficiados":
                            estudiantes,

                        "Proyecto_Linea_Investigacion":
                            proyecto,

                        "Descripcion_Aportacion":
                            aportacion,

                        "Relevancia_Pertinencia":
                            relevancia,

                        "Impacto_Beneficio_Social":
                            impacto,

                        "Caracteristicas_SNII":
                            "; ".join(
                                caracteristicas
                            ),

                        "Arbitrado":
                            arbitrado,

                        "Publicado":
                            publicado,

                        "Revista_Editorial":
                            revista,

                        "Volumen_Numero":
                            volumen,

                        "Paginas":
                            paginas,

                        "ISBN_ISSN":
                            isbn,

                        "DOI_URL":
                            doi,

                        "Nombre_Archivo_PDF":
                            nombre_archivo,

                        "Enlace_Drive_Probatorio":
                            enlace,

                        "ID_Drive_Probatorio":
                            drive_id,

                        "Estado_Probatorio":
                            estado,

                        "Incluir_en_CV":
                            incluir_cv,

                        "Incluir_en_CV_SNII":
                            incluir_snii,

                        "Notas_Observaciones":
                            notas,
                    }

                    nueva_fila[
                        "Redaccion_CV"
                    ] = redactar_registro(
                        pd.Series(
                            nueva_fila
                        )
                    )

                    # Compatibilidad con Excel anterior.

                    for columna in df.columns:

                        if columna not in nueva_fila:

                            nueva_fila[
                                columna
                            ] = ""

                    df_actualizado = pd.concat(
                        [
                            df,
                            pd.DataFrame(
                                [nueva_fila]
                            )
                        ],
                        ignore_index=True
                    )

                    actualizar_excel_drive(
                        service,
                        excel_id,
                        df_actualizado
                    )

                st.success(
                    f"Actividad '{titulo}' "
                    "registrada correctamente."
                )

                st.info(
                    "📁 Probatorio guardado en: "
                    f"02 — Probatorios / "
                    f"{anio} / {categoria}"
                )

                st.balloons()

                st.rerun()

        # ====================================================
        # GENERADOR DE CV
        # ====================================================

        with tab_cv:

            st.subheader(
                "📄 Generador de CV académico SNII"
            )

            seleccionados = df[
                df[
                    "Incluir_en_CV_SNII"
                ]
                .astype(str)
                .str.strip()
                .str.lower()
                == "sí"
            ]

            verificados = seleccionados[
                seleccionados[
                    "Estado_Probatorio"
                ]
                .astype(str)
                .str.lower()
                .str.contains(
                    "verificado"
                )
            ]

            c1, c2, c3 = st.columns(
                3
            )

            c1.metric(
                "Seleccionados",
                len(seleccionados)
            )

            c2.metric(
                "Probatorios verificados",
                len(verificados)
            )

            c3.metric(
                "Pendientes",
                max(
                    0,
                    len(seleccionados)
                    - len(verificados)
                )
            )

            if (
                len(seleccionados)
                > len(verificados)
            ):

                st.warning(
                    "⚠️ Hay actividades seleccionadas "
                    "para el CV SNII cuyo probatorio "
                    "todavía no aparece como verificado."
                )

            if seleccionados.empty:

                st.info(
                    "Marca actividades como "
                    "'Incluir en CV SNII = Sí' "
                    "para generar el documento."
                )

            else:

                archivo_word = crear_cv_word(
                    df
                )

                st.download_button(
                    label=(
                        "📥 Descargar CV académico SNII (.docx)"
                    ),
                    data=archivo_word,
                    file_name=(
                        "CV_Academico_SNII_"
                        "Dra_Maria_Griselda_Gunther.docx"
                    ),
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "wordprocessingml.document"
                    )
                )

                st.markdown("---")

                st.markdown(
                    "### Vista previa"
                )

                columnas_preview = [
                    c
                    for c in [
                        "Año",
                        "Componente_SNII",
                        "Tipo_Producto_SNII",
                        "Titulo_Actividad_o_Publicacion",
                        "Estado_Probatorio"
                    ]
                    if c in seleccionados.columns
                ]

                st.dataframe(
                    seleccionados[
                        columnas_preview
                    ],
                    use_container_width=True
                )
