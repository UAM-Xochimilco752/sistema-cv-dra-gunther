import io
import os
import re
import pickle
import mimetypes
import zipfile
from datetime import datetime

import pandas as pd
import streamlit as st

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor

from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload


# ============================================================
# CONFIGURACIÓN
# ============================================================

st.set_page_config(
    page_title="Control de CV y Probatorios - Dra. Günther",
    page_icon="📄",
    layout="wide",
)


CATEGORIAS = [
    "Coordinación de Libros",
    "Capítulos de Libros / Artículos",
    "Ponencias y Conferencias",
    "Dirección de Tesis y Sinodalías",
    "Proyectos de Investigación",
    "Organización de Eventos Académicos",
    "Arbitraje y Trabajo Editorial (Dictaminación)",
    "Gestión Académico-Administrativa",
    "Difusión y Divulgación",
    "Diseño de Planes y Programas de Estudio",
    "Presentaciones de Libros",
    "Comisiones",
    "Cursos e Impartición de Clases",
    "Premios y Reconocimientos",
    "Asesorías",
]


NOMBRE_CARPETA_RAIZ = "CV — Sistema de Gestión"


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


ANIOS_PROBATORIOS = list(range(2020, datetime.now().year + 2))


# ============================================================
# ESQUEMA LIMPIO DE LA BASE DE DATOS
# ============================================================

COLUMNAS_BASE = [
    "ID",
    "Año",
    "Fecha",
    "Categoría",
    "Rol",
    "Título",
    "Institución",
    "Lugar",
    "Estado_Probatorio",
    "Incluir_en_CV",
    "Notas_Observaciones",
    "Nombre_Archivo_PDF",
    "Enlace_Drive_Probatorio",
    "ID_Drive_Probatorio",
]


# Columnas antiguas que NO queremos conservar.
COLUMNAS_OBSOLETAS = [
    "Componente_SNII",
    "Tipo_Producto_SNII",
    "Categoría_CV",
    "Rol_Participación",
    "Título_Actividad_o_Publicación",
    "Evento_Revista_Libro",
    "Institución_Organización",
    "Modalidad",
    "Autores",
    "Coautores",
    "Nivel_Formación",
    "Estudiantes_Beneficiados",
    "Proyecto_Línea_Investigación",
    "Descripción_Aportación",
    "Impacto_Beneficio_Social",
    "Características_SNII",
    "Arbitrado",
    "Publicado",
    "Revista_Editorial",
    "Volumen_Número",
    "Páginas",
    "ISBN_ISSN",
    "DOI_URL",
    "Incluir_en_CV_SNII",
    "Redacción",
]


# Separador utilizado para múltiples archivos dentro de una celda.
SEPARADOR_ARCHIVOS = " || "


# ============================================================
# GOOGLE DRIVE
# ============================================================

@st.cache_resource
def obtener_servicio_drive():
    """Obtiene el servicio autenticado de Google Drive."""

    creds = None

    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)

    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
        except Exception as e:
            st.error(
                f"Error al refrescar las credenciales de Google: {e}"
            )
            return None

    if not creds:
        st.error(
            "⚠️ No se encontró el archivo 'token.pickle'."
        )
        return None

    try:
        return build(
            "drive",
            "v3",
            credentials=creds,
        )
    except Exception as e:
        st.error(
            f"Error al crear el servicio de Google Drive: {e}"
        )
        return None


# ============================================================
# CARPETAS
# ============================================================

def buscar_carpeta(service, nombre, parent_id=None):

    try:

        nombre_escapado = nombre.replace("'", "\\'")

        if parent_id:
            query = (
                f"name = '{nombre_escapado}' "
                "and mimeType = 'application/vnd.google-apps.folder' "
                "and trashed = false "
                f"and '{parent_id}' in parents"
            )
        else:
            query = (
                f"name = '{nombre_escapado}' "
                "and mimeType = 'application/vnd.google-apps.folder' "
                "and trashed = false "
                "and 'root' in parents"
            )

        resultado = service.files().list(
            q=query,
            spaces="drive",
            fields="files(id,name,webViewLink)",
            pageSize=10,
        ).execute()

        carpetas = resultado.get("files", [])

        return carpetas[0] if carpetas else None

    except Exception as e:

        st.error(
            f"Error al buscar la carpeta '{nombre}': {e}"
        )

        return None


def crear_carpeta(service, nombre, parent_id=None):

    try:

        metadata = {
            "name": nombre,
            "mimeType": "application/vnd.google-apps.folder",
        }

        if parent_id:
            metadata["parents"] = [parent_id]

        return service.files().create(
            body=metadata,
            fields="id,name,webViewLink",
        ).execute()

    except Exception as e:

        st.error(
            f"Error al crear la carpeta '{nombre}': {e}"
        )

        return None


def obtener_o_crear_carpeta(
    service,
    nombre,
    parent_id=None,
):

    carpeta = buscar_carpeta(
        service,
        nombre,
        parent_id,
    )

    if carpeta:
        return carpeta

    return crear_carpeta(
        service,
        nombre,
        parent_id,
    )


@st.cache_resource
def inicializar_estructura_drive(_service):

    estructura = {}

    raiz = obtener_o_crear_carpeta(
        _service,
        NOMBRE_CARPETA_RAIZ,
    )

    if not raiz:
        return {}

    estructura["raiz"] = raiz

    for nombre in ESTRUCTURA_CARPETAS:

        carpeta = obtener_o_crear_carpeta(
            _service,
            nombre,
            raiz["id"],
        )

        if carpeta:
            estructura[nombre] = carpeta

    carpeta_probatorios = estructura.get(
        "02 — Probatorios"
    )

    if not carpeta_probatorios:
        return estructura

    estructura["probatorios"] = carpeta_probatorios

    for anio in ANIOS_PROBATORIOS:

        carpeta_anio = obtener_o_crear_carpeta(
            _service,
            str(anio),
            carpeta_probatorios["id"],
        )

        if not carpeta_anio:
            continue

        estructura[f"probatorios_{anio}"] = carpeta_anio

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


def obtener_carpeta_probatorio(
    service,
    estructura,
    anio,
    categoria,
):

    clave = f"probatorios_{anio}_{categoria}"

    if clave in estructura:
        return estructura[clave]

    carpeta_probatorios = estructura.get(
        "probatorios"
    )

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


# ============================================================
# ARCHIVOS DRIVE
# ============================================================

def obtener_mimetype(nombre_archivo):

    mimetype, _ = mimetypes.guess_type(
        nombre_archivo
    )

    return (
        mimetype
        or "application/octet-stream"
    )


def subir_a_google_drive(
    service,
    nombre_archivo,
    bytes_archivo,
    carpeta_destino,
):

    try:

        metadata = {
            "name": nombre_archivo,
            "parents": [
                carpeta_destino["id"]
            ],
        }

        media = MediaIoBaseUpload(
            io.BytesIO(bytes_archivo),
            mimetype=obtener_mimetype(
                nombre_archivo
            ),
            resumable=True,
        )

        archivo = service.files().create(
            body=metadata,
            media_body=media,
            fields="id,name,webViewLink,parents",
        ).execute()

        return (
            archivo.get("name", nombre_archivo),
            archivo.get("webViewLink", ""),
            archivo.get("id", ""),
        )

    except Exception as e:

        st.error(
            f"Error al subir '{nombre_archivo}': {e}"
        )

        return None, None, None


def obtener_archivo_drive(
    service,
    archivo_id,
):

    try:

        metadata = service.files().get(
            fileId=archivo_id,
            fields="id,name,mimeType",
        ).execute()

        request = service.files().get_media(
            fileId=archivo_id
        )

        buffer = io.BytesIO()

        downloader = MediaIoBaseDownload(
            buffer,
            request,
        )

        done = False

        while not done:
            _, done = downloader.next_chunk()

        buffer.seek(0)

        return (
            metadata,
            buffer.getvalue(),
        )

    except Exception as e:

        st.error(
            f"No se pudo descargar el archivo "
            f"{archivo_id}: {e}"
        )

        return None, None


def eliminar_archivo_drive(
    service,
    archivo_id,
):

    if not archivo_id:
        return True

    try:

        service.files().update(
            fileId=str(archivo_id).strip(),
            body={"trashed": True},
        ).execute()

        return True

    except Exception as e:

        st.error(
            f"Error al enviar archivo a papelera: {e}"
        )

        return False


def mover_archivo_drive(
    service,
    archivo_id,
    carpeta_destino,
):

    if not archivo_id or not carpeta_destino:
        return False

    try:

        archivo = service.files().get(
            fileId=archivo_id,
            fields="parents",
        ).execute()

        padres_actuales = archivo.get(
            "parents",
            [],
        )

        service.files().update(
            fileId=archivo_id,
            addParents=carpeta_destino["id"],
            removeParents=",".join(
                padres_actuales
            ) if padres_actuales else None,
            fields="id,parents",
        ).execute()

        return True

    except Exception as e:

        st.error(
            f"Error al mover archivo: {e}"
        )

        return False


def renombrar_archivo_drive(
    service,
    archivo_id,
    nuevo_nombre,
):

    if not archivo_id:
        return False

    try:

        service.files().update(
            fileId=archivo_id,
            body={"name": nuevo_nombre},
        ).execute()

        return True

    except Exception as e:

        st.error(
            f"Error al renombrar archivo: {e}"
        )

        return False


# ============================================================
# MANEJO DE MÚLTIPLES PROBATORIOS
# ============================================================

def convertir_celda_a_lista(valor):

    if valor is None:
        return []

    if pd.isna(valor):
        return []

    texto = str(valor).strip()

    if not texto:
        return []

    valores = [
        x.strip()
        for x in texto.split(
            SEPARADOR_ARCHIVOS
        )
    ]

    return [
        x for x in valores
        if x and x.lower() not in [
            "nan",
            "none",
            "sin_pdf",
            "sin_enlace",
        ]
    ]


def lista_a_celda(lista):

    if not lista:
        return ""

    return SEPARADOR_ARCHIVOS.join(
        str(x).strip()
        for x in lista
        if str(x).strip()
    )


def obtener_probatorios_de_fila(
    fila,
):

    nombres = convertir_celda_a_lista(
        fila.get(
            "Nombre_Archivo_PDF",
            "",
        )
    )

    enlaces = convertir_celda_a_lista(
        fila.get(
            "Enlace_Drive_Probatorio",
            "",
        )
    )

    ids = convertir_celda_a_lista(
        fila.get(
            "ID_Drive_Probatorio",
            "",
        )
    )

    probatorios = []

    cantidad = max(
        len(nombres),
        len(enlaces),
        len(ids),
    )

    for i in range(cantidad):

        probatorios.append(
            {
                "nombre": (
                    nombres[i]
                    if i < len(nombres)
                    else ""
                ),
                "enlace": (
                    enlaces[i]
                    if i < len(enlaces)
                    else ""
                ),
                "id": (
                    ids[i]
                    if i < len(ids)
                    else ""
                ),
            }
        )

    return probatorios


def actualizar_probatorios_fila(
    fila,
    probatorios,
):

    fila["Nombre_Archivo_PDF"] = lista_a_celda(
        [p["nombre"] for p in probatorios]
    )

    fila["Enlace_Drive_Probatorio"] = lista_a_celda(
        [p["enlace"] for p in probatorios]
    )

    fila["ID_Drive_Probatorio"] = lista_a_celda(
        [p["id"] for p in probatorios]
    )

    return fila


# ============================================================
# NOMBRES
# ============================================================

def limpiar_nombre_archivo(texto):

    texto = str(texto)

    texto = re.sub(
        r'[\\/:*?"<>|]+',
        "_",
        texto,
    )

    texto = re.sub(
        r"\s+",
        " ",
        texto,
    ).strip()

    return texto


def nombre_probatorio(
    titulo,
    anio,
    categoria,
    numero,
    extension,
):

    titulo_limpio = limpiar_nombre_archivo(
        titulo
    )

    titulo_limpio = titulo_limpio[:100]

    categoria_limpia = limpiar_nombre_archivo(
        categoria
    ).replace(
        " ",
        "_",
    )

    return (
        f"{anio}_"
        f"{categoria_limpia}_"
        f"{titulo_limpio}_"
        f"Probatorio_{numero}"
        f"{extension.lower()}"
    )


# ============================================================
# EXCEL
# ============================================================

def buscar_excel_en_drive(service):

    try:

        resultado = service.files().list(
            q=(
                "name contains "
                "'Base_de_Datos_Probatorios_y_CV' "
                "and trashed = false"
            ),
            fields="files(id,name)",
            pageSize=100,
        ).execute()

        archivos = resultado.get(
            "files",
            [],
        )

        if not archivos:
            return None, None

        archivo = archivos[0]

        return (
            archivo["id"],
            archivo["name"],
        )

    except Exception as e:

        st.error(
            f"Error al buscar Excel: {e}"
        )

        return None, None


def cargar_datos_drive(
    service,
    file_id,
):

    request = service.files().get_media(
        fileId=file_id
    )

    buffer = io.BytesIO()

    downloader = MediaIoBaseDownload(
        buffer,
        request,
    )

    done = False

    while not done:
        _, done = downloader.next_chunk()

    buffer.seek(0)

    return pd.read_excel(buffer)


def actualizar_excel_drive(
    service,
    file_id,
    df,
):

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl",
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="Base de datos",
        )

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


# ============================================================
# MIGRACIÓN / LIMPIEZA DE BASE EXISTENTE
# ============================================================

def preparar_dataframe(df):

    df = df.copy()

    # --------------------------------------------------------
    # Migrar nombres antiguos a nombres nuevos
    # --------------------------------------------------------

    equivalencias = {

        "Categoría_CV": "Categoría",

        "Rol_Participación": "Rol",

        "Título_Actividad_o_Publicación": "Título",

        "Institución_Organización": "Institución",

    }

    for antigua, nueva in equivalencias.items():

        if (
            antigua in df.columns
            and nueva not in df.columns
        ):

            df[nueva] = df[antigua]

    # --------------------------------------------------------
    # Crear columnas nuevas que falten
    # --------------------------------------------------------

    for columna in COLUMNAS_BASE:

        if columna not in df.columns:
            df[columna] = ""

    # --------------------------------------------------------
    # Eliminar columnas basura
    # --------------------------------------------------------

    columnas_a_eliminar = [
        c
        for c in df.columns
        if c in COLUMNAS_OBSOLETAS
    ]

    if columnas_a_eliminar:

        df = df.drop(
            columns=columnas_a_eliminar,
            errors="ignore",
        )

    # --------------------------------------------------------
    # Asegurar exactamente nuestro esquema
    # --------------------------------------------------------

    df = df[
        [
            c
            for c in COLUMNAS_BASE
            if c in df.columns
        ]
    ]

    return df


# ============================================================
# UTILIDADES
# ============================================================

def valor_fila(
    fila,
    columna,
    default="",
):

    if columna not in fila.index:
        return default

    valor = fila[columna]

    if pd.isna(valor):
        return default

    return valor


def siguiente_id_registro(df):

    numeros = []

    if "ID" in df.columns:

        for valor in df["ID"].dropna():

            match = re.fullmatch(
                r"ACT-(\d+)",
                str(valor).strip(),
                re.IGNORECASE,
            )

            if match:
                numeros.append(
                    int(match.group(1))
                )

    siguiente = max(
        numeros,
        default=0,
    ) + 1

    return f"ACT-{siguiente:03d}"


def formatear_fecha(
    valor,
):

    try:

        return pd.to_datetime(
            valor
        ).strftime(
            "%d/%m/%Y"
        )

    except Exception:

        return str(valor)


# ============================================================
# CREACIÓN DE ZIP
# ============================================================

def crear_zip_probatorios(
    service,
    filas,
):

    buffer_zip = io.BytesIO()

    archivos_agregados = 0

    nombres_zip = set()

    with zipfile.ZipFile(
        buffer_zip,
        "w",
        zipfile.ZIP_DEFLATED,
    ) as zip_file:

        for _, fila in filas.iterrows():

            registro_id = str(
                valor_fila(
                    fila,
                    "ID",
                    "SIN_ID",
                )
            )

            titulo = limpiar_nombre_archivo(
                valor_fila(
                    fila,
                    "Título",
                    "Sin título",
                )
            )

            probatorios = (
                obtener_probatorios_de_fila(
                    fila
                )
            )

            for numero, probatorio in enumerate(
                probatorios,
                start=1,
            ):

                archivo_id = probatorio["id"]

                if not archivo_id:
                    continue

                metadata, contenido = (
                    obtener_archivo_drive(
                        service,
                        archivo_id,
                    )
                )

                if not contenido:
                    continue

                nombre_original = (
                    metadata.get(
                        "name",
                        f"probatorio_{numero}",
                    )
                    if metadata
                    else f"probatorio_{numero}"
                )

                nombre_original = (
                    limpiar_nombre_archivo(
                        nombre_original
                    )
                )

                nombre_zip = (
                    f"{registro_id}/"
                    f"{numero:02d}_"
                    f"{nombre_original}"
                )

                # Evitar duplicados accidentales
                contador = 2

                nombre_base = nombre_zip

                while nombre_zip in nombres_zip:

                    raiz, ext = os.path.splitext(
                        nombre_base
                    )

                    nombre_zip = (
                        f"{raiz}_{contador}{ext}"
                    )

                    contador += 1

                nombres_zip.add(
                    nombre_zip
                )

                zip_file.writestr(
                    nombre_zip,
                    contenido,
                )

                archivos_agregados += 1

    buffer_zip.seek(0)

    return (
        buffer_zip.getvalue(),
        archivos_agregados,
    )


# ============================================================
# WORD
# ============================================================

def crear_cv_word(df):

    doc = Document()

    for section in doc.sections:

        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    style = doc.styles["Normal"]

    style.font.name = "Calibri"
    style.font.size = Pt(11)

    p = doc.add_paragraph()

    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = p.add_run(
        "DRA. MARÍA GRISELDA GÜNTHER"
    )

    run.bold = True
    run.font.size = Pt(16)
    run.font.color.rgb = RGBColor(
        0,
        51,
        102,
    )

    p2 = doc.add_paragraph()

    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run2 = p2.add_run(
        "CURRÍCULUM VITAE — SÍNTESIS EJECUTIVA"
    )

    run2.font.size = Pt(10.5)
    run2.font.italic = True

    df_cv = df[
        df["Incluir_en_CV"]
        .astype(str)
        .str.strip()
        .str.lower()
        == "sí"
    ].copy()

    if df_cv.empty:

        doc.add_paragraph(
            "No existen actividades marcadas "
            "para incluir en el CV."
        )

    else:

        df_cv["Año_num"] = pd.to_numeric(
            df_cv["Año"],
            errors="coerce",
        )

        df_cv = df_cv.sort_values(
            "Año_num",
            ascending=False,
        )

        categorias_presentes = (
            df_cv["Categoría"]
            .dropna()
            .unique()
        )

        for categoria in CATEGORIAS:

            if categoria not in categorias_presentes:
                continue

            sub_df = df_cv[
                df_cv["Categoría"]
                == categoria
            ]

            if sub_df.empty:
                continue

            p_cat = doc.add_paragraph()

            p_cat.paragraph_format.space_before = Pt(14)
            p_cat.paragraph_format.space_after = Pt(6)

            run_cat = p_cat.add_run(
                categoria
            )

            run_cat.bold = True
            run_cat.font.size = Pt(12.5)
            run_cat.font.color.rgb = RGBColor(
                0,
                51,
                102,
            )

            for _, fila in sub_df.iterrows():

                titulo = str(
                    valor_fila(
                        fila,
                        "Título",
                        "",
                    )
                ).strip()

                rol = str(
                    valor_fila(
                        fila,
                        "Rol",
                        "",
                    )
                ).strip()

                institucion = str(
                    valor_fila(
                        fila,
                        "Institución",
                        "",
                    )
                ).strip()

                lugar = str(
                    valor_fila(
                        fila,
                        "Lugar",
                        "",
                    )
                ).strip()

                fecha = formatear_fecha(
                    valor_fila(
                        fila,
                        "Fecha",
                        "",
                    )
                )

                if not titulo:
                    continue

                p_item = doc.add_paragraph(
                    style="List Bullet"
                )

                run_t = p_item.add_run(
                    titulo
                )

                run_t.bold = True

                detalles = []

                if rol:
                    detalles.append(
                        f"Rol: {rol}"
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

                    p_item.add_run(
                        ". "
                        + ", ".join(detalles)
                        + "."
                    )

    buffer = io.BytesIO()

    doc.save(buffer)

    buffer.seek(0)

    return buffer


# ============================================================
# APLICACIÓN
# ============================================================

st.title(
    "📄 Sistema de Gestión de CV "
    "— Dra. María Griselda Günther"
)


service = obtener_servicio_drive()


if service:

    # --------------------------------------------------------
    # DRIVE
    # --------------------------------------------------------

    with st.spinner(
        "🔧 Verificando estructura de Google Drive..."
    ):

        estructura_drive = (
            inicializar_estructura_drive(
                service
            )
        )

    if not estructura_drive:

        st.error(
            "No fue posible inicializar "
            "Google Drive."
        )

        st.stop()

    # --------------------------------------------------------
    # EXCEL
    # --------------------------------------------------------

    excel_id, found_name = (
        buscar_excel_en_drive(
            service
        )
    )

    if not excel_id:

        st.warning(
            "No se encontró la base de datos."
        )

        archivo_excel = st.file_uploader(
            "Sube Base_de_Datos_Probatorios_y_CV.xlsx",
            type=["xlsx"],
        )

        if archivo_excel:

            metadata = {
                "name":
                "Base_de_Datos_Probatorios_y_CV.xlsx"
            }

            media = MediaIoBaseUpload(
                io.BytesIO(
                    archivo_excel.getvalue()
                ),
                mimetype=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                resumable=True,
            )

            service.files().create(
                body=metadata,
                media_body=media,
                fields="id",
            ).execute()

            st.success(
                "Base de datos subida correctamente."
            )

            st.rerun()

        st.stop()

    # --------------------------------------------------------
    # CARGAR Y LIMPIAR DATAFRAME
    # --------------------------------------------------------

    df_original = cargar_datos_drive(
        service,
        excel_id,
    )

    df = preparar_dataframe(
        df_original
    )

    # --------------------------------------------------------
    # DETECTAR SI HAY QUE MIGRAR EL EXCEL
    # --------------------------------------------------------

    columnas_originales = list(
        df_original.columns
    )

    columnas_nuevas = list(
        df.columns
    )

    necesita_migracion = (
        columnas_originales
        != columnas_nuevas
    )

    if necesita_migracion:

        st.warning(
            "⚠️ La base de datos contiene columnas "
            "antiguas. La aplicación ha preparado "
            "una estructura limpia."
        )

        if st.button(
            "🧹 Aplicar limpieza y actualizar Excel",
            type="primary",
        ):

            with st.spinner(
                "Limpiando estructura del Excel..."
            ):

                actualizar_excel_drive(
                    service,
                    excel_id,
                    df,
                )

            st.success(
                "Excel actualizado correctamente."
            )

            st.rerun()

    # ========================================================
    # TABS
    # ========================================================

    (
        tab_buscar,
        tab_editar,
        tab_nuevo,
        tab_paquetes,
        tab_cv,
    ) = st.tabs(
        [
            "🔍 Buscar",
            "✏️ Editar",
            "➕ Nueva actividad",
            "📦 Paquetes ZIP",
            "📄 Generar CV",
        ]
    )

    # ========================================================
    # BUSCADOR
    # ========================================================

    with tab_buscar:

        st.subheader(
            "🔍 Buscador de actividades"
        )

        col1, col2, col3 = st.columns(3)

        with col1:

            anios = sorted(
                pd.to_numeric(
                    df["Año"],
                    errors="coerce",
                )
                .dropna()
                .astype(int)
                .unique()
                .tolist(),
                reverse=True,
            )

            filtro_anio = st.selectbox(
                "Año",
                ["Todos"] + anios,
            )

        with col2:

            filtro_categoria = st.selectbox(
                "Categoría",
                ["Todas"] + CATEGORIAS,
            )

        with col3:

            filtro_texto = st.text_input(
                "Buscar",
                placeholder=(
                    "Título, institución, ID..."
                ),
            )

        resultado = df.copy()

        if filtro_anio != "Todos":

            resultado = resultado[
                pd.to_numeric(
                    resultado["Año"],
                    errors="coerce",
                )
                == int(filtro_anio)
            ]

        if filtro_categoria != "Todas":

            resultado = resultado[
                resultado["Categoría"]
                == filtro_categoria
            ]

        if filtro_texto:

            mask = resultado.apply(
                lambda fila:
                fila.astype(str)
                .str.contains(
                    filtro_texto,
                    case=False,
                    na=False,
                    regex=False,
                )
                .any(),
                axis=1,
            )

            resultado = resultado[mask]

        st.write(
            f"**{len(resultado)} registros encontrados.**"
        )

        columnas_visibles = [
            "ID",
            "Año",
            "Fecha",
            "Categoría",
            "Rol",
            "Título",
            "Institución",
            "Lugar",
            "Estado_Probatorio",
            "Incluir_en_CV",
            "Nombre_Archivo_PDF",
        ]

        st.dataframe(
            resultado[columnas_visibles],
            use_container_width=True,
            hide_index=True,
        )


    # ========================================================
    # EDITAR
    # ========================================================

    with tab_editar:

        st.subheader(
            "✏️ Editar actividad y administrar probatorios"
        )

        if df.empty:

            st.info(
                "No existen registros."
            )

        else:

            busqueda = st.text_input(
                "Buscar registro",
                placeholder=(
                    "ID, título, institución..."
                ),
                key="buscar_edicion",
            )

            seleccion = df.copy()

            if busqueda:

                mask = seleccion.apply(
                    lambda fila:
                    fila.astype(str)
                    .str.contains(
                        busqueda,
                        case=False,
                        na=False,
                        regex=False,
                    )
                    .any(),
                    axis=1,
                )

                seleccion = seleccion[mask]

            if seleccion.empty:

                st.warning(
                    "No se encontraron registros."
                )

            else:

                def etiqueta_registro(indice):

                    fila = df.loc[indice]

                    return (
                        f"{valor_fila(fila, 'ID')} — "
                        f"{valor_fila(fila, 'Título', 'Sin título')}"
                    )

                indice = st.selectbox(
                    "Selecciona una actividad",
                    seleccion.index.tolist(),
                    format_func=etiqueta_registro,
                )

                fila = df.loc[indice].copy()

                probatorios = (
                    obtener_probatorios_de_fila(
                        fila
                    )
                )

                st.markdown(
                    f"### {valor_fila(fila, 'ID')}"
                )

                st.info(
                    f"Esta actividad tiene "
                    f"**{len(probatorios)} probatorio(s)**."
                )

                # --------------------------------------------
                # DATOS
                # --------------------------------------------

                with st.form(
                    "form_editar"
                ):

                    col1, col2 = st.columns(2)

                    with col1:

                        anio = st.number_input(
                            "Año",
                            min_value=1900,
                            max_value=2100,
                            value=int(
                                float(
                                    valor_fila(
                                        fila,
                                        "Año",
                                        datetime.now().year,
                                    )
                                )
                            ),
                            step=1,
                        )

                        fecha_raw = valor_fila(
                            fila,
                            "Fecha",
                            "",
                        )

                        try:
                            fecha = pd.to_datetime(
                                fecha_raw
                            ).date()
                        except Exception:
                            fecha = datetime.now().date()

                        fecha = st.date_input(
                            "Fecha",
                            value=fecha,
                        )

                        categoria = st.selectbox(
                            "Categoría",
                            CATEGORIAS,
                            index=(
                                CATEGORIAS.index(
                                    valor_fila(
                                        fila,
                                        "Categoría",
                                        CATEGORIAS[0],
                                    )
                                )
                                if valor_fila(
                                    fila,
                                    "Categoría",
                                    CATEGORIAS[0],
                                ) in CATEGORIAS
                                else 0
                            ),
                        )

                        rol = st.text_input(
                            "Rol",
                            value=str(
                                valor_fila(
                                    fila,
                                    "Rol",
                                    "",
                                )
                            ),
                        )

                    with col2:

                        titulo = st.text_input(
                            "Título *",
                            value=str(
                                valor_fila(
                                    fila,
                                    "Título",
                                    "",
                                )
                            ),
                        )

                        institucion = st.text_input(
                            "Institución",
                            value=str(
                                valor_fila(
                                    fila,
                                    "Institución",
                                    "",
                                )
                            ),
                        )

                        lugar = st.text_input(
                            "Lugar / Sede",
                            value=str(
                                valor_fila(
                                    fila,
                                    "Lugar",
                                    "",
                                )
                            ),
                        )

                        estado = st.selectbox(
                            "Estado",
                            [
                                "Verificado / En Drive",
                                "Pendiente de Escanear",
                                "En Trámite",
                            ],
                            index=(
                                [
                                    "Verificado / En Drive",
                                    "Pendiente de Escanear",
                                    "En Trámite",
                                ].index(
                                    valor_fila(
                                        fila,
                                        "Estado_Probatorio",
                                        "Verificado / En Drive",
                                    )
                                )
                                if valor_fila(
                                    fila,
                                    "Estado_Probatorio",
                                    "Verificado / En Drive",
                                )
                                in [
                                    "Verificado / En Drive",
                                    "Pendiente de Escanear",
                                    "En Trámite",
                                ]
                                else 0
                            ),
                        )

                    incluir = st.radio(
                        "¿Incluir en CV?",
                        ["Sí", "No"],
                        horizontal=True,
                        index=(
                            0
                            if valor_fila(
                                fila,
                                "Incluir_en_CV",
                                "No",
                            )
                            == "Sí"
                            else 1
                        ),
                    )

                    notas = st.text_area(
                        "Notas / Observaciones",
                        value=str(
                            valor_fila(
                                fila,
                                "Notas_Observaciones",
                                "",
                            )
                        ),
                    )

                    # ----------------------------------------
                    # PROBATORIOS EXISTENTES
                    # ----------------------------------------

                    st.markdown(
                        "### 📎 Probatorios actuales"
                    )

                    if probatorios:

                        for i, p in enumerate(
                            probatorios,
                            start=1,
                        ):

                            st.markdown(
                                f"**{i}. {p['nombre']}**"
                            )

                            if p["enlace"]:

                                st.markdown(
                                    f"[🔗 Abrir en Drive]"
                                    f"({p['enlace']})"
                                )

                    else:

                        st.warning(
                            "Esta actividad todavía "
                            "no tiene probatorios."
                        )

                    st.markdown(
                        "### ➕ Agregar probatorios"
                    )

                    nuevos_archivos = st.file_uploader(
                        "Puedes seleccionar uno o varios archivos",
                        type=[
                            "pdf",
                            "png",
                            "jpg",
                            "jpeg",
                        ],
                        accept_multiple_files=True,
                        key=f"edit_files_{indice}",
                    )

                    reemplazar_todos = st.checkbox(
                        "Reemplazar TODOS los probatorios actuales",
                        value=False,
                    )

                    guardar = st.form_submit_button(
                        "💾 Guardar cambios",
                        type="primary",
                    )

                # --------------------------------------------
                # GUARDAR
                # --------------------------------------------

                if guardar:

                    if not titulo.strip():

                        st.error(
                            "El título es obligatorio."
                        )

                        st.stop()

                    with st.spinner(
                        "Actualizando actividad y Drive..."
                    ):

                        probatorios_finales = []

                        # ------------------------------------
                        # SI CONSERVAMOS LOS EXISTENTES
                        # ------------------------------------

                        if not reemplazar_todos:

                            probatorios_finales = (
                                probatorios.copy()
                            )

                        # ------------------------------------
                        # SUBIR NUEVOS
                        # ------------------------------------

                        for numero, archivo in enumerate(
                            nuevos_archivos or [],
                            start=len(
                                probatorios_finales
                            ) + 1,
                        ):

                            extension = os.path.splitext(
                                archivo.name
                            )[1].lower()

                            nombre = nombre_probatorio(
                                titulo,
                                anio,
                                categoria,
                                numero,
                                extension,
                            )

                            carpeta = (
                                obtener_carpeta_probatorio(
                                    service,
                                    estructura_drive,
                                    anio,
                                    categoria,
                                )
                            )

                            resultado_archivo = (
                                subir_a_google_drive(
                                    service,
                                    nombre,
                                    archivo.getvalue(),
                                    carpeta,
                                )
                            )

                            if not resultado_archivo[2]:

                                st.error(
                                    f"No se pudo subir "
                                    f"{archivo.name}"
                                )

                                st.stop()

                            (
                                nombre_drive,
                                enlace,
                                archivo_id,
                            ) = resultado_archivo

                            probatorios_finales.append(
                                {
                                    "nombre": nombre_drive,
                                    "enlace": enlace,
                                    "id": archivo_id,
                                }
                            )

                        # ------------------------------------
                        # SI REEMPLAZAMOS
                        # MANDAR LOS ANTIGUOS A PAPELERA
                        # ------------------------------------

                        if reemplazar_todos:

                            for p in probatorios:

                                if p["id"]:

                                    eliminar_archivo_drive(
                                        service,
                                        p["id"],
                                    )

                        # ------------------------------------
                        # SI NO SE SUBIÓ NADA NUEVO
                        # Y CAMBIÓ AÑO/CATEGORÍA
                        # ------------------------------------

                        elif (
                            not nuevos_archivos
                            and (
                                int(
                                    valor_fila(
                                        fila,
                                        "Año",
                                        anio,
                                    )
                                )
                                != anio
                                or valor_fila(
                                    fila,
                                    "Categoría",
                                    "",
                                )
                                != categoria
                            )
                        ):

                            carpeta_destino = (
                                obtener_carpeta_probatorio(
                                    service,
                                    estructura_drive,
                                    anio,
                                    categoria,
                                )
                            )

                            for p in probatorios_finales:

                                if p["id"]:

                                    mover_archivo_drive(
                                        service,
                                        p["id"],
                                        carpeta_destino,
                                    )

                        # ------------------------------------
                        # ACTUALIZAR FILA
                        # ------------------------------------

                        df_actualizado = df.copy()

                        df_actualizado.at[
                            indice,
                            "Año",
                        ] = anio

                        df_actualizado.at[
                            indice,
                            "Fecha",
                        ] = str(fecha)

                        df_actualizado.at[
                            indice,
                            "Categoría",
                        ] = categoria

                        df_actualizado.at[
                            indice,
                            "Rol",
                        ] = rol

                        df_actualizado.at[
                            indice,
                            "Título",
                        ] = titulo.strip()

                        df_actualizado.at[
                            indice,
                            "Institución",
                        ] = institucion

                        df_actualizado.at[
                            indice,
                            "Lugar",
                        ] = lugar

                        df_actualizado.at[
                            indice,
                            "Estado_Probatorio",
                        ] = (
                            "Verificado / En Drive"
                            if probatorios_finales
                            else estado
                        )

                        df_actualizado.at[
                            indice,
                            "Incluir_en_CV",
                        ] = incluir

                        df_actualizado.at[
                            indice,
                            "Notas_Observaciones",
                        ] = notas

                        df_actualizado.loc[
                            indice
                        ] = actualizar_probatorios_fila(
                            df_actualizado.loc[indice],
                            probatorios_finales,
                        )

                        actualizar_excel_drive(
                            service,
                            excel_id,
                            df_actualizado,
                        )

                    st.success(
                        "✅ Actividad actualizada correctamente."
                    )

                    st.rerun()


    # ========================================================
    # NUEVA ACTIVIDAD
    # ========================================================

    with tab_nuevo:

        st.subheader(
            "➕ Registrar nueva actividad"
        )

        with st.form(
            "form_nueva_actividad"
        ):

            col1, col2 = st.columns(2)

            with col1:

                nuevo_id = st.text_input(
                    "ID",
                    value=siguiente_id_registro(df),
                )

                anio = st.number_input(
                    "Año",
                    min_value=1900,
                    max_value=2100,
                    value=datetime.now().year,
                    step=1,
                )

                fecha = st.date_input(
                    "Fecha"
                )

                categoria = st.selectbox(
                    "Categoría",
                    CATEGORIAS,
                )

                rol = st.text_input(
                    "Rol"
                )

            with col2:

                titulo = st.text_input(
                    "Título *"
                )

                institucion = st.text_input(
                    "Institución"
                )

                lugar = st.text_input(
                    "Lugar / Sede"
                )

                estado = st.selectbox(
                    "Estado",
                    [
                        "Verificado / En Drive",
                        "Pendiente de Escanear",
                        "En Trámite",
                    ],
                )

                incluir = st.radio(
                    "¿Incluir en CV?",
                    ["Sí", "No"],
                    horizontal=True,
                )

            st.markdown("---")

            st.subheader(
                "📎 Probatorios"
            )

            archivos = st.file_uploader(
                "Selecciona uno o varios probatorios",
                type=[
                    "pdf",
                    "png",
                    "jpg",
                    "jpeg",
                ],
                accept_multiple_files=True,
            )

            notas = st.text_area(
                "Notas / Observaciones"
            )

            guardar = st.form_submit_button(
                "💾 Guardar actividad",
                type="primary",
            )

        if guardar:

            if not titulo.strip():

                st.error(
                    "El título es obligatorio."
                )

                st.stop()

            with st.spinner(
                "Subiendo actividad y probatorios..."
            ):

                probatorios = []

                for numero, archivo in enumerate(
                    archivos or [],
                    start=1,
                ):

                    extension = os.path.splitext(
                        archivo.name
                    )[1].lower()

                    nombre = nombre_probatorio(
                        titulo,
                        anio,
                        categoria,
                        numero,
                        extension,
                    )

                    carpeta = (
                        obtener_carpeta_probatorio(
                            service,
                            estructura_drive,
                            anio,
                            categoria,
                        )
                    )

                    resultado_archivo = (
                        subir_a_google_drive(
                            service,
                            nombre,
                            archivo.getvalue(),
                            carpeta,
                        )
                    )

                    if not resultado_archivo[2]:

                        st.error(
                            f"No se pudo subir "
                            f"{archivo.name}"
                        )

                        st.stop()

                    (
                        nombre_drive,
                        enlace,
                        archivo_id,
                    ) = resultado_archivo

                    probatorios.append(
                        {
                            "nombre": nombre_drive,
                            "enlace": enlace,
                            "id": archivo_id,
                        }
                    )

                nueva_fila = {
                    "ID": nuevo_id.strip(),
                    "Año": anio,
                    "Fecha": str(fecha),
                    "Categoría": categoria,
                    "Rol": rol,
                    "Título": titulo.strip(),
                    "Institución": institucion,
                    "Lugar": lugar,
                    "Estado_Probatorio": (
                        "Verificado / En Drive"
                        if probatorios
                        else estado
                    ),
                    "Incluir_en_CV": incluir,
                    "Notas_Observaciones": notas,
                    "Nombre_Archivo_PDF": lista_a_celda(
                        [
                            p["nombre"]
                            for p in probatorios
                        ]
                    ),
                    "Enlace_Drive_Probatorio": lista_a_celda(
                        [
                            p["enlace"]
                            for p in probatorios
                        ]
                    ),
                    "ID_Drive_Probatorio": lista_a_celda(
                        [
                            p["id"]
                            for p in probatorios
                        ]
                    ),
                }

                df_actualizado = pd.concat(
                    [
                        df,
                        pd.DataFrame(
                            [nueva_fila]
                        ),
                    ],
                    ignore_index=True,
                )

                actualizar_excel_drive(
                    service,
                    excel_id,
                    df_actualizado,
                )

            st.success(
                f"✅ Actividad '{titulo}' registrada."
            )

            st.info(
                f"Se guardaron "
                f"**{len(probatorios)} probatorios** "
                f"en Drive."
            )

            st.rerun()


    # ========================================================
    # PAQUETES ZIP
    # ========================================================

    with tab_paquetes:

        st.subheader(
            "📦 Generador de paquetes de probatorios"
        )

        st.write(
            "Selecciona actividades o utiliza filtros "
            "para construir un ZIP con todos sus "
            "probatorios."
        )

        st.markdown(
            "### 🎯 Filtrar actividades"
        )

        col1, col2, col3 = st.columns(3)

        with col1:

            anios_zip = sorted(
                pd.to_numeric(
                    df["Año"],
                    errors="coerce",
                )
                .dropna()
                .astype(int)
                .unique()
                .tolist(),
                reverse=True,
            )

            filtro_zip_anio = st.selectbox(
                "Año",
                ["Todos"] + anios_zip,
                key="zip_anio",
            )

        with col2:

            filtro_zip_categoria = (
                st.selectbox(
                    "Categoría",
                    ["Todas"] + CATEGORIAS,
                    key="zip_categoria",
                )
            )

        with col3:

            filtro_zip_texto = st.text_input(
                "Buscar",
                key="zip_texto",
            )

        df_zip = df.copy()

        if filtro_zip_anio != "Todos":

            df_zip = df_zip[
                pd.to_numeric(
                    df_zip["Año"],
                    errors="coerce",
                )
                == int(filtro_zip_anio)
            ]

        if filtro_zip_categoria != "Todas":

            df_zip = df_zip[
                df_zip["Categoría"]
                == filtro_zip_categoria
            ]

        if filtro_zip_texto:

            mask = df_zip.apply(
                lambda fila:
                fila.astype(str)
                .str.contains(
                    filtro_zip_texto,
                    case=False,
                    na=False,
                    regex=False,
                )
                .any(),
                axis=1,
            )

            df_zip = df_zip[mask]

        st.write(
            f"**{len(df_zip)} actividades encontradas.**"
        )

        if not df_zip.empty:

            opciones = df_zip[
                "ID"
            ].astype(str).tolist()

            seleccion_ids = st.multiselect(
                "Selecciona las actividades",
                opciones,
                format_func=lambda x:
                f"{x} — "
                f"{valor_fila(df[df['ID'].astype(str) == str(x)].iloc[0], 'Título', 'Sin título')}",
            )

            col_a, col_b = st.columns(2)

            with col_a:

                if st.button(
                    "☑️ Seleccionar todas las actividades filtradas"
                ):

                    st.session_state[
                        "zip_seleccionadas"
                    ] = opciones

                    st.rerun()

            with col_b:

                if st.button(
                    "🧹 Limpiar selección"
                ):

                    st.session_state[
                        "zip_seleccionadas"
                    ] = []

                    st.rerun()

            seleccion_ids = st.session_state.get(
                "zip_seleccionadas",
                seleccion_ids,
            )

            if seleccion_ids:

                df_seleccion_zip = df[
                    df["ID"]
                    .astype(str)
                    .isin(
                        [str(x) for x in seleccion_ids]
                    )
                ].copy()

            else:

                df_seleccion_zip = pd.DataFrame()

            st.write(
                f"**{len(df_seleccion_zip)} actividades "
                f"seleccionadas.**"
            )

            if st.button(
                "📦 Generar ZIP",
                type="primary",
                disabled=df_seleccion_zip.empty,
            ):

                with st.spinner(
                    "Descargando probatorios y "
                    "construyendo ZIP..."
                ):

                    zip_bytes, cantidad = (
                        crear_zip_probatorios(
                            service,
                            df_seleccion_zip,
                        )
                    )

                if cantidad == 0:

                    st.error(
                        "No se encontraron archivos "
                        "válidos para estas actividades."
                    )

                else:

                    fecha_zip = datetime.now().strftime(
                        "%Y%m%d_%H%M"
                    )

                    nombre_zip = (
                        f"Paquete_Probatorios_"
                        f"{fecha_zip}.zip"
                    )

                    st.success(
                        f"✅ ZIP generado con "
                        f"**{cantidad} archivos**."
                    )

                    st.download_button(
                        "📥 Descargar ZIP",
                        data=zip_bytes,
                        file_name=nombre_zip,
                        mime="application/zip",
                        type="primary",
                    )

        else:

            st.info(
                "No hay actividades que coincidan "
                "con los filtros."
            )


    # ========================================================
    # GENERADOR DE CV
    # ========================================================

    with tab_cv:

        st.subheader(
            "📄 Generar CV en Word"
        )

        df_cv = df[
            df["Incluir_en_CV"]
            .astype(str)
            .str.strip()
            .str.lower()
            == "sí"
        ].copy()

        st.info(
            f"Actualmente hay "
            f"**{len(df_cv)} actividades** "
            f"marcadas para el CV."
        )

        if not df_cv.empty:

            documento = crear_cv_word(
                df
            )

            st.download_button(
                "📥 Descargar CV (.docx)",
                data=documento,
                file_name=(
                    "CV_Dra_Maria_Griselda_Gunther.docx"
                ),
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "wordprocessingml.document"
                ),
                type="primary",
            )

            st.markdown(
                "### 👁️ Actividades incluidas"
            )

            st.dataframe(
                df_cv[
                    [
                        "ID",
                        "Año",
                        "Fecha",
                        "Categoría",
                        "Rol",
                        "Título",
                        "Institución",
                    ]
                ],
                use_container_width=True,
                hide_index=True,
            )

        else:

            st.warning(
                "No hay actividades marcadas "
                "para incluir en el CV."
            )
