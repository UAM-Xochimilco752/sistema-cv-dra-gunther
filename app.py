import io
import os
import pickle
import mimetypes
import re
from datetime import date

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
    page_title="Expediente Académico y CV",
    page_icon="🎓",
    layout="wide",
)

NOMBRE_CARPETA_RAIZ = "CV — Sistema de Gestión"
ANIOS_PROBATORIOS = list(range(2020, 2031))

# Las actividades pueden ser SNII, complementarias o ambas.
COMPONENTES_SNII = [
    "Componente 1 — Producción de investigación",
    "Componente 2 — Fortalecimiento y consolidación de la comunidad",
    "Componente 3 — Divulgación",
    "Actividad complementaria / no aplica",
]

# Estas categorías siguen sirviendo para ordenar probatorios,
# pero ahora existe una opción abierta para actividades que no encajan.
CATEGORIAS = [
    "Coordinación de Libros",
    "Capítulos de Libros / Artículos",
    "Ponencias y Conferencias",
    "Presentaciones de Libros",
    "Comisiones y Arbitrajes",
    "Cursos e Impartición de Clases",
    "Premios y Reconocimientos",
    "Asesorías",
    "Otros / Sin clasificación",
]

TIPOS_PRODUCTO = [
    "Artículo", "Libro", "Capítulo", "Reporte", "Informe",
    "Dossier o número temático", "Antología", "Traducción",
    "Prólogo o estudio introductorio", "Curaduría", "Datos primarios",
    "Software", "Producto tecnológico", "Proceso o estrategia digital",
    "Patente", "Derecho de autor", "Diseño industrial", "Modelo de utilidad",
    "Esquema de trazado", "Marca", "Licenciamiento",
    "Consultoría o asesoría técnica especializada", "Creación de empresa",
    "Asociación estratégica", "Informe técnico", "Curso", "Diplomado",
    "Capacitación", "Taller", "Seminario", "Tutoría", "Tesis", "Tesina",
    "Portafolio", "Proyecto de investigación", "Plan de estudios",
    "Acuerdo o convenio", "Coordinación de programa o centro", "Jurado",
    "Evaluación de programa o proyecto SECIHTI", "Dictaminación de publicación",
    "Dictaminación especializada", "Medios escritos",
    "Medios audiovisuales/radiofónicos/digitales",
    "Museografía / educación no formal", "Evento o comunicación",
    "Comisión académica", "Constancia", "Conferencia", "Participación académica",
    "Coeficiente / evaluación académica", "Docencia", "Otro",
]

SUBTIPOS = [
    "Investigación científica/humanística", "Desarrollo tecnológico",
    "Innovación", "Docencia", "Trabajo de titulación",
    "Desarrollo institucional", "Evaluación académica",
    "Divulgación", "Actividad académica complementaria", "Otro",
]

MODALIDADES = [
    "No aplica", "Nacional", "Internacional",
    "Presencial", "Virtual", "Híbrida",
]

CARACTERISTICAS_SNII = [
    "Relevante", "Pertinente", "Sostenida", "Diversa",
    "Comprometida con la formación", "Colaborativa", "Constante",
    "Inclusiva", "Gratuita", "Comprensible",
]

COLUMNAS_NUEVAS = [
    "ID", "Año", "Fecha",
    "Componente_SNII", "Tipo_Producto_SNII", "Subtipo_SNII",
    "Categoria_CV", "Rol_Participacion", "Titulo_Actividad_o_Publicacion",
    "Evento_Revista_Libro", "Institucion_Organizacion", "Lugar_Sede",
    "Modalidad", "Autores", "Coautores", "Nivel_Formacion",
    "Estudiantes_Beneficiados", "Proyecto_Linea_Investigacion",
    "Descripcion_Aportacion", "Relevancia_Pertinencia",
    "Impacto_Beneficio_Social", "Caracteristicas_SNII", "Arbitrado",
    "Publicado", "Revista_Editorial", "Volumen_Numero", "Paginas",
    "ISBN_ISSN", "DOI_URL", "Nombre_Archivo_PDF",
    "Enlace_Drive_Probatorio", "ID_Drive_Probatorio",
    "Estado_Probatorio", "Incluir_en_CV", "Incluir_en_CV_SNII",
    "Redaccion_CV", "Notas_Observaciones",
]


# ============================================================
# UTILIDADES
# ============================================================

def limpiar_texto(valor):
    if pd.isna(valor):
        return ""
    texto = str(valor).strip()
    if texto.lower() in {"", "nan", "none", "ninguno", "ninguna", "n/a", "na", "-"}:
        return ""
    return texto


def obtener_valor(row, columna):
    if columna not in row.index:
        return ""
    return limpiar_texto(row.get(columna))


def valor_si(valor, default="No"):
    return "Sí" if str(valor).strip().lower() in {"sí", "si", "yes", "true"} else default


def formatear_fecha(row):
    fecha = obtener_valor(row, "Fecha")
    anio = obtener_valor(row, "Año")
    if fecha:
        try:
            return pd.to_datetime(fecha).strftime("%d/%m/%Y")
        except Exception:
            return fecha
    return anio


def normalizar_df(df):
    """Hace compatible una base vieja con la nueva estructura."""
    df = df.copy()

    # Alias de columnas antiguas, por si existen.
    aliases = {
        "Categoria": "Categoria_CV",
        "Categoría": "Categoria_CV",
        "Titulo": "Titulo_Actividad_o_Publicacion",
        "Título": "Titulo_Actividad_o_Publicacion",
        "Enlace_Probatorio": "Enlace_Drive_Probatorio",
        "ID_Drive": "ID_Drive_Probatorio",
    }

    for antigua, nueva in aliases.items():
        if antigua in df.columns and nueva not in df.columns:
            df[nueva] = df[antigua]

    for columna in COLUMNAS_NUEVAS:
        if columna not in df.columns:
            df[columna] = ""

    # Mantener columnas adicionales que ya existan en la base.
    return df


def generar_id(df):
    existentes = set(df["ID"].astype(str).str.strip()) if "ID" in df.columns else set()
    numeros = []
    for x in existentes:
        m = re.match(r"ACT-(\d+)$", x)
        if m:
            numeros.append(int(m.group(1)))
    n = max(numeros, default=0) + 1
    nuevo = f"ACT-{n:04d}"
    while nuevo in existentes:
        n += 1
        nuevo = f"ACT-{n:04d}"
    return nuevo


def limpiar_nombre_archivo(texto):
    texto = re.sub(r'[\\/:*?"<>|]+', "", str(texto))
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto[:100] or "Sin_Titulo"


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
            st.error(f"Error al refrescar credenciales de Google: {e}")
            return None

    if not creds:
        st.error("⚠️ No se encontró el archivo 'token.pickle'.")
        return None

    return build("drive", "v3", credentials=creds)


def buscar_carpeta(service, nombre, parent_id=None):
    nombre_escapado = nombre.replace("'", "\\'")
    query = (
        f"name = '{nombre_escapado}' "
        "and mimeType = 'application/vnd.google-apps.folder' "
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
            pageSize=10,
        ).execute()
        archivos = resultado.get("files", [])
        return archivos[0] if archivos else None
    except Exception as e:
        st.error(f"Error buscando carpeta '{nombre}': {e}")
        return None


def crear_carpeta(service, nombre, parent_id=None):
    metadata = {"name": nombre, "mimeType": "application/vnd.google-apps.folder"}
    if parent_id:
        metadata["parents"] = [parent_id]
    try:
        return service.files().create(
            body=metadata,
            fields="id,name,webViewLink"
        ).execute()
    except Exception as e:
        st.error(f"Error creando carpeta '{nombre}': {e}")
        return None


def obtener_o_crear_carpeta(service, nombre, parent_id=None):
    carpeta = buscar_carpeta(service, nombre, parent_id)
    return carpeta or crear_carpeta(service, nombre, parent_id)


@st.cache_resource
def inicializar_estructura_drive(_service):
    estructura = {}

    raiz = obtener_o_crear_carpeta(_service, NOMBRE_CARPETA_RAIZ)
    if not raiz:
        return {}
    estructura["raiz"] = raiz

    principales = [
        "00 — Administración",
        "01 — Datos personales y CV",
        "02 — Probatorios",
        "10 — CV generados",
    ]

    for nombre in principales:
        carpeta = obtener_o_crear_carpeta(_service, nombre, raiz["id"])
        if carpeta:
            estructura[nombre] = carpeta

    probatorios = estructura.get("02 — Probatorios")
    if probatorios:
        for anio in ANIOS_PROBATORIOS:
            carpeta_anio = obtener_o_crear_carpeta(
                _service, str(anio), probatorios["id"]
            )
            if not carpeta_anio:
                continue
            estructura[f"probatorios_{anio}"] = carpeta_anio

            for categoria in CATEGORIAS:
                carpeta_categoria = obtener_o_crear_carpeta(
                    _service, categoria, carpeta_anio["id"]
                )
                if carpeta_categoria:
                    estructura[f"probatorios_{anio}_{categoria}"] = carpeta_categoria

    return estructura


def obtener_carpeta_probatorio(service, estructura, anio, categoria):
    clave = f"probatorios_{anio}_{categoria}"
    if clave in estructura:
        return estructura[clave]

    probatorios = estructura.get("02 — Probatorios")
    if not probatorios:
        return None

    carpeta_anio = obtener_o_crear_carpeta(
        service, str(anio), probatorios["id"]
    )
    if not carpeta_anio:
        return None

    carpeta = obtener_o_crear_carpeta(
        service, categoria, carpeta_anio["id"]
    )
    if carpeta:
        estructura[clave] = carpeta
    return carpeta


def obtener_mimetype(nombre):
    tipo, _ = mimetypes.guess_type(nombre)
    return tipo or "application/octet-stream"


def subir_a_google_drive(service, nombre_archivo, bytes_archivo, carpeta_destino):
    metadata = {"name": nombre_archivo, "parents": [carpeta_destino["id"]]}
    media = MediaIoBaseUpload(
        io.BytesIO(bytes_archivo),
        mimetype=obtener_mimetype(nombre_archivo),
        resumable=True,
    )
    try:
        archivo = service.files().create(
            body=metadata,
            media_body=media,
            fields="id,name,webViewLink,parents",
        ).execute()
        return archivo.get("webViewLink", ""), archivo.get("id", "")
    except Exception as e:
        st.error(f"Error al subir archivo: {e}")
        return "", ""


def eliminar_archivo_drive(service, file_id):
    if not file_id:
        return True
    try:
        service.files().delete(fileId=file_id).execute()
        return True
    except Exception as e:
        st.warning(f"No se pudo eliminar el probatorio de Drive: {e}")
        return False


# ============================================================
# EXCEL EN DRIVE
# ============================================================

def buscar_excel_en_drive(service):
    try:
        resultado = service.files().list(
            q="trashed = false",
            fields="files(id,name,mimeType,modifiedTime)",
            pageSize=100,
        ).execute()

        for archivo in resultado.get("files", []):
            nombre = archivo["name"]
            if "Base_de_Datos_Probatorios_y_CV" in nombre and nombre.lower().endswith(".xlsx"):
                return archivo["id"], nombre
        return None, None
    except Exception as e:
        st.error(f"Error consultando Google Drive: {e}")
        return None, None


def cargar_datos_drive(service, file_id):
    request = service.files().get_media(fileId=file_id)
    buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(buffer, request)
    done = False

    while not done:
        _, done = downloader.next_chunk()

    buffer.seek(0)
    df = pd.read_excel(buffer)
    return normalizar_df(df)


def actualizar_excel_drive(service, file_id, df):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    buffer.seek(0)

    media = MediaIoBaseUpload(
        buffer,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=True,
    )

    service.files().update(fileId=file_id, media_body=media).execute()


# ============================================================
# REDACCIÓN
# ============================================================

def redactar_registro(row):
    tipo = obtener_valor(row, "Tipo_Producto_SNII")
    titulo = obtener_valor(row, "Titulo_Actividad_o_Publicacion")
    autores = obtener_valor(row, "Autores")
    rol = obtener_valor(row, "Rol_Participacion")
    evento = obtener_valor(row, "Evento_Revista_Libro")
    institucion = obtener_valor(row, "Institucion_Organizacion")
    lugar = obtener_valor(row, "Lugar_Sede")
    fecha = formatear_fecha(row)
    revista = obtener_valor(row, "Revista_Editorial")
    volumen = obtener_valor(row, "Volumen_Numero")
    paginas = obtener_valor(row, "Paginas")
    isbn = obtener_valor(row, "ISBN_ISSN")
    doi = obtener_valor(row, "DOI_URL")
    arbitrado = obtener_valor(row, "Arbitrado")

    if not titulo:
        return ""

    if tipo == "Artículo":
        texto = f"{autores}. " if autores else ""
        texto += f"({fecha}). {titulo}."
        if revista:
            texto += f" {revista}."
        if volumen:
            texto += f" {volumen}."
        if paginas:
            texto += f" pp. {paginas}."
        if arbitrado == "Sí":
            texto += " Publicación arbitrada."
        if isbn:
            texto += f" ISSN/ISBN: {isbn}."
        if doi:
            texto += f" DOI/URL: {doi}."
        return texto.strip()

    if tipo in {"Libro", "Capítulo", "Antología", "Traducción", "Prólogo o estudio introductorio"}:
        texto = f"{autores}. " if autores else ""
        texto += f"({fecha}). {titulo}."
        if revista:
            texto += f" {revista}."
        if paginas:
            texto += f" pp. {paginas}."
        if isbn:
            texto += f" ISBN/ISSN: {isbn}."
        if doi:
            texto += f" DOI/URL: {doi}."
        return texto.strip()

    texto = f"{autores}. " if autores else ""
    texto += f"{titulo}."
    detalles = []

    if rol:
        detalles.append(f"Rol: {rol}")
    if tipo:
        detalles.append(tipo)
    if evento:
        detalles.append(evento)
    if institucion:
        detalles.append(institucion)
    if lugar:
        detalles.append(lugar)
    if fecha:
        detalles.append(fecha)

    if detalles:
        texto += " " + "; ".join(detalles) + "."

    if isbn:
        texto += f" ISBN/ISSN: {isbn}."
    if doi:
        texto += f" DOI/URL: {doi}."

    return texto.strip()


# ============================================================
# DOCUMENTOS WORD
# ============================================================

def configurar_documento():
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.9)
        section.right_margin = Inches(0.9)

    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(10.5)
    return doc


def agregar_encabezado_cv(doc, subtitulo):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("DRA. MARÍA GRISELDA GÜNTHER")
    r.bold = True
    r.font.size = Pt(17)
    r.font.color.rgb = RGBColor(0, 51, 102)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("CURRÍCULUM VITAE ACADÉMICO")
    r.bold = True
    r.font.size = Pt(12)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(subtitulo)
    r.italic = True
    r.font.size = Pt(9.5)


def preparar_trabajo(df, columna_inclusion):
    trabajo = normalizar_df(df.copy())
    trabajo["_anio_num"] = pd.to_numeric(trabajo["Año"], errors="coerce")
    trabajo["_fecha_num"] = pd.to_datetime(trabajo["Fecha"], errors="coerce")
    trabajo = trabajo.sort_values(
        ["_anio_num", "_fecha_num"],
        ascending=[False, False],
        na_position="last",
    )

    if columna_inclusion in trabajo.columns:
        trabajo = trabajo[
            trabajo[columna_inclusion].astype(str).str.strip().str.lower() == "sí"
        ]
    return trabajo


def texto_cv(row):
    texto = obtener_valor(row, "Redaccion_CV")
    return texto or redactar_registro(row)


def crear_cv_snii_word(df):
    """CV nuevo: organizado por componentes SNII y tipos de producto."""
    doc = configurar_documento()
    agregar_encabezado_cv(
        doc,
        "Documento de trabajo organizado conforme a la clasificación SNII",
    )

    trabajo = preparar_trabajo(df, "Incluir_en_CV_SNII")

    p = doc.add_paragraph()
    r = p.add_run("TRAYECTORIA ACADÉMICA Y CIENTÍFICA")
    r.bold = True
    r.font.size = Pt(13)
    r.font.color.rgb = RGBColor(0, 51, 102)

    doc.add_paragraph(
        "El documento reúne las actividades marcadas para el CV SNII. "
        "Las actividades complementarias pueden conservarse en el expediente "
        "sin necesidad de forzarlas dentro de un rubro SNII."
    )

    componentes = [
        (
            "Componente 1 — Producción de investigación",
            "I. PRODUCCIÓN DE INVESTIGACIÓN CIENTÍFICA, HUMANÍSTICA Y TECNOLÓGICA",
        ),
        (
            "Componente 2 — Fortalecimiento y consolidación de la comunidad",
            "II. FORTALECIMIENTO Y CONSOLIDACIÓN DE LA COMUNIDAD",
        ),
        (
            "Componente 3 — Divulgación",
            "III. DIVULGACIÓN Y ACCESO UNIVERSAL AL CONOCIMIENTO",
        ),
        (
            "Actividad complementaria / no aplica",
            "IV. ACTIVIDADES ACADÉMICAS COMPLEMENTARIAS",
        ),
    ]

    for componente, titulo_componente in componentes:
        registros = trabajo[trabajo["Componente_SNII"] == componente]
        if registros.empty:
            continue

        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(16)
        r = p.add_run(titulo_componente)
        r.bold = True
        r.font.size = Pt(13)
        r.font.color.rgb = RGBColor(0, 51, 102)

        tipos_presentes = registros["Tipo_Producto_SNII"].dropna().astype(str).unique()

        for tipo in TIPOS_PRODUCTO:
            if tipo not in tipos_presentes:
                continue
            sub = registros[registros["Tipo_Producto_SNII"] == tipo]
            if sub.empty:
                continue

            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(8)
            r = p.add_run(tipo.upper())
            r.bold = True
            r.font.size = Pt(11)

            for _, row in sub.iterrows():
                texto = texto_cv(row)
                if texto:
                    p = doc.add_paragraph(style="List Bullet")
                    p.paragraph_format.space_after = Pt(4)
                    p.paragraph_format.line_spacing = 1.08
                    p.add_run(texto)

    agregar_anexo_control(doc, trabajo)
    return guardar_docx(doc)


def crear_cv_general_word(df):
    """CV general: no depende de componentes SNII; usa Categoria_CV."""
    doc = configurar_documento()
    agregar_encabezado_cv(
        doc,
        "Versión general del currículum vitae académico",
    )

    trabajo = preparar_trabajo(df, "Incluir_en_CV")

    p = doc.add_paragraph()
    r = p.add_run("TRAYECTORIA ACADÉMICA")
    r.bold = True
    r.font.size = Pt(13)
    r.font.color.rgb = RGBColor(0, 51, 102)

    doc.add_paragraph(
        "Esta versión conserva la lógica del CV general: las actividades "
        "se agrupan por categoría curricular y no se exige que pertenezcan "
        "a un componente SNII."
    )

    categorias_presentes = (
        trabajo["Categoria_CV"].fillna("Otros / Sin clasificación")
        .astype(str)
        .replace("", "Otros / Sin clasificación")
        .unique()
    )

    for categoria in CATEGORIAS:
        if categoria not in categorias_presentes:
            continue

        sub = trabajo[
            trabajo["Categoria_CV"].fillna("Otros / Sin clasificación").astype(str) == categoria
        ]
        if sub.empty:
            continue

        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        r = p.add_run(categoria.upper())
        r.bold = True
        r.font.size = Pt(13)
        r.font.color.rgb = RGBColor(0, 51, 102)

        for _, row in sub.iterrows():
            texto = texto_cv(row)
            if texto:
                p = doc.add_paragraph(style="List Bullet")
                p.paragraph_format.space_after = Pt(4)
                p.paragraph_format.line_spacing = 1.08
                p.add_run(texto)

    # Por seguridad, cualquier categoría nueva que no esté en la lista
    # también aparece.
    conocidas = set(CATEGORIAS)
    extras = sorted(set(categorias_presentes) - conocidas)

    for categoria in extras:
        sub = trabajo[
            trabajo["Categoria_CV"].fillna("").astype(str) == categoria
        ]
        if sub.empty:
            continue

        p = doc.add_paragraph()
        r = p.add_run(str(categoria).upper())
        r.bold = True
        r.font.size = Pt(13)

        for _, row in sub.iterrows():
            texto = texto_cv(row)
            if texto:
                p = doc.add_paragraph(style="List Bullet")
                p.add_run(texto)

    agregar_anexo_control(doc, trabajo)
    return guardar_docx(doc)


def agregar_anexo_control(doc, trabajo):
    doc.add_page_break()
    p = doc.add_paragraph()
    r = p.add_run("ANEXO — CONTROL DOCUMENTAL")
    r.bold = True
    r.font.size = Pt(13)
    r.font.color.rgb = RGBColor(0, 51, 102)

    columnas = [
        "ID", "Año", "Categoria_CV", "Componente_SNII",
        "Tipo_Producto_SNII", "Titulo_Actividad_o_Publicacion",
        "Estado_Probatorio",
    ]

    tabla = doc.add_table(rows=1, cols=len(columnas))
    tabla.style = "Table Grid"

    for i, columna in enumerate(columnas):
        tabla.rows[0].cells[i].text = columna

    for _, row in trabajo.iterrows():
        cells = tabla.add_row().cells
        for i, columna in enumerate(columnas):
            cells[i].text = obtener_valor(row, columna)


def guardar_docx(doc):
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# ============================================================
# FORMULARIO DE ACTIVIDAD
# ============================================================

def formulario_actividad(df, valores=None, key_prefix="nuevo"):
    """
    Formulario reutilizable para alta y edición.
    Solo el título es obligatorio.
    La clasificación SNII puede ser 'Actividad complementaria / no aplica'.
    """
    valores = valores or {}
    es_edicion = bool(valores)

    def v(col, default=""):
        return valores.get(col, default)

    with st.form(f"form_{key_prefix}", clear_on_submit=not es_edicion):
        st.markdown("### 1. Clasificación")

        c1, c2, c3 = st.columns(3)

        with c1:
            id_default = v("ID") or generar_id(df)
            registro_id = st.text_input(
                "ID",
                value=id_default,
                disabled=es_edicion,
            )

            anio_default = int(pd.to_numeric(v("Año"), errors="coerce")) if str(v("Año")).strip() else date.today().year
            opciones_anio = sorted(set(ANIOS_PROBATORIOS + [anio_default]), reverse=True)
            anio = st.selectbox(
                "Año",
                opciones_anio,
                index=opciones_anio.index(anio_default),
            )

            componente_default = v(
                "Componente_SNII",
                "Actividad complementaria / no aplica",
            )
            componente = st.selectbox(
                "Componente SNII",
                COMPONENTES_SNII,
                index=COMPONENTES_SNII.index(componente_default)
                if componente_default in COMPONENTES_SNII else len(COMPONENTES_SNII) - 1,
            )

        with c2:
            tipo_default = v("Tipo_Producto_SNII", "Otro")
            tipo_producto = st.selectbox(
                "Tipo de producto / actividad",
                TIPOS_PRODUCTO,
                index=TIPOS_PRODUCTO.index(tipo_default)
                if tipo_default in TIPOS_PRODUCTO else TIPOS_PRODUCTO.index("Otro"),
            )

            subtipo_default = v("Subtipo_SNII", "Actividad académica complementaria")
            subtipo = st.selectbox(
                "Subtipo",
                SUBTIPOS,
                index=SUBTIPOS.index(subtipo_default)
                if subtipo_default in SUBTIPOS else SUBTIPOS.index("Actividad académica complementaria"),
            )

            modalidad_default = v("Modalidad", "No aplica")
            modalidad = st.selectbox(
                "Modalidad",
                MODALIDADES,
                index=MODALIDADES.index(modalidad_default)
                if modalidad_default in MODALIDADES else 0,
            )

        with c3:
            categoria_default = v("Categoria_CV", "Otros / Sin clasificación")
            categoria = st.selectbox(
                "Categoría documental",
                CATEGORIAS,
                index=CATEGORIAS.index(categoria_default)
                if categoria_default in CATEGORIAS else CATEGORIAS.index("Otros / Sin clasificación"),
            )

            rol = st.text_input("Rol / participación", value=v("Rol_Participacion"))

            fecha_default = pd.to_datetime(v("Fecha"), errors="coerce")
            if pd.isna(fecha_default):
                fecha_default = date(anio_default, 1, 1)
            else:
                fecha_default = fecha_default.date()
            fecha = st.date_input("Fecha", value=fecha_default)

        st.markdown("### 2. Información académica")

        titulo = st.text_input(
            "Título de la actividad o publicación *",
            value=v("Titulo_Actividad_o_Publicacion"),
        )
        evento = st.text_input("Evento / Revista / Libro", value=v("Evento_Revista_Libro"))
        institucion = st.text_input("Institución / organización", value=v("Institucion_Organizacion"))
        lugar = st.text_input("Lugar / sede", value=v("Lugar_Sede"))
        autores = st.text_input("Autores", value=v("Autores"))
        coautores = st.text_input("Coautores", value=v("Coautores"))
        nivel_formacion = st.text_input("Nivel de formación relacionado", value=v("Nivel_Formacion"))
        estudiantes = st.text_input("Estudiantes beneficiados / dirigidos", value=v("Estudiantes_Beneficiados"))
        proyecto = st.text_input("Proyecto / línea de investigación", value=v("Proyecto_Linea_Investigacion"))

        st.markdown("### 3. Caracterización")

        aportacion = st.text_area("Descripción de la aportación", value=v("Descripcion_Aportacion"))
        relevancia = st.text_area("Relevancia / pertinencia", value=v("Relevancia_Pertinencia"))
        impacto = st.text_area("Impacto / beneficio social", value=v("Impacto_Beneficio_Social"))

        caracteristicas_previas = [
            x.strip() for x in v("Caracteristicas_SNII").split(";") if x.strip()
        ]
        caracteristicas = st.multiselect(
            "Características SNII",
            CARACTERISTICAS_SNII,
            default=[x for x in caracteristicas_previas if x in CARACTERISTICAS_SNII],
        )

        st.markdown("### 4. Información bibliográfica")

        c1, c2, c3 = st.columns(3)
        with c1:
            opciones_arbitrado = ["No aplica", "Sí", "No", "No especificado"]
            arbitrado_default = v("Arbitrado", "No aplica")
            arbitrado = st.selectbox(
                "¿Arbitrado / pares?",
                opciones_arbitrado,
                index=opciones_arbitrado.index(arbitrado_default)
                if arbitrado_default in opciones_arbitrado else 0,
            )

            opciones_publicado = [
                "No aplica", "Publicado", "Aceptado", "En prensa", "No publicado"
            ]
            publicado_default = v("Publicado", "No aplica")
            publicado = st.selectbox(
                "Estado",
                opciones_publicado,
                index=opciones_publicado.index(publicado_default)
                if publicado_default in opciones_publicado else 0,
            )

        with c2:
            revista = st.text_input("Revista / editorial", value=v("Revista_Editorial"))
            volumen = st.text_input("Volumen / número", value=v("Volumen_Numero"))
            paginas = st.text_input("Páginas", value=v("Paginas"))

        with c3:
            isbn = st.text_input("ISBN / ISSN", value=v("ISBN_ISSN"))
            doi = st.text_input("DOI / URL", value=v("DOI_URL"))

        st.markdown("### 5. Probatorio")

        archivo = st.file_uploader(
            "Subir probatorio nuevo (opcional)",
            type=["pdf", "png", "jpg", "jpeg"],
            key=f"archivo_{key_prefix}",
        )

        estados = [
            "Verificado / En Drive",
            "Pendiente de verificar",
            "Pendiente de escanear",
            "En trámite",
            "Sin probatorio",
        ]
        estado_default = v("Estado_Probatorio", "Sin probatorio")
        estado = st.selectbox(
            "Estado del probatorio",
            estados,
            index=estados.index(estado_default)
            if estado_default in estados else 0,
        )

        if es_edicion and v("Nombre_Archivo_PDF"):
            st.caption(f"Probatorio actual: {v('Nombre_Archivo_PDF')}")

        st.markdown("### 6. Inclusión en CV")

        c1, c2 = st.columns(2)
        with c1:
            incluir_cv = st.radio(
                "¿Incluir en CV general?",
                ["Sí", "No"],
                index=0 if str(v("Incluir_en_CV", "Sí")).lower() in {"sí", "si"} else 1,
                horizontal=True,
            )

        with c2:
            incluir_snii = st.radio(
                "¿Incluir en CV SNII?",
                ["Sí", "No"],
                index=0 if str(v("Incluir_en_CV_SNII", "No")).lower() in {"sí", "si"} else 1,
                horizontal=True,
            )

        notas = st.text_area("Notas / observaciones", value=v("Notas_Observaciones"))

        guardar = st.form_submit_button(
            "💾 Guardar cambios" if es_edicion else "💾 Guardar actividad"
        )

    if not guardar:
        return None

    if not titulo.strip():
        st.error("⚠️ El título es el único campo obligatorio.")
        return None

    fila = {
        "ID": registro_id,
        "Año": anio,
        "Fecha": str(fecha),
        "Componente_SNII": componente,
        "Tipo_Producto_SNII": tipo_producto,
        "Subtipo_SNII": subtipo,
        "Categoria_CV": categoria,
        "Rol_Participacion": rol,
        "Titulo_Actividad_o_Publicacion": titulo,
        "Evento_Revista_Libro": evento,
        "Institucion_Organizacion": institucion,
        "Lugar_Sede": lugar,
        "Modalidad": modalidad,
        "Autores": autores,
        "Coautores": coautores,
        "Nivel_Formacion": nivel_formacion,
        "Estudiantes_Beneficiados": estudiantes,
        "Proyecto_Linea_Investigacion": proyecto,
        "Descripcion_Aportacion": aportacion,
        "Relevancia_Pertinencia": relevancia,
        "Impacto_Beneficio_Social": impacto,
        "Caracteristicas_SNII": "; ".join(caracteristicas),
        "Arbitrado": arbitrado,
        "Publicado": publicado,
        "Revista_Editorial": revista,
        "Volumen_Numero": volumen,
        "Paginas": paginas,
        "ISBN_ISSN": isbn,
        "DOI_URL": doi,
        "Estado_Probatorio": estado,
        "Incluir_en_CV": incluir_cv,
        "Incluir_en_CV_SNII": incluir_snii,
        "Notas_Observaciones": notas,
    }

    # Conserva el probatorio anterior durante una edición.
    for col in ["Nombre_Archivo_PDF", "Enlace_Drive_Probatorio", "ID_Drive_Probatorio"]:
        fila[col] = v(col, "")

    fila["Redaccion_CV"] = redactar_registro(pd.Series(fila))

    return fila, archivo


# ============================================================
# OPERACIONES CRUD
# ============================================================

def guardar_nueva_actividad(service, estructura_drive, excel_id, df, fila, archivo):
    fila = fila.copy()

    if archivo:
        extension = os.path.splitext(archivo.name)[1].lower()
        titulo_limpio = limpiar_nombre_archivo(fila["Titulo_Actividad_o_Publicacion"])
        nombre_archivo = (
            f"{fila['Año']}_{fila['Categoria_CV'].replace(' ', '_')}_"
            f"{titulo_limpio}{extension}"
        )

        carpeta = obtener_carpeta_probatorio(
            service,
            estructura_drive,
            fila["Año"],
            fila["Categoria_CV"],
        )
        if not carpeta:
            st.error("No se encontró la carpeta de destino.")
            return False

        enlace, drive_id = subir_a_google_drive(
            service,
            nombre_archivo,
            archivo.getvalue(),
            carpeta,
        )

        if not drive_id:
            return False

        fila["Nombre_Archivo_PDF"] = nombre_archivo
        fila["Enlace_Drive_Probatorio"] = enlace
        fila["ID_Drive_Probatorio"] = drive_id

    # Compatibilidad con cualquier columna adicional de la base.
    for columna in df.columns:
        if columna not in fila:
            fila[columna] = ""

    df_nuevo = pd.concat([df, pd.DataFrame([fila])], ignore_index=True)
    actualizar_excel_drive(service, excel_id, df_nuevo)
    return True


def actualizar_actividad(service, estructura_drive, excel_id, df, indice, fila, archivo):
    df_nuevo = df.copy()
    fila = fila.copy()

    old_drive_id = limpiar_texto(df.loc[indice, "ID_Drive_Probatorio"])
    old_anio = limpiar_texto(df.loc[indice, "Año"])
    old_categoria = limpiar_texto(df.loc[indice, "Categoria_CV"])

    # Si se sube un nuevo probatorio, se reemplaza el anterior.
    if archivo:
        extension = os.path.splitext(archivo.name)[1].lower()
        titulo_limpio = limpiar_nombre_archivo(fila["Titulo_Actividad_o_Publicacion"])
        nombre_archivo = (
            f"{fila['Año']}_{fila['Categoria_CV'].replace(' ', '_')}_"
            f"{titulo_limpio}{extension}"
        )

        carpeta = obtener_carpeta_probatorio(
            service,
            estructura_drive,
            fila["Año"],
            fila["Categoria_CV"],
        )
        if not carpeta:
            st.error("No se encontró la carpeta de destino.")
            return False

        enlace, drive_id = subir_a_google_drive(
            service,
            nombre_archivo,
            archivo.getvalue(),
            carpeta,
        )
        if not drive_id:
            return False

        fila["Nombre_Archivo_PDF"] = nombre_archivo
        fila["Enlace_Drive_Probatorio"] = enlace
        fila["ID_Drive_Probatorio"] = drive_id

        if old_drive_id:
            eliminar_archivo_drive(service, old_drive_id)

    # Si se cambió de año/categoría, el archivo anterior se queda en
    # su carpeta original. No se mueve automáticamente para no arriesgar
    # documentos; el usuario puede subir uno nuevo si desea reubicarlo.

    for columna in df_nuevo.columns:
        if columna not in fila:
            fila[columna] = ""

    for columna in df_nuevo.columns:
        df_nuevo.at[indice, columna] = fila.get(columna, "")

    actualizar_excel_drive(service, excel_id, df_nuevo)
    return True


def eliminar_actividad(service, excel_id, df, indice, eliminar_probatorio=True):
    drive_id = limpiar_texto(df.loc[indice, "ID_Drive_Probatorio"])

    if eliminar_probatorio and drive_id:
        eliminar_archivo_drive(service, drive_id)

    df_nuevo = df.drop(index=indice).reset_index(drop=True)
    actualizar_excel_drive(service, excel_id, df_nuevo)
    return True


# ============================================================
# TABLERO
# ============================================================

def mostrar_tablero_snii(df):
    st.subheader("📊 Estado del expediente")

    total = len(df)
    if total == 0:
        st.info("Todavía no existen actividades registradas.")
        return

    verificados = (
        df["Estado_Probatorio"].astype(str).str.lower().str.contains("verificado").sum()
    )
    pendientes = total - verificados

    produccion = (df["Componente_SNII"] == COMPONENTES_SNII[0]).sum()
    comunidad = (df["Componente_SNII"] == COMPONENTES_SNII[1]).sum()
    divulgacion = (df["Componente_SNII"] == COMPONENTES_SNII[2]).sum()
    complementarias = (df["Componente_SNII"] == COMPONENTES_SNII[3]).sum()

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Actividades", total)
    c2.metric("Probatorios verificados", verificados)
    c3.metric("Pendientes", pendientes)
    c4.metric("Producción", produccion)
    c5.metric("Complementarias", complementarias)

    st.markdown("### Distribución por componente")
    resumen = pd.DataFrame({
        "Componente": ["Producción", "Comunidad", "Divulgación", "Complementarias"],
        "Registros": [produccion, comunidad, divulgacion, complementarias],
    })
    st.bar_chart(resumen.set_index("Componente"))


# ============================================================
# APP
# ============================================================

st.title("🎓 Sistema de Gestión de CV y Expediente SNII")
st.caption(
    "Dra. María Griselda Günther — expediente documental, control académico "
    "y generación de CV general y CV SNII"
)

service = obtener_servicio_drive()

if service:
    with st.spinner("☁️ Verificando estructura de Google Drive..."):
        estructura_drive = inicializar_estructura_drive(service)

    if not estructura_drive:
        st.error("No fue posible inicializar la estructura de Google Drive.")
        st.stop()

    excel_id, found_name = buscar_excel_en_drive(service)

    if not excel_id:
        st.warning("⚠️ No se encontró la base de datos en Google Drive.")

        archivo_excel_nuevo = st.file_uploader(
            "Sube Base_de_Datos_Probatorios_y_CV.xlsx",
            type=["xlsx"],
        )

        if archivo_excel_nuevo:
            metadata = {"name": "Base_de_Datos_Probatorios_y_CV.xlsx"}
            media = MediaIoBaseUpload(
                io.BytesIO(archivo_excel_nuevo.getvalue()),
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            service.files().create(
                body=metadata,
                media_body=media,
                fields="id",
            ).execute()

            st.success("Base de datos vinculada correctamente.")
            st.rerun()

    else:
        df = cargar_datos_drive(service, excel_id)

        (
            tab_dashboard,
            tab_consulta,
            tab_registro,
            tab_gestion,
            tab_cv,
        ) = st.tabs([
            "📊 Expediente",
            "🔍 Buscar",
            "➕ Nueva actividad",
            "✏️ Modificar / borrar",
            "📄 Generar CV",
        ])

        # ----------------------------------------------------
        # DASHBOARD
        # ----------------------------------------------------
        with tab_dashboard:
            mostrar_tablero_snii(df)
            st.markdown("---")
            st.info(
                "Ahora las actividades que no encajan en un rubro SNII pueden "
                "registrarse como 'Actividad complementaria / no aplica' y "
                "'Otros / Sin clasificación'. No es necesario llenar campos "
                "que no correspondan a la actividad."
            )

        # ----------------------------------------------------
        # BUSCADOR
        # ----------------------------------------------------
        with tab_consulta:
            st.subheader("🔍 Buscador de actividades y probatorios")

            c1, c2, c3, c4 = st.columns(4)

            with c1:
                anios = ["Todos"]
                valores = pd.to_numeric(df["Año"], errors="coerce").dropna().astype(int).unique()
                anios += sorted(valores.tolist(), reverse=True)
                filtro_anio = st.selectbox("Año", anios)

            with c2:
                filtro_componente = st.selectbox(
                    "Componente SNII",
                    ["Todos"] + COMPONENTES_SNII,
                )

            with c3:
                filtro_categoria = st.selectbox(
                    "Categoría",
                    ["Todas"] + CATEGORIAS,
                )

            with c4:
                filtro_texto = st.text_input(
                    "Buscar",
                    placeholder="Título, autor, revista, institución...",
                )

            resultado = df.copy()

            if filtro_anio != "Todos":
                resultado = resultado[
                    pd.to_numeric(resultado["Año"], errors="coerce") == int(filtro_anio)
                ]

            if filtro_componente != "Todos":
                resultado = resultado[
                    resultado["Componente_SNII"] == filtro_componente
                ]

            if filtro_categoria != "Todas":
                resultado = resultado[
                    resultado["Categoria_CV"] == filtro_categoria
                ]

            if filtro_texto:
                mascara = resultado.apply(
                    lambda row: row.astype(str).str.contains(
                        filtro_texto, case=False, na=False
                    ).any(),
                    axis=1,
                )
                resultado = resultado[mascara]

            st.write(f"**{len(resultado)} registros encontrados.**")

            columnas = [
                "ID", "Año", "Fecha", "Categoria_CV",
                "Componente_SNII", "Tipo_Producto_SNII",
                "Titulo_Actividad_o_Publicacion",
                "Institucion_Organizacion", "Estado_Probatorio",
                "Enlace_Drive_Probatorio",
            ]
            columnas = [c for c in columnas if c in resultado.columns]

            config = {}
            if "Enlace_Drive_Probatorio" in resultado.columns:
                config["Enlace_Drive_Probatorio"] = st.column_config.LinkColumn(
                    "Probatorio en Drive"
                )

            st.dataframe(
                resultado[columnas],
                use_container_width=True,
                column_config=config,
                hide_index=True,
            )

            st.caption(
                "Para modificar o borrar un registro usa la pestaña "
                "'✏️ Modificar / borrar'."
            )

        # ----------------------------------------------------
        # NUEVA ACTIVIDAD
        # ----------------------------------------------------
        with tab_registro:
            st.subheader("➕ Registrar nueva actividad académica")

            resultado_form = formulario_actividad(
                df,
                valores=None,
                key_prefix="nueva_actividad",
            )

            if resultado_form:
                fila, archivo = resultado_form

                with st.spinner("📚 Guardando actividad..."):
                    ok = guardar_nueva_actividad(
                        service,
                        estructura_drive,
                        excel_id,
                        df,
                        fila,
                        archivo,
                    )

                if ok:
                    st.success(
                        f"Actividad '{fila['Titulo_Actividad_o_Publicacion']}' "
                        "registrada correctamente."
                    )
                    st.rerun()

        # ----------------------------------------------------
        # MODIFICAR / BORRAR
        # ----------------------------------------------------
        with tab_gestion:
            st.subheader("✏️ Modificar o borrar actividades")

            if df.empty:
                st.info("No hay registros para modificar.")
            else:
                ids = df["ID"].astype(str).tolist()
                id_seleccionado = st.selectbox(
                    "Selecciona el ID del registro",
                    ids,
                    key="id_gestion",
                )

                indice = df.index[df["ID"].astype(str) == str(id_seleccionado)][0]
                registro = df.loc[indice].to_dict()

                st.markdown("#### Resumen del registro seleccionado")
                c1, c2, c3 = st.columns(3)
                c1.write(f"**Título:** {obtener_valor(df.loc[indice], 'Titulo_Actividad_o_Publicacion')}")
                c2.write(f"**Año:** {obtener_valor(df.loc[indice], 'Año')}")
                c3.write(f"**Probatorio:** {obtener_valor(df.loc[indice], 'Nombre_Archivo_PDF') or 'No registrado'}")

                st.markdown("---")
                st.markdown("### Editar registro")

                resultado_edicion = formulario_actividad(
                    df,
                    valores=registro,
                    key_prefix=f"editar_{id_seleccionado}",
                )

                if resultado_edicion:
                    fila_editada, archivo_nuevo = resultado_edicion

                    with st.spinner("✏️ Actualizando registro..."):
                        ok = actualizar_actividad(
                            service,
                            estructura_drive,
                            excel_id,
                            df,
                            indice,
                            fila_editada,
                            archivo_nuevo,
                        )

                    if ok:
                        st.success("Registro actualizado correctamente.")
                        st.rerun()

                st.markdown("---")
                st.markdown("### 🗑️ Eliminar registro")

                st.warning(
                    "Eliminar un registro lo quitará de la base de datos. "
                    "Si marcas la opción siguiente, también se eliminará "
                    "el probatorio correspondiente de Google Drive."
                )

                eliminar_pdf = st.checkbox(
                    "Eliminar también el probatorio de Google Drive",
                    value=True,
                    key=f"eliminar_pdf_{id_seleccionado}",
                )

                confirmar = st.checkbox(
                    "Confirmo que deseo eliminar este registro",
                    key=f"confirmar_eliminar_{id_seleccionado}",
                )

                if st.button(
                    "🗑️ Eliminar definitivamente",
                    type="secondary",
                    disabled=not confirmar,
                    key=f"boton_eliminar_{id_seleccionado}",
                ):
                    with st.spinner("Eliminando registro..."):
                        eliminar_actividad(
                            service,
                            excel_id,
                            df,
                            indice,
                            eliminar_probatorio=eliminar_pdf,
                        )

                    st.success("Registro eliminado correctamente.")
                    st.rerun()

        # ----------------------------------------------------
        # CV
        # ----------------------------------------------------
        with tab_cv:
            st.subheader("📄 Generador de dos versiones del CV")

            seleccionados_general = df[
                df["Incluir_en_CV"].astype(str).str.strip().str.lower() == "sí"
            ]

            seleccionados_snii = df[
                df["Incluir_en_CV_SNII"].astype(str).str.strip().str.lower() == "sí"
            ]

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("CV general", len(seleccionados_general))
            c2.metric("CV SNII", len(seleccionados_snii))
            c3.metric(
                "Probatorios verificados",
                df["Estado_Probatorio"].astype(str).str.lower().str.contains("verificado").sum(),
            )
            c4.metric(
                "Sin probatorio",
                df["ID_Drive_Probatorio"].astype(str).str.strip().eq("").sum(),
            )

            st.markdown("### 1. CV general")
            st.caption(
                "Usa 'Incluir en CV general = Sí'. No exige clasificación SNII."
            )

            if seleccionados_general.empty:
                st.info("No hay actividades marcadas para el CV general.")
            else:
                archivo_general = crear_cv_general_word(df)
                st.download_button(
                    "📥 Descargar CV general (.docx)",
                    data=archivo_general,
                    file_name="CV_General_Dra_Maria_Griselda_Gunther.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="descargar_cv_general",
                )

            st.markdown("---")
            st.markdown("### 2. CV SNII")
            st.caption(
                "Usa 'Incluir en CV SNII = Sí'. Las actividades complementarias "
                "también pueden aparecer sin forzarlas dentro de los tres componentes."
            )

            if seleccionados_snii.empty:
                st.info("No hay actividades marcadas para el CV SNII.")
            else:
                archivo_snii = crear_cv_snii_word(df)
                st.download_button(
                    "📥 Descargar CV SNII (.docx)",
                    data=archivo_snii,
                    file_name="CV_Academico_SNII_Dra_Maria_Griselda_Gunther.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="descargar_cv_snii",
                )

            st.markdown("---")
            st.markdown("### Vista previa de registros seleccionados")

            vista = df[
                df["Incluir_en_CV"].astype(str).str.strip().str.lower().eq("sí")
                | df["Incluir_en_CV_SNII"].astype(str).str.strip().str.lower().eq("sí")
            ].copy()

            columnas_preview = [
                "ID", "Año", "Categoria_CV", "Componente_SNII",
                "Tipo_Producto_SNII", "Titulo_Actividad_o_Publicacion",
                "Estado_Probatorio", "Incluir_en_CV", "Incluir_en_CV_SNII",
            ]
            columnas_preview = [c for c in columnas_preview if c in vista.columns]

            st.dataframe(
                vista[columnas_preview],
                use_container_width=True,
                hide_index=True,
            )
