import io
import os
import re
import zipfile
import mimetypes
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from googleapiclient.errors import HttpError

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor


# ============================================================
# SISTEMA DE GESTIÓN DE CV Y PROBATORIOS
# Arquitectura:
#   ACTIVIDADES 1 ---- N PROBATORIOS
#
# Cada probatorio conserva:
#   - nombre de archivo
#   - ID de Drive
#   - enlace directo a Drive
#   - actividad a la que pertenece
#
# La base se guarda en un único XLSX con dos hojas:
#   1) Actividades
#   2) Probatorios
# ============================================================


st.set_page_config(
    page_title="Control de CV y Probatorios - Dra. Günther",
    page_icon="📚",
    layout="wide",
)


# ============================================================
# CONFIGURACIÓN
# ============================================================

ROOT_FOLDER = "CV — Sistema de Gestión"
DB_FILENAME = "Base_de_Datos_Probatorios_y_CV.xlsx"

FOLDER_STRUCTURE = [
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

CATEGORIES = [
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

CURRENT_YEAR = datetime.now().year
YEARS = list(range(2020, CURRENT_YEAR + 2))

SCOPES = ["https://www.googleapis.com/auth/drive"]

ACTIVITY_COLUMNS = [
    "ID_Actividad",
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
    "Fecha_Alta",
    "Fecha_Actualización",
]

EVIDENCE_COLUMNS = [
    "ID_Probatorio",
    "ID_Actividad",
    "Nombre_Archivo_PDF",
    "Enlace_Drive_Probatorio",
    "ID_Drive_Probatorio",
    "Año_Drive",
    "Categoría_Drive",
    "Fecha_Alta",
    "Fecha_Actualización",
    "Estado_Archivo",
]

# Columnas antiguas que no deben sobrevivir a la migración.
LEGACY_COLUMNS = [
    "Componente_SNII",
    "Componente SNII",
    "Tipo_Producto_SNII",
    "Tipo producto SNII",
    "Categoría_CV",
    "Categoría CV",
    "Rol_Participación",
    "Rol participacion",
    "Título_Actividad_o_Publicación",
    "Titulo actividad o publicacion",
    "Evento_Revista_Libro",
    "Evento revista libro",
    "Institución_Organización",
    "Institucion organizacion",
    "Modalidad",
    "Autores",
    "Coautores",
    "Nivel_Formación",
    "Nivel formacion",
    "Estudiantes_Beneficiados",
    "Estudiantes beneficiados",
    "Proyecto_Línea_Investigación",
    "Proyecto linea investigacion",
    "Descripción_Aportación",
    "Descripcion aportacion",
    "Impacto_Beneficio_Social",
    "Impacto beneficio social",
    "Características_SNII",
    "Caracteristicas SNII",
    "Arbitrado",
    "Publicado",
    "Revista_Editorial",
    "Revista editorial",
    "Volumen_Número",
    "Volumen numero",
    "Páginas",
    "Paginas",
    "ISBN_ISSN",
    "ISBN ISSN",
    "DOI_URL",
    "DOI URL",
    "Incluir_en_CV_SNII",
    "Incluir en CV SNII",
    "Redacción",
    "Redaccion",
    "Nombre_Archivo_PDF",
    "Enlace_Drive_Probatorio",
    "ID_Drive_Probatorio",
]

COLUMN_ALIASES = {
    "ID": "ID_Actividad",
    "ID Actividad": "ID_Actividad",
    "Año": "Año",
    "Fecha": "Fecha",
    "Categoría": "Categoría",
    "Categoría_CV": "Categoría",
    "Categoría CV": "Categoría",
    "Rol": "Rol",
    "Rol_Participación": "Rol",
    "Rol participacion": "Rol",
    "Título": "Título",
    "Titulo": "Título",
    "Título_Actividad_o_Publicación": "Título",
    "Titulo actividad o publicacion": "Título",
    "Institución": "Institución",
    "Institución_Organización": "Institución",
    "Institucion organizacion": "Institución",
    "Lugar": "Lugar",
    "Estado_Probatorio": "Estado_Probatorio",
    "Incluir_en_CV": "Incluir_en_CV",
    "Incluir en CV": "Incluir_en_CV",
    "Notas_Observaciones": "Notas_Observaciones",
    "Notas / Observaciones": "Notas_Observaciones",
}


# ============================================================
# UTILIDADES GENERALES
# ============================================================

def clean_text(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()


def safe_filename(text):
    text = clean_text(text)
    text = re.sub(r'[\\/:*?"<>|]+', "_", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text[:150] or "archivo"


def now_string():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def parse_date(value):
    if not clean_text(value):
        return datetime.now().date()
    try:
        return pd.to_datetime(value).date()
    except Exception:
        return datetime.now().date()


def next_id(prefix, values, width=4):
    nums = []
    pattern = re.compile(rf"^{re.escape(prefix)}-(\d+)$", re.I)
    for value in values:
        match = pattern.match(clean_text(value))
        if match:
            nums.append(int(match.group(1)))
    return f"{prefix}-{max(nums, default=0) + 1:0{width}d}"


def get_row_value(row, column, default=""):
    if column not in row.index:
        return default
    return clean_text(row[column]) or default


def normalize_columns(df, required_columns):
    df = df.copy()
    for column in required_columns:
        if column not in df.columns:
            df[column] = ""
    return df[required_columns]


# ============================================================
# GOOGLE DRIVE
# ============================================================

@st.cache_resource
def get_drive_service():
    """
    Usa token.json si existe. Si no existe, intenta OAuth con credentials.json.
    """
    creds = None
    token_path = Path("token.json")
    credentials_path = Path("credentials.json")

    if token_path.exists():
        try:
            creds = Credentials.from_authorized_user_file(
                str(token_path),
                SCOPES,
            )
        except Exception:
            creds = None

    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
            token_path.write_text(creds.to_json(), encoding="utf-8")
        except Exception:
            creds = None

    if not creds or not creds.valid:
        if not credentials_path.exists():
            st.error(
                "No se encontró credentials.json. "
                "Coloca tus credenciales OAuth de Google en la carpeta de la aplicación."
            )
            return None

        try:
            flow = InstalledAppFlow.from_client_secrets_file(
                str(credentials_path),
                SCOPES,
            )
            creds = flow.run_local_server(port=0)
            token_path.write_text(creds.to_json(), encoding="utf-8")
        except Exception as exc:
            st.error(f"No fue posible autenticar Google Drive: {exc}")
            return None

    try:
        return build("drive", "v3", credentials=creds)
    except Exception as exc:
        st.error(f"No fue posible crear el servicio de Google Drive: {exc}")
        return None


def drive_find_folder(service, name, parent_id=None):
    escaped = name.replace("'", "\\'")
    query = (
        f"name = '{escaped}' "
        "and mimeType = 'application/vnd.google-apps.folder' "
        "and trashed = false"
    )
    if parent_id:
        query += f" and '{parent_id}' in parents"
    else:
        query += " and 'root' in parents"

    try:
        result = service.files().list(
            q=query,
            spaces="drive",
            fields="files(id,name,webViewLink,parents)",
            pageSize=10,
        ).execute()
        files = result.get("files", [])
        return files[0] if files else None
    except HttpError as exc:
        st.error(f"Error buscando carpeta '{name}': {exc}")
        return None


def drive_create_folder(service, name, parent_id=None):
    metadata = {
        "name": name,
        "mimeType": "application/vnd.google-apps.folder",
    }
    if parent_id:
        metadata["parents"] = [parent_id]

    try:
        return service.files().create(
            body=metadata,
            fields="id,name,webViewLink,parents",
        ).execute()
    except HttpError as exc:
        st.error(f"Error creando carpeta '{name}': {exc}")
        return None


def drive_get_or_create_folder(service, name, parent_id=None):
    folder = drive_find_folder(service, name, parent_id)
    return folder or drive_create_folder(service, name, parent_id)


@st.cache_resource
def build_drive_structure(_service):
    structure = {}

    root = drive_get_or_create_folder(_service, ROOT_FOLDER)
    if not root:
        return {}

    structure["root"] = root

    for folder_name in FOLDER_STRUCTURE:
        folder = drive_get_or_create_folder(
            _service,
            folder_name,
            root["id"],
        )
        if folder:
            structure[folder_name] = folder

    evidence_root = structure.get("02 — Probatorios")
    if not evidence_root:
        return structure

    structure["evidence_root"] = evidence_root

    for year in YEARS:
        year_folder = drive_get_or_create_folder(
            _service,
            str(year),
            evidence_root["id"],
        )
        if not year_folder:
            continue

        structure[f"year_{year}"] = year_folder

        for category in CATEGORIES:
            category_folder = drive_get_or_create_folder(
                _service,
                category,
                year_folder["id"],
            )
            if category_folder:
                structure[f"evidence_{year}_{category}"] = category_folder

    return structure


def get_evidence_folder(service, structure, year, category):
    key = f"evidence_{year}_{category}"
    if key in structure:
        return structure[key]

    evidence_root = structure.get("evidence_root")
    if not evidence_root:
        return None

    year_folder = drive_get_or_create_folder(
        service,
        str(year),
        evidence_root["id"],
    )
    if not year_folder:
        return None

    category_folder = drive_get_or_create_folder(
        service,
        category,
        year_folder["id"],
    )
    return category_folder


def drive_upload_file(service, filename, content, folder):
    if not folder:
        return None

    metadata = {
        "name": filename,
        "parents": [folder["id"]],
    }

    media = MediaIoBaseUpload(
        io.BytesIO(content),
        mimetype=mimetypes.guess_type(filename)[0] or "application/octet-stream",
        resumable=True,
    )

    try:
        result = service.files().create(
            body=metadata,
            media_body=media,
            fields="id,name,webViewLink,webContentLink,parents,mimeType",
        ).execute()

        return result
    except HttpError as exc:
        st.error(f"Error subiendo '{filename}': {exc}")
        return None


def drive_download_file(service, file_id):
    try:
        metadata = service.files().get(
            fileId=file_id,
            fields="id,name,mimeType,webViewLink,parents",
        ).execute()

        request = service.files().get_media(fileId=file_id)
        buffer = io.BytesIO()
        downloader = MediaIoBaseDownload(buffer, request)

        done = False
        while not done:
            _, done = downloader.next_chunk()

        buffer.seek(0)
        return metadata, buffer.getvalue()
    except HttpError as exc:
        st.warning(f"No se pudo descargar el archivo de Drive {file_id}: {exc}")
        return None, None


def drive_trash_file(service, file_id):
    if not clean_text(file_id):
        return False
    try:
        service.files().update(
            fileId=file_id,
            body={"trashed": True},
        ).execute()
        return True
    except HttpError as exc:
        st.warning(f"No se pudo enviar a papelera {file_id}: {exc}")
        return False


def drive_move_file(service, file_id, destination_folder):
    if not file_id or not destination_folder:
        return False

    try:
        metadata = service.files().get(
            fileId=file_id,
            fields="parents",
        ).execute()

        old_parents = metadata.get("parents", [])
        kwargs = {
            "fileId": file_id,
            "addParents": destination_folder["id"],
            "fields": "id,parents",
        }
        if old_parents:
            kwargs["removeParents"] = ",".join(old_parents)

        service.files().update(**kwargs).execute()
        return True
    except HttpError as exc:
        st.warning(f"No se pudo mover el archivo {file_id}: {exc}")
        return False


def drive_rename_file(service, file_id, new_name):
    try:
        service.files().update(
            fileId=file_id,
            body={"name": new_name},
        ).execute()
        return True
    except HttpError as exc:
        st.warning(f"No se pudo renombrar el archivo: {exc}")
        return False


# ============================================================
# BASE DE DATOS: LECTURA / ESCRITURA
# ============================================================

def find_database_file(service):
    escaped = DB_FILENAME.replace("'", "\\'")
    query = (
        f"name = '{escaped}' "
        "and trashed = false"
    )
    try:
        result = service.files().list(
            q=query,
            spaces="drive",
            fields="files(id,name,mimeType,webViewLink)",
            pageSize=10,
        ).execute()
        files = result.get("files", [])
        return files[0] if files else None
    except HttpError as exc:
        st.error(f"Error buscando la base de datos: {exc}")
        return None


def dataframe_to_xlsx_bytes(activities_df, evidence_df):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        activities_df.to_excel(
            writer,
            index=False,
            sheet_name="Actividades",
        )
        evidence_df.to_excel(
            writer,
            index=False,
            sheet_name="Probatorios",
        )

    output.seek(0)

    # Ajustes de legibilidad con openpyxl.
    wb = load_workbook(output)
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for col_idx, column_cells in enumerate(ws.iter_cols(), start=1):
            max_len = 0
            for cell in column_cells:
                value = clean_text(cell.value)
                max_len = max(max_len, len(value))
            ws.column_dimensions[get_column_letter(col_idx)].width = min(
                max(max_len + 2, 12),
                60,
            )

    final = io.BytesIO()
    wb.save(final)
    final.seek(0)
    return final.getvalue()


def upload_database(service, activities_df, evidence_df):
    content = dataframe_to_xlsx_bytes(
        activities_df,
        evidence_df,
    )

    existing = find_database_file(service)

    media = MediaIoBaseUpload(
        io.BytesIO(content),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=True,
    )

    try:
        if existing:
            service.files().update(
                fileId=existing["id"],
                media_body=media,
            ).execute()
            return existing["id"]

        metadata = {
            "name": DB_FILENAME,
            "parents": [ROOT_FOLDER_ID],
        }
        result = service.files().create(
            body=metadata,
            media_body=media,
            fields="id",
        ).execute()
        return result["id"]
    except HttpError as exc:
        st.error(f"No se pudo guardar la base de datos: {exc}")
        return None


def download_database_bytes(service, file_id):
    try:
        request = service.files().get_media(fileId=file_id)
        buffer = io.BytesIO()
        downloader = MediaIoBaseDownload(buffer, request)

        done = False
        while not done:
            _, done = downloader.next_chunk()

        buffer.seek(0)
        return buffer.getvalue()
    except HttpError as exc:
        st.error(f"No se pudo descargar la base de datos: {exc}")
        return None


def read_database(service, file_id):
    content = download_database_bytes(service, file_id)
    if not content:
        return empty_database()

    try:
        xls = pd.ExcelFile(io.BytesIO(content))
        sheets = xls.sheet_names

        activities = (
            pd.read_excel(
                io.BytesIO(content),
                sheet_name="Actividades",
            )
            if "Actividades" in sheets
            else pd.DataFrame()
        )

        evidence = (
            pd.read_excel(
                io.BytesIO(content),
                sheet_name="Probatorios",
            )
            if "Probatorios" in sheets
            else pd.DataFrame()
        )

        # Compatibilidad con bases anteriores.
        if activities.empty and "Base de datos" in sheets:
            old = pd.read_excel(
                io.BytesIO(content),
                sheet_name="Base de datos",
            )
            activities, evidence = migrate_legacy_database(old)

        activities = prepare_activities(activities)
        evidence = prepare_evidence(evidence)

        return activities, evidence

    except Exception as exc:
        st.error(f"No se pudo leer la base de datos: {exc}")
        return empty_database()


def empty_database():
    return (
        pd.DataFrame(columns=ACTIVITY_COLUMNS),
        pd.DataFrame(columns=EVIDENCE_COLUMNS),
    )


# ============================================================
# MIGRACIÓN DE LA BASE ANTERIOR
# ============================================================

def apply_aliases(df):
    df = df.copy()

    rename_map = {}
    for old_name, new_name in COLUMN_ALIASES.items():
        if old_name in df.columns and new_name not in df.columns:
            rename_map[old_name] = new_name

    if rename_map:
        df = df.rename(columns=rename_map)

    return df


def migrate_legacy_database(old_df):
    """
    Convierte el modelo antiguo:
      1 fila = actividad + un probatorio
    al modelo:
      Actividades
      Probatorios
    """
    old_df = apply_aliases(old_df)

    activities_rows = []
    evidence_rows = []

    activity_id_values = []
    evidence_id_values = []

    for _, row in old_df.iterrows():
        activity_id = clean_text(
            row.get("ID_Actividad")
            or row.get("ID")
        )

        if not activity_id:
            activity_id = next_id(
                "ACT",
                activity_id_values,
                width=4,
            )

        activity_id_values.append(activity_id)

        activity = {
            "ID_Actividad": activity_id,
            "Año": clean_text(row.get("Año")),
            "Fecha": clean_text(row.get("Fecha")),
            "Categoría": clean_text(row.get("Categoría")),
            "Rol": clean_text(row.get("Rol")),
            "Título": clean_text(row.get("Título")),
            "Institución": clean_text(row.get("Institución")),
            "Lugar": clean_text(row.get("Lugar")),
            "Estado_Probatorio": clean_text(
                row.get("Estado_Probatorio")
            ),
            "Incluir_en_CV": clean_text(
                row.get("Incluir_en_CV")
            ) or "No",
            "Notas_Observaciones": clean_text(
                row.get("Notas_Observaciones")
            ),
            "Fecha_Alta": now_string(),
            "Fecha_Actualización": now_string(),
        }

        activities_rows.append(activity)

        old_file = clean_text(row.get("Nombre_Archivo_PDF"))
        old_link = clean_text(row.get("Enlace_Drive_Probatorio"))
        old_drive_id = clean_text(row.get("ID_Drive_Probatorio"))

        # La versión anterior podía haber guardado varios archivos
        # en una celda. Se recuperan respetando el separador.
        names = [x.strip() for x in old_file.split(" || ")] if old_file else []
        links = [x.strip() for x in old_link.split(" || ")] if old_link else []
        ids = [x.strip() for x in old_drive_id.split(" || ")] if old_drive_id else []

        count = max(len(names), len(links), len(ids))

        for i in range(count):
            evidence_id = next_id(
                "PROB",
                evidence_id_values,
                width=5,
            )
            evidence_id_values.append(evidence_id)

            evidence_rows.append({
                "ID_Probatorio": evidence_id,
                "ID_Actividad": activity_id,
                "Nombre_Archivo_PDF": names[i] if i < len(names) else "",
                "Enlace_Drive_Probatorio": links[i] if i < len(links) else "",
                "ID_Drive_Probatorio": ids[i] if i < len(ids) else "",
                "Año_Drive": clean_text(row.get("Año")),
                "Categoría_Drive": clean_text(row.get("Categoría")),
                "Fecha_Alta": now_string(),
                "Fecha_Actualización": now_string(),
                "Estado_Archivo": "Registrado",
            })

    activities = prepare_activities(pd.DataFrame(activities_rows))
    evidence = prepare_evidence(pd.DataFrame(evidence_rows))
    return activities, evidence


def prepare_activities(df):
    df = apply_aliases(df)

    for column in ACTIVITY_COLUMNS:
        if column not in df.columns:
            df[column] = ""

    df = df[ACTIVITY_COLUMNS].copy()

    if df.empty:
        return df

    df["ID_Actividad"] = df["ID_Actividad"].astype(str).str.strip()
    df["Incluir_en_CV"] = (
        df["Incluir_en_CV"]
        .fillna("No")
        .astype(str)
        .str.strip()
    )

    return df


def prepare_evidence(df):
    for column in EVIDENCE_COLUMNS:
        if column not in df.columns:
            df[column] = ""

    df = df[EVIDENCE_COLUMNS].copy()

    if df.empty:
        return df

    for column in EVIDENCE_COLUMNS:
        df[column] = df[column].fillna("").astype(str).str.strip()

    return df


# ============================================================
# ÍNDICES / VALIDACIÓN
# ============================================================

def evidence_for_activity(evidence_df, activity_id):
    return evidence_df[
        evidence_df["ID_Actividad"].astype(str)
        == str(activity_id)
    ].copy()


def validate_database(activities_df, evidence_df):
    problems = []

    activity_ids = activities_df["ID_Actividad"].astype(str).tolist()
    duplicates = pd.Series(activity_ids).duplicated()
    if duplicates.any():
        dup_ids = pd.Series(activity_ids)[duplicates].tolist()
        problems.append(
            f"IDs de actividad duplicados: {', '.join(map(str, dup_ids))}"
        )

    evidence_ids = evidence_df["ID_Probatorio"].astype(str).tolist()
    duplicates_e = pd.Series(evidence_ids).duplicated()
    if duplicates_e.any():
        dup_ids = pd.Series(evidence_ids)[duplicates_e].tolist()
        problems.append(
            f"IDs de probatorio duplicados: {', '.join(map(str, dup_ids))}"
        )

    valid_activities = set(activity_ids)
    orphan = evidence_df[
        ~evidence_df["ID_Actividad"].astype(str).isin(valid_activities)
    ]

    if not orphan.empty:
        problems.append(
            f"Hay {len(orphan)} probatorio(s) sin actividad asociada."
        )

    no_drive_id = evidence_df[
        evidence_df["ID_Drive_Probatorio"].astype(str).str.strip() == ""
    ]

    if not no_drive_id.empty:
        problems.append(
            f"Hay {len(no_drive_id)} probatorio(s) sin ID de Drive."
        )

    no_link = evidence_df[
        evidence_df["Enlace_Drive_Probatorio"].astype(str).str.strip() == ""
    ]

    if not no_link.empty:
        problems.append(
            f"Hay {len(no_link)} probatorio(s) sin enlace de Drive."
        )

    return problems


# ============================================================
# OPERACIONES DE PROBATORIOS
# ============================================================

def generate_evidence_name(activity, index, original_name):
    extension = Path(original_name).suffix.lower() or ".pdf"

    year = clean_text(activity["Año"]) or str(CURRENT_YEAR)
    category = safe_filename(activity["Categoría"]).replace(" ", "_")
    title = safe_filename(activity["Título"])

    return (
        f"{year}_{category}_{title}"
        f"_Probatorio_{index}{extension}"
    )


def add_uploaded_evidence(
    service,
    structure,
    activities_df,
    evidence_df,
    activity_row,
    uploaded_files,
):
    activity_id = clean_text(activity_row["ID_Actividad"])

    current = evidence_for_activity(
        evidence_df,
        activity_id,
    )

    next_number = len(current) + 1

    new_rows = []

    folder = get_evidence_folder(
        service,
        structure,
        int(float(activity_row["Año"])),
        activity_row["Categoría"],
    )

    if not folder:
        st.error("No se encontró la carpeta de destino en Drive.")
        return evidence_df

    for file_obj in uploaded_files:
        filename = generate_evidence_name(
            activity_row,
            next_number,
            file_obj.name,
        )

        uploaded = drive_upload_file(
            service,
            filename,
            file_obj.getvalue(),
            folder,
        )

        if not uploaded:
            continue

        evidence_id = next_id(
            "PROB",
            evidence_df["ID_Probatorio"].tolist()
            + [row["ID_Probatorio"] for row in new_rows],
            width=5,
        )

        new_rows.append({
            "ID_Probatorio": evidence_id,
            "ID_Actividad": activity_id,
            "Nombre_Archivo_PDF": uploaded.get("name", filename),
            "Enlace_Drive_Probatorio": (
                uploaded.get("webViewLink")
                or f"https://drive.google.com/file/d/{uploaded['id']}/view"
            ),
            "ID_Drive_Probatorio": uploaded["id"],
            "Año_Drive": clean_text(activity_row["Año"]),
            "Categoría_Drive": clean_text(activity_row["Categoría"]),
            "Fecha_Alta": now_string(),
            "Fecha_Actualización": now_string(),
            "Estado_Archivo": "Activo",
        })

        next_number += 1

    if new_rows:
        evidence_df = pd.concat(
            [
                evidence_df,
                pd.DataFrame(new_rows),
            ],
            ignore_index=True,
        )

    return prepare_evidence(evidence_df)


def move_activity_evidence_if_needed(
    service,
    structure,
    evidence_df,
    activity_id,
    old_year,
    old_category,
    new_year,
    new_category,
):
    if (
        str(old_year) == str(new_year)
        and clean_text(old_category) == clean_text(new_category)
    ):
        return evidence_df

    destination = get_evidence_folder(
        service,
        structure,
        int(float(new_year)),
        new_category,
    )

    if not destination:
        return evidence_df

    mask = evidence_df["ID_Actividad"].astype(str) == str(activity_id)

    for idx in evidence_df[mask].index:
        drive_id = clean_text(
            evidence_df.at[idx, "ID_Drive_Probatorio"]
        )
        if drive_id:
            moved = drive_move_file(
                service,
                drive_id,
                destination,
            )
            if moved:
                evidence_df.at[idx, "Año_Drive"] = str(new_year)
                evidence_df.at[idx, "Categoría_Drive"] = new_category
                evidence_df.at[idx, "Fecha_Actualización"] = now_string()

    return evidence_df


def delete_evidence(
    service,
    evidence_df,
    evidence_id,
):
    row = evidence_df[
        evidence_df["ID_Probatorio"].astype(str)
        == str(evidence_id)
    ]

    if row.empty:
        return evidence_df

    drive_id = clean_text(
        row.iloc[0]["ID_Drive_Probatorio"]
    )

    if drive_id:
        drive_trash_file(
            service,
            drive_id,
        )

    return evidence_df[
        evidence_df["ID_Probatorio"].astype(str)
        != str(evidence_id)
    ].copy()


def delete_activity(
    service,
    activities_df,
    evidence_df,
    activity_id,
):
    related = evidence_for_activity(
        evidence_df,
        activity_id,
    )

    for _, row in related.iterrows():
        drive_id = clean_text(
            row["ID_Drive_Probatorio"]
        )
        if drive_id:
            drive_trash_file(
                service,
                drive_id,
            )

    activities_df = activities_df[
        activities_df["ID_Actividad"].astype(str)
        != str(activity_id)
    ].copy()

    evidence_df = evidence_df[
        evidence_df["ID_Actividad"].astype(str)
        != str(activity_id)
    ].copy()

    return activities_df, evidence_df


# ============================================================
# ZIP
# ============================================================

def build_zip(service, activities_df, evidence_df, selected_ids):
    output = io.BytesIO()
    count = 0
    failures = []

    with zipfile.ZipFile(
        output,
        "w",
        zipfile.ZIP_DEFLATED,
    ) as zf:

        selected = activities_df[
            activities_df["ID_Actividad"].astype(str).isin(
                [str(x) for x in selected_ids]
            )
        ]

        for _, activity in selected.iterrows():
            activity_id = clean_text(activity["ID_Actividad"])
            title = safe_filename(activity["Título"])

            related = evidence_for_activity(
                evidence_df,
                activity_id,
            )

            for _, evidence in related.iterrows():
                drive_id = clean_text(
                    evidence["ID_Drive_Probatorio"]
                )

                if not drive_id:
                    failures.append(
                        f"{activity_id}: {evidence['Nombre_Archivo_PDF']} sin ID de Drive"
                    )
                    continue

                metadata, content = drive_download_file(
                    service,
                    drive_id,
                )

                if not content:
                    failures.append(
                        f"{activity_id}: {evidence['Nombre_Archivo_PDF']}"
                    )
                    continue

                original_name = safe_filename(
                    metadata.get(
                        "name",
                        evidence["Nombre_Archivo_PDF"],
                    )
                )

                archive_path = (
                    f"{activity_id} - {title}/"
                    f"{original_name}"
                )

                zf.writestr(
                    archive_path,
                    content,
                )
                count += 1

    output.seek(0)
    return output.getvalue(), count, failures


# ============================================================
# CV WORD
# ============================================================

def build_cv_docx(activities_df):
    doc = Document()

    for section in doc.sections:
        section.top_margin = Inches(0.85)
        section.bottom_margin = Inches(0.85)
        section.left_margin = Inches(0.9)
        section.right_margin = Inches(0.9)

    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(10.5)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run(
        "DRA. MARÍA GRISELDA GÜNTHER"
    )
    run.bold = True
    run.font.size = Pt(17)
    run.font.color.rgb = RGBColor(31, 78, 121)

    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle.add_run(
        "CURRÍCULUM VITAE — SÍNTESIS DE ACTIVIDADES"
    )
    run.italic = True
    run.font.size = Pt(10)

    selected = activities_df[
        activities_df["Incluir_en_CV"].astype(str).str.lower().str.strip() == "sí"
    ].copy()

    if selected.empty:
        doc.add_paragraph(
            "No hay actividades marcadas para incluir en el CV."
        )
        return doc

    selected["Año_num"] = pd.to_numeric(
        selected["Año"],
        errors="coerce",
    )

    selected = selected.sort_values(
        ["Categoría", "Año_num"],
        ascending=[True, False],
    )

    for category in CATEGORIES:
        subset = selected[
            selected["Categoría"] == category
        ]

        if subset.empty:
            continue

        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)

        r = p.add_run(category)
        r.bold = True
        r.font.size = Pt(12.5)
        r.font.color.rgb = RGBColor(31, 78, 121)

        for _, row in subset.iterrows():
            p = doc.add_paragraph(style="List Bullet")
            title_text = clean_text(row["Título"])

            r = p.add_run(title_text)
            r.bold = True

            details = []

            if clean_text(row["Rol"]):
                details.append(f"Rol: {row['Rol']}")

            if clean_text(row["Institución"]):
                details.append(clean_text(row["Institución"]))

            if clean_text(row["Lugar"]):
                details.append(clean_text(row["Lugar"]))

            if clean_text(row["Fecha"]):
                details.append(
                    parse_date(row["Fecha"]).strftime("%d/%m/%Y")
                )

            if details:
                p.add_run(". " + " · ".join(details) + ".")

    return doc


# ============================================================
# EXPORTACIÓN DE VISTAS
# ============================================================

def create_selection_xlsx(activities_df, evidence_df, selected_ids):
    activities = activities_df[
        activities_df["ID_Actividad"].astype(str).isin(
            [str(x) for x in selected_ids]
        )
    ].copy()

    evidence = evidence_df[
        evidence_df["ID_Actividad"].astype(str).isin(
            [str(x) for x in selected_ids]
        )
    ].copy()

    return dataframe_to_xlsx_bytes(
        activities,
        evidence,
    )


# ============================================================
# INICIALIZACIÓN
# ============================================================

service = get_drive_service()

if not service:
    st.stop()

with st.spinner("Preparando estructura de Google Drive..."):
    structure = build_drive_structure(service)

if not structure:
    st.error("No fue posible preparar la estructura de Google Drive.")
    st.stop()

ROOT_FOLDER_ID = structure["root"]["id"]

database_file = find_database_file(service)

if database_file:
    activities_df, evidence_df = read_database(
        service,
        database_file["id"],
    )
else:
    activities_df, evidence_df = empty_database()

# Si la base es antigua, mostrar una migración explícita.
legacy_detected = (
    database_file is not None
    and "Actividades" not in (
        pd.ExcelFile(
            io.BytesIO(
                download_database_bytes(
                    service,
                    database_file["id"],
                )
            )
        ).sheet_names
    )
)

if legacy_detected:
    st.warning(
        "⚠️ La base de datos actual utiliza la estructura antigua. "
        "El sistema la ha convertido en memoria a la arquitectura "
        "Actividades → Probatorios."
    )

    if st.button(
        "🔄 Migrar y guardar definitivamente la nueva estructura",
        type="primary",
    ):
        upload_database(
            service,
            activities_df,
            evidence_df,
        )
        st.success(
            "Migración terminada. La base ahora tiene las hojas "
            "'Actividades' y 'Probatorios'."
        )
        st.rerun()

# ============================================================
# CABECERA
# ============================================================

st.title("📚 Sistema de Gestión de CV y Probatorios")
st.caption(
    "Arquitectura 1 actividad → N probatorios · Google Drive + Excel"
)

# ============================================================
# MÉTRICAS
# ============================================================

col1, col2, col3, col4 = st.columns(4)

with col1:
    st.metric(
        "Actividades",
        len(activities_df),
    )

with col2:
    st.metric(
        "Probatorios",
        len(evidence_df),
    )

with col3:
    activities_with_evidence = (
        evidence_df["ID_Actividad"].astype(str).nunique()
        if not evidence_df.empty
        else 0
    )
    st.metric(
        "Actividades con evidencia",
        activities_with_evidence,
    )

with col4:
    pending = len(
        activities_df[
            ~activities_df["ID_Actividad"].astype(str).isin(
                evidence_df["ID_Actividad"].astype(str)
            )
        ]
    )
    st.metric(
        "Sin probatorio",
        pending,
    )

# ============================================================
# NAVEGACIÓN
# ============================================================

(
    tab_search,
    tab_new,
    tab_edit,
    tab_evidence,
    tab_packages,
    tab_cv,
    tab_maintenance,
) = st.tabs(
    [
        "🔎 Buscar",
        "➕ Nueva actividad",
        "✏️ Editar actividad",
        "📎 Probatorios",
        "📦 Paquetes",
        "📄 Generar CV",
        "🛠️ Mantenimiento",
    ]
)


# ============================================================
# BUSCAR
# ============================================================

with tab_search:
    st.subheader("🔎 Buscador")

    c1, c2, c3 = st.columns(3)

    with c1:
        years_available = sorted(
            pd.to_numeric(
                activities_df["Año"],
                errors="coerce",
            )
            .dropna()
            .astype(int)
            .unique()
            .tolist(),
            reverse=True,
        )
        year_filter = st.selectbox(
            "Año",
            ["Todos"] + years_available,
        )

    with c2:
        category_filter = st.selectbox(
            "Categoría",
            ["Todas"] + CATEGORIES,
        )

    with c3:
        search_text = st.text_input(
            "Buscar",
            placeholder="ID, título, institución, rol...",
        )

    result = activities_df.copy()

    if year_filter != "Todos":
        result = result[
            pd.to_numeric(
                result["Año"],
                errors="coerce",
            ) == int(year_filter)
        ]

    if category_filter != "Todas":
        result = result[
            result["Categoría"] == category_filter
        ]

    if search_text:
        mask = result.apply(
            lambda row: row.astype(str)
            .str.contains(
                search_text,
                case=False,
                na=False,
                regex=False,
            ).any(),
            axis=1,
        )
        result = result[mask]

    st.write(f"**{len(result)} actividad(es)**")

    display = result.copy()
    display["N_Probatorios"] = display["ID_Actividad"].apply(
        lambda x: len(evidence_for_activity(evidence_df, x))
    )

    st.dataframe(
        display[
            [
                "ID_Actividad",
                "Año",
                "Fecha",
                "Categoría",
                "Rol",
                "Título",
                "Institución",
                "Estado_Probatorio",
                "Incluir_en_CV",
                "N_Probatorios",
            ]
        ],
        use_container_width=True,
        hide_index=True,
    )


# ============================================================
# NUEVA ACTIVIDAD
# ============================================================

with tab_new:
    st.subheader("➕ Nueva actividad")

    with st.form("new_activity_form"):
        c1, c2 = st.columns(2)

        with c1:
            new_id = st.text_input(
                "ID de actividad",
                value=next_id(
                    "ACT",
                    activities_df["ID_Actividad"].tolist(),
                    width=4,
                ),
            )

            new_year = st.number_input(
                "Año",
                min_value=1900,
                max_value=2100,
                value=CURRENT_YEAR,
            )

            new_date = st.date_input(
                "Fecha",
                value=datetime.now().date(),
            )

            new_category = st.selectbox(
                "Categoría",
                CATEGORIES,
            )

            new_role = st.text_input("Rol")

        with c2:
            new_title = st.text_input(
                "Título *",
            )

            new_institution = st.text_input(
                "Institución / Organización",
            )

            new_place = st.text_input(
                "Lugar / Sede",
            )

            new_status = st.selectbox(
                "Estado del probatorio",
                [
                    "Pendiente de probatorio",
                    "Verificado / En Drive",
                    "En trámite",
                ],
            )

            new_include = st.radio(
                "Incluir en CV",
                ["Sí", "No"],
                horizontal=True,
            )

        new_notes = st.text_area(
            "Notas / Observaciones",
        )

        new_files = st.file_uploader(
            "Probatorios — puedes seleccionar varios",
            type=["pdf", "png", "jpg", "jpeg"],
            accept_multiple_files=True,
        )

        submitted = st.form_submit_button(
            "💾 Crear actividad",
            type="primary",
        )

    if submitted:
        if not clean_text(new_title):
            st.error("El título es obligatorio.")
            st.stop()

        if new_id in activities_df["ID_Actividad"].astype(str).tolist():
            st.error(f"El ID {new_id} ya existe.")
            st.stop()

        activity = {
            "ID_Actividad": new_id.strip(),
            "Año": str(new_year),
            "Fecha": new_date.isoformat(),
            "Categoría": new_category,
            "Rol": new_role,
            "Título": new_title.strip(),
            "Institución": new_institution,
            "Lugar": new_place,
            "Estado_Probatorio": (
                "Verificado / En Drive"
                if new_files
                else new_status
            ),
            "Incluir_en_CV": new_include,
            "Notas_Observaciones": new_notes,
            "Fecha_Alta": now_string(),
            "Fecha_Actualización": now_string(),
        }

        new_activities = pd.concat(
            [
                activities_df,
                pd.DataFrame([activity]),
            ],
            ignore_index=True,
        )

        new_evidence = add_uploaded_evidence(
            service,
            structure,
            new_activities,
            evidence_df,
            activity,
            new_files or [],
        )

        if new_files:
            new_activities.loc[
                new_activities["ID_Actividad"].astype(str) == new_id,
                "Estado_Probatorio",
            ] = "Verificado / En Drive"

        new_activities.loc[
            new_activities["ID_Actividad"].astype(str) == new_id,
            "Fecha_Actualización",
        ] = now_string()

        upload_database(
            service,
            prepare_activities(new_activities),
            prepare_evidence(new_evidence),
        )

        st.success(
            f"Actividad {new_id} creada correctamente con "
            f"{len(new_files or [])} probatorio(s)."
        )
        st.rerun()


# ============================================================
# EDITAR ACTIVIDAD
# ============================================================

with tab_edit:
    st.subheader("✏️ Editar actividad")

    if activities_df.empty:
        st.info("No hay actividades.")
    else:
        edit_search = st.text_input(
            "Buscar actividad",
            placeholder="ACT-0001, título, institución...",
            key="edit_search",
        )

        candidates = activities_df.copy()

        if edit_search:
            mask = candidates.apply(
                lambda row: row.astype(str)
                .str.contains(
                    edit_search,
                    case=False,
                    na=False,
                    regex=False,
                ).any(),
                axis=1,
            )
            candidates = candidates[mask]

        if candidates.empty:
            st.warning("No se encontraron actividades.")
        else:
            selected_id = st.selectbox(
                "Actividad",
                candidates["ID_Actividad"].astype(str).tolist(),
                format_func=lambda x: (
                    f"{x} — "
                    f"{get_row_value(
                        candidates[candidates['ID_Actividad'].astype(str) == str(x)].iloc[0],
                        'Título',
                        'Sin título'
                    )}"
                ),
            )

            row_idx = activities_df[
                activities_df["ID_Actividad"].astype(str) == str(selected_id)
            ].index[0]

            row = activities_df.loc[row_idx].copy()
            related = evidence_for_activity(
                evidence_df,
                selected_id,
            )

            old_year = get_row_value(row, "Año")
            old_category = get_row_value(row, "Categoría")

            with st.form("edit_activity_form"):
                c1, c2 = st.columns(2)

                with c1:
                    edit_year = st.number_input(
                        "Año",
                        min_value=1900,
                        max_value=2100,
                        value=int(float(old_year or CURRENT_YEAR)),
                    )

                    edit_date = st.date_input(
                        "Fecha",
                        value=parse_date(row["Fecha"]),
                    )

                    edit_category = st.selectbox(
                        "Categoría",
                        CATEGORIES,
                        index=(
                            CATEGORIES.index(old_category)
                            if old_category in CATEGORIES
                            else 0
                        ),
                    )

                    edit_role = st.text_input(
                        "Rol",
                        value=get_row_value(row, "Rol"),
                    )

                with c2:
                    edit_title = st.text_input(
                        "Título *",
                        value=get_row_value(row, "Título"),
                    )

                    edit_institution = st.text_input(
                        "Institución",
                        value=get_row_value(row, "Institución"),
                    )

                    edit_place = st.text_input(
                        "Lugar / Sede",
                        value=get_row_value(row, "Lugar"),
                    )

                    status_options = [
                        "Pendiente de probatorio",
                        "Verificado / En Drive",
                        "En trámite",
                    ]

                    old_status = get_row_value(
                        row,
                        "Estado_Probatorio",
                        status_options[0],
                    )

                    edit_status = st.selectbox(
                        "Estado",
                        status_options,
                        index=(
                            status_options.index(old_status)
                            if old_status in status_options
                            else 0
                        ),
                    )

                old_include = get_row_value(
                    row,
                    "Incluir_en_CV",
                    "No",
                )

                edit_include = st.radio(
                    "Incluir en CV",
                    ["Sí", "No"],
                    index=0 if old_include.lower() == "sí" else 1,
                    horizontal=True,
                )

                edit_notes = st.text_area(
                    "Notas / Observaciones",
                    value=get_row_value(
                        row,
                        "Notas_Observaciones",
                    ),
                )

                save_edit = st.form_submit_button(
                    "💾 Guardar cambios",
                    type="primary",
                )

            if save_edit:
                activities_df.at[row_idx, "Año"] = str(edit_year)
                activities_df.at[row_idx, "Fecha"] = edit_date.isoformat()
                activities_df.at[row_idx, "Categoría"] = edit_category
                activities_df.at[row_idx, "Rol"] = edit_role
                activities_df.at[row_idx, "Título"] = edit_title.strip()
                activities_df.at[row_idx, "Institución"] = edit_institution
                activities_df.at[row_idx, "Lugar"] = edit_place
                activities_df.at[row_idx, "Estado_Probatorio"] = (
                    "Verificado / En Drive"
                    if not related.empty
                    else edit_status
                )
                activities_df.at[row_idx, "Incluir_en_CV"] = edit_include
                activities_df.at[row_idx, "Notas_Observaciones"] = edit_notes
                activities_df.at[row_idx, "Fecha_Actualización"] = now_string()

                evidence_df = move_activity_evidence_if_needed(
                    service,
                    structure,
                    evidence_df,
                    selected_id,
                    old_year,
                    old_category,
                    str(edit_year),
                    edit_category,
                )

                upload_database(
                    service,
                    prepare_activities(activities_df),
                    prepare_evidence(evidence_df),
                )

                st.success("Actividad actualizada.")
                st.rerun()

            st.divider()

            st.subheader(
                f"📎 Probatorios de {selected_id}"
            )

            if related.empty:
                st.info("Esta actividad todavía no tiene probatorios.")
            else:
                for _, evidence in related.iterrows():
                    c1, c2, c3 = st.columns([5, 2, 1])

                    with c1:
                        st.markdown(
                            f"**{evidence['Nombre_Archivo_PDF']}**"
                        )
                        st.caption(
                            f"ID: {evidence['ID_Probatorio']} · "
                            f"Drive ID: {evidence['ID_Drive_Probatorio']}"
                        )

                    with c2:
                        link = evidence["Enlace_Drive_Probatorio"]
                        if link:
                            st.link_button(
                                "🔗 Abrir",
                                link,
                            )

                    with c3:
                        if st.button(
                            "🗑️",
                            key=f"delete_evidence_{evidence['ID_Probatorio']}",
                        ):
                            evidence_df = delete_evidence(
                                service,
                                evidence_df,
                                evidence["ID_Probatorio"],
                            )

                            activities_df.loc[
                                activities_df["ID_Actividad"].astype(str) == selected_id,
                                "Estado_Probatorio",
                            ] = (
                                "Verificado / En Drive"
                                if not evidence_for_activity(
                                    evidence_df,
                                    selected_id,
                                ).empty
                                else "Pendiente de probatorio"
                            )

                            upload_database(
                                service,
                                prepare_activities(activities_df),
                                prepare_evidence(evidence_df),
                            )

                            st.success("Probatorio enviado a la papelera de Drive.")
                            st.rerun()

            st.divider()

            st.subheader("➕ Agregar más probatorios")

            more_files = st.file_uploader(
                "Selecciona uno o varios archivos",
                type=["pdf", "png", "jpg", "jpeg"],
                accept_multiple_files=True,
                key=f"more_files_{selected_id}",
            )

            if st.button(
                "⬆️ Subir probatorios",
                type="primary",
                disabled=not more_files,
            ):
                evidence_df = add_uploaded_evidence(
                    service,
                    structure,
                    activities_df,
                    evidence_df,
                    row,
                    more_files,
                )

                activities_df.loc[
                    activities_df["ID_Actividad"].astype(str) == selected_id,
                    "Estado_Probatorio",
                ] = "Verificado / En Drive"

                activities_df.loc[
                    activities_df["ID_Actividad"].astype(str) == selected_id,
                    "Fecha_Actualización",
                ] = now_string()

                upload_database(
                    service,
                    prepare_activities(activities_df),
                    prepare_evidence(evidence_df),
                )

                st.success(
                    f"{len(more_files)} probatorio(s) agregado(s)."
                )
                st.rerun()

            st.divider()

            if st.button(
                "🗑️ Eliminar actividad completa",
                type="secondary",
            ):
                st.session_state[
                    f"confirm_delete_{selected_id}"
                ] = True

            if st.session_state.get(
                f"confirm_delete_{selected_id}",
                False,
            ):
                st.warning(
                    "Esto eliminará la actividad de la base y enviará "
                    "todos sus probatorios a la papelera de Drive."
                )

                c1, c2 = st.columns(2)

                with c1:
                    if st.button(
                        "Sí, eliminar definitivamente",
                        type="primary",
                    ):
                        activities_df, evidence_df = delete_activity(
                            service,
                            activities_df,
                            evidence_df,
                            selected_id,
                        )

                        upload_database(
                            service,
                            prepare_activities(activities_df),
                            prepare_evidence(evidence_df),
                        )

                        st.session_state[
                            f"confirm_delete_{selected_id}"
                        ] = False

                        st.success("Actividad eliminada.")
                        st.rerun()

                with c2:
                    if st.button("Cancelar"):
                        st.session_state[
                            f"confirm_delete_{selected_id}"
                        ] = False
                        st.rerun()


# ============================================================
# ADMINISTRACIÓN DE PROBATORIOS
# ============================================================

with tab_evidence:
    st.subheader("📎 Inventario de probatorios")

    if evidence_df.empty:
        st.info("No hay probatorios registrados.")
    else:
        evidence_view = evidence_df.merge(
            activities_df[
                [
                    "ID_Actividad",
                    "Título",
                    "Categoría",
                    "Año",
                ]
            ],
            on="ID_Actividad",
            how="left",
        )

        search_evidence = st.text_input(
            "Buscar probatorio",
            placeholder="Nombre, ID, actividad, título...",
        )

        if search_evidence:
            mask = evidence_view.apply(
                lambda row: row.astype(str)
                .str.contains(
                    search_evidence,
                    case=False,
                    na=False,
                    regex=False,
                ).any(),
                axis=1,
            )
            evidence_view = evidence_view[mask]

        st.dataframe(
            evidence_view[
                [
                    "ID_Probatorio",
                    "ID_Actividad",
                    "Título",
                    "Año",
                    "Categoría",
                    "Nombre_Archivo_PDF",
                    "Enlace_Drive_Probatorio",
                    "ID_Drive_Probatorio",
                    "Estado_Archivo",
                ]
            ],
            use_container_width=True,
            hide_index=True,
            column_config={
                "Enlace_Drive_Probatorio": st.column_config.LinkColumn(
                    "Enlace Drive"
                )
            },
        )

        st.download_button(
            "📥 Descargar inventario de probatorios (.xlsx)",
            data=dataframe_to_xlsx_bytes(
                activities_df,
                evidence_view[
                    EVIDENCE_COLUMNS
                ],
            ),
            file_name="Inventario_Probatorios.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# ============================================================
# PAQUETES
# ============================================================

with tab_packages:
    st.subheader("📦 Generador de paquetes")

    if activities_df.empty:
        st.info("No hay actividades.")
    else:
        c1, c2, c3 = st.columns(3)

        with c1:
            package_year = st.selectbox(
                "Año",
                ["Todos"]
                + sorted(
                    pd.to_numeric(
                        activities_df["Año"],
                        errors="coerce",
                    )
                    .dropna()
                    .astype(int)
                    .unique()
                    .tolist(),
                    reverse=True,
                ),
            )

        with c2:
            package_category = st.selectbox(
                "Categoría",
                ["Todas"] + CATEGORIES,
            )

        with c3:
            package_search = st.text_input(
                "Texto",
            )

        package_df = activities_df.copy()

        if package_year != "Todos":
            package_df = package_df[
                pd.to_numeric(
                    package_df["Año"],
                    errors="coerce",
                ) == int(package_year)
            ]

        if package_category != "Todas":
            package_df = package_df[
                package_df["Categoría"] == package_category
            ]

        if package_search:
            mask = package_df.apply(
                lambda row: row.astype(str)
                .str.contains(
                    package_search,
                    case=False,
                    na=False,
                    regex=False,
                ).any(),
                axis=1,
            )
            package_df = package_df[mask]

        package_ids = st.multiselect(
            "Actividades a incluir",
            package_df["ID_Actividad"].astype(str).tolist(),
            format_func=lambda x: (
                f"{x} — "
                f"{get_row_value(
                    package_df[
                        package_df['ID_Actividad'].astype(str) == str(x)
                    ].iloc[0],
                    'Título'
                )}"
            ),
        )

        if package_ids:
            count_evidence = sum(
                len(evidence_for_activity(evidence_df, x))
                for x in package_ids
            )

            st.info(
                f"{len(package_ids)} actividad(es) · "
                f"{count_evidence} probatorio(s)"
            )

            if st.button(
                "📦 Crear ZIP",
                type="primary",
            ):
                with st.spinner(
                    "Construyendo ZIP desde Google Drive..."
                ):
                    zip_bytes, count, failures = build_zip(
                        service,
                        activities_df,
                        evidence_df,
                        package_ids,
                    )

                if count:
                    stamp = datetime.now().strftime("%Y%m%d_%H%M")
                    st.download_button(
                        "📥 Descargar ZIP",
                        data=zip_bytes,
                        file_name=f"Paquete_Probatorios_{stamp}.zip",
                        mime="application/zip",
                        type="primary",
                    )

                    st.success(
                        f"ZIP listo: {count} archivo(s)."
                    )

                if failures:
                    st.warning(
                        "Algunos archivos no pudieron incluirse:"
                    )
                    for failure in failures:
                        st.write(f"- {failure}")

            st.divider()

            if st.button(
                "📊 Exportar selección a Excel",
            ):
                xlsx = create_selection_xlsx(
                    activities_df,
                    evidence_df,
                    package_ids,
                )

                st.download_button(
                    "📥 Descargar Excel de la selección",
                    data=xlsx,
                    file_name="Seleccion_CV_Probatorios.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        else:
            st.info("Selecciona al menos una actividad.")


# ============================================================
# CV
# ============================================================

with tab_cv:
    st.subheader("📄 Generación de CV")

    cv_df = activities_df[
        activities_df["Incluir_en_CV"].astype(str).str.lower().str.strip() == "sí"
    ].copy()

    st.metric(
        "Actividades marcadas para CV",
        len(cv_df),
    )

    if not cv_df.empty:
        doc = build_cv_docx(
            activities_df,
        )

        output = io.BytesIO()
        doc.save(output)
        output.seek(0)

        st.download_button(
            "📥 Descargar CV Word",
            data=output.getvalue(),
            file_name="CV_Dra_Maria_Griselda_Gunther.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
        )

        st.dataframe(
            cv_df[
                [
                    "ID_Actividad",
                    "Año",
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
            "No hay actividades marcadas para incluir en el CV."
        )


# ============================================================
# MANTENIMIENTO
# ============================================================

with tab_maintenance:
    st.subheader("🛠️ Mantenimiento y control de integridad")

    problems = validate_database(
        activities_df,
        evidence_df,
    )

    if problems:
        st.warning(
            "Se encontraron observaciones:"
        )
        for problem in problems:
            st.write(f"- {problem}")
    else:
        st.success(
            "✅ La estructura de la base no presenta inconsistencias "
            "detectables por las validaciones automáticas."
        )

    st.divider()

    st.write(
        "### Copia de seguridad"
    )

    db_bytes = dataframe_to_xlsx_bytes(
        activities_df,
        evidence_df,
    )

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    st.download_button(
        "💾 Descargar copia de seguridad de la base",
        data=db_bytes,
        file_name=f"Backup_CV_Probatorios_{stamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.divider()

    st.write(
        "### Base de datos"
    )

    st.write(
        f"Archivo: `{DB_FILENAME}`"
    )

    if database_file:
        st.link_button(
            "🔗 Abrir base de datos en Drive",
            database_file.get(
                "webViewLink",
                f"https://drive.google.com/open?id={database_file['id']}",
            ),
        )

    st.write(
        "### Estructura"
    )

    st.code(
        """CV — Sistema de Gestión/
├── 00 — Administración/
├── 01 — Datos personales y CV/
├── 02 — Probatorios/
│   ├── 2024/
│   ├── 2025/
│   └── 2026/
├── 03 — Formación académica/
├── 04 — Docencia/
├── 05 — Investigación/
├── 06 — Ponencias y congresos/
├── 07 — Publicaciones/
├── 08 — Gestión y cargos/
├── 09 — Reconocimientos/
└── 10 — CV generados/

Base_de_Datos_Probatorios_y_CV.xlsx
├── Actividades
└── Probatorios
""",
        language="text",
    )

st.caption(
    "Sistema de gestión documental para CV académico · "
    "Modelo 1:N Actividades → Probatorios"
)
