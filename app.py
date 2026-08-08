import io
import os
import pickle
import pandas as pd
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# ----------------------------------------------------
# CONFIGURACIÓN GENERAL Y ARCHIVOS
# ----------------------------------------------------
st.set_page_config(
    page_title="Control de CV y Probatorios - Dra. Günther",
    page_icon="📄",
    layout="wide",
)

EXCEL_FILE = "Base_de_Datos_Probatorios_y_CV.xlsx"
SHEET_NAME = "Base_de_Datos_Probatorios_y_CV"
CARPETA_PROBATORIOS = "probatorios"

# Google Drive Config
CLIENT_SECRET_FILE = "credentials_oauth.json"
TOKEN_FILE = "token.pickle"
DRIVE_FOLDER_ID = "1mMTyf1QtX7WgrLI10fuO-dHuPCNwLDon"
SCOPES = ["https://www.googleapis.com/auth/drive.file"]

CATEGORIAS = [
    "Coordinación de Libros",
    "Capítulos de Libros / Artículos",
    "Ponencias y Conferencias",
    "Presentaciones de Libros",
    "Comisiones y Arbitrajes",
    "Cursos e Impartición de Clases",
    "Premios y Reconocimientos",
]

if not os.path.exists(CARPETA_PROBATORIOS):
    os.makedirs(CARPETA_PROBATORIOS)


# ----------------------------------------------------
# FUNCIONES AUXILIARES Y CONEXIÓN GOOGLE DRIVE
# ----------------------------------------------------
def cargar_datos():
    return pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)


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


def obtener_servicio_drive():
    """Autentica al usuario usando OAuth 2.0."""
    creds = None
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, "rb") as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception:
                creds = None

        if not creds:
            if not os.path.exists(CLIENT_SECRET_FILE):
                st.error(
                    f"⚠️ No se encontró el archivo `{CLIENT_SECRET_FILE}`."
                )
                return None
            flow = InstalledAppFlow.from_client_secrets_file(
                CLIENT_SECRET_FILE, SCOPES
            )
            creds = flow.run_local_server(port=0)

        with open(TOKEN_FILE, "wb") as token:
            pickle.dump(creds, token)

    return build("drive", "v3", credentials=creds)


def subir_a_google_drive(nombre_archivo, bytes_archivo):
    """Sube el archivo PDF a la carpeta de Google Drive configurada."""
    try:
        service = obtener_servicio_drive()
        if not service:
            return None

        file_metadata = {
            "name": nombre_archivo,
            "parents": [DRIVE_FOLDER_ID],
        }

        media = MediaIoBaseUpload(
            io.BytesIO(bytes_archivo),
            mimetype="application/pdf",
            resumable=True,
        )

        file = (
            service.files()
            .create(
                body=file_metadata,
                media_body=media,
                fields="id, webViewLink",
                supportsAllDrives=True,
            )
            .execute()
        )

        return file.get("webViewLink")
    except Exception as e:
        st.error(f"Error al subir a Google Drive: {e}")
        return None


def crear_cv_word(df):
    """Genera el documento Word del CV con formato académico limpio, ejecutivo y organizado."""
    doc = Document()

    # Margenes de página profesionales (2.5 cm)
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # Estilo base del documento
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
    run_nombre.font.color.rgb = RGBColor(0, 51, 102)  # Azul Marino

    # Subtítulo
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(16)

    run_sub = p_sub.add_run("CURRÍCULUM VITAE — SÍNTESIS EJECUTIVA")
    run_sub.font.size = Pt(10.5)
    run_sub.font.italic = True
    run_sub.font.color.rgb = RGBColor(100, 100, 100)

    # Filtrar solo actividades aprobadas para el CV
    df_cv = df[
        df["Incluir_en_CV"].astype(str).str.strip().str.lower() == "sí"
    ].copy()
    df_cv["Año_num"] = pd.to_numeric(df_cv["Año"], errors="coerce")
    df_cv = df_cv.sort_values(by=["Año_num", "Fecha"], ascending=[False, False])

    categorias_presentes = df_cv["Categoría_CV"].unique()

    for cat in CATEGORIAS:
        if cat in categorias_presentes:
            sub_df = df_cv[df_cv["Categoría_CV"] == cat]
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
                titulo = limpiar_texto(row.get("Título_Actividad_o_Publicación"))
                rol = limpiar_texto(row.get("Rol_Participación"))
                inst = limpiar_texto(row.get("Institución_Organización"))
                lugar = limpiar_texto(row.get("Lugar_Sede"))
                fecha_str = formatear_fecha_cv(row)

                # Si el registro no tiene información válida, omitir
                if not titulo and not rol and not inst:
                    continue

                p_item = doc.add_paragraph(style="List Bullet")
                p_item.paragraph_format.space_after = Pt(4)
                p_item.paragraph_format.space_before = Pt(0)
                p_item.paragraph_format.line_spacing = 1.15

                # Título de la actividad en negrita
                if titulo:
                    run_t = p_item.add_run(titulo)
                    run_t.bold = True

                # Ensamblar los detalles de forma fluida
                detalles = []
                if rol:
                    detalles.append(
                        rol
                        if rol.lower().startswith(("rol", "participación"))
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
                    p_item.add_run(", ".join(detalles) + ".")

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# ----------------------------------------------------
# INTERFAZ WEB (STREAMLIT)
# ----------------------------------------------------
st.title("📄 Sistema de Gestión de CV - Dra. María Griselda Günther")

tab_consulta, tab_registro, tab_cv_word = st.tabs(
    [
        "🔍 Buscar y Consultar Probatorios",
        "➕ Registrar Nueva Actividad",
        "📄 Generar CV en Word",
    ]
)

# --- PESTAÑA 1: BUSCADOR ---
with tab_consulta:
    df = cargar_datos()
    st.subheader("Buscador de Actividades y Documentos")

    st.markdown("##### 🎯 Filtros de Búsqueda")
    col_f1, col_f2, col_f3 = st.columns(3)

    with col_f1:
        filtro_anio = st.selectbox(
            "Filtrar por Año", ["Todos"] + [2023, 2024, 2025, 2026]
        )
    with col_f2:
        filtro_categoria = st.selectbox(
            "Filtrar por Categoría", ["Todas"] + CATEGORIAS
        )
    with col_f3:
        filtro_texto = st.text_input(
            "🔍 Buscar por palabra clave",
            placeholder="Ej. Congreso, Comisión, Libro...",
        )

    df_filtrado = df.copy()

    if filtro_anio != "Todos":
        df_filtrado = df_filtrado[df_filtrado["Año"] == int(filtro_anio)]
    if filtro_categoria != "Todas":
        df_filtrado = df_filtrado[
            df_filtrado["Categoría_CV"] == filtro_categoria
        ]
    if filtro_texto:
        mask = df_filtrado.apply(
            lambda row: row.astype(str)
            .str.contains(filtro_texto, case=False)
            .any(),
            axis=1,
        )
        df_filtrado = df_filtrado[mask]

    st.markdown("---")
    st.write(
        f"Se encontraron **{len(df_filtrado)}** registros coincidentes:"
    )

    st.dataframe(
        df_filtrado,
        use_container_width=True,
        column_config={
            "Enlace_Drive_Probatorio": st.column_config.LinkColumn(
                "Enlace Drive / Local"
            )
        },
    )

    st.markdown("---")
    st.subheader("📥 Visor / Descarga de Probatorio PDF Local")

    if not df_filtrado.empty:
        opciones_actividades = df_filtrado[
            df_filtrado["Nombre_Archivo_PDF"] != "Sin_PDF"
        ]

        if not opciones_actividades.empty:
            actividad_seleccionada = st.selectbox(
                "Selecciona una actividad para descargar su PDF local:",
                opciones_actividades["Título_Actividad_o_Publicación"].tolist(),
            )

            datos_act = opciones_actividades[
                opciones_actividades["Título_Actividad_o_Publicación"]
                == actividad_seleccionada
            ].iloc[0]
            nombre_pdf = datos_act["Nombre_Archivo_PDF"]
            ruta_pdf = os.path.join(CARPETA_PROBATORIOS, nombre_pdf)

            if os.path.exists(ruta_pdf):
                with open(ruta_pdf, "rb") as pdf_file:
                    bytes_pdf = pdf_file.read()

                col_pdf1, col_pdf2 = st.columns([1, 2])
                with col_pdf1:
                    st.success(f"📄 Archivo encontrado: `{nombre_pdf}`")
                    st.download_button(
                        label="⬇️ Descargar PDF Escaneado",
                        data=bytes_pdf,
                        file_name=nombre_pdf,
                        mime="application/pdf",
                    )
            else:
                st.warning(
                    f"⚠️ El archivo `{nombre_pdf}` está registrado pero no se encontró físicamente en la carpeta local."
                )

# --- PESTAÑA 2: CAPTURA ---
with tab_registro:
    st.subheader("Formulario de Captura de Actividades y Constancias")

    with st.form("form_nueva_actividad", clear_on_submit=True):
        col1, col2 = st.columns(2)

        with col1:
            nuevo_id = st.text_input(
                "ID de Registro", value=f"ACT-00{len(cargar_datos()) + 1}"
            )
            anio = st.selectbox("Año", [2023, 2024, 2025, 2026])
            fecha = st.date_input("Fecha")
            categoria = st.selectbox("Categoría del CV", CATEGORIAS)
            rol = st.text_input(
                "Rol / Participación (ej. Ponente, Autora, Coordinadora)"
            )

        with col2:
            titulo = st.text_input("Título de la Actividad o Publicación")
            institucion = st.text_input("Institución u Organización")
            lugar = st.text_input("Lugar / Sede")
            estado = st.selectbox(
                "Estado del Probatorio",
                [
                    "Verificado / En Drive",
                    "Pendiente de Escanear",
                    "En Trámite",
                ],
            )
            incluir = st.radio(
                "¿Incluir en el CV?", ["Sí", "No"], horizontal=True
            )

        st.markdown("---")
        st.subheader("📎 Documento Probatorio (PDF)")
        archivo_pdf = st.file_uploader(
            "Sube el PDF escaneado de la constancia", type=["pdf"]
        )
        notas = st.text_area("Notas / Observaciones de control interno")

        boton_guardar = st.form_submit_button("💾 Guardar Registro y Subir PDF")

    if boton_guardar:
        if not titulo:
            st.error("⚠️ El campo 'Título de la Actividad' es obligatorio.")
        else:
            nombre_pdf_guardado = "Sin_PDF"
            enlace_drive = "Sin_Enlace"

            if archivo_pdf is not None:
                bytes_pdf = archivo_pdf.getbuffer()
                titulo_limpio = "".join(
                    x for x in titulo if x.isalnum() or x in " _-"
                )[:20]
                nombre_pdf_guardado = (
                    f"{anio}_{categoria.replace(' ', '_')}_{titulo_limpio}.pdf"
                )

                # 1. Guardar copia local
                ruta_local = os.path.join(
                    CARPETA_PROBATORIOS, nombre_pdf_guardado
                )
                with open(ruta_local, "wb") as f:
                    f.write(bytes_pdf)

                # 2. Subir a Google Drive
                st.info("☁️ Subiendo archivo a Google Drive...")
                url_drive = subir_a_google_drive(
                    nombre_pdf_guardado, bytes_pdf
                )

                if url_drive:
                    enlace_drive = url_drive
                    st.success("☁️ ¡PDF subido con éxito a Google Drive!")
                else:
                    enlace_drive = ruta_local

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

            try:
                df_actual = cargar_datos()
                df_actualizado = pd.concat(
                    [df_actual, pd.DataFrame([nueva_fila])], ignore_index=True
                )

                with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
                    df_actualizado.to_excel(
                        writer, sheet_name=SHEET_NAME, index=False
                    )

                st.success(
                    f"¡Excelente! Registro '{titulo}' guardado exitosamente."
                )
                st.rerun()
            except PermissionError:
                st.error(
                    "⚠️ **¡Archivo de Excel abierto!** Cierra Microsoft Excel e inténtalo de nuevo."
                )

# --- PESTAÑA 3: GENERADOR DE CV ---
with tab_cv_word:
    st.subheader("📄 Generador Automático de CV impreso (Word)")
    st.write(
        "Este módulo compila automáticamente todas las actividades marcadas con **'¿Incluir en el CV? = Sí'**, formateadas en estilo académico profesional."
    )

    df_total = cargar_datos()
    df_cv_aprobados = df_total[
        df_total["Incluir_en_CV"].astype(str).str.strip().str.lower() == "sí"
    ]

    col_w1, col_w2 = st.columns([2, 1])

    with col_w1:
        st.info(
            f"📌 Actualmente hay **{len(df_cv_aprobados)} actividades** listas para incluirse en el currículum."
        )

    with col_w2:
        if not df_cv_aprobados.empty:
            archivo_word_bytes = crear_cv_word(df_total)

            st.download_button(
                label="📥 Descargar CV Actualizado (.docx)",
                data=archivo_word_bytes,
                file_name="CV_Dra_Maria_Griselda_Gunther.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
        else:
            st.warning("No hay registros marcados con 'Incluir en CV = Sí'.")

    st.markdown("---")
    st.markdown("##### 👁️ Vista previa de los datos que se incluirán:")
    st.dataframe(
        df_cv_aprobados[
            [
                "Año",
                "Categoría_CV",
                "Título_Actividad_o_Publicación",
                "Rol_Participación",
                "Institución_Organización",
            ]
        ],
        use_container_width=True,
    )