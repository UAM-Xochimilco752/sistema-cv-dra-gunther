import io
import os
import pickle
import mimetypes
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
ANIOS_PROBATORIOS = list(range(2020, 2031))

COLUMNAS_ACTIVIDADES = [
    "ID",
    "Año",
    "Fecha",
    "Categoría_CV",
    "Rol_Participación",
    "Título_Actividad_o_Publicación",
    "Institución_Organización",
    "Lugar_Sede",
    "Incluir_en_CV",
    "Notas_Observaciones",
]

COLUMNAS_PROBATORIOS = [
    "ID_Probatorio",
    "ID_Actividad",
    "Nombre_Archivo",
    "Tipo_Probatorio",
    "Enlace_Drive",
    "ID_Drive",
    "Fecha_Alta",
    "Notas",
]

TIPOS_PROBATORIO = [
    "Constancia",
    "Reconocimiento",
    "Programa",
    "Carta",
    "Portada",
    "Artículo / Publicación",
    "Lista de asistencia",
    "Otro",
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
            st.error(f"Error al refrescar credenciales de Google: {e}")
            return None

    if not creds:
        st.error("⚠️ No se encontró el archivo 'token.pickle' en el proyecto.")
        return None

    return build("drive", "v3", credentials=creds)


def buscar_carpeta(service, nombre, parent_id=None):
    nombre_escapado = nombre.replace("'", "\\'")
    if parent_id:
        query = (
            f"name = '{nombre_escapado}' and "
            "mimeType = 'application/vnd.google-apps.folder' and "
            "trashed = false and "
            f"'{parent_id}' in parents"
        )
    else:
        query = (
            f"name = '{nombre_escapado}' and "
            "mimeType = 'application/vnd.google-apps.folder' and "
            "trashed = false and 'root' in parents"
        )
    try:
        result = service.files().list(
            q=query,
            spaces="drive",
            fields="files(id,name,webViewLink)",
            pageSize=10,
        ).execute()
        files = result.get("files", [])
        return files[0] if files else None
    except Exception as e:
        st.error(f"Error al buscar la carpeta '{nombre}': {e}")
        return None


def crear_carpeta(service, nombre, parent_id=None):
    metadata = {
        "name": nombre,
        "mimeType": "application/vnd.google-apps.folder",
    }
    if parent_id:
        metadata["parents"] = [parent_id]
    try:
        return service.files().create(
            body=metadata,
            fields="id,name,webViewLink",
        ).execute()
    except Exception as e:
        st.error(f"Error al crear la carpeta '{nombre}': {e}")
        return None


def obtener_o_crear_carpeta(service, nombre, parent_id=None):
    return buscar_carpeta(service, nombre, parent_id) or crear_carpeta(
        service, nombre, parent_id
    )


@st.cache_resource
def inicializar_estructura_drive(_service):
    estructura = {}
    raiz = obtener_o_crear_carpeta(_service, NOMBRE_CARPETA_RAIZ)
    if not raiz:
        return {}
    estructura["raiz"] = raiz

    for nombre in ESTRUCTURA_CARPETAS:
        carpeta = obtener_o_crear_carpeta(_service, nombre, raiz["id"])
        if carpeta:
            estructura[nombre] = carpeta

    base = estructura.get("02 — Probatorios")
    if not base:
        return estructura

    estructura["probatorios"] = base
    for anio in ANIOS_PROBATORIOS:
        ca = obtener_o_crear_carpeta(_service, str(anio), base["id"])
        if not ca:
            continue
        estructura[f"probatorios_{anio}"] = ca
        for categoria in CATEGORIAS:
            cc = obtener_o_crear_carpeta(_service, categoria, ca["id"])
            if cc:
                estructura[f"probatorios_{anio}_{categoria}"] = cc
    return estructura


def obtener_carpeta_probatorio(service, estructura, anio, categoria):
    clave = f"probatorios_{int(anio)}_{categoria}"
    if clave in estructura:
        return estructura[clave]
    base = estructura.get("probatorios")
    if not base:
        return None
    ca = obtener_o_crear_carpeta(service, str(int(anio)), base["id"])
    if not ca:
        return None
    cc = obtener_o_crear_carpeta(service, categoria, ca["id"])
    if cc:
        estructura[clave] = cc
    return cc


def obtener_mimetype(nombre):
    return mimetypes.guess_type(nombre)[0] or "application/octet-stream"


def subir_a_google_drive(service, nombre, datos, carpeta):
    try:
        metadata = {"name": nombre, "parents": [carpeta["id"]]}
        media = MediaIoBaseUpload(
            io.BytesIO(datos),
            mimetype=obtener_mimetype(nombre),
            resumable=True,
        )
        f = service.files().create(
            body=metadata,
            media_body=media,
            fields="id,name,webViewLink,parents",
        ).execute()
        return f.get("webViewLink", ""), f.get("id", "")
    except Exception as e:
        st.error(f"Error al subir archivo a Google Drive: {e}")
        return "", ""


def eliminar_archivo_drive(service, file_id):
    if not limpiar_texto(file_id):
        return True
    try:
        service.files().update(
            fileId=str(file_id), body={"trashed": True}
        ).execute()
        return True
    except Exception as e:
        st.error(f"Error al mandar el archivo a la papelera: {e}")
        return False


def mover_archivo_drive(service, file_id, nueva_carpeta_id):
    if not limpiar_texto(file_id):
        return False
    try:
        info = service.files().get(fileId=str(file_id), fields="parents").execute()
        padres = info.get("parents", [])
        service.files().update(
            fileId=str(file_id),
            addParents=nueva_carpeta_id,
            removeParents=",".join(padres),
            fields="id,parents",
        ).execute()
        return True
    except Exception as e:
        st.error(f"Error al mover el probatorio: {e}")
        return False


def renombrar_archivo_drive(service, file_id, nuevo_nombre):
    if not limpiar_texto(file_id):
        return False
    try:
        service.files().update(
            fileId=str(file_id),
            body={"name": nuevo_nombre},
            fields="id,name",
        ).execute()
        return True
    except Exception as e:
        st.error(f"Error al renombrar el probatorio: {e}")
        return False

# ============================================================
# EXCEL: DOS TABLAS RELACIONADAS
# ============================================================
def buscar_excel_en_drive(service):
    try:
        result = service.files().list(
            q="trashed = false",
            fields="files(id,name)",
            pageSize=100,
        ).execute()
        for f in result.get("files", []):
            if "Base_de_Datos_Probatorios_y_CV" in f["name"]:
                return f["id"], f["name"]
    except Exception as e:
        st.error(f"Error al consultar Google Drive: {e}")
    return None, None


def descargar_excel(service, file_id):
    request = service.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    buf.seek(0)
    return buf


def columnas_base(df, columnas):
    df = df.copy()
    for col in columnas:
        if col not in df.columns:
            df[col] = ""
    return df


def cargar_datos_drive(service, file_id):
    buf = descargar_excel(service, file_id)
    try:
        xls = pd.ExcelFile(buf)
        if "Actividades" in xls.sheet_names:
            actividades = pd.read_excel(buf, sheet_name="Actividades")
            if "Probatorios" in xls.sheet_names:
                probatorios = pd.read_excel(buf, sheet_name="Probatorios")
            else:
                probatorios = pd.DataFrame(columns=COLUMNAS_PROBATORIOS)
            return columnas_base(actividades, COLUMNAS_ACTIVIDADES), columnas_base(
                probatorios, COLUMNAS_PROBATORIOS
            ), False

        # Compatibilidad con la base anterior: 1 actividad = 1 probatorio.
        antiguo = pd.read_excel(buf, sheet_name=0)
        actividades = antiguo.copy()
        actividades = columnas_base(actividades, COLUMNAS_ACTIVIDADES)
        probatorios = pd.DataFrame(columns=COLUMNAS_PROBATORIOS)

        for _, row in antiguo.iterrows():
            act_id = limpiar_texto(row.get("ID"))
            nombre = limpiar_texto(row.get("Nombre_Archivo_PDF"))
            enlace = limpiar_texto(row.get("Enlace_Drive_Probatorio"))
            drive_id = limpiar_texto(row.get("ID_Drive_Probatorio"))
            if nombre or enlace or drive_id:
                probatorios.loc[len(probatorios)] = [
                    siguiente_id(probatorios, "ID_Probatorio", "PRB-"),
                    act_id,
                    nombre,
                    "Probatorio",
                    enlace,
                    drive_id,
                    "",
                    "Migrado de la base anterior",
                ]

        actividades = actividades.drop(
            columns=[
                c for c in [
                    "Nombre_Archivo_PDF",
                    "Enlace_Drive_Probatorio",
                    "ID_Drive_Probatorio",
                ]
                if c in actividades.columns
            ]
        )
        return actividades, probatorios, True
    except Exception as e:
        st.error(f"Error al leer el Excel: {e}")
        return pd.DataFrame(columns=COLUMNAS_ACTIVIDADES), pd.DataFrame(
            columns=COLUMNAS_PROBATORIOS
        ), False


def actualizar_excel_drive(service, file_id, actividades, probatorios):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        columnas_base(actividades, COLUMNAS_ACTIVIDADES).to_excel(
            writer, sheet_name="Actividades", index=False
        )
        columnas_base(probatorios, COLUMNAS_PROBATORIOS).to_excel(
            writer, sheet_name="Probatorios", index=False
        )
    output.seek(0)
    media = MediaIoBaseUpload(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=True,
    )
    service.files().update(fileId=file_id, media_body=media).execute()

# ============================================================
# UTILIDADES
# ============================================================
def limpiar_texto(valor):
    if pd.isna(valor):
        return ""
    texto = str(valor).strip()
    if texto.lower() in {"nan", "none", "n/a", "na", "ninguno", "ninguna", "-"}:
        return ""
    return texto


def siguiente_id(df, columna, prefijo):
    usados = set(df[columna].dropna().astype(str).str.strip()) if columna in df.columns else set()
    n = 1
    while f"{prefijo}{n:03d}" in usados:
        n += 1
    return f"{prefijo}{n:03d}"


def nombre_probatorio(anio, categoria, titulo, original, actividad_id):
    ext = os.path.splitext(original)[1].lower()
    limpio = "".join(x for x in str(titulo) if x.isalnum() or x in " _-").strip()[:70]
    limpio = limpio or "Sin_Titulo"
    return f"{actividad_id}_{anio}_{categoria.replace(' ', '_')}_{limpio}{ext}"


def fecha_cv(row):
    valor = limpiar_texto(row.get("Fecha"))
    if valor:
        try:
            return pd.to_datetime(valor).strftime("%d/%m/%Y")
        except Exception:
            return valor.split(" ")[0]
    anio = limpiar_texto(row.get("Año"))
    return anio

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

    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("DRA. MARÍA GRISELDA GÜNTHER")
    r.bold = True
    r.font.size = Pt(16)
    r.font.color.rgb = RGBColor(0, 51, 102)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("CURRÍCULUM VITAE — SÍNTESIS EJECUTIVA")
    r.font.size = Pt(10.5)
    r.font.italic = True
    r.font.color.rgb = RGBColor(100, 100, 100)

    if "Incluir_en_CV" not in df.columns:
        return io.BytesIO()

    df_cv = df[df["Incluir_en_CV"].astype(str).str.strip().str.lower() == "sí"].copy()
    df_cv["Año_num"] = pd.to_numeric(df_cv.get("Año"), errors="coerce")
    df_cv = df_cv.sort_values("Año_num", ascending=False)

    for categoria in CATEGORIAS:
        sub = df_cv[df_cv["Categoría_CV"] == categoria]
        if sub.empty:
            continue
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        r = p.add_run(categoria)
        r.bold = True
        r.font.size = Pt(12.5)
        r.font.color.rgb = RGBColor(0, 51, 102)

        for _, row in sub.iterrows():
            titulo = limpiar_texto(row.get("Título_Actividad_o_Publicación"))
            rol = limpiar_texto(row.get("Rol_Participación"))
            inst = limpiar_texto(row.get("Institución_Organización"))
            lugar = limpiar_texto(row.get("Lugar_Sede"))
            fecha = fecha_cv(row)
            if not (titulo or rol or inst):
                continue
            p = doc.add_paragraph(style="List Bullet")
            if titulo:
                r = p.add_run(titulo)
                r.bold = True
            detalles = []
            if rol:
                detalles.append(rol if rol.lower().startswith(("rol", "participación")) else f"Rol: {rol}")
            if inst:
                detalles.append(inst)
            if lugar:
                detalles.append(lugar)
            if fecha:
                detalles.append(fecha)
            if detalles:
                if titulo:
                    p.add_run(". ")
                p.add_run(", ".join(detalles) + ".")

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# ============================================================
# INICIO DE APP
# ============================================================
st.title("📄 Sistema de Gestión de CV - Dra. María Griselda Günther")
service = obtener_servicio_drive()

if service:
    with st.spinner("🔧 Verificando estructura de Google Drive..."):
        estructura_drive = inicializar_estructura_drive(service)

    excel_id, found_name = buscar_excel_en_drive(service)

    if not excel_id:
        st.warning("⚠️ No se encontró la base de datos en Google Drive.")
        archivo = st.file_uploader(
            "Sube Base_de_Datos_Probatorios_y_CV.xlsx",
            type=["xlsx"],
        )
        if archivo:
            media = MediaIoBaseUpload(
                io.BytesIO(archivo.getvalue()),
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                resumable=True,
            )
            service.files().create(
                body={"name": "Base_de_Datos_Probatorios_y_CV.xlsx"},
                media_body=media,
                fields="id",
            ).execute()
            st.success("¡Base de datos vinculada!")
            st.rerun()
    else:
        actividades, probatorios, migrar = cargar_datos_drive(service, excel_id)

        if migrar:
            st.warning(
                "🔄 Se detectó tu formato anterior. Se migrará automáticamente a "
                "dos hojas: Actividades y Probatorios."
            )
            actualizar_excel_drive(service, excel_id, actividades, probatorios)
            st.success("✅ Migración completada. No se modificaron las carpetas de Drive.")
            st.rerun()

        tab_buscar, tab_nueva, tab_cv = st.tabs([
            "🔍 Buscar y Administrar",
            "➕ Registrar Actividad",
            "📄 Generar CV",
        ])

        # --------------------------------------------------------
        # BUSCADOR / ADMINISTRACIÓN
        # --------------------------------------------------------
        with tab_buscar:
            st.subheader("Buscar y administrar actividades")

            vista = actividades.copy()
            conteos = probatorios.groupby("ID_Actividad").size() if not probatorios.empty else pd.Series(dtype=int)
            vista["Cantidad_Probatorios"] = vista["ID"].astype(str).map(conteos).fillna(0).astype(int)
            vista["Estado_Documental"] = vista["Cantidad_Probatorios"].apply(
                lambda n: "🟡 Sin probatorio" if n == 0 else "🟢 1 probatorio" if n == 1 else f"🔵 {n} probatorios"
            )

            f1, f2, f3 = st.columns(3)
            with f1:
                anios = pd.to_numeric(vista["Año"], errors="coerce").dropna().astype(int).unique().tolist()
                anio_filtro = st.selectbox("Año", ["Todos"] + sorted(anios, reverse=True))
            with f2:
                categoria_filtro = st.selectbox("Categoría", ["Todas"] + CATEGORIAS)
            with f3:
                texto_filtro = st.text_input("🔍 Buscar", placeholder="ID, título, institución...")

            filtrada = vista.copy()
            if anio_filtro != "Todos":
                filtrada = filtrada[pd.to_numeric(filtrada["Año"], errors="coerce") == int(anio_filtro)]
            if categoria_filtro != "Todas":
                filtrada = filtrada[filtrada["Categoría_CV"] == categoria_filtro]
            if texto_filtro:
                mask = filtrada.astype(str).apply(
                    lambda col: col.str.contains(texto_filtro, case=False, na=False, regex=False)
                ).any(axis=1)
                filtrada = filtrada[mask]

            columnas = [
                "ID", "Año", "Categoría_CV", "Título_Actividad_o_Publicación",
                "Institución_Organización", "Cantidad_Probatorios", "Estado_Documental"
            ]
            st.dataframe(filtrada[columnas], use_container_width=True, hide_index=True)

            if actividades.empty:
                st.info("No hay actividades registradas.")
            else:
                ids = actividades["ID"].astype(str).tolist()
                seleccion = st.selectbox(
                    "Selecciona una actividad para administrarla",
                    ids,
                    format_func=lambda x: f"{x} — {limpiar_texto(actividades.loc[actividades['ID'].astype(str) == x, 'Título_Actividad_o_Publicación'].iloc[0])}",
                )
                idx = actividades[actividades["ID"].astype(str) == seleccion].index[0]
                act = actividades.loc[idx].copy()
                prob_act = probatorios[probatorios["ID_Actividad"].astype(str) == seleccion].copy()

                st.markdown(f"### {limpiar_texto(act.get('Título_Actividad_o_Publicación'))}")
                st.info(f"📎 Esta actividad tiene **{len(prob_act)} probatorio(s)** asociados.")

                # EDITAR ACTIVIDAD
                with st.expander("✏️ Editar actividad"):
                    try:
                        anio_actual = int(float(act.get("Año")))
                    except Exception:
                        anio_actual = date.today().year
                    try:
                        fecha_actual = pd.to_datetime(act.get("Fecha")).date()
                    except Exception:
                        fecha_actual = date.today()

                    with st.form(f"editar_{seleccion}"):
                        c1, c2 = st.columns(2)
                        with c1:
                            e_anio = st.number_input("Año", min_value=1900, max_value=2100, value=anio_actual, step=1)
                            e_fecha = st.date_input("Fecha", value=fecha_actual)
                            e_categoria = st.selectbox(
                                "Categoría",
                                CATEGORIAS,
                                index=CATEGORIAS.index(act.get("Categoría_CV")) if act.get("Categoría_CV") in CATEGORIAS else 0,
                            )
                            e_rol = st.text_input("Rol / Participación", value=limpiar_texto(act.get("Rol_Participación")))
                        with c2:
                            e_titulo = st.text_input("Título *", value=limpiar_texto(act.get("Título_Actividad_o_Publicación")))
                            e_inst = st.text_input("Institución / Organización", value=limpiar_texto(act.get("Institución_Organización")))
                            e_lugar = st.text_input("Lugar / Sede", value=limpiar_texto(act.get("Lugar_Sede")))
                            e_incluir = st.radio(
                                "¿Incluir en CV?", ["Sí", "No"], horizontal=True,
                                index=0 if limpiar_texto(act.get("Incluir_en_CV")).lower() == "sí" else 1,
                            )
                        e_notas = st.text_area("Notas / Observaciones", value=limpiar_texto(act.get("Notas_Observaciones")))
                        guardar = st.form_submit_button("💾 Guardar cambios")

                    if guardar:
                        if not e_titulo.strip():
                            st.error("El título es obligatorio.")
                        else:
                            # Si cambia año/categoría, mover todos los probatorios de la actividad.
                            cambio_carpeta = (
                                int(pd.to_numeric(act.get("Año"), errors="coerce")) != int(e_anio)
                                or limpiar_texto(act.get("Categoría_CV")) != e_categoria
                            )
                            actividades.at[idx, "Año"] = int(e_anio)
                            actividades.at[idx, "Fecha"] = str(e_fecha)
                            actividades.at[idx, "Categoría_CV"] = e_categoria
                            actividades.at[idx, "Rol_Participación"] = e_rol
                            actividades.at[idx, "Título_Actividad_o_Publicación"] = e_titulo
                            actividades.at[idx, "Institución_Organización"] = e_inst
                            actividades.at[idx, "Lugar_Sede"] = e_lugar
                            actividades.at[idx, "Incluir_en_CV"] = e_incluir
                            actividades.at[idx, "Notas_Observaciones"] = e_notas

                            if cambio_carpeta and not prob_act.empty:
                                nueva_carpeta = obtener_carpeta_probatorio(service, estructura_drive, int(e_anio), e_categoria)
                                if nueva_carpeta:
                                    for _, p in prob_act.iterrows():
                                        mover_archivo_drive(service, p.get("ID_Drive"), nueva_carpeta["id"])

                            actualizar_excel_drive(service, excel_id, actividades, probatorios)
                            st.success("✅ Actividad actualizada.")
                            st.rerun()

                # PROBATORIOS
                st.markdown("### 📎 Probatorios asociados")
                if prob_act.empty:
                    st.warning("Esta actividad no tiene probatorios todavía.")
                else:
                    for _, p in prob_act.iterrows():
                        pid = limpiar_texto(p.get("ID_Probatorio"))
                        nombre = limpiar_texto(p.get("Nombre_Archivo"))
                        tipo = limpiar_texto(p.get("Tipo_Probatorio")) or "Probatorio"
                        enlace = limpiar_texto(p.get("Enlace_Drive"))
                        c1, c2, c3, c4 = st.columns([1.2, 4, 2, 1.5])
                        c1.write(f"**{pid}**")
                        c2.write(nombre or "Sin nombre")
                        c3.write(tipo)
                        if enlace:
                            c4.link_button("🔗 Abrir", enlace)

                        with st.expander(f"⚙️ Administrar {pid}"):
                            confirmar = st.checkbox("Confirmo que quiero mandar este probatorio a la papelera.", key=f"conf_p_{pid}")
                            if st.button("🗑️ Eliminar probatorio", key=f"del_p_{pid}", disabled=not confirmar):
                                eliminar_archivo_drive(service, p.get("ID_Drive"))
                                probatorios = probatorios[probatorios["ID_Probatorio"].astype(str) != pid].copy()
                                actualizar_excel_drive(service, excel_id, actividades, probatorios)
                                st.success("Probatorio eliminado.")
                                st.rerun()

                # AGREGAR UNO O VARIOS EN UNA MISMA OPERACIÓN
                st.markdown("### ➕ Agregar probatorios")
                with st.form(f"form_multiples_{seleccion}"):
                    archivos = st.file_uploader(
                        "Selecciona uno o varios documentos",
                        type=["pdf", "png", "jpg", "jpeg"],
                        accept_multiple_files=True,
                        key=f"multi_{seleccion}",
                    )
                    tipo = st.selectbox("Tipo que tendrán estos archivos", TIPOS_PROBATORIO, key=f"tipo_{seleccion}")
                    notas = st.text_area("Notas para estos probatorios", key=f"notas_{seleccion}")
                    subir = st.form_submit_button("📎 Subir y asociar archivos")

                if subir:
                    if not archivos:
                        st.error("Selecciona al menos un archivo.")
                    else:
                        try:
                            anio = int(float(act.get("Año")))
                        except Exception:
                            anio = date.today().year
                        categoria = limpiar_texto(act.get("Categoría_CV"))
                        titulo = limpiar_texto(act.get("Título_Actividad_o_Publicación"))
                        carpeta = obtener_carpeta_probatorio(service, estructura_drive, anio, categoria)
                        if not carpeta:
                            st.error("No se pudo determinar la carpeta de destino.")
                        else:
                            exitos = 0
                            for archivo in archivos:
                                nombre = nombre_probatorio(anio, categoria, titulo, archivo.name, seleccion)
                                enlace, drive_id = subir_a_google_drive(service, nombre, archivo.getvalue(), carpeta)
                                if drive_id:
                                    probatorios.loc[len(probatorios)] = [
                                        siguiente_id(probatorios, "ID_Probatorio", "PRB-"),
                                        seleccion,
                                        nombre,
                                        tipo,
                                        enlace,
                                        drive_id,
                                        pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
                                        notas,
                                    ]
                                    exitos += 1
                            actualizar_excel_drive(service, excel_id, actividades, probatorios)
                            st.success(f"✅ Se agregaron {exitos} probatorio(s) a {seleccion}.")
                            st.rerun()

                # ELIMINAR ACTIVIDAD Y TODOS SUS PROBATORIOS
                st.markdown("---")
                st.subheader("🗑️ Eliminar actividad completa")
                st.warning("Esto mandará a la papelera de Drive todos los probatorios asociados y eliminará la actividad de la hoja Actividades.")
                confirmar_act = st.checkbox("Confirmo que quiero eliminar la actividad y todos sus probatorios.", key=f"conf_a_{seleccion}")
                if st.button("🗑️ Eliminar actividad completa", key=f"del_a_{seleccion}", disabled=not confirmar_act, type="primary"):
                    for _, p in prob_act.iterrows():
                        eliminar_archivo_drive(service, p.get("ID_Drive"))
                    actividades = actividades[actividades["ID"].astype(str) != seleccion].copy()
                    probatorios = probatorios[probatorios["ID_Actividad"].astype(str) != seleccion].copy()
                    actualizar_excel_drive(service, excel_id, actividades, probatorios)
                    st.success("Actividad y probatorios eliminados.")
                    st.rerun()

        # --------------------------------------------------------
        # NUEVA ACTIVIDAD
        # --------------------------------------------------------
        with tab_nueva:
            st.subheader("➕ Registrar nueva actividad")
            nuevo_id = siguiente_id(actividades, "ID", "ACT-")
            st.info(f"El sistema asignará automáticamente el ID **{nuevo_id}**.")

            with st.form("nueva_actividad", clear_on_submit=True):
                c1, c2 = st.columns(2)
                with c1:
                    n_anio = st.number_input("Año", min_value=1900, max_value=2100, value=date.today().year, step=1)
                    n_fecha = st.date_input("Fecha", value=date.today())
                    n_categoria = st.selectbox("Categoría", CATEGORIAS)
                    n_rol = st.text_input("Rol / Participación")
                with c2:
                    n_titulo = st.text_input("Título de la Actividad o Publicación *")
                    n_inst = st.text_input("Institución / Organización")
                    n_lugar = st.text_input("Lugar / Sede")
                    n_incluir = st.radio("¿Incluir en el CV?", ["Sí", "No"], horizontal=True)
                n_notas = st.text_area("Notas / Observaciones")
                guardar_nueva = st.form_submit_button("💾 Guardar actividad")

            if guardar_nueva:
                if not n_titulo.strip():
                    st.error("El título es obligatorio.")
                else:
                    actividades.loc[len(actividades)] = [
                        nuevo_id,
                        int(n_anio),
                        str(n_fecha),
                        n_categoria,
                        n_rol,
                        n_titulo,
                        n_inst,
                        n_lugar,
                        n_incluir,
                        n_notas,
                    ]
                    actualizar_excel_drive(service, excel_id, actividades, probatorios)
                    st.success(f"✅ Actividad {nuevo_id} guardada.")
                    st.info("Ahora puedes asociarle uno o varios probatorios desde 'Buscar y Administrar'.")
                    st.rerun()

        # --------------------------------------------------------
        # CV
        # --------------------------------------------------------
        with tab_cv:
            st.subheader("📄 Generador Automático de CV")
            df_cv = actividades[actividades["Incluir_en_CV"].astype(str).str.strip().str.lower() == "sí"]
            st.info(f"Hay **{len(df_cv)} actividades** marcadas para el CV.")
            if not df_cv.empty:
                st.download_button(
                    "📥 Descargar CV Actualizado (.docx)",
                    data=crear_cv_word(actividades),
                    file_name="CV_Dra_Maria_Griselda_Gunther.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
            else:
                st.warning("No hay actividades marcadas para incluir en el CV.")

            st.markdown("---")
            st.subheader("📊 Estado documental")
            conteos = actividades["ID"].astype(str).map(
                probatorios.groupby("ID_Actividad").size() if not probatorios.empty else pd.Series(dtype=int)
            ).fillna(0).astype(int)
            a, b, c, d = st.columns(4)
            a.metric("Actividades", len(actividades))
            b.metric("Sin probatorio", int((conteos == 0).sum()))
            c.metric("Con 1 probatorio", int((conteos == 1).sum()))
            d.metric("Con múltiples", int((conteos > 1).sum()))

            st.markdown("### Vista previa")
            cols = [
                "ID", "Año", "Categoría_CV", "Título_Actividad_o_Publicación",
                "Rol_Participación", "Institución_Organización", "Lugar_Sede"
            ]
            st.dataframe(df_cv[cols], use_container_width=True, hide_index=True)
