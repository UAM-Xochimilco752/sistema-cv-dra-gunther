import streamlit as st
import pandas as pd
import os
import pickle
import io
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from google.auth.transport.requests import Request

# Configuración de la página
st.set_page_config(
    page_title="Sistema de CV y Probatorios - Dra. Gunther",
    page_icon="📚",
    layout="wide"
)

# ---------------------------------------------------------
# FUNCIONES DE CONEXIÓN CON GOOGLE DRIVE API
# ---------------------------------------------------------

@st.cache_resource
def get_drive_service():
    """Conecta con la API de Google Drive usando token.pickle"""
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    
    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
        except Exception as e:
            st.error(f"Error al refrescar credenciales de Google: {e}")
            return None
            
    if not creds:
        st.error("No se encontró el archivo de sesión 'token.pickle' en el proyecto.")
        return None
        
    return build('drive', 'v3', credentials=creds)


def find_excel_file_in_drive(service):
    """Busca el archivo Excel probando coincidencia flexible"""
    try:
        results = service.files().list(
            q="trashed = false",
            fields="files(id, name, mimeType)"
        ).execute()
        files = results.get('files', [])
        
        for f in files:
            if "Base_de_Datos_Probatorios_y_CV" in f['name']:
                return f['id'], f['name'], files
                
        return None, None, files
    except Exception as e:
        st.error(f"Error al consultar la API de Google Drive: {e}")
        return None, None, []


def load_data_from_drive(service, file_id):
    """Lee el Excel guardado en Google Drive y lo pasa a DataFrame"""
    request = service.files().get_media(fileId=file_id)
    file_buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(file_buffer, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    file_buffer.seek(0)
    return pd.read_excel(file_buffer)


def upload_pdf_to_drive(service, uploaded_file):
    """Sube un archivo PDF/Imagen a Google Drive y devuelve su enlace público"""
    file_metadata = {'name': uploaded_file.name}
    media = MediaIoBaseUpload(io.BytesIO(uploaded_file.getvalue()), mimetype=uploaded_file.type)
    
    file = service.files().create(body=file_metadata, media_body=media, fields='id, webViewLink').execute()
    file_id = file.get('id')
    
    service.permissions().create(
        fileId=file_id,
        body={'type': 'anyone', 'role': 'reader'}
    ).execute()
    
    return file.get('webViewLink')


def update_excel_in_drive(service, file_id, df):
    """Sobreescribe el Excel en Google Drive con los datos actualizados"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    
    media = MediaIoBaseUpload(
        output, 
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    service.files().update(fileId=file_id, media_body=media).execute()

# ---------------------------------------------------------
# INTERFAZ PRINCIPAL DE LA APLICACIÓN
# ---------------------------------------------------------

st.title("📚 Sistema de Gestión de CV y Probatorios")
st.subheader("Dra. Gunther — UAM Xochimilco")

service = get_drive_service()

if service:
    excel_id, found_name, all_files = find_excel_file_in_drive(service)
    
    if not excel_id:
        st.warning("⚠️ No se encontró la base de datos en el Google Drive conectado.")
        st.info("💡 **Solución fácil:** Sube tu archivo Excel aquí abajo una sola vez y la aplicación lo colocarán en el Drive correcto por ti.")
        
        # Botón directo para vincular el archivo
        archivo_excel_nuevo = st.file_uploader(
            "Selecciona el archivo 'Base_de_Datos_Probatorios_y_CV.xlsx' de tu computadora:", 
            type=["xlsx"]
        )
        
        if archivo_excel_nuevo is not None:
            with st.spinner("Vinculando y guardando Excel en Google Drive..."):
                file_metadata = {'name': 'Base_de_Datos_Probatorios_y_CV.xlsx'}
                media = MediaIoBaseUpload(
                    io.BytesIO(archivo_excel_nuevo.getvalue()), 
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
                service.files().create(body=file_metadata, media_body=media, fields='id').execute()
                st.success("¡Base de datos guardada y vinculada con éxito!")
                st.rerun()
                
    else:
        st.success(f" Conectado exitosamente a la base de datos: `{found_name}`")
        
        # Cargar datos
        df = load_data_from_drive(service, excel_id)
        
        # Pestañas principales
        tab1, tab2 = st.tabs(["🔍 Consultar Base de Datos y Probatorios", "➕ Cargar Nueva Constancia / Registro"])
        
        # -----------------------------------------------------
        # PESTAÑA 1: VISUALIZADOR Y BÚSQUEDA
        # -----------------------------------------------------
        with tab1:
            st.markdown("### 📋 Registro General de Actividades")
            
            if 'Categoría' in df.columns:
                categorias = ["Todas"] + list(df['Categoría'].dropna().unique())
                cat_seleccionada = st.selectbox("Filtrar por Categoría:", categorias)
                
                if cat_seleccionada != "Todas":
                    df_filtrado = df[df['Categoría'] == cat_seleccionada]
                else:
                    df_filtrado = df
            else:
                df_filtrado = df
            
            busqueda = st.text_input("🔎 Buscar en cualquier campo:")
            if busqueda:
                mask = df_filtrado.astype(str).apply(lambda x: x.str.contains(busqueda, case=False)).any(axis=1)
                df_filtrado = df_filtrado[mask]
                
            st.dataframe(df_filtrado, use_container_width=True)
            st.caption(f"Total de registros mostrados: {len(df_filtrado)}")
            
        # -----------------------------------------------------
        # PESTAÑA 2: NUEVO REGISTRO / CARGA
        # -----------------------------------------------------
        with tab2:
            st.markdown("### 📝 Registrar Nueva Actividad / Constancia")
            st.write("Llena los datos del probatorio. El documento se guardará automáticamente en Google Drive.")
            
            with st.form("form_nueva_constancia", clear_on_submit=True):
                col1, col2 = st.columns(2)
                
                with col1:
                    titulo = st.text_input("Título de la Actividad / Documento *")
                    categoria = st.selectbox("Categoría *", ["Docencia", "Investigación", "Difusión/Congresos", "Asesoría/Tesis", "Gestión/Otros"])
                    ano = st.number_input("Año *", min_value=1970, max_value=2030, value=2026)
                
                with col2:
                    institucion = st.text_input("Institución / Entidad Organizadora")
                    detalles = st.text_area("Detalles / Observaciones opcionales")
                    archivo_pdf = st.file_uploader("Adjuntar Constancia o PDF Probatorio", type=["pdf", "png", "jpg", "jpeg"])
                
                enviado = st.form_submit_button("💾 Guardar Registro y Subir a Google Drive")
                
                if enviado:
                    if not titulo:
                        st.error("Por favor ingresa al menos el título de la actividad.")
                    else:
                        with st.spinner("Subiendo archivo a Google Drive y actualizando la base de datos..."):
                            link_drive = "Sin probatorio"
                            
                            if archivo_pdf is not None:
                                link_drive = upload_pdf_to_drive(service, archivo_pdf)
                            
                            nueva_fila = {
                                'Título': titulo,
                                'Categoría': categoria,
                                'Año': ano,
                                'Institución': institucion,
                                'Detalles': detalles,
                                'Enlace_Probatorio': link_drive
                            }
                            
                            for col in df.columns:
                                if col not in nueva_fila:
                                    nueva_fila[col] = ""
                                    
                            df_nuevo_registro = pd.DataFrame([nueva_fila])
                            df_actualizado = pd.concat([df, df_nuevo_registro], ignore_index=True)
                            
                            update_excel_in_drive(service, excel_id, df_actualizado)
                            
                            st.success("¡Registro guardado con éxito! Google Drive ya ha sido actualizado.")
                            st.balloons()
