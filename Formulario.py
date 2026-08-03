import os
import sqlite3
import pandas as pd
import streamlit as st
from datetime import datetime
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# ============================================================
# CONFIGURACIÓN Y CONEXIÓN A GOOGLE DRIVE (OAUTH 2.0)
# ============================================================
SCOPES = ["https://www.googleapis.com/auth/drive.file"]

def obtener_servicio_drive():
    """Conecta a Google Drive usando OAuth 2.0 (Secrets de Streamlit)."""
    if "gcp_oauth" not in st.secrets:
        st.error("❌ No se encontró la sección [gcp_oauth] en los Secrets de Streamlit.")
        return None

    try:
        creds = Credentials(
            token=None,
            refresh_token=st.secrets["gcp_oauth"]["refresh_token"],
            token_uri="https://oauth2.googleapis.com/token",
            client_id=st.secrets["gcp_oauth"]["client_id"],
            client_secret=st.secrets["gcp_oauth"]["client_secret"],
            scopes=SCOPES,
        )
        return build("drive", "v3", credentials=creds)
    except Exception as e:
        st.error(f"Error al autenticar con Google Drive: {e}")
        return None

def subir_archivo_a_drive(ruta_archivo_local, nombre_en_drive, mime_type="application/octet-stream"):
    """Subes un archivo local a la carpeta especificada en FOLDER_ID."""
    servicio = obtener_servicio_drive()
    if not servicio:
        return None

    try:
        folder_id = st.secrets.get("FOLDER_ID")
        
        metadata = {
            "name": nombre_en_drive,
            "parents": [folder_id] if folder_id else []
        }

        media = MediaFileUpload(ruta_archivo_local, mimetype=mime_type, resumable=True)

        archivo = (
            servicio.files()
            .create(
                body=metadata,
                media_body=media,
                fields="id",
                supportsAllDrives=True,
            )
            .execute()
        )
        return archivo.get("id")

    except Exception as e:
        st.error(f"Error al subir {nombre_en_drive} a Drive: {e}")
        return None

# ============================================================
# BASE DE DATOS LOCAL (SQLITE)
# ============================================================
DB_FILE = "inscripciones.db"

def inicializar_bd():
    """Crea la tabla si no existe en SQLite."""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS participantes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL,
            correo TEXT NOT NULL,
            telefono TEXT,
            fecha_registro TEXT
        )
    """)
    conn.commit()
    conn.close()

def guardar_participante(nombre, correo, telefono):
    """Guarda una nueva inscripción en SQLite."""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cursor.execute("""
        INSERT INTO participantes (nombre, correo, telefono, fecha_registro)
        VALUES (?, ?, ?, ?)
    """, (nombre, correo, telefono, fecha_actual))
    conn.commit()
    conn.close()

def exportar_a_excel():
    """Convierte la BBDD a un archivo Excel local."""
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query("SELECT * FROM participantes", conn)
    conn.close()
    
    excel_path = "reporte_inscripciones.xlsx"
    df.to_excel(excel_path, index=False)
    return excel_path

# ============================================================
# INTERFAZ DE STREAMLIT
# ============================================================
def main():
    st.set_page_config(page_title="Formulario de Inscripción", page_icon="📝")
    st.title("📝 Formulario de Inscripción")

    # Inicializamos la Base de Datos local
    inicializar_bd()

    with st.form("form_registro", clear_on_submit=True):
        nombre = st.text_input("Nombre completo *")
        correo = st.text_input("Correo electrónico *")
        telefono = st.text_input("Teléfono")
        
        enviado = st.form_submit_button("Guardar e Inscribir")

        if enviado:
            if not nombre or not correo:
                st.warning("Por favor completa los campos obligatorios (*).")
            else:
                # 1. Guardar localmente en SQLite
                guardar_participante(nombre, correo, telefono)
                st.success("¡Registro guardado localmente con éxito!")

                # 2. Generar Excel
                excel_local = exportar_a_excel()

                # 3. Subir respaldo en Excel a Google Drive
                with st.spinner("Subiendo reporte en Excel a Google Drive..."):
                    id_excel = subir_archivo_a_drive(
                        ruta_archivo_local=excel_local,
                        nombre_en_drive=f"Inscripciones_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    if id_excel:
                        st.info(f"🟢 Excel respaldado en Drive (ID: `{id_excel}`)")

                # 4. Subir respaldo de la BBDD SQLite a Google Drive
                with st.spinner("Respaldando Base de Datos SQLite en Google Drive..."):
                    id_db = subir_archivo_a_drive(
                        ruta_archivo_local=DB_FILE,
                        nombre_en_drive=f"respaldo_bd_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db",
                        mime_type="application/octet-stream"
                    )
                    if id_db:
                        st.info(f"🟢 BBDD respaldada en Drive (ID: `{id_db}`)")

if __name__ == "__main__":
    main()