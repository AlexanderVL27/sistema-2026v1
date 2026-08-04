import datetime
import os
import sqlite3
from io import BytesIO
import openpyxl
from openpyxl.styles import Alignment, Font
import pandas as pd
import streamlit as st

# Configuración inicial de la página
st.set_page_config(
    page_title="Sistema de Inscripción 2026", page_icon="📝", layout="wide"
)

PASSWORD_ADMIN = st.secrets.get("PASSWORD_ADMIN", "admin123")
DB_FILE = "inscripciones.db"
PLANTILLA_EXCEL = "SOLIC INSCRIP NVO 2026.xlsx"


# ==========================================
# GESTIÓN BASE DE DATOS LOCAL (SQLITE)
# ==========================================
@st.cache_resource
def get_db_connection():
    """Mantiene una conexión persistente y segura para hilos a SQLite."""
    conn = sqlite3.connect(DB_FILE, check_same_thread=False)
    return conn


def inicializar_db():
    """Crea la tabla de alumnos si no existe en la BBDD local."""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS alumnos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha_registro TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            nombre_alumno TEXT, curp TEXT UNIQUE, fecha_nacimiento TEXT, edad TEXT,
            sexo TEXT, lugar_nacimiento TEXT, celular_alumno TEXT, correo TEXT,
            red_social TEXT, secundaria TEXT, cct TEXT, promedio TEXT, carrera TEXT,
            turno TEXT, estatus TEXT, observaciones TEXT, nombre_tutor TEXT,
            domicilio TEXT, celular_tutor TEXT, tel_casa TEXT, tel_emergencia TEXT,
            ocupacion TEXT, docs_entregados TEXT
        )
    """)
    conn.commit()


def vaciar_db():
    """Elimina todos los registros de la tabla de alumnos."""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM alumnos")
    conn.commit()


def guardar_en_db(datos):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(
        """
        INSERT INTO alumnos (
            nombre_alumno, curp, fecha_nacimiento, edad, sexo, lugar_nacimiento,
            celular_alumno, correo, red_social, secundaria, cct, promedio,
            carrera, turno, estatus, observaciones, nombre_tutor, domicilio,
            celular_tutor, tel_casa, tel_emergencia, ocupacion, docs_entregados
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """,
        datos,
    )
    conn.commit()


def actualizar_en_db(id_alumno, datos):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(
        """
        UPDATE alumnos SET
            nombre_alumno=?, curp=?, fecha_nacimiento=?, edad=?, sexo=?, lugar_nacimiento=?,
            celular_alumno=?, correo=?, red_social=?, secundaria=?, cct=?, promedio=?,
            carrera=?, turno=?, estatus=?, observaciones=?, nombre_tutor=?, domicilio=?,
            celular_tutor=?, tel_casa=?, tel_emergencia=?, ocupacion=?, docs_entregados=?
        WHERE id=?
    """,
        (*datos, id_alumno),
    )
    conn.commit()


def obtener_alumnos():
    conn = get_db_connection()
    return pd.read_sql_query(
        "SELECT * FROM alumnos ORDER BY fecha_registro DESC", conn
    )


def escribir_celda_segura(sheet, celda, valor, alineacion=None, fuente=None):
    """Escribe en una celda conservando el formato o aplicando nuevos estilos."""
    try:
        cell = sheet[celda]
        cell.value = valor
        if alineacion:
            cell.alignment = alineacion
        if fuente:
            cell.font = fuente
    except AttributeError:
        pass


# ==========================================
# FUNCIÓN PARA GENERAR EXCEL RELLENADO
# ==========================================
def generar_excel_alumno(
    nombre_alumno,
    dia_nac,
    mes_nac,
    anio_nac,
    edad,
    lugar_nac,
    sexo,
    celular_alumno,
    correo,
    red_social,
    curp,
    secundaria,
    cct,
    promedio,
    carrera,
    turno,
    estatus,
    observaciones,
    nombre_tutor,
    domicilio,
    celular_tutor,
    tel_casa,
    tel_emergencia,
    ocupacion,
    docs_list,
):
    """Rellena la plantilla Excel con los datos del alumno y mantiene el diseño exacto para impresión."""
    if not os.path.exists(PLANTILLA_EXCEL):
        return None

    wb = openpyxl.load_workbook(PLANTILLA_EXCEL)
    sheet = wb.active

    # --- AJUSTES DE IMPRESIÓN EXACTOS ---
    sheet.page_setup.orientation = sheet.ORIENTATION_PORTRAIT
    sheet.page_setup.paperSize = sheet.PAPERSIZE_LETTER
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 1

    escribir_celda_segura(sheet, "G2", nombre_alumno)
    escribir_celda_segura(sheet, "I5", dia_nac)
    escribir_celda_segura(sheet, "M5", mes_nac)
    escribir_celda_segura(sheet, "Q5", anio_nac)
    escribir_celda_segura(sheet, "U5", edad)
    escribir_celda_segura(sheet, "I8", lugar_nac)

    # --- CELDAS DE SEXO (V8 Y V9) ---
    sexo_clean = str(sexo).strip().upper() if sexo else ""
    if sexo_clean == "FEMENINO":
        escribir_celda_segura(sheet, "V8", "X")
    elif sexo_clean == "MASCULINO":
        escribir_celda_segura(sheet, "V9", "X")

    escribir_celda_segura(sheet, "I10", celular_alumno)
    escribir_celda_segura(sheet, "I11", correo)
    escribir_celda_segura(sheet, "J12", red_social)

    escribir_celda_segura(sheet, "E14", curp)
    escribir_celda_segura(sheet, "I15", secundaria)
    escribir_celda_segura(sheet, "H16", cct)
    escribir_celda_segura(sheet, "T16", promedio)
    escribir_celda_segura(sheet, "H17", carrera)

    if turno == "MATUTINO":
        escribir_celda_segura(sheet, "H18", "X")
    elif turno == "VESPERTINO":
        escribir_celda_segura(sheet, "M18", "X")

    map_estatus = {
        "ASIGNADO": "D19",
        "CAMBIO": "H19",
        "OTRO RESULTADO": "N19",
        "SIN PROCESO": "V19",
    }
    if estatus in map_estatus:
        escribir_celda_segura(sheet, map_estatus[estatus], "X")

    escribir_celda_segura(sheet, "Q18", observaciones)
    escribir_celda_segura(sheet, "G22", nombre_tutor)
    escribir_celda_segura(sheet, "H24", domicilio)
    escribir_celda_segura(sheet, "I27", celular_tutor)
    escribir_celda_segura(sheet, "I28", tel_casa)
    escribir_celda_segura(sheet, "I29", tel_emergencia)
    escribir_celda_segura(sheet, "I30", ocupacion)

    # Mapeo de casillas de documentos entregados (del 1 al 11)
    cell_docs_map = {
        "1": "J32",
        "2": "J33",
        "3": "J34",
        "4": "J35",
        "5": "J36",
        "6": "J37",
        "7": "J38",
        "8": "V32",
        "9": "V34",
        "10": "V36",
        "11": "V37",
    }
    for doc_num in docs_list:
        doc_num_str = str(doc_num).strip()
        if doc_num_str in cell_docs_map:
            escribir_celda_segura(sheet, cell_docs_map[doc_num_str], "X")

    # --- FECHA EN FILA 41 (IGUALITA AL FORMATO ORIGINAL) ---
    hoy = datetime.date.today()
    meses_es = [
        "ENERO",
        "FEBRERO",
        "MARZO",
        "ABRIL",
        "MAYO",
        "JUNIO",
        "JULIO",
        "AGOSTO",
        "SEPTIEMBRE",
        "OCTUBRE",
        "NOVIEMBRE",
        "DICIEMBRE",
    ]
    nombre_mes = meses_es[hoy.month - 1]
    texto_fecha_completa = f"COL. NETZAHUALCOYOTL, TEXCOCO, MEXICO. A {hoy.day:02d} DE {nombre_mes} DEL {hoy.year}"

    # Aplicar con el estilo exacto de la plantilla original (Aptos Narrow, 11pt, Centrado)
    fuente_original = Font(name="Aptos Narrow", size=11, bold=False)
    alineacion_centrada = Alignment(
        horizontal="center", vertical="center", wrap_text=False
    )

    escribir_celda_segura(
        sheet,
        "A41",
        texto_fecha_completa,
        alineacion=alineacion_centrada,
        fuente=fuente_original,
    )

    excel_buffer = BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)
    return excel_buffer.getvalue()


# Inicializar tabla al arrancar
inicializar_db()

# ==========================================
# INTERFAZ Y NAVEGACIÓN
# ==========================================
tab1, tab2 = st.tabs(["📝 Formulario de Inscripción", "🔒 Panel Administrador"])

# --- TAB 1: FORMULARIO DE INSCRIPCIÓN ---
with tab1:
    st.title("📝 Solicitud de Inscripción 2026")

    with st.form("form_inscripcion", clear_on_submit=False):
        st.header("1. Datos Personales del Alumno")
        nombre_alumno = st.text_input("Nombre completo del Alumno:")

        col_f1, col_f2 = st.columns([2, 1])
        fecha_nac = col_f1.date_input(
            "Fecha de Nacimiento:",
            value=datetime.date(2010, 1, 1),
            min_value=datetime.date(1990, 1, 1),
            max_value=datetime.date.today(),
            format="DD/MM/YYYY",
        )
        edad = col_f2.text_input("Años (Edad):")

        lugar_nac = st.text_input("Lugar de Nacimiento:")
        sexo = st.radio("Sexo:", ["FEMENINO", "MASCULINO"], horizontal=True)

        col_c1, col_c2 = st.columns(2)
        celular_alumno = col_c1.text_input("Celular del Alumno:")
        correo = col_c2.text_input("Correo electrónico:")
        red_social = st.text_input("Red Social:")

        st.header("2. Datos Académicos")
        col_a1, col_a2 = st.columns(2)
        curp = col_a1.text_input("CURP del Alumno:")
        promedio = col_a2.text_input("Promedio de Secundaria:")

        secundaria = st.text_input("Secundaria de procedencia:")
        cct = st.text_input("CCT de la Secundaria:")
        carrera = st.text_input("Aceptado en la Carrera de:")

        col_t1, col_t2 = st.columns(2)
        turno = col_t1.radio(
            "Turno:", ["MATUTINO", "VESPERTINO"], horizontal=True
        )
        estatus = col_t2.radio(
            "Resultado:",
            ["ASIGNADO", "CAMBIO", "OTRO RESULTADO", "SIN PROCESO"],
            horizontal=True,
        )
        observaciones = st.text_input("Observaciones:")

        st.header("3. Datos del Tutor")
        nombre_tutor = st.text_input("Nombre completo del Tutor:")
        domicilio = st.text_input("Domicilio particular:")

        col_tut1, col_tut2, col_tut3 = st.columns(3)
        celular_tutor = col_tut1.text_input("Celular del Tutor:")
        tel_casa = col_tut2.text_input("Teléfono de Casa:")
        tel_emergencia = col_tut3.text_input("Teléfono de Emergencia:")
        ocupacion = st.text_input("Ocupación del Tutor:")

        st.header("4. Documentación Entregada")
        col_d1, col_d2 = st.columns(2)
        doc1 = col_d1.checkbox("1.- Voucher de Pago Original")
        doc2 = col_d1.checkbox("2.- Comprobante de Asignación")
        doc3 = col_d1.checkbox("3.- Certificado de Secundaria")
        doc4 = col_d1.checkbox("4.- Boleta de 3er Año")
        doc5 = col_d1.checkbox("5.- CURP Alumno")
        doc6 = col_d1.checkbox("6.- Acta de Nacimiento")
        doc7 = col_d1.checkbox("7.- Certificado Médico")

        doc8 = col_d2.checkbox("8.- Comprobante Domicilio")
        doc9 = col_d2.checkbox("9.- INE Tutor")
        doc10 = col_d2.checkbox("10.- CURP Tutor")
        doc11 = col_d2.checkbox("11.- 3 Fotografías Infantil")

        enviado = st.form_submit_button("💾 GUARDAR SOLICITUD LOCALMENTE")

    if enviado:
        if not nombre_alumno or not curp:
            st.error("⚠️ Debes llenar al menos Nombre del Alumno y CURP.")
        else:
            try:
                # Extraer día, mes y año de la fecha seleccionada
                dia_nac = f"{fecha_nac.day:02d}"
                mes_nac = f"{fecha_nac.month:02d}"
                anio_nac = str(fecha_nac.year)

                # Recopilar documentos marcados
                docs_checks = [
                    doc1,
                    doc2,
                    doc3,
                    doc4,
                    doc5,
                    doc6,
                    doc7,
                    doc8,
                    doc9,
                    doc10,
                    doc11,
                ]
                lista_docs = [
                    str(i) for i, chk in enumerate(docs_checks, 1) if chk
                ]

                # 1. Guardar en la Base de Datos SQLite local
                datos_alumno = (
                    nombre_alumno,
                    curp,
                    f"{dia_nac}/{mes_nac}/{anio_nac}",
                    edad,
                    sexo,
                    lugar_nac,
                    celular_alumno,
                    correo,
                    red_social,
                    secundaria,
                    cct,
                    promedio,
                    carrera,
                    turno,
                    estatus,
                    observaciones,
                    nombre_tutor,
                    domicilio,
                    celular_tutor,
                    tel_casa,
                    tel_emergencia,
                    ocupacion,
                    ",".join(lista_docs),
                )
                guardar_en_db(datos_alumno)

                st.success(
                    "✅ ¡Inscripción guardada correctamente en la BBDD local!"
                )

                # 2. Generar Excel individual para descarga inmediata
                bytes_excel = generar_excel_alumno(
                    nombre_alumno,
                    dia_nac,
                    mes_nac,
                    anio_nac,
                    edad,
                    lugar_nac,
                    sexo,
                    celular_alumno,
                    correo,
                    red_social,
                    curp,
                    secundaria,
                    cct,
                    promedio,
                    carrera,
                    turno,
                    estatus,
                    observaciones,
                    nombre_tutor,
                    domicilio,
                    celular_tutor,
                    tel_casa,
                    tel_emergencia,
                    ocupacion,
                    lista_docs,
                )

                if bytes_excel:
                    nombre_limpio = "".join(
                        c for c in nombre_alumno if c.isalnum() or c == " "
                    ).strip()
                    st.download_button(
                        label="📄 Descargar Solicitud en Excel (.xlsx)",
                        data=bytes_excel,
                        file_name=f"SOLICITUD_{nombre_limpio}_{curp}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                else:
                    st.warning(
                        f"⚠️ No se encontró la plantilla `{PLANTILLA_EXCEL}` en el servidor para generar el archivo."
                    )

            except sqlite3.IntegrityError:
                st.error(f"⚠️ La CURP **{curp}** ya se encuentra registrada.")
            except Exception as e:
                st.error(f"Error al guardar la solicitud: {e}")

# --- TAB 2: PANEL ADMINISTRADOR ---
with tab2:
    st.title("🔒 Panel Administrador")
    pass_input = st.text_input("Contraseña:", type="password")

    if pass_input == PASSWORD_ADMIN:

        # --- SUBIR BBDD MANUAL ---
        with st.expander(
            "📤 Cargar / Reemplazar Base de Datos (.db)", expanded=False
        ):
            st.info(
                "Sube una copia previa de tu archivo `inscripciones.db` para consultar o actualizar los datos."
            )
            uploaded_db = st.file_uploader(
                "Selecciona un archivo .db local:", type=["db"]
            )
            if uploaded_db is not None:
                if st.button("🔄 Cargar esta Base de Datos ahora"):
                    with open(DB_FILE, "wb") as f:
                        f.write(uploaded_db.getbuffer())
                    st.cache_resource.clear()
                    st.success(
                        "✅ Base de datos cargada y actualizada con éxito."
                    )
                    st.rerun()

        st.markdown("---")
        df_alumnos = obtener_alumnos()
        st.write(f"**Total de alumnos registrados:** {len(df_alumnos)}")

        col_d1, col_d2 = st.columns(2)

        # Exportar Padrón Consolidado
        output_excel = BytesIO()
        with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
            df_alumnos.to_excel(
                writer, index=False, sheet_name="Padrón Completo"
            )

        col_d1.download_button(
            label="📥 Exportar Padrón Completo (.xlsx)",
            data=output_excel.getvalue(),
            file_name="Padron_Inscripciones_2026.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # Descargar copia de la BBDD SQLite local (.db)
        if os.path.exists(DB_FILE):
            with open(DB_FILE, "rb") as f_db:
                col_d2.download_button(
                    label="🗄️ Descargar Copia BBDD (.db)",
                    data=f_db.read(),
                    file_name="inscripciones.db",
                    mime="application/x-sqlite3",
                )

        # --- DESCARGAR EXCEL INDIVIDUAL DESDE EL ADMIN ---
        st.markdown("---")
        st.subheader("📄 Descargar Solicitud Individual en Excel")
        if not df_alumnos.empty:
            col_sel1, col_sel2 = st.columns([3, 1])

            opciones_alumnos_excel = {
                f"{r['nombre_alumno']} - CURP: {r['curp']}": r
                for _, r in df_alumnos.iterrows()
            }

            alumno_escogido = col_sel1.selectbox(
                "Selecciona un alumno para generar su Excel rellenado:",
                list(opciones_alumnos_excel.keys()),
            )

            r_al = opciones_alumnos_excel[alumno_escogido]

            # Parsear fecha de nacimiento guardada "DD/MM/AAAA"
            fnac_parts = (
                str(r_al["fecha_nacimiento"]).split("/")
                if r_al["fecha_nacimiento"]
                else ["", "", ""]
            )
            d_nac = fnac_parts[0] if len(fnac_parts) > 0 else ""
            m_nac = fnac_parts[1] if len(fnac_parts) > 1 else ""
            a_nac = fnac_parts[2] if len(fnac_parts) > 2 else ""

            docs_entregados_list = (
                str(r_al["docs_entregados"]).split(",")
                if r_al["docs_entregados"]
                else []
            )

            bytes_excel_admin = generar_excel_alumno(
                r_al["nombre_alumno"],
                d_nac,
                m_nac,
                a_nac,
                r_al["edad"],
                r_al["lugar_nacimiento"],
                r_al["sexo"],
                r_al["celular_alumno"],
                r_al["correo"],
                r_al["red_social"],
                r_al["curp"],
                r_al["secundaria"],
                r_al["cct"],
                r_al["promedio"],
                r_al["carrera"],
                r_al["turno"],
                r_al["estatus"],
                r_al["observaciones"],
                r_al["nombre_tutor"],
                r_al["domicilio"],
                r_al["celular_tutor"],
                r_al["tel_casa"],
                r_al["tel_emergencia"],
                r_al["ocupacion"],
                docs_entregados_list,
            )

            if bytes_excel_admin:
                nom_clean = "".join(
                    c for c in r_al["nombre_alumno"] if c.isalnum() or c == " "
                ).strip()
                col_sel2.download_button(
                    label="📄 Descargar Excel Solicitud",
                    data=bytes_excel_admin,
                    file_name=f"SOLICITUD_{nom_clean}_{r_al['curp']}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                col_sel2.warning("No se encontró la plantilla de Excel.")

        # Vaciar base de datos
        with st.expander("⚠️ Opción Temporal: Vaciar / Eliminar Base de Datos"):
            st.warning(
                "Esta acción borrará TODOS los registros de la base de datos local."
            )
            confirmar_vaciar = st.checkbox(
                "Entiendo que esta acción es irreversible"
            )
            if st.button("🗑️ VACIAR BASE DE DATOS AHORA") and confirmar_vaciar:
                vaciar_db()
                st.success("✅ Base de datos vaciada con éxito.")
                st.rerun()

        st.markdown("---")
        st.dataframe(df_alumnos, use_container_width=True)

        # EDICIÓN COMPLETA DE CAMPOS
        st.markdown("---")
        st.subheader("✏️ Editar Cualquier Campo de un Alumno Registrado")
        if not df_alumnos.empty:
            opciones = {
                f"{r['nombre_alumno']} (CURP: {r['curp']})": r["id"]
                for _, r in df_alumnos.iterrows()
            }
            sel_alumno = st.selectbox(
                "Selecciona alumno a modificar:", list(opciones.keys())
            )
            id_sel = opciones[sel_alumno]
            row_sel = df_alumnos[df_alumnos["id"] == id_sel].iloc[0]

            with st.form("form_edit_admin_completo"):
                st.markdown("##### Datos del Alumno")
                col_e1, col_e2 = st.columns(2)
                e_nombre = col_e1.text_input(
                    "Nombre:", value=row_sel["nombre_alumno"]
                )
                e_curp = col_e2.text_input("CURP:", value=row_sel["curp"])

                col_e3, col_e4, col_e5 = st.columns(3)
                e_fnac = col_e3.text_input(
                    "Fecha Nac. (DD/MM/AAAA):",
                    value=str(row_sel["fecha_nacimiento"] or ""),
                )
                e_edad = col_e4.text_input(
                    "Edad:", value=str(row_sel["edad"] or "")
                )
                e_sexo = col_e5.selectbox(
                    "Sexo:",
                    ["FEMENINO", "MASCULINO"],
                    index=0 if row_sel["sexo"] == "FEMENINO" else 1,
                )

                col_e6, col_e7, col_e8 = st.columns(3)
                e_lugar_nac = col_e6.text_input(
                    "Lugar Nacimiento:",
                    value=str(row_sel["lugar_nacimiento"] or ""),
                )
                e_cel_alum = col_e7.text_input(
                    "Celular Alumno:",
                    value=str(row_sel["celular_alumno"] or ""),
                )
                e_correo = col_e8.text_input(
                    "Correo:", value=str(row_sel["correo"] or "")
                )

                e_red_social = st.text_input(
                    "Red Social:", value=str(row_sel["red_social"] or "")
                )

                st.markdown("##### Datos Académicos")
                col_ea1, col_ea2, col_ea3 = st.columns(3)
                e_secundaria = col_ea1.text_input(
                    "Secundaria:", value=str(row_sel["secundaria"] or "")
                )
                e_cct = col_ea2.text_input(
                    "CCT:", value=str(row_sel["cct"] or "")
                )
                e_promedio = col_ea3.text_input(
                    "Promedio:", value=str(row_sel["promedio"] or "")
                )

                col_ea4, col_ea5, col_ea6 = st.columns(3)
                e_carrera = col_ea4.text_input(
                    "Carrera:", value=str(row_sel["carrera"] or "")
                )

                opt_turno = ["MATUTINO", "VESPERTINO"]
                idx_turno = (
                    opt_turno.index(row_sel["turno"])
                    if row_sel["turno"] in opt_turno
                    else 0
                )
                e_turno = col_ea5.selectbox(
                    "Turno:", opt_turno, index=idx_turno
                )

                opt_estatus = [
                    "ASIGNADO",
                    "CAMBIO",
                    "OTRO RESULTADO",
                    "SIN PROCESO",
                ]
                idx_estatus = (
                    opt_estatus.index(row_sel["estatus"])
                    if row_sel["estatus"] in opt_estatus
                    else 0
                )
                e_estatus = col_ea6.selectbox(
                    "Estatus:", opt_estatus, index=idx_estatus
                )

                e_obs = st.text_input(
                    "Observaciones:", value=str(row_sel["observaciones"] or "")
                )

                st.markdown("##### Datos del Tutor")
                col_et1, col_et2 = st.columns(2)
                e_tutor = col_et1.text_input(
                    "Nombre Tutor:", value=str(row_sel["nombre_tutor"] or "")
                )
                e_domicilio = col_et2.text_input(
                    "Domicilio:", value=str(row_sel["domicilio"] or "")
                )

                col_et3, col_et4, col_et5, col_et6 = st.columns(4)
                e_cel_tut = col_et3.text_input(
                    "Celular Tutor:", value=str(row_sel["celular_tutor"] or "")
                )
                e_tel_casa = col_et4.text_input(
                    "Tel. Casa:", value=str(row_sel["tel_casa"] or "")
                )
                e_tel_emerg = col_et5.text_input(
                    "Tel. Emergencia:",
                    value=str(row_sel["tel_emergencia"] or ""),
                )
                e_ocupacion = col_et6.text_input(
                    "Ocupación:", value=str(row_sel["ocupacion"] or "")
                )

                e_docs = st.text_input(
                    "Documentos Entregados (números separados por coma):",
                    value=str(row_sel["docs_entregados"] or ""),
                )

                btn_actualizar_todo = st.form_submit_button(
                    "🔄 Guardar Cambios en BBDD Local"
                )

            if btn_actualizar_todo:
                datos_actualizados = (
                    e_nombre,
                    e_curp,
                    e_fnac,
                    e_edad,
                    e_sexo,
                    e_lugar_nac,
                    e_cel_alum,
                    e_correo,
                    e_red_social,
                    e_secundaria,
                    e_cct,
                    e_promedio,
                    e_carrera,
                    e_turno,
                    e_estatus,
                    e_obs,
                    e_tutor,
                    e_domicilio,
                    e_cel_tut,
                    e_tel_casa,
                    e_tel_emerg,
                    e_ocupacion,
                    e_docs,
                )
                actualizar_en_db(id_sel, datos_actualizados)
                st.success("✅ Registro actualizado correctamente.")
                st.rerun()