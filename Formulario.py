import os
import sqlite3
from io import BytesIO
import openpyxl
import pandas as pd
import streamlit as st

# Configuración inicial de la página
st.set_page_config(
    page_title="Sistema de Inscripción 2026", page_icon="📝", layout="wide"
)

PASSWORD_ADMIN = st.secrets.get("PASSWORD_ADMIN", "admin123")
DB_FILE = "inscripciones.db"


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

        col_f1, col_f2, col_f3, col_f4 = st.columns(4)
        dia_nac = col_f1.text_input("Día (DD):")
        mes_nac = col_f2.text_input("Mes (MM):")
        anio_nac = col_f3.text_input("Año (AAAA):")
        edad = col_f4.text_input("Años (Edad):")

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
                # Modificar plantilla Excel si existe localmente
                if os.path.exists("SOLIC INSCRIP NVO 2026.xlsx"):
                    wb = openpyxl.load_workbook("SOLIC INSCRIP NVO 2026.xlsx")
                    sheet = wb.active

                    sheet["G2"] = nombre_alumno
                    sheet["I5"], sheet["M5"], sheet["Q5"], sheet["U5"] = (
                        dia_nac,
                        mes_nac,
                        anio_nac,
                        edad,
                    )
                    sheet["I8"] = lugar_nac
                    if sexo == "FEMENINO":
                        sheet["R8"] = "X"
                    elif sexo == "MASCULINO":
                        sheet["R9"] = "X"

                    sheet["I10"], sheet["I11"], sheet["J12"] = (
                        celular_alumno,
                        correo,
                        red_social,
                    )
                    sheet["E14"], sheet["I15"], sheet["H16"], sheet["T16"], (
                        sheet["H17"]
                    ) = (
                        curp,
                        secundaria,
                        cct,
                        promedio,
                        carrera,
                    )

                    if turno == "MATUTINO":
                        sheet["H18"] = "X"
                    elif turno == "VESPERTINO":
                        sheet["M18"] = "X"

                    map_estatus = {
                        "ASIGNADO": "D19",
                        "CAMBIO": "H19",
                        "OTRO RESULTADO": "N19",
                        "SIN PROCESO": "V19",
                    }
                    if estatus in map_estatus:
                        sheet[map_estatus[estatus]] = "X"

                    sheet["Q18"] = observaciones
                    sheet["G22"], sheet["H24"], sheet["I27"], (
                        sheet["I28"]
                    ), sheet["I29"], sheet["I30"] = (
                        nombre_tutor,
                        domicilio,
                        celular_tutor,
                        tel_casa,
                        tel_emergencia,
                        ocupacion,
                    )

                    docs_map = {
                        doc1: "J32",
                        doc2: "J33",
                        doc3: "J34",
                        doc4: "J35",
                        doc5: "J36",
                        doc6: "J37",
                        doc7: "J38",
                        doc8: "V32",
                        doc9: "V34",
                        doc10: "V36",
                        doc11: "V37",
                    }
                    lista_docs = []
                    for idx, (check, cell) in enumerate(docs_map.items(), 1):
                        if check:
                            sheet[cell] = "X"
                            lista_docs.append(str(idx))

                    # Guardar archivo Excel en una carpeta local opcional o memoria
                    nombre_limpio = "".join(
                        c for c in nombre_alumno if c.isalnum() or c == " "
                    ).strip()
                    nombre_excel = f"SOLICITUD_{nombre_limpio}_{curp}.xlsx"
                    os.makedirs("solicitudes_excel", exist_ok=True)
                    wb.save(os.path.join("solicitudes_excel", nombre_excel))
                else:
                    lista_docs = []

                # Guardar registro en la Base de Datos SQLite local
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

                st.success("✅ ¡Inscripción guardada correctamente en la BBDD local!")

            except sqlite3.IntegrityError:
                st.error(f"⚠️ La CURP **{curp}** ya se encuentra registrada.")
            except Exception as e:
                st.error(f"Error al guardar la solicitud: {e}")

# --- TAB 2: PANEL ADMINISTRADOR ---
with tab2:
    st.title("🔒 Panel Administrador")
    pass_input = st.text_input("Contraseña:", type="password")

    if pass_input == PASSWORD_ADMIN:
        
        # --- NUBES DE ARCHIVO / SUBIR BBDD MANUAL ---
        with st.expander("📤 Cargar / Reemplazar Base de Datos (.db)", expanded=False):
            st.info("Sube una copia previa de tu archivo `inscripciones.db` para consultar o actualizar los datos.")
            uploaded_db = st.file_uploader("Selecciona un archivo .db local:", type=["db"])
            if uploaded_db is not None:
                if st.button("🔄 Cargar esta Base de Datos ahora"):
                    with open(DB_FILE, "wb") as f:
                        f.write(uploaded_db.getbuffer())
                    st.cache_resource.clear()
                    st.success("✅ Base de datos cargada y actualizada con éxito.")
                    st.rerun()

        st.markdown("---")
        df_alumnos = obtener_alumnos()
        st.write(f"**Total de alumnos registrados:** {len(df_alumnos)}")

        col_d1, col_d2 = st.columns(2)

        # Botón 1: Descargar Padrón Consolidado en Excel
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

        # Botón 2: Descargar la BBDD SQLite local (.db)
        if os.path.exists(DB_FILE):
            with open(DB_FILE, "rb") as f_db:
                col_d2.download_button(
                    label="🗄️ Descargar Copia BBDD (.db)",
                    data=f_db.read(),
                    file_name="inscripciones.db",
                    mime="application/x-sqlite3",
                )

        # Vaciar base de datos
        with st.expander("⚠️ Opción Temporal: Vaciar / Eliminar Base de Datos"):
            st.warning("Esta acción borrará TODOS los registros de la base de datos local.")
            confirmar_vaciar = st.checkbox("Entiendo que esta acción es irreversible")
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