import streamlit as st
import pandas as pd
import openpyxl
import json
import os
from io import BytesIO
from datetime import datetime, date

st.set_page_config(page_title="Sistema Nacional de Bitácoras - Procuraduría Agraria", layout="wide")

# Archivos persistentes y carpetas de almacenamiento
USUARIOS_FILE = "usuarios.json"
FOTOS_DIR = "fotos_perfil"
LOGO_FILE = "logo_pa.png"
MUN_FILE = "MUNICIPIOS_202606.xlsx"
LOC_FILE = "LOCALIDADES_202606.xlsx"
os.makedirs(FOTOS_DIR, exist_ok=True)

# Listado oficial de los 32 Estados de la República Mexicana
ESTADOS_REPUBLICA = [
    "Aguascalientes", "Baja California", "Baja California Sur", "Campeche", 
    "Chiapas", "Chihuahua", "Ciudad de México", "Coahuila", "Colima", 
    "Durango", "Estado de México", "Guanajuato", "Guerrero", "Hidalgo", 
    "Jalisco", "Michoacán", "Morelos", "Nayarit", "Nuevo León", "Oaxaca", 
    "Puebla", "Querétaro", "Quintana Roo", "San Luis Potosí", "Sinaloa", 
    "Sonora", "Tabasco", "Tamaulipas", "Tlaxcala", "Veracruz", "Yucatán", "Zacatecas"
]

# Carga de catálogos geográficos nacionales
@st.cache_data
def cargar_catalogos_geograficos():
    muns_df = pd.read_excel(MUN_FILE) if os.path.exists(MUN_FILE) else pd.DataFrame()
    locs_df = pd.read_excel(LOC_FILE) if os.path.exists(LOC_FILE) else pd.DataFrame()
    return muns_df, locs_df

muns_global, locs_global = cargar_catalogos_geograficos()

def obtener_municipios_estado(estado_nombre):
    if muns_global.empty or estado_nombre not in ESTADOS_REPUBLICA:
        return []
    efe_key = ESTADOS_REPUBLICA.index(estado_nombre) + 1
    muns = muns_global[muns_global['EFE_KEY'] == efe_key]['MUNICIPIO'].tolist()
    return sorted(muns)

def obtener_localidades_municipio(estado_nombre, municipio_nombre):
    if muns_global.empty or locs_global.empty or estado_nombre not in ESTADOS_REPUBLICA:
        return []
    efe_key = ESTADOS_REPUBLICA.index(estado_nombre) + 1
    mun_row = muns_global[(muns_global['EFE_KEY'] == efe_key) & (muns_global['MUNICIPIO'] == municipio_nombre)]
    if mun_row.empty:
        return []
    cat_key = mun_row['CATALOG_KEY'].values[0]
    locs = locs_global[(locs_global['EFE_KEY'] == efe_key) & (locs_global['MUN_KEY'] == cat_key)]['LOCALIDAD'].tolist()
    return sorted(locs)

def cargar_usuarios():
    usuarios_base = {
        "victor.olmedo@pa.gob.mx": {
            "nombre": "VÍCTOR LEONARDO OLMEDO GONZALEZ",
            "pass": "Leonardo",
            "licencia": "0101P3402484l",
            "rol": "Administrador Nacional",
            "estado": "Estado de México",
            "foto": "",
            "activo": True
        },
        "marichuy.duarte@pa.gob.mx": {
            "nombre": "MARICHUY DUARTE SALAMANCA",
            "pass": "Marichuy2026",
            "licencia": "12345678",
            "rol": "Administrador Nacional",
            "estado": "Estado de México",
            "foto": "",
            "activo": True
        },
        "marichuy@pa.gob.mx": {
            "nombre": "MARICHUY",
            "pass": "Marichuy2026",
            "licencia": "0000000000000",
            "rol": "Administrador Nacional",
            "estado": "Michoacán",
            "foto": "",
            "activo": True
        }
    }
    
    if os.path.exists(USUARIOS_FILE):
        try:
            with open(USUARIOS_FILE, "r", encoding="utf-8") as f:
                usuarios_guardados = json.load(f)
                for email_admin in ["victor.olmedo@pa.gob.mx", "marichuy.duarte@pa.gob.mx", "marichuy@pa.gob.mx"]:
                    if email_admin in usuarios_guardados:
                        usuarios_guardados[email_admin]["rol"] = "Administrador Nacional"
                        if "activo" not in usuarios_guardados[email_admin]:
                            usuarios_guardados[email_admin]["activo"] = True
                        if "estado" not in usuarios_guardados[email_admin]:
                            usuarios_guardados[email_admin]["estado"] = usuarios_base[email_admin]["estado"]
                    elif email_admin in usuarios_base:
                        usuarios_guardados[email_admin] = usuarios_base[email_admin]
                return usuarios_guardados
        except:
            pass
            
    return usuarios_base

def guardar_usuarios(usuarios_dict):
    with open(USUARIOS_FILE, "w", encoding="utf-8") as f:
        json.dump(usuarios_dict, f, ensure_ascii=False, indent=4)

# Semáforos visuales de gasolina
OPCIONES_GASOLINA = {
    "🔴 1/4 de Tanque": "1/4",
    "🟡 1/2 Tanque": "1/2",
    "🟢 3/4 de Tanque": "3/4",
    "🟢 Tanque Lleno (1/1)": "1/1",
    "🔴 Reserva / Vacío (V)": "V"
}

# Inicialización de estados en sesión
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
if "usuarios" not in st.session_state:
    st.session_state["usuarios"] = cargar_usuarios()
if "registros_acumulados" not in st.session_state:
    st.session_state["registros_acumulados"] = []

# --- PANTALLA DE LOGIN ---
if not st.session_state["logged_in"]:
    col_l_logo, col_l_title = st.columns([0.15, 2.85])
    with col_l_logo:
        if os.path.exists(LOGO_FILE):
            st.image(LOGO_FILE, width=60)
    with col_l_title:
        st.title("Acceso al Sistema Nacional de Bitácoras")
    st.markdown("---")
    
    col_l1, col_l2, col_l3 = st.columns([1, 2, 1])
    with col_l2:
        st.subheader("Iniciar Sesión")
        email_input = st.text_input("Correo Electrónico Institucional", max_chars=300).strip().lower()
        pass_input = st.text_input("Contraseña", type="password", max_chars=300)
        
        if st.button("🔑 Ingresar al Sistema", use_container_width=True):
            usuarios_actuales = cargar_usuarios()
            if email_input in usuarios_actuales and usuarios_actuales[email_input]["pass"] == pass_input:
                if not usuarios_actuales[email_input].get("activo", True):
                    st.error("⚠️ Tu cuenta se encuentra desactivada. Contacta al Administrador Nacional.")
                else:
                    st.session_state["logged_in"] = True
                    st.session_state["current_email"] = email_input
                    st.session_state["current_user"] = usuarios_actuales[email_input]["nombre"]
                    st.session_state["current_licencia"] = usuarios_actuales[email_input]["licencia"]
                    st.session_state["current_rol"] = usuarios_actuales[email_input].get("rol", "Operador de Residencia")
                    st.session_state["current_estado"] = usuarios_actuales[email_input].get("estado", "Estado de México")
                    st.session_state["current_foto"] = usuarios_actuales[email_input].get("foto", "")
                    st.success("¡Acceso concedido! Cargando sistema...")
                    st.rerun()
            else:
                st.error("⚠️ Usuario or contraseña incorrectos. Verifica tus datos.")
    st.stop()

# --- APLICACIÓN PRINCIPAL (UNA VEZ LOGUEADO) ---
col_m_logo, col_m_title = st.columns([0.15, 2.85])
with col_m_logo:
    if os.path.exists(LOGO_FILE):
        st.image(LOGO_FILE, width=60)
with col_m_title:
    st.title("Sistema Nacional de Control Vehicular y Bitácoras")
st.markdown("---")

# Panel lateral: Logotipo, foto de perfil y datos de sesión
if os.path.exists(LOGO_FILE):
    st.sidebar.image(LOGO_FILE, use_container_width=True)

st.sidebar.title("👤 Sesión Activa")
current_email_key = st.session_state['current_email']
usuarios_actuales_sidebar = cargar_usuarios()
foto_actual = usuarios_actuales_sidebar.get(current_email_key, {}).get("foto", "")

if foto_actual and os.path.exists(foto_actual):
    st.sidebar.image(foto_actual, width=120)
else:
    st.sidebar.info("Sin foto de perfil asignada.")

st.sidebar.write(f"**Usuario:** {st.session_state['current_user']}")
st.sidebar.write(f"**Correo:** {st.session_state['current_email']}")
st.sidebar.write(f"**Estado Adscripción:** {st.session_state.get('current_estado', 'N/A')}")
st.sidebar.write(f"**Rol:** {st.session_state.get('current_rol', 'Operador')}")

# Control de módulos según el rol del usuario
rol_actual = st.session_state.get("current_rol", "Operador de Residencia")
if rol_actual in ["Administrador Nacional", "Administrador Estatal"]:
    modulos_disponibles = ["Operador de Residencia", "Mi Perfil / Foto", "Administrador Nacional (Sede)"]
else:
    modulos_disponibles = ["Operador de Residencia", "Mi Perfil / Foto"]

perfil = st.sidebar.selectbox("Selecciona tu módulo:", modulos_disponibles)

if st.sidebar.button("🚪 Cerrar Sesión"):
    st.session_state["logged_in"] = False
    st.rerun()

if perfil == "Mi Perfil / Foto":
    st.subheader("🖼️ Configuración de Perfil y Fotografía de Adscripción")
    st.markdown("Actualiza tu fotografía de perfil institucional y verifica tu estado de adscripción.")
    
    usuarios_dict_perfil = cargar_usuarios()
    datos_u_actual = usuarios_dict_perfil.get(current_email_key, {})
    
    st.info(f"**Estado de Adscripción Actual:** {datos_u_actual.get('estado', 'Estado de México')}")
    st.markdown("*(El Estado de Adscripción y perfil general pueden ser actualizados por el Administrador Nacional desde su panel).*")
    
    with st.form("form_foto_perfil"):
        archivo_foto = st.file_uploader("Seleccionar imagen de perfil (JPG, PNG)", type=["jpg", "jpeg", "png"])
        btn_subir_foto = st.form_submit_button("💾 Guardar Fotografía")
        
        if btn_subir_foto:
            if archivo_foto is not None:
                extension = archivo_foto.name.split(".")[-1]
                nombre_archivo = f"{current_email_key.replace('@', '_').replace('.', '_')}.{extension}"
                ruta_destino = os.path.join(FOTOS_DIR, nombre_archivo)
                
                with open(ruta_destino, "wb") as f:
                    f.write(archivo_foto.getbuffer())
                
                all_users = cargar_usuarios()
                if current_email_key in all_users:
                    all_users[current_email_key]["foto"] = ruta_destino
                    guardar_usuarios(all_users)
                    st.session_state["current_foto"] = ruta_destino
                    st.success("¡Fotografía de perfil actualizada con éxito! Vuelve a cargar o navega por el sistema para verla.")
            else:
                st.warning("⚠️ Selecciona un archivo de imagen válido antes de guardar.")

elif perfil == "Operador de Residencia":
    st.subheader("📝 Módulo de Captura por Día - Residencia")
    estado_usuario_actual = st.session_state.get("current_estado", "Estado de México")
    st.markdown(f"Ingresa los datos de tu recorrido diario. Los menús de **Municipios y Localidades** están filtrados automáticamente para **{estado_usuario_actual}**.")
    
    # Obtener municipios y localidades para el estado de adscripción del usuario
    lista_municipios = obtener_municipios_estado(estado_usuario_actual)
    if not lista_municipios:
        lista_municipios = ["Toluca", "Naucalpan de Juárez", "Metepec"] # Fallback
        
    with st.form("form_captura_dia"):
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            fecha = st.date_input("Fecha de registro del uso del vehículo")
            municipio = st.selectbox(f"Municipio ({estado_usuario_actual})", lista_municipios)
            
            # Obtener localidades dinámicamente según el municipio seleccionado
            lista_localidades = obtener_localidades_municipio(estado_usuario_actual, municipio)
            if not lista_localidades:
                lista_localidades = ["Cabecera Municipal", "Comisión Oficial"]
            poblado = st.selectbox("Poblado / Localidad", lista_localidades)
            
            folio_ciia = st.text_input("Folio CIIA", value="", max_chars=300)
        with col2:
            h_salida = st.text_input("Hora de Salida (Formato 24h, ej. 09:00)", value="09:00", max_chars=300)
            km_inicial = st.number_input("KM Inicial / Salida", min_value=0.0, value=0.0, step=1.0)
            h_llegada = st.text_input("Hora de Llegada (Formato 24h, ej. 17:00)", value="17:00", max_chars=300)
            km_final = st.number_input("KM Final / Llegada", min_value=0.0, value=0.0, step=1.0)
        with col3:
            residencia = st.selectbox("Área de Adscripción", [
                f"RESIDENCIA {estado_usuario_actual.upper()}", 
                "RESIDENCIA NAUCALPAN", 
                "RESIDENCIA TOLUCA", 
                "RESIDENCIA ATLACOMULCO", 
                "RESIDENCIA TEXCOCO", 
                "RESIDENCIA VALLE DE BRAVO", 
                "RESIDENCIA TENANCINGO"
            ])
            vehiculo = st.selectbox("Tipo de Vehículo", ["NISSAN VERSA", "PickUp", "Estacas"])
            placas = st.text_input("Placas", value="MGX-543-A", max_chars=300)
            licencia = st.text_input("No. De Licencia", value=st.session_state["current_licencia"], max_chars=300)
        with col4:
            usuario = st.text_input("Usuario Responsable", value=st.session_state["current_user"], max_chars=300)
            dotacion = st.number_input("Dotación de Gasolina ($)", min_value=0.0, value=200.0, step=1.0)
            
            gas_salida_label = st.selectbox("Gasolina de Salida (Nivel)", list(OPCIONES_GASOLINA.keys()))
            gas_llegada_label = st.selectbox("Gasolina de Llegada (Nivel)", list(OPCIONES_GASOLINA.keys()))
        
        st.markdown("---")
        col_o1, col_o2, col_o3 = st.columns(3)
        with col_o1:
            oficio_num = st.text_input("Oficio Número (Opcional)", value="", max_chars=300)
        with col_o2:
            oficio_fecha = st.text_input("Oficio Fecha (Opcional)", value="", max_chars=300)
        with col_o3:
            observaciones = st.text_input("Observaciones / Ruta", value="", max_chars=300)
        
        guardar_dia = st.form_submit_button("💾 Guardar Día")
        
        if guardar_dia:
            if h_salida.strip() == h_llegada.strip():
                st.error("⚠️ Error: La Hora de Salida y la Hora de Llegada no pueden ser iguales en formato 24 horas.")
            elif km_final <= km_inicial and km_final != 0.0:
                st.warning("⚠️ Aviso: El KM Final es menor o igual al KM Inicial. Verifica tus datos.")
            else:
                recorrido = km_final - km_inicial
                nuevo_reg = {
                    "FECHA COMPLETA": fecha.strftime("%d/%m/%Y"),
                    "MES": fecha.strftime("%B").upper(),
                    "MUNICIPIO": municipio,
                    "POBLADO": poblado,
                    "folio CIIA": folio_ciia,
                    "HORA DE SALIDA": h_salida,
                    "KM INICIAL / Km de Salida": km_inicial,
                    "RECORRIDO": recorrido,
                    "HORA DE LLEGADA": h_llegada,
                    "KM FINAL / Km de Llegada": km_final,
                    "GASTO COMBUSTI": dotacion,
                    "Gasolina de Salida": OPCIONES_GASOLINA[gas_salida_label],
                    "Gasolina de Llegada": OPCIONES_GASOLINA[gas_llegada_label],
                    "Dotación de Gasolina(LLENAR GASTO DE COMBUSTIBLE)": dotacion,
                    "Oficio Numero": oficio_num if oficio_num else None,
                    "Oficio Fecha": oficio_fecha if oficio_fecha else None,
                    "observaciones": observaciones if observaciones else None,
                    "Usuario Responsable": usuario,
                    "Áreas de Adscripción": residencia,
                    "Tipo de Vehículo": vehiculo,
                    "Placas": placas,
                    "No. De Licencia": licencia,
                    "ESTADO_ADSCRIPCION": estado_usuario_actual
                }
                st.session_state["registros_acumulados"].append(nuevo_reg)
                st.success(f"✅ ¡Día {fecha.strftime('%d/%m/%Y')} en {municipio} ({poblado}) guardado correctamente!")

    if len(st.session_state["registros_acumulados"]) > 0:
        st.markdown("---")
        st.subheader("📋 Días Guardados (Acumulado)")
        df_acumulado = pd.DataFrame(st.session_state["registros_acumulados"])
        st.dataframe(df_acumulado, use_container_width=True)
        
        col_acc1, col_acc2 = st.columns(2)
        with col_acc1:
            if st.button("🗑️ Limpiar lista y empezar de nuevo"):
                st.session_state["registros_acumulados"] = []
                st.rerun()
        
        with col_acc2:
            if st.button("🚀 Guardar y Generar 3 Bitácoras Definitivas"):
                try:
                    wb = openpyxl.load_workbook("Prueba unificación.xlsx")
                    ws = wb["BASE_DE_DATOS"]
                    
                    if ws.max_row >= 2:
                        ws.delete_rows(2, ws.max_row)
                    
                    for i, reg in enumerate(st.session_state["registros_acumulados"]):
                        row_idx = 2 + i
                        fila_datos = [
                            reg["FECHA COMPLETA"],
                            f'=UPPER(TEXT(A{row_idx}, "MMMM"))',
                            reg["MUNICIPIO"],
                            reg["POBLADO"],
                            reg["folio CIIA"],
                            reg["HORA DE SALIDA"],
                            reg["KM INICIAL / Km de Salida"],
                            reg["RECORRIDO"],
                            reg["HORA DE LLEGADA"],
                            reg["KM FINAL / Km de Llegada"],
                            reg["GASTO COMBUSTI"],
                            reg["Gasolina de Salida"],
                            reg["Gasolina de Llegada"],
                            reg["Dotación de Gasolina(LLENAR GASTO DE COMBUSTIBLE)"],
                            reg["Oficio Numero"],
                            reg["Oficio Fecha"],
                            reg["observaciones"],
                            reg["Usuario Responsable"],
                            reg["Áreas de Adscripción"],
                            reg["Tipo de Vehículo"],
                            reg["Placas"],
                            reg["No. De Licencia"]
                        ]
                        ws.append(fila_datos)
                    
                    output = BytesIO()
                    wb.save(output)
                    output.seek(0)
                    
                    st.success("¡Archivo generado con éxito y listo para descarga!")
                    st.download_button(
                        label="⬇️ Descargar Archivo Definitivo (Incluye las 3 Bitácoras)",
                        data=output,
                        file_name="BITACORAS_OFICIALES_DEFINITIVAS.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"Error al generar el archivo definitivo: {e}")

elif perfil == "Administrador Nacional (Sede)":
    st.subheader("📊 Panel de Administración Nacional y Gestión de Usuarios")
    
    # --- PESTAÑAS ADMINISTRATIVAS ---
    tab_reg_user, tab_edit_user, tab_ctrl_user, tab_resumen_auditoria = st.tabs([
        "➕ Alta de Usuario", 
        "✏️ Editar Usuario", 
        "👥 Control y Estatus", 
        "📈 Resumen Ejecutivo y Auditoría"
    ])
    
    with tab_reg_user:
        with st.form("form_nuevo_usuario"):
            col_u1, col_u2 = st.columns(2)
            with col_u1:
                c_email = st.text_input("Correo Electrónico (Usuario)", max_chars=300)
                c_nombre = st.text_input("Nombre Completo (Mayúsculas)", max_chars=300)
                c_pass = st.text_input("Contraseña Asignada", type="password", max_chars=300)
            with col_u2:
                c_licencia = st.text_input("Número de Licencia de Conducir", max_chars=300)
                c_estado = st.selectbox("Estado de Adscripción", ESTADOS_REPUBLICA)
                c_rol = st.selectbox("Rol en el Sistema", ["Operador de Residencia", "Administrador Estatal", "Administrador Nacional"])
            
            btn_crear = st.form_submit_button("Registrar Usuario")
            if btn_crear:
                if c_email and c_nombre and c_pass:
                    usuarios_actuales = cargar_usuarios()
                    email_limpio = c_email.strip().lower()
                    usuarios_actuales[email_limpio] = {
                        "nombre": c_nombre.strip().upper(),
                        "pass": c_pass.strip(),
                        "licencia": c_licencia.strip(),
                        "estado": c_estado,
                        "rol": c_rol,
                        "foto": "",
                        "activo": True
                    }
                    guardar_usuarios(usuarios_actuales)
                    st.session_state["usuarios"] = usuarios_actuales
                    st.success(f"¡Usuario {c_nombre} ({c_rol}) con adscripción en {c_estado} registrado exitosamente!")
                    st.rerun()
                else:
                    st.error("⚠️ Por favor completa los campos obligatorios (Correo, Nombre y Contraseña).")

    with tab_edit_user:
        usuarios_actuales_edit = cargar_usuarios()
        lista_emails = list(usuarios_actuales_edit.keys())
        if lista_emails:
            email_a_editar = st.selectbox("Selecciona el correo del usuario a modificar", lista_emails)
            if email_a_editar:
                u_data = usuarios_actuales_edit[email_a_editar]
                with st.form("form_editar_usuario"):
                    col_e1, col_e2 = st.columns(2)
                    with col_e1:
                        e_nombre = st.text_input("Nombre Completo (Mayúsculas)", value=u_data.get("nombre", ""), max_chars=300)
                        e_pass = st.text_input("Contraseña Asignada", value=u_data.get("pass", ""), type="password", max_chars=300)
                        e_licencia = st.text_input("Número de Licencia de Conducir", value=u_data.get("licencia", ""), max_chars=300)
                    with col_e2:
                        estado_actual = u_data.get("estado", "Estado de México")
                        idx_estado = ESTADOS_REPUBLICA.index(estado_actual) if estado_actual in ESTADOS_REPUBLICA else 0
                        e_estado = st.selectbox("Estado de Adscripción", ESTADOS_REPUBLICA, index=idx_estado)
                        
                        roles_disponibles = ["Operador de Residencia", "Administrador Estatal", "Administrador Nacional"]
                        rol_actual_u = u_data.get("rol", "Operador de Residencia")
                        idx_rol = roles_disponibles.index(rol_actual_u) if rol_actual_u in roles_disponibles else 0
                        e_rol = st.selectbox("Rol en el Sistema", roles_disponibles, index=idx_rol)
                    
                    btn_actualizar = st.form_submit_button("💾 Guardar Cambios de Usuario")
                    if btn_actualizar:
                        usuarios_actuales_edit[email_a_editar]["nombre"] = e_nombre.strip().upper()
                        usuarios_actuales_edit[email_a_editar]["pass"] = e_pass.strip()
                        usuarios_actuales_edit[email_a_editar]["licencia"] = e_licencia.strip()
                        usuarios_actuales_edit[email_a_editar]["estado"] = e_estado
                        usuarios_actuales_edit[email_a_editar]["rol"] = e_rol
                        guardar_usuarios(usuarios_actuales_edit)
                        st.success(f"¡Información de {email_a_editar} actualizada exitosamente!")
                        st.rerun()
        else:
            st.info("No hay usuarios registrados para editar.")

    with tab_ctrl_user:
        st.subheader("📋 Control, Estatus y Eliminación de Usuarios en la Red Nacional")
        st.markdown("Listado general de la red con controles de activación, desactivación y eliminación:")
        
        usuarios_actuales_tabla = cargar_usuarios()
        
        # Encabezado estilo Excel
        st.markdown(
            """
            <div style="background-color: #343a40; color: white; padding: 10px 15px; border-radius: 6px; font-weight: bold; display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px; font-size: 13px;">
                <span style="flex: 1.5; text-align: center;">ACCIONES</span>
                <span style="flex: 2;">CORREO</span>
                <span style="flex: 2.5;">NOMBRE</span>
                <span style="flex: 2;">ROL</span>
                <span style="flex: 1.5;">ESTADO ADS.</span>
                <span style="flex: 1; text-align: right;">ESTATUS</span>
            </div>
            """,
            unsafe_allow_html=True
        )
        
        for email, datos in usuarios_actuales_tabla.items():
            estado_activo = datos.get("activo", True)
            color_fondo = "#d4edda" if estado_activo else "#e2e3e5" # Verde claro si activo, gris si inactivo
            texto_estado = "🟢 Activo" if estado_activo else "🔴 Desactivado"
            
            col_btns, col_info = st.columns([1.6, 8.4])
            
            with col_btns:
                b1, b2, b3 = st.columns(3)
                with b1:
                    if st.button("👍", key=f"activar_{email}", help="Activar"):
                        usuarios_actuales_tabla[email]["activo"] = True
                        guardar_usuarios(usuarios_actuales_tabla)
                        st.rerun()
                with b2:
                    if st.button("👎", key=f"desactivar_{email}", help="Desactivar"):
                        usuarios_actuales_tabla[email]["activo"] = False
                        guardar_usuarios(usuarios_actuales_tabla)
                        st.rerun()
                with b3:
                    if st.button("🗑️", key=f"eliminar_{email}", help="Eliminar perfil"):
                        if email == st.session_state["current_email"]:
                            st.error("No puedes eliminar tu propia cuenta.")
                        else:
                            del usuarios_actuales_tabla[email]
                            guardar_usuarios(usuarios_actuales_tabla)
                            st.rerun()
            
            with col_info:
                st.markdown(
                    f"""
                    <div style="background-color: {color_fondo}; padding: 8px 12px; border-radius: 6px; margin-bottom: 6px; border: 1px solid #c0c0c0; display: flex; justify-content: space-between; align-items: center; font-size: 13px;">
                        <span style="flex: 2; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;"><b>{email}</b></span>
                        <span style="flex: 2.5; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">{datos.get('nombre')}</span>
                        <span style="flex: 2; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">{datos.get('rol')}</span>
                        <span style="flex: 1.5; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">{datos.get('estado')}</span>
                        <span style="flex: 1; text-align: right; white-space: nowrap;"><b>{texto_estado}</b></span>
                    </div>
                    """, 
                    unsafe_allow_html=True
                )

    with tab_resumen_auditoria:
        st.subheader("📈 Panel de Resumen Ejecutivo y Auditoría (Control Vehicular)")
        st.markdown("Consulta global del uso general de vehículos por rango de fechas, kilómetros recorridos, consumo de combustible y trazabilidad geográfica.")
        
        # Simulación de registros globales o acumulados en sesión
        if len(st.session_state["registros_acumulados"]) > 0:
            df_global = pd.DataFrame(st.session_state["registros_acumulados"])
            df_global['FECHA_DT'] = pd.to_datetime(df_global['FECHA COMPLETA'], format='%d/%m/%Y').dt.date
            
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                f_inicio = st.date_input("Fecha Inicial", value=df_global['FECHA_DT'].min())
            with col_f2:
                f_fin = st.date_input("Fecha Final", value=df_global['FECHA_DT'].max())
                
            # Filtrar por rango de fechas
            df_filtrado = df_global[(df_global['FECHA_DT'] >= f_inicio) & (df_global['FECHA_DT'] <= f_fin)]
            
            if not df_filtrado.empty:
                st.markdown("---")
                # Métricas principales
                total_km = df_filtrado['RECORRIDO'].sum()
                total_gas = df_filtrado['GASTO COMBUSTI'].sum()
                total_viajes = len(df_filtrado)
                
                m1, m2, m3 = st.columns(3)
                m1.metric("KM Totales Recorridos", f"{total_km:,.1f} km")
                m2.metric("Gasto Total de Combustible", f"${total_gas:,.2f} MXN")
                m3.metric("Comisiones Registradas", f"{total_viajes}")
                
                st.markdown("---")
                st.subheader("🗺️ Trazabilidad Geográfica y Municipios Visitados")
                
                # Agrupación y listado cronológico de municipios visitados
                df_resumen_mun = df_filtrado.groupby(['FECHA COMPLETA', 'MUNICIPIO', 'POBLADO', 'Usuario Responsable', 'ESTADO_ADSCRIPCION']).agg({
                    'RECORRIDO': 'sum',
                    'GASTO COMBUSTI': 'sum'
                }).reset_index().sort_values(by='FECHA COMPLETA')
                
                st.dataframe(df_resumen_mun, use_container_width=True)
                
                # Resumen único de municipios visitados en el periodo
                municipios_visitados = df_filtrado['MUNICIPIO'].unique().tolist()
                st.markdown(f"**Municipios únicos visitados en el periodo ({len(municipios_visitados)}):** " + ", ".join(municipios_visitados))
                
            else:
                st.warning("⚠️ No se encontraron registros en el rango de fechas seleccionado.")
        else:
            st.info("Aún no hay registros de recorridos acumulados en la sesión actual para generar el reporte de auditoría. Realiza capturas en el módulo de operador.")
