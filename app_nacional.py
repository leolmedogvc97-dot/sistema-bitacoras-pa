import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font
import json
import os
import random
import zipfile
import pydeck as pdk
from io import BytesIO
from datetime import datetime, date

st.set_page_config(page_title="Sistema Nacional de Bitácoras - Procuraduría Agraria", layout="wide")

# Inyección de estilos CSS personalizados (Colores Institucionales / Guinda Morena)
st.markdown("""
    <style>
    h1, h2, h3 {
        color: #6B1D2F !important;
    }
    .stButton>button {
        background-color: #6B1D2F !important;
        color: white !important;
        border-radius: 6px !important;
        border: none !important;
        font-weight: bold;
    }
    .stButton>button:hover {
        background-color: #8A253D !important;
        color: white !important;
    }
    [data-testid="stFormSubmitButton"]>button {
        background-color: #6B1D2F !important;
        color: white !important;
        border-radius: 6px !important;
        font-weight: bold;
    }
    [data-testid="stFormSubmitButton"]>button:hover {
        background-color: #8A253D !important;
    }
    section[data-testid="stSidebar"] {
        background-color: #faf6f7;
        border-right: 1px solid #e0d0d3;
    }
    .frase-agraria {
        background-color: #fcf5f6;
        border-left: 5px solid #6B1D2F;
        padding: 12px 18px;
        border-radius: 4px;
        margin-bottom: 20px;
        font-style: italic;
        color: #4a2c33;
    }
    </style>
""", unsafe_allow_html=True)

# Rutas persistentes y carpetas de almacenamiento
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
USUARIOS_FILE = os.path.join(BASE_DIR, "usuarios.json")
REGISTROS_FILE = os.path.join(BASE_DIR, "registros.json")
SOLICITUDES_FILE = os.path.join(BASE_DIR, "solicitudes_gasolina.json")
INCIDENCIAS_FILE = os.path.join(BASE_DIR, "incidencias_mecanicas.json")
PROGRAMACION_FILE = os.path.join(BASE_DIR, "programacion_combustible.json")
AUDIT_FILE = os.path.join(BASE_DIR, "audit_log.json")
FOTOS_DIR = os.path.join(BASE_DIR, "fotos_perfil")
LOGO_FILE = os.path.join(BASE_DIR, "logo_pa.png")
MUN_FILE = os.path.join(BASE_DIR, "MUNICIPIOS_202606.xlsx")
LOC_FILE = os.path.join(BASE_DIR, "localidades.xlsx")
os.makedirs(FOTOS_DIR, exist_ok=True)

# Detección automática y flexible de la plantilla Excel institucional en la raíz
excel_files = [os.path.join(BASE_DIR, f) for f in os.listdir(BASE_DIR) if f.endswith('.xlsx') and not f.startswith("BITACORAS_OFICIALES") and f not in ["MUNICIPIOS_202606.xlsx", "LOCALIDADES_202606.xlsx", "localidades.xlsx"]]
PLANTILLA_EXCEL = excel_files[0] if excel_files else os.path.join(BASE_DIR, "Bitacora_Actualizada_Formula (2).xlsx")

CODS_MUNICIPIOS = {
    "Toluca": {"lat": 19.2826, "lon": -99.6556},
    "Naucalpan de Juárez": {"lat": 19.4781, "lon": -99.2363},
    "Naucalpan": {"lat": 19.4781, "lon": -99.2363},
    "Atizapán de Zaragoza": {"lat": 19.5667, "lon": -99.2500},
    "Tlalnepantla de Baz": {"lat": 19.5400, "lon": -99.1900},
    "Ecatepec de Morelos": {"lat": 19.6010, "lon": -99.0558},
    "Texcoco": {"lat": 19.5100, "lon": -98.8800},
    "Valle de Bravo": {"lat": 19.1944, "lon": -100.1333},
    "Tenancingo": {"lat": 18.9667, "lon": -99.5833},
    "Atlacomulco": {"lat": 19.8000, "lon": -99.8833},
    "Morelia": {"lat": 19.7027, "lon": -101.1924},
    "Uruapan": {"lat": 19.4128, "lon": -102.0647}
}

ESTADOS_REPUBLICA = [
    "Aguascalientes", "Baja California", "Baja California Sur", "Campeche", 
    "Chiapas", "Chihuahua", "Ciudad de México", "Coahuila", "Colima", 
    "Durango", "Estado de México", "Guanajuato", "Guerrero", "Hidalgo", 
    "Jalisco", "Michoacán", "Morelos", "Nayarit", "Nuevo León", "Oaxaca", 
    "Puebla", "Querétaro", "Quintana Roo", "San Luis Potosí", "Sinaloa", 
    "Sonora", "Tabasco", "Tamaulipas", "Tlaxcala", "Veracruz", "Yucatán", "Zacatecas"
]

JEFATURAS_RESIDENCIA = [
    "RESIDENCIA NAUCALPAN",
    "RESIDENCIA TOLUCA",
    "RESIDENCIA ATLACOMULCO",
    "RESIDENCIA TEXCOCO",
    "RESIDENCIA VALLE DE BRAVO",
    "RESIDENCIA TENANCINGO",
    "RESIDENCIA MORELIA"
]

ROLES_SISTEMA = [
    "Organizador Agrario (Operador)", 
    "Analista de Información", 
    "Administrador de Vehículos (Estatal)",
    "Jefe de Residencia", 
    "Administrador Estatal", 
    "Administrador Nacional"
]

FRASES_AGRARIAS = [
    "“La tierra no pertenece al hombre; el hombre pertenece a la tierra.” — Jefe Seattle",
    "“El agricultor es la única persona en el mundo que gasta dinero en esperar y arriesgarse a cosechar.” — E.W. Howe",
    "“La agricultura es la base de todas las manufacturas y el sustento de la nación.” — Daniel Webster",
    "“Quien planta un árbol cree en el mañana.” — Audrey Hepburn",
    "“El campo da pan a todos, y la justicia agraria da paz y dignidad a nuestra gente.”",
    "“Cultivar la tierra es servir a la patria con las manos y el corazón.”",
    "“La tierra generosa recompensa siempre el esfuerzo honesto de quien la trabaja.”",
    "“El verdadero desarrollo del país florece desde sus raíces ejidales y comunales.”",
    "“Sembrar conciencia en el campo es cosechar soberanía y bienestar social.”",
    "“La labor del agrónomo y del servidor agrario es transformar la esperanza en frutos tangibles.”"
]

def cargar_usuarios():
    usuarios_base = {
        "victor.olmedo@pa.gob.mx": {"nombre": "VÍCTOR LEONARDO OLMEDO GONZALEZ", "pass": "Leonardo", "licencia": "0101P3402484l", "rol": "Administrador Nacional", "estado": "Estado de México", "jefatura": "RESIDENCIA NAUCALPAN", "jefe_residencia": "N/A", "foto": "", "activo": True},
        "marichuy.duarte@pa.gob.mx": {"nombre": "MARICHUY DUARTE SALAMANCA", "pass": "Marichuy2026", "licencia": "12345678", "rol": "Administrador Nacional", "estado": "Estado de México", "jefatura": "RESIDENCIA TOLUCA", "jefe_residencia": "N/A", "foto": "", "activo": True},
        "marichuy@pa.gob.mx": {"nombre": "MARICHUY", "pass": "Marichuy2026", "licencia": "0000000000000", "rol": "Administrador Nacional", "estado": "Michoacán", "jefatura": "RESIDENCIA MORELIA", "jefe_residencia": "N/A", "foto": "", "activo": True},
        "jcarlos.patino@pa.gob.mx": {"nombre": "JUAN CARLOS PATINO PEREZ", "pass": "Perez", "licencia": "", "rol": "Administrador Nacional", "estado": "Estado de México", "jefatura": "RESIDENCIA TOLUCA", "jefe_residencia": "NINGUNO", "foto": "", "activo": True},
        "arosales@pa.gob.mx": {"nombre": "ARISVE LESEIE ESLAVA ROSALES", "pass": "Arosales", "licencia": "", "rol": "Administrador Nacional", "estado": "Estado de México", "jefatura": "RESIDENCIA TOLUCA", "jefe_residencia": "NINGUNO", "foto": "", "activo": True},
        "carmen.lara@pa.gob.mx": {"nombre": "CARMEN LARA", "pass": "Carmen", "licencia": "", "rol": "Administrador Nacional", "estado": "Aguascalientes", "jefatura": "RESIDENCIA NAUCALPAN", "jefe_residencia": "ING. GABRIEL ESTRADA", "foto": "", "activo": True},
        "josue.rodriguez@pa.gob.mx": {"nombre": "JOSUE RODRIGUEZ", "pass": "Josue", "licencia": "", "rol": "Administrador Nacional", "estado": "Aguascalientes", "jefatura": "RESIDENCIA NAUCALPAN", "jefe_residencia": "ING. GABRIEL ESTRADA", "foto": "", "activo": True},
        "nsalgado@pa.gob.mx": {"nombre": "NANCY SALGADO ANTUNEZ", "pass": "Nancy", "licencia": "", "rol": "Analista de Información", "estado": "Estado de México", "jefatura": "RESIDENCIA NAUCALPAN", "jefe_residencia": "ING. GABRIEL ESTRADA", "foto": "", "activo": True},
        "esperanza.ramos@pa.gob.mx": {"nombre": "ESPERANZA WENDY RAMOS RODRIGUEZ", "pass": "Wendy", "licencia": "", "rol": "Administrador Nacional", "estado": "Estado de México", "jefatura": "RESIDENCIA NAUCALPAN", "jefe_residencia": "ADAN JIMENEZ", "foto": "", "activo": True},
        "carlos.javg.96@gmail.com": {"nombre": "CARLOS JAVIER GALVEZ GONZALEZ", "pass": "Carlos", "licencia": "", "rol": "Administrador Nacional", "estado": "Estado de México", "jefatura": "RESIDENCIA NAUCALPAN", "jefe_residencia": "ADAN JIMENEZ", "foto": "", "activo": True},
        "dehnny.vasquez@pa.gob.mx": {"nombre": "DEHNNY VAZQUEZ FLORES", "pass": "Dehnny", "licencia": "", "rol": "Organizador Agrario (Operador)", "estado": "Estado de México", "jefatura": "RESIDENCIA TOLUCA", "jefe_residencia": "", "foto": "", "activo": True}
    }
    if os.path.exists(USUARIOS_FILE):
        try:
            with open(USUARIOS_FILE, "r", encoding="utf-8") as f:
                usuarios_guardados = json.load(f)
                for email, data in usuarios_base.items():
                    if email not in usuarios_guardados:
                        usuarios_guardados[email] = data
                return usuarios_guardados
        except:
            pass
    return usuarios_base

def guardar_usuarios(usuarios_dict):
    with open(USUARIOS_FILE, "w", encoding="utf-8") as f:
        json.dump(usuarios_dict, f, ensure_ascii=False, indent=4)

def cargar_registros_acumulados():
    if os.path.exists(REGISTROS_FILE):
        try:
            with open(REGISTROS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            pass
    return []

def guardar_registros_acumulados(registros_list):
    with open(REGISTROS_FILE, "w", encoding="utf-8") as f:
        json.dump(registros_list, f, ensure_ascii=False, indent=4)

def cargar_solicitudes():
    if os.path.exists(SOLICITUDES_FILE):
        try:
            with open(SOLICITUDES_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            pass
    return []

def guardar_solicitudes(solicitudes_list):
    with open(SOLICITUDES_FILE, "w", encoding="utf-8") as f:
        json.dump(solicitudes_list, f, ensure_ascii=False, indent=4)

def cargar_incidencias():
    if os.path.exists(INCIDENCIAS_FILE):
        try:
            with open(INCIDENCIAS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            pass
    return []

def guardar_incidencias(incidencias_list):
    with open(INCIDENCIAS_FILE, "w", encoding="utf-8") as f:
        json.dump(incidencias_list, f, ensure_ascii=False, indent=4)

def registrar_auditoria(accion, detalle):
    log_entry = {
        "FECHA_HORA": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "USUARIO": st.session_state.get("current_email", "Sistema"),
        "ROL": st.session_state.get("current_rol", "N/A"),
        "ACCION": accion,
        "DETALLE": detalle
    }
    logs = []
    if os.path.exists(AUDIT_FILE):
        try:
            with open(AUDIT_FILE, "r", encoding="utf-8") as f:
                logs = json.load(f)
        except:
            pass
    logs.append(log_entry)
    with open(AUDIT_FILE, "w", encoding="utf-8") as f:
        json.dump(logs, f, ensure_ascii=False, indent=4)

OPCIONES_GASOLINA = {
    "🔴 1/4 de Tanque": "1/4",
    "🟡 1/2 Tanque": "1/2",
    "🟢 3/4 de Tanque": "3/4",
    "🟢 Tanque Lleno (1/1)": "1/1",
    "🔴 Reserva / Vacío (V)": "V"
}

if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
if "current_email" not in st.session_state:
    st.session_state["current_email"] = ""
if "current_user" not in st.session_state:
    st.session_state["current_user"] = ""
if "current_licencia" not in st.session_state:
    st.session_state["current_licencia"] = ""
if "current_rol" not in st.session_state:
    st.session_state["current_rol"] = ""
if "current_estado" not in st.session_state:
    st.session_state["current_estado"] = ""
if "current_jefatura" not in st.session_state:
    st.session_state["current_jefatura"] = ""
if "current_jefe_residencia" not in st.session_state:
    st.session_state["current_jefe_residencia"] = ""
if "frase_dia" not in st.session_state:
    st.session_state["frase_dia"] = random.choice(FRASES_AGRARIAS)

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
                    st.session_state["current_rol"] = usuarios_actuales[email_input].get("rol", "Organizador Agrario (Operador)")
                    st.session_state["current_estado"] = usuarios_actuales[email_input].get("estado", "Estado de México")
                    st.session_state["current_jefatura"] = usuarios_actuales[email_input].get("jefatura", "RESIDENCIA NAUCALPAN")
                    st.session_state["current_jefe_residencia"] = usuarios_actuales[email_input].get("jefe_residencia", "N/A")
                    st.session_state["frase_dia"] = random.choice(FRASES_AGRARIAS)
                    registrar_auditoria("INICIO DE SESION", f"Acceso exitoso de {email_input}")
                    st.success("¡Acceso concedido! Cargando sistema...")
                    st.rerun()
            else:
                st.error("⚠️ Usuario o contraseña incorrectos. Verifica tus datos.")
    st.stop()

# --- APLICACIÓN PRINCIPAL ---
col_m_logo, col_m_title = st.columns([0.15, 2.85])
with col_m_logo:
    if os.path.exists(LOGO_FILE):
        st.image(LOGO_FILE, width=60)
with col_m_title:
    st.title("Sistema Nacional de Control Vehicular y Bitácoras")

st.markdown(f'<div class="frase-agraria">🌾 <b>Reflexión Agraria del Día:</b> {st.session_state["frase_dia"]}</div>', unsafe_allow_html=True)
st.markdown("---")

if os.path.exists(LOGO_FILE):
    st.sidebar.image(LOGO_FILE, use_container_width=True)

current_email_key = st.session_state['current_email']
usuarios_actuales_sidebar = cargar_usuarios()
rol_actual = st.session_state.get("current_rol", "Organizador Agrario (Operador)")

st.sidebar.title("📌 Menú de Navegación")
if rol_actual == "Administrador Nacional":
    modulos_disponibles = ["Módulo de Captura (Recorrido)", "Mi Perfil / Foto", "Solicitud de Recurso de Gasolina", "Reporte de Incidencias en Ruta", "Programación de Combustible y Viáticos", "Panel de Administración y Auditoría"]
elif rol_actual in ["Administrador Estatal", "Administrador de Vehículos (Estatal)", "Jefe de Residencia", "Analista de Información"]:
    modulos_disponibles = ["Módulo de Captura (Recorrido)", "Mi Perfil / Foto", "Solicitud de Recurso de Gasolina", "Reporte de Incidencias en Ruta", "Programación de Combustible y Viáticos", "Panel de Supervisión (Estatal/Residencia)"]
else:
    modulos_disponibles = ["Módulo de Captura (Recorrido)", "Mi Perfil / Foto", "Solicitud de Recurso de Gasolina", "Reporte de Incidencias en Ruta", "Programación de Combustible y Viáticos"]

perfil = st.sidebar.radio("Selecciona tu módulo:", modulos_disponibles)

st.sidebar.markdown("---")
with st.sidebar.expander("👤 Sesión Activa (Ver Datos)", expanded=False):
    foto_actual = usuarios_actuales_sidebar.get(current_email_key, {}).get("foto", "")
    if foto_actual and os.path.exists(foto_actual):
        st.image(foto_actual, width=100)
    else:
        st.info("Sin foto de perfil.")
    st.write(f"**Usuario:** {st.session_state['current_user']}")
    st.write(f"**Correo:** {st.session_state['current_email']}")
    st.write(f"**Rol:** {st.session_state.get('current_rol', 'Organizador')}")
    st.write(f"**Estado:** {st.session_state.get('current_estado', 'N/A')}")
    st.write(f"**Jefatura:** {st.session_state.get('current_jefatura', 'N/A')}")
    st.write(f"**Jefe Resp.:** {st.session_state.get('current_jefe_residencia', 'N/A')}")

st.sidebar.markdown("---")
if st.sidebar.button("🚪 Cerrar Sesión"):
    registrar_auditoria("CIERRE DE SESION", f"Cierre de sesión de {current_email_key}")
    st.session_state["logged_in"] = False
    st.session_state["current_email"] = ""
    st.rerun()

if perfil == "Mi Perfil / Foto":
    st.subheader("🖼️ Configuración de Perfil y Adscripción Institucional")
    st.markdown("Actualiza tu fotografía de perfil y verifica tu pertenencia territorial, jefatura y jefe de residencia correspondiente.")
    
    usuarios_dict_perfil = cargar_usuarios()
    datos_u_actual = usuarios_dict_perfil.get(current_email_key, {})
    
    st.info(f"""
    * **Estado de Adscripción:** {datos_u_actual.get('estado', 'N/A')}
    * **Jefatura de Residencia:** {datos_u_actual.get('jefatura', 'N/A')}
    * **Jefe de Residencia Asignado:** {datos_u_actual.get('jefe_residencia', 'N/A')}
    * **Rol en la Red:** {datos_u_actual.get('rol', 'N/A')}
    """)
    
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
                    registrar_auditoria("ACTUALIZACION FOTO", f"Actualización de foto para {current_email_key}")
                    st.success("¡Fotografía de perfil actualizada con éxito!")
            else:
                st.warning("⚠️ Selecciona un archivo de imagen válido antes de guardar.")

    st.markdown("---")
    st.subheader("📊 Resumen de Actividad y Uso de la Red (Accesos y Bitácoras)")
    st.markdown("Auditoría general del sistema: Conoce qué usuarios han ingresado, cuántas veces han accedido, cuántas bitácoras oficiales han generado y cuántos registros acumulan.")
    
    logs_audit = []
    if os.path.exists(AUDIT_FILE):
        try:
            with open(AUDIT_FILE, "r", encoding="utf-8") as f:
                logs_audit = json.load(f)
        except:
            pass
            
    registros_todos = cargar_registros_acumulados()
    
    if usuarios_dict_perfil:
        resumen_uso_dict = {}
        for email_u, info_u in usuarios_dict_perfil.items():
            resumen_uso_dict[email_u] = {
                "Correo": email_u,
                "Nombre": info_u.get("nombre", "N/A"),
                "Rol": info_u.get("rol", "N/A"),
                "Estado": info_u.get("estado", "N/A"),
                "Ingresos (Logins)": 0,
                "Último Acceso": "Sin registro",
                "Bitácoras Generadas": 0,
                "Registros Guardados": 0
            }
            
        for reg in registros_todos:
            c_org = reg.get("CORREO_ORGANIZADOR", "")
            if c_org in resumen_uso_dict:
                resumen_uso_dict[c_org]["Registros Guardados"] += 1
                
        for log in logs_audit:
            usr_log = log.get("USUARIO", "").strip().lower()
            acc_log = log.get("ACCION", "")
            fh_log = log.get("FECHA_HORA", "")
            
            if usr_log in resumen_uso_dict:
                if acc_log == "INICIO DE SESION":
                    resumen_uso_dict[usr_log]["Ingresos (Logins)"] += 1
                    resumen_uso_dict[usr_log]["Último Acceso"] = fh_log
                elif acc_log == "GENERAR BITACORAS":
                    resumen_uso_dict[usr_log]["Bitácoras Generadas"] += 1
                    
        df_uso_red = pd.DataFrame(list(resumen_uso_dict.values()))
        st.dataframe(df_uso_red, use_container_width=True)
    else:
        st.info("No hay usuarios registrados en el sistema.")

elif perfil == "Solicitud de Recurso de Gasolina":
    st.error("⚠️ **[PRUEBA]** - Módulo en Fase de Pruebas Operativas")
    st.subheader("⛽ Solicitud de Recurso de Gasolina para Comisión Oficial")
    st.markdown("Apartado institucional para la gestión, solicitud y aprobación de asignación presupuestal de combustible.")
    
    todas_sols_notif = cargar_solicitudes()
    sols_mias = [s for s in todas_sols_notif if s.get("CORREO") == current_email_key]
    for s in sols_mias:
        est = s.get("ESTATUS", "")
        if est == "APROBADO":
            st.success(f"🎉 ¡Notificación! Tu solicitud para el destino **{s.get('DESTINO')}** ha sido **APROBADA**.")
        elif est == "RECHAZADO":
            st.error(f"⚠️ Aviso: Tu solicitud para el destino **{s.get('DESTINO')}** fue **RECHAZADA**.")

    estado_usuario_actual = st.session_state.get("current_estado", "Estado de México")
    jefatura_actual = st.session_state.get("current_jefatura", "RESIDENCIA NAUCALPAN")
    solicitante_actual = st.session_state.get("current_user", "")
        
    with st.form("form_solicitud_gasolina"):
        col_s1, col_s2 = st.columns(2)
        with col_s1:
            fecha_solicitud = st.date_input("Fecha de Solicitud", value=date.today())
            analista_solicitante = st.text_input("Analista de Información / Solicitante", value=solicitante_actual, max_chars=300)
            residencia_adscripcion = st.text_input("Jefatura de Residencia", value=jefatura_actual, max_chars=300)
            funcionario_comisionado = st.text_input("Funcionario / Organizador Asignado a Comisión", value="", max_chars=300)
        with col_s2:
            municipio_destino = st.text_input(f"Municipio de Destino ({estado_usuario_actual}) - Libre", value="", max_chars=300, placeholder="Escribe el municipio libremente...")
            localidad_destino = st.text_input("Localidad / Poblado de Destino - Libre", value="", max_chars=300, placeholder="Escribe la localidad libremente...")
            vehiculo_asignado = st.text_input("Vehículo Oficial Asignado - Libre", value="", max_chars=300, placeholder="Ej. Nissan Versa, PickUp...")
            placas_vehiculo = st.text_input("Placas del Vehículo", value="", max_chars=300)
            
        st.markdown("---")
        col_s3, col_s4, col_s5 = st.columns(3)
        with col_s3:
            f_inicio_com = st.date_input("Fecha Inicio de Comisión", value=date.today())
        with col_s4:
            f_fin_com = st.date_input("Fecha Término de Comisión", value=date.today())
        with col_s5:
            monto_solicitado = st.number_input("Monto Solicitado para Combustible ($ MXN)", min_value=0.0, value=1500.0, step=50.0)
            
        col_s6, col_s7 = st.columns(2)
        with col_s6:
            oficio_asociado = st.text_input("Número de Oficio de Comisión", value="", max_chars=300)
        with col_s7:
            motivo_comision = st.text_input("Motivo / Descripción de la Comisión Oficial", value="", max_chars=300)
            
        btn_enviar_solicitud = st.form_submit_button("📥 Enviar Solicitud de Recurso")
        if btn_enviar_solicitud:
            if not funcionario_comisionado.strip() or not motivo_comision.strip():
                st.error("⚠️ Debes completar el nombre del funcionario comisionado y el motivo.")
            else:
                nueva_solicitud = {
                    "FECHA SOLICITUD": fecha_solicitud.strftime("%d/%m/%Y"),
                    "SOLICITANTE": analista_solicitante,
                    "ROL SOLICITANTE": rol_actual,
                    "JEFATURA": residencia_adscripcion,
                    "ESTADO": estado_usuario_actual,
                    "FUNCIONARIO": funcionario_comisionado.upper(),
                    "DESTINO": f"{localidad_destino.strip().upper()}, {municipio_destino.strip().upper()}",
                    "VEHICULO": vehiculo_asignado.strip().upper(),
                    "PLACAS": placas_vehiculo.upper(),
                    "FECHA INICIO": f_inicio_com.strftime("%d/%m/%Y"),
                    "FECHA TERMINO": f_fin_com.strftime("%d/%m/%Y"),
                    "MONTO SOLICITADO": monto_solicitado,
                    "OFICIO": oficio_asociado if oficio_asociado else "N/A",
                    "MOTIVO": motivo_comision,
                    "ESTATUS": "PENDIENTE DE APROBACIÓN",
                    "CORREO": current_email_key
                }
                lista_sols = cargar_solicitudes()
                lista_sols.append(nueva_solicitud)
                guardar_solicitudes(lista_sols)
                registrar_auditoria("SOLICITUD GASOLINA", f"Solicitud de ${monto_solicitado} para {funcionario_comisionado} en {estado_usuario_actual}")
                st.success("✅ Solicitud de recurso enviada y registrada correctamente.")

    solicitudes_historial = cargar_solicitudes()
    if len(solicitudes_historial) > 0:
        st.markdown("---")
        st.subheader("📋 Historial y Gestión de Solicitudes de Gasolina")
        df_sols = pd.DataFrame(solicitudes_historial)
        if rol_actual == "Analista de Información":
            df_sols_filtrado = df_sols[df_sols['ESTADO'] == estado_usuario_actual]
        elif rol_actual in ["Jefe de Residencia", "Administrador de Vehículos (Estatal)"]:
            df_sols_filtrado = df_sols[df_sols['JEFATURA'] == jefatura_actual]
        elif rol_actual in ["Administrador Estatal", "Administrador Nacional"]:
            df_sols_filtrado = df_sols
        else:
            df_sols_filtrado = df_sols[df_sols['CORREO'] == current_email_key]
        st.dataframe(df_sols_filtrado, use_container_width=True)
        
        if rol_actual in ["Jefe de Residencia", "Administrador de Vehículos (Estatal)", "Administrador Estatal", "Administrador Nacional"]:
            st.markdown("---")
            st.subheader("⚙️ Panel de Aprobación de Solicitudes")
            for idx, sol in enumerate(solicitudes_historial):
                if sol.get("ESTATUS") == "PENDIENTE DE APROBACIÓN":
                    col_info_s, col_btn_ap, col_btn_re = st.columns([5, 1, 1])
                    with col_info_s:
                        st.write(f"**{sol['FUNCIONARIO']}** | Destino: {sol['DESTINO']} | Monto: ${sol['MONTO SOLICITADO']:,.2f} MXN ({sol['JEFATURA']})")
                    with col_btn_ap:
                        if st.button("✅ Aprobar", key=f"aprobar_{idx}"):
                            solicitudes_historial[idx]["ESTATUS"] = "APROBADO"
                            guardar_solicitudes(solicitudes_historial)
                            registrar_auditoria("APROBAR SOLICITUD", f"Aprobación de recurso para {sol['FUNCIONARIO']}")
                            st.rerun()
                    with col_btn_re:
                        if st.button("❌ Rechazar", key=f"rechazar_{idx}"):
                            solicitudes_historial[idx]["ESTATUS"] = "RECHAZADO"
                            guardar_solicitudes(solicitudes_historial)
                            registrar_auditoria("RECHAZAR SOLICITUD", f"Rechazo de recurso para {sol['FUNCIONARIO']}")
                            st.rerun()

elif perfil == "Reporte de Incidencias en Ruta":
    st.error("⚠️ **[PRUEBA]** - Módulo en Fase de Pruebas Operativas")
    st.subheader("🚨 Reporte y Gestión de Incidencias y Fallas Mecánicas")
    st.markdown("Módulo para registrar y dar seguimiento a averías, fallas mecánicas o imprevistos durante las comisiones oficiales.")
    
    estado_usuario_actual = st.session_state.get("current_estado", "Estado de México")
    jefatura_actual = st.session_state.get("current_jefatura", "RESIDENCIA NAUCALPAN")
    usuario_actual_nombre = st.session_state.get("current_user", "")
    
    with st.form("form_incidencia"):
        col_i1, col_i2 = st.columns(2)
        with col_i1:
            fecha_inc = st.date_input("Fecha de la Incidencia", value=date.today())
            placas_inc = st.text_input("Placas del Vehículo Afectado", value="", max_chars=300)
            tipo_falla = st.selectbox("Tipo de Incidencia / Falla", [
                "🔴 Falla Mecánica Mayor (Motor/Transmisión)",
                "🟡 Falla Eléctrica / Batería",
                "🟠 Pinchadura de Neumático / Llanta",
                "🔵 Calentamiento / Sistema de Enfriamiento",
                "⚪ Otro imprevisto en ruta"
            ])
        with col_i2:
            municipio_inc = st.text_input("Municipio / Lugar de la Avería", value="", max_chars=300)
            gravedad_inc = st.selectbox("Nivel de Urgencia", ["Baja", "Media", "Alta (Requiere Grúa/Auxilio Inmediato)"])
            
        descripcion_inc = st.text_area("Descripción detallada de la incidencia mecánica o imprevisto")
        btn_enviar_inc = st.form_submit_button("🚨 Enviar Reporte de Incidencia")
        if btn_enviar_inc:
            if not municipio_inc.strip() or not descripcion_inc.strip() or not placas_inc.strip():
                st.error("⚠️ Debes completar las placas, el municipio y la descripción de la falla.")
            else:
                nueva_inc = {
                    "FECHA": fecha_inc.strftime("%d/%m/%Y"),
                    "CONDUCTOR": usuario_actual_nombre,
                    "CORREO": current_email_key,
                    "JEFATURA": jefatura_actual,
                    "ESTADO": estado_usuario_actual,
                    "PLACAS": placas_inc.strip().upper(),
                    "TIPO": tipo_falla,
                    "MUNICIPIO": municipio_inc.upper(),
                    "URGENCIA": gravedad_inc,
                    "DESCRIPCION": descripcion_inc,
                    "ESTATUS": "PENDIENTE DE ATENCIÓN"
                }
                lista_incs = cargar_incidencias()
                lista_incs.append(nueva_inc)
                guardar_incidencias(lista_incs)
                registrar_auditoria("REPORTE INCIDENCIA", f"Falla reportada en vehículo {placas_inc} ({tipo_falla})")
                st.success("✅ Incidencia reportada exitosamente. Se ha notificado al Administrador de Vehículos y Jefatura.")

    todas_incs = cargar_incidencias()
    if todas_incs:
        st.markdown("---")
        st.subheader("📋 Historial de Incidencias en Ruta")
        df_incs = pd.DataFrame(todas_incs)
        if rol_actual in ["Jefe de Residencia", "Administrador de Vehículos (Estatal)"]:
            df_incs_filtrado = df_incs[df_incs['ESTADO'] == estado_usuario_actual]
        elif rol_actual == "Administrador Estatal":
            df_incs_filtrado = df_incs[df_incs['ESTADO'] == estado_usuario_actual]
        elif rol_actual == "Administrador Nacional":
            df_incs_filtrado = df_incs
        else:
            df_incs_filtrado = df_incs[df_incs['CORREO'] == current_email_key]
        st.dataframe(df_incs_filtrado, use_container_
