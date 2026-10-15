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

# Diccionarios de meses en español
MESES_MAP_ES = {
    1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril', 5: 'Mayo', 6: 'Junio',
    7: 'Julio', 8: 'Agosto', 9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'
}

MESES_MAP_UPPER = {
    1: 'ENERO', 2: 'FEBRERO', 3: 'MARZO', 4: 'ABRIL', 5: 'MAYO', 6: 'JUNIO',
    7: 'JULIO', 8: 'AGOSTO', 9: 'SEPTIEMBRE', 10: 'OCTUBRE', 11: 'NOVIEMBRE', 12: 'DICIEMBRE'
}

# Rutas persistentes y carpetas de almacenamiento
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
USUARIOS_FILE = os.path.join(BASE_DIR, "usuarios.json")
REGISTROS_FILE = os.path.join(BASE_DIR, "registros.json")
PROGRAMACION_FILE = os.path.join(BASE_DIR, "programacion_combustible.json")
AUDIT_FILE = os.path.join(BASE_DIR, "audit_log.json")
FOTOS_DIR = os.path.join(BASE_DIR, "fotos_perfil")
LOGO_FILE = os.path.join(BASE_DIR, "logo_pa.png")
MUN_FILE = os.path.join(BASE_DIR, "MUNICIPIOS_202606.xlsx")
LOC_FILE = os.path.join(BASE_DIR, "localidades.xlsx")
os.makedirs(FOTOS_DIR, exist_ok=True)

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
    "Atlacomulco": {"lat": 19.8000, "lon": -99.8833}
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
    "RESIDENCIA TENANCINGO"
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
        "jcarlos.patino@pa.gob.mx": {"nombre": "JUAN CARLOS PATINO PEREZ", "pass": "Perez", "licencia": "", "rol": "Administrador Nacional", "estado": "Estado de México", "jefatura": "RESIDENCIA TOLUCA", "jefe_residencia": "NINGUNO", "foto": "", "activo": True},
        "arosales@pa.gob.mx": {"nombre": "ARISVE LESEIE ESLAVA ROSALES", "pass": "Arosales", "licencia": "", "rol": "Administrador Nacional", "estado": "Estado de México", "jefatura": "RESIDENCIA TOLUCA", "jefe_residencia": "NINGUNO", "foto": "", "activo": True},
        "carmen.lara@pa.gob.mx": {"nombre": "CARMEN LARA", "pass": "Carmen", "licencia": "", "rol": "Administrador Nacional", "estado": "Aguascalientes", "jefatura": "RESIDENCIA NAUCALPAN", "jefe_residencia": "ING. GABRIEL ESTRADA", "foto": "", "activo": True},
        "josue.rodriguez@pa.gob.mx": {"nombre": "JOSUE RODRIGUEZ", "pass": "Josue", "licencia": "", "rol": "Administrador Nacional", "estado": "Aguascalientes", "jefatura": "RESIDENCIA NAUCALPAN", "jefe_residencia": "ING. GABRIEL ESTRADA", "foto": "", "activo": True},
        "nsalgado@pa.gob.mx": {"nombre": "NANCY SALGADO ANTUNEZ", "pass": "Nancy", "licencia": "", "rol": "Analista de Información", "estado": "Estado de México", "jefatura": "RESIDENCIA NAUCALPAN", "jefe_residencia": "ING. GABRIEL ESTRADA", "foto": "", "activo": True},
        "esperanza.ramos@pa.gob.mx": {"nombre": "ESPERANZA WENDY RAMOS RODRIGUEZ", "pass": "Wendy", "licencia": "", "rol": "Administrador Nacional", "estado": "Estado de México", "jefatura": "RESIDENCIA NAUCALPAN", "jefe_residencia": "ADAN JIMENEZ", "foto": "", "activo": True},
        "carlos.javg.96@gmail.com": {"nombre": "CARLOS JAVIER GALVEZ GONZALEZ", "pass": "Carlos", "licencia": "", "rol": "Administrador Nacional", "estado": "Estado de México", "jefatura": "RESIDENCIA NAUCALPAN", "jefe_residencia": "ADAN JIMENEZ", "foto": "", "activo": True},
        "dehnny.vasquez@pa.gob.mx": {"nombre": "DEHNNY VAZQUEZ FLORES", "pass": "Dehnny", "licencia": "", "rol": "Organizador Agrario (Operador)", "estado": "Estado de México", "jefatura": "RESIDENCIA TOLUCA", "jefe_residencia": "", "foto": "", "activo": True},
        # --- USUARIOS DE RESIDENCIAS (SIN ARROBA) ---
        "toluca.residencia": {"nombre": "OSCAR MENDOZA RUIZ", "pass": "Toluca2026", "licencia": "TOL123456", "rol": "Organizador Agrario (Operador)", "estado": "Estado de México", "jefatura": "RESIDENCIA TOLUCA", "jefe_residencia": "ING. GABRIEL ESTRADA", "foto": "", "activo": True},
        "atlacomulco.residencia": {"nombre": "BRENDA LÓPEZ SÁNCHEZ", "pass": "Atlacomulco2026", "licencia": "ATL123456", "rol": "Organizador Agrario (Operador)", "estado": "Estado de México", "jefatura": "RESIDENCIA ATLACOMULCO", "jefe_residencia": "ING. GABRIEL ESTRADA", "foto": "", "activo": True},
        "naucalpan.residencia": {"nombre": "ADÁN JIMÉNEZ PÉREZ", "pass": "Naucalpan2026", "licencia": "NAU123456", "rol": "Organizador Agrario (Operador)", "estado": "Estado de México", "jefatura": "RESIDENCIA NAUCALPAN", "jefe_residencia": "ING. GABRIEL ESTRADA", "foto": "", "activo": True},
        "texcoco.residencia": {"nombre": "GUADALUPE RAMÍREZ FLORES", "pass": "Texcoco2026", "licencia": "TEX123456", "rol": "Organizador Agrario (Operador)", "estado": "Estado de México", "jefatura": "RESIDENCIA TEXCOCO", "jefe_residencia": "ING. GABRIEL ESTRADA", "foto": "", "activo": True},
        "vallebravo.residencia": {"nombre": "ROBERTO GÓMEZ CRUZ", "pass": "Valle2026", "licencia": "VAL123456", "rol": "Organizador Agrario (Operador)", "estado": "Estado de México", "jefatura": "RESIDENCIA VALLE DE BRAVO", "jefe_residencia": "ING. GABRIEL ESTRADA", "foto": "", "activo": True},
        "tenancingo.residencia": {"nombre": "MARÍA ELENA SOTO VILLA", "pass": "Tenancingo2026", "licencia": "TEN123456", "rol": "Organizador Agrario (Operador)", "estado": "Estado de México", "jefatura": "RESIDENCIA TENANCINGO", "jefe_residencia": "ING. GABRIEL ESTRADA", "foto": "", "activo": True}
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
        email_input = st.text_input("Usuario o Correo Institucional", max_chars=300).strip().lower()
        pass_input = st.text_input("Contraseña", type="password", max_chars=300)
        
        if st.button("🔑 Ingresar al Sistema", use_container_width=True):
            usuarios_actuales = cargar_usuarios()
            if email_input in usuarios_actuales and usuarios_actuales[email_input]["pass"] == pass_input:
                if not usuarios_actuales[email_input].get("activo", True):
                    st.error("Tu cuenta se encuentra desactivada. Contacta al Administrador Nacional.")
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
                st.error("Usuario o contraseña incorrectos. Verifica tus datos.")
    st.stop()

# --- APLICACIÓN PRINCIPAL ---
col_m_logo, col_m_title = st.columns([0.15, 2.85])
with col_m_logo:
    if os.path.exists(LOGO_FILE):
        st.image(LOGO_FILE, width=60)
with col_m_title:
    st.title("Sistema Nacional de Control Vehicular y Bitácoras")

st.markdown(f'<div class="frase-agraria"><b>Reflexión Agraria del Día:</b> {st.session_state["frase_dia"]}</div>', unsafe_allow_html=True)
st.markdown("---")

if os.path.exists(LOGO_FILE):
    st.sidebar.image(LOGO_FILE, use_container_width=True)

current_email_key = st.session_state['current_email']
usuarios_actuales_sidebar = cargar_usuarios()
rol_actual = st.session_state.get("current_rol", "Organizador Agrario (Operador)")
jefatura_sesion_actual = st.session_state.get("current_jefatura", "RESIDENCIA NAUCALPAN")

st.sidebar.title("Menú de Navegación")
if current_email_key == "victor.olmedo@pa.gob.mx":
    modulos_disponibles = ["Módulo de Captura (Recorrido)", "Mi Perfil / Foto", "Programación de Combustible y Viáticos", "Panel de Administración y Auditoría"]
elif rol_actual in ["Administrador Estatal", "Administrador de Vehículos (Estatal)", "Jefe de Residencia", "Analista de Información"]:
    modulos_disponibles = ["Módulo de Captura (Recorrido)", "Mi Perfil / Foto", "Programación de Combustible y Viáticos", "Panel de Supervisión (Estatal/Residencia)"]
else:
    modulos_disponibles = ["Módulo de Captura (Recorrido)", "Mi Perfil / Foto", "Programación de Combustible y Viáticos"]

perfil = st.sidebar.radio("Selecciona tu módulo:", modulos_disponibles)

st.sidebar.markdown("---")
with st.sidebar.expander("Sesión Activa (Ver Datos)", expanded=False):
    foto_actual = usuarios_actuales_sidebar.get(current_email_key, {}).get("foto", "")
    if foto_actual and os.path.exists(foto_actual):
        st.image(foto_actual, width=100)
    else:
        st.info("Sin foto de perfil.")
    st.write(f"**Usuario:** {st.session_state['current_user']}")
    st.write(f"**Credencial/Correo:** {st.session_state['current_email']}")
    st.write(f"**Rol:** {st.session_state.get('current_rol', 'Organizador')}")
    st.write(f"**Estado:** {st.session_state.get('current_estado', 'N/A')}")
    st.write(f"**Jefatura:** {st.session_state.get('current_jefatura', 'N/A')}")
    st.write(f"**Jefe Resp.:** {st.session_state.get('current_jefe_residencia', 'N/A')}")

st.sidebar.markdown("---")
if st.sidebar.button("Cerrar Sesión"):
    registrar_auditoria("CIERRE DE SESION", f"Cierre de sesión de {current_email_key}")
    st.session_state["logged_in"] = False
    st.session_state["current_email"] = ""
    st.rerun()

if perfil == "Mi Perfil / Foto":
    st.subheader("Configuración de Perfil y Adscripción Institucional")
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
        btn_subir_foto = st.form_submit_button("Guardar Fotografía")
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
                st.warning("Selecciona un archivo de imagen válido antes de guardar.")

    st.markdown("---")
    st.subheader("Resumen de Actividad y Uso de la Red (Accesos y Bitácoras)")
    st.markdown("Auditoría general del sistema: Conoce qué usuarios han ingresado, cuántas veces han accedido, cuántas bitácoras oficiales han generado y cuántos registros acumulan.")
    
    logs_audit = []
    if os.path.exists(AUDIT_FILE):
        try:
            with open(AUDIT_FILE, "r", encoding="utf-8") as f:
                logs_audit = json.load(f)
        except:
            pass
            
    registros_todos = cargar_registros_acumulados()
    if current_email_key != "victor.olmedo@pa.gob.mx":
        registros_todos = [r for r in registros_todos if r.get("JEFATURA") == jefatura_sesion_actual]
    
    if usuarios_dict_perfil:
        resumen_uso_dict = {}
        for email_u, info_u in usuarios_dict_perfil.items():
            if current_email_key != "victor.olmedo@pa.gob.mx" and info_u.get("jefatura") != jefatura_sesion_actual:
                continue
            resumen_uso_dict[email_u] = {
                "Usuario / Correo": email_u,
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

elif perfil == "Programación de Combustible y Viáticos":
    st.subheader("Programación Mensual de Combustible y Viáticos por Residencia")
    st.markdown("Módulo institucional para la programación, registro y supervisión de actividades, combustible y viáticos por Residencia (Registro Manual).")

    def cargar_programacion():
        if os.path.exists(PROGRAMACION_FILE):
            try:
                with open(PROGRAMACION_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except:
                pass
        return []

    def guardar_programacion(lista_prog):
        with open(PROGRAMACION_FILE, "w", encoding="utf-8") as f:
            json.dump(lista_prog, f, ensure_ascii=False, indent=4)

    residencia_usuario_actual = st.session_state.get("current_jefatura", "RESIDENCIA TENANCINGO")
    estado_usuario_actual = st.session_state.get("current_estado", "Estado de México")
    usuario_actual_nombre = st.session_state.get("current_user", "")

    tab_prog_reg, tab_prog_hist = st.tabs(["Registrar Programación (Manual)", "Historial y Supervisión"])

    with tab_prog_reg:
        st.markdown("### Registro Manual Estilo Tabla Institucional")
        
        if current_email_key == "victor.olmedo@pa.gob.mx":
            cols_res = st.columns(6)
            lista_res_botones = ["Toluca", "Atlacomulco", "Naucalpan", "Tenancingo", "Valle de Bravo", "Texcoco"]
            if "residencia_seleccionada_prog" not in st.session_state:
                st.session_state["residencia_seleccionada_prog"] = "Tenancingo"
                
            for idx_r, res_nombre in enumerate(lista_res_botones):
                with cols_res[idx_r]:
                    if st.button(res_nombre, key=f"btn_res_{res_nombre}", use_container_width=True):
                        st.session_state["residencia_seleccionada_prog"] = res_nombre
            st.info(f"Residencia seleccionada para este registro: **{st.session_state['residencia_seleccionada_prog']}**")
        else:
            res_short = residencia_usuario_actual.replace("RESIDENCIA", "").strip().title()
            st.session_state["residencia_seleccionada_prog"] = res_short
            st.info(f"Residencia asignada para este registro: **{res_short}** (Bloqueada a tu adscripción)")

        if "editor_data" not in st.session_state:
            st.session_state["editor_data"] = pd.DataFrame([{
                "FECHA APROXIMADA DE ACTIVIDAD": date.today().strftime("%d/%m/%Y"),
                "NOMBRE DEL RESPONSABLE DE LA ACTIVIDAD": usuario_actual_nombre,
                "CARGO DEL RESPONSABLE": "Organizador Agrario",
                "PLACAS DEL VEHÍCULO A UTILIZAR": "",
                "DESCRIPCIÓN DE LA ACTIVIDAD A REALIZAR": "",
                "MUNICIPIO": "",
                "EJIDO": "",
                "INDICADOR DEL POA EN QUE REPERCUTE EL USO DEL COMBUSTIBLE": "",
                "¿SE REQUIEREN VIÁTICOS? (SÍ/NO)": "Sí",
                "IMPORTE DE VIÁTICOS ($)": 0.0,
                "IMPORTE DE COMBUSTIBLE": 0.0
            }])

        edited_df = st.data_editor(
            st.session_state["editor_data"],
            num_rows="dynamic",
            use_container_width=True,
            key="prog_data_editor",
            column_config={
                "FECHA APROXIMADA DE ACTIVIDAD": st.column_config.TextColumn("FECHA APROXIMADA DE ACTIVIDAD", required=True),
                "NOMBRE DEL RESPONSABLE DE LA ACTIVIDAD": st.column_config.TextColumn("NOMBRE DEL RESPONSABLE DE LA ACTIVIDAD", required=True),
                "CARGO DEL RESPONSABLE": st.column_config.TextColumn("CARGO DEL RESPONSABLE", required=True),
                "PLACAS DEL VEHÍCULO A UTILIZAR": st.column_config.TextColumn("PLACAS DEL VEHÍCULO A UTILIZAR", required=True),
                "DESCRIPCIÓN DE LA ACTIVIDAD A REALIZAR": st.column_config.TextColumn("DESCRIPCIÓN DE LA ACTIVIDAD A REALIZAR", required=True),
                "MUNICIPIO": st.column_config.TextColumn("MUNICIPIO", required=True),
                "EJIDO": st.column_config.TextColumn("EJIDO"),
                "INDICADOR DEL POA EN QUE REPERCUTE EL USO DEL COMBUSTIBLE": st.column_config.TextColumn("INDICADOR DEL POA EN QUE REPERCUTE EL USO DEL COMBUSTIBLE"),
                "¿SE REQUIEREN VIÁTICOS? (SÍ/NO)": st.column_config.SelectboxColumn("¿SE REQUIEREN VIÁTICOS? (SÍ/NO)", options=["Sí", "No"], required=True),
                "IMPORTE DE VIÁTICOS ($)": st.column_config.NumberColumn("IMPORTE DE VIÁTICOS ($)", min_value=0.0, step=50.0),
                "IMPORTE DE COMBUSTIBLE": st.column_config.NumberColumn("IMPORTE DE COMBUSTIBLE", min_value=0.0, step=50.0)
            }
        )

        if st.button("Guardar Actividades Capturadas", use_container_width=True):
            lista_prog = cargar_programacion()
            res_asignada = st.session_state.get("residencia_seleccionada_prog", "Tenancingo")
            guardados_count = 0
            errores = 0

            for idx, row in edited_df.iterrows():
                placas_val = str(row.get("PLACAS DEL VEHÍCULO A UTILIZAR", "")).strip()
                mun_val = str(row.get("MUNICIPIO", "")).strip()
                desc_val = str(row.get("DESCRIPCIÓN DE LA ACTIVIDAD A REALIZAR", "")).strip()

                if placas_val and mun_val and desc_val:
                    nuevo_reg = {
                        "FECHA": str(row.get("FECHA APROXIMADA DE ACTIVIDAD", date.today().strftime("%d/%m/%Y"))).strip(),
                        "RESPONSABLE": str(row.get("NOMBRE DEL RESPONSABLE DE LA ACTIVIDAD", usuario_actual_nombre)).strip().upper(),
                        "CARGO": str(row.get("CARGO DEL RESPONSABLE", "Organizador Agrario")).strip(),
                        "PLACAS": placas_val.upper(),
                        "DESCRIPCION": desc_val,
                        "MUNICIPIO": mun_val.upper(),
                        "EJIDO": str(row.get("EJIDO", "GENERAL")).strip().upper() if str(row.get("EJIDO", "")).strip() else "GENERAL",
                        "INDICADOR_POA": str(row.get("INDICADOR DEL POA EN QUE REPERCUTE EL USO DEL COMBUSTIBLE", "")).strip(),
                        "VIATICOS_REQ": str(row.get("¿SE REQUIEREN VIÁTICOS? (SÍ/NO)", "SÍ")).strip().upper(),
                        "IMPORTE_VIATICOS": float(row.get("IMPORTE DE VIÁTICOS ($)", 0.0)),
                        "IMPORTE_COMBUSTIBLE": float(row.get("IMPORTE DE COMBUSTIBLE", 0.0)),
                        "RESIDENCIA": res_asignada,
                        "ESTADO": estado_usuario_actual,
                        "ESTATUS": "REGISTRADO",
                        "CORREO": current_email_key
                    }
                    lista_prog.append(nuevo_reg)
                    guardados_count += 1
                else:
                    if placas_val or mun_val or desc_val:
                        errores += 1

            if guardados_count > 0:
                guardar_programacion(lista_prog)
                registrar_auditoria("REGISTRO PROGRAMACION TABLA", f"Guardados {guardados_count} registros manuales para {res_asignada}")
                st.success(f"¡Se guardaron {guardados_count} actividades programadas con éxito!")
                if errores > 0:
                    st.warning(f"Se omitieron {errores} filas por estar incompletas (requieren Placas, Municipio y Descripción).")
                st.session_state["editor_data"] = pd.DataFrame([{
                    "FECHA APROXIMADA DE ACTIVIDAD": date.today().strftime("%d/%m/%Y"),
                    "NOMBRE DEL RESPONSABLE DE LA ACTIVIDAD": usuario_actual_nombre,
                    "CARGO DEL RESPONSABLE": "Organizador Agrario",
                    "PLACAS DEL VEHÍCULO A UTILIZAR": "",
                    "DESCRIPCIÓN DE LA ACTIVIDAD A REALIZAR": "",
                    "MUNICIPIO": "",
                    "EJIDO": "",
                    "INDICADOR DEL POA EN QUE REPERCUTE EL USO DEL COMBUSTIBLE": "",
                    "¿SE REQUIEREN VIÁTICOS? (SÍ/NO)": "Sí",
                    "IMPORTE DE VIÁTICOS ($)": 0.0,
                    "IMPORTE DE COMBUSTIBLE": 0.0
                }])
                st.rerun()
            else:
                st.warning("No hay filas válidas para guardar. Asegúrate de llenar al menos Placas, Municipio y Descripción.")

    with tab_prog_hist:
        st.markdown("### Historial y Supervisión de Programaciones")
        lista_prog_hist = cargar_programacion()
        if lista_prog_hist:
            df_prog = pd.DataFrame(lista_prog_hist)
            if current_email_key == "victor.olmedo@pa.gob.mx":
                df_prog_filtro = df_prog
            else:
                res_clean = residencia_usuario_actual.replace("RESIDENCIA", "").strip().upper()
                df_prog_filtro = df_prog[df_prog['RESIDENCIA'].str.upper().str.contains(res_clean, na=False)]

            if not df_prog_filtro.empty:
                df_prog_filtro['FECHA_DT'] = pd.to_datetime(df_prog_filtro['FECHA'], format='%d/%m/%Y', errors='coerce')
                df_prog_filtro['MES_AÑO_STR'] = df_prog_filtro['FECHA_DT'].dt.month.map(MESES_MAP_ES) + ' ' + df_prog_filtro['FECHA_DT'].dt.year.astype(str)

                meses_disp = ["Todos los meses"] + sorted(df_prog_filtro['MES_AÑO_STR'].dropna().unique().tolist())
                if current_email_key == "victor.olmedo@pa.gob.mx":
                    residencias_disp = ["Todas las residencias"] + sorted(df_prog_filtro['RESIDENCIA'].dropna().unique().tolist())
                else:
                    residencias_disp = [residencia_usuario_actual.replace("RESIDENCIA", "").strip().title()]
                personas_disp = ["Todas las personas"] + sorted(df_prog_filtro['RESPONSABLE'].dropna().unique().tolist())

                col_f1, col_f2, col_f3 = st.columns(3)
                with col_f1:
                    mes_sel = st.selectbox("Filtrar por Mes", meses_disp)
                with col_f2:
                    if current_email_key == "victor.olmedo@pa.gob.mx":
                        res_sel = st.selectbox("Filtrar por Residencia", residencias_disp)
                    else:
                        res_sel = residencia_usuario_actual.replace("RESIDENCIA", "").strip().title()
                        st.text_input("Filtrar por Residencia", value=res_sel, disabled=True)
                with col_f3:
                    pers_sel = st.selectbox("Filtrar por Persona (Responsable)", personas_disp)

                if mes_sel != "Todos los meses":
                    df_prog_filtro = df_prog_filtro[df_prog_filtro['MES_AÑO_STR'] == mes_sel]
                if current_email_key == "victor.olmedo@pa.gob.mx" and res_sel != "Todas las residencias":
                    df_prog_filtro = df_prog_filtro[df_prog_filtro['RESIDENCIA'] == res_sel]
                if pers_sel != "Todas las personas":
                    df_prog_filtro = df_prog_filtro[df_prog_filtro['RESPONSABLE'] == pers_sel]

                df_prog_filtro['ORIGINAL_IDX'] = df_prog_filtro.index
                df_prog_filtro = df_prog_filtro.reset_index(drop=True)
                df_prog_filtro.index = range(1, len(df_prog_filtro) + 1)

                st.dataframe(df_prog_filtro.drop(columns=['FECHA_DT', 'MES_AÑO_STR', 'ORIGINAL_IDX'], errors='ignore'), use_container_width=True)

                if not df_prog_filtro.empty:
                    tot_viaticos = df_prog_filtro['IMPORTE_VIATICOS'].sum()
                    tot_combustible = df_prog_filtro['IMPORTE_COMBUSTIBLE'].sum()

                    m1, m2, m3 = st.columns(3)
                    m1.metric("Total Importe Viáticos", f"${tot_viaticos:,.2f} MXN")
                    m2.metric("Total Importe Combustible", f"${tot_combustible:,.2f} MXN")
                    m3.metric("Total Actividades", f"{len(df_prog_filtro)}")

                    if current_email_key == "victor.olmedo@pa.gob.mx" or rol_actual in ["Administrador Nacional", "Administrador Estatal", "Jefe de Residencia"]:
                        st.markdown("---")
                        st.subheader("Panel de Eliminación Autorizada")
                        tipo_borrado = st.radio("Selecciona el modo de eliminación", ["Eliminar registro específico", "Borrar todos los registros"])

                        if tipo_borrado == "Eliminar registro específico":
                            with st.expander("Eliminar registro individual"):
                                num_a_borrar = st.number_input("Número exacto del registro a eliminar", min_value=1, max_value=max(1, len(df_prog_filtro)), value=1, step=1)

                                if num_a_borrar in df_prog_filtro.index:
                                    reg_sel = df_prog_filtro.loc[num_a_borrar]
                                    orig_idx = int(reg_sel['ORIGINAL_IDX'])
                                    st.warning(f"**Registro seleccionado (#{num_a_borrar}):** Fecha: {reg_sel.get('FECHA')} | Responsable: {reg_sel.get('RESPONSABLE')} | Municipio: {reg_sel.get('MUNICIPIO')}")
                                    confirmar_del = st.checkbox("Confirmo que deseo eliminar este registro permanentemente")
                                    if confirmar_del and st.button("Ejecutar Eliminación Individual"):
                                        lista_prog_hist = [item for i, item in enumerate(lista_prog_hist) if i != orig_idx]
                                        guardar_programacion(lista_prog_hist)
                                        registrar_auditoria("ELIMINAR PROGRAMACION", f"Eliminación de registro #{num_a_borrar} por {st.session_state.get('current_email')}")
                                        st.success("Registro eliminado correctamente.")
                                        st.rerun()
                        else:
                            with st.expander("Borrar todos los registros"):
                                st.error(f"¡Atención! Estás a punto de eliminar **{len(df_prog_filtro)} registros** correspondientes a los filtros actuales.")
                                conf_todo = st.checkbox("Confirmo bajo mi responsabilidad la eliminación masiva de todos los registros")
                                if conf_todo and st.button("Ejecutar Eliminación Masiva"):
                                    indices_a_eliminar = set(df_prog_filtro['ORIGINAL_IDX'].tolist())
                                    lista_prog_hist = [item for i, item in enumerate(lista_prog_hist) if i not in indices_a_eliminar]
                                    guardar_programacion(lista_prog_hist)
                                    registrar_auditoria("ELIMINACION MASIVA", f"Eliminación masiva de {len(indices_a_eliminar)} registros por {st.session_state.get('current_email')}")
                                    st.success("¡Registros eliminados exitosamente!")
                                    st.rerun()

                    st.markdown("---")
                    if st.button("Descargar Programación en Excel Institucional"):
                        try:
                            output_prog = BytesIO()
                            with pd.ExcelWriter(output_prog, engine='openpyxl') as writer:
                                df_prog_filtro.drop(columns=['FECHA_DT', 'MES_AÑO_STR', 'ORIGINAL_IDX'], errors='ignore').to_excel(writer, index=False, sheet_name='PROGRAMACION')
                            output_prog.seek(0)
                            st.download_button(
                                label="Descargar Archivo Excel de Programación",
                                data=output_prog,
                                file_name=f"PROGRAMACION_COMBUSTIBLE_VIATICOS_{residencia_usuario_actual.replace(' ', '_')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        except Exception as e:
                            st.error(f"Error al generar archivo Excel: {e}")
        else:
            st.info("No hay registros de programación de combustible y viáticos en el sistema.")

elif perfil == "Módulo de Captura (Recorrido)":
    st.subheader("Módulo de Captura por Día - Organizador / Operador")
    estado_usuario_actual = st.session_state.get("current_estado", "Estado de México")
    jefatura_actual = st.session_state.get("current_jefatura", "RESIDENCIA NAUCALPAN")
    jefe_actual = st.session_state.get("current_jefe_residencia", "N/A")
    
    st.markdown(f"Ingresa los datos de tu recorrido. Ubicación filtrada para **{estado_usuario_actual}** | Jefatura: **{jefatura_actual}** | Jefe: **{jefe_actual}**.")
    
    registros_previos_user = [r for r in cargar_registros_acumulados() if r.get("CORREO_ORGANIZADOR") == current_email_key]
    km_sugerido = 0.0
    if registros_previos_user:
        ultimo_reg = registros_previos_user[-1]
        km_sugerido = float(ultimo_reg.get("KM FINAL / Km de Llegada", 0.0))
        
    with st.form("form_captura_dia"):
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            fecha = st.date_input("Fecha de registro del uso del vehículo")
            municipio = st.text_input(f"Municipio ({estado_usuario_actual}) - Libre", value="", max_chars=300, placeholder="Escribe el municipio libremente...")
            poblado = st.text_input("Poblado / Localidad - Libre", value="", max_chars=300, placeholder="Escribe la localidad libremente...")
            folio_ciia = st.text_input("Folio CIIA", value="", max_chars=300)
        with col2:
            h_salida = st.text_input("Hora de Salida (Formato 24h, ej. 09:00)", value="09:00", max_chars=300)
            km_inicial = st.number_input("KM Inicial / Salida", min_value=0.0, value=km_sugerido, step=1.0)
            h_llegada = st.text_input("Hora de Llegada (Formato 24h, ej. 17:00)", value="17:00", max_chars=300)
            km_final = st.number_input("KM Final / Llegada", min_value=0.0, value=km_sugerido, step=1.0)
        with col3:
            if current_email_key == "victor.olmedo@pa.gob.mx":
                lista_adscripciones_unicas = sorted(list(set([
                    jefatura_actual, 
                    "RESIDENCIA NAUCALPAN", 
                    "RESIDENCIA TOLUCA", 
                    "RESIDENCIA ATLACOMULCO", 
                    "RESIDENCIA TEXCOCO", 
                    "RESIDENCIA VALLE DE BRAVO", 
                    "RESIDENCIA TENANCINGO"
                ])))
                residencia = st.selectbox("Área de Adscripción", lista_adscripciones_unicas)
            else:
                residencia = jefatura_actual
                st.text_input("Área de Adscripción (Bloqueada)", value=residencia, disabled=True)

            vehiculo = st.text_input("Tipo de Vehículo - Libre", value="", max_chars=300, placeholder="Ej. Nissan Versa, PickUp, Estacas...")
            placas = st.text_input("Placas", value="", max_chars=300)
            licencia = st.text_input("No. De Licencia", value="", max_chars=300)
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
        
        guardar_dia = st.form_submit_button("Guardar Día")
        if guardar_dia:
            if not placas.strip():
                st.error("Debes ingresar el número de placas del vehículo (no hay placa predeterminada).")
            elif not municipio.strip():
                st.error("Debes escribir el municipio.")
            elif not poblado.strip():
                st.error("Debes escribir la localidad o poblado.")
            elif not vehiculo.strip():
                st.error("Debes escribir el tipo de vehículo.")
            elif h_salida.strip() == h_llegada.strip():
                st.error("Error: La Hora de Salida y la Hora de Llegada no pueden ser iguales en formato 24 horas.")
            elif km_final < km_inicial:
                st.warning("Aviso: El KM Final no puede ser menor al KM Inicial.")
            else:
                fecha_str = fecha.strftime("%d/%m/%Y")
                placas_upper = placas.strip().upper()
                mun_upper = municipio.strip().upper()
                pob_upper = poblado.strip().upper()
                veh_upper = vehiculo.strip().upper()
                
                registros_actuales = cargar_registros_acumulados()
                
                conflicto_duplicado = False
                mensaje_error = ""
                for r in registros_actuales:
                    r_fecha = r.get("FECHA COMPLETA")
                    r_placas = str(r.get("Placas", "")).strip().upper()
                    r_km_ini = float(r.get("KM INICIAL / Km de Salida", 0))
                    r_km_fin = float(r.get("KM FINAL / Km de Llegada", 0))
                    if r_fecha == fecha_str and r_placas == placas_upper:
                        if r_km_ini == km_inicial and r_km_fin == km_final and str(r.get("MUNICIPIO")).upper() == mun_upper:
                            conflicto_duplicado = True
                            mensaje_error = "Registro duplicado: Ya existe un recorrido guardado con exactamente la misma información."
                            break
                        elif km_inicial < r_km_fin and km_final > r_km_ini:
                            conflicto_duplicado = True
                            mensaje_error = f"Conflicto de Kilometraje: El vehículo {placas_upper} ya registra un recorrido el {fecha_str} de {r_km_ini:,.0f} a {r_km_fin:,.0f} km."
                            break

                if conflicto_duplicado:
                    st.error(mensaje_error)
                else:
                    recorrido = km_final - km_inicial
                    nuevo_reg = {
                        "FECHA COMPLETA": fecha_str,
                        "MES": MESES_MAP_UPPER[fecha.month],
                        "MUNICIPIO": mun_upper,
                        "POBLADO": pob_upper,
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
                        "Tipo de Vehículo": veh_upper,
                        "Placas": placas_upper,
                        "No. De Licencia": licencia,
                        "ESTADO_ADSCRIPCION": estado_usuario_actual,
                        "JEFATURA": jefatura_actual,
                        "JEFE_RESIDENCIA": jefe_actual,
                        "CORREO_ORGANIZADOR": current_email_key
                    }
                    registros_actuales.append(nuevo_reg)
                    guardar_registros_acumulados(registros_actuales)
                    registrar_auditoria("CAPTURA RECORRIDO", f"Registro de {recorrido} km en {mun_upper} ({pob_upper})")
                    st.success(f"¡Día {fecha_str} en {mun_upper} ({pob_upper}) guardado correctamente!")

    registros_totales = cargar_registros_acumulados()
    if current_email_key == "victor.olmedo@pa.gob.mx":
        registros_filtrados_vista = registros_totales
    else:
        registros_filtrados_vista = [r for r in registros_totales if r.get("JEFATURA") == jefatura_actual]

    if len(registros_totales) > 0:
        st
