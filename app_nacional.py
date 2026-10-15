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

MESES_CORTOS = {
    1: 'ENE', 2: 'FEB', 3: 'MAR', 4: 'ABR', 5: 'MAY', 6: 'JUN',
    7: 'JUL', 8: 'AGO', 9: 'SEP', 10: 'OCT', 11: 'NOV', 12: 'DIC'
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
    "RESIDENCIA TENANCINGO",
    "REPRESENTACIÓN"
]

ROLES_SISTEMA = [
    "Organizador Agrario (Operador)", 
    "Analista de Información", 
    "Administrador de Vehículos (Estatal)",
    "Jefe de Residencia",
    "Representación",
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
        # --- USUARIOS DE RESIDENCIAS Y REPRESENTACIÓN ---
        "toluca.residencia": {"nombre": "OSCAR MENDOZA RUIZ", "pass": "Toluca2026", "licencia": "TOL123456", "rol": "Organizador Agrario (Operador)", "estado": "Estado de México", "jefatura": "RESIDENCIA TOLUCA", "jefe_residencia": "ING. GABRIEL ESTRADA", "foto": "", "activo": True},
        "atlacomulco.residencia": {"nombre": "BRENDA LÓPEZ SÁNCHEZ", "pass": "Atlacomulco2026", "licencia": "ATL123456", "rol": "Organizador Agrario (Operador)", "estado": "Estado de México", "jefatura": "RESIDENCIA ATLACOMULCO", "jefe_residencia": "ING. GABRIEL ESTRADA", "foto": "", "activo": True},
        "naucalpan.residencia": {"nombre": "ADÁN JIMÉNEZ PÉREZ", "pass": "Naucalpan2026", "licencia": "NAU123456", "rol": "Organizador Agrario (Operador)", "estado": "Estado de México", "jefatura": "RESIDENCIA NAUCALPAN", "jefe_residencia": "ING. GABRIEL ESTRADA", "foto": "", "activo": True},
        "texcoco.residencia": {"nombre": "GUADALUPE RAMÍREZ FLORES", "pass": "Texcoco2026", "licencia": "TEX123456", "rol": "Organizador Agrario (Operador)", "estado": "Estado de México", "jefatura": "RESIDENCIA TEXCOCO", "jefe_residencia": "ING. GABRIEL ESTRADA", "foto": "", "activo": True},
        "vallebravo.residencia": {"nombre": "ROBERTO GÓMEZ CRUZ", "pass": "Valle2026", "licencia": "VAL123456", "rol": "Organizador Agrario (Operador)", "estado": "Estado de México", "jefatura": "RESIDENCIA VALLE DE BRAVO", "jefe_residencia": "ING. GABRIEL ESTRADA", "foto": "", "activo": True},
        "tenancingo.residencia": {"nombre": "MARÍA ELENA SOTO VILLA", "pass": "Tenancingo2026", "licencia": "TEN123456", "rol": "Organizador Agrario (Operador)", "estado": "Estado de México", "jefatura": "RESIDENCIA TENANCINGO", "jefe_residencia": "ING. GABRIEL ESTRADA", "foto": "", "activo": True},
        "representacion.residencia": {"nombre": "REPRESENTACIÓN OFICIAL", "pass": "Representacion2026", "licencia": "REP123456", "rol": "Representación", "estado": "Estado de México", "jefatura": "REPRESENTACIÓN", "jefe_residencia": "ING. GABRIEL ESTRADA", "foto": "", "activo": True}
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
elif rol_actual in ["Administrador Estatal", "Administrador de Vehículos (Estatal)", "Jefe de Residencia", "Representación", "Analista de Información"]:
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
    st.subheader("Programación Mensual de Combustible y Viáticos por Residencia y Representación")
    st.markdown("Selecciona el **Año**, el **Área (Residencia o Representación)** y el **Mes** para consultar, registrar o editar la programación de actividades.")

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

    tab_prog_meses, tab_prog_hist = st.tabs(["Programación por Año, Área y Mes", "Historial y Supervisión"])

    with tab_prog_meses:
        # --- SELECCIÓN DE AÑO (2026 AL 2028) ---
        st.markdown("### 1. Selección de Año Institucional")
        if "anio_activo_prog" not in st.session_state:
            st.session_state["anio_activo_prog"] = datetime.now().year if datetime.now().year in [2026, 2027, 2028] else 2026

        cols_anios = st.columns(3)
        anios_disponibles = [2026, 2027, 2028]
        for idx_a, anio_val in enumerate(anios_disponibles):
            with cols_anios[idx_a]:
                if st.button(f"📅 Año {anio_val}", key=f"btn_anio_{anio_val}", use_container_width=True):
                    st.session_state["anio_activo_prog"] = anio_val
                    st.rerun()

        anio_activo = st.session_state["anio_activo_prog"]
        st.info(f"Año seleccionado: **{anio_activo}**")
        st.markdown("---")

        # --- SELECCIÓN DE RESIDENCIA O REPRESENTACIÓN ---
        st.markdown("### 2. Selección de Residencia o Representación")
        if current_email_key == "victor.olmedo@pa.gob.mx":
            cols_res = st.columns(7)
            lista_res_botones = ["Toluca", "Atlacomulco", "Naucalpan", "Tenancingo", "Valle de Bravo", "Texcoco", "Representación"]
            if "residencia_seleccionada_prog" not in st.session_state:
                st.session_state["residencia_seleccionada_prog"] = "Tenancingo"
                
            for idx_r, res_nombre in enumerate(lista_res_botones):
                with cols_res[idx_r]:
                    if st.button(res_nombre, key=f"btn_res_m_{res_nombre}", use_container_width=True):
                        st.session_state["residencia_seleccionada_prog"] = res_nombre
            res_asignada = st.session_state['residencia_seleccionada_prog']
            st.info(f"Área seleccionada para el registro: **{res_asignada}**")
        else:
            if rol_actual == "Representación" or "REPRESENTACION" in residencia_usuario_actual.upper():
                res_asignada = "Representación"
            else:
                res_asignada = residencia_usuario_actual.replace("RESIDENCIA", "").strip().title()
            st.info(f"Adscripción asignada: **{res_asignada}** (Bloqueada a tu perfil)")

        st.markdown("---")

        # --- SELECCIÓN DE MES INSTITUCIONAL ---
        st.markdown("### 3. Selección de Mes Institucional")
        if "mes_activo_prog" not in st.session_state:
            st.session_state["mes_activo_prog"] = datetime.now().month

        st.markdown(f"#### Haz clic en un mes del año **{anio_activo}** para ver y editar su información ({res_asignada}):")
        cols_meses_1 = st.columns(6)
        cols_meses_2 = st.columns(6)

        for m_num in range(1, 7):
            with cols_meses_1[m_num - 1]:
                m_nombre = MESES_CORTOS[m_num]
                if st.button(m_nombre, key=f"mes_btn_{m_num}", use_container_width=True):
                    st.session_state["mes_activo_prog"] = m_num
                    st.rerun()

        for m_num in range(7, 13):
            with cols_meses_2[m_num - 7]:
                m_nombre = MESES_CORTOS[m_num]
                if st.button(m_nombre, key=f"mes_btn_{m_num}", use_container_width=True):
                    st.session_state["mes_activo_prog"] = m_num
                    st.rerun()

        mes_activo = st.session_state["mes_activo_prog"]
        nombre_mes_largo = MESES_MAP_ES[mes_activo]

        st.markdown("---")
        st.markdown(f"### Actividades programadas para: **{nombre_mes_largo.upper()} {anio_activo}** — *{res_asignada}*")

        lista_todos_prog = cargar_programacion()
        registros_mes_actual = []
        for item in lista_todos_prog:
            r_res = str(item.get("RESIDENCIA", "")).upper()
            r_fecha = str(item.get("FECHA", ""))
            if res_asignada.upper() in r_res:
                try:
                    dt_item = datetime.strptime(r_fecha, "%d/%m/%Y")
                    if dt_item.month == mes_activo and dt_item.year == anio_activo:
                        registros_mes_actual.append({
                            "FECHA APROXIMADA DE ACTIVIDAD": r_fecha,
                            "NOMBRE DEL RESPONSABLE DE LA ACTIVIDAD": item.get("RESPONSABLE", usuario_actual_nombre),
                            "CARGO DEL RESPONSABLE": item.get("CARGO", "Organizador Agrario"),
                            "PLACAS DEL VEHÍCULO A UTILIZAR": item.get("PLACAS", ""),
                            "DESCRIPCIÓN DE LA ACTIVIDAD A REALIZAR": item.get("DESCRIPCION", ""),
                            "MUNICIPIO": item.get("MUNICIPIO", ""),
                            "EJIDO": item.get("EJIDO", "GENERAL"),
                            "INDICADOR DEL POA EN QUE REPERCUTE EL USO DEL COMBUSTIBLE": item.get("INDICADOR_POA", ""),
                            "¿SE REQUIEREN VIÁTICOS? (SÍ/NO)": item.get("VIATICOS_REQ", "Sí"),
                            "IMPORTE DE VIÁTICOS ($)": float(item.get("IMPORTE_VIATICOS", 0.0)),
                            "IMPORTE DE COMBUSTIBLE": float(item.get("IMPORTE_COMBUSTIBLE", 0.0))
                        })
                except:
                    pass

        if not registros_mes_actual:
            registros_mes_actual = [{
                "FECHA APROXIMADA DE ACTIVIDAD": f"01/{str(mes_activo).zfill(2)}/{anio_activo}",
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
            }]

        df_mes_editor = pd.DataFrame(registros_mes_actual)

        edited_mes_df = st.data_editor(
            df_mes_editor,
            num_rows="dynamic",
            use_container_width=True,
            key=f"editor_mes_{anio_activo}_{mes_activo}",
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

        if st.button(f"Guardar Actividades de {nombre_mes_largo.upper()} {anio_activo}", use_container_width=True):
            lista_nueva = []
            for item in lista_todos_prog:
                r_res = str(item.get("RESIDENCIA", "")).upper()
                r_fecha = str(item.get("FECHA", ""))
                es_este_mes = False
                if res_asignada.upper() in r_res:
                    try:
                        dt_item = datetime.strptime(r_fecha, "%d/%m/%Y")
                        if dt_item.month == mes_activo and dt_item.year == anio_activo:
                            es_este_mes = True
                    except:
                        pass
                if not es_este_mes:
                    lista_nueva.append(item)

            guardados_count = 0
            for idx, row in edited_mes_df.iterrows():
                placas_val = str(row.get("PLACAS DEL VEHÍCULO A UTILIZAR", "")).strip()
                mun_val = str(row.get("MUNICIPIO", "")).strip()
                desc_val = str(row.get("DESCRIPCIÓN DE LA ACTIVIDAD A REALIZAR", "")).strip()
                f_val = str(row.get("FECHA APROXIMADA DE ACTIVIDAD", f"01/{str(mes_activo).zfill(2)}/{anio_activo}")).strip()

                if placas_val and mun_val and desc_val:
                    nuevo_reg = {
                        "FECHA": f_val,
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
                    lista_nueva.append(nuevo_reg)
                    guardados_count += 1

            guardar_programacion(lista_nueva)
            registrar_auditoria("GUARDAR PROGRAMACION MES", f"Actualización de {guardados_count} registros para {res_asignada} en {nombre_mes_largo} {anio_activo}")
            st.success(f"¡Se ha guardado con éxito la programación de {nombre_mes_largo} {anio_activo} ({guardados_count} actividades) para {res_asignada}!")
            st.rerun()

    with tab_prog_hist:
        st.markdown("### Historial, Supervisión y Corrección de Duplicados")
        st.info("ℹ️ Aquí puedes consultar todo el historial de actividades programadas y eliminar registros específicos en caso de haber guardado por error o duplicado.")
        
        lista_prog_hist = cargar_programacion()
        if lista_prog_hist:
            df_prog = pd.DataFrame(lista_prog_hist)
            
            # --- APARTADO DE DESCARGA GENERAL POR MES (EXCLUSIVO ADMINISTRADOR NACIONAL) ---
            if st.session_state.get("current_rol") == "Administrador Nacional":
                st.markdown("---")
                st.subheader("📥 Descarga General por Mes (Todas las Residencias y Representación)")
                df_prog['FECHA_DT_TEMP'] = pd.to_datetime(df_prog['FECHA'], format='%d/%m/%Y', errors='coerce')
                df_prog['MES_AÑO_STR_TEMP'] = df_prog['FECHA_DT_TEMP'].dt.month.map(MESES_MAP_ES) + ' ' + df_prog['FECHA_DT_TEMP'].dt.year.astype(str)
                meses_globales_disp = sorted(df_prog['MES_AÑO_STR_TEMP'].dropna().unique().tolist()) if not df_prog.empty else []
                
                col_dg1, col_dg2 = st.columns([2, 1])
                with col_dg1:
                    mes_descarga_gen = st.selectbox("Seleccionar Mes y Año para Descarga Consolidada (Global)", meses_globales_disp if meses_globales_disp else ["Sin datos"], key="sel_mes_gen_nac")
                with col_dg2:
                    st.write("")
                    st.write("")
                    if st.button("Generar Consolidado Mensual", use_container_width=True, key="btn_gen_cons_nac") and mes_descarga_gen != "Sin datos":
                        df_mes_gen = df_prog[df_prog['MES_AÑO_STR_TEMP'] == mes_descarga_gen]
                        if not df_mes_gen.empty:
                            try:
                                output_gen = BytesIO()
                                with pd.ExcelWriter(output_gen, engine='openpyxl') as writer:
                                    df_mes_gen.drop(columns=['FECHA_DT_TEMP', 'MES_AÑO_STR_TEMP'], errors='ignore').to_excel(writer, index=False, sheet_name='CONSOLIDADO_GENERAL')
                                output_gen.seek(0)
                                st.download_button(
                                    label=f"📥 Descargar Archivo Consolidado ({mes_descarga_gen})",
                                    data=output_gen,
                                    file_name=f"PROGRAMACION_CONSOLIDADA_GLOBAL_{mes_descarga_gen.replace(' ', '_')}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    key="btn_dl_gen_excel_nac"
                                )
                            except Exception as e:
                                st.error(f"Error al generar archivo Excel consolidado: {e}")
                        else:
                            st.warning("No hay registros para el periodo seleccionado.")
                st.markdown("---")
            # -------------------------------------------------------------------------------------------------

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
                    residencias_disp = ["Todas las áreas"] + sorted(df_prog_filtro['RESIDENCIA'].dropna().unique().tolist())
                else:
                    residencias_disp = [res_asignada]
                personas_disp = ["Todas las personas"] + sorted(df_prog_filtro['RESPONSABLE'].dropna().unique().tolist())

                col_f1, col_f2, col_f3 = st.columns(3)
                with col_f1:
                    mes_sel = st.selectbox("Filtrar por Mes y Año", meses_disp, key="filtro_mes_prog_h")
                with col_f2:
                    if current_email_key == "victor.olmedo@pa.gob.mx":
                        res_sel = st.selectbox("Filtrar por Residencia / Representación", residencias_disp, key="filtro_res_prog_h")
                    else:
                        res_sel = res_asignada
                        st.text_input("Filtrar por Área", value=res_sel, disabled=True, key="filtro_res_prog_dis_h")
                with col_f3:
                    pers_sel = st.selectbox("Filtrar por Persona (Responsable)", personas_disp, key="filtro_pers_prog_h")

                if mes_sel != "Todos los meses":
                    df_prog_filtro = df_prog_filtro[df_prog_filtro['MES_AÑO_STR'] == mes_sel]
                if current_email_key == "victor.olmedo@pa.gob.mx" and res_sel != "Todas las áreas":
                    df_prog_filtro = df_prog_filtro[df_prog_filtro['RESIDENCIA'] == res_sel]
                if pers_sel != "Todas las personas":
                    df_prog_filtro = df_prog_filtro[df_prog_filtro['RESPONSABLE'] == pers_sel]

                df_prog_filtro['ORIGINAL_IDX'] = df_prog_filtro.index
                df_prog_filtro = df_prog_filtro.reset_index(drop=True)
                df_prog_filtro.index = range(1, len(df_prog_filtro) + 1)

                st.dataframe(df_prog_filtro.drop(columns=['FECHA_DT', 'MES_AÑO_STR', 'ORIGINAL_IDX', 'FECHA_DT_TEMP', 'MES_AÑO_STR_TEMP'], errors='ignore'), use_container_width=True)

                if not df_prog_filtro.empty:
                    tot_viaticos = df_prog_filtro['IMPORTE_VIATICOS'].sum()
                    tot_combustible = df_prog_filtro['IMPORTE_COMBUSTIBLE'].sum()

                    m1, m2, m3 = st.columns(3)
                    m1.metric("Total Importe Viáticos", f"${tot_viaticos:,.2f} MXN")
                    m2.metric("Total Importe Combustible", f"${tot_combustible:,.2f} MXN")
                    m3.metric("Total Actividades", f"{len(df_prog_filtro)}")

                    st.markdown("---")
                    st.subheader("Eliminación de Registros por Error / Duplicados")
                    with st.expander("Eliminar un registro específico del historial"):
                        num_a_borrar = st.number_input("Número exacto del registro en la tabla a eliminar", min_value=1, max_value=max(1, len(df_prog_filtro)), value=1, step=1)
                        if num_a_borrar in df_prog_filtro.index:
                            reg_sel = df_prog_filtro.loc[num_a_borrar]
                            orig_idx = int(reg_sel['ORIGINAL_IDX'])
                            st.warning(f"**Registro seleccionado (#{num_a_borrar}):** Fecha: {reg_sel.get('FECHA')} | Responsable: {reg_sel.get('RESPONSABLE')} | Municipio: {reg_sel.get('MUNICIPIO')}")
                            if st.button("Eliminar este registro específico"):
                                lista_prog_hist = [item for i, item in enumerate(lista_prog_hist) if i != orig_idx]
                                guardar_programacion(lista_prog_hist)
                                registrar_auditoria("ELIMINAR REGISTRO PROGRAMACION", f"Eliminación de registro #{num_a_borrar}")
                                st.success("¡Registro eliminado correctamente!")
                                st.rerun()

                    st.markdown("---")
                    if st.button("Descargar Programación en Excel Institucional"):
                        try:
                            output_prog = BytesIO()
                            with pd.ExcelWriter(output_prog, engine='openpyxl') as writer:
                                df_prog_filtro.drop(columns=['FECHA_DT', 'MES_AÑO_STR', 'ORIGINAL_IDX', 'FECHA_DT_TEMP', 'MES_AÑO_STR_TEMP'], errors='ignore').to_excel(writer, index=False, sheet_name='PROGRAMACION')
                            output_prog.seek(0)
                            st.download_button(
                                label="Descargar Archivo Excel de Programación",
                                data=output_prog,
                                file_name=f"PROGRAMACION_COMBUSTIBLE_VIATICOS_{res_asignada.replace(' ', '_')}.xlsx",
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
                    "RESIDENCIA TENANCINGO",
                    "REPRESENTACIÓN"
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
        st.markdown("---")
        st.subheader("Días Guardados (Historial Acumulado Permanente)")
        busqueda_texto = st.text_input("Buscar en historial (Folio CIIA, Oficio, Poblado o Municipio):").strip().lower()
        df_acumulado = pd.DataFrame(registros_filtrados_vista)
        if not df_acumulado.empty and busqueda_texto:
            mask = df_acumulado.apply(lambda row: row.astype(str).str.lower().str.contains(busqueda_texto).any(), axis=1)
            df_acumulado = df_acumulado[mask]
        
        if not df_acumulado.empty:
            st.dataframe(df_acumulado, use_container_width=True)
        else:
            st.info("No hay registros guardados para tu residencia.")
        
        col_acc1, col_acc2 = st.columns(2)
        with col_acc1:
            if current_email_key == "victor.olmedo@pa.gob.mx" and st.button("Limpiar historial completo"):
                guardar_registros_acumulados([])
                registrar_auditoria("LIMPIAR HISTORIAL", "Se eliminó todo el historial acumulado de recorridos")
                st.rerun()
        
        with col_acc2:
            if st.button("Generar y Descargar 3 Bitácoras Oficiales Definitivas"):
                if not os.path.exists(PLANTILLA_EXCEL):
                    st.error(f"Error crítico: No se encuentra la plantilla oficial. Ruta buscada: {PLANTILLA_EXCEL}")
                else:
                    try:
                        wb = openpyxl.load_workbook(PLANTILLA_EXCEL)
                        
                        ws_b = wb["BASE_DE_DATOS"]
                        max_r_existente = max(100, ws_b.max_row)
                        for r in range(2, max_r_existente + 1):
                            for col in ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V']:
                                ws_b[f'{col}{r}'] = None

                        for i, reg in enumerate(registros_filtrados_vista, start=2):
                            fecha_dt = datetime.strptime(reg["FECHA COMPLETA"], "%d/%m/%Y").date()
                            ws_b[f'A{i}'] = fecha_dt
                            ws_b[f'A{i}'].number_format = "DD/MM/YYYY"
                            ws_b[f'B{i}'] = MESES_MAP_UPPER[fecha_dt.month]
                            ws_b[f'C{i}'] = reg["MUNICIPIO"]
                            ws_b[f'D{i}'] = reg["POBLADO"]
                            ws_b[f'E{i}'] = reg["folio CIIA"]
                            ws_b[f'F{i}'] = reg["HORA DE SALIDA"]
                            
                            if i == 2:
                                ws_b[f'G{i}'] = reg["KM INICIAL / Km de Salida"]
                            else:
                                ws_b[f'G{i}'] = f'=J{i-1}'
                                
                            ws_b[f'H{i}'] = reg["RECORRIDO"]
                            ws_b[f'I{i}'] = reg["HORA DE LLEGADA"]
                            ws_b[f'J{i}'] = f'=G{i}+H{i}'
                            ws_b[f'K{i}'] = 12.0
                            ws_b[f'L{i}'] = reg["Gasolina de Salida"]
                            ws_b[f'M{i}'] = f'=ROUND((H{i}/K{i})*23.99, 2)'
                            ws_b[f'N{i}'] = reg["GASTO COMBUSTI"]
                            ws_b[f'O{i}'] = reg["Oficio Numero"]
                            ws_b[f'P{i}'] = reg["Oficio Fecha"]
                            ws_b[f'Q{i}'] = reg["observaciones"]
                            ws_b[f'R{i}'] = reg["Usuario Responsable"]
                            ws_b[f'S{i}'] = reg["Áreas de Adscripción"]
                            ws_b[f'T{i}'] = reg["Tipo de Vehículo"]
                            ws_b[f'U{i}'] = reg["Placas"]
                            ws_b[f'V{i}'] = reg["No. De Licencia"]

                        row_tot = len(registros_filtrados_vista) + 2
                        ws_b[f'A{row_tot}'] = "TOTALES"
                        ws_b[f'H{row_tot}'] = f'=SUM(H2:H{row_tot-1})'

                        if "2_BITACORA_VERSA" in wb.sheetnames:
                            ws_v = wb["2_BITACORA_VERSA"]
                            for r in range(3, 100):
                                for col in ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W']:
                                    ws_v[f'{col}{r}'] = None

                            for idx, r_db in enumerate(range(2, 2 + len(registros_filtrados_vista))):
                                r_v = idx + 3
                                ws_v[f'A{r_v}'] = f'=IF(BASE_DE_DATOS!A{r_db}="","",BASE_DE_DATOS!A{r_db})'
                                ws_v[f'B{r_v}'] = f'=IF(BASE_DE_DATOS!R{r_db}="","",BASE_DE_DATOS!R{r_db})'
                                ws_v[f'C{r_v}'] = f'=IF(BASE_DE_DATOS!S{r_db}="","",BASE_DE_DATOS!S{r_db})'
                                ws_v[f'D{r_v}'] = f'=IF(BASE_DE_DATOS!T{r_db}="","",BASE_DE_DATOS!T{r_db})'
                                ws_v[f'E{r_v}'] = f'=IF(BASE_DE_DATOS!U{r_db}="","",BASE_DE_DATOS!U{r_db})'
                                ws_v[f'F{r_v}'] = f'=IF(BASE_DE_DATOS!O{r_db}="","",BASE_DE_DATOS!O{r_db})'
                                ws_v[f'G{r_v}'] = f'=IF(BASE_DE_DATOS!P{r_db}="","",BASE_DE_DATOS!P{r_db})'
                                ws_v[f'H{r_v}'] = f'=IF(BASE_DE_DATOS!F{r_db}="","",BASE_DE_DATOS!F{r_db})'
                                ws_v[f'I{r_v}'] = f'=IF(BASE_DE_DATOS!G{r_db}="","",BASE_DE_DATOS!G{r_db})'
                                ws_v[f'J{r_v}'] = f'=IF(BASE_DE_DATOS!L{r_db}="v","X","")'
                                ws_v[f'K{r_v}'] = f'=IF(BASE_DE_DATOS!L{r_db}="1/4","X","")'
                                ws_v[f'L{r_v}'] = f'=IF(BASE_DE_DATOS!L{r_db}="1/2","X","")'
                                ws_v[f'M{r_v}'] = f'=IF(BASE_DE_DATOS!L{r_db}="1","X","")'
                                ws_v[f'N{r_v}'] = f'=IF(BASE_DE_DATOS!I{r_db}="","",BASE_DE_DATOS!I{r_db})'
                                ws_v[f'O{r_v}'] = f'=IF(BASE_DE_DATOS!J{r_db}="","",BASE_DE_DATOS!J{r_db})'
                                ws_v[f'P{r_v}'] = f'=IF(BASE_DE_DATOS!M{r_db}="v","X","")'
                                ws_v[f'Q{r_v}'] = f'=IF(BASE_DE_DATOS!M{r_db}="1/4","X","")'
                                ws_v[f'R{r_v}'] = f'=IF(BASE_DE_DATOS!M{r_db}="1/2","X","")'
                                ws_v[f'S{r_v}'] = f'=IF(BASE_DE_DATOS!M{r_db}="1","X","")'
                                ws_v[f'T{r_v}'] = f'=IF(BASE_DE_DATOS!N{r_db}="","",BASE_DE_DATOS!N{r_db})'
                                ws_v[f'V{r_v}'] = f'=IF(BASE_DE_DATOS!V{r_db}="","",BASE_DE_DATOS!V{r_db})'
                                ws_v[f'W{r_v}'] = f'=IF(BASE_DE_DATOS!Q{r_db}="","",BASE_DE_DATOS!Q{r_db})'

                        if "3_BITACORAS_INDIVIDUALES" in wb.sheetnames:
                            ws_i = wb["3_BITACORAS_INDIVIDUALES"]
                            for r in range(1, 500):
                                for c in range(1, 10):
                                    ws_i.cell(row=r, column=c, value=None)

                            for idx, r_db in enumerate(range(2, 2 + len(registros_filtrados_vista))):
                                start_row = 1 + idx * 12
                                fecha_reg_dt = datetime.strptime(registros_filtrados_vista[idx]["FECHA COMPLETA"], "%d/%m/%Y").date()
                                mes_nombre_upper = MESES_MAP_UPPER[fecha_reg_dt.month]
                                ws_i.cell(row=start_row, column=1, value=f'="BITÁCORA " & "{mes_nombre_upper}"')
                                ws_i.cell(row=start_row+1, column=4, value=f'=IF(BASE_DE_DATOS!A{r_db}="","",BASE_DE_DATOS!A{r_db})')
                                ws_i.cell(row=start_row+2, column=2, value=f'=IF(BASE_DE_DATOS!T{r_db}="","",BASE_DE_DATOS!T{r_db})')
                                ws_i.cell(row=start_row+2, column=4, value=f'=IF(BASE_DE_DATOS!U{r_db}="","",BASE_DE_DATOS!U{r_db})')
                                ws_i.cell(row=start_row+3, column=2, value=f'=IF(BASE_DE_DATOS!R{r_db}="","",BASE_DE_DATOS!R{r_db})')
                                ws_i.cell(row=start_row+3, column=4, value=f'=IF(BASE_DE_DATOS!V{r_db}="","",BASE_DE_DATOS!V{r_db})')
                                ws_i.cell(row=start_row+4, column=2, value=f'=IF(BASE_DE_DATOS!F{r_db}="","",BASE_DE_DATOS!F{r_db})')
                                ws_i.cell(row=start_row+4, column=4, value=f'=IF(BASE_DE_DATOS!O{r_db}="","",BASE_DE_DATOS!O{r_db})')
                                ws_i.cell(row=start_row+5, column=4, value=f'=IF(BASE_DE_DATOS!G{r_db}="","",BASE_DE_DATOS!G{r_db})')
                                ws_i.cell(row=start_row+6, column=2, value=f'=IF(BASE_DE_DATOS!I{r_db}="","",BASE_DE_DATOS!I{r_db})')
                                ws_i.cell(row=start_row+6, column=4, value=f'=IF(BASE_DE_DATOS!J{r_db}="","",BASE_DE_DATOS!J{r_db})')
                                ws_i.cell(row=start_row+7, column=3, value=f'=IF(BASE_DE_DATOS!N{r_db}="","",BASE_DE_DATOS!N{r_db})')
                                ws_i.cell(row=start_row+8, column=2, value=f'=IF(BASE_DE_DATOS!C{r_db}="","",BASE_DE_DATOS!C{r_db})')
                                ws_i.cell(row=start_row+9, column=2, value=f'=IF(BASE_DE_DATOS!Q{r_db}="","",BASE_DE_DATOS!Q{r_db})')

                        wb.calculation.fullCalcOnLoad = True
                        wb.calculation.forceFullCalc = True
                        wb.calculation.calcMode = "auto"

                        output = BytesIO()
                        wb.save(output)
                        output.seek(0)
                        
                        registrar_auditoria("GENERAR BITACORAS", "Generación nativa completa exitosa")
                        st.success("¡Las 3 bitácoras oficiales se generaron y vincularon perfectamente sin residuos!")
                        st.download_button(
                            label="Descargar Archivo Oficial Definitivo",
                            data=output,
                            file_name="BITACORAS_OFICIALES_DEFINITIVAS.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    except Exception as e:
                        st.error(f"Error crítico al procesar la plantilla Excel: {e}")

elif perfil in ["Panel de Administración y Auditoría", "Panel de Supervisión (Estatal/Residencia)"]:
    st.subheader("Panel de Gestión, Supervisión y Auditoría")
    
    if current_email_key == "victor.olmedo@pa.gob.mx":
        st.markdown("""
            <div style="background-color: #fcf5f6; border: 2px solid #6B1D2F; padding: 15px; border-radius: 8px; margin-bottom: 20px;">
                <h4 style="color: #6B1D2F; margin-top: 0;">Panel de Respaldo Maestro Exclusivo (Víctor Olmedo)</h4>
                <p style="font-size: 13px; color: #4a2c33;">Este botón descarga un archivo ZIP con absolutamente todo: el código fuente actual, todas las bases de datos de registros y auditoría, y las fotografías de perfil de todos los usuarios de la red.</p>
            </div>
        """, unsafe_allow_html=True)
        
        if st.button("Generar y Descargar ZIP Maestro (Sistema + Bases de Datos + Fotos)"):
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                script_actual = os.path.join(BASE_DIR, "app_nacional.py")
                if os.path.exists(script_actual):
                    zip_file.write(script_actual, "app_nacional.py")
                
                for json_f in [USUARIOS_FILE, REGISTROS_FILE, PROGRAMACION_FILE, AUDIT_FILE]:
                    if os.path.exists(json_f):
                        zip_file.write(json_f, os.path.basename(json_f))
                
                if os.path.exists(FOTOS_DIR):
                    for root, dirs, files in os.walk(FOTOS_DIR):
                        for file in files:
                            file_path = os.path.join(root, file)
                            arcname = os.path.join("fotos_perfil", file)
                            zip_file.write(file_path, arcname)
                            
                for ext_f in [PLANTILLA_EXCEL, MUN_FILE, LOC_FILE]:
                    if ext_f and os.path.exists(ext_f):
                        zip_file.write(ext_f, os.path.basename(ext_f))
                        
            zip_buffer.seek(0)
            registrar_auditoria("RESPALDO MAESTRO", "Descarga de ZIP maestro de sistema y bases de datos")
            st.success("¡Paquete de respaldo maestro generado exitosamente!")
            st.download_button(
                label="Descargar Archivo ZIP Maestro Institucional",
                data=zip_buffer,
                file_name=f"RESPALDO_MAESTRO_PA_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                mime="application/zip"
            )
        st.markdown("---")

    if current_email_key == "victor.olmedo@pa.gob.mx":
        tab_reg_user, tab_edit_user, tab_ctrl_user, tab_resumen_auditoria, tab_bitacora_audit = st.tabs([
            "Alta de Usuario", "Editar Usuario", "Control y Estatus", 
            "Resumen Ejecutivo y Auditoría", "Bitácora de Auditoría"
        ])
    else:
        tab_resumen_auditoria = st.container()
        st.markdown("---")
        st.subheader(f"Supervisión de Registros de tu Residencia ({jefatura_sesion_actual})")
        
    def render_alta_usuario():
        with st.form("form_nuevo_usuario"):
            col_u1, col_u2 = st.columns(2)
            with col_u1:
                c_email = st.text_input("Correo / Usuario", max_chars=300)
                c_nombre = st.text_input("Nombre Completo (Mayúsculas)", max_chars=300)
                c_pass = st.text_input("Contraseña Asignada", type="password", max_chars=300)
                c_licencia = st.text_input("Número de Licencia de Conducir", max_chars=300)
            with col_u2:
                c_estado = st.selectbox("Estado de Adscripción", ESTADOS_REPUBLICA)
                c_jefatura = st.selectbox("Jefatura de Residencia", JEFATURAS_RESIDENCIA)
                c_jefe_res = st.text_input("Nombre del Jefe de Residencia Asignado", value="ING. GABRIEL ESTRADA", max_chars=300)
                c_rol = st.selectbox("Rol en el Sistema", ROLES_SISTEMA)
            btn_crear = st.form_submit_button("Registrar Usuario en la Red")
            if btn_crear:
                if c_email and c_nombre and c_pass:
                    usuarios_actuales = cargar_usuarios()
                    email_limpio = c_email.strip().lower()
                    usuarios_actuales[email_limpio] = {
                        "nombre": c_nombre.strip().upper(), "pass": c_pass.strip(), "licencia": c_licencia.strip(),
                        "estado": c_estado, "jefatura": c_jefatura, "jefe_residencia": c_jefe_res.strip().upper(),
                        "rol": c_rol, "foto": "", "activo": True
                    }
                    guardar_usuarios(usuarios_actuales)
                    registrar_auditoria("ALTA USUARIO", f"Creación de usuario {email_limpio} con rol {c_rol}")
                    st.success(f"¡Usuario {c_nombre} ({c_rol}) registrado exitosamente para {c_estado}!")
                    st.rerun()
                else:
                    st.error("Por favor completa los campos obligatorios (Usuario/Correo, Nombre y Contraseña).")

    def render_editar_usuario():
        usuarios_actuales_edit = cargar_usuarios()
        lista_emails = list(usuarios_actuales_edit.keys())
        if lista_emails:
            email_a_editar = st.selectbox("Selecciona el usuario a modificar", lista_emails)
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
                        jef_actual = u_data.get("jefatura", "RESIDENCIA NAUCALPAN")
                        idx_jef = JEFATURAS_RESIDENCIA.index(jef_actual) if jef_actual in JEFATURAS_RESIDENCIA else 0
                        e_jefatura = st.selectbox("Jefatura de Residencia", JEFATURAS_RESIDENCIA, index=idx_jef)
                        e_jefe_res = st.text_input("Jefe de Residencia Asignado", value=u_data.get("jefe_residencia", ""), max_chars=300)
                        idx_rol = ROLES_SISTEMA.index(u_data.get("rol", "Organizador Agrario (Operador)")) if u_data.get("rol") in ROLES_SISTEMA else 0
                        e_rol = st.selectbox("Rol en el Sistema", ROLES_SISTEMA, index=idx_rol)
                    btn_actualizar = st.form_submit_button("Guardar Cambios")
                    if btn_actualizar:
                        usuarios_actuales_edit[email_a_editar]["nombre"] = e_nombre.strip().upper()
                        usuarios_actuales_edit[email_a_editar]["pass"] = e_pass.strip()
                        usuarios_actuales_edit[email_a_editar]["licencia"] = e_licencia.strip()
                        usuarios_actuales_edit[email_a_editar]["estado"] = e_estado
                        usuarios_actuales_edit[email_a_editar]["jefatura"] = e_jefatura
                        usuarios_actuales_edit[email_a_editar]["jefe_residencia"] = e_jefe_res.strip().upper()
                        usuarios_actuales_edit[email_a_editar]["rol"] = e_rol
                        guardar_usuarios(usuarios_actuales_edit)
                        registrar_auditoria("EDITAR USUARIO", f"Actualización de datos para {email_a_editar}")
                        st.success(f"¡Información de {email_a_editar} actualizada exitosamente!")
                        st.rerun()

    def render_control_estatus():
        st.subheader("Control, Estatus y Eliminación de Usuarios en la Red Nacional")
        usuarios_actuales_tabla = cargar_usuarios()
        filas_html = ""
        for email, datos in usuarios_actuales_tabla.items():
            estado_activo = datos.get("activo", True)
            color_fondo = "#d4edda" if estado_activo else "#e2e3e5"
            texto_estado = "Activo" if estado_activo else "Desac."
            filas_html += f"""
            <div style="background-color: {color_fondo}; padding: 8px 12px; border-radius: 6px; margin-bottom: 6px; border: 1px solid #c0c0c0; display: flex; justify-content: space-between; align-items: center; font-size: 12px;">
                <span style="flex: 2.0; text-align: left;"><b>{email}</b></span>
                <span style="flex: 2.5; text-align: left;">{datos.get('nombre')}</span>
                <span style="flex: 2.0; text-align: left;">{datos.get('rol')}</span>
                <span style="flex: 1.3; text-align: left;">{datos.get('estado')}</span>
                <span style="flex: 1.8; text-align: left;">{datos.get('jefatura')}</span>
                <span style="flex: 0.9; text-align: right;"><b>{texto_estado}</b></span>
            </div>
            """
        st.markdown(f"""
            <div style="background-color: #6B1D2F; color: white; padding: 10px 15px; border-radius: 6px; font-weight: bold; display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px; font-size: 13px;">
                <span style="flex: 2.0; text-align: left;">USUARIO / CORREO</span>
                <span style="flex: 2.5; text-align: left;">NOMBRE</span>
                <span style="flex: 2.0; text-align: left;">ROL</span>
                <span style="flex: 1.3; text-align: left;">ESTADO</span>
                <span style="flex: 1.8; text-align: left;">JEFATURA</span>
                <span style="flex: 0.9; text-align: right;">ESTATUS</span>
            </div>
            {filas_html}
        """, unsafe_allow_html=True)
        
        st.markdown("---")
        st.subheader("Botones de Acción por Usuario")
        for email, datos in usuarios_actuales_tabla.items():
            col_acc_lbl, col_b1, col_b2, col_b3 = st.columns([3, 1, 1, 1])
            with col_acc_lbl:
                st.write(f"**{email}** ({datos.get('nombre')})")
            with col_b1:
                if st.button("Encender", key=f"activar_{email}"):
                    usuarios_actuales_tabla[email]["activo"] = True
                    guardar_usuarios(usuarios_actuales_tabla)
                    registrar_auditoria("ACTIVAR USUARIO", f"Activación de cuenta para {email}")
                    st.rerun()
            with col_b2:
                if st.button("Apagar", key=f"desactivar_{email}"):
                    usuarios_actuales_tabla[email]["activo"] = False
                    guardar_usuarios(usuarios_actuales_tabla)
                    registrar_auditoria("DESACTIVAR USUARIO", f"Desactivación de cuenta para {email}")
                    st.rerun()
            with col_b3:
                if st.button("Eliminar", key=f"eliminar_{email}"):
                    if email == st.session_state["current_email"]:
                        st.error("No puedes eliminar tu propia cuenta.")
                    else:
                        del usuarios_actuales_tabla[email]
                        guardar_usuarios(usuarios_actuales_tabla)
                        registrar_auditoria("ELIMINAR USUARIO", f"Eliminación permanente de cuenta para {email}")
                        st.rerun()

    def render_resumen_auditoria():
        st.subheader("Resumen Ejecutivo y Auditoría (Control Vehicular)")
        registros_persisted = cargar_registros_acumulados()
        if len(registros_persisted) > 0:
            df_global = pd.DataFrame(registros_persisted)
            for col in ['JEFATURA', 'ESTADO_ADSCRIPCION', 'JEFE_RESIDENCIA', 'Áreas de Adscripción', 'Usuario Responsable']:
                if col not in df_global.columns:
                    df_global[col] = 'N/A'
            df_global['FECHA_DT'] = pd.to_datetime(df_global['FECHA COMPLETA'], format='%d/%m/%Y', errors='coerce').dt.date
            
            if current_email_key != "victor.olmedo@pa.gob.mx":
                df_global = df_global[df_global['JEFATURA'] == jefatura_sesion_actual]
            
            col_f1, col_f2, col_f3, col_f4, col_f5 = st.columns(5)
            with col_f1:
                f_inicio = st.date_input("Fecha Inicial", value=df_global['FECHA_DT'].min() if not df_global.empty else date.today())
            with col_f2:
                f_fin = st.date_input("Fecha Final", value=df_global['FECHA_DT'].max() if not df_global.empty else date.today())
            with col_f3:
                orgs_disp = ["Todos los organizadores"] + sorted(df_global['Usuario Responsable'].dropna().unique().tolist()) if not df_global.empty else ["Todos"]
                filtro_usuario = st.selectbox("Organizador Agrario", orgs_disp)
            with col_f4:
                if current_email_key == "victor.olmedo@pa.gob.mx":
                    est_disp = ["Todos los estados"] + sorted(df_global['ESTADO_ADSCRIPCION'].dropna().unique().tolist()) if not df_global.empty else ["Todos"]
                    filtro_estado = st.selectbox("Estado", est_disp)
                else:
                    filtro_estado = estado_usuario_actual
            with col_f5:
                if current_email_key == "victor.olmedo@pa.gob.mx":
                    jef_disp = ["Todas las jefaturas"] + sorted(df_global['JEFATURA'].dropna().unique().tolist()) if not df_global.empty else ["Todas"]
                    filtro_jefatura = st.selectbox("Jefatura de Residencia", jef_disp)
                else:
                    filtro_jefatura = jefatura_sesion_actual
                
            if not df_global.empty:
                df_filtrado = df_global[(df_global['FECHA_DT'] >= f_inicio) & (df_global['FECHA_DT'] <= f_fin)]
                if filtro_usuario != "Todos los organizadores":
                    df_filtrado = df_filtrado[df_filtrado['Usuario Responsable'] == filtro_usuario]
                if current_email_key == "victor.olmedo@pa.gob.mx" and filtro_estado != "Todos los estados":
                    df_filtrado = df_filtrado[df_filtrado['ESTADO_ADSCRIPCION'] == filtro_estado]
                if current_email_key == "victor.olmedo@pa.gob.mx" and filtro_jefatura != "Todas las jefaturas":
                    df_filtrado = df_filtrado[df_filtrado['JEFATURA'] == filtro_jefatura]
                
                if not df_filtrado.empty:
                    st.markdown("---")
                    total_km = df_filtrado['RECORRIDO'].sum()
                    total_gas = df_filtrado['GASTO COMBUSTI'].sum()
                    total_viajes = len(df_filtrado)
                    
                    m1, m2, m3 = st.columns(3)
                    m1.metric("KM Totales Recorridos", f"{total_km:,.1f} km")
                    m2.metric("Gasto Total de Combustible", f"${total_gas:,.2f} MXN")
                    m3.metric("Comisiones Registradas", f"{total_viajes}")
                    
                    st.markdown("---")
                    st.subheader("Alertas de Mantenimiento Preventivo (Flotilla)")
                    df_mantenimiento = df_filtrado.groupby("Placas")["RECORRIDO"].sum().reset_index()
                    for _, row in df_mantenimiento.iterrows():
                        km_total_veh = row["RECORRIDO"]
                        placa_v = row["Placas"]
                        if km_total_veh >= 5000:
                            st.warning(f"ALERTA DE MANTENIMIENTO: El vehículo **{placa_v}** suma **{km_total_veh:,.1f} km**. Requiere servicio.")
                        else:
                            st.info(f"El vehículo **{placa_v}** registra **{km_total_veh:,.1f} km** (En rango operativo normal).")
                    
                    st.markdown("---")
                    col_g1, col_g2 = st.columns(2)
                    with col_g1:
                        st.write("**Kilómetros por Municipio**")
                        st.bar_chart(df_filtrado.groupby("MUNICIPIO")["RECORRIDO"].sum())
                    with col_g2:
                        st.write("**Gasto de Combustible por Jefatura**")
                        st.bar_chart(df_filtrado.groupby("JEFATURA")["GASTO COMBUSTI"].sum())
                    
                    st.markdown("---")
                    st.subheader("Mapa Interactivo de Movimiento y Rutas")
                    mapa_rows = []
                    for _, r in df_filtrado.iterrows():
                        mun = r["MUNICIPIO"]
                        mun_key = next((k for k in CODS_MUNICIPIOS if k.upper() in mun.upper()), None)
                        if mun_key:
                            mapa_rows.append({
                                "lat": CODS_MUNICIPIOS[mun_key]["lat"], "lon": CODS_MUNICIPIOS[mun_key]["lon"],
                                "MUNICIPIO": mun, "POBLADO": r["POBLADO"], "FECHA": r["FECHA COMPLETA"],
                                "USUARIO": r["Usuario Responsable"], "RECORRIDO": r["RECORRIDO"]
                            })
                    if mapa_rows:
                        df_mapa = pd.DataFrame(mapa_rows)
                        arcos_rows = []
                        if len(df_mapa) > 1:
                            for i in range(len(df_mapa) - 1):
                                arcos_rows.append({
                                    "start_lon": df_mapa.iloc[i]["lon"], "start_lat": df_mapa.iloc[i]["lat"],
                                    "end_lon": df_mapa.iloc[i+1]["lon"], "end_lat": df_mapa.iloc[i+1]["lat"]
                                })
                        df_arcos = pd.DataFrame(arcos_rows) if arcos_rows else pd.DataFrame()
                        
                        capa_puntos = pdk.Layer("ScatterplotLayer", data=df_mapa, get_position='[lon, lat]', get_color='[107, 29, 47, 180]', get_radius=3000, pickable=True, auto_highlight=True)
                        layers = [capa_puntos]
                        if not df_arcos.empty:
                            capa_lineas = pdk.Layer("ArcLayer", data=df_arcos, get_source_position='[start_lon, start_lat]', get_target_position='[end_lon, end_lat]', get_source_color='[107, 29, 47]', get_target_color='[138, 37, 61]', get_width=3)
                            layers.append(capa_lineas)
                            
                        r_deck = pdk.Deck(
                            layers=layers,
                            initial_view_state=pdk.ViewState(latitude=df_mapa["lat"].mean(), longitude=df_mapa["lon"].mean(), zoom=8, pitch=30),
                            tooltip={"html": "<b>Municipio:</b> {MUNICIPIO}<br/><b>Poblado:</b> {POBLADO}<br/><b>Fecha:</b> {FECHA}<br/><b>Responsable:</b> {USUARIO}", "style": {"backgroundColor": "#6B1D2F", "color": "white"}}
                        )
                        st.pydeck_chart(r_deck)
                    
                    st.markdown("---")
                    df_resumen_mun = df_filtrado.groupby(['FECHA COMPLETA', 'MUNICIPIO', 'POBLADO', 'Usuario Responsable', 'JEFATURA', 'ESTADO_ADSCRIPCION']).agg({
                        'RECORRIDO': 'sum', 'GASTO COMBUSTI': 'sum'
                    }).reset_index().sort_values(by='FECHA COMPLETA')
                    st.dataframe(df_resumen_mun, use_container_width=True)
                    
                    output_filtrado = BytesIO()
                    df_res_mun = df_resumen_mun.copy()
                    df_res_mun.to_excel(output_filtrado, index=False, sheet_name="REPORTE_FILTRADO")
                    output_filtrado.seek(0)
                    st.download_button("Descargar Reporte Filtrado in Excel", data=output_filtrado, file_name="REPORTE_EJECUTIVO_FILTRADO.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                else:
                    st.warning("No se encontraron registros con los filtros seleccionados.")
            else:
                st.warning("No hay registros disponibles para tu nivel de adscripción.")
        else:
            st.info("Aún no hay registros de recorridos acumulados en el sistema.")

    def render_bitacora_audit():
        st.subheader("Bitácora de Auditoría del Sistema")
        if os.path.exists(AUDIT_FILE):
            try:
                with open(AUDIT_FILE, "r", encoding="utf-8") as f:
                    logs_data = json.load(f)
                if logs_data:
                    st.dataframe(pd.DataFrame(logs_data), use_container_width=True)
                else:
                    st.info("No hay registros de auditoría almacenados.")
            except:
                st.info("No se pudo leer el archivo de auditoría.")
        else:
            st.info("Aún no se han generado eventos de auditoría.")

    if current_email_key == "victor.olmedo@pa.gob.mx":
        with tab_reg_user: render_alta_usuario()
        with tab_edit_user: render_editar_usuario()
        with tab_ctrl_user: render_control_estatus()
        with tab_resumen_auditoria: render_resumen_auditoria()
        with tab_bitacora_audit: render_bitacora_audit()
    else:
        render_resumen_auditoria()
