import streamlit as st
import pandas as pd
import openpyxl
import json
import os
import random
import pydeck as pdk
from io import BytesIO
from datetime import datetime, date

st.set_page_config(page_title="Sistema Nacional de Bitácoras - Procuraduría Agraria", layout="wide")

# Estilos institucionales (Guinda Morena)
st.markdown("""
    <style>
    h1, h2, h3 { color: #6B1D2F !important; }
    .stButton>button {
        background-color: #6B1D2F !important;
        color: white !important;
        border-radius: 6px !important;
        border: none !important;
        font-weight: bold;
    }
    .stButton>button:hover { background-color: #8A253D !important; color: white !important; }
    [data-testid="stFormSubmitButton"]>button {
        background-color: #6B1D2F !important;
        color: white !important;
        border-radius: 6px !important;
        font-weight: bold;
    }
    [data-testid="stFormSubmitButton"]>button:hover { background-color: #8A253D !important; }
    section[data-testid="stSidebar"] { background-color: #faf6f7; border-right: 1px solid #e0d0d3; }
    .frase-agraria {
        background-color: #fcf5f6; border-left: 5px solid #6B1D2F;
        padding: 12px 18px; border-radius: 4px; margin-bottom: 20px;
        font-style: italic; color: #4a2c33;
    }
    </style>
""", unsafe_allow_html=True)

# Rutas base
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
USUARIOS_FILE = os.path.join(BASE_DIR, "usuarios.json")
REGISTROS_FILE = os.path.join(BASE_DIR, "registros.json")
SOLICITUDES_FILE = os.path.join(BASE_DIR, "solicitudes_gasolina.json")
INCIDENCIAS_FILE = os.path.join(BASE_DIR, "incidencias_mecanicas.json")
AUDIT_FILE = os.path.join(BASE_DIR, "audit_log.json")

# --- DETECCIÓN AUTOMÁTICA DE PLANTILLA EXCEL ---
excel_files = [os.path.join(BASE_DIR, f) for f in os.listdir(BASE_DIR) if f.endswith('.xlsx') and not f.startswith("BITACORAS_OFICIALES")]
PLANTILLA_EXCEL = excel_files[0] if excel_files else os.path.join(BASE_DIR, "Bitacora_Actualizada_Formula (2).xlsx")

FOTOS_DIR = os.path.join(BASE_DIR, "fotos_perfil")
LOGO_FILE = os.path.join(BASE_DIR, "logo_pa.png")
MUN_FILE = os.path.join(BASE_DIR, "MUNICIPIOS_202606.xlsx")
LOC_FILE = os.path.join(BASE_DIR, "LOCALIDADES_202606.xlsx")
os.makedirs(FOTOS_DIR, exist_ok=True)

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

MUNICIPIOS_EDOMEX = [
    "ACAMBAY DE RUIZ CASTAÑEDA", "ACOLMAN", "ACULCO", "ALMOLOYA DE ALQUISIRAS",
    "ALMOLOYA DE JUÁREZ", "ALMOLOYA DEL RÍO", "AMANALCO", "AMATEPEC", "AMECAMECA",
    "APAXCO", "ATENCO", "ATIZAPÁN", "ATIZAPÁN DE ZARAGOZA", "ATLACOMULCO",
    "ATLAUTLA", "AXAPUSCO", "AYAPANGO", "CALIMAYA", "CAPULHUAC",
    "COACALCO DE BERRIOZÁBAL", "COATEPEC HARINAS", "COCOTITLÁN", "COYOTEPEC",
    "CUAUTITLÁN", "CUAUTITLÁN IZCALLI", "CHALCO", "CHAPA DE MOTA", "CHAPULTEPEC",
    "CHIAUTLA", "CHICOLOAPAN", "CHICONCUAC", "CHIMALHUACÁN", "DONATO GUERRA",
    "ECATEPEC DE MORELOS", "ECATZINGO", "HUEHUETOCA", "HUEYPOXTLA", "HUIXQUILUCAN",
    "ISIDRO FABELA", "IXTAPALUCA", "IXTAPAN DE LA SAL", "IXTAPAN DEL ORO",
    "IXTLAHUACA", "XALATLACO", "JALTENCO", "JILOTEPEC", "JILOTZINGO", "JIQUIPILCO",
    "JOCOTITLÁN", "JOQUICINGO", "JUCHITEPEC", "LERMA", "MALINALCO", "MELCHOR OCAMPO",
    "METEPEC", "MEXICALTZINGO", "MORELOS", "NAUCALPAN DE JUÁREZ", "NEXTLALPAN",
    "NEZAHUALCÓYOTL", "NICOLÁS ROMERO", "NOPALTEPEC", "OCOYOACAC", "OCUILAN",
    "EL ORO", "OTUMBA", "OTZOLOAPAN", "OTZOLOTEPEC", "OZUMBA", "PAPALOTLA",
    "LA PAZ", "POLOTITLÁN", "RAYÓN", "SAN ANTONIO LA ISLA", "SAN FELIPE DEL PROGRESO",
    "SAN MARTÍN DE LAS PIRÁMIDES", "SAN MATEO ATENCO", "SAN SIMÓN DE GUERRERO",
    "SANTO TOMÁS", "SOYANIQUILPAN DE JUÁREZ", "SULTEPEC", "TECÁMAC", "TEJUPILCO",
    "TEMAMATLA", "TEMASCALAPA", "TEMASCALCINGO", "TEMASCALTEPEC", "TEMOAYA",
    "TENANCINGO", "TENANGO DEL AIRE", "TENANGO DEL VALLE", "TEOLOYUCAN",
    "TEOTIHUACÁN", "TEPETLAOXTOC", "TEPETLIXPA", "TEPOTZOTLÁN", "TEQUIXQUIAC",
    "TEXCALTITLÁN", "TEXCALYACAC", "TEXCOCO", "TEZOYUCA", "TIANGUISTENCO",
    "TIMILPAN", "TLALMANALCO", "TLALNEPANTLA DE BAZ", "TLATLAYA", "TOLUCA",
    "TONATICO", "TULTEPEC", "TULTITLÁN", "VALLE DE BRAVO", "VILLA DE ALLENDE",
    "VILLA DEL CARBÓN", "VILLA GUERRERO", "VILLA VICTORIA", "XONACATLÁN",
    "ZACAZONAPAN", "ZACUALPAN", "ZINACANTEPEC", "ZUMPAHUACÁN", "ZUMPANGO",
    "VALLE DE CHALCO SOLIDARIDAD", "LUVIANOS", "SAN JOSÉ DEL RINCÓN", "TONANITLA"
]

JEFATURAS_RESIDENCIA = [
    "RESIDENCIA NAUCALPAN", "RESIDENCIA TOLUCA", "RESIDENCIA ATLACOMULCO",
    "RESIDENCIA TEXCOCO", "RESIDENCIA VALLE DE BRAVO", "RESIDENCIA TENANCINGO", "RESIDENCIA MORELIA"
]

ROLES_SISTEMA = [
    "Organizador Agrario (Operador)", "Analista de Información", 
    "Administrador de Vehículos (Estatal)", "Jefe de Residencia", 
    "Administrador Estatal", "Administrador Nacional"
]

FRASES_AGRARIAS = [
    "“La tierra no pertenece al hombre; el hombre pertenece a la tierra.” — Jefe Seattle",
    "“El agricultor es la única persona en el mundo que gasta dinero en esperar y arriesgarse a cosechar.” — E.W. Howe",
    "“Cultivar la tierra es servir a la patria con las manos y el corazón.”",
    "“El campo da pan a todos, y la justicia agraria da paz y dignidad a nuestra gente.”"
]

@st.cache_data
def cargar_catalogos_geograficos():
    muns_df = pd.read_excel(MUN_FILE) if os.path.exists(MUN_FILE) else pd.DataFrame()
    locs_df = pd.read_excel(LOC_FILE) if os.path.exists(LOC_FILE) else pd.DataFrame()
    return muns_df, locs_df

muns_global, locs_global = cargar_catalogos_geograficos()

def obtener_municipios_estado(estado_nombre):
    if muns_global.empty or estado_nombre not in ESTADOS_REPUBLICA:
        return MUNICIPIOS_EDOMEX if estado_nombre == "Estado de México" else ["Cabecera Municipal"]
    efe_key = ESTADOS_REPUBLICA.index(estado_nombre) + 1
    if 'EFE_KEY' in muns_global.columns:
        muns = muns_global[muns_global['EFE_KEY'] == efe_key]['MUNICIPIO'].dropna().tolist()
    else:
        muns = muns_global.iloc[:, 1].dropna().tolist()
    return sorted(list(set(muns))) if muns else ["Cabecera Municipal"]

def obtener_localidades_municipio(estado_nombre, municipio_nombre):
    if locs_global.empty:
        return ["Cabecera Municipal"]
    efe_key = ESTADOS_REPUBLICA.index(estado_nombre) + 1 if estado_nombre in ESTADOS_REPUBLICA else 1
    df_locs = locs_global.copy()
    if 'MUNICIPIO' in df_locs.columns:
        locs = df_locs[df_locs['MUNICIPIO'].str.upper() == municipio_nombre.upper()]['LOCALIDAD'].dropna().tolist()
    else:
        locs = df_locs.iloc[:, 2].dropna().tolist()
    return sorted(list(set(locs))) if locs else ["Cabecera Municipal"]

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
                saved = json.load(f)
                for email, data in usuarios_base.items():
                    if email not in saved:
                        saved[email] = data
                return saved
        except:
            pass
    return usuarios_base

def guardar_usuarios(d):
    with open(USUARIOS_FILE, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=4)

def cargar_registros():
    if os.path.exists(REGISTROS_FILE):
        try:
            with open(REGISTROS_FILE, "r", encoding="utf-8") as f: return json.load(f)
        except: pass
    return []

def guardar_registros(r):
    with open(REGISTROS_FILE, "w", encoding="utf-8") as f: json.dump(r, f, ensure_ascii=False, indent=4)

def registrar_auditoria(acc, det):
    entry = {"FECHA_HORA": datetime.now().strftime("%d/%m/%Y %H:%M:%S"), "USUARIO": st.session_state.get("current_email", "Sistema"), "ACCION": acc, "DETALLE": det}
    logs = []
    if os.path.exists(AUDIT_FILE):
        try:
            with open(AUDIT_FILE, "r", encoding="utf-8") as f: logs = json.load(f)
        except: pass
    logs.append(entry)
    with open(AUDIT_FILE, "w", encoding="utf-8") as f: json.dump(logs, f, ensure_ascii=False, indent=4)

OPCIONES_GASOLINA = {"🔴 1/4": "1/4", "🟡 1/2": "1/2", "🟢 3/4": "3/4", "🟢 1/1": "1/1", "🔴 V": "V"}

for k, v in [("logged_in", False), ("current_email", ""), ("current_user", ""), ("current_licencia", ""), ("current_rol", ""), ("current_estado", ""), ("current_jefatura", ""), ("current_jefe_residencia", "")]:
    if k not in st.session_state: st.session_state[k] = v

if not st.session_state["logged_in"]:
    col1, col2 = st.columns([0.15, 2.85])
    with col1:
        if os.path.exists(LOGO_FILE): st.image(LOGO_FILE, width=60)
    with col2: st.title("Acceso al Sistema Nacional de Bitácoras")
    st.markdown("---")
    _, col_mid, _ = st.columns([1, 2, 1])
    with col_mid:
        st.subheader("Iniciar Sesión")
        email_in = st.text_input("Correo Institucional").strip().lower()
        pass_in = st.text_input("Contraseña", type="password")
        if st.button("🔑 Ingresar", use_container_width=True):
            users = cargar_usuarios()
            if email_in in users and users[email_in]["pass"] == pass_in and users[email_in].get("activo", True):
                st.session_state["logged_in"] = True
                st.session_state["current_email"] = email_in
                st.session_state["current_user"] = users[email_in]["nombre"]
                st.session_state["current_licencia"] = users[email_in]["licencia"]
                st.session_state["current_rol"] = users[email_in].get("rol", "Organizador Agrario (Operador)")
                st.session_state["current_estado"] = users[email_in].get("estado", "Estado de México")
                st.session_state["current_jefatura"] = users[email_in].get("jefatura", "RESIDENCIA NAUCALPAN")
                st.session_state["current_jefe_residencia"] = users[email_in].get("jefe_residencia", "N/A")
                registrar_auditoria("LOGIN", f"Acceso de {email_in}")
                st.success("¡Acceso concedido!")
                st.rerun()
            else:
                st.error("⚠️ Usuario o contraseña incorrectos.")
    st.stop()

# --- APLICACIÓN PRINCIPAL ---
col1, col2 = st.columns([0.15, 2.85])
with col1:
    if os.path.exists(LOGO_FILE): st.image(LOGO_FILE, width=60)
with col2: st.title("Sistema Nacional de Control Vehicular y Bitácoras")

st.markdown(f'<div class="frase-agraria">🌾 <b>Reflexión del Día:</b> {random.choice(FRASES_AGRARIAS)}</div>', unsafe_allow_html=True)
st.markdown("---")

if os.path.exists(LOGO_FILE): st.sidebar.image(LOGO_FILE, use_container_width=True)
rol_actual = st.session_state.get("current_rol", "")

st.sidebar.title("📌 Menú")
modulos = ["Módulo de Captura (Recorrido)"]
if rol_actual == "Administrador Nacional":
    modulos.append("Panel de Administración")
perfil = st.sidebar.radio("Selecciona módulo:", modulos)

st.sidebar.markdown("---")
if st.sidebar.button("🚪 Cerrar Sesión"):
    st.session_state["logged_in"] = False
    st.rerun()

if perfil == "Módulo de Captura (Recorrido)":
    st.subheader("📝 Captura de Recorrido Diario")
    est_act = st.session_state["current_estado"]
    jef_act = st.session_state["current_jefatura"]
    
    muns = obtener_municipios_estado(est_act)
    
    with st.form("form_captura"):
        c1, c2, c3 = st.columns(3)
        with c1:
            fecha = st.date_input("Fecha", value=date.today())
            mun = st.selectbox("Municipio", muns)
            loc = st.selectbox("Localidad", obtener_localidades_municipio(est_act, mun))
            folio = st.text_input("Folio CIIA")
        with c2:
            h_sal = st.text_input("Hora Salida", value="09:00")
            km_ini = st.number_input("KM Inicial", min_value=0.0, value=0.0, step=1.0)
            h_lle = st.text_input("Hora Llegada", value="17:00")
            km_fin = st.number_input("KM Final", min_value=0.0, value=0.0, step=1.0)
        with c3:
            veh = st.selectbox("Vehículo", ["NISSAN VERSA", "PickUp", "Estacas"])
            placas = st.text_input("Placas", value="MGX-543-A")
            dot = st.number_input("Dotación Gasolina ($)", min_value=0.0, value=200.0)
            obs = st.text_input("Observaciones / Ruta")
            
        guardar = st.form_submit_button("💾 Guardar Día")
        if guardar:
            if km_fin < km_ini:
                st.error("⚠️ El KM Final no puede ser menor al KM Inicial.")
            else:
                recorrido = km_fin - km_ini
                nuevo = {
                    "FECHA COMPLETA": fecha.strftime("%d/%m/%Y"),
                    "MES": fecha.strftime("%B").upper(),
                    "MUNICIPIO": mun,
                    "POBLADO": loc,
                    "folio CIIA": folio,
                    "HORA DE SALIDA": h_sal,
                    "KM INICIAL / Km de Salida": km_ini,
                    "RECORRIDO": recorrido,
                    "HORA DE LLEGADA": h_lle,
                    "KM FINAL / Km de Llegada": km_fin,
                    "GASTO COMBUSTI": dot,
                    "Gasolina de Salida": "1/4",
                    "Gasolina de Llegada": "1/2",
                    "Dotación de Gasolina(LLENAR GASTO DE COMBUSTIBLE)": dot,
                    "Oficio Numero": "N/A", "Oficio Fecha": "N/A",
                    "observaciones": obs,
                    "Usuario Responsable": st.session_state["current_user"],
                    "Áreas de Adscripción": jef_act,
                    "Tipo de Vehículo": veh,
                    "Placas": placas.upper(),
                    "No. De Licencia": st.session_state["current_licencia"],
                    "ESTADO_ADSCRIPCION": est_act,
                    "JEFATURA": jef_act
                }
                regs = cargar_registros()
                regs.append(nuevo)
                guardar_registros(regs)
                st.success("✅ ¡Recorrido guardado con éxito!")

    regs = cargar_registros()
    if regs:
        st.markdown("---")
        st.subheader("📋 Historial y Generación de Bitácoras Oficiales")
        st.dataframe(pd.DataFrame(regs), use_container_width=True)
        
        c_del, c_gen = st.columns(2)
        with c_del:
            if st.button("🗑️ Limpiar Historial"):
                guardar_registros([])
                st.rerun()
        with c_gen:
            if st.button("🚀 Generar y Descargar 3 Bitácoras Oficiales"):
                if not os.path.exists(PLANTILLA_EXCEL):
                    st.error(f"⚠️ No se encontró la plantilla en: {PLANTILLA_EXCEL}")
                else:
                    try:
                        wb = openpyxl.load_workbook(PLANTILLA_EXCEL)
                        ws_b = wb["BASE_DE_DATOS"]
                        
                        # Limpiar filas previas
                        for r in range(2, max(33, ws_b.max_row) + 1):
                            for col in ['A','B','C','D','E','F','G','H','I','J','L','M','O','P','Q','R','S','T','U','V']:
                                ws_b[f'{col}{r}'] = None

                        for i, reg in enumerate(regs, start=2):
                            ws_b[f'A{i}'] = datetime.strptime(reg["FECHA COMPLETA"], "%d/%m/%Y").date()
                            ws_b[f'A{i}'].number_format = "dd/mm/yyyy"
                            ws_b[f'B{i}'] = f'=UPPER(TEXT(A{i}, "MMMM"))'
                            ws_b[f'C{i}'] = reg["MUNICIPIO"]
                            ws_b[f'D{i}'] = reg["POBLADO"]
                            ws_b[f'E{i}'] = reg["folio CIIA"]
                            ws_b[f'F{i}'] = reg["HORA DE SALIDA"]
                            ws_b[f'G{i}'] = reg["KM INICIAL / Km de Salida"] if i == 2 else f'=J{i-1}'
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

                        row_tot = len(regs) + 2
                        ws_b[f'A{row_tot}'] = "TOTALES"
                        ws_b[f'H{row_tot}'] = f'=SUM(H2:H{row_tot-1})'

                        output = BytesIO()
                        wb.save(output)
                        output.seek(0)
                        st.success("✅ ¡Bitácoras generadas con éxito!")
                        st.download_button("⬇️ Descargar Archivo Oficial Definitivo", data=output, file_name="BITACORAS_OFICIALES_DEFINITIVAS.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    except Exception as e:
                        st.error(f"Error al procesar el Excel: {e}")

elif perfil == "Panel de Administración":
    st.subheader("⚙️ Panel de Administración de Usuarios")
    users = cargar_usuarios()
    for email, d in users.items():
        st.write(f"**{email}** - {d['nombre']} ({d['rol']})")
