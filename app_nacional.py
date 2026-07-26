import streamlit as st
import pandas as pd
import openpyxl
import json
import os
from io import BytesIO

st.set_page_config(page_title="Sistema Nacional de Bitácoras - Procuraduría Agraria", layout="wide")

# Archivo persistente para guardar usuarios
USUARIOS_FILE = "usuarios.json"

def cargar_usuarios():
    if os.path.exists(USUARIOS_FILE):
        try:
            with open(USUARIOS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            pass
    # Usuario por defecto si no existe el archivo
    return {
        "victor.olmedo@pa.gob.mx": {
            "nombre": "VÍCTOR LEONARDO OLMEDO GONZALEZ",
            "pass": "Leonardo",
            "licencia": "0101P3402484l"
        }
    }

def guardar_usuarios(usuarios_dict):
    with open(USUARIOS_FILE, "w", encoding="utf-8") as f:
        json.dump(usuarios_dict, f, ensure_ascii=False, indent=4)

# Catálogo oficial de los 125 municipios del Estado de México (Sin duplicados)
MUNICIPIOS_EDOMEX = sorted([
    "Acambay de Ruíz Castañeda", "Acolman", "Aculco", "Almoloya de Alquisiras", "Almoloya de Juárez", 
    "Almoloya del Río", "Amanalco", "Amatepec", "Amecameca", "Apaxco", "Atenco", "Atizapán", 
    "Atizapán de Zaragoza", "Atlautla", "Atlacomulco", "Axapusco", "Ayapango", "Calimaya", 
    "Capulhuac", "Chalco", "Chapa de Mota", "Chapultepec", "Chiautla", "Chicoloapan", 
    "Chiconcuac", "Chimalhuacán", "Coacalco de Berriozábal", "Coatepec Harinas", "Cocotitlán", 
    "Coyotepec", "Cuautitlán", "Cuautitlán Izcalli", "Donato Guerra", "Ecatepec de Morelos", 
    "Ecatzingo", "El Oro", "Hueypoxtla", "Huixquilucan", "Isidro Fabela", "Ixtapaluca", 
    "Ixtapan de la Sal", "Ixtapan del Oro", "Ixtlahuaca", "Xalatlaco", "Jaltenco", "Jilotepec", 
    "Jilotzingo", "Jiquipilco", "Jocotitlán", "Joquicingo", "Juchitepec", "La Paz", "Lerma", 
    "Luvianos", "Malinalco", "Melchor Ocampo", "Metepec", "Mexicaltzingo", "Morelos", 
    "Naucalpan de Juárez", "Nezahualcóyotl", "Nextlalpan", "Nicolás Romero", "Nopaltepec", 
    "Ocoyoacac", "Ocuilan", "Otumba", "Otzoloapan", "Otzolotepec", "Ozumba", "Papalotla", 
    "Polotitlán", "Rayón", "San Antonio la Isla", "San Felipe del Progreso", "San José del Rincón", 
    "San Martín de las Pirámides", "San Mateo Atenco", "San Simón de Guerrero", "Santo Tomás", 
    "Soyaniquilpan de Juárez", "Sultepec", "Tecámac", "Tejupilco", "Temamatla", "Temascalapa", 
    "Temascalcingo", "Temascaltepec", "Temoaya", "Tenancingo", "Tenango del Aire", 
    "Tenango del Valle", "Teoloyucan", "Teotihuacán", "Tepetlaoxtoc", "Tepetlixpa", 
    "Tepotzotlán", "Tequixquiac", "Texcaltitlán", "Texcalyacac", "Texcoco", "Tezoyuca", 
    "Tianguistenco", "Timilpan", "Tlalmanalco", "Tlalnepantla de Baz", "Tlatlaya", "Toluca", 
    "Tonanitla", "Tonatico", "Tultepec", "Tultitlán", "Valle de Bravo", "Valle de Chalco Solidaridad", 
    "Villa de Allende", "Villa del Carbón", "Villa Guerrero", "Villa Victoria", "Xonacatlán", 
    "Zacazonapan", "Zacualpan", "Zinacantepec", "Zumpahuacán", "Zumpango"
])

# Inicialización de estados en sesión
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
if "usuarios" not in st.session_state:
    st.session_state["usuarios"] = cargar_usuarios()
if "registros_acumulados" not in st.session_state:
    st.session_state["registros_acumulados"] = []

# --- PANTALLA DE LOGIN ---
if not st.session_state["logged_in"]:
    st.title("🔐 Acceso al Sistema Nacional de Bitácoras")
    st.markdown("Procuraduría Agraria - Módulo de Autenticación")
    st.markdown("---")
    
    col_l1, col_l2, col_l3 = st.columns([1, 2, 1])
    with col_l2:
        st.subheader("Iniciar Sesión")
        email_input = st.text_input("Correo Electrónico Institucional").strip().lower()
        pass_input = st.text_input("Contraseña", type="password")
        
        if st.button("🔑 Ingresar al Sistema", use_container_width=True):
            usuarios_actuales = cargar_usuarios()
            if email_input in usuarios_actuales and usuarios_actuales[email_input]["pass"] == pass_input:
                st.session_state["logged_in"] = True
                st.session_state["current_email"] = email_input
                st.session_state["current_user"] = usuarios_actuales[email_input]["nombre"]
                st.session_state["current_licencia"] = usuarios_actuales[email_input]["licencia"]
                st.success("¡Acceso concedido! Cargando sistema...")
                st.rerun()
            else:
                st.error("⚠️ Usuario o contraseña incorrectos. Verifica tus datos.")
    st.stop()

# --- APLICACIÓN PRINCIPAL (UNA VEZ LOGUEADO) ---
st.title("🌐 Sistema Nacional de Control Vehicular y Bitácoras")
st.markdown("---")

st.sidebar.title("👤 Sesión Activa")
st.sidebar.write(f"**Usuario:** {st.session_state['current_user']}")
st.sidebar.write(f"**Correo:** {st.session_state['current_email']}")

perfil = st.sidebar.selectbox("Selecciona tu rol:", ["Operador de Residencia", "Administrador Nacional (Sede)"])

if st.sidebar.button("🚪 Cerrar Sesión"):
    st.session_state["logged_in"] = False
    st.rerun()

if perfil == "Operador de Residencia":
    st.subheader("📝 Módulo de Captura por Día - Residencia")
    st.markdown("Ingresa los datos de tu recorrido diario. Las horas deben ser en formato de 24 horas y diferentes entre sí.")
    
    with st.form("form_captura_dia"):
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            fecha = st.date_input("Fecha Completa")
            municipio = st.selectbox("Municipio (Edo. Méx.)", MUNICIPIOS_EDOMEX, index=MUNICIPIOS_EDOMEX.index("Toluca") if "Toluca" in MUNICIPIOS_EDOMEX else 0)
            poblado = st.text_input("Poblado / Comisión", value="")
            folio_ciia = st.text_input("Folio CIIA", value="")
        with col2:
            h_salida = st.text_input("Hora de Salida (Formato 24h, ej. 09:00)", value="09:00")
            km_inicial = st.number_input("KM Inicial / Salida", min_value=0.0, value=23379.0, step=1.0)
            h_llegada = st.text_input("Hora de Llegada (Formato 24h, ej. 17:00)", value="17:00")
            km_final = st.number_input("KM Final / Llegada", min_value=0.0, value=23479.0, step=1.0)
        with col3:
            residencia = st.selectbox("Área de Adscripción", [
                "RESIDENCIA NAUCALPAN", 
                "RESIDENCIA TOLUCA", 
                "RESIDENCIA ATLACOMULCO", 
                "RESIDENCIA TEXCOCO", 
                "RESIDENCIA VALLE DE BRAVO", 
                "RESIDENCIA TENANCINGO"
            ])
            vehiculo = st.selectbox("Tipo de Vehículo", ["NISSAN VERSA", "PickUp", "Estacas"])
            placas = st.text_input("Placas", value="MGX-543-A")
            licencia = st.text_input("No. De Licencia", value=st.session_state["current_licencia"])
        with col4:
            usuario = st.text_input("Usuario Responsable", value=st.session_state["current_user"])
            dotacion = st.number_input("Dotación de Gasolina ($)", min_value=0.0, value=200.0, step=1.0)
            gas_salida = st.selectbox("Gasolina Salida", ["1/4", "1/2", "3/4", "1/1", "V"])
            gas_llegada = st.selectbox("Gasolina Llegada", ["1/4", "1/2", "3/4", "1/1", "V"])
        
        st.markdown("---")
        col_o1, col_o2, col_o3 = st.columns(3)
        with col_o1:
            oficio_num = st.text_input("Oficio Número (Opcional)", value="")
        with col_o2:
            oficio_fecha = st.text_input("Oficio Fecha (Opcional)", value="")
        with col_o3:
            observaciones = st.text_input("Observaciones / Ruta", value="")
        
        guardar_dia = st.form_submit_button("💾 Guardar Día")
        
        if guardar_dia:
            if h_salida.strip() == h_llegada.strip():
                st.error("⚠️ Error: La Hora de Salida y la Hora de Llegada no pueden ser iguales en formato 24 horas.")
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
                    "Gasolina de Salida": gas_salida,
                    "Gasolina de Llegada": gas_llegada,
                    "Dotación de Gasolina(LLENAR GASTO DE COMBUSTIBLE)": dotacion,
                    "Oficio Numero": oficio_num if oficio_num else None,
                    "Oficio Fecha": oficio_fecha if oficio_fecha else None,
                    "observaciones": observaciones if observaciones else None,
                    "Usuario Responsable": usuario,
                    "Áreas de Adscripción": residencia,
                    "Tipo de Vehículo": vehiculo,
                    "Placas": placas,
                    "No. De Licencia": licencia
                }
                st.session_state["registros_acumulados"].append(nuevo_reg)
                st.success(f"✅ ¡Día {fecha.strftime('%d/%m/%Y')} en {municipio} guardado correctamente!")

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
    
    with st.expander("➕ Dar de alta nuevo usuario al sistema", expanded=True):
        with st.form("form_nuevo_usuario"):
            c_email = st.text_input("Correo Electrónico (Usuario)")
            c_nombre = st.text_input("Nombre Completo (Mayúsculas)")
            c_pass = st.text_input("Contraseña Asignada", type="password")
            c_licencia = st.text_input("Número de Licencia de Conducir")
            
            btn_crear = st.form_submit_button("Registrar Usuario")
            if btn_crear:
                if c_email and c_nombre and c_pass:
                    usuarios_actuales = cargar_usuarios()
                    email_limpio = c_email.strip().lower()
                    usuarios_actuales[email_limpio] = {
                        "nombre": c_nombre.strip().upper(),
                        "pass": c_pass.strip(),
                        "licencia": c_licencia.strip()
                    }
                    guardar_usuarios(usuarios_actuales)
                    st.session_state["usuarios"] = usuarios_actuales
                    st.success(f"¡Usuario {c_nombre} registrado exitosamente y guardado en el sistema!")
                else:
                    st.error("⚠️ Por favor completa los campos obligatorios (Correo, Nombre y Contraseña).")
    
    st.markdown("---")
    st.subheader("📋 Usuarios Registrados en el Sistema")
    usuarios_actuales = cargar_usuarios()
    df_users = pd.DataFrame([
        {"Correo": k, "Nombre": v["nombre"], "Licencia": v["licencia"]} 
        for k, v in usuarios_actuales.items()
    ])
    st.dataframe(df_users, use_container_width=True)
    
    st.markdown("---")
    st.subheader("📋 Registros Acumulados en Sesión")
    if len(st.session_state["registros_acumulados"]) > 0:
        df_sede = pd.DataFrame(st.session_state["registros_acumulados"])
        st.dataframe(df_sede, use_container_width=True)
        st.metric("Total de Días Registrados", len(df_sede))
    else:
        st.info("Aún no hay días guardados en la sesión actual.")
