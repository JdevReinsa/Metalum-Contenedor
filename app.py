import streamlit as st
import pandas as pd
import time
import urllib.parse
import os
import io
import hashlib

# 1. Configuración de la página
st.set_page_config(
    page_title="Carga contenedor", 
    page_icon="🚛", 
    layout="centered",
    initial_sidebar_state="collapsed"
)

# 🌓 CONTROL DE MODO OSCURO / MODO CLARO
if "tema_oscuro" not in st.session_state:
    st.session_state.tema_oscuro = False  # Inicia por defecto en modo claro

def cambiar_tema():
    st.session_state.tema_oscuro = not st.session_state.tema_oscuro

# 🔐 DETECTOR DE DISPOSITIVO INVIOLABLE
def obtener_id_dispositivo_unico():
    try:
        headers = st.context.headers
        user_agent = headers.get("User-Agent", "dispositivo_desconocido")
        host = headers.get("Host", "local")
        firma_secreta = f"{user_agent}_{host}"
        id_anonimo = hashlib.md5(firma_secreta.encode()).hexdigest()[:10]
        return id_anonimo
    except Exception:
        return "dispositivo_fijo"

ID_DISPOSITIVO = obtener_id_dispositivo_unico()
ARCHIVO_DATOS = f"registro_fardos_{ID_DISPOSITIVO}.csv"

# 📱 CONEXIÓN CON EL MANIFEST PARA EL ACCESO DIRECTO MÓVIL
st.markdown('<link rel="manifest" href="./manifest.json">', unsafe_allow_html=True)
st.markdown('<meta name="apple-mobile-web-app-title" content="Carga contenedor">', unsafe_allow_html=True)
st.markdown('<meta name="apple-mobile-web-app-capable" content="yes">', unsafe_allow_html=True)
st.markdown('<link rel="apple-touch-icon" href="https://cdn-icons-png.flaticon.com/512/3066/3066514.png">', unsafe_allow_html=True)

# 🎨 VARIABLES DE COLOR CORREGIDAS (Sin romper la estructura interna de los inputs)
if st.session_state.tema_oscuro:
    variables_color = """
    .stApp {
        background-color: #0e1117 !important;
        color: #ffffff !important;
    }
    p, h1, h2, h3, h4, h5, h6, span, label, small {
        color: #ffffff !important;
    }
    /* Estilo limpio para inputs en modo oscuro sin romper layouts */
    .stTextInput input, .stSelectbox div[data-baseweb="select"] {
        background-color: #1a1c23 !important;
        color: #ffffff !important;
        border: 1px solid #3f4452 !important;
    }
    """
    icono_tema = "☀️"
    label_tema = "Modo Claro"
else:
    variables_color = ""  # Usa el tema claro nativo y limpio de Streamlit
    icono_tema = "🌙"
    label_tema = "Modo Oscuro"

# Estilo CSS general corregido
st.markdown(f"""
    <style>
    {variables_color}
    
    [data-testid="stHeader"] {{
        visibility: hidden;
        height: 0% !important;
    }}
    [data-testid="viewerBadge"] {{
        visibility: hidden;
    }}
    [data-testid="stSidebarCollapse"] {{
        display: none !important;
    }}
    
    [data-testid="stDataFrame"] {{
        width: 100% !important;
        overflow-x: auto !important;
    }}
    
    /* 📌 FIJAR EL BOTÓN EN LA ESQUINA SUPERIOR DERECHA ABSOLUTA */
    div[data-testid="stVerticalBlock"] > div:has(button[key="btn_tema"]) {{
        position: fixed !important;
        top: 20px !important;
        right: 20px !important;
        z-index: 999999 !important;
        width: auto !important;
    }}
    
    button[key="btn_tema"] {{
        background-color: #1e3a8a !important;
        color: white !important;
        border: none !important;
        border-radius: 50% !important;
        width: 40px !important;
        height: 40px !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        font-size: 18px !important;
        box-shadow: 0px 4px 10px rgba(0,0,0,0.3) !important;
        cursor: pointer !important;
        transition: transform 0.2s ease !important;
    }}
    
    button[key="btn_tema"]:hover {{
        transform: scale(1.1) !important;
    }}
    
    .stButton>button, .stDownloadButton>button {{
        width: 100%;
        height: 55px;
        font-size: 16px;
        border-radius: 8px !important;
        transition: all 0.3s ease;
    }}
    
    .boton-normal>div>button {{ background-color: #1e3a8a !important; color: white !important; }}
    .boton-exito>div>button {{ background-color: #2e7d32 !important; color: white !important; font-weight: bold !important; }}
    .boton-error>div>button {{ background-color: #d32f2f !important; color: white !important; font-weight: bold !important; }}
    .boton-borrar>div>button {{ background-color: #555555 !important; color: white !important; height: 55px !important; }}
    
    .boton-wsp>div>a {{
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 10px;
        background-color: #25D366 !important;
        color: white !important;
        font-weight: bold !important;
        height: 55px;
        width: 100%;
        text-decoration: none;
        border-radius: 8px;
        font-size: 18px;
    }}
    
    .boton-excel-wsp>div>button {{
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        gap: 10px !important;
        background-color: #107c41 !important;
        color: white !important;
        font-weight: bold !important;
        font-size: 18px !important;
    }}
    
    /* Input height fixed sin romper paddings internos */
    .stTextInput input {{
        height: 45px !important;
    }}
    </style>
    """, unsafe_allow_html=True)

# BOTÓN EN LA ESQUINA SUPERIOR DERECHA
st.button(icono_tema, key="btn_tema", on_click=cambiar_tema, help=f"Cambiar a {label_tema}")

st.title("🏭 METALUM")
st.subheader("Registro de Contenedor")
st.caption(f"🔒 Conexión Protegida Privada (ID: {ID_DISPOSITIVO})")
st.divider()

# --- CARGAR DATOS ---
def cargar_datos():
    columnas_correctas = ["Ítem", "Folio", "Peso (Kg)", "Producto"]
    if os.path.exists(ARCHIVO_DATOS):
        try:
            df = pd.read_csv(ARCHIVO_DATOS)
            df["Ítem"] = df["Ítem"].astype(int)
            df["Folio"] = df["Folio"].astype(str)
            df = df.reindex(columns=columnas_correctas)
            return df
        except Exception:
            return pd.DataFrame(columns=columnas_correctas)
    return pd.DataFrame(columns=columnas_correctas)

# --- GUARDAR DATOS ---
def guardar_datos(df):
    df.to_csv(ARCHIVO_DATOS, index=False)

# --- LÓGICA DE BORRADO ---
def ejecutar_borrado_directo():
    folio_a_borrar = st.session_state.get("folio_a_borrar_input", "").strip()
    if not folio_a_borrar:
        st.session_state.mensaje_borrado = "❌ Error: Debes ingresar un número de Folio válido."
        return

    st.session_state.tabla_carga["Folio"] = st.session_state.tabla_carga["Folio"].astype(str)
    if folio_a_borrar not in st.session_state.tabla_carga["Folio"].values:
        st.session_state.mensaje_borrado = f"❌ Error: El Folio N° {folio_a_borrar} no existe en la carga actual."
    else:
        fila_seleccionada = st.session_state.tabla_carga[st.session_state.tabla_carga["Folio"] == folio_a_borrar].iloc[0]
        producto_borrado = fila_seleccionada["Producto"]
        
        st.session_state.tabla_carga = st.session_state.tabla_carga[st.session_state.tabla_carga["Folio"] != folio_a_borrar]
        st.session_state.tabla_carga["Ítem"] = range(1, len(st.session_state.tabla_carga) + 1)
        
        guardar_datos(st.session_state.tabla_carga)
        st.session_state.mensaje_borrado = f"✅ ¡Registro eliminado con éxito! Producto: {producto_borrado} (Folio/Código: {folio_a_borrar})"
        st.session_state["folio_a_borrar_input"] = ""

# --- GENERAR EXCEL ---
def generar_excel(df, patente_nom, total_b, total_k, total_p):
    output = io.BytesIO()
    df_excel = df.copy()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_excel.to_excel(writer, index=False, sheet_name='Reporte Carga')
        
        workbook  = writer.book
        worksheet = writer.sheets['Reporte Carga']
        
        header_format = workbook.add_format({
            'bold': True, 'text_wrap': True, 'align': 'left', 'valign': 'vcenter',
            'fg_color': '#107C41', 'font_color': 'white', 'border': 1
        })
        cell_format = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1})
        bold_label = workbook.add_format({'bold': True, 'align': 'left'})
        normal_label = workbook.add_format({'align': 'left'})
        
        for col_num, header in enumerate(df_excel.columns):
            worksheet.write(0, col_num, header, header_format)
            
        for row_idx in range(len(df_excel)):
            worksheet.write(row_idx + 1, 0, int(df_excel.iloc[row_idx, 0]), cell_format)
            worksheet.write(row_idx + 1, 1, str(df_excel.iloc[row_idx, 1]), cell_format)
            
            val_peso = df_excel.iloc[row_idx, 2]
            try:
                worksheet.write(row_idx + 1, 2, int(val_peso), cell_format)
            except ValueError:
                worksheet.write(row_idx + 1, 2, str(val_peso), cell_format)
                
            worksheet.write(row_idx + 1, 3, str(df_excel.iloc[row_idx, 3]), cell_format)

        for col_num, col_name in enumerate(df_excel.columns):
            max_len = len(str(col_name))
            for val in df_excel[col_name]:
                max_len = max(max_len, len(str(val)))
            worksheet.set_column(col_num, col_num, max_len + 5)
            
        start_row = len(df_excel) + 2
        worksheet.write(start_row, 1, "Patente Camión:", bold_label)
        worksheet.write(start_row, 2, patente_nom, normal_label)
        worksheet.write(start_row + 1, 1, "Total Bultos (Fardos):", bold_label)
        worksheet.write(start_row + 1, 2, f"{total_b} fardos", normal_label)
        worksheet.write(start_row + 2, 1, "Peso Total Fardos:", bold_label)
        worksheet.write(start_row + 2, 2, f"{total_k:,} Kg", normal_label)
        worksheet.write(start_row + 3, 1, "Total Pallets:", bold_label)
        worksheet.write(start_row + 3, 2, f"{total_p} unidades", normal_label)
        
    return output.getvalue()

# Inicializar Estados
if "tabla_carga" not in st.session_state:
    st.session_state.tabla_carga = cargar_datos()
if "estado_ultimo_fardo" not in st.session_state:
    st.session_state.estado_ultimo_fardo = "normal"
if "ultimo_folio_processed" not in st.session_state:
    st.session_state.ultimo_folio_processed = 0
if "folio_intentado" not in st.session_state:
    st.session_state.folio_intentado = 0
if "cantidad_pallets_processed" not in st.session_state:
    st.session_state.cantidad_pallets_processed = 0
if "form_reset_counter" not in st.session_state:
    st.session_state.form_reset_counter = 0
if "mensaje_borrado" not in st.session_state:
    st.session_state.mensaje_borrado = ""

st.session_state.tabla_carga = cargar_datos()

# 2. SELECCIÓN DE PRODUCTO
producto = st.selectbox(
    "Selecciona el Producto:",
    ["UBC", "Perfil", "Tense", "Taint Tabor", "Radiador", "Acero", "Offset", "Pallets"],
    key="producto_seleccionado"
)

# 3. FORMULARIO DE CARGA DINÁMICO
with st.form(key="formulario_fardo"):
    ctr = st.session_state.form_reset_counter
    
    if producto == "Pallets":
        cantidad_pallets_raw = st.text_input("Cantidad de Pallets:", placeholder="Escribe el número de pallets...", max_chars=5, key=f"pallets_{ctr}")
        peso_raw = "0"
        folio_raw = "0"
    else:
        peso_raw = st.text_input("Peso (Kg):", placeholder="Escribe el peso...", max_chars=8, key=f"peso_{ctr}")
        folio_raw = st.text_input("Número de Folio:", placeholder="Escribe el folio...", max_chars=8, key=f"folio_{ctr}")
        cantidad_pallets_raw = "0"
    
    f_proc = st.session_state.ultimo_folio_processed
    f_rep = st.session_state.folio_intentado
    p_proc = st.session_state.cantidad_pallets_processed
    
    if st.session_state.estado_ultimo_fardo == "exito":
        clase_boton = "boton-exito"
        texto_boton = f"✅ ¡FARDO #{f_proc} SUBIDO! (ENTER)"
    elif st.session_state.estado_ultimo_fardo == "exito_pallets":
        clase_boton = "boton-exito"
        texto_boton = f"✅ ¡{p_proc} PALLETS AÑADIDOS! (ENTER)"
    elif st.session_state.estado_ultimo_fardo == "error_duplicado":
        clase_boton = "boton-error"
        texto_boton = f"❌ ¡FOLIO #{f_rep} REPETIDO! ✖️"
    elif st.session_state.estado_ultimo_fardo == "error_vacio":
        clase_boton = "boton-error"
        texto_boton = "❌ ERROR: ¡DATOS VACÍOS O INVÁLIDOS! ✖️"
    else:
        clase_boton = "boton-normal"
        texto_boton = "➕ AGREGAR REGISTRO"

    st.markdown(f'<div class="{clase_boton}">', unsafe_allow_html=True)
    boton_guardar = st.form_submit_button(label=texto_boton)
    st.markdown('</div>', unsafe_allow_html=True)

# 4. LÓGICA DE PROCESAMIENTO
if boton_guardar:
    if producto == "Pallets":
        try:
            cant_p_val = int(cantidad_pallets_raw.strip()) if cantidad_pallets_raw else 0
        except ValueError:
            cant_p_val = 0
            
        if cant_p_val > 0:
            st.session_state.cantidad_pallets_processed = cant_p_val
            siguiente_item = st.session_state.tabla_carga["Ítem"].max() + 1 if len(st.session_state.tabla_carga) > 0 else 1
            
            nueva_fila = pd.DataFrame([{
                "Ítem": int(siguiente_item),
                "Folio": f"PL-{cant_p_val}", 
                "Peso (Kg)": "-",
                "Producto": "Pallets"
            }])
            st.session_state.tabla_carga = pd.concat([st.session_state.tabla_carga, nueva_fila], ignore_index=True)
            guardar_datos(st.session_state.tabla_carga)
            st.session_state.form_reset_counter += 1
            st.session_state.estado_ultimo_fardo = "exito_pallets"
            st.rerun()
        else:
            st.session_state.estado_ultimo_fardo = "error_vacio"
            st.rerun()
    else:
        try:
            peso_val = int(peso_raw.strip()) if peso_raw else 0
            folio_val = int(folio_raw.strip()) if folio_raw else 0
        except ValueError:
            peso_val = 0
            folio_val = 0
            
        st.session_state.folio_intentado = folio_val
        
        if peso_val > 0 and folio_val > 0:
            folios_existentes = st.session_state.tabla_carga["Folio"].astype(str).values
            if str(folio_val) in folios_existentes:
                st.session_state.estado_ultimo_fardo = "error_duplicado"
                st.rerun()
            else:
                st.session_state.ultimo_folio_processed = folio_val
                siguiente_item = st.session_state.tabla_carga["Ítem"].max() + 1 if len(st.session_state.tabla_carga) > 0 else 1
                
                nueva_fila = pd.DataFrame([{
                    "Ítem": int(siguiente_item),
                    "Folio": str(folio_val),
                    "Peso (Kg)": peso_val,
                    "Producto": producto
                }])
                
                st.session_state.tabla_carga = pd.concat([st.session_state.tabla_carga, nueva_fila], ignore_index=True)
                guardar_datos(st.session_state.tabla_carga)
                st.session_state.form_reset_counter += 1
                st.session_state.estado_ultimo_fardo = "exito"
                st.rerun()
        else:
            st.session_state.estado_ultimo_fardo = "error_vacio"
            st.rerun()

if st.session_state.estado_ultimo_fardo != "normal":
    time.sleep(1.2)
    st.session_state.estado_ultimo_fardo = "normal"
    st.rerun()

st.divider()

patente_texto = "NO REGISTRADA"

# 5. MONITOREO EN TIEMPO REAL
if not st.session_state.tabla_carga.empty:
    try:
        df_fardos = st.session_state.tabla_carga[st.session_state.tabla_carga["Producto"] != "Pallets"]
        df_pallets = st.session_state.tabla_carga[st.session_state.tabla_carga["Producto"] == "Pallets"]
        
        pesos_numericos = pd.to_numeric(df_fardos["Peso (Kg)"], errors='coerce').fillna(0)
        total_kg = int(pesos_numericos.sum())
        total_bultos = len(df_fardos)
        
        total_pallets_unidades = 0
        for _, fila in df_pallets.iterrows():
            partes = str(fila["Folio"]).split("-")
            if len(partes) == 2:
                try:
                    total_pallets_unidades += int(partes[1].strip())
                except ValueError:
                    pass
    except Exception:
        total_kg = 0
        total_bultos = 0
        total_pallets_unidades = 0

    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(label="TOTAL KILOS", value=f"{total_kg:,} Kg")
    with col2:
        st.metric(label="TOTAL BULTOS", value=f"{total_bultos} fardos")
    with col3:
        st.metric(label="TOTAL PALLETS", value=f"{total_pallets_unidades} und")
    
    st.divider()
    
    col_izq, col_der = st.columns([3, 2])
    with col_izq:
        st.write("### Detalle de la Carga Actual")
    with col_der:
        patente = st.text_input("Patente del Camión:", key="patente_camion", placeholder="EJ: AB-CD-12")
        patente_texto = patente.strip().upper() if patente.strip() != "" else "NO REGISTRADA"
    
    st.dataframe(
        st.session_state.tabla_carga, 
        use_container_width=True, 
        hide_index=True,
        column_config={
            "Ítem": st.column_config.Column(disabled=True, help="Fijo del sistema"),
            "Folio": st.column_config.Column(disabled=True, help="Fijo del sistema"),
            "Peso (Kg)": st.column_config.Column(disabled=True, help="Fijo del sistema"),
            "Producto": st.column_config.Column(disabled=False)
        }
    )
    
    st.divider()

    # REPORTES DE SALIDA
    st.write("### 📤 Reporte de Salida")
    
    mensaje_wsp = f"🚛 *REPORTE DE CARGA - METALUM*\n"
    mensaje_wsp += f"🔹 *Patente:* {patente_texto}\n"
    mensaje_wsp += f"----------------------------------------\n"
    mensaje_wsp += f"`Ítem | Folio/Cant | Peso | Producto`\n"
    for idx, fila in st.session_state.tabla_carga.iterrows():
        peso_txt = f"{fila['Peso (Kg)']} Kg" if fila['Peso (Kg)'] != "-" else "-"
        mensaje_wsp += f"{int(fila['Ítem'])} | {fila['Folio']} | {peso_txt} | {fila['Producto']}\n"
    mensaje_wsp += f"----------------------------------------\n"
    mensaje_wsp += f"📦 *Total Bultos (Fardos):* {total_bultos}\n"
    mensaje_wsp += f"⚖️ *Peso Total:* {total_kg:,} Kg\n"
    mensaje_wsp += f"🪵 *Total Pallets:* {total_pallets_unidades} unidades"
    
    texto_codificado = urllib.parse.quote(mensaje_wsp)
    enlace_whatsapp = f"https://api.whatsapp.com/send?text={texto_codificado}"
    
    st.markdown(f"""
        <div class="boton-wsp">
            <a href="{enlace_whatsapp}" target="_blank">
                <svg xmlns="http://www.w3.org/2000/svg" width="22" height="22" fill="currentColor" viewBox="0 0 16 16">
                    <path d="M13.601 2.326A7.85 7.85 0 0 0 7.994 0C3.627 0 .068 3.558.064 7.926c0 1.399.366 2.76 1.057 3.965L0 16l4.204-1.102a7.9 7.9 0 0 0 3.79.965h.004c4.368 0 7.926-3.558 7.93-7.93A7.9 7.9 0 0 0 13.6 2.326zM7.994 14.521a6.6 6.6 0 0 1-3.356-.92l-.24-.144-2.494.654.666-2.433-.156-.251a6.56 6.56 0 0 1-1.007-3.505c0-3.626 2.957-6.584 6.591-6.584a6.56 6.56 0 0 1 4.66 1.931 6.56 6.56 0 0 1 1.928 4.66c-.004 3.639-2.961 6.592-6.592 6.592m3.69-4.936c-.202-.101-1.202-.594-1.392-.661-.19-.068-.33-.101-.47.101-.14.202-.54.661-.661.8-.12.138-.24.154-.442.053-1.648-.826-2.617-1.963-3.064-2.731-.12-.207-.013-.319.09-.421.093-.092.202-.24.302-.36.101-.12.135-.2.203-.34.067-.137.034-.257-.017-.359-.051-.101-.47-1.136-.645-1.545-.171-.413-.344-.358-.47-.364-.121-.006-.26-.006-.4-.006-.14 0-.368.053-.56.26-.191.207-.73.714-.73 1.74s.747 2.009.85 2.147c.101.137 1.47 2.246 3.562 3.149.497.215.885.342 1.184.438.5.158.955.135 1.314.08.4-.061 1.202-.493 1.37-.967.169-.473.169-.88.119-.967-.051-.088-.18-.137-.383-.238"/>
                </svg>
                COMPARTIR REPORTE TEXTO
            </a>
        </div>
    """, unsafe_allow_html=True)
    
    st.write("") 
    
    data_excel = generar_excel(st.session_state.tabla_carga, patente_texto, total_bultos, total_kg, total_pallets_unidades)
    st.markdown('<div class="boton-excel-wsp">', unsafe_allow_html=True)
    st.download_button(
        label="📊 DESCARGAR EXCEL",  
        data=data_excel,
        file_name=f"Reporte_Metalum_{patente_texto}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.markdown('</div>', unsafe_allow_html=True)
    st.divider()
else:
    st.info("El contenedor está vacío. Empieza a registrar los fardos arriba.")

# --- 7. SECCIÓN CORREGIR ERRORES ---
st.write("### 🛠️ Corregir Errores")
folio_a_borrar_raw = st.text_input(
    "Digita el N° de Folio que deseas eliminar:", 
    placeholder="Escribe el número de folio aquí...",
    key="folio_a_borrar_input"
)

folio_a_borrar_limpio = folio_a_borrar_raw.strip()
if not folio_a_borrar_limpio:
    texto_boton = "🗑️ INGRESA UN FOLIO"
else:
    texto_boton = f"🗑️ ELIMINAR FOLIO N° {folio_a_borrar_limpio}"

st.markdown('<div class="boton-borrar">', unsafe_allow_html=True)
st.button(texto_boton, on_click=ejecutar_borrado_directo)
st.markdown('</div>', unsafe_allow_html=True)

if st.session_state.mensaje_borrado:
    if "❌" in st.session_state.mensaje_borrado:
        st.error(st.session_state.mensaje_borrado)
    else:
        st.success(st.session_state.mensaje_borrado)
    st.session_state.mensaje_borrado = ""

st.write("") 

if st.button("⚠️ Reiniciar Todo (Camión Nuevo)"):
    st.session_state.tabla_carga = pd.DataFrame(columns=["Ítem", "Folio", "Peso (Kg)", "Producto"])
    if os.path.exists(ARCHIVO_DATOS):
        os.remove(ARCHIVO_DATOS)
    st.session_state.estado_ultimo_fardo = "normal"
    st.session_state.ultimo_folio_processed = 0
    st.session_state.folio_intentado = 0
    st.session_state.cantidad_pallets_processed = 0
    st.session_state.mensaje_borrado = ""
    st.rerun()
