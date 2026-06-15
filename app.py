import streamlit as st
import pandas as pd
import time
import urllib.parse
import os
import io  # Requerido para generar el Excel en memoria

# Nombre del archivo donde se guardará la información de forma permanente
ARCHIVO_DATOS = "registro_fardos_metalum.csv"

# 1. Configuración de la página (Oculta el menú lateral por defecto)
st.set_page_config(
    page_title="Metalum - Carga", 
    page_icon="🚛", 
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Estilo CSS inyectado para optimizar el tamaño en celulares y quitar elementos molestos
st.markdown("""
    <style>
    /* 1. Esconde por completo la barra superior (Menú, tu foto de perfil y estado de conexión) */
    [data-testid="stHeader"] {
        visibility: hidden;
        height: 0% !important;
    }
    
    /* 2. Esconde el logo de Streamlit ("Made with Streamlit") en el pie de página */
    [data-testid="viewerBadge"] {
        visibility: hidden;
    }
    
    /* Esconde el botón del menú de hamburguesa de la esquina superior izquierda */
    [data-testid="stSidebarCollapse"] {
        display: none !important;
    }
    
    /* Estilo base para el botón unificado alargado */
    .stButton>button {
        width: 100%;
        height: 55px;
        font-size: 16px;
        border-radius: 8px !important;
        transition: all 0.3s ease;
    }
    
    /* ESTILOS DINÁMICOS PARA EL BOTÓN UNIFICADO */
    .boton-normal>div>button {
        background-color: #1e3a8a !important; 
        color: white !important;
    }
    .boton-exito>div>button {
        background-color: #2e7d32 !important; 
        color: white !important;
        font-weight: bold !important;
        font-size: 15px;
    }
    .boton-error>div>button {
        background-color: #d32f2f !important; 
        color: white !important;
        font-weight: bold !important;
        font-size: 15px;
    }
    
    .boton-borrar>div>button {
        background-color: #555555 !important;
        color: white !important;
        height: 45px !important;
    }
    
    /* Estilo para el botón verde de WhatsApp optimizado con ícono */
    .boton-wsp>div>a {
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 8px;
        background-color: #25D366 !important;
        color: white !important;
        font-weight: bold !important;
        height: 55px;
        width: 100%;
        text-decoration: none;
        border-radius: 8px;
        font-size: 18px;
    }
    input {
        height: 45px !important;
        font-size: 16px !important;
    }
    </style>
    """, unsafe_allow_html=True)

# Título de la App
st.title("🏭 METALUM")
st.subheader("Registro de Contenedor")
st.divider()

# --- FUNCIÓN: CARGAR DATOS AL INICIAR/RECARGAR ---
def cargar_datos():
    # CAMBIO: Se reordenaron las columnas para que Folio vaya antes de Peso
    columnas_correctas = ["Ítem", "Folio", "Peso (Kg)", "Producto"]
    if os.path.exists(ARCHIVO_DATOS):
        try:
            df = pd.read_csv(ARCHIVO_DATOS)
            # Forzamos a que tanto Peso como Folio se carguen estrictamente como enteros
            df["Peso (Kg)"] = df["Peso (Kg)"].astype(int)
            df["Folio"] = df["Folio"].astype(int)
            df = df.reindex(columns=columnas_correctas)
            return df
        except Exception:
            return pd.DataFrame(columns=columnas_correctas)
    return pd.DataFrame(columns=columnas_correctas)

# --- FUNCIÓN: GUARDAR DATOS EN EL ARCHIVO LOCAL ---
def guardar_datos(df):
    df.to_csv(ARCHIVO_DATOS, index=False)

# --- FUNCIÓN: GENERAR EXCEL EN MEMORIA ---
def generar_excel(df, patente_nom):
    output = io.BytesIO()
    # Copia del dataframe reseteando el índice para que "Ítem" sea una columna normal en el Excel
    df_excel = df.copy()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_excel.to_excel(writer, index=False, sheet_name='Reporte Carga')
        # Puedes personalizar el diseño aquí si lo deseas en el futuro
    return output.getvalue()

# 2. Inicializar variables en memoria (con respaldo de archivo)
if "tabla_carga" not in st.session_state:
    st.session_state.tabla_carga = cargar_datos()

if "estado_ultimo_fardo" not in st.session_state:
    st.session_state.estado_ultimo_fardo = "normal"

if "ultimo_folio_processed" not in st.session_state:
    st.session_state.ultimo_folio_processed = 0

if "folio_intentado" not in st.session_state:
    st.session_state.folio_intentado = 0

# Contador interno que usamos para limpiar el Peso y Folio de forma controlada sin perder el Producto
if "form_reset_counter" not in st.session_state:
    st.session_state.form_reset_counter = 0

# 3. FORMULARIO DE CARGA (Sin clear_on_submit automático para mantener el Producto seleccionado)
with st.form(key="formulario_fardo"):
    # CAMBIO: Se actualizó "Taint" por "Taint Tabor"
    producto = st.selectbox(
        "Selecciona el Producto:",
        ["UBC", "Perfil", "Tense", "Taint Tabor", "Radiador", "Acero", "Offset"],
        key="producto_seleccionado"
    )
    
    # El peso y el folio cambian su key dinámicamente cuando se sube un fardo con éxito, obligándolos a quedar en blanco
    ctr = st.session_state.form_reset_counter
    peso = st.number_input("Peso (Kg):", min_value=0, step=1, value=None, placeholder="Escribe el peso...", key=f"peso_{ctr}")
    folio = st.number_input("Número de Folio:", min_value=0, step=1, value=None, placeholder="Escribe el folio...", key=f"folio_{ctr}")
    
    # --- FUSIÓN: UN SOLO BOTÓN ALARGADO DINÁMICO ---
    f_proc = st.session_state.ultimo_folio_processed
    f_rep = st.session_state.folio_intentado
    
    if st.session_state.estado_ultimo_fardo == "exito":
        clase_boton = "boton-exito"
        texto_boton = f"✅ ¡FARDO #{f_proc} SUBIDO CON ÉXITO! ✔️"
    elif st.session_state.estado_ultimo_fardo == "error_duplicado":
        clase_boton = "boton-error"
        texto_boton = f"❌ ¡FOLIO #{f_rep} RECHAZADO (REPETIDO)! ✖️"
    elif st.session_state.estado_ultimo_fardo == "error_vacio":
        clase_boton = "boton-error"
        texto_boton = "❌ ERROR: ¡DATOS INCOMPLETOS O VACÍOS! ✖️"
    else:
        clase_boton = "boton-normal"
        texto_boton = "➕ AGREGAR FARDO"

    st.markdown(f'<div class="{clase_boton}">', unsafe_allow_html=True)
    boton_guardar = st.form_submit_button(label=texto_boton)
    st.markdown('</div>', unsafe_allow_html=True)

# 4. LÓGICA DE PROCESAMIENTO
if boton_guardar:
    peso_val = int(peso) if peso is not None else 0
    folio_val = int(folio) if folio is not None else 0
    
    st.session_state.folio_intentado = folio_val
    
    # Ambos campos deben ser enteros mayores a 0 para ser válidos
    if peso_val > 0 and folio_val > 0:
        folios_existentes = st.session_state.tabla_carga["Folio"].astype(int).values
        
        if folio_val in folios_existentes:
            st.session_state.estado_ultimo_fardo = "error_duplicado"
            st.rerun()
        else:
            st.session_state.ultimo_folio_processed = folio_val
            
            if len(st.session_state.tabla_carga) > 0:
                siguiente_item = st.session_state.tabla_carga["Ítem"].max() + 1
            else:
                siguiente_item = 1
            
            nueva_fila = pd.DataFrame([{
                "Ítem": int(siguiente_item),
                "Folio": folio_val,        # CAMBIO: Sigue el nuevo orden estructural
                "Peso (Kg)": peso_val,
                "Producto": producto
            }])
            
            st.session_state.tabla_carga = pd.concat([st.session_state.tabla_carga, nueva_fila], ignore_index=True)
            guardar_datos(st.session_state.tabla_carga)
            
            # Éxito: Cambiamos el contador para limpiar Peso y Folio, manteniendo intacto el material (ej. UBC)
            st.session_state.form_reset_counter += 1
            st.session_state.estado_ultimo_fardo = "exito"
            st.rerun()
    else:
        # Se presionó el botón pero faltaron datos reales
        st.session_state.estado_ultimo_fardo = "error_vacio"
        st.rerun()

# --- TEMPORIZADOR AUTOMÁTICO SEGURO ---
if st.session_state.estado_ultimo_fardo != "normal":
    time.sleep(1.5)
    st.session_state.estado_ultimo_fardo = "normal"
    st.rerun()

st.divider()

# 5. MONITOREO EN TIEMPO REAL
if not st.session_state.tabla_carga.empty:
    total_kg = int(st.session_state.tabla_carga["Peso (Kg)"].sum())
    total_bultos = len(st.session_state.tabla_carga)
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric(label="TOTAL KILOS", value=f"{total_kg:,} Kg")
    with col2:
        st.metric(label="TOTAL BULTOS", value=f"{total_bultos} fardos")
    
    st.divider()
    
    # Registro de Patente
    col_izq, col_der = st.columns([3, 2])
    with col_izq:
        st.write("### Detalle de la Carga Actual")
    with col_der:
        patente = st.text_input("Patente del Camión:", key="patente_camion", placeholder="EJ: AB-CD-12")
    
    # CAMBIO: Se reordenó el set_index y el column_order para la visualización en pantalla
    df_pantalla = st.session_state.tabla_carga[["Ítem", "Folio", "Peso (Kg)", "Producto"]].set_index("Ítem")
    
    st.dataframe(
        df_pantalla, 
        use_container_width=True,
        column_order=["Folio", "Peso (Kg)", "Producto"]
    )
    
    st.divider()

    # Botón WhatsApp Texto
    st.write("### 📤 Reporte de Salida")
    patente_texto = patente.strip().upper() if patente.strip() != "" else "NO REGISTRADA"
    
    mensaje_wsp = f"🚛 *REPORTE DE CARGA - METALUM*\n"
    mensaje_wsp += f"🔹 *Patente:* {patente_texto}\n"
    mensaje_wsp += f"----------------------------------------\n"
    mensaje_wsp += f"`Ítem | Folio | Peso(Kg) | Producto`\n"  # CAMBIO: Ajuste de orden en el texto
    
    for idx, fila in st.session_state.tabla_carga.iterrows():
        mensaje_wsp += f"{int(fila['Ítem'])} | F:{int(fila['Folio'])} | {int(fila['Peso (Kg)'])} Kg | {fila['Producto']}\n"
        
    mensaje_wsp += f"----------------------------------------\n"
    mensaje_wsp += f"📦 *Total Bultos:* {total_bultos}\n"
    mensaje_wsp += f"⚖️ *Peso Total:* {total_kg:,} Kg"
    
    texto_codificado = urllib.parse.quote(mensaje_wsp)
    enlace_whatsapp = f"https://api.whatsapp.com/send?text={texto_codificado}"
    
    st.markdown(f"""
        <div class="boton-wsp">
            <a href="{enlace_whatsapp}" target="_blank">
                <svg xmlns="http://www.w3.org/2000/svg" width="22" height="22" fill="currentColor" viewBox="0 0 16 16">
                    <path d="M13.601 2.326A7.85 7.85 0 0 0 7.994 0C3.627 0 .068 3.558.064 7.926c0 1.399.366 2.76 1.057 3.965L0 16l4.204-1.102a7.9 7.9 0 0 0 3.79.965h.004c4.368 0 7.926-3.558 7.93-7.93A7.9 7.9 0 0 0 13.6 2.326zM7.994 14.521a6.6 6.6 0 0 1-3.356-.92l-.24-.144-2.494.654.666-2.433-.156-.251a6.56 6.56 0 0 1-1.007-3.505c0-3.626 2.957-6.584 6.591-6.584a6.56 6.56 0 0 1 4.66 1.931 6.56 6.56 0 0 1 1.928 4.66c-.004 3.639-2.961 6.592-6.592 6.592m3.69-4.936c-.202-.101-1.202-.594-1.392-.661-.19-.068-.33-.101-.47.101-.14.202-.54.661-.661.8-.12.138-.24.154-.442.053-1.648-.826-2.617-1.963-3.064-2.731-.12-.207-.013-.319.09-.421.093-.092.202-.24.302-.36.101-.12.135-.2.203-.34.067-.137.034-.257-.017-.359-.051-.101-.47-1.136-.645-1.545-.171-.413-.344-.358-.47-.364-.121-.006-.26-.006-.4-.006-.14 0-.368.053-.56.26-.191.207-.73.714-.73 1.74s.747 2.009.85 2.147c.101.137 1.47 2.246 3.562 3.149.497.215.885.342 1.184.438.5.158.955.135 1.314.08.4-.061 1.202-.493 1.37-.967.169-.473.169-.88.119-.967-.051-.088-.18-.137-.383-.238"/>
                </svg>
                COMPARTIR TEXTO POR WSP
            </a>
        </div>
    """, unsafe_allow_html=True)
    
    st.write("") # Espaciador
    
    # CAMBIO: Nueva función para descargar la tabla directamente en formato Excel corporativo
    data_excel = generar_excel(st.session_state.tabla_carga, patente_texto)
    st.download_button(
        label="📊 DESCARGAR EXCEL (Para enviar por Wsp)",
        data=data_excel,
        file_name=f"Reporte_Metalum_{patente_texto}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.divider()
    
    # Sección para corregir errores
    st.write("### 🛠️ Corregir Errores")
    lista_items = [0] + st.session_state.tabla_carga["Ítem"].tolist()
    item_a_borrar = st.selectbox("Selecciona el N° de Ítem que deseas eliminar:", lista_items, index=0)
    
    st.markdown('<div class="boton-borrar">', unsafe_allow_html=True)
    if item_a_borrar == 0:
        texto_boton = "🗑️ SELECCIONA UN ÍTEM"
    else:
        texto_boton = f"🗑️ ELIMINAR ÍTEM N° {item_a_borrar}"
        
    if st.button(texto_boton):
        if item_a_borrar == 0:
            st.error("Debes seleccionar un número de Ítem válido.")
        else:
            st.session_state.tabla_carga = st.session_state.tabla_carga[st.session_state.tabla_carga["Ítem"] != item_a_borrar]
            st.session_state.tabla_carga["Ítem"] = range(1, len(st.session_state.tabla_carga) + 1)
            
            guardar_datos(st.session_state.tabla_carga)
            st.session_state.estado_ultimo_fardo = "normal"
            st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.write("") 
    
    if st.button("⚠️ Reiniciar Todo (Camión Nuevo)"):
        st.session_state.tabla_carga = pd.DataFrame(columns=["Ítem", "Folio", "Peso (Kg)", "Producto"])
        
        if os.path.exists(ARCHIVO_DATOS):
            os.remove(ARCHIVO_DATOS)
            
        st.session_state.estado_ultimo_fardo = "normal"
        st.session_state.ultimo_folio_processed = 0
        st.session_state.folio_intentado = 0
        st.rerun()
else:
    st.info("El contenedor está vacío. Empieza a registrar los fardos arriba.")
