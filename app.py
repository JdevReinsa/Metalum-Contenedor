import streamlit as st
import pandas as pd
import time
import urllib.parse
import os
import io

# Nombre del archivo donde se guardará la información de forma permanente
ARCHIVO_DATOS = "registro_fardos_metalum.csv"

# 1. Configuración de la página
st.set_page_config(
    page_title="Metalum - Carga", 
    page_icon="🚛", 
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Inicializar estados para el tema interactivo y clics
if "tema_oscuro" not in st.session_state:
    st.session_state.tema_oscuro = True  # Inicia por defecto en modo oscuro

if "contador_clicks_industria" not in st.session_state:
    st.session_state.contador_clicks_industria = 0

# --- INYECCIÓN DE CSS DINÁMICO SEGÚN EL MODO SELECCIONADO ---
if st.session_state.tema_oscuro:
    # Estilo Modo Oscuro
    css_tema = """
    html, body, [data-testid="stAppViewContainer"], [data-testid="stHeader"] {
        background-color: #0e1117 !important;
        color: #ffffff !important;
    }
    h1, h2, h3, h4, h5, h6, p, label, .stMarkdown, [data-testid="stMetricLabel"] {
        color: #ffffff !important;
    }
    input, select, div[data-baseweb="select"] {
        background-color: #1f2937 !important;
        color: #ffffff !important;
        border: 1px solid #4b5563 !important;
    }
    """
else:
    # Estilo Modo Claro
    css_tema = """
    html, body, [data-testid="stAppViewContainer"], [data-testid="stHeader"] {
        background-color: #ffffff !important;
        color: #111827 !important;
    }
    h1, h2, h3, h4, h5, h6, p, label, .stMarkdown, [data-testid="stMetricLabel"] {
        color: #111827 !important;
    }
    input, select, div[data-baseweb="select"] {
        background-color: #f3f4f6 !important;
        color: #111827 !important;
        border: 1px solid #d1d5db !important;
    }
    """

# CSS general de la app (Botones, ocultar cabeceras nativas)
st.markdown(f"""
    <style>
    {css_tema}
    
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
    
    /* Estilo general para botones */
    .stButton>button, .stDownloadButton>button {{
        width: 100%;
        height: 55px;
        font-size: 16px;
        border-radius: 8px !important;
        transition: all 0.3s ease;
    }}
    
    .boton-normal>div>button {{ background-color: #1e3a8a !important; color: white !important; border: none !important; }}
    .boton-exito>div>button {{ background-color: #2e7d32 !important; color: white !important; font-weight: bold !important; border: none !important; }}
    .boton-error>div>button {{ background-color: #d32f2f !important; color: white !important; font-weight: bold !important; border: none !important; }}
    .boton-borrar>div>button {{ background-color: #4b5563 !important; color: white !important; height: 55px !important; border: none !important; }}
    
    /* Estilo del link-botón de WhatsApp */
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
    
    /* Estilo del botón nativo de Excel */
    .boton-excel-wsp>div>button {{
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        gap: 10px !important;
        background-color: #107c41 !important;
        color: white !important;
        font-weight: bold !important;
        font-size: 18px !important;
        border: none !important;
    }}
    
    input {{
        height: 45px !important;
        font-size: 16px !important;
    }}
    
    /* Estilo para el botón invisible de la industria */
    .btn-industria {{
        background: none !important;
        border: none !important;
        padding: 0 !important;
        font-size: 32px !important;
        cursor: pointer;
        line-height: 1;
    }}
    </style>
    """, unsafe_allow_html=True)


# --- LÓGICA DE DETECCIÓN DE TRIPLE CLICK EN LA INDUSTRIA ---
col_logo, col_titulo = st.columns([1, 7])

with col_logo:
    # Un botón nativo estilizado como texto plano para capturar los clics en la industria sin romper el diseño
    if st.button("🏭", key="click_secreto_industria", help="Cambiar modo visual"):
        st.session_state.contador_clicks_industria += 1
        if st.session_state.contador_clicks_industria >= 3:
            st.session_state.tema_oscuro = not st.session_state.tema_oscuro
            st.session_state.contador_clicks_industria = 0
            st.rerun()

with col_titulo:
    st.markdown("<h1 style='margin:0; padding:0; line-height:1.1;'>METALUM</h1>", unsafe_allow_html=True)

st.subheader("Registro de Contenedor")
st.divider()

# --- CARGAR DATOS ---
def cargar_datos():
    columnas_correctas = ["Ítem", "Folio", "Peso (Kg)", "Producto"]
    if os.path.exists(ARCHIVO_DATOS):
        try:
            df = pd.read_csv(ARCHIVO_DATOS)
            df["Peso (Kg)"] = df["Peso (Kg)"].astype(int)
            df["Folio"] = df["Folio"].astype(int)
            df = df.reindex(columns=columnas_correctas)
            return df
        except Exception:
            return pd.DataFrame(columns=columnas_correctas)
    return pd.DataFrame(columns=columnas_correctas)

# --- GUARDAR DATOS ---
def guardar_datos(df):
    df.to_csv(ARCHIVO_DATOS, index=False)

# --- GENERAR EXCEL PREMIUM DE RESPALDO Y APTO PARA IMPRESIÓN ---
def generar_excel(df, patente_nom, total_b, total_k):
    output = io.BytesIO()
    df_excel = df.copy()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_excel.to_excel(writer, index=False, sheet_name='Reporte Carga')
        
        workbook  = writer.book
        worksheet = writer.sheets['Reporte Carga']
        
        # Formato de la cabecera (Estilo corporativo Excel, texto blanco, fondo verde oscuro, centrado y con bordes)
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'vcenter',
            'align': 'center',
            'fg_color': '#107C41',
            'font_color': 'white',
            'border': 1
        })
        
        # Formatos para el cuerpo de la tabla (Garantiza que se impriman todas las líneas divisorias de cuadrícula)
        center_format = workbook.add_format({'align': 'center', 'border': 1})
        right_format = workbook.add_format({'align': 'right', 'border': 1, 'num_format': '#,##0'})
        left_format = workbook.add_format({'align': 'left', 'border': 1})
        
        # Formatos especiales para resaltar los Totales al pie de página
        bold_label = workbook.add_format({'bold': True, 'align': 'right', 'font_size': 11})
        bold_value = workbook.add_format({'bold': True, 'align': 'left', 'font_size': 11})
        
        # Aplicar el formato premium a los encabezados
        for col_num, header in enumerate(df_excel.columns):
            worksheet.write(0, col_num, header, header_format)
            
        # Forzar la escritura de datos aplicando alineación y cuadrículas completas
        for row_idx in range(len(df_excel)):
            worksheet.write(row_idx + 1, 0, df_excel.iloc[row_idx, 0], center_format) # Ítem
            worksheet.write(row_idx + 1, 1, df_excel.iloc[row_idx, 1], center_format) # Folio
            worksheet.write(row_idx + 1, 2, df_excel.iloc[row_idx, 2], right_format)  # Peso (Kg)
            worksheet.write(row_idx + 1, 3, df_excel.iloc[row_idx, 3], left_format)   # Producto

        # Auto-ajuste de columnas adaptativo con un margen seguro (+5 de ancho) para evitar los símbolos '###'
        for col_num, col_name in enumerate(df_excel.columns):
            max_len = len(str(col_name))
            for val in df_excel[col_name]:
                max_len = max(max_len, len(str(val)))
            worksheet.set_column(col_num, col_num, max_len + 5)
            
        # Cuadro de resumen final de carga estilizado para la firma o aprobación del jefe
        start_row = len(df_excel) + 2
        worksheet.write(start_row, 1, "Patente Camión:", bold_label)
        worksheet.write(start_row, 2, patente_nom, bold_value)
        worksheet.write(start_row + 1, 1, "Total Bultos:", bold_label)
        worksheet.write(start_row + 1, 2, f"{total_b} fardos", bold_value)
        worksheet.write(start_row + 2, 1, "Peso Total:", bold_label)
        worksheet.write(start_row + 2, 2, f"{total_k:,} Kg", bold_value)
        
    return output.getvalue()

# Inicializar Estados de la tabla
if "tabla_carga" not in st.session_state:
    st.session_state.tabla_carga = cargar_datos()
if "estado_ultimo_fardo" not in st.session_state:
    st.session_state.estado_ultimo_fardo = "normal"
if "ultimo_folio_processed" not in st.session_state:
    st.session_state.ultimo_folio_processed = 0
if "folio_intentado" not in st.session_state:
    st.session_state.folio_intentado = 0
if "form_reset_counter" not in st.session_state:
    st.session_state.form_reset_counter = 0

# 3. FORMULARIO DE CARGA
with st.form(key="formulario_fardo"):
    producto = st.selectbox(
        "Selecciona el Producto:",
        ["UBC", "Perfil", "Tense", "Taint Tabor", "Radiador", "Acero", "Offset"],
        key="producto_seleccionado"
    )
    
    ctr = st.session_state.form_reset_counter
    peso_raw = st.text_input("Peso (Kg):", placeholder="Escribe el peso...", key=f"peso_{ctr}")
    folio_raw = st.text_input("Número de Folio:", placeholder="Escribe el folio...", key=f"folio_{ctr}")
    
    f_proc = st.session_state.ultimo_folio_processed
    f_rep = st.session_state.folio_intentado
    
    if st.session_state.estado_ultimo_fardo == "exito":
        clase_boton = "boton-exito"
        texto_boton = f"✅ ¡FARDO #{f_proc} SUBIDO! (ENTER)"
    elif st.session_state.estado_ultimo_fardo == "error_duplicado":
        clase_boton = "boton-error"
        texto_boton = f"❌ ¡FOLIO #{f_rep} REPETIDO! ✖️"
    elif st.session_state.estado_ultimo_fardo == "error_vacio":
        clase_boton = "boton-error"
        texto_boton = "❌ ERROR: ¡DATOS VACÍOS O INVÁLIDOS! ✖️"
    else:
        clase_boton = "boton-normal"
        texto_boton = "➕ AGREGAR FARDO"

    st.markdown(f'<div class="{clase_boton}">', unsafe_allow_html=True)
    boton_guardar = st.form_submit_button(label=texto_boton)
    st.markdown('</div>', unsafe_allow_html=True)

# 4. LÓGICA DE PROCESAMIENTO
if boton_guardar:
    try:
        peso_val = int(peso_raw.strip()) if peso_raw else 0
        folio_val = int(folio_raw.strip()) if folio_raw else 0
    except ValueError:
        peso_val = 0
        folio_val = 0
        
    st.session_state.folio_intentado = folio_val
    
    if peso_val > 0 and folio_val > 0:
        folios_existentes = st.session_state.tabla_carga["Folio"].astype(int).values
        if folio_val in folios_existentes:
            st.session_state.estado_ultimo_fardo = "error_duplicado"
            st.rerun()
        else:
            st.session_state.ultimo_folio_processed = folio_val
            siguiente_item = st.session_state.tabla_carga["Ítem"].max() + 1 if len(st.session_state.tabla_carga) > 0 else 1
            
            nueva_fila = pd.DataFrame([{
                "Ítem": int(siguiente_item),
                "Folio": folio_val,
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
    
    col_izq, col_der = st.columns([3, 2])
    with col_izq:
        st.write("### Detalle de la Carga Actual")
    with col_der:
        patente = st.text_input("Patente del Camión:", key="patente_camion", placeholder="EJ: AB-CD-12")
    
    mostrar_producto = st.toggle("👁️ Mostrar columna Producto", value=False)
    columnas_visibles = ["Folio", "Peso (Kg)", "Producto"] if mostrar_producto else ["Folio", "Peso (Kg)"]
    
    df_pantalla = st.session_state.tabla_carga[["Ítem"] + columnas_visibles].set_index("Ítem")
    st.dataframe(df_pantalla, use_container_width=True, column_order=columnas_visibles)
    
    st.divider()

    # REPORTES DE SALIDA
    st.write("### 📤 Reporte de Salida")
    patente_texto = patente.strip().upper() if patente.strip() != "" else "NO REGISTRADA"
    
    # Mensaje de WhatsApp Texto estándar
    mensaje_wsp = f"🚛 *REPORTE DE CARGA - METALUM*\n"
    mensaje_wsp += f"🔹 *Patente:* {patente_texto}\n"
    mensaje_wsp += f"----------------------------------------\n"
    mensaje_wsp += f"`Ítem | Folio | Peso(Kg) | Producto`\n"
    for idx, fila in st.session_state.tabla_carga.iterrows():
        mensaje_wsp += f"{int(fila['Ítem'])} | F:{int(fila['Folio'])} | {int(fila['Peso (Kg)'])} Kg | {fila['Producto']}\n"
    mensaje_wsp += f"----------------------------------------\n"
    mensaje_wsp += f"📦 *Total Bultos:* {total_bultos}\n"
    mensaje_wsp += f"⚖️ *Peso Total:* {total_kg:,} Kg"
    
    texto_codificado = urllib.parse.quote(mensaje_wsp)
    enlace_whatsapp = f"https://api.whatsapp.com/send?text={texto_codificado}"
    
    # Botón 1: WhatsApp Texto
    st.markdown(f"""
        <div class="boton-wsp">
            <a href="{enlace_whatsapp}" target="_blank">
                <svg xmlns="http://www.w3.org/2000/svg" width="22" height="22" fill="currentColor" viewBox="0 0 16 16">
                    <path d="M13.601 2.326A7.85 7.85 0 0 0 7.994 0C3.627 0 .068 3.558.064 7.926c0 1.399.366 2.76 1.057 3.965L0 16l4.204-1.102a7.9 7.9 0 0 0 3.79.965h.004c4.368 0 7.926-3.558 7.93-7.93A7.9 7.9 0 0 0 13.6 2.326zM7.994 14.521a6.6 6.6 0 0 1-3.356-.92l-.24-.144-2.494.654.666-2.433-.156-.251a6.56 6.56 0 0 1-1.007-3.505c0-3.626 2.957-6.584 6.591-6.584a6.56 6.56 0 0 1 4.66 1.931 6.56 6.56 0 0 1 1.928 4.66c-.004 3.639-2.961 6.592-6.592 6.592m3.69-4.936c-.202-.101-1.202-.594-1.392-.661-.19-.068-.33-.101-.47.101-.14.202-.54.661-.661.8-.12.138-.24.154-.442.053-1.648-.826-2.617-1.963-3.064-2.731-.12-.207-.013-.319.09-.421.093-.092.202-.24.302-.36.101-.12.135-.2.203-.34.067-.137.034-.257-.017-.359-.051-.101-.47-1.136-.645-1.545-.171-.413-.344-.358-.47-.364-.121-.006-.26-.006-.4-.006-.14 0-.368.053-.56.26-.191.207-.73.714-.73 1.74s.747 2.009.85 2.147c.101.137 1.47 2.246 3.562 3.149.497.215.885.342 1.184.438.5.158.955.135 1.314.08.4-.061 1.202-.493 1.37-.967.169-.473.169-.88.119-.967-.051-.088-.18-.137-.383-.238"/>
                </svg>
                COMPARTIR
            </a>
        </div>
    """, unsafe_allow_html=True)
    
    st.write("") 
    
    # Botón 2: Excel Premium con Descarga Directa Ultra Estable y Logotipo de Libro de Excel
    data_excel = generar_excel(st.session_state.tabla_carga, patente_texto, total_bultos, total_kg)
    st.markdown('<div class="boton-excel-wsp">', unsafe_allow_html=True)
    st.download_button(
        label="📊 COMPARTIR",  
        data=data_excel,
        file_name=f"Reporte_Metalum_{patente_texto}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.divider()
    
    # Sección para corregir errores
    st.write("### 🛠️ Corregir Errores")
    item_a_borrar_raw = st.text_input(
        "Digita el N° de Ítem que deseas eliminar:", 
        placeholder="Escribe el número de ítem aquí..."
    )
    
    try:
        item_a_borrar = int(item_a_borrar_raw.strip()) if item_a_borrar_raw else 0
    except ValueError:
        item_a_borrar = 0
        
    st.markdown('<div class="boton-borrar">', unsafe_allow_html=True)
    if item_a_borrar == 0:
        texto_boton = "🗑️ INGRESA UN ÍTEM"
    else:
        texto_boton = f"🗑️ ELIMINAR ÍTEM N° {item_a_borrar}"
        
    if st.button(texto_boton):
        if item_a_borrar == 0:
            st.error("Debes ingresar un número de Ítem válido.")
        else:
            items_existentes = st.session_state.tabla_carga["Ítem"].tolist()
            if item_a_borrar not in items_existentes:
                st.error(f"El Ítem N° {item_a_borrar} no existe en la carga actual.")
            else:
                st.session_state.tabla_carga = st.session_state.tabla_carga[st.session_state.tabla_carga["Ítem"] != item_a_borrar]
                st.session_state.tabla_carga["Ítem"] = range(1, len(st.session_state.tabla_carga) + 1)
                
                guardar_datos(st.session_state.tabla_carga)
                st.success(f"¡Ítem N° {item_a_borrar} eliminado con éxito!")
                time.sleep(1)
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
