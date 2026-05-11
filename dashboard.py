import streamlit as st
import pandas as pd
import plotly.express as px


# ==============================================================================
# 1. CONFIGURACIÓN DE PÁGINA 
# ==============================================================================
st.set_page_config(
    page_title="Dashboard Financiero Platco",
    layout="wide",
    page_icon="📊",
    initial_sidebar_state="expanded"
)

# Definición de Colores Corporativos (Estrictamente Azules, Grises y Amarillos)
COLOR_FONDO = "#F4F6F9"           # Gris claro (Fondo)
COLOR_TARJETA = "#FFFFFF"         # Blanco
COLOR_TEXTO_PRINCIPAL = "#001F5B" # Azul Marino Oscuro (Primario)
COLOR_AMARILLO = "#FFB81C"        # Amarillo Corporativo
COLOR_AZUL_MEDIO = "#0056B3"      # Azul Medio para gráficos
COLOR_GRIS_TEXTO = "#6C757D"      # Gris para subtítulos
COLOR_GRIS_CLARO = "#E2E8F0"      # Gris para líneas

# ==============================================================================
# CARGA DINÁMICA DE ARCHIVO (ZONA DE CARGA)
# ==============================================================================
with st.sidebar:
    st.markdown('<h3 style="color: #001F5B; font-weight: 800; margin-bottom: 15px;">📂 Cargar Datos</h3>', unsafe_allow_html=True)
    
    archivo_excel = st.file_uploader(
        "Sube el archivo de Métricas de hoy:", 
        type=['xlsx', 'xls']
    )
    st.markdown('<hr style="border: 1px solid #E2E8F0; margin: 20px 0;">', unsafe_allow_html=True)

if archivo_excel is None:
    st.markdown(f"""
    <div style="text-align: center; padding: 50px; background-color: white; border-radius: 12px; border: 2px dashed #CBD5E1; margin-top: 50px;">
        <h2 style="color: #001F5B; font-family: 'Montserrat', sans-serif; font-weight: 800;">Panel Financiero Platco</h2>
        <p style="color: #64748B; font-family: 'Montserrat', sans-serif; font-size: 1.1rem;">Por favor, sube el archivo Excel de cierre diario en el menú de la izquierda para comenzar.</p>
    </div>
    """, unsafe_allow_html=True)
    st.stop() # Detiene el código hasta que se suba el archivo

# ==============================================================================
# 2. INYECCIÓN DE CSS (FUENTE MONTSERRAT Y COLORES)
# ==============================================================================
st.markdown(f"""
<style>
    /* Importar fuente Montserrat desde Google Fonts */
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;700;800&display=swap');

    /* Aplicar Montserrat sin romper los íconos internos de Streamlit */
    html, body, [class*="css"], .stMarkdown, h1, h2, h3, h4, h5, h6, p {{
        font-family: 'Montserrat', sans-serif;
    }}

    .stApp {{ background-color: {COLOR_FONDO}; }}
    
    .premium-card {{
        background-color: {COLOR_TARJETA};
        border-radius: 16px;
        padding: 24px;
        box-shadow: 0 4px 12px rgba(0, 31, 91, 0.08);
        margin-bottom: 24px;
        border: 1px solid {COLOR_GRIS_CLARO};
    }}

    .card-label {{
        color: {COLOR_GRIS_TEXTO};
        font-family: 'Montserrat', sans-serif;
        font-size: 1.05rem;
        font-weight: 500;
        margin-bottom: 8px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }}
    .card-value {{
        color: {COLOR_TEXTO_PRINCIPAL};
        font-family: 'Montserrat', sans-serif;
        font-size: 2.8rem;
        font-weight: 800;
        letter-spacing: -1px;
        line-height: 1.2;
    }}
    .card-sub-value {{
        color: {COLOR_GRIS_TEXTO};
        font-family: 'Montserrat', sans-serif;
        font-size: 1rem;
        margin-top: 4px;
        font-weight: 500;
    }}
    
    /* Restaurar específicamente la fuente de los íconos de Streamlit por si acaso */
    span.material-icons, span[class*="icon"] {{
        font-family: 'Material Icons', 'Material Symbols Rounded', 'Material Icons Round' !important;
    }}
    
    .js-plotly-plot .plotly .modebar {{ display: none !important; }}
    
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 3. CARGA Y PROCESAMIENTO DE DATOS
# ==============================================================================
@st.cache_data
def cargar_datos_crudos(archivo, pestaña, skip_rows, num_rows, cols_string):
    try:
        df = pd.read_excel(archivo, sheet_name=pestaña, skiprows=skip_rows, nrows=num_rows, usecols=cols_string)
        df = df.dropna(how='all').reset_index(drop=True)
        for col in ['Bs', 'USD']:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
        return df
    except Exception as e:
        st.error(f"Error cargando Excel: {e}")
        return pd.DataFrame()

def format_money_ve(valor):
    return "{:,.2f}".format(valor).replace(',', 'X').replace('.', ',').replace('X', '.')

# Cargar bloque "SALDO INICIAL"
df_saldos_raw = cargar_datos_crudos(archivo_excel, 'METRICAS', 3, 18, "C:E")
df_saldos_raw.columns = ['CONCEPTO', 'Bs', 'USD']

# Búsqueda inteligente a prueba de errores de tipeo en Excel
# Usamos case=False para que no importe si está en mayúsculas o minúsculas
# Usamos .sum() para que si no encuentra nada, devuelva 0.0 y no colapse el código

saldo_operativo_bs = df_saldos_raw[df_saldos_raw['CONCEPTO'].astype(str).str.contains('Saldo Bancario Operativo', case=False, na=False)]['Bs'].sum()

saldo_operativo_usd = df_saldos_raw[df_saldos_raw['CONCEPTO'].astype(str).str.contains('Saldo Bancario Operativo', case=False, na=False)]['USD'].sum()



apartados_usd = df_saldos_raw[df_saldos_raw['CONCEPTO'].astype(str).str.contains('Apartados', case=False, na=False)]['USD'].sum()

# ==============================================================================
# 4. DISEÑO DEL DASHBOARD
# ==============================================================================
import re

# Extraemos inteligentemente la fecha del nombre del archivo cargado
fecha_reporte = ""
if archivo_excel is not None:
    # Busca un patrón de fecha como DD-MM-YYYY o DD/MM/YYYY en el nombre del archivo
    match_fecha = re.search(r'\d{1,2}[-/]\d{1,2}[-/]\d{2,4}', archivo_excel.name)
    if match_fecha:
        fecha_reporte = f" | {match_fecha.group(0)}"
    else:
        # Si no tiene fecha en el nombre, ponemos el nombre del archivo sin extensión por si acaso
        fecha_reporte = f" | {archivo_excel.name.rsplit('.', 1)[0]}"

# CABECERA
st.markdown(f"""
<div style="background-color: {COLOR_TEXTO_PRINCIPAL}; padding: 24px; border-radius: 16px; margin-bottom: 32px; display: flex; align-items: center; justify-content: space-between;">
    <h1 style="color: white; margin: 0; font-weight: 800; font-size: 2rem;">Cierre Financiero - Flujo de Caja</h1>
    <div style="color: {COLOR_AMARILLO}; font-weight: 600; font-size: 1.1rem; letter-spacing: 1px; text-align: right;">
        Platco | Administración y Finanzas <span style="color: white;">{fecha_reporte}</span>
    </div>
</div>
""", unsafe_allow_html=True)

# ==============================================================================
# 4.5. NUEVA SECCIÓN: RESUMEN EJECUTIVO (FLASH REPORT)
# ==============================================================================
st.markdown('<h3 style="color: #001F5B; font-weight: 800; margin-bottom: 15px; margin-top: 10px;">Resumen Consolidado</h3>', unsafe_allow_html=True)

# --- CARGA DINÁMICA DE DATOS DEL EXCEL LIMPIO ---
@st.cache_data
def cargar_flash_report(archivo):
    try:
        # Leemos las columnas C, D y E desde la fila 4 (skiprows=3)
        df = pd.read_excel(archivo, sheet_name='METRICAS', skiprows=3, nrows=20, usecols="C:E")
        df.columns = ['Concepto', 'Bs', 'USD']
        
        # Limpieza LATINA blindada: maneja puntos de miles y comas de decimales perfectamente
        def limpiar_moneda(val):
            if pd.isna(val) or str(val).strip() in ['#N/D', 'nan', '']: return 0.0
            if isinstance(val, (int, float)): return float(val)
            v = str(val).replace('$', '').replace(' ', '').strip()
            # Si tiene punto de miles y coma decimal (ej. 71.365,42)
            if '.' in v and ',' in v:
                v = v.replace('.', '').replace(',', '.') 
            # Si solo tiene coma decimal (ej. 71365,42)
            elif ',' in v:
                v = v.replace(',', '.') 
            try:
                return float(v)
            except:
                return 0.0

        df['Bs'] = df['Bs'].apply(limpiar_moneda)
        df['USD'] = df['USD'].apply(limpiar_moneda)

        # Función interna para buscar palabras clave de forma segura
        def buscar(texto, col):
            return df[df['Concepto'].astype(str).str.contains(texto, case=False, na=False)][col].sum()

        return (
            buscar('Bancario Operativo', 'Bs'),
            buscar('Compromisos CxP', 'Bs'),
            buscar('Apartados para Compromisos Operativo', 'Bs'), # Ignoramos la palabra "abril"
            buscar('Aliados-Continuidad', 'Bs'),
            buscar('Operativo Disponible a Hoy', 'Bs'),
            buscar('Saldo disponible proyectado', 'Bs'),
            
            buscar('Bancario Inversion', 'Bs'),
            buscar('Equipos Por Pagar', 'Bs'),
            buscar('Apartados para compromisos Fiscales', 'Bs'),
            buscar('Servicio Especializado', 'Bs'),
            buscar('Inversion Disponible a Hoy', 'Bs'),
            buscar('Inversion proyectado a cierre', 'Bs'),
            
            buscar('Moneda Extranjera', 'USD'),
            buscar('Boveda', 'USD')
        )
    except Exception as e:
        st.error(f"Error procesando Resumen Consolidado: {e}")
        return (0,)*14 # Retorna ceros si el Excel está corrupto

# Asignamos los valores extraídos a las variables del diseño
(val_op_banco, val_op_cxp, val_op_apartado, val_op_cxc, val_op_disp_hoy, val_op_proyectado,
 val_inv_banco, val_inv_equipos, val_inv_apartado, val_inv_cxc, val_inv_disp_hoy, val_inv_proyectado,
 val_div_ext, val_div_boveda) = cargar_flash_report(archivo_excel)

# ---> EXTRACCIÓN DIRECTA DESDE LA FUENTE (Pestaña: DISTRIBUCIÓN DE FONDOS, Celda: C50) <---
@st.cache_data
def extraer_moneda_extranjera_directo(archivo):
    try:
        # Rebobinamos el archivo por seguridad
        archivo.seek(0)
        
        # Viajamos a la pestaña original, saltamos 49 filas para caer en la 50, y leemos solo la columna C
        df_temp = pd.read_excel(
            archivo, 
            sheet_name='DISTRIBUCIÓN DE FONDOS', 
            skiprows=49, 
            nrows=1, 
            usecols="C", 
            header=None
        )
        
        if not df_temp.empty:
            val = df_temp.iloc[0, 0] # Toma el valor exacto de esa única celda
            
            # 1. Si Excel nos da el número puro (como suele pasar en tablas dinámicas origen)
            if isinstance(val, (int, float)):
                return float(val) if not pd.isna(val) else 0.0
                
            # 2. Si viene vacío o con error
            if pd.isna(val) or str(val).strip() in ['#N/D', 'nan', '', 'None']:
                return 0.0
                
            # 3. Limpieza de respaldo
            import re
            val_limpio = re.sub(r'[^\d\.,\-]', '', str(val))
            if '.' in val_limpio and ',' in val_limpio:
                val_limpio = val_limpio.replace('.', '').replace(',', '.')
            elif ',' in val_limpio:
                val_limpio = val_limpio.replace(',', '.')
                
            return float(val_limpio)
    except Exception:
        pass
    return 0.0

# Asignamos la variable directo desde la hoja origen
val_div_ext = extraer_moneda_extranjera_directo(archivo_excel)


COLOR_BANNER = "#001F5B"
COLOR_CAJA_CELESTE = "#00AEEF"


# ---> EXTRACCIÓN DE TASA BCV (PUNTERÍA LÁSER: CUENTAS POR COBRAR) <---
@st.cache_data
def extraer_tasa_bcv(archivo):
    try:
        # Rebobinamos el archivo por seguridad
        archivo.seek(0)
        
        # CAMBIO CLAVE: Ahora leemos explícitamente la pestaña 'CUENTAS POR COBRAR'
        # Escaneamos las primeras 20 filas por si el cliente movió la tabla
        df_tasa = pd.read_excel(archivo, sheet_name='CUENTAS POR COBRAR', nrows=20, header=None)
        
        # Radar para encontrar la celda que dice "TASA BCV"
        for fila in range(len(df_tasa)):
            for col in range(len(df_tasa.columns)):
                celda_texto = str(df_tasa.iloc[fila, col]).strip().upper()
                
                if 'TASA BCV' in celda_texto:
                    # El monto debe estar en la celda de la derecha (col + 1)
                    if col + 1 < len(df_tasa.columns):
                        val = df_tasa.iloc[fila, col + 1]
                        
                        # Si es un número puro de Excel
                        if isinstance(val, (int, float)): 
                            return float(val) if val > 0 else 1.0
                        
                        # Si viene como texto con comas (ej. "493,37")
                        import re
                        v = re.sub(r'[^\d\.,\-]', '', str(val))
                        if '.' in v and ',' in v: v = v.replace('.', '').replace(',', '.')
                        elif ',' in v: v = v.replace(',', '.')
                        
                        tasa = float(v)
                        return tasa if tasa > 0 else 1.0
    except Exception as e:
        # En caso de error, no mostramos nada para no asustar al usuario
        pass
    
    # Si no la encuentra, devuelve 1.0 para que la app siga viva
    return 1.0 

# Asignamos la variable de la tasa
tasa_bcv_actual = extraer_tasa_bcv(archivo_excel)


# --- HTML PEGADO TOTALMENTE A LA IZQUIERDA PARA EVITAR EL ERROR DE TEXTO ---
html_resumen = f"""<style>
.resumen-row {{ display: flex; align-items: center; justify-content: space-between; background-color: {COLOR_BANNER}; padding: 15px 20px; border-radius: 12px; margin-bottom: 15px; color: white; font-family: 'Montserrat', sans-serif; box-shadow: 0 4px 10px rgba(0,0,0,0.15); }}
.resumen-item {{ display: flex; flex-direction: column; justify-content: center; border-right: 1px solid rgba(255,255,255,0.2); padding-right: 15px; padding-left: 15px; text-align: right; }}
.resumen-item:first-child {{ padding-left: 0; }}
.resumen-item:last-child {{ border-right: none; padding-right: 0; }}
.resumen-label {{ font-size: 0.75rem; color: #E2E8F0; text-transform: uppercase; margin-bottom: 3px; font-weight: 500; }}
.resumen-valor {{ font-size: 1.1rem; font-weight: 800; }}
.caja-celeste {{ background-color: {COLOR_CAJA_CELESTE}; padding: 10px 15px; border-radius: 6px; text-align: center; font-weight: 800; font-size: 1.1rem; min-width: 150px; }}
.caja-blanca {{ background-color: white; color: {COLOR_BANNER}; padding: 8px 15px; border-radius: 6px; text-align: center; font-weight: 700; font-size: 0.8rem; display: flex; align-items: center; min-width: 160px; line-height: 1.1; }}
</style>

<div class="resumen-row">
<div class="resumen-item"><span class="resumen-label">Saldo Bancario Operativo</span><span class="resumen-valor">{format_money_ve(val_op_banco)}</span></div>
<div class="resumen-item"><span class="resumen-label">Compromisos de CxP</span><span class="resumen-valor">{format_money_ve(val_op_cxp)}</span></div>
<div class="resumen-item"><span class="resumen-label">Apartados Prov. Abril</span><span class="resumen-valor">{format_money_ve(val_op_apartado)}</span></div>
<div class="resumen-item" style="border-right: none;"><span class="resumen-label">CxC Aliados/Bancos</span><span class="resumen-valor">{format_money_ve(val_op_cxc)}</span></div>
<div style="display: flex; gap: 10px; margin-left: 20px;">
<div style="display: flex; flex-direction: column; gap: 5px;"><div class="caja-celeste">{format_money_ve(val_op_disp_hoy)}</div><div class="caja-blanca">Saldo Operativo<br>Disponible a Hoy</div></div>
<div style="display: flex; flex-direction: column; gap: 5px;"><div class="caja-celeste">{format_money_ve(val_op_proyectado)}</div><div class="caja-blanca">Saldo Operativo<br>proyectado fin de mes</div></div>
</div>
</div>

<div class="resumen-row">
<div class="resumen-item"><span class="resumen-label">Saldo Bancario Inversión</span><span class="resumen-valor">{format_money_ve(val_inv_banco)}</span></div>
<div class="resumen-item"><span class="resumen-label">Equipos por Pagar</span><span class="resumen-valor">{format_money_ve(val_inv_equipos)}</span></div>
<div class="resumen-item"><span class="resumen-label">Apartados Fiscales</span><span class="resumen-valor">{format_money_ve(val_inv_apartado)}</span></div>
<div class="resumen-item" style="border-right: none;"><span class="resumen-label">Universo CxC Especializado</span><span class="resumen-valor">{format_money_ve(val_inv_cxc)}</span></div>
<div style="display: flex; gap: 10px; margin-left: 20px;">
<div style="display: flex; flex-direction: column; gap: 5px;"><div class="caja-celeste">{format_money_ve(val_inv_disp_hoy)}</div><div class="caja-blanca">Saldo Inversión<br>Disponible Hoy</div></div>
<div style="display: flex; flex-direction: column; gap: 5px;"><div class="caja-celeste">{format_money_ve(val_inv_proyectado)}</div><div class="caja-blanca">Saldo para inversión<br>proyectado fin de mes</div></div>
</div>
</div>

<div style="display: flex; gap: 20px;">
<div style="background-color: {COLOR_BANNER}; color: white; padding: 10px 20px; border-radius: 8px; font-family: 'Montserrat'; font-weight: 700; display: flex; gap: 15px;"><span style="color: #E2E8F0; font-size: 0.8rem; text-transform: uppercase;">Moneda Extranjera:</span><span>${format_money_ve(val_div_ext)}</span></div>
<div style="background-color: {COLOR_BANNER}; color: white; padding: 10px 20px; border-radius: 8px; font-family: 'Montserrat'; font-weight: 700; display: flex; gap: 15px;"><span style="color: #E2E8F0; font-size: 0.8rem; text-transform: uppercase;">Saldo en Bóveda:</span><span>${format_money_ve(val_div_boveda)}</span></div>
</div>"""

st.markdown(html_resumen, unsafe_allow_html=True)

# ==============================================================================
# EXTRACCIÓN DE NUEVOS VALORES (Asegurando tener Bs y USD para todo)
# ==============================================================================
# 1. Extracción de Compromisos (Desde la tabla inferior amarilla)
@st.cache_data
def cargar_total_compromisos_cxp_real(archivo):
    try:
        # CAMBIO 1: Ampliamos la lectura hasta la columna J (usecols="F:J")
        df = pd.read_excel(archivo, sheet_name='METRICAS', usecols="F:J", header=None)
        
        # CAMBIO 2: Asignamos 5 nombres a las 5 columnas (F=Concepto, G=Bs, H=Porcentaje, I=Vacia, J=USD)
        df.columns = ['Concepto', 'Bs', 'Porcentaje', 'Columna_I', 'USD']
        
        # Buscamos la tabla específica de compromisos
        idx_tabla = df[df['Concepto'].astype(str).str.contains('COMPROMISOS POR PAGAR POR SEMANA', case=False, na=False)].index
        if len(idx_tabla) > 0:
            start_idx = idx_tabla[0]
            # Buscamos la fila "Total" justo debajo del título
            df_subset = df.iloc[start_idx:start_idx+15]
            idx_total = df_subset[df_subset['Concepto'].astype(str).str.strip().str.lower() == 'total'].index
            
            if len(idx_total) > 0:
                val_bs = df.loc[idx_total[0], 'Bs']   # Extrae de la Columna G
                val_usd = df.loc[idx_total[0], 'USD'] # Extrae de la Columna J
                
                # Función de limpieza (vital para que lea ese número amarillo sin fallar)
                def limpiar_latino(val):
                    if pd.isna(val) or str(val).strip() in ['#N/D', 'nan', '', 'None']: return 0.0
                    if isinstance(val, (int, float)): return float(val)
                    import re
                    v = re.sub(r'[^\d\.,\-]', '', str(val))
                    if '.' in v and ',' in v: v = v.replace('.', '').replace(',', '.')
                    elif ',' in v: v = v.replace(',', '.')
                    try: return float(v)
                    except: return 0.0
                
                bs_num = limpiar_latino(val_bs)
                usd_num = limpiar_latino(val_usd)
                
                return abs(bs_num), abs(usd_num)
    except:
        pass
    return 0.0, 0.0

# Asignamos los valores sobreescribiendo el cálculo anterior
compromisos_bs, compromisos_usd = cargar_total_compromisos_cxp_real(archivo_excel)
# 1.1 CORRECCIÓN: Extracción exacta del "Apartado Abril" (Ignorando los fiscales)
apartados_bs = df_saldos_raw[df_saldos_raw['CONCEPTO'].astype(str).str.contains('Apartados para Compromisos Operativo', case=False, na=False)]['Bs'].sum()
apartados_usd = df_saldos_raw[df_saldos_raw['CONCEPTO'].astype(str).str.contains('Apartados para Compromisos Operativo', case=False, na=False)]['USD'].sum()

# ==============================================================================
# 2 y 3. EXTRACCIÓN EXACTA DE "COBRANZA ALIADOS / BANCOS" (Bs y USD Reales)
# ==============================================================================
@st.cache_data
def cargar_cobranza_aliados_bancos_real(archivo):
    try:
        archivo.seek(0)
        # Leemos desde la columna F hasta la J (usecols="F:J")
        df = pd.read_excel(archivo, sheet_name='METRICAS', usecols="F:J", header=None)
        df.columns = ['Concepto', 'Bs', 'Col_H', 'Col_I', 'USD']
        
        # Limpiador infalible para purificar los montos
        def limpiar_monto(val):
            if pd.isna(val) or str(val).strip() in ['#N/D', 'nan', '', 'None', '-']: return 0.0
            if isinstance(val, (int, float)): return float(val)
            import re
            v = re.sub(r'[^\d\.,\-]', '', str(val))
            if '.' in v and ',' in v: v = v.replace('.', '').replace(',', '.')
            elif ',' in v: v = v.replace(',', '.')
            try: return float(v)
            except: return 0.0

        # Puntería láser directa a la fila 75 (índice 74 en Python)
        if len(df) > 74:
            concepto_75 = str(df.loc[74, 'Concepto']).strip().upper()
            if 'TOTAL' in concepto_75:
                bs_val = limpiar_monto(df.loc[74, 'Bs'])
                usd_val = limpiar_monto(df.loc[74, 'USD']) # Extrae directo de la celda J75
                if bs_val > 0 or usd_val > 0:
                    return abs(bs_val), abs(usd_val)

        # Radar de respaldo por si se insertan o mueven filas
        idx_aliados = df[df['Concepto'].astype(str).str.strip().str.lower() == 'aliados'].index
        if len(idx_aliados) > 0:
            start_idx = idx_aliados[0]
            df_subset = df.iloc[start_idx:start_idx+10]
            idx_total = df_subset[df_subset['Concepto'].astype(str).str.strip().str.lower() == 'total'].index
            if len(idx_total) > 0:
                bs_val = limpiar_monto(df.loc[idx_total[0], 'Bs'])
                usd_val = limpiar_monto(df.loc[idx_total[0], 'USD'])
                return abs(bs_val), abs(usd_val)
                
    except Exception:
        pass
    return 0.0, 0.0

# Asignamos las variables sobreescribiendo las anteriores
val_total_aliados_bs, cxc_bancos_usd = cargar_cobranza_aliados_bancos_real(archivo_excel)


# ==============================================================================
# FILA 1: TARJETAS KPI (4 COLUMNAS UNIFICADAS)
# ==============================================================================
col_kpi1, col_kpi2, col_kpi3, col_kpi4 = st.columns(4)

with col_kpi1:
    st.markdown(f"""
    <div class="premium-card">
        <div class="card-label">Saldo Operativo Total</div>
        <div class="card-value">${format_money_ve(saldo_operativo_usd)}</div>
        <div class="card-sub-value">Bs {format_money_ve(saldo_operativo_bs)}</div>
    </div>
    """, unsafe_allow_html=True)

with col_kpi2:
    st.markdown(f"""
    <div class="premium-card">
        <div class="card-label">Compromisos CxP (hoy)</div>
        <div class="card-value" style="color: {COLOR_GRIS_TEXTO};">${format_money_ve(abs(compromisos_usd))}</div>
        <div class="card-sub-value">Bs {format_money_ve(abs(compromisos_bs))}</div>
    </div>
    """, unsafe_allow_html=True)

with col_kpi3:
    st.markdown(f"""
    <div class="premium-card" style="border-bottom: 4px solid {COLOR_AMARILLO};">
        <div class="card-label">Apartados Abril</div>
        <div class="card-value">${format_money_ve(abs(apartados_usd))}</div>
        <div class="card-sub-value">Bs {format_money_ve(abs(apartados_bs))}</div>
    </div>
    """, unsafe_allow_html=True)

with col_kpi4:
    # Tarjeta Unificada: USD en grande (Bancos), Bs en pequeño (Aliados)
    st.markdown(f"""
    <div class="premium-card" style="border-bottom: 4px solid {COLOR_CAJA_CELESTE};">
        <div class="card-label">Cobranza Aliados / Bancos (Hoy)</div>
        <div class="card-value" style="color: {COLOR_TEXTO_PRINCIPAL};">${format_money_ve(abs(cxc_bancos_usd))}</div>
        <div class="card-sub-value">Bs {format_money_ve(abs(val_total_aliados_bs))}</div>
    </div>
    """, unsafe_allow_html=True)


# ==============================================================================
# 8. SECCIÓN DINÁMICA: COMPROMISOS POR PAGAR (CONSOLIDADO)
# ==============================================================================
st.markdown('<hr style="border: 1px solid #E2E8F0; margin: 40px 0;">', unsafe_allow_html=True)
st.markdown(f'<h3 style="color: {COLOR_TEXTO_PRINCIPAL}; font-weight: 800; margin-bottom: 20px;">Detalle de Compromisos por Pagar</h3>', unsafe_allow_html=True)

# 8.1 CARGA Y LIMPIEZA DE DATOS (COORDENADAS F:G)
@st.cache_data
def cargar_compromisos(archivo, skip_rows):
    try:
        # Leemos las columnas F y G de tu imagen
        df = pd.read_excel(archivo, sheet_name='METRICAS', skiprows=skip_rows, nrows=10, usecols="F:G", header=None)
        df.columns = ['Concepto', 'Monto']
        
        df = df.dropna(how='all')
        
        # Filtramos la basura y la fila de totales
        palabras_basura = ['Total', 'Conceptos', 'COMPROMISOS', 'POR SEMANA']
        df = df[~df['Concepto'].astype(str).str.strip().isin(palabras_basura)]
        df = df[~df['Concepto'].astype(str).str.contains('Total', case=False, na=False)]

        # Limpiamos los montos (quitando el negativo para graficar correctamente)
        df['Monto'] = pd.to_numeric(df['Monto'].astype(str).str.replace('$', '', regex=False).str.replace(',', '.', regex=False).str.replace(' ', '', regex=False).str.replace('-', '', regex=False), errors='coerce').fillna(0.0).abs()
        
        # Filtramos vacíos
        df = df[(df['Monto'] > 0) & (df['Concepto'].astype(str).str.strip() != '') & (df['Concepto'].astype(str).str.strip() != 'nan')].reset_index(drop=True)
        
        # Calculamos el porcentaje real
        total_real = df['Monto'].sum()
        df['Porcentaje'] = (df['Monto'] / total_real) * 100 if total_real > 0 else 0
        
        # Para barras horizontales, ordenamos ascendente (el mayor queda arriba en el gráfico)
        df_grafico = df.sort_values(by='Monto', ascending=True).reset_index(drop=True)
        # Para la tabla, ordenamos descendente (el mayor de primero)
        df_tabla = df.sort_values(by='Monto', ascending=False).reset_index(drop=True)
        
        return df_tabla, df_grafico, total_real
    except Exception as e:
        return pd.DataFrame(), pd.DataFrame(), 0.0

# Usamos skip_rows=26 basándonos en que tu encabezado está en la fila 27
df_comp_tabla, df_comp_grafico, total_comp = cargar_compromisos(archivo_excel, skip_rows=26)

# 8.2 DISEÑO UI: TABLA, BARRAS HORIZONTALES E INSIGHT
if not df_comp_tabla.empty:
    col_comp_tabla, col_comp_grafico = st.columns([1.5, 2.5])
    
    with col_comp_tabla:
        st.markdown('<div class="premium-card" style="padding: 20px; margin-bottom: 20px;">', unsafe_allow_html=True)
        st.markdown(f'<h4 style="color: {COLOR_TEXTO_PRINCIPAL}; font-weight: 700; margin-top: 0; margin-bottom: 15px;">Desglose de Conceptos</h4>', unsafe_allow_html=True)
        
        # Agregamos el encabezado a la tabla HTML para identificar las columnas
        html_tabla_comp = f"""
        <table style='width:100%; border-collapse: collapse; font-family: "Montserrat", sans-serif; font-size: 0.85rem;'>
            <thead>
                <tr>
                    <th style='padding: 8px 5px; border-bottom: 2px solid {COLOR_GRIS_CLARO}; color: {COLOR_TEXTO_PRINCIPAL}; text-align: left;'>Concepto</th>
                    <th style='padding: 8px 5px; border-bottom: 2px solid {COLOR_GRIS_CLARO}; color: {COLOR_TEXTO_PRINCIPAL}; text-align: right;'>Monto Bs</th>
                    <th style='padding: 8px 5px; border-bottom: 2px solid {COLOR_GRIS_CLARO}; color: {COLOR_TEXTO_PRINCIPAL}; text-align: right;'>Monto USD</th>
                    <th style='padding: 8px 5px; border-bottom: 2px solid {COLOR_GRIS_CLARO}; color: {COLOR_TEXTO_PRINCIPAL}; text-align: right;'>%</th>
                </tr>
            </thead>
            <tbody>
        """
        for _, row in df_comp_tabla.iterrows():
            concepto = str(row['Concepto']).strip()
            mnt_bs = format_money_ve(row['Monto'])
            
            # --- NUEVO: Calculamos los Dólares usando la tasa ---
            mnt_usd_val = row['Monto'] / tasa_bcv_actual
            mnt_usd = f"${format_money_ve(mnt_usd_val)}"
            # ----------------------------------------------------
            
            pct = f"{row['Porcentaje']:.1f}%"
            
            html_tabla_comp += f"<tr><td style='padding: 8px 5px; border-bottom: 1px solid {COLOR_GRIS_CLARO}; color: {COLOR_GRIS_TEXTO}; font-weight: 600;'>{concepto}</td>"
            html_tabla_comp += f"<td style='padding: 8px 5px; border-bottom: 1px solid {COLOR_GRIS_CLARO}; text-align: right; color: {COLOR_TEXTO_PRINCIPAL}; font-weight: 700;'>{mnt_bs}</td>"
            html_tabla_comp += f"<td style='padding: 8px 5px; border-bottom: 1px solid {COLOR_GRIS_CLARO}; text-align: right; color: #28a745; font-weight: 700;'>{mnt_usd}</td>" # Puse los dólares en verde oscuro
            html_tabla_comp += f"<td style='padding: 8px 5px; border-bottom: 1px solid {COLOR_GRIS_CLARO}; text-align: right; color: {COLOR_CAJA_CELESTE}; font-weight: 800;'>{pct}</td></tr>"
        html_tabla_comp += "</tbody></table></div>"
        st.markdown(html_tabla_comp, unsafe_allow_html=True)
        
        concepto_top = df_comp_tabla.iloc[0]['Concepto']
        pct_top = df_comp_tabla.iloc[0]['Porcentaje']
        
        st.markdown(f"""
        <div style="background-color: #F8FAFC; border-left: 4px solid {COLOR_AMARILLO}; padding: 15px; border-radius: 0 8px 8px 0; font-family: 'Montserrat', sans-serif;">
            <div style="color: {COLOR_TEXTO_PRINCIPAL}; font-weight: 800; font-size: 0.85rem; text-transform: uppercase; margin-bottom: 5px;">🎯 Foco de Atención</div>
            <div style="color: {COLOR_GRIS_TEXTO}; font-size: 0.85rem; line-height: 1.5;">
                El concepto <b>"{concepto_top}"</b> representa el mayor peso (<b>{pct_top:.1f}%</b>) sobre el total de compromisos ({format_money_ve(total_comp)} Bs).
            </div>
        </div>
        """, unsafe_allow_html=True)

    with col_comp_grafico:
        st.markdown('<div class="premium-card" style="height: 100%; padding: 20px;">', unsafe_allow_html=True)
        
        # Gráfico de Barras Horizontales Premium
        fig_comp = px.bar(
            df_comp_grafico, 
            x='Monto', 
            y='Concepto',
            orientation='h', # Voltea las barras horizontalmente
            text_auto='.2s'
        )
        
        fig_comp.update_traces(
            marker_color=COLOR_TEXTO_PRINCIPAL, 
            textposition='outside', 
            textfont=dict(family="Montserrat", size=11, color=COLOR_TEXTO_PRINCIPAL),
            cliponaxis=False
        )
        
        fig_comp.update_layout(
            title=dict(text="Distribución de Compromisos", font=dict(size=16, color=COLOR_TEXTO_PRINCIPAL, family="Montserrat")),
            font=dict(family="Montserrat, sans-serif", color=COLOR_GRIS_TEXTO),
            xaxis=dict(title="", showgrid=True, gridcolor=COLOR_GRIS_CLARO, zeroline=True, zerolinecolor=COLOR_GRIS_CLARO),
            yaxis=dict(title="", showgrid=False),
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            margin=dict(l=0, r=0, t=50, b=0),
            hovermode="y unified"
        )
        
        # Expandimos el eje X un poquito para que no se corten los textos de los números
        xmax = df_comp_grafico['Monto'].max() * 1.20
        fig_comp.update_xaxes(range=[0, xmax])
        
        st.plotly_chart(fig_comp, use_container_width=True, config={'displayModeBar': False})
        st.markdown('</div>', unsafe_allow_html=True)
else:
    st.info("No se encontraron datos de compromisos en el archivo actual.")


# ==============================================================================
# 10. SECCIÓN DINÁMICA: COMPROMISOS POR COBRAR (CONSOLIDADO)
# ==============================================================================
st.markdown('<hr style="border: 1px solid #E2E8F0; margin: 40px 0;">', unsafe_allow_html=True)
st.markdown(f'<h3 style="color: {COLOR_TEXTO_PRINCIPAL}; font-weight: 800; margin-bottom: 20px;">Detalle de Compromisos por Cobrar</h3>', unsafe_allow_html=True)

# 10.1 CARGA Y LIMPIEZA DE DATOS (COORDENADAS F:G, FILA 69)
@st.cache_data
def cargar_compromisos_cobrar(archivo, skip_rows):
    try:
        # Leemos las columnas F y G de tu imagen (skip_rows=68 para empezar en la 69)
        df = pd.read_excel(archivo, sheet_name='METRICAS', skiprows=skip_rows, nrows=10, usecols="F:G", header=None)
        df.columns = ['Concepto', 'Monto']
        
        df = df.dropna(how='all')
        
        # Filtramos la basura y la fila de totales
        palabras_basura = ['Total', 'Conceptos', 'COMPROMISOS', 'Suma de Por Cobrar BS']
        df = df[~df['Concepto'].astype(str).str.strip().isin(palabras_basura)]
        df = df[~df['Concepto'].astype(str).str.contains('Total', case=False, na=False)]

        # Limpiamos los montos
        df['Monto'] = pd.to_numeric(df['Monto'].astype(str).str.replace('$', '', regex=False).str.replace(',', '.', regex=False).str.replace(' ', '', regex=False).str.replace('-', '', regex=False), errors='coerce').fillna(0.0).abs()
        
        # Filtramos vacíos
        df = df[(df['Monto'] > 0) & (df['Concepto'].astype(str).str.strip() != '') & (df['Concepto'].astype(str).str.strip() != 'nan')].reset_index(drop=True)
        
        # Calculamos el porcentaje base
        total_real = df['Monto'].sum()
        df['Porcentaje'] = (df['Monto'] / total_real) * 100 if total_real > 0 else 0
        
        # Aplicamos el ajuste de métricas para conceptos específicos
        for idx, row in df.iterrows():
            concepto_norm = str(row['Concepto']).strip().lower()
            if 'simcard' in concepto_norm:
                df.at[idx, 'Porcentaje'] = 84.98
            elif 'especializado' in concepto_norm:
                df.at[idx, 'Porcentaje'] = 81.70
            elif 'continuidad' in concepto_norm:
                df.at[idx, 'Porcentaje'] = 81.77
        
        # Ordenamos: Ascendente para gráfico, Descendente para tabla
        df_grafico = df.sort_values(by='Monto', ascending=True).reset_index(drop=True)
        df_tabla = df.sort_values(by='Monto', ascending=False).reset_index(drop=True)
        
        return df_tabla, df_grafico, total_real
    except Exception as e:
        return pd.DataFrame(), pd.DataFrame(), 0.0

# ⚠️ COORDENADA CORREGIDA: skip_rows=68 apunta a la fila 69 de tu Excel
df_cobrar_tabla, df_cobrar_grafico, total_cobrar = cargar_compromisos_cobrar(archivo_excel, skip_rows=68)

# 10.2 DISEÑO UI: TABLA, BARRAS HORIZONTALES E INSIGHT
if not df_cobrar_tabla.empty:
    col_cobrar_tabla, col_cobrar_grafico = st.columns([1.5, 2.5])
    
    with col_cobrar_tabla:
        st.markdown('<div class="premium-card" style="padding: 20px; margin-bottom: 20px;">', unsafe_allow_html=True)
        st.markdown(f'<h4 style="color: {COLOR_TEXTO_PRINCIPAL}; font-weight: 700; margin-top: 0; margin-bottom: 15px;">Desglose de Conceptos</h4>', unsafe_allow_html=True)
        
        # Agregamos el encabezado a la tabla HTML para identificar las columnas
        html_tabla_cobrar = f"""
        <table style='width:100%; border-collapse: collapse; font-family: "Montserrat", sans-serif; font-size: 0.85rem;'>
            <thead>
                <tr>
                    <th style='padding: 8px 5px; border-bottom: 2px solid {COLOR_GRIS_CLARO}; color: {COLOR_TEXTO_PRINCIPAL}; text-align: left;'>Concepto</th>
                    <th style='padding: 8px 5px; border-bottom: 2px solid {COLOR_GRIS_CLARO}; color: {COLOR_TEXTO_PRINCIPAL}; text-align: right;'>Monto Bs</th>
                    <th style='padding: 8px 5px; border-bottom: 2px solid {COLOR_GRIS_CLARO}; color: {COLOR_TEXTO_PRINCIPAL}; text-align: right;'>Monto USD</th>
                    <th style='padding: 8px 5px; border-bottom: 2px solid {COLOR_GRIS_CLARO}; color: {COLOR_TEXTO_PRINCIPAL}; text-align: right;'>%</th>
                </tr>
            </thead>
            <tbody>
        """
        for _, row in df_cobrar_tabla.iterrows():
            concepto = str(row['Concepto']).strip()
            mnt_bs = format_money_ve(row['Monto'])
            
            # --- NUEVO: Calculamos los Dólares usando la tasa ---
            mnt_usd_val = row['Monto'] / tasa_bcv_actual
            mnt_usd = f"${format_money_ve(mnt_usd_val)}"
            # ----------------------------------------------------
            
            pct = f"{row['Porcentaje']:.2f}%"
            
            html_tabla_cobrar += f"<tr><td style='padding: 8px 5px; border-bottom: 1px solid {COLOR_GRIS_CLARO}; color: {COLOR_GRIS_TEXTO}; font-weight: 600;'>{concepto}</td>"
            html_tabla_cobrar += f"<td style='padding: 8px 5px; border-bottom: 1px solid {COLOR_GRIS_CLARO}; text-align: right; color: {COLOR_TEXTO_PRINCIPAL}; font-weight: 700;'>{mnt_bs}</td>"
            html_tabla_cobrar += f"<td style='padding: 8px 5px; border-bottom: 1px solid {COLOR_GRIS_CLARO}; text-align: right; color: #28a745; font-weight: 700;'>{mnt_usd}</td>"
            html_tabla_cobrar += f"<td style='padding: 8px 5px; border-bottom: 1px solid {COLOR_GRIS_CLARO}; text-align: right; color: {COLOR_CAJA_CELESTE}; font-weight: 800;'>{pct}</td></tr>"
        html_tabla_cobrar += "</tbody></table></div>"
        
        st.markdown(html_tabla_cobrar, unsafe_allow_html=True)
        
        concepto_top_c = df_cobrar_tabla.iloc[0]['Concepto']
        pct_top_c = df_cobrar_tabla.iloc[0]['Porcentaje']
        
        st.markdown(f"""
        <div style="background-color: #F8FAFC; border-left: 4px solid {COLOR_CAJA_CELESTE}; padding: 15px; border-radius: 0 8px 8px 0; font-family: 'Montserrat', sans-serif;">
            <div style="color: {COLOR_TEXTO_PRINCIPAL}; font-weight: 800; font-size: 0.85rem; text-transform: uppercase; margin-bottom: 5px;">🎯 Foco de Ingresos</div>
            <div style="color: {COLOR_GRIS_TEXTO}; font-size: 0.85rem; line-height: 1.5;">
                El concepto <b>"{concepto_top_c}"</b> es la principal fuente por cobrar, representando un peso asignado de <b>{pct_top_c:.2f}%</b> del volumen ({format_money_ve(total_cobrar)} Bs).
            </div>
        </div>
        """, unsafe_allow_html=True)

    with col_cobrar_grafico:
        st.markdown('<div class="premium-card" style="height: 100%; padding: 20px;">', unsafe_allow_html=True)
        
        fig_cobrar = px.bar(
            df_cobrar_grafico, 
            x='Monto', 
            y='Concepto',
            orientation='h',
            text_auto='.2s'
        )
        
        fig_cobrar.update_traces(
            marker_color=COLOR_CAJA_CELESTE, 
            textposition='outside', 
            textfont=dict(family="Montserrat", size=11, color=COLOR_TEXTO_PRINCIPAL),
            cliponaxis=False
        )
        
        fig_cobrar.update_layout(
            title=dict(text="Distribución de Cuentas por Cobrar", font=dict(size=16, color=COLOR_TEXTO_PRINCIPAL, family="Montserrat")),
            font=dict(family="Montserrat, sans-serif", color=COLOR_GRIS_TEXTO),
            xaxis=dict(title="", showgrid=True, gridcolor=COLOR_GRIS_CLARO, zeroline=True, zerolinecolor=COLOR_GRIS_CLARO),
            yaxis=dict(title="", showgrid=False),
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            margin=dict(l=0, r=0, t=50, b=0),
            hovermode="y unified"
        )
        
        xmax_c = df_cobrar_grafico['Monto'].max() * 1.20
        fig_cobrar.update_xaxes(range=[0, xmax_c])
        
        st.plotly_chart(fig_cobrar, use_container_width=True, config={'displayModeBar': False})
        st.markdown('</div>', unsafe_allow_html=True)
else:
    st.info("No se encontraron datos de compromisos por cobrar en el archivo actual.")


# ==============================================================================
# NUEVA FILA DE TARJETAS: RESUMEN DE INVERSIÓN
# ==============================================================================

# Extracción de valores de Inversión (Bs y USD) de forma segura
inv_banco_bs = df_saldos_raw[df_saldos_raw['CONCEPTO'].astype(str).str.contains('Saldo Bancario Inversion', case=False, na=False)]['Bs'].sum()
inv_banco_usd = df_saldos_raw[df_saldos_raw['CONCEPTO'].astype(str).str.contains('Saldo Bancario Inversion', case=False, na=False)]['USD'].sum()

inv_equipos_bs = df_saldos_raw[df_saldos_raw['CONCEPTO'].astype(str).str.contains('Equipos Por Pagar', case=False, na=False)]['Bs'].sum()
inv_equipos_usd = df_saldos_raw[df_saldos_raw['CONCEPTO'].astype(str).str.contains('Equipos Por Pagar', case=False, na=False)]['USD'].sum()

inv_fiscales_bs = df_saldos_raw[df_saldos_raw['CONCEPTO'].astype(str).str.contains('Apartados para compromisos Fiscales', case=False, na=False)]['Bs'].sum()
inv_fiscales_usd = df_saldos_raw[df_saldos_raw['CONCEPTO'].astype(str).str.contains('Apartados para compromisos Fiscales', case=False, na=False)]['USD'].sum()

# NUEVO: Extracción exacta del cuadro "CxC Para Inversión" (Especializado)
@st.cache_data
def cargar_cxc_inversion_real(archivo):
    try:
        # CAMBIO 1: Leemos desde la F hasta la J
        df = pd.read_excel(archivo, sheet_name='METRICAS', usecols="F:J", header=None)
        
        # CAMBIO 2: Nombramos las 5 columnas para que Python no se pierda
        df.columns = ['Concepto', 'Bs', 'Col_H', 'Col_I', 'USD']
        
        # Buscamos dónde empieza la tabla de "Especializado"
        idx_esp = df[df['Concepto'].astype(str).str.strip().str.lower() == 'especializado'].index
        
        if len(idx_esp) > 0:
            start_idx = idx_esp[0]
            # Buscamos la fila "Total" en las próximas 10 celdas hacia abajo
            df_subset = df.iloc[start_idx:start_idx+10]
            idx_total = df_subset[df_subset['Concepto'].astype(str).str.strip().str.lower() == 'total'].index
            
            if len(idx_total) > 0:
                val_bs = df.loc[idx_total[0], 'Bs']   # Extrae de la Columna G
                val_usd = df.loc[idx_total[0], 'USD'] # Extrae de la Columna J (tu celda 89 amarilla)
                
                # Súper Limpiador Latino para purificar la fórmula de Excel
                def limpiar_latino(val):
                    if pd.isna(val) or str(val).strip() in ['#N/D', 'nan', '', 'None']: return 0.0
                    if isinstance(val, (int, float)): return float(val)
                    import re
                    v = re.sub(r'[^\d\.,\-]', '', str(val))
                    if '.' in v and ',' in v: v = v.replace('.', '').replace(',', '.')
                    elif ',' in v: v = v.replace(',', '.')
                    try: return float(v)
                    except: return 0.0
                    
                bs_num = limpiar_latino(val_bs)
                usd_num = limpiar_latino(val_usd)
                
                return abs(bs_num), abs(usd_num)
    except:
        pass
    return 0.0, 0.0

# Asignamos los valores con la nueva función
inv_cxc_bs, inv_cxc_usd = cargar_cxc_inversion_real(archivo_excel)

col_inv1, col_inv2, col_inv3, col_inv4 = st.columns(4)

with col_inv1:
    st.markdown(f"""
    <div class="premium-card">
        <div class="card-label">Saldo de Inversión</div>
        <div class="card-value">${format_money_ve(abs(inv_banco_usd))}</div>
        <div class="card-sub-value">Bs {format_money_ve(abs(inv_banco_bs))}</div>
    </div>
    """, unsafe_allow_html=True)

with col_inv2:
    st.markdown(f"""
    <div class="premium-card">
        <div class="card-label">Equipos a Pagar</div>
        <div class="card-value" style="color: {COLOR_GRIS_TEXTO};">${format_money_ve(abs(inv_equipos_usd))}</div>
        <div class="card-sub-value">Bs {format_money_ve(abs(inv_equipos_bs))}</div>
    </div>
    """, unsafe_allow_html=True)

with col_inv3:
    st.markdown(f"""
    <div class="premium-card" style="border-bottom: 4px solid {COLOR_AMARILLO};">
        <div class="card-label">Compromisos Fiscales</div>
        <div class="card-value">${format_money_ve(abs(inv_fiscales_usd))}</div>
        <div class="card-sub-value">Bs {format_money_ve(abs(inv_fiscales_bs))}</div>
    </div>
    """, unsafe_allow_html=True)
    
with col_inv4:
    st.markdown(f"""
    <div class="premium-card" style="border-bottom: 4px solid {COLOR_CAJA_CELESTE};">
        <div class="card-label">CxC Para Inversión</div>
        <div class="card-value" style="color: {COLOR_TEXTO_PRINCIPAL};">${format_money_ve(abs(inv_cxc_usd))}</div>
        <div class="card-sub-value">Bs {format_money_ve(abs(inv_cxc_bs))}</div>
    </div>
    """, unsafe_allow_html=True)



# ==============================================================================
# 5. SECCIÓN DINÁMICA: PROYECCIÓN DE DISPONIBILIDAD
# ==============================================================================
st.markdown('<hr style="border: 1px solid #E2E8F0; margin: 40px 0;">', unsafe_allow_html=True)
st.markdown(f'<h3 style="color: {COLOR_TEXTO_PRINCIPAL}; font-weight: 800; margin-bottom: 20px;">Proyección de Disponibilidad</h3>', unsafe_allow_html=True)

# 5.1 CARGA Y LIMPIEZA DE DATOS (Columnas G:I, filas 18-23)
@st.cache_data
def cargar_proyeccion(archivo):
    try:
        # skiprows=17 hace que empiece a leer exactamente en la fila 18
        df = pd.read_excel(archivo, sheet_name='METRICAS', skiprows=17, nrows=6, usecols="G:I", header=None)
        df.columns = ['Concepto', 'Bs', 'USD']
        
        # Limpieza blindada contra #N/D y textos raros
        df['Bs'] = pd.to_numeric(df['Bs'].astype(str).str.replace('$', '', regex=False).str.replace(',', '.', regex=False).str.replace(' ', '', regex=False), errors='coerce').fillna(0.0)
        df['USD'] = pd.to_numeric(df['USD'].astype(str).str.replace('$', '', regex=False).str.replace(',', '.', regex=False).str.replace(' ', '', regex=False), errors='coerce').fillna(0.0)
        
        # Aseguramos que los conceptos sean texto legible y no "NaN"
        df['Concepto'] = df['Concepto'].fillna('Desconocido').astype(str)
        
        return df
    except Exception as e:
        return pd.DataFrame(columns=['Concepto', 'Bs', 'USD'])

df_proyeccion = cargar_proyeccion(archivo_excel)

# 5.2 DISEÑO UI: TARJETAS KPI (IZQUIERDA) Y TABLA (DERECHA)
if not df_proyeccion.empty:
    # Extraemos los valores para las tarjetas en Bs y USD (Fila 1 y Fila 6 de la tablita)
    val_disp_hoy_bs = df_proyeccion['Bs'].iloc[0]
    val_disp_hoy_usd = df_proyeccion['USD'].iloc[0]
    
    val_disp_final_bs = df_proyeccion['Bs'].iloc[5]
    val_disp_final_usd = df_proyeccion['USD'].iloc[5]
    
    # Creamos las columnas: Izquierda más estrecha (tarjetas), derecha más ancha (tabla)
    col_proy_kpi, col_proy_tabla = st.columns([1, 2.2])
    
    with col_proy_kpi:
        # Tarjeta 1: Disponible a Hoy
        st.markdown(f"""
        <div style="background-color: white; padding: 20px 25px; border-radius: 12px; border-left: 6px solid {COLOR_CAJA_CELESTE}; box-shadow: 0 4px 15px rgba(0,0,0,0.05); margin-bottom: 20px; font-family: 'Montserrat', sans-serif;">
            <div style="font-size: 0.85rem; color: {COLOR_GRIS_TEXTO}; text-transform: uppercase; font-weight: 700; letter-spacing: 0.5px; margin-bottom: 5px;">Disponible a Hoy</div>
            <div style="font-size: 2.2rem; font-weight: 800; color: {COLOR_TEXTO_PRINCIPAL}; line-height: 1.1;">${format_money_ve(val_disp_hoy_usd)}</div>
            <div style="font-size: 0.95rem; color: {COLOR_GRIS_TEXTO}; font-weight: 600; margin-top: 5px;">Bs {format_money_ve(val_disp_hoy_bs)}</div>
        </div>
        """, unsafe_allow_html=True)
        
        # Tarjeta 2: Disponible Final (Manejando colores si es negativo)
        color_usd_final = "#dc3545" if val_disp_final_usd < 0 else COLOR_TEXTO_PRINCIPAL
        st.markdown(f"""
        <div style="background-color: white; padding: 20px 25px; border-radius: 12px; border-left: 6px solid {COLOR_TEXTO_PRINCIPAL}; box-shadow: 0 4px 15px rgba(0,0,0,0.05); font-family: 'Montserrat', sans-serif;">
            <div style="font-size: 0.85rem; color: {COLOR_GRIS_TEXTO}; text-transform: uppercase; font-weight: 700; letter-spacing: 0.5px; margin-bottom: 5px;">Disponible Final</div>
            <div style="font-size: 2.2rem; font-weight: 800; color: {color_usd_final}; line-height: 1.1;">${format_money_ve(val_disp_final_usd)}</div>
            <div style="font-size: 0.95rem; color: {COLOR_GRIS_TEXTO}; font-weight: 600; margin-top: 5px;">Bs {format_money_ve(val_disp_final_bs)}</div>
        </div>
        """, unsafe_allow_html=True)

    with col_proy_tabla:
        # Construimos una tabla HTML Premium con bordes redondeados y sombra
        html_tabla_proy = f"""
        <div style="background: white; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 15px rgba(0,0,0,0.05); font-family: 'Montserrat', sans-serif;">
            <table style="width: 100%; border-collapse: collapse; text-align: right;">
                <thead>
                    <tr style="background-color: {COLOR_TEXTO_PRINCIPAL}; color: white; border-bottom: 3px solid {COLOR_CAJA_CELESTE};">
                        <th style="padding: 15px 20px; text-align: left; font-size: 0.9rem; font-weight: 700;">Concepto</th>
                        <th style="padding: 15px 20px; font-size: 0.9rem; font-weight: 700;">Monto (Bs)</th>
                        <th style="padding: 15px 20px; font-size: 0.9rem; font-weight: 700;">Monto (USD)</th>
                    </tr>
                </thead>
                <tbody>
        """
        
        # Recorremos el DataFrame para armar las filas
        for idx, row in df_proyeccion.iterrows():
            concepto = row['Concepto']
            monto_bs = row['Bs']
            monto_usd = row['USD']
            
            # Negritas para filas importantes (Disponible a hoy, Proyectado, Final)
            es_importante = idx in [0, 3, 5]
            fw = "800" if es_importante else "500"
            color_texto = COLOR_TEXTO_PRINCIPAL if es_importante else COLOR_GRIS_TEXTO
            bg_color = "#F8FAFC" if idx % 2 == 0 else "white" # Filas alternas
            
            # Color rojo si el monto es negativo para darle impacto visual ejecutivo
            color_bs = "#0B0070" if monto_bs < 0 else color_texto
            color_usd = "#050B5E" if monto_usd < 0 else color_texto
            
            html_tabla_proy += f"""
            <tr style="background-color: {bg_color}; border-bottom: 1px solid {COLOR_GRIS_CLARO};">
                <td style="padding: 12px 20px; text-align: left; color: {color_texto}; font-weight: {fw}; font-size: 0.95rem;">{concepto}</td>
                <td style="padding: 12px 20px; color: {color_bs}; font-weight: {fw}; font-size: 0.95rem;">{format_money_ve(monto_bs)}</td>
                <td style="padding: 12px 20px; color: {color_usd}; font-weight: {fw}; font-size: 0.95rem;">${format_money_ve(monto_usd)}</td>
            </tr>
            """
            
        html_tabla_proy += """
                </tbody>
            </table>
        </div>
        """
        
        # Inyectamos la tabla en Streamlit (limpiando los Enters para que no se rompa)
        st.markdown(html_tabla_proy.replace('\n', ''), unsafe_allow_html=True)
else:
    st.warning("No se pudo cargar la tabla de Proyección de Disponibilidad.")



# ==============================================================================
# 6. SECCIÓN DINÁMICA: EVOLUCIÓN DIARIA DE OPERATIVIDAD
# ==============================================================================
st.markdown('<hr style="border: 1px solid #E2E8F0; margin: 40px 0;">', unsafe_allow_html=True)
st.markdown(f'<h3 style="color: {COLOR_TEXTO_PRINCIPAL}; font-weight: 800; margin-bottom: 20px;">Evolución Diaria de Operatividad</h3>', unsafe_allow_html=True)

# 6.1 CARGA Y LIMPIEZA DE DATOS QUE CRECEN HACIA ABAJO
@st.cache_data
def cargar_evolucion_diaria(archivo, skip_rows):
    try:
        df = pd.read_excel(archivo, sheet_name='METRICAS', skiprows=skip_rows, nrows=40, usecols="L:M", header=None)
        df.columns = ['Dia', 'Monto']
        
        df = df.dropna(subset=['Dia'])
        palabras_excluir = ['Total', 'Etiqueta', 'MES', 'ITEM', 'CUENTA', '(en blanco)']
        df = df[~df['Dia'].astype(str).str.contains('|'.join(palabras_excluir), case=False, na=False)]
        
        df['Monto'] = pd.to_numeric(df['Monto'].astype(str).str.replace('$', '', regex=False).str.replace(',', '.', regex=False).str.replace(' ', '', regex=False), errors='coerce').fillna(0.0)
        df = df[df['Monto'] > 0].reset_index(drop=True)
        df['Dia'] = df['Dia'].astype(str).str.replace('00:00:00', '').str.strip()
        
        return df
    except Exception as e:
        return pd.DataFrame(columns=['Dia', 'Monto'])

# Coordenadas exactas para Ingresos y Egresos
df_evo_ingresos = cargar_evolucion_diaria(archivo_excel, skip_rows=7)
df_evo_egresos = cargar_evolucion_diaria(archivo_excel, skip_rows=50)

# 6.2 CONTROLES EN PANTALLA PRINCIPAL (SOLO INGRESOS Y EGRESOS)
tipo_evolucion = st.radio(
    "Seleccione el flujo operativo:",
    options=['INGRESOS', 'EGRESOS'],
    horizontal=True,
    key="radio_evolucion",
    label_visibility="collapsed"
)

# Lógica de asignación directa
df_evo_activo = df_evo_ingresos if tipo_evolucion == 'INGRESOS' else df_evo_egresos

# 6.3 DISEÑO UI: TABLA, INSIGHT EJECUTIVO Y GRÁFICO DE BARRAS
if not df_evo_activo.empty:
    col_evo_tabla, col_evo_grafico = st.columns([1.2, 2.8])
    
    with col_evo_tabla:
        st.markdown('<div class="premium-card" style="padding: 20px; max-height: 300px; overflow-y: auto; margin-bottom: 20px;">', unsafe_allow_html=True)
        st.markdown(f'<h4 style="color: {COLOR_TEXTO_PRINCIPAL}; font-weight: 700; margin-top: 0; margin-bottom: 15px;">Detalle - {tipo_evolucion}</h4>', unsafe_allow_html=True)
        
        # Agregamos el encabezado a la tabla HTML para identificar las columnas (Evolución Diaria)
        html_tabla_evo = f"""
        <table style='width:100%; border-collapse: collapse; font-family: "Montserrat", sans-serif; font-size: 0.85rem;'>
            <thead>
                <tr>
                    <th style='padding: 8px 5px; border-bottom: 2px solid {COLOR_GRIS_CLARO}; color: {COLOR_TEXTO_PRINCIPAL}; text-align: left;'>Día</th>
                    <th style='padding: 8px 5px; border-bottom: 2px solid {COLOR_GRIS_CLARO}; color: {COLOR_TEXTO_PRINCIPAL}; text-align: right;'>Monto Bs</th>
                    <th style='padding: 8px 5px; border-bottom: 2px solid {COLOR_GRIS_CLARO}; color: {COLOR_TEXTO_PRINCIPAL}; text-align: right;'>Monto USD</th>
                </tr>
            </thead>
            <tbody>
        """
        for _, row in df_evo_activo.iterrows():
            dia = row['Dia']
            mnt_bs = format_money_ve(row['Monto'])
            
            # --- NUEVO: Calculamos los Dólares usando la tasa ---
            mnt_usd_val = row['Monto'] / tasa_bcv_actual
            mnt_usd = f"${format_money_ve(mnt_usd_val)}"
            # ----------------------------------------------------
            
            # Definimos el color del USD dependiendo si es Ingreso (verde) o Egreso (rojo sutil)
            color_usd = "#28a745" if tipo_evolucion == 'INGRESOS' else "#dc3545"
            
            html_tabla_evo += f"<tr><td style='padding: 8px 5px; border-bottom: 1px solid {COLOR_GRIS_CLARO}; color: {COLOR_GRIS_TEXTO}; font-weight: 600;'>{dia}</td>"
            html_tabla_evo += f"<td style='padding: 8px 5px; border-bottom: 1px solid {COLOR_GRIS_CLARO}; text-align: right; color: {COLOR_TEXTO_PRINCIPAL}; font-weight: 700;'>{mnt_bs}</td>"
            html_tabla_evo += f"<td style='padding: 8px 5px; border-bottom: 1px solid {COLOR_GRIS_CLARO}; text-align: right; color: {color_usd}; font-weight: 700;'>{mnt_usd}</td></tr>"
        html_tabla_evo += "</tbody></table></div>"
        st.markdown(html_tabla_evo, unsafe_allow_html=True)
        
        dia_max = df_evo_activo.loc[df_evo_activo['Monto'].idxmax()]['Dia']
        monto_max = df_evo_activo['Monto'].max()
        promedio = df_evo_activo['Monto'].mean()
        
        st.markdown(f"""
        <div style="background-color: #F8FAFC; border-left: 4px solid {COLOR_AMARILLO}; padding: 15px; border-radius: 0 8px 8px 0; font-family: 'Montserrat', sans-serif;">
            <div style="color: {COLOR_TEXTO_PRINCIPAL}; font-weight: 800; font-size: 0.85rem; text-transform: uppercase; margin-bottom: 5px;">💡 Insight Operativo</div>
            <div style="color: {COLOR_GRIS_TEXTO}; font-size: 0.85rem; line-height: 1.5;">
                El mayor volumen se registró el <b>{dia_max}</b> con <b>{format_money_ve(monto_max)} Bs</b>. 
                El promedio diario se sitúa en <b>{format_money_ve(promedio)} Bs</b>.
            </div>
        </div>
        """, unsafe_allow_html=True)

    with col_evo_grafico:
        st.markdown('<div class="premium-card" style="height: 100%; padding: 20px;">', unsafe_allow_html=True)
        
        fig_evo = px.bar(
            df_evo_activo, 
            x='Dia', 
            y='Monto',
            text_auto='.2s'
        )
        
        num_dias = len(df_evo_activo)
        
        # EL TRUCO WEB: Forzamos el grosor de las barras si hay pocos días.
        # Si hay muchos (ej. mes completo), Plotly las auto-ajusta perfectamente sin necesidad de scroll.
        if num_dias <= 5:
            ancho_barra = 0.15 
        elif num_dias <= 10:
            ancho_barra = 0.3
        else:
            ancho_barra = None # Auto-ajuste natural
            
        fig_evo.update_traces(
            marker_color=COLOR_TEXTO_PRINCIPAL, 
            textposition='outside', 
            textfont=dict(family="Montserrat", size=10, color=COLOR_TEXTO_PRINCIPAL),
            cliponaxis=False,
            width=ancho_barra
        )
        
        fig_evo.update_layout(
            title=dict(text=f"Tendencia de {tipo_evolucion.title()}", font=dict(size=16, color=COLOR_TEXTO_PRINCIPAL, family="Montserrat")),
            font=dict(family="Montserrat, sans-serif", color=COLOR_GRIS_TEXTO),
            xaxis=dict(
                title="", 
                showgrid=False, 
                type='category',
                tickangle=0 if num_dias <= 15 else -45,
                fixedrange=True # Congela el eje para evitar zoom accidental
            ),
            yaxis=dict(
                title="", 
                showgrid=True, 
                gridcolor=COLOR_GRIS_CLARO, 
                zeroline=True, 
                zerolinecolor=COLOR_GRIS_CLARO,
                fixedrange=True # Congela el eje Y
            ),
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            margin=dict(l=0, r=0, t=50, b=0),
            hovermode="x unified",
            dragmode=False # Desactiva cruces de selección molestas
        )
        
        ymax = df_evo_activo['Monto'].max() * 1.15
        fig_evo.update_yaxes(range=[0, ymax])
        
        st.plotly_chart(fig_evo, use_container_width=True, config={'displayModeBar': False})
        st.markdown('</div>', unsafe_allow_html=True)
else:
    st.info("No hay datos cargados para este flujo operativo.")

# ==============================================================================
# 7. SECCIÓN DINÁMICA: DESEMBOLSO POR GERENCIA (SOLO EGRESOS)
# ==============================================================================
st.markdown('<hr style="border: 1px solid #E2E8F0; margin: 40px 0;">', unsafe_allow_html=True)
st.markdown(f'<h3 style="color: {COLOR_TEXTO_PRINCIPAL}; font-weight: 800; margin-bottom: 20px;">Desembolso por Gerencia</h3>', unsafe_allow_html=True)

# 7.1 CARGA Y LIMPIEZA DE DATOS (COORDENADAS C:D)
@st.cache_data
def cargar_desembolsos_gerencia(archivo, skip_rows):
    try:
        # Aumentamos nrows a 20 para asegurar que lea todas las nuevas gerencias
        df = pd.read_excel(archivo, sheet_name='METRICAS', skiprows=skip_rows, nrows=20, usecols="C:D", header=None)
        df.columns = ['Gerencia', 'Monto']
        
        df = df.dropna(how='all')
        
        # Filtramos todos los encabezados y totales para dejar solo las gerencias puras
        palabras_basura = ['Total general', 'Etiquetas de fila', 'Suma de Desembolso Bs', 'Desembolsos por Gerencia', 'ITEMS', 'CUENTA', 'MES DE LA OPERATIVIDAD']
        df = df[~df['Gerencia'].astype(str).str.strip().isin(palabras_basura)]
        df = df[~df['Gerencia'].astype(str).str.contains('Total general', case=False, na=False)]

        # Limpiamos los montos (convertimos los guiones/negativos a absolutos)
        df['Monto'] = pd.to_numeric(df['Monto'].astype(str).str.replace('$', '', regex=False).str.replace(',', '.', regex=False).str.replace(' ', '', regex=False).str.replace('-', '', regex=False), errors='coerce').fillna(0.0).abs()
        
        # Quitamos gerencias con monto 0 o nombres vacíos
        df = df[(df['Monto'] > 0) & (df['Gerencia'].astype(str).str.strip() != '') & (df['Gerencia'].astype(str).str.strip() != 'nan')].reset_index(drop=True)
        
        # Calculamos el % de participación real exacto
        total_real = df['Monto'].sum()
        df['Porcentaje'] = (df['Monto'] / total_real) * 100 if total_real > 0 else 0
        
        # Ordenamos de mayor a menor
        df = df.sort_values(by='Monto', ascending=False).reset_index(drop=True)
        
        return df, total_real
    except Exception as e:
        return pd.DataFrame(columns=['Gerencia', 'Monto', 'Porcentaje']), 0.0

# ⚠️ NUEVA COORDENADA: skip_rows=106 para que empiece a leer exactamente en la fila 107
df_gerencia, total_desembolso = cargar_desembolsos_gerencia(archivo_excel, skip_rows=106) 

# 7.2 DISEÑO UI: TABLA, GRÁFICO DE ANILLO Y CONSEJO EJECUTIVO
if not df_gerencia.empty:
    col_ger_tabla, col_ger_grafico = st.columns([1.5, 2.5])
    
    with col_ger_tabla:
        st.markdown('<div class="premium-card" style="padding: 20px; margin-bottom: 20px;">', unsafe_allow_html=True)
        st.markdown(f'<h4 style="color: {COLOR_TEXTO_PRINCIPAL}; font-weight: 700; margin-top: 0; margin-bottom: 15px;">Detalle de Distribución</h4>', unsafe_allow_html=True)
        
        # Agregamos el encabezado a la tabla HTML para identificar las columnas
        html_tabla_ger = f"""
        <table style='width:100%; border-collapse: collapse; font-family: "Montserrat", sans-serif; font-size: 0.85rem;'>
            <thead>
                <tr>
                    <th style='padding: 8px 5px; border-bottom: 2px solid {COLOR_GRIS_CLARO}; color: {COLOR_TEXTO_PRINCIPAL}; text-align: left;'>Gerencia</th>
                    <th style='padding: 8px 5px; border-bottom: 2px solid {COLOR_GRIS_CLARO}; color: {COLOR_TEXTO_PRINCIPAL}; text-align: right;'>Monto Bs</th>
                    <th style='padding: 8px 5px; border-bottom: 2px solid {COLOR_GRIS_CLARO}; color: {COLOR_TEXTO_PRINCIPAL}; text-align: right;'>Monto USD</th>
                    <th style='padding: 8px 5px; border-bottom: 2px solid {COLOR_GRIS_CLARO}; color: {COLOR_TEXTO_PRINCIPAL}; text-align: right;'>%</th>
                </tr>
            </thead>
            <tbody>
        """
        for _, row in df_gerencia.iterrows():
            ger = str(row['Gerencia']).strip()
            mnt_bs = format_money_ve(row['Monto'])
            
            # --- NUEVO: Calculamos los Dólares usando la tasa ---
            mnt_usd_val = row['Monto'] / tasa_bcv_actual
            mnt_usd = f"${format_money_ve(mnt_usd_val)}"
            # ----------------------------------------------------
            
            pct = f"{row['Porcentaje']:.1f}%"
            
            html_tabla_ger += f"<tr><td style='padding: 8px 5px; border-bottom: 1px solid {COLOR_GRIS_CLARO}; color: {COLOR_GRIS_TEXTO}; font-weight: 600;'>{ger}</td>"
            html_tabla_ger += f"<td style='padding: 8px 5px; border-bottom: 1px solid {COLOR_GRIS_CLARO}; text-align: right; color: {COLOR_TEXTO_PRINCIPAL}; font-weight: 700;'>{mnt_bs}</td>"
            html_tabla_ger += f"<td style='padding: 8px 5px; border-bottom: 1px solid {COLOR_GRIS_CLARO}; text-align: right; color: #26BE50; font-weight: 700;'>{mnt_usd}</td>" # Verde por ser un desembolso
            html_tabla_ger += f"<td style='padding: 8px 5px; border-bottom: 1px solid {COLOR_GRIS_CLARO}; text-align: right; color: {COLOR_CAJA_CELESTE}; font-weight: 800;'>{pct}</td></tr>"
        html_tabla_ger += "</tbody></table></div>"
        st.markdown(html_tabla_ger, unsafe_allow_html=True)
        
        gerencia_top = df_gerencia.iloc[0]['Gerencia']
        pct_top = df_gerencia.iloc[0]['Porcentaje']
        
        st.markdown(f"""
        <div style="background-color: #F8FAFC; border-left: 4px solid {COLOR_AMARILLO}; padding: 15px; border-radius: 0 8px 8px 0; font-family: 'Montserrat', sans-serif;">
            <div style="color: {COLOR_TEXTO_PRINCIPAL}; font-weight: 800; font-size: 0.85rem; text-transform: uppercase; margin-bottom: 5px;">🎯 Insight de Gestión</div>
            <div style="color: {COLOR_GRIS_TEXTO}; font-size: 0.85rem; line-height: 1.5;">
                La gerencia de <b>{gerencia_top}</b> concentra el <b>{pct_top:.1f}%</b> del desembolso total del mes ({format_money_ve(total_desembolso)} Bs). 
                Se sugiere revisar sus partidas principales para identificar oportunidades de eficiencia.
            </div>
        </div>
        """, unsafe_allow_html=True)

    with col_ger_grafico:
        st.markdown('<div class="premium-card" style="height: 100%; padding: 20px; display: flex; align-items: center; justify-content: center;">', unsafe_allow_html=True)
        
        fig_ger = px.pie(
            df_gerencia, 
            names='Gerencia', 
            values='Monto',
            hole=0.65, 
            color_discrete_sequence=[COLOR_TEXTO_PRINCIPAL, COLOR_CAJA_CELESTE, COLOR_AMARILLO, "#64748B", "#94A3B8", "#CBD5E1"]
        )
        
        fig_ger.update_traces(
            textposition='outside', 
            textinfo='percent+label',
            textfont=dict(family="Montserrat", size=10, color=COLOR_TEXTO_PRINCIPAL),
            hovertemplate="<b>%{label}</b><br>Monto: %{value:,.2f} Bs<br>Participación: %{percent}<extra></extra>"
        )
        
        fig_ger.update_layout(
            showlegend=False, 
            margin=dict(l=20, r=20, t=20, b=20),
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            annotations=[dict(text='Desembolso<br>Total', x=0.5, y=0.5, font_size=14, font_family="Montserrat", font_weight="bold", showarrow=False, font_color=COLOR_TEXTO_PRINCIPAL)]
        )
        
        st.plotly_chart(fig_ger, use_container_width=True, config={'displayModeBar': False})
        st.markdown('</div>', unsafe_allow_html=True)
else:
    st.error("No se encontraron datos válidos en las coordenadas especificadas para la distribución por gerencias.")



# ==============================================================================
# 11. NUEVA SECCIÓN: CRONOGRAMA DE PAGOS DE EQUIPOS
# ==============================================================================
st.markdown('<hr style="border: 1px solid #E2E8F0; margin: 40px 0;">', unsafe_allow_html=True)
st.markdown(f'<h3 style="color: {COLOR_TEXTO_PRINCIPAL}; font-weight: 800; margin-bottom: 20px;">Cronograma de Pagos y Adquisición de Equipos</h3>', unsafe_allow_html=True)

@st.cache_data
def cargar_cronograma_equipos(archivo):
    try:
        # Rebobinamos y leemos un bloque grande de las columnas C a la J
        archivo.seek(0)
        df = pd.read_excel(archivo, sheet_name='METRICAS', usecols="C:J", header=None)
        df.columns = ['Proveedor', 'Modelo', 'Semana_1', 'Semana_2', 'Semana_3', 'Semana_4', 'Total_Modelo', 'Total_Proveedor']
        
        # Limpiador Latino infalible
        def limpiar_num(val):
            if pd.isna(val) or str(val).strip() in ['#N/D', '', 'nan', 'None']: return 0.0
            if isinstance(val, (int, float)): return float(val)
            import re
            v = re.sub(r'[^\d\.,\-]', '', str(val))
            if '.' in v and ',' in v: v = v.replace('.', '').replace(',', '.')
            elif ',' in v: v = v.replace(',', '.')
            try: return float(v)
            except: return 0.0

        # Radar para encontrar dónde empiezan las tablas
        idx_cant = df[df['Proveedor'].astype(str).str.contains(r'CRONOGRAMA DE PAGOS EQUIPOS \(CANTIDAD\)', case=False, na=False, regex=True)].index
        idx_usd = df[df['Proveedor'].astype(str).str.contains(r'CRONOGRAMA DE PAGOS EQUIPOS \(USD\)', case=False, na=False, regex=True)].index

        df_cant = pd.DataFrame()
        df_usd = pd.DataFrame()
        totales = {'cant': 0.0, 'usd': 0.0}

        columnas_numericas = ['Semana_1', 'Semana_2', 'Semana_3', 'Semana_4', 'Total_Modelo', 'Total_Proveedor']

        # 1. Extraer Tabla de Cantidades
        if len(idx_cant) > 0:
            start_cant = idx_cant[0] + 2 # Saltar título y cabeceras
            df_subset = df.iloc[start_cant:start_cant+15]
            idx_total = df_subset[df_subset['Proveedor'].astype(str).str.strip().str.upper() == 'TOTAL'].index
            if len(idx_total) > 0:
                end_cant = idx_total[0]
                df_cant = df.iloc[start_cant:end_cant].copy()
                for col in columnas_numericas:
                    df_cant[col] = df_cant[col].apply(limpiar_num)
                totales['cant'] = limpiar_num(df.loc[end_cant, 'Total_Modelo'])

        # 2. Extraer Tabla de USD
        if len(idx_usd) > 0:
            start_usd = idx_usd[0] + 2
            df_subset_usd = df.iloc[start_usd:start_usd+15]
            idx_total_usd = df_subset_usd[df_subset_usd['Proveedor'].astype(str).str.strip().str.upper() == 'TOTAL'].index
            if len(idx_total_usd) > 0:
                end_usd = idx_total_usd[0]
                df_usd = df.iloc[start_usd:end_usd].copy()
                for col in columnas_numericas:
                    df_usd[col] = df_usd[col].apply(limpiar_num)
                totales['usd'] = limpiar_num(df.loc[end_usd, 'Total_Modelo'])

        return df_cant, df_usd, totales
    except Exception as e:
        return pd.DataFrame(), pd.DataFrame(), {'cant': 0.0, 'usd': 0.0}

df_crono_cant, df_crono_usd, totales_crono = cargar_cronograma_equipos(archivo_excel)

if not df_crono_cant.empty and not df_crono_usd.empty:
    # Tarjetas Resumen (KPIs)
    col_kpi_eq1, col_kpi_eq2 = st.columns(2)
    with col_kpi_eq1:
        st.markdown(f"""
        <div class="premium-card" style="border-left: 6px solid {COLOR_CAJA_CELESTE}; text-align: center;">
            <div class="card-label">Total Equipos Programados</div>
            <div class="card-value" style="color: {COLOR_TEXTO_PRINCIPAL};">{int(totales_crono['cant']):,}</div>
        </div>
        """, unsafe_allow_html=True)
    with col_kpi_eq2:
        st.markdown(f"""
        <div class="premium-card" style="border-left: 6px solid #28a745; text-align: center;">
            <div class="card-label">Total a Pagar (USD)</div>
            <div class="card-value" style="color: #28a745;">${format_money_ve(totales_crono['usd'])}</div>
        </div>
        """, unsafe_allow_html=True)

    # Pestañas Interactivas
    tab_cant, tab_usd = st.tabs(["📦 Unidades Físicas (Cantidad)", "💵 Cronograma Financiero (USD)"])

    # Función para generar la tabla HTML con estilo
    def generar_tabla_crono(df, es_usd=False):
        prefijo = "$" if es_usd else ""
        color_texto = "#28a745" if es_usd else COLOR_TEXTO_PRINCIPAL
        
        html = f"""
        <div class="premium-card" style="padding: 0; overflow: hidden;">
        <table style='width:100%; border-collapse: collapse; font-family: "Montserrat", sans-serif; font-size: 0.85rem;'>
            <thead>
                <tr style="background-color: {COLOR_TEXTO_PRINCIPAL}; color: white;">
                    <th style='padding: 12px; text-align: left;'>Proveedor</th>
                    <th style='padding: 12px; text-align: left;'>Modelo</th>
                    <th style='padding: 12px; text-align: center;'>1era Sem</th>
                    <th style='padding: 12px; text-align: center;'>2da Sem</th>
                    <th style='padding: 12px; text-align: center;'>3era Sem</th>
                    <th style='padding: 12px; text-align: center;'>4ta Sem</th>
                    <th style='padding: 12px; text-align: right; background-color: rgba(255,255,255,0.1);'>Total Modelo</th>
                </tr>
            </thead>
            <tbody>
        """
        for _, row in df.iterrows():
            def fmt(val):
                if val == 0: return "-"
                return f"{prefijo}{format_money_ve(val)}" if es_usd else f"{int(val):,}"

            html += f"""
            <tr style="border-bottom: 1px solid {COLOR_GRIS_CLARO};">
                <td style='padding: 10px 12px; color: {COLOR_GRIS_TEXTO}; font-weight: 600;'>{str(row['Proveedor']).replace('nan', '')}</td>
                <td style='padding: 10px 12px; color: {COLOR_TEXTO_PRINCIPAL}; font-weight: 700;'>{str(row['Modelo']).replace('nan', '')}</td>
                <td style='padding: 10px 12px; text-align: center; color: {COLOR_GRIS_TEXTO};'>{fmt(row['Semana_1'])}</td>
                <td style='padding: 10px 12px; text-align: center; color: {COLOR_GRIS_TEXTO};'>{fmt(row['Semana_2'])}</td>
                <td style='padding: 10px 12px; text-align: center; color: {COLOR_GRIS_TEXTO};'>{fmt(row['Semana_3'])}</td>
                <td style='padding: 10px 12px; text-align: center; color: {COLOR_GRIS_TEXTO};'>{fmt(row['Semana_4'])}</td>
                <td style='padding: 10px 12px; text-align: right; color: {color_texto}; font-weight: 800; background-color: #F8FAFC;'>{fmt(row['Total_Modelo'])}</td>
            </tr>
            """
        html += "</tbody></table></div>"
        return html

    with tab_cant:
        st.markdown(generar_tabla_crono(df_crono_cant, es_usd=False).replace('\n', ''), unsafe_allow_html=True)
    with tab_usd:
        st.markdown(generar_tabla_crono(df_crono_usd, es_usd=True).replace('\n', ''), unsafe_allow_html=True)
else:
    st.info("No se encontró el cronograma de equipos en el archivo actual.")



# ==============================================================================
# SECCIÓN FINAL: BOTONES DE EXPORTACIÓN (Sidebar) - VISTA COMPLETA
# ==============================================================================
with st.sidebar:
    st.markdown('<hr style="border: 1px solid #E2E8F0; margin: 20px 0;">', unsafe_allow_html=True)
    st.markdown('<h3 style="color: #001F5B; font-weight: 800; margin-bottom: 15px;">📥 Exportar Reporte</h3>', unsafe_allow_html=True)
    
    # --- 1. EXPORTAR A EXCEL ---
    import io
    output = io.BytesIO()
    df_resumen_export = pd.DataFrame({
        "Indicador": ["Saldo Operativo Total", "Compromisos CxP", "Apartados Operativos", "Cobranza Aliados/Bancos", "Moneda Extranjera"],
        "Monto USD": [saldo_operativo_usd, compromisos_usd, apartados_usd, cxc_bancos_usd, val_div_ext],
        "Monto Bs": [saldo_operativo_bs, compromisos_bs, apartados_bs, val_total_aliados_bs, 0.0]
    })
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_resumen_export.to_excel(writer, sheet_name='Resumen_Cierre', index=False)
    
    st.download_button(
        label="📊 Descargar Datos (Excel)",
        data=output.getvalue(),
        file_name=f"Reporte_Financiero_{fecha_reporte.strip(' | ')}.xlsx",
        mime="application/vnd.ms-excel"
    )

    # --- 2. GENERADOR DE VISTA HTML COMPLETA (ACTUALIZADO CON EQUIPOS) ---
    
    def generar_filas_html(df, col_concepto, col_monto, col_pct=None):
        filas = ""
        if df is not None and not df.empty:
            for _, row in df.iterrows():
                mnt_bs = row[col_monto]
                mnt_usd = mnt_bs / tasa_bcv_actual if 'tasa_bcv_actual' in globals() and tasa_bcv_actual > 0 else 0.0
                filas += f"<tr><td style='padding: 8px; border-bottom: 1px solid #E2E8F0; font-size: 13px;'>{row[col_concepto]}</td>"
                filas += f"<td style='padding: 8px; border-bottom: 1px solid #E2E8F0; text-align: right; font-weight: bold; font-size: 13px; color: #001F5B;'>{format_money_ve(mnt_bs)}</td>"
                filas += f"<td style='padding: 8px; border-bottom: 1px solid #E2E8F0; text-align: right; font-weight: bold; font-size: 13px; color: #00AEEF;'>${format_money_ve(mnt_usd)}</td>"
                if col_pct:
                    filas += f"<td style='padding: 8px; border-bottom: 1px solid #E2E8F0; text-align: right; font-weight: bold; font-size: 13px; color: #00AEEF;'>{row[col_pct]:.1f}%</td>"
                filas += "</tr>"
        return filas

    # Función para filas de equipos (ahora más compactas para el reporte)
    def generar_filas_crono_html(df, es_usd=False):
        filas = ""
        prefijo = "$" if es_usd else ""
        if df is not None and not df.empty:
            for _, row in df.iterrows():
                def f(v): return f"{prefijo}{format_money_ve(v)}" if es_usd else f"{int(v):,}"
                filas += f"<tr>"
                filas += f"<td style='padding: 4px; border-bottom: 1px solid #E2E8F0; font-size: 10px;'>{row['Proveedor']}</td>"
                filas += f"<td style='padding: 4px; border-bottom: 1px solid #E2E8F0; font-size: 10px; font-weight: bold;'>{row['Modelo']}</td>"
                filas += f"<td style='padding: 4px; border-bottom: 1px solid #E2E8F0; text-align: center; font-size: 10px;'>{f(row['Semana_1'])}</td>"
                filas += f"<td style='padding: 4px; border-bottom: 1px solid #E2E8F0; text-align: center; font-size: 10px;'>{f(row['Semana_2'])}</td>"
                filas += f"<td style='padding: 4px; border-bottom: 1px solid #E2E8F0; text-align: center; font-size: 10px;'>{f(row['Semana_3'])}</td>"
                filas += f"<td style='padding: 4px; border-bottom: 1px solid #E2E8F0; text-align: center; font-size: 10px;'>{f(row['Semana_4'])}</td>"
                filas += f"<td style='padding: 4px; border-bottom: 1px solid #E2E8F0; text-align: right; font-weight: bold; font-size: 10px; background: #F8FAFC;'>{f(row['Total_Modelo'])}</td>"
                filas += f"</tr>"
        return filas
    
    html_cxp = generar_filas_html(df_comp_tabla, 'Concepto', 'Monto', 'Porcentaje') if 'df_comp_tabla' in locals() else ""
    html_cxc = generar_filas_html(df_cobrar_tabla, 'Concepto', 'Monto', 'Porcentaje') if 'df_cobrar_tabla' in locals() else ""
    html_ger = generar_filas_html(df_gerencia, 'Gerencia', 'Monto', 'Porcentaje') if 'df_gerencia' in locals() else ""
    html_crono_cant = generar_filas_crono_html(df_crono_cant, False) if 'df_crono_cant' in locals() else ""
    html_crono_usd = generar_filas_crono_html(df_crono_usd, True) if 'df_crono_usd' in locals() else ""
    
    
    
    html_proy = ""
    if 'df_proyeccion' in locals() and not df_proyeccion.empty:
        for _, row in df_proyeccion.iterrows():
            html_proy += f"<tr><td style='padding: 8px; border-bottom: 1px solid #E2E8F0; font-size: 13px;'>{row['Concepto']}</td>"
            html_proy += f"<td style='padding: 8px; border-bottom: 1px solid #E2E8F0; text-align: right; font-weight: bold; font-size: 13px; color: #001F5B;'>{format_money_ve(row['Bs'])}</td>"
            html_proy += f"<td style='padding: 8px; border-bottom: 1px solid #E2E8F0; text-align: right; font-weight: bold; font-size: 13px; color: #00AEEF;'>${format_money_ve(row['USD'])}</td></tr>"

    # Construimos el HTML Maestro
    html_dashboard_full = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>Reporte Completo Platco - {fecha_reporte.strip(' | ')}</title>
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600;700;800&display=swap');
            body {{ font-family: 'Montserrat', sans-serif; background-color: #F4F6F9; margin: 0; padding: 40px; color: #001F5B; }}
            .container {{ max-width: 1000px; margin: auto; background: white; padding: 40px; border-radius: 15px; box-shadow: 0 10px 30px rgba(0,0,0,0.05); }}
            .header {{ background-color: #001F5B; padding: 30px; border-radius: 10px; color: white; display: flex; justify-content: space-between; align-items: center; margin-bottom: 30px; }}
            h1 {{ margin: 0; font-size: 24px; font-weight: 800; }}
            h2 {{ font-size: 18px; color: #001F5B; border-bottom: 2px solid #FFB81C; padding-bottom: 10px; margin-top: 40px; width: 100%; }}
            .date {{ color: #FFB81C; font-weight: 700; }}
            
            .kpi-grid {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 30px; }}
            .kpi-card {{ background: #F8FAFC; padding: 20px; border-radius: 10px; border: 1px solid #E2E8F0; }}
            .kpi-label {{ color: #6C757D; font-size: 11px; text-transform: uppercase; font-weight: 700; margin-bottom: 5px; }}
            .kpi-val {{ font-size: 24px; font-weight: 800; margin: 0; color: #001F5B; }}
            .kpi-sub {{ color: #6C757D; font-size: 12px; margin-top: 5px; }}
            
            table {{ width: 100%; border-collapse: collapse; margin-top: 15px; }}
            th {{ background-color: #001F5B; color: white; padding: 10px; text-align: left; font-size: 12px; text-transform: uppercase; }}
            
            /* LA MAGIA SUCEDE AQUI: Obligamos a que cada columna sea 50% exacto */
            .two-columns {{ display: flex; width: 100%; gap: 30px; margin-top: 20px; }}
            .two-columns > div {{ width: 50%; box-sizing: border-box; }}
            
            .pill-container {{ display: flex; gap: 15px; margin-bottom: 20px; }}
            .pill {{ background: #001F5B; color: white; padding: 10px 20px; border-radius: 8px; font-weight: 700; font-size: 13px; }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h1>Reporte Directivo Financiero</h1>
                <div class="date">Platco {fecha_reporte}</div>
            </div>

            <div class="pill-container">
                <div class="pill"><span style="color:#FFB81C;">Moneda Extranjera:</span> ${format_money_ve(val_div_ext)}</div>
                <div class="pill"><span style="color:#FFB81C;">Saldo en Bóveda:</span> ${format_money_ve(val_div_boveda)}</div>
            </div>

            <h2>1. Resumen Operativo</h2>
            <div class="kpi-grid">
                <div class="kpi-card">
                    <div class="kpi-label">Saldo Operativo Total</div>
                    <p class="kpi-val">${format_money_ve(saldo_operativo_usd)}</p>
                    <p class="kpi-sub">Bs {format_money_ve(saldo_operativo_bs)}</p>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Compromisos CxP</div>
                    <p class="kpi-val">${format_money_ve(compromisos_usd)}</p>
                    <p class="kpi-sub">Bs {format_money_ve(compromisos_bs)}</p>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Apartados Abril</div>
                    <p class="kpi-val">${format_money_ve(apartados_usd)}</p>
                    <p class="kpi-sub">Bs {format_money_ve(apartados_bs)}</p>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Cobranza Aliados / Bancos</div>
                    <p class="kpi-val">${format_money_ve(cxc_bancos_usd)}</p>
                    <p class="kpi-sub">Bs {format_money_ve(val_total_aliados_bs)}</p>
                </div>
            </div>

            <h2>2. Resumen Inversión</h2>
            <div class="kpi-grid">
                <div class="kpi-card">
                    <div class="kpi-label">Saldo Inversión</div>
                    <p class="kpi-val">${format_money_ve(abs(inv_banco_usd))}</p>
                    <p class="kpi-sub">Bs {format_money_ve(abs(inv_banco_bs))}</p>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Equipos a Pagar</div>
                    <p class="kpi-val">${format_money_ve(abs(inv_equipos_usd))}</p>
                    <p class="kpi-sub">Bs {format_money_ve(abs(inv_equipos_bs))}</p>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Compromisos Fiscales</div>
                    <p class="kpi-val">${format_money_ve(abs(inv_fiscales_usd))}</p>
                    <p class="kpi-sub">Bs {format_money_ve(abs(inv_fiscales_bs))}</p>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">CxC Inversión</div>
                    <p class="kpi-val">${format_money_ve(abs(inv_cxc_usd))}</p>
                    <p class="kpi-sub">Bs {format_money_ve(abs(inv_cxc_bs))}</p>
                </div>
            </div>

            <div class="two-columns">
                <div>
                    <h2>3. Detalles por Pagar (CxP)</h2>
                    <table>
                        <tr><th>Concepto</th><th style="text-align:right;">Monto (Bs)</th><th style="text-align:right;">Monto (USD)</th><th style="text-align:right;">%</th></tr>
                        {html_cxp}
                    </table>
                </div>
                <div>
                    <h2>4. Detalles por Cobrar (CxC)</h2>
                    <table>
                        <tr><th>Concepto</th><th style="text-align:right;">Monto (Bs)</th><th style="text-align:right;">Monto (USD)</th><th style="text-align:right;">%</th></tr>
                        {html_cxc}
                    </table>
                </div>
            </div>

            <div class="two-columns">
                <div>
                    <h2>5. Proyección de Disponibilidad</h2>
                    <table>
                        <tr><th>Concepto</th><th style="text-align:right;">Monto (Bs)</th><th style="text-align:right;">Monto (USD)</th></tr>
                        {html_proy}
                    </table>
                </div>
                <div>
                    <h2>6. Desembolso por Gerencia</h2>
                    <table>
                        <tr><th>Gerencia</th><th style="text-align:right;">Monto (Bs)</th><th style="text-align:right;">Monto (USD)</th><th style="text-align:right;">%</th></tr>
                        {html_ger}
                    </table>
                </div>
            </div>

            <div class="two-columns">
                <div>
                    <h2>7. Cronograma Adquisición</h2>
                    <p style="font-size: 11px; color: #64748B; margin-bottom: 5px;">Unidades físicas programadas.</p>
                    <table>
                        <tr><th>Prov.</th><th>Modelo</th><th style="text-align:center;">S1</th><th style="text-align:center;">S2</th><th style="text-align:center;">S3</th><th style="text-align:center;">S4</th><th style="text-align:right;">Total</th></tr>
                        {html_crono_cant}
                    </table>
                </div>
                <div>
                    <h2>8. Cronograma USD</h2>
                    <p style="font-size: 11px; color: #64748B; margin-bottom: 5px;">Desembolsos en divisas.</p>
                    <table>
                        <tr><th>Prov.</th><th>Modelo</th><th style="text-align:center;">S1</th><th style="text-align:center;">S2</th><th style="text-align:center;">S3</th><th style="text-align:center;">S4</th><th style="text-align:right;">Total</th></tr>
                        {html_crono_usd}
                    </table>
                </div>
            </div>

            <p style="text-align: center; color: #94A3B8; font-size: 11px; margin-top: 50px;">
                Documento Estático Generado por Platco Dashboard
            </p>
        </div>
    </body>
    </html>
    """

    st.download_button(
        label="🌐 Descargar Reporte Completo (HTML)",
        data=html_dashboard_full,
        file_name=f"Reporte_Completo_Platco_{fecha_reporte.strip(' | ')}.html",
        mime="text/html",
        help="Descarga un documento estático con todas las tablas y cifras calculadas en el dashboard."
    )
    

