import streamlit as st
import pandas as pd
import io
import unicodedata
import re
import zipfile

st.set_page_config(page_title="Consolidador Remuneraciones Pro", layout="wide")

# Inicializar memoria de la aplicación
if 'datos_acumulados' not in st.session_state:
    st.session_state['datos_acumulados'] = []

def normalizar_extremo(texto):
    """Limpia tildes, espacios, signos y plurales para comparar nombres de columnas."""
    if pd.isna(texto): return ""
    texto = str(texto).lower().strip()
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    texto = re.sub(r'[^a-z0-9]', '', texto)
    if texto.endswith('s'): texto = texto[:-1] # Trata 'dias' y 'dia' como igual
    return texto

def clean_rut(rut):
    if pd.isna(rut): return None
    rut = str(rut).upper()
    return "".join(filter(lambda x: x.isdigit() or x == 'K', rut))

def parse_informe(df):
    """Extrae conceptos del informe detallado clasificando en HABERES y DESCUENTOS."""
    data = []
    current_section = 'HABERES'
    current_category = None
    for i in range(len(df)):
        row = df.iloc[i]
        val_0 = str(row.iloc[0]) if pd.notna(row.iloc[0]) else None
        if val_0 == 'HABERES': current_section = 'HABERES'
        elif val_0 == 'DESCUENTOS': current_section = 'DESCUENTOS'
        
        val_2 = str(row.iloc[2]) if pd.notna(row.iloc[2]) else None
        val_10 = row.iloc[10]
        
        if val_0 and not val_0.startswith('Total') and val_0 not in ['Rut', 'Nombre', 'Razón Social', 'R.U.T.', 'Dirección', 'HABERES', 'DESCUENTOS']:
             current_category = val_0
        
        if val_2 and pd.notna(val_10) and '-' in val_2:
            data.append({
                'Rut': clean_rut(val_2),
                'Concepto': current_category,
                'Monto': val_10,
                'Seccion': current_section
            })
    return pd.DataFrame(data)

def procesar_mes(file_lib, file_inf, rut_empresa, mes_nombre, anio):
    """Procesa un par de archivos aplicando toda la lógica de negocio."""
    df_lib = pd.read_excel(file_lib)
    df_inf_raw = pd.read_excel(file_inf)
    
    # 1. Filtro estricto del Libro (Nombres y Conteos N°)
    id_cols = ['Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres']
    conteos = [c for c in df_lib.columns if re.search(r'^[Nn][°º\.\s]', str(c))]
    
    # Columnas para cálculo de Seguro de Cesantía
    cols_lib = id_cols + conteos
    if 'Imponible' in df_lib.columns: cols_lib.append('Imponible')
    if 'Seguro de Cesantía' in df_lib.columns: cols_lib.append('Seguro de Cesantía')
    
    df_lib_clean = df_lib[[c for c in cols_lib if c in df_lib.columns]].copy()
    
    # 2. Lógica de Seguro de Cesantía e Imponible
    if 'Imponible' in df_lib_clean.columns:
        df_lib_clean['Total Imponible (H)'] = df_lib_clean['Imponible']
        # Detección de contrato: Si tiene descuento de trabajador (>0) es 2.4%, sino 3.0%
        df_lib_clean['Seguro Cesantía Empleador (D)'] = df_lib_clean.apply(
            lambda x: round(x['Imponible'] * 0.024) if x.get('Seguro de Cesantía', 0) > 0 else round(x['Imponible'] * 0.030), axis=1
        )
        df_lib_clean = df_lib_clean.drop(columns=['Imponible', 'Seguro de Cesantía'], errors='ignore')

    # 3. Procesar Informe y Etiquetar (H)/(D)
    df_parsed = parse_informe(df_inf_raw)
    df_parsed['Concepto_Final'] = df_parsed.apply(
        lambda x: f"{x['Concepto']} (H)" if x['Seccion'] == 'HABERES' else f"{x['Concepto']} (D)", axis=1
    )
    df_pivot = df_parsed.pivot_table(index='Rut', columns='Concepto_Final', values='Monto', aggfunc='sum').reset_index()

    # 4. Unión y Limpieza de duplicados
    rut_col_lib = [c for c in df_lib_clean.columns if 'Rut' in str(c)][0]
    df_lib_clean['rut_key'] = df_lib_clean[rut_col_lib].apply(clean_rut)
    df_final = pd.merge(df_lib_clean, df_pivot, left_on='rut_key', right_on='Rut', how='left').drop(columns=['rut_key', 'Rut'])
    
    # Insertar metadatos
    df_final.insert(0, 'Año', anio)
    df_final.insert(0, 'Mes', mes_nombre)
    df_final.insert(0, 'Empresa', rut_empresa)
    
    # Ordenar columnas lógicamente
    ids = ['Empresa', 'Mes', 'Año', 'Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres']
    d_n = [c for c in df_final.columns if re.search(r'^[Nn][°º\.\s]', str(c))]
    habs = sorted([c for c in df_final.columns if '(H)' in c])
    dess = sorted([c for c in df_final.columns if '(D)' in c])
    resto = [c for c in df_final.columns if c not in ids + d_n + habs + dess]
    
    return df_final[ids + d_n + habs + dess + resto]

# --- INTERFAZ STREAMLIT ---
st.title("🚀 Consolidador de Remuneraciones Definitivo")

with st.sidebar:
    st.header("1. Configuración del Excel")
    rut_global = st.text_input("RUT Empresa para el archivo", value="76.455.680-1")
    anio_global = st.number_input("Año de los datos", 2025)
    st.divider()
    if st.button("🗑️ Limpiar Memoria"):
        st.session_state['datos_acumulados'] = []
        st.rerun()

st.markdown("### 📥 Carga de Archivos")
st.info("Sube todos tus archivos. El sistema los emparejará si el nombre del archivo contiene el mes (ej: 'Libro Enero.xlsx' e 'Informe Enero.xlsx').")

col1, col2 = st.columns(2)
with col1:
    files_l = st.file_uploader("Subir LIBROS DE REMUNERACIONES", type=["xlsx"], accept_multiple_files=True)
with col2:
    files_i = st.file_uploader("Subir INFORMES HABERES/DESCUENTOS", type=["xlsx"], accept_multiple_files=True)

if files_l and files_i:
    if st.button("🔄 Procesar y Consolidar"):
        lib_map = {f.name: f for f in files_l}
        inf_map = {f.name: f for f in files_i}
        meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
        
        for mes in meses:
            m_l = next((f for n, f in lib_map.items() if mes.lower() in n.lower()), None)
            m_i = next((f for n, f in inf_map.items() if mes.lower() in n.lower()), None)
            
            if m_l and m_i:
                try:
                    df_res = procesar_mes(m_l, m_i, rut_global, mes, anio_global)
                    st.session_state['datos_acumulados'].append({'mes': mes, 'df': df_res})
                    st.success(f"✅ Mes procesado: {mes}")
                except Exception as e:
                    st.error(f"❌ Error procesando {mes}: {e}")

# --- DESCARGAS ---
if st.session_state['datos_acumulados']:
    st.divider()
    st.subheader("📦 Descargas Disponibles")
    
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED, False) as zip_f:
        for item in st.session_state['datos_acumulados']:
            mes = item['mes']
            df = item['df']
            
            # Excel Individual
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            
            c1, c2 = st.columns([3, 1])
            c1.write(f"📁 Consolidado_{mes}_{anio_global}.xlsx")
            c2.download_button(f"Descargar {mes}", out.getvalue(), f"Consolidado_{mes}_{anio_global}.xlsx", key=f"dl_{mes}")
            
            zip_f.writestr(f"Consolidado_{mes}_{anio_global}.xlsx", out.getvalue())

    st.divider()
    st.download_button("📥 DESCARGAR TODOS LOS MESES (ZIP)", zip_buf.getvalue(), f"Remuneraciones_{anio_global}.zip", "application/zip")
    
    if st.button("📄 Unir todos los meses en una Gran Tabla"):
        df_grande = pd.concat([i['df'] for i in st.session_state['datos_acumulados']], ignore_index=True)
        # Limpieza final de columnas repetidas que puedan venir de distintos meses
        df_grande = df_grande.loc[:, ~df_grande.columns.duplicated()]
        
        out_g = io.BytesIO()
        with pd.ExcelWriter(out_g, engine='openpyxl') as writer:
            df_grande.to_excel(writer, index=False)
        st.download_button("Descargar Tabla Anual Única", out_g.getvalue(), "Consolidado_Anual_Completo.xlsx")
