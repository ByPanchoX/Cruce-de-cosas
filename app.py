import streamlit as st
import pandas as pd
import io
import unicodedata
import re
import zipfile

st.set_page_config(page_title="Consolidador Remuneraciones Inteligente", layout="wide")

if 'datos_acumulados' not in st.session_state:
    st.session_state['datos_acumulados'] = []

def clean_rut(rut):
    if pd.isna(rut): return None
    rut = str(rut).upper()
    return "".join(filter(lambda x: x.isdigit() or x == 'K', rut))

def parse_informe(df):
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
            data.append({'Rut': clean_rut(val_2), 'Concepto': current_category, 'Monto': val_10, 'Seccion': current_section})
    return pd.DataFrame(data)

def procesar_par(file_lib, file_inf, rut_empresa, mes_nombre, anio):
    df_lib = pd.read_excel(file_lib)
    df_inf_raw = pd.read_excel(file_inf)
    
    # 1. Filtro del Libro (Identificación, Días, Imponible y Seguro Cesantía del trabajador)
    id_cols = ['Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres']
    conteos = [c for c in df_lib.columns if re.search(r'^[Nn][°º\.\s]', str(c))]
    
    # Columnas necesarias para el cálculo automático
    cols_calculo = []
    if 'Imponible' in df_lib.columns: cols_calculo.append('Imponible')
    if 'Seguro de Cesantía' in df_lib.columns: cols_calculo.append('Seguro de Cesantía')
    
    df_lib_clean = df_lib[id_cols + conteos + cols_calculo].copy()
    
    # --- DETECCIÓN AUTOMÁTICA DE TASA ---
    def calcular_seguro_empleador(row):
        imponible = row.get('Imponible', 0)
        seguro_trabajador = row.get('Seguro de Cesantía', 0)
        
        if imponible <= 0:
            return 0
        
        # Si el trabajador aporta (Seguro > 0), es Indefinido -> Empleador paga 2.4%
        if seguro_trabajador > 0:
            return round(imponible * 0.024)
        # Si el trabajador no aporta, es Plazo Fijo/Obra -> Empleador paga 3.0%
        else:
            return round(imponible * 0.030)

    if 'Imponible' in df_lib_clean.columns:
        df_lib_clean['Total Imponible (H)'] = df_lib_clean['Imponible']
        df_lib_clean['Seguro Cesantía Empleador (D)'] = df_lib_clean.apply(calcular_seguro_empleador, axis=1)
        # Borramos las columnas originales del libro para que no ensucien
        df_lib_clean = df_lib_clean.drop(columns=['Imponible', 'Seguro de Cesantía'], errors='ignore')

    # 2. Procesar Informe (Haberes y Descuentos desglosados)
    df_parsed = parse_informe(df_inf_raw)
    df_parsed['Concepto_Final'] = df_parsed.apply(lambda x: f"{x['Concepto']} (H)" if x['Seccion'] == 'HABERES' else f"{x['Concepto']} (D)", axis=1)
    df_pivot = df_parsed.pivot_table(index='Rut', columns='Concepto_Final', values='Monto', aggfunc='sum').reset_index()

    # 3. Unión
    rut_c = [c for c in df_lib_clean.columns if 'Rut' in str(c)][0]
    df_lib_clean['rut_key'] = df_lib_clean[rut_c].apply(clean_rut)
    df_merged = pd.merge(df_lib_clean, df_pivot, left_on='rut_key', right_on='Rut', how='left').drop(columns=['rut_key', 'Rut'])
    
    df_merged.insert(0, 'Año', anio)
    df_merged.insert(0, 'Mes', mes_nombre)
    df_merged.insert(0, 'Empresa', rut_empresa)
    
    # Ordenar columnas
    ids = ['Empresa', 'Mes', 'Año', 'Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres']
    d_n = [c for c in df_merged.columns if re.search(r'^[Nn][°º\.\s]', str(c))]
    h_f = sorted([c for c in df_merged.columns if '(H)' in c])
    d_f = sorted([c for c in df_merged.columns if '(D)' in c])
    
    return df_merged[ids + d_n + h_f + d_f]

# --- INTERFAZ ---
st.title("📊 Consolidador con Detección Automática de Contratos")

with st.sidebar:
    st.header("Parámetros Generales")
    rut_unico = st.text_input("RUT Empresa para el Excel", value="76.455.680-1")
    anio_global = st.number_input("Año", 2025)
    if st.button("🗑️ Limpiar Memoria"):
        st.session_state['datos_acumulados'] = []
        st.rerun()

st.info("ℹ️ El sistema detecta el tipo de contrato analizando si el trabajador tiene descuento de Seguro de Cesantía. (Indefinido: 2.4% / Plazo Fijo: 3.0%)")

col1, col2 = st.columns(2)
with col1:
    files_libros = st.file_uploader("Subir LIBROS (.xlsx)", type=["xlsx"], accept_multiple_files=True)
with col2:
    files_informes = st.file_uploader("Subir INFORMES (.xlsx)", type=["xlsx"], accept_multiple_files=True)

if files_libros and files_informes:
    if st.button("🚀 Procesar Todo"):
        lib_dict = {f.name: f for f in files_libros}
        inf_dict = {f.name: f for f in files_informes}
        meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
        
        for mes in meses:
            match_l = next((f for n, f in lib_dict.items() if mes.lower() in n.lower()), None)
            match_i = next((f for n, f in inf_dict.items() if mes.lower() in n.lower()), None)
            
            if match_l and match_i:
                try:
                    df_res = procesar_par(match_l, match_i, rut_unico, mes, anio_global)
                    st.session_state['datos_acumulados'].append({'mes': mes, 'df': df_res})
                    st.success(f"✅ Procesado: {mes}")
                except Exception as e:
                    st.error(f"❌ Error en {mes}: {e}")

# --- DESCARGAS ---
if st.session_state['datos_acumulados']:
    st.divider()
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED, False) as zip_f:
        for item in st.session_state['datos_acumulados']:
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as writer:
                item['df'].to_excel(writer, index=False)
            zip_f.writestr(f"Consolidado_{item['mes']}.xlsx", out.getvalue())
            
            # Botón de descarga individual
            st.download_button(f"📥 Descargar {item['mes']}", out.getvalue(), f"Consolidado_{item['mes']}.xlsx", key=f"dl_{item['mes']}")

    st.divider()
    st.download_button("📦 DESCARGAR TODO (ZIP)", zip_buf.getvalue(), "Remuneraciones_Consolidado.zip", "application/zip")
