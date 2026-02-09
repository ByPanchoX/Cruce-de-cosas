import streamlit as st
import pandas as pd
import io
import unicodedata
import re

st.set_page_config(page_title="Consolidador Remuneraciones (H/D)", layout="wide")

# Inicializar memoria para acumular empresas del mismo mes
if 'datos_mes_actual' not in st.session_state:
    st.session_state['datos_mes_actual'] = []

def clean_rut(rut):
    if pd.isna(rut): return None
    rut = str(rut).upper()
    return "".join(filter(lambda x: x.isdigit() or x == 'K', rut))

def parse_informe(df):
    """Extrae conceptos del informe clasificando en HABERES y DESCUENTOS"""
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

def procesar_par(file_lib, file_inf, rut_empresa, mes_nombre, anio):
    """Lógica de procesamiento con cálculo de Seguro Cesantía e Imponible"""
    df_lib = pd.read_excel(file_lib)
    df_inf_raw = pd.read_excel(file_inf)
    
    # 1. Filtro del Libro (ID + Conteos N° + Imponible + Seguro para cálculo)
    id_cols = ['Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres']
    # Buscamos columnas de conteo (N°, Nº, N.)
    conteos = [c for c in df_lib.columns if re.search(r'^[Nn][°º\.\s]', str(c))]
    
    # Columnas necesarias para cálculos internos
    cols_calculo = []
    if 'Imponible' in df_lib.columns: cols_calculo.append('Imponible')
    if 'Seguro de Cesantía' in df_lib.columns: cols_calculo.append('Seguro de Cesantía')
    
    df_lib_clean = df_lib[id_cols + conteos + cols_calculo].copy()
    
    # --- CÁLCULO DE IMPONIBLE Y SEGURO CESANTÍA EMPLEADOR ---
    if 'Imponible' in df_lib_clean.columns:
        df_lib_clean['Total Imponible (H)'] = df_lib_clean['Imponible']
        
        # Detección automática: Si trabajador aporta (>0) es 2.4%, si no (0) es 3.0%
        def calc_cesantia_emp(row):
            imp = row.get('Imponible', 0)
            seg_trab = row.get('Seguro de Cesantía', 0)
            if imp <= 0: return 0
            tasa = 0.024 if seg_trab > 0 else 0.030
            return round(imp * tasa)
            
        df_lib_clean['Seguro Cesantía Empleador (D)'] = df_lib_clean.apply(calc_cesantia_emp, axis=1)
        # Quitar columnas originales del libro para evitar duplicados
        df_lib_clean = df_lib_clean.drop(columns=['Imponible', 'Seguro de Cesantía'], errors='ignore')

    # 2. Procesar Informe (Etiquetado H/D)
    df_parsed = parse_informe(df_inf_raw)
    df_parsed['Concepto_Final'] = df_parsed.apply(lambda x: f"{x['Concepto']} (H)" if x['Seccion'] == 'HABERES' else f"{x['Concepto']} (D)", axis=1)
    df_pivot = df_parsed.pivot_table(index='Rut', columns='Concepto_Final', values='Monto', aggfunc='sum').reset_index()

    # 3. Unión
    rut_c = [c for c in df_lib_clean.columns if 'Rut' in str(c)][0]
    df_lib_clean['rut_key'] = df_lib_clean[rut_c].apply(clean_rut)
    df_merged = pd.merge(df_lib_clean, df_pivot, left_on='rut_key', right_on='Rut', how='left').drop(columns=['rut_key', 'Rut'])
    
    # Metadatos
    df_merged.insert(0, 'Año', anio)
    df_merged.insert(0, 'Mes', mes_nombre)
    df_merged.insert(0, 'Empresa', rut_empresa)
    
    return df_merged

# --- INTERFAZ ---
st.title("📊 Consolidador de Remuneraciones: Carga por Empresa")

with st.sidebar:
    st.header("Configuración Global")
    rut_unico = st.text_input("RUT Empresa para el Excel", value="76.455.680-1")
    anio_global = st.number_input("Año", 2025)
    mes_global = st.selectbox("Mes de Proceso", ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"])
    
    st.divider()
    if st.button("🗑️ Reiniciar Mes Actual"):
        st.session_state['datos_mes_actual'] = []
        st.rerun()

st.info(f"📅 Procesando: **{mes_global} {anio_global}**. Puedes subir los archivos de ambas empresas por separado y se juntarán en el archivo final.")

# Carga de archivos
col1, col2 = st.columns(2)
with col1:
    file_lib = st.file_uploader("1. Subir Libro de Remuneraciones (.xlsx)", type=["xlsx"])
with col2:
    file_inf = st.file_uploader("2. Subir Informe Haberes y Descuentos (.xlsx)", type=["xlsx"])

if file_lib and file_inf:
    if st.button("➕ Procesar y Agregar esta Empresa"):
        try:
            df_res = procesar_par(file_lib, file_inf, rut_unico, mes_global, anio_global)
            st.session_state['datos_mes_actual'].append(df_res)
            st.success("✅ Empresa agregada a la lista. Puedes subir la siguiente o generar el archivo.")
        except Exception as e:
            st.error(f"❌ Error: {e}")

# Sección de descarga
if st.session_state['datos_mes_actual']:
    st.divider()
    st.subheader(f"📋 Resumen {mes_global}")
    
    # Consolidar todo lo cargado para el mes
    df_final = pd.concat(st.session_state['datos_mes_actual'], ignore_index=True)
    
    # Limpieza final de columnas y ordenamiento
    df_final = df_final.loc[:, ~df_final.columns.duplicated()]
    
    ids = ['Empresa', 'Mes', 'Año', 'Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres']
    dias_n = [c for c in df_final.columns if re.search(r'^[Nn][°º\.\s]', str(c))]
    habs = sorted([c for c in df_final.columns if '(H)' in c])
    dess = sorted([c for c in df_final.columns if '(D)' in c])
    resto = sorted([c for c in df_final.columns if c not in ids + dias_n + habs + dess])
    
    df_final = df_final[ids + dias_n + habs + dess + resto]
    
    st.write(f"Total trabajadores cargados: **{len(df_final)}**")
    st.dataframe(df_final.head(10))

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df_final.to_excel(writer, index=False)
    
    st.download_button(
        label=f"📥 Descargar Consolidado {mes_global} (Ambas Empresas)",
        data=buf.getvalue(),
        file_name=f"Consolidado_{mes_global}_{anio_global}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
