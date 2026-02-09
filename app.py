import streamlit as st
import pandas as pd
import io
import unicodedata
import re
import zipfile

st.set_page_config(page_title="Consolidador Masivo Pro (H/D)", layout="wide")

if 'datos_acumulados' not in st.session_state:
    st.session_state['datos_acumulados'] = []

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
    """Lógica de procesamiento de un mes individual con cálculos automáticos"""
    df_lib = pd.read_excel(file_lib)
    df_inf_raw = pd.read_excel(file_inf)
    
    # 1. Filtro del Libro (Solo identificación, días N° e Imponible)
    id_cols = ['Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres']
    conteos = [c for c in df_lib.columns if re.search(r'^[Nn][°º\.\s]', str(c))]
    
    # Columnas necesarias para el cálculo de Seguro de Cesantía
    cols_calculo = []
    if 'Imponible' in df_lib.columns: cols_calculo.append('Imponible')
    if 'Seguro de Cesantía' in df_lib.columns: cols_calculo.append('Seguro de Cesantía')
    
    df_lib_clean = df_lib[id_cols + conteos + cols_calculo].copy()
    
    # Limpieza: eliminar columnas de días que no tengan el prefijo N° (evita duplicados con el informe)
    drop_c = [c for c in df_lib_clean.columns if ('dia' in str(c).lower() or 'día' in str(c).lower()) and not re.search(r'^[Nn][°º\.\s]', str(c)) and c not in cols_calculo]
    df_lib_clean = df_lib_clean.drop(columns=drop_c)
    
    # --- CÁLCULO DE IMPONIBLE Y SEGURO CESANTÍA EMPLEADOR ---
    if 'Imponible' in df_lib_clean.columns:
        df_lib_clean['Total Imponible (H)'] = df_lib_clean['Imponible']
        
        # Detección automática: Si trabajador aporta (>0) es 2.4%, si no (0) es 3.0%
        def calc_cesantia(row):
            imp = row.get('Imponible', 0)
            seg_trab = row.get('Seguro de Cesantía', 0)
            if imp <= 0: return 0
            tasa = 0.024 if seg_trab > 0 else 0.030
            return round(imp * tasa)
            
        df_lib_clean['Seguro Cesantía Empleador (D)'] = df_lib_clean.apply(calc_cesantia, axis=1)
        # Quitar originales para no duplicar
        df_lib_clean = df_lib_clean.drop(columns=['Imponible', 'Seguro de Cesantía'], errors='ignore')

    # 2. Procesar Informe (Etiquetado H/D)
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
    dias_n = [c for c in df_merged.columns if re.search(r'^[Nn][°º\.\s]', str(c))]
    haberes = sorted([c for c in df_merged.columns if '(H)' in c])
    descuentos = sorted([c for c in df_merged.columns if '(D)' in c])
    resto = sorted([c for c in df_merged.columns if c not in ids + dias_n + haberes + descuentos])
    
    return df_merged[ids + dias_n + haberes + descuentos + resto]

# --- INTERFAZ ---
st.title("📊 Procesador Masivo de Remuneraciones")

with st.sidebar:
    st.header("1. Datos del Excel Final")
    rut_unico = st.text_input("RUT Empresa para el Excel", value="76.455.680-1")
    anio_global = st.number_input("Año", 2025)
    if st.button("🗑️ Limpiar Todo"):
        st.session_state['datos_acumulados'] = []
        st.rerun()

st.markdown("### 📥 Carga Masiva de Archivos")
st.info("Sube Libros e Informes. El sistema detectará automáticamente el tipo de contrato para el Seguro de Cesantía (2.4% o 3.0%).")

col1, col2 = st.columns(2)
with col1:
    files_libros = st.file_uploader("Subir todos los LIBROS (.xlsx)", type=["xlsx"], accept_multiple_files=True)
with col2:
    files_informes = st.file_uploader("Subir todos los INFORMES (.xlsx)", type=["xlsx"], accept_multiple_files=True)

if files_libros and files_informes:
    if st.button("🚀 Procesar Todos los Archivos"):
        libros_dict = {f.name: f for f in files_libros}
        informes_dict = {f.name: f for f in files_informes}
        
        meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
        
        for mes in meses:
            libro_match = next((f for name, f in libros_dict.items() if mes.lower() in name.lower()), None)
            informe_match = next((f for name, f in informes_dict.items() if mes.lower() in name.lower()), None)
            
            if libro_match and informe_match:
                try:
                    df_res = procesar_par(libro_match, informe_match, rut_unico, mes, anio_global)
                    st.session_state['datos_acumulados'].append({'mes': mes, 'df': df_res})
                    st.success(f"✅ Procesado: {mes}")
                except Exception as e:
                    st.error(f"❌ Error en {mes}: {e}")

# --- SECCIÓN DE DESCARGAS ---
if st.session_state['datos_acumulados']:
    st.divider()
    st.subheader("📦 Archivos Procesados")
    
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        for item in st.session_state['datos_acumulados']:
            mes = item['mes']
            df = item['df']
            
            c1, c2 = st.columns([3, 1])
            c1.write(f"📁 Consolidado_{mes}_{anio_global}.xlsx")
            
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            
            c2.download_button(
                label=f"Descargar {mes}",
                data=buf.getvalue(),
                file_name=f"Consolidado_{mes}_{anio_global}.xlsx",
                key=f"dl_{mes}"
            )
            zip_file.writestr(f"Consolidado_{mes}_{anio_global}.xlsx", buf.getvalue())

    st.divider()
    st.download_button(
        label="📥 DESCARGAR TODOS LOS MESES (ZIP)",
        data=zip_buffer.getvalue(),
        file_name=f"Remuneraciones_{anio_global}_Completo.zip",
        mime="application/zip"
    )
    
    if st.button("📄 Generar Consolidado Total (Todos los meses en una tabla)"):
        df_total = pd.concat([item['df'] for item in st.session_state['datos_acumulados']], ignore_index=True)
        buf_total = io.BytesIO()
        with pd.ExcelWriter(buf_total, engine='openpyxl') as writer:
            df_total.to_excel(writer, index=False)
        st.download_button("Descargar Tabla Única Total", buf_total.getvalue(), "Consolidado_Total_Anual.xlsx")
