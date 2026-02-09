import streamlit as st
import pandas as pd
import io
import unicodedata
import re
import zipfile

st.set_page_config(page_title="Consolidador Remuneraciones con Topes", layout="wide")

# TABLA DE UF (Al último día de cada mes)
VALORES_UF = {
    2024: {
        "Enero": 36863.81, "Febrero": 36929.54, "Marzo": 37046.22, "Abril": 37178.50,
        "Mayo": 37297.64, "Junio": 37422.58, "Julio": 37562.81, "Agosto": 37645.45,
        "Septiembre": 37733.62, "Octubre": 37844.21, "Noviembre": 37967.50, "Diciembre": 38103.39
    },
    2025: {
        "Enero": 38284.86, "Febrero": 38647.94, "Marzo": 38800.00, "Abril": 39050.00,
        "Mayo": 39200.00, "Junio": 39267.07, "Julio": 39179.01, "Agosto": 39383.07,
        "Septiembre": 39500.00, "Octubre": 39597.67, "Noviembre": 39643.59, "Diciembre": 39727.96
    }
}

# TOPES EN UF
TOPES_UF = {
    2024: {"prevision": 84.3, "cesantia": 126.6},
    2025: {"prevision": 87.8, "cesantia": 131.8}
}

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
    
    id_cols = ['Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres']
    conteos = [c for c in df_lib.columns if re.search(r'^[Nn][°º\.\s]', str(c))]
    
    cols_calculo = []
    if 'Imponible' in df_lib.columns: cols_calculo.append('Imponible')
    if 'Seguro de Cesantía' in df_lib.columns: cols_calculo.append('Seguro de Cesantía')
    
    df_lib_clean = df_lib[id_cols + conteos + cols_calculo].copy()
    
    # Obtener UF y Topes del periodo
    valor_uf = VALORES_UF.get(anio, {}).get(mes_nombre, 38500.0) # Valor por defecto si no existe
    tope_afc_uf = TOPES_UF.get(anio, TOPES_UF[2025])["cesantia"]
    tope_afc_pesos = valor_uf * tope_afc_uf

    if 'Imponible' in df_lib_clean.columns:
        df_lib_clean['Total Imponible (H)'] = df_lib_clean['Imponible']
        
        def calc_seguro_empleador(row):
            imp = row.get('Imponible', 0)
            seg_trab = row.get('Seguro de Cesantía', 0)
            if imp <= 0: return 0
            
            # APLICAR TOPE: El cálculo se hace sobre el menor entre el sueldo y el tope
            base_calculo = min(imp, tope_afc_pesos)
            
            # Tasa: 2.4% (Indefinido) o 3.0% (Plazo Fijo)
            tasa = 0.024 if seg_trab > 0 else 0.030
            return round(base_calculo * tasa)
            
        df_lib_clean['Seguro Cesantía Empleador (D)'] = df_lib_clean.apply(calc_seguro_empleador, axis=1)
        df_lib_clean = df_lib_clean.drop(columns=['Imponible', 'Seguro de Cesantía'], errors='ignore')

    df_parsed = parse_informe(df_inf_raw)
    df_parsed['Concepto_Final'] = df_parsed.apply(lambda x: f"{x['Concepto']} (H)" if x['Seccion'] == 'HABERES' else f"{x['Concepto']} (D)", axis=1)
    df_pivot = df_parsed.pivot_table(index='Rut', columns='Concepto_Final', values='Monto', aggfunc='sum').reset_index()

    rut_c = [c for c in df_lib_clean.columns if 'Rut' in str(c)][0]
    df_lib_clean['rut_key'] = df_lib_clean[rut_c].apply(clean_rut)
    df_merged = pd.merge(df_lib_clean, df_pivot, left_on='rut_key', right_on='Rut', how='left').drop(columns=['rut_key', 'Rut'])
    
    df_merged.insert(0, 'Año', anio)
    df_merged.insert(0, 'Mes', mes_nombre)
    df_merged.insert(0, 'Empresa', rut_empresa)
    
    return df_merged

# --- INTERFAZ ---
st.title("📊 Procesador Masivo con Topes Imponibles")

with st.sidebar:
    st.header("1. Datos Globales")
    rut_unico = st.text_input("RUT Empresa para el Excel", value="76.455.680-1")
    anio_global = st.number_input("Año", 2025)
    if st.button("🗑️ Limpiar Memoria"):
        st.session_state['datos_acumulados'] = []
        st.rerun()

st.info(f"✅ Los cálculos de 2.4% y 3.0% ahora consideran automáticamente el **Tope Imponible de Cesantía** del periodo.")

col1, col2 = st.columns(2)
with col1:
    files_libros = st.file_uploader("Subir LIBROS", type=["xlsx"], accept_multiple_files=True)
with col2:
    files_informes = st.file_uploader("Subir INFORMES", type=["xlsx"], accept_multiple_files=True)

if files_libros and files_informes:
    if st.button("🚀 Procesar Todo"):
        lib_dict = {f.name: f for f in files_libros}
        inf_dict = {f.name: f for f in files_informes}
        meses = list(VALORES_UF[2024].keys())
        
        for mes in meses:
            match_l = next((f for n, f in lib_dict.items() if mes.lower() in n.lower()), None)
            match_i = next((f for n, f in inf_map.items() if mes.lower() in n.lower()), None) if 'inf_map' in locals() else next((f for n, f in inf_dict.items() if mes.lower() in n.lower()), None)
            
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
            mes = item['mes']
            df = item['df']
            
            # Ordenar columnas
            ids = ['Empresa', 'Mes', 'Año', 'Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres']
            dias_n = [c for c in df.columns if re.search(r'^[Nn][°º\.\s]', str(c))]
            habs = sorted([c for c in df.columns if '(H)' in c])
            dess = sorted([c for c in df.columns if '(D)' in c])
            resto = [c for c in df.columns if c not in ids + dias_n + habs + dess]
            df = df[ids + dias_n + habs + dess + resto]
            
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            zip_f.writestr(f"Consolidado_{mes}_{anio_global}.xlsx", out.getvalue())
            st.download_button(f"📥 Descargar {mes}", out.getvalue(), f"Consolidado_{mes}.xlsx", key=f"dl_{mes}")

    st.download_button("📦 DESCARGAR TODO (ZIP)", zip_buf.getvalue(), "Remuneraciones_Final.zip", "application/zip")
