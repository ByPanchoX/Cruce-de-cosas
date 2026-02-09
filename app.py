import streamlit as st
import pandas as pd
import io
import re

# Configuración avanzada de la interfaz
st.set_page_config(page_title="Consolidador Pro 2025", layout="wide", page_icon="📈")

# --- BASE DE DATOS LEGAL 2025 ---
TOPES_AFC_UF_2025 = {
    "Enero": 131.8, "Febrero": 131.9, "Marzo": 131.9, "Abril": 131.9,
    "Mayo": 131.9, "Junio": 131.9, "Julio": 131.9, "Agosto": 131.9,
    "Septiembre": 131.9, "Octubre": 131.9, "Noviembre": 131.9, "Diciembre": 131.9
}

VALORES_UF_2025 = {
    "Enero": 38284.86, "Febrero": 38647.94, "Marzo": 38800.00, "Abril": 39050.00,
    "Mayo": 39200.00, "Junio": 39267.07, "Julio": 39179.01, "Agosto": 39383.07,
    "Septiembre": 39500.00, "Octubre": 39597.67, "Noviembre": 39643.59, "Diciembre": 39727.96
}

# --- FUNCIONES DE PROCESAMIENTO ---
def clean_rut(rut):
    if pd.isna(rut): return None
    return "".join(filter(lambda x: x.isdigit() or x == 'K', str(rut).upper()))

def procesar_informe_detallado(df):
    """Transforma el informe vertical en columnas horizontales (H)/(D)"""
    data = []
    seccion_actual = 'HABERES'
    categoria_actual = None
    
    for _, row in df.iterrows():
        col0 = str(row.iloc[0]) if pd.notna(row.iloc[0]) else ""
        if 'HABERES' in col0.upper(): seccion_actual = 'HABERES'
        elif 'DESCUENTOS' in col0.upper(): seccion_actual = 'DESCUENTOS'
        
        # Detectar categoría del concepto
        if col0 and not col0.startswith('Total') and col0 not in ['Rut', 'Nombre', 'HABERES', 'DESCUENTOS']:
            categoria_actual = col0
            
        rut_raw = str(row.iloc[2]) if pd.notna(row.iloc[2]) else ""
        monto = row.iloc[10]
        
        if '-' in rut_raw and pd.notna(monto):
            data.append({
                'Rut': clean_rut(rut_raw),
                'Concepto': f"{categoria_actual} ({'H' if seccion_actual == 'HABERES' else 'D'})",
                'Monto': monto
            })
    
    if not data: return pd.DataFrame()
    df_res = pd.DataFrame(data)
    return df_res.pivot_table(index='Rut', columns='Concepto', values='Monto', aggfunc='sum').reset_index()

# --- APLICACIÓN PRINCIPAL ---
if 'db' not in st.session_state: st.session_state['db'] = []

st.title("📊 Consolidador de Remuneraciones 2025")
st.markdown("---")

with st.sidebar:
    st.header("⚙️ Configuración")
    rut_empresa = st.text_input("RUT Empresa Principal", "76.455.680-1")
    mes = st.selectbox("Mes de Proceso", list(TOPES_AFC_UF_2025.keys()))
    if st.button("🗑️ Borrar Memoria"):
        st.session_state['db'] = []
        st.rerun()

# Carga de archivos
col_a, col_b = st.columns(2)
with col_a:
    entidad = st.radio("Empresa a cargar:", ["INGEMARS", "ENAP", "Otro"])
    nombre_entidad = st.text_input("Especifique:") if entidad == "Otro" else entidad
with col_b:
    f_libro = st.file_uploader("Libro de Remuneraciones", type=['xlsx'])
    f_informe = st.file_uploader("Informe Haberes/Descuentos", type=['xlsx'])

if f_libro and f_informe:
    if st.button(f"📥 Procesar y Guardar {nombre_entidad}"):
        try:
            lib = pd.read_excel(f_libro)
            inf = pd.read_excel(f_informe)
            
            # 1. Extraer columnas fijas y de días
            columnas_fijas = ['Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres']
            columnas_dias = [c for c in lib.columns if re.search(r'^[Nn][°º\.\s]', str(c))]
            
            # 2. Lógica de Seguro de Cesantía
            uf_mes = VALORES_UF_2025[mes]
            tope_pesos = uf_mes * TOPES_AFC_UF_2025[mes]
            
            res_lib = lib[columnas_fijas + columnas_dias].copy()
            res_lib['Total Imponible (H)'] = lib.get('Imponible', 0)
            
            # Cálculo de la tasa (2.4% o 3.0%)
            def calcular_afc(row_idx):
                imp = lib.loc[row_idx, 'Imponible'] if 'Imponible' in lib.columns else 0
                seg = lib.loc[row_idx, 'Seguro de Cesantía'] if 'Seguro de Cesantía' in lib.columns else 0
                base = min(imp, tope_pesos)
                tasa = 0.024 if seg > 0 else 0.030
                return round(base * tasa)
            
            res_lib['Seguro Cesantía Empleador (D)'] = [calcular_afc(i) for i in lib.index]
            
            # 3. Cruzar con Informe
            inf_pivot = procesar_informe_detallado(inf)
            res_lib['key'] = res_lib['Rut Trabajador'].apply(clean_rut)
            final = pd.merge(res_lib, inf_pivot, left_on='key', right_on='Rut', how='left').drop(columns=['key', 'Rut'])
            
            # 4. Metadatos
            final.insert(0, 'Empresa', rut_empresa)
            final.insert(1, 'Mes', mes)
            final.insert(2, 'Año', 2025)
            final['Origen'] = nombre_entidad
            
            st.session_state['db'].append(final)
            st.success(f"✅ {nombre_entidad} agregada correctamente.")
            
        except Exception as e:
            st.error(f"Error crítico: {e}")

# --- BOTÓN DE DESCARGA FINAL ---
if st.session_state['db']:
    st.markdown("---")
    if st.button("🚀 GENERAR EXCEL FINAL CONSOLIDADO"):
        df_final = pd.concat(st.session_state['db'], ignore_index=True)
        
        # Limpieza de nombres de columnas para evitar el error de openpyxl
        df_final.columns = [str(c).strip() for c in df_final.columns if c is not None]
        df_final = df_final.loc[:, ~df_final.columns.str.contains('^Unnamed')]
        
        # Orden inteligente
        cols = df_final.columns.tolist()
        inicio = ['Empresa', 'Mes', 'Año', 'Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres']
        dias = [c for c in cols if re.search(r'^[Nn][°º\.\s]', c)]
        haberes = sorted([c for c in cols if '(H)' in c])
        descuentos = sorted([c for c in cols if '(D)' in c])
        otros = [c for c in cols if c not in inicio + dias + haberes + descuentos + ['Origen']]
        
        df_final = df_final[inicio + dias + haberes + descuentos + otros]
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False)
        
        st.download_button("📥 Descargar Consolidado 2025", output.getvalue(), f"Consolidado_{mes}.xlsx")
        st.dataframe(df_final)
