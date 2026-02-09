import streamlit as st
import pandas as pd
import io
import re

# Configuración 2025
st.set_page_config(page_title="Consolidador Auditoría 2025", layout="wide", page_icon="📊")

TOPES_AFC_UF_2025 = {"Enero": 131.8, "Febrero": 131.9, "Marzo": 131.9, "Abril": 131.9, "Mayo": 131.9, "Junio": 131.9, "Julio": 131.9, "Agosto": 131.9, "Septiembre": 131.9, "Octubre": 131.9, "Noviembre": 131.9, "Diciembre": 131.9}
VALORES_UF_2025 = {"Enero": 38284.86, "Febrero": 38647.94, "Marzo": 38800.00, "Abril": 39050.00, "Mayo": 39200.00, "Junio": 39267.07, "Julio": 39179.01, "Agosto": 39383.07, "Septiembre": 39500.00, "Octubre": 39597.67, "Noviembre": 39643.59, "Diciembre": 39727.96}

if 'db' not in st.session_state: st.session_state['db'] = []

def clean_rut(rut):
    if pd.isna(rut): return None
    return "".join(filter(lambda x: x.isdigit() or x == 'K', str(rut).upper()))

def procesar_informe(df):
    data = []
    seccion, categoria = 'HABERES', None
    for _, row in df.iterrows():
        c0 = str(row.iloc[0]) if pd.notna(row.iloc[0]) else ""
        if 'HABERES' in c0.upper(): seccion = 'HABERES'
        elif 'DESCUENTOS' in c0.upper(): seccion = 'DESCUENTOS'
        if c0 and not c0.startswith('Total') and c0 not in ['Rut', 'Nombre', 'HABERES', 'DESCUENTOS']: categoria = c0
        rut_r = str(row.iloc[2]) if pd.notna(row.iloc[2]) else ""
        monto = row.iloc[10]
        if '-' in rut_r and pd.notna(monto):
            data.append({'Rut': clean_rut(rut_r), 'Concepto': f"{categoria} ({'H' if seccion == 'HABERES' else 'D'})", 'Monto': monto})
    if not data: return pd.DataFrame()
    return pd.DataFrame(data).pivot_table(index='Rut', columns='Concepto', values='Monto', aggfunc='sum').reset_index()

st.title("📊 Consolidador Pro 2025: Auditoría Total")

with st.sidebar:
    st.header("⚙️ Configuración")
    mes = st.selectbox("Mes de Proceso", list(TOPES_AFC_UF_2025.keys()))
    rut_emp = st.text_input("RUT Empresa Principal", "76.455.680-1")
    if st.button("🗑️ Limpiar Todo"):
        st.session_state['db'] = []
        st.rerun()

st.info("💡 Consejo: Para mejores resultados, sube los 3 archivos de la misma empresa antes de presionar 'Procesar'.")

col_empresa = st.selectbox("Empresa:", ["INGEMARS", "ENAP", "Otro"])
nombre_entidad = st.text_input("Especifique:") if col_empresa == "Otro" else col_empresa

c1, c2, c3 = st.columns(3)
with c1: f_lib = st.file_uploader("1. Libro Remuneraciones", type=['xlsx'])
with c2: f_inf = st.file_uploader("2. Informe Haberes/Desc", type=['xlsx'])
with c3: f_pat = st.file_uploader("3. Aporte Patronal (Opcional)", type=['xlsx'])

if f_lib or f_inf or f_pat:
    if st.button(f"📥 Procesar y Agregar {nombre_entidad}"):
        try:
            # Procesar Libro
            res_lib = pd.DataFrame()
            if f_lib:
                lib = pd.read_excel(f_lib)
                dias = [c for c in lib.columns if re.search(r'^[Nn][°º\.\s]', str(c))]
                res_lib = lib[['Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres'] + dias].copy()
                res_lib['Total Imponible (H)'] = lib.get('Imponible', 0)
                res_lib['key'] = res_lib['Rut Trabajador'].apply(clean_rut)
                
                # Cálculo Seguro
                tope_p = VALORES_UF_2025[mes] * TOPES_AFC_UF_2025[mes]
                def calc_afc(idx):
                    imp = lib.loc[idx, 'Imponible'] if 'Imponible' in lib.columns else 0
                    seg_t = lib.loc[idx, 'Seguro de Cesantía'] if 'Seguro de Cesantía' in lib.columns else 0
                    tasa = 0.024 if seg_t > 0 else 0.030
                    return round(min(imp, tope_p) * tasa)
                res_lib['Seguro Cesantía Empleador (P)'] = [calc_afc(i) for i in lib.index]

            # Procesar Informe
            inf_pivot = procesar_informe(pd.read_excel(f_inf)) if f_inf else pd.DataFrame()
            
            # Procesar Patronal
            pat_data = pd.DataFrame()
            if f_pat:
                pat_df = pd.read_excel(f_pat)
                pat_df['key_rut'] = pat_df['Rut Trabajador'].apply(clean_rut)
                cols_pat = ['key_rut', 'Seguro de Invalidez y Sobrevivencia Empleador', 'Centro de Costo', 'Cargo', 'Tipo de Contrato']
                pat_data = pat_df[[c for c in cols_pat if c in pat_df.columns]].rename(columns={'Seguro de Invalidez y Sobrevivencia Empleador': 'SIS (P)'})

            # Cruzar datos de la carga actual
            df_merged = res_lib if not res_lib.empty else (inf_pivot if not inf_pivot.empty else pat_data)
            if not res_lib.empty and not inf_pivot.empty:
                df_merged = pd.merge(df_merged, inf_pivot, left_on='key', right_on='Rut', how='outer')
            if not pat_data.empty:
                df_merged = pd.merge(df_merged, pat_data, left_on='key' if 'key' in df_merged.columns else 'Rut', right_on='key_rut', how='outer')

            # Limpiar y preparar
            if 'key' not in df_merged.columns and 'Rut' in df_merged.columns: df_merged['key'] = df_merged['Rut']
            if 'key' not in df_merged.columns and 'key_rut' in df_merged.columns: df_merged['key'] = df_merged['key_rut']
            
            df_merged.insert(0, 'Empresa', rut_emp)
            df_merged.insert(1, 'Mes', mes)
            df_merged.insert(2, 'Año', 2025)
            df_merged['Origen'] = nombre_entidad
            
            st.session_state['db'].append(df_merged)
            st.success(f"✅ {nombre_entidad} agregada.")
            
        except Exception as e:
            st.error(f"Error: {e}")

if st.session_state['db']:
    if st.button("🚀 GENERAR EXCEL CONSOLIDADO FINAL"):
        # 1. Unir todas las cargas
        df_all = pd.concat(st.session_state['db'], ignore_index=True)
        
        # 2. FUSIONAR DUPLICADOS: Agrupar por RUT y combinar datos
        # Esto junta las filas divididas en una sola fila por persona
        df_final = df_all.groupby('key').first().reset_index()
        
        # Limpieza de nombres
        df_final.columns = [str(c).strip() for c in df_final.columns]
        df_final = df_final.loc[:, ~df_final.columns.str.contains('^Unnamed|^key|^Rut|^key_rut')]
        
        # Reordenar
        fijas = ['Empresa', 'Mes', 'Año', 'Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres', 'Cargo', 'Centro de Costo', 'Tipo de Contrato']
        dias = [c for c in df_final.columns if re.search(r'^[Nn][°º\.\s]', c)]
        habs = sorted([c for c in df_final.columns if '(H)' in c])
        dess = sorted([c for c in df_final.columns if '(D)' in c])
        pats = sorted([c for c in df_final.columns if '(P)' in c or 'SIS' in c])
        otros = [c for c in df_final.columns if c not in fijas+dias+habs+dess+pats+['Origen']]
        
        df_final = df_final[[c for c in fijas+dias+habs+dess+pats+otros if c in df_final.columns]]
        
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False)
        st.download_button("📥 Descargar Excel Fusionado", out.getvalue(), f"Consolidado_{mes}_2025_Final.xlsx")
        st.dataframe(df_final)
