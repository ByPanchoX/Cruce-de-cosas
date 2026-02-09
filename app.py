import streamlit as st
import pandas as pd
import io
import re

# Configuración
st.set_page_config(page_title="Consolidador Pro + SIS 2025", layout="wide", page_icon="📈")

# DATOS LEGALES 2025
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

st.title("📊 Consolidador con Auditoría SIS y Aporte Patronal")

with st.sidebar:
    st.header("⚙️ Periodo 2025")
    mes = st.selectbox("Mes", list(TOPES_AFC_UF_2025.keys()))
    rut_emp = st.text_input("RUT Empresa Principal", "76.455.680-1")
    if st.button("🗑️ Limpiar Memoria"):
        st.session_state['db'] = []
        st.rerun()

entidad = st.selectbox("Empresa:", ["INGEMARS", "ENAP", "Otro"])
nombre_entidad = st.text_input("Nombre:") if entidad == "Otro" else entidad

c1, c2, c3 = st.columns(3)
with c1: f_lib = st.file_uploader("1. Libro Remuneraciones", type=['xlsx'])
with c2: f_inf = st.file_uploader("2. Informe Haberes/Desc", type=['xlsx'])
with c3: f_pat = st.file_uploader("3. Aporte Patronal (Opcional)", type=['xlsx'])

if f_lib and f_inf:
    if st.button(f"📥 Procesar {nombre_entidad}"):
        try:
            lib = pd.read_excel(f_lib)
            inf_pivot = procesar_informe(pd.read_excel(f_inf))
            
            # Datos de Aporte Patronal (Opcional)
            pat_data = pd.DataFrame()
            if f_pat:
                pat_df = pd.read_excel(f_pat)
                pat_df['key_rut'] = pat_df['Rut Trabajador'].apply(clean_rut)
                # Extraer SIS, Cargo, Centro Costo y Contrato
                cols_pat = ['key_rut', 'Seguro de Invalidez y Sobrevivencia Empleador', 'Centro de Costo', 'Cargo', 'Tipo de Contrato']
                pat_data = pat_df[[c for c in cols_pat if c in pat_df.columns]]

            # Procesar Libro
            dias = [c for c in lib.columns if re.search(r'^[Nn][°º\.\s]', str(c))]
            res = lib[['Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres'] + dias].copy()
            res['Total Imponible (H)'] = lib.get('Imponible', 0)
            res['key'] = res['Rut Trabajador'].apply(clean_rut)

            # Cálculo Seguro Cesantía con Topes
            tope_p = VALORES_UF_2025[mes] * TOPES_AFC_UF_2025[mes]
            
            def calc_afc(idx):
                imp = lib.loc[idx, 'Imponible'] if 'Imponible' in lib.columns else 0
                seg_t = lib.loc[idx, 'Seguro de Cesantía'] if 'Seguro de Cesantía' in lib.columns else 0
                # Si tenemos archivo patronal, usamos el tipo de contrato real
                tasa = 0.024
                if not pat_data.empty:
                    rut_actual = clean_rut(lib.loc[idx, 'Rut Trabajador'])
                    match = pat_data[pat_data['key_rut'] == rut_actual]
                    if not match.empty and 'A Plazo' in str(match['Tipo de Contrato'].values[0]): tasa = 0.030
                    else: tasa = 0.024 if seg_t > 0 else 0.030
                else:
                    tasa = 0.024 if seg_t > 0 else 0.030
                return round(min(imp, tope_p) * tasa)

            res['Seguro Cesantía Empleador (P)'] = [calc_afc(i) for i in lib.index]

            # Cruzar Todo
            m1 = pd.merge(res, inf_pivot, left_on='key', right_on='Rut', how='left')
            if not pat_data.empty:
                m1 = pd.merge(m1, pat_data, left_on='key', right_on='key_rut', how='left').drop(columns=['key_rut'])
                m1.rename(columns={'Seguro de Invalidez y Sobrevivencia Empleador': 'SIS (P)'}, inplace=True)
            
            final = m1.drop(columns=['key', 'Rut'], errors='ignore')
            final.insert(0, 'Empresa', rut_emp)
            final.insert(1, 'Mes', mes)
            final.insert(2, 'Año', 2025)
            final['Origen'] = nombre_entidad
            
            st.session_state['db'].append(final)
            st.success(f"✅ Procesado con éxito.")
            
        except Exception as e:
            st.error(f"Error: {e}")

if st.session_state['db']:
    if st.button("🚀 GENERAR EXCEL FINAL"):
        df = pd.concat(st.session_state['db'], ignore_index=True)
        df.columns = [str(c).strip() for c in df.columns]
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
        
        # Orden de Columnas
        fijas = ['Empresa', 'Mes', 'Año', 'Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres', 'Cargo', 'Centro de Costo']
        dias = [c for c in df.columns if re.search(r'^[Nn][°º\.\s]', c)]
        habs = sorted([c for c in df.columns if '(H)' in c])
        dess = sorted([c for c in df.columns if '(D)' in c])
        pats = sorted([c for c in df.columns if '(P)' in c])
        otros = [c for c in df.columns if c not in fijas+dias+habs+dess+pats+['Origen']]
        
        df = df[[c for c in fijas+dias+habs+dess+pats+otros if c in df.columns]]
        
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        st.download_button("📥 Descargar Consolidado", out.getvalue(), f"Consolidado_{mes}.xlsx")
        st.dataframe(df)
