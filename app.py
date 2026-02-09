import streamlit as st
import pandas as pd
import io
import re

# Configuración 2025
st.set_page_config(page_title="Consolidador Auditoría 2025", layout="wide", page_icon="🏦")

# TOPES Y VALORES 2025
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

st.title("📊 Consolidador con RUT al Inicio y Aportes Patronales")

with st.sidebar:
    st.header("⚙️ Configuración")
    mes = st.selectbox("Mes de Proceso", list(TOPES_AFC_UF_2025.keys()))
    rut_emp = st.text_input("RUT Empresa (Fijo)", "76.455.680-1")
    if st.button("🗑️ Limpiar Memoria"):
        st.session_state['db'] = []
        st.rerun()

entidad = st.selectbox("Empresa:", ["INGEMARS", "ENAP", "Otro"])
nombre_entidad = st.text_input("Especifique:") if entidad == "Otro" else entidad

c1, c2, c3 = st.columns(3)
with c1: f_lib = st.file_uploader("1. Libro Remuneraciones", type=['xlsx'])
with c2: f_inf = st.file_uploader("2. Informe Haberes/Desc", type=['xlsx'])
with c3: f_pat = st.file_uploader("3. Aporte Patronal", type=['xlsx'])

if f_lib or f_inf or f_pat:
    if st.button(f"➕ Procesar y Fusionar {nombre_entidad}"):
        try:
            # 1. LIBRO
            df_l = pd.DataFrame()
            if f_lib:
                lib = pd.read_excel(f_lib)
                dias = [c for c in lib.columns if re.search(r'^[Nn][°º\.\s]', str(c))]
                df_l = lib[['Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres'] + dias].copy()
                df_l['Total Imponible (H)'] = lib.get('Imponible', 0)
                df_l['key'] = df_l['Rut Trabajador'].apply(clean_rut)

            # 2. INFORME
            inf_p = procesar_informe(pd.read_excel(f_inf)) if f_inf else pd.DataFrame()
            
            # 3. APORTE PATRONAL (Columnas específicas solicitadas)
            pat_d = pd.DataFrame()
            if f_pat:
                pat_df = pd.read_excel(f_pat)
                pat_df['key_rut'] = pat_df['Rut Trabajador'].apply(clean_rut)
                mapeo = {
                    'Aporte Accidentes de Trabajo IPS': 'Accidentes Trabajo IPS (P)',
                    'Aporte Accidentes de Trabajo Mutual': 'Accidentes Trabajo Mutual (P)',
                    'Adicional de Capitalización Individual AFP': 'Cap. Adicional AFP (P)',
                    'Expectativa de Vida Seguro Social': 'Expectativa Vida SS (P)',
                    'Seguro de Cesantía Empleador': 'Seguro Cesantía Empleador (P)',
                    'Seguro de Invalidez y Sobrevivencia Empleador': 'SIS (P)',
                    'Centro de Costo': 'Centro de Costo', 'Cargo': 'Cargo', 'Tipo de Contrato': 'Tipo de Contrato'
                }
                cols_ok = [c for c in mapeo.keys() if c in pat_df.columns]
                pat_d = pat_df[['key_rut'] + cols_ok].rename(columns=mapeo)

            # --- UNIÓN TEMPORAL ---
            df_m = df_l if not df_l.empty else (inf_p if not inf_p.empty else pat_d)
            if not df_l.empty and not inf_p.empty:
                df_m = pd.merge(df_m, inf_p, left_on='key', right_on='Rut', how='outer')
            if not pat_d.empty:
                llave = 'key' if 'key' in df_m.columns else ('Rut' if 'Rut' in df_m.columns else 'key_rut')
                df_m = pd.merge(df_m, pat_d, left_on=llave, right_on='key_rut', how='outer')

            if 'key' not in df_m.columns:
                df_m['key'] = df_m['Rut'] if 'Rut' in df_m.columns else df_m['key_rut']
            
            df_m.insert(0, 'Empresa_RUT', rut_emp)
            df_m.insert(1, 'Mes', mes)
            df_m.insert(2, 'Año', 2025)
            df_m['Origen'] = nombre_entidad
            
            st.session_state['db'].append(df_m)
            st.success(f"✅ Datos de {nombre_entidad} agregados.")
            
        except Exception as e:
            st.error(f"Error: {e}")

if st.session_state['db']:
    st.divider()
    if st.button("🚀 GENERAR EXCEL FINAL CON RUT AL INICIO"):
        df_all = pd.concat(st.session_state['db'], ignore_index=True)
        
        # FUSIÓN POR RUT: Una sola fila por trabajador
        df_final = df_all.groupby('key').first().reset_index()
        
        # Limpieza de columnas técnicas
        df_final.columns = [str(c).strip() for c in df_final.columns]
        df_final = df_final.loc[:, ~df_final.columns.str.contains('^Unnamed|^key|^Rut|^key_rut')]
        
        # --- ORDENACIÓN: RUT TRABAJADOR AL INICIO ---
        fijas = ['Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres', 'Empresa_RUT', 'Mes', 'Año', 'Cargo', 'Centro de Costo', 'Tipo de Contrato']
        dias = [c for c in df_final.columns if re.search(r'^[Nn][°º\.\s]', c)]
        habs = sorted([c for c in df_final.columns if '(H)' in c])
        dess = sorted([c for c in df_final.columns if '(D)' in c])
        pats = sorted([c for c in df_final.columns if '(P)' in c or 'SIS' in c])
        otros = [c for c in df_final.columns if c not in fijas+dias+habs+dess+pats+['Origen']]
        
        # Filtrar solo columnas que existan
        cols_finales = [c for c in fijas+dias+habs+dess+pats+otros if c in df_final.columns]
        df_final = df_final[cols_finales]

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False)
        
        st.download_button(
            label="📥 Descargar Excel Final (RUT al Inicio)",
            data=output.getvalue(),
            file_name=f"Consolidado_RUT_Primero_{mes}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.dataframe(df_final.head(10))
