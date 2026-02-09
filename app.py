import streamlit as st
import pandas as pd
import io
import re

# Configuración 2025
st.set_page_config(page_title="Consolidador Auditoría Pro 2025", layout="wide", page_icon="🏦")

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
            data.append({'RutKey': clean_rut(rut_r), 'Concepto': f"{categoria} ({'H' if seccion == 'HABERES' else 'D'})", 'Monto': monto})
    if not data: return pd.DataFrame()
    return pd.DataFrame(data).pivot_table(index='RutKey', columns='Concepto', values='Monto', aggfunc='sum').reset_index()

st.title("📊 Consolidador: Estructura Empresa - Periodo - Trabajador")

with st.sidebar:
    st.header("⚙️ Configuración")
    mes = st.selectbox("Mes de Proceso", list(VALORES_UF_2025.keys()))
    rut_emp_input = st.text_input("RUT Empresa", "76.455.680-1")
    if st.button("🗑️ Limpiar Todo"):
        st.session_state['db'] = []
        st.rerun()

entidad = st.selectbox("Origen de Datos:", ["INGEMARS", "ENAP", "Otro"])
nombre_entidad = st.text_input("Especifique:") if entidad == "Otro" else entidad

c1, c2, c3 = st.columns(3)
with c1: f_lib = st.file_uploader("1. Libro Remuneraciones", type=['xlsx'])
with c2: f_inf = st.file_uploader("2. Informe Haberes/Desc", type=['xlsx'])
with c3: f_pat = st.file_uploader("3. Aporte Patronal", type=['xlsx'])

if f_lib or f_inf or f_pat:
    if st.button(f"➕ Procesar {nombre_entidad}"):
        try:
            res_lib = pd.DataFrame()
            if f_lib:
                lib = pd.read_excel(f_lib)
                dias = [c for c in lib.columns if re.search(r'^[Nn][°º\.\s]', str(c))]
                res_lib = lib[['Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres'] + dias].copy()
                res_lib['Total Imponible (H)'] = lib.get('Imponible', 0)
                res_lib['RutKey'] = res_lib['Rut Trabajador'].apply(clean_rut)

            inf_pivot = procesar_informe(pd.read_excel(f_inf)) if f_inf else pd.DataFrame()
            
            pat_data = pd.DataFrame()
            if f_pat:
                pat_df = pd.read_excel(f_pat)
                pat_df['RutKey'] = pat_df['Rut Trabajador'].apply(clean_rut)
                mapeo = {
                    'Rut Trabajador': 'Rut Trabajador',
                    'Aporte Accidentes de Trabajo IPS': 'Acc. Trabajo IPS (P)',
                    'Aporte Accidentes de Trabajo Mutual': 'Acc. Trabajo Mutual (P)',
                    'Adicional de Capitalización Individual AFP': 'Cap. Adicional AFP (P)',
                    'Expectativa de Vida Seguro Social': 'Expectativa Vida SS (P)',
                    'Seguro de Cesantía Empleador': 'Seguro Cesantía Empleador (P)',
                    'Seguro de Invalidez y Sobrevivencia Empleador': 'SIS (P)',
                    'Cargo': 'Cargo', 'Centro de Costo': 'Centro de Costo', 'Tipo de Contrato': 'Tipo de Contrato'
                }
                cols_encontradas = [c for c in mapeo.keys() if c in pat_df.columns]
                pat_data = pat_df[['RutKey'] + cols_encontradas].rename(columns=mapeo)

            # Fusión
            df_curr = res_lib if not res_lib.empty else (pat_data if not pat_data.empty else inf_pivot)
            if not res_lib.empty and not inf_pivot.empty:
                df_curr = pd.merge(df_curr, inf_pivot, on='RutKey', how='outer')
            if not pat_data.empty and df_curr is not pat_data:
                df_curr = pd.merge(df_curr, pat_data, on='RutKey', how='outer', suffixes=('', '_drop'))

            # Definir columnas de cabecera solicitadas
            df_curr.insert(0, 'RUT EMPRESA', rut_emp_input)
            df_curr.insert(1, 'MES', mes)
            df_curr.insert(2, 'AÑO', 2025)
            
            st.session_state['db'].append(df_curr)
            st.success(f"✅ {nombre_entidad} cargada.")
            
        except Exception as e:
            st.error(f"Error: {e}")

if st.session_state['db']:
    if st.button("🚀 GENERAR EXCEL CONSOLIDADO"):
        df_all = pd.concat(st.session_state['db'], ignore_index=True)
        df_final = df_all.groupby('RutKey').first().reset_index()
        df_final = df_final.loc[:, ~df_final.columns.str.contains('_drop$|^RutKey$')]
        
        # ORDEN ESTRICTO SOLICITADO
        # 1. RUT Empresa, 2. Mes, 3. Año, 4. RUT Trabajador, 5. Otros
        cols_id = ['RUT EMPRESA', 'MES', 'AÑO', 'Rut Trabajador']
        cols_personales = ['Apellido Paterno', 'Apellido Materno', 'Nombres', 'Cargo', 'Centro de Costo', 'Tipo de Contrato']
        cols_dias = [c for c in df_final.columns if re.search(r'^[Nn][°º\.\s]', c)]
        cols_hab = sorted([c for c in df_final.columns if '(H)' in c])
        cols_des = sorted([c for c in df_final.columns if '(D)' in c])
        cols_pat = sorted([c for c in df_final.columns if '(P)' in c or 'SIS' in c])
        resto = [c for c in df_final.columns if c not in cols_id + cols_personales + cols_dias + cols_hab + cols_des + cols_pat]
        
        orden_final = [c for c in cols_id + cols_personales + cols_dias + cols_hab + cols_des + cols_pat + resto if c in df_final.columns]
        df_final = df_final[orden_final]
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False)
        st.download_button("📥 Descargar Excel", output.getvalue(), f"Consolidado_{mes}_2025.xlsx")
        st.dataframe(df_final)
