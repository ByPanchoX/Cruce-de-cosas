import streamlit as st
import pandas as pd
import io
import re

# Configuración de página
st.set_page_config(page_title="Consolidador Auditoría Final 2025", layout="wide", page_icon="🏦")

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

st.title("📊 Sistema de Consolidación 2025")

with st.sidebar:
    st.header("⚙️ Configuración")
    mes_sel = st.selectbox("Mes de Proceso", list(VALORES_UF_2025.keys()))
    rut_emp_global = st.text_input("RUT Empresa Principal", "76.455.680-1")
    if st.button("🗑️ Reiniciar Todo"):
        st.session_state['db'] = []
        st.rerun()

entidad = st.selectbox("Origen de Datos:", ["INGEMARS", "ENAP", "Otro"])
nombre_entidad = st.text_input("Especifique:") if entidad == "Otro" else entidad

c1, c2, c3 = st.columns(3)
with c1: f_lib = st.file_uploader("1. Libro de Remuneraciones", type=['xlsx'])
with c2: f_inf = st.file_uploader("2. Informe Haberes/Descuentos", type=['xlsx'])
with c3: f_pat = st.file_uploader("3. Aporte Patronal (Opcional)", type=['xlsx'])

if f_lib or f_inf or f_pat:
    if st.button(f"➕ Procesar y Validar {nombre_entidad}", type="primary"):
        try:
            # 1. LIBRO
            df_lib = pd.DataFrame()
            if f_lib:
                raw_lib = pd.read_excel(f_lib)
                dias = [c for c in raw_lib.columns if re.search(r'^[Nn][°º\.\s]', str(c))]
                df_lib = raw_lib[['Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres'] + dias].copy()
                df_lib['Total Imponible (H)'] = raw_lib.get('Imponible', 0)
                df_lib['RutKey'] = df_lib['Rut Trabajador'].apply(clean_rut)

            # 2. INFORME
            df_inf_pivot = procesar_informe(pd.read_excel(f_inf)) if f_inf else pd.DataFrame()
            
            # 3. APORTE PATRONAL
            df_pat_clean = pd.DataFrame()
            if f_pat:
                pat_raw = pd.read_excel(f_pat)
                pat_raw['RutKey'] = pat_raw['Rut Trabajador'].apply(clean_rut)
                mapeo = {
                    'Rut Trabajador': 'Rut Trabajador',
                    'Aporte Accidentes de Trabajo IPS': 'Accidentes Trabajo IPS (P)',
                    'Aporte Accidentes de Trabajo Mutual': 'Accidentes Trabajo Mutual (P)',
                    'Adicional de Capitalización Individual AFP': 'Cap. Adicional AFP (P)',
                    'Expectativa de Vida Seguro Social': 'Expectativa Vida SS (P)',
                    'Seguro de Cesantía Empleador': 'Seguro Cesantía Empleador (P)',
                    'Seguro de Invalidez y Sobrevivencia Empleador': 'SIS (P)',
                    'Cargo': 'Cargo', 'Centro de Costo': 'Centro de Costo', 'Tipo de Contrato': 'Tipo de Contrato'
                }
                cols_ok = [c for c in mapeo.keys() if c in pat_raw.columns]
                df_pat_clean = pat_raw[['RutKey'] + cols_ok].rename(columns=mapeo)

            # --- UNIÓN INTELIGENTE ---
            df_curr = df_lib if not df_lib.empty else (df_pat_clean if not df_pat_clean.empty else df_inf_pivot)
            
            if not df_lib.empty and not df_inf_pivot.empty:
                df_curr = pd.merge(df_curr, df_inf_pivot, on='RutKey', how='outer')
            
            if not df_pat_clean.empty and df_curr is not df_pat_clean:
                df_curr = pd.merge(df_curr, df_pat_clean, on='RutKey', how='outer', suffixes=('', '_tecnico'))

            # Cabeceras finales
            df_curr['RUT EMPRESA'] = rut_emp_global
            df_curr['MES'] = mes_sel
            df_curr['AÑO'] = 2025
            df_curr['__origen__'] = nombre_entidad
            
            st.session_state['db'].append(df_curr)
            st.success(f"✅ {nombre_entidad} procesada con éxito.")
            
        except Exception as e:
            st.error(f"Error al procesar: {e}")

if st.session_state['db']:
    st.divider()
    if st.button("🚀 GENERAR EXCEL FINAL LIMPIO", use_container_width=True):
        df_all = pd.concat(st.session_state['db'], ignore_index=True)
        
        # Fusión por RUT para que no haya duplicados
        df_final = df_all.groupby('RutKey').first().reset_index()
        
        # --- ELIMINACIÓN DE COLUMNAS BASURA ---
        # Borramos RutKey y cualquier columna que contenga nombres técnicos
        cols_a_borrar = [c for c in df_final.columns if 'RutKey' in c or '_tecnico' in c or '__origen__' in c or 'Empresa' == c or 'Mes' == c or 'Año' == c]
        df_final = df_final.drop(columns=cols_a_borrar, errors='ignore')

        # --- ORDEN ESTRICTO ---
        # 1. RUT EMPRESA, 2. MES, 3. AÑO, 4. RUT TRABAJADOR
        orden_id = ['RUT EMPRESA', 'MES', 'AÑO', 'Rut Trabajador']
        orden_personal = ['Apellido Paterno', 'Apellido Materno', 'Nombres', 'Cargo', 'Centro de Costo', 'Tipo de Contrato']
        
        # Clasificar el resto
        cols_dias = [c for c in df_final.columns if re.search(r'^[Nn][°º\.\s]', str(c))]
        cols_hab = sorted([c for c in df_final.columns if '(H)' in c])
        cols_des = sorted([c for c in df_final.columns if '(D)' in c])
        cols_pat = sorted([c for c in df_final.columns if '(P)' in c or 'SIS' in c])
        
        # Columnas que no entran en ninguna categoría (seguridad)
        ya_ordenadas = orden_id + orden_personal + cols_dias + cols_hab + cols_des + cols_pat
        resto = [c for c in df_final.columns if c not in ya_ordenadas]
        
        # Construir orden final
        final_cols = [c for c in orden_id + orden_personal + cols_dias + cols_hab + cols_des + cols_pat + resto if c in df_final.columns]
        df_final = df_final[final_cols]
        
        # Exportar
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False)
        
        st.download_button(
            label="📥 DESCARGAR EXCEL CONSOLIDADO FINAL",
            data=output.getvalue(),
            file_name=f"Consolidado_Auditoria_{mes_sel}_2025.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.dataframe(df_final)
