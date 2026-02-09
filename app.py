import streamlit as st
import pandas as pd
import io
import unicodedata
import re

st.set_page_config(page_title="Consolidador Remuneraciones Pro", layout="wide")

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

# --- INTERFAZ ---
st.title("📊 Consolidador de Remuneraciones")

with st.sidebar:
    st.header("Configuración")
    empresa_opcion = st.selectbox("Empresa", ["INGEMARS", "ENAP", "Otra..."])
    empresa_final = st.text_input("Nombre Personalizado") if empresa_opcion == "Otra..." else empresa_opcion
    mes_sel = st.selectbox("Mes", ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"])
    anio_sel = st.number_input("Año", 2025)
    if st.button("🗑️ Reiniciar Todo"):
        st.session_state['datos_acumulados'] = []
        st.rerun()

col1, col2 = st.columns(2)
with col1:
    file_libro = st.file_uploader("1. Libro de Remuneraciones (.xlsx)", type=["xlsx"])
with col2:
    file_informe = st.file_uploader("2. Informe Haberes y Descuentos (.xlsx)", type=["xlsx"])

if file_libro and file_informe:
    if st.button(f"➕ Procesar {empresa_final}"):
        try:
            df_lib = pd.read_excel(file_libro)
            df_inf_raw = pd.read_excel(file_informe)
            
            # --- FILTRO DEL LIBRO: SOLO IDENTIFICACIÓN Y DÍAS (EMPIEZAN CON N°) ---
            identificacion = ['Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres']
            cols_conteo = [c for c in df_lib.columns if str(c).strip().startswith('N°')]
            df_lib_clean = df_lib[[c for c in identificacion + cols_conteo if c in df_lib.columns]].copy()
            
            # --- PROCESAR INFORME (DINERO) ---
            df_parsed = parse_informe(df_inf_raw)
            df_pivot = df_parsed.pivot_table(index='Rut', columns='Concepto', values='Monto', aggfunc='sum').reset_index()
            
            # Renombrar para evitar choques con nombres de días
            renames = {}
            for col in df_pivot.columns:
                if col != 'Rut' and ('dia' in col.lower() or 'día' in col.lower()):
                    renames[col] = f"{col} (Monto)"
            df_pivot = df_pivot.rename(columns=renames)

            # --- UNIÓN ---
            rut_col_lib = [c for c in df_lib_clean.columns if 'Rut' in str(c)][0]
            df_lib_clean['rut_key'] = df_lib_clean[rut_col_lib].apply(clean_rut)
            df_merged = pd.merge(df_lib_clean, df_pivot, left_on='rut_key', right_on='Rut', how='left').drop(columns=['rut_key', 'Rut'])
            
            # Metadatos
            df_merged.insert(0, 'Año', anio_sel); df_merged.insert(0, 'Mes', mes_sel); df_merged.insert(0, 'Empresa', empresa_final)
            
            # Guardamos el resultado en un diccionario para mantener control
            st.session_state['datos_acumulados'].append({
                'df': df_merged,
                'haberes': [renames.get(c, c) for c in df_parsed[df_parsed['Seccion'] == 'HABERES']['Concepto'].unique()],
                'descuentos': [renames.get(c, c) for c in df_parsed[df_parsed['Seccion'] == 'DESCUENTOS']['Concepto'].unique()]
            })
            st.success(f"✅ {empresa_final} agregada correctamente.")
            
        except Exception as e:
            st.error(f"Error: {e}")

# --- DESCARGA ---
if st.session_state['datos_acumulados']:
    st.divider()
    if st.button("🚀 GENERAR EXCEL FINAL"):
        # CORRECCIÓN: Extraer solo los DataFrames de la lista de diccionarios
        lista_solo_dfs = [item['df'] for item in st.session_state['datos_acumulados']]
        df_total = pd.concat(lista_solo_dfs, ignore_index=True)
        
        # Eliminar cualquier columna duplicada exacta
        df_total = df_total.loc[:, ~df_total.columns.duplicated()]

        # Ordenar columnas
        ids = ['Empresa', 'Mes', 'Año', 'Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres']
        conteos = [c for c in df_total.columns if str(c).startswith('N°')]
        
        h_all, d_all = [], []
        for item in st.session_state['datos_acumulados']:
            h_all.extend(item['haberes']); d_all.extend(item['descuentos'])
        
        h_f = [c for c in sorted(list(set(h_all))) if c in df_total.columns and c not in ids]
        d_f = [c for c in sorted(list(set(d_all))) if c in df_total.columns and c not in ids]
        
        final_order = []
        for c in (ids + conteos + h_f + d_f):
            if c in df_total.columns and c not in final_order: final_order.append(c)
        
        # Agregar columnas que no entraron en las categorías anteriores
        resto = [c for c in df_total.columns if c not in final_order]
        df_total = df_total[final_order + resto]
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_total.to_excel(writer, index=False, sheet_name='Consolidado')
        
        nombre_xls = f"Consolidado_{mes_sel}_{anio_sel}.xlsx"
        st.download_button(label=f"📥 Descargar {nombre_xls}", data=output.getvalue(), file_name=nombre_xls)
        st.dataframe(df_total.head(10))
