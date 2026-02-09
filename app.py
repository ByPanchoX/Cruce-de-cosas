import streamlit as st
import pandas as pd
import io
import unicodedata
import re

st.set_page_config(page_title="Consolidador Remuneraciones (H) (D)", layout="wide")

if 'datos_acumulados' not in st.session_state:
    st.session_state['datos_acumulados'] = []

def clean_rut(rut):
    if pd.isna(rut): return None
    rut = str(rut).upper()
    return "".join(filter(lambda x: x.isdigit() or x == 'K', rut))

def parse_informe(df):
    """Extrae conceptos y detecta si son HABERES o DESCUENTOS"""
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

# --- INTERFAZ ---
st.title("📊 Consolidador: Identificación + Días + (H)/(D)")

with st.sidebar:
    st.header("1. Datos Globales")
    # RUT que aparecerá en la columna Empresa para todos
    rut_unico_excel = st.text_input("RUT Empresa para el Excel", value="76.455.680-1")
    
    st.divider()
    st.header("2. Periodo")
    mes_sel = st.selectbox("Mes", ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"])
    anio_sel = st.number_input("Año", 2025)
    
    st.divider()
    if st.button("🗑️ Reiniciar Todo"):
        st.session_state['datos_acumulados'] = []
        st.rerun()

st.subheader("Carga por Empresa")
identificador_carga = st.selectbox("Subiendo datos de:", ["INGEMARS", "ENAP", "Otra..."])
nombre_carga = st.text_input("Nombre de la carga") if identificador_carga == "Otra..." else identificador_carga

col1, col2 = st.columns(2)
with col1:
    file_libro = st.file_uploader(f"Libro de Remuneraciones - {nombre_carga}", type=["xlsx"])
with col2:
    file_informe = st.file_uploader(f"Informe Haberes/Descuentos - {nombre_carga}", type=["xlsx"])

if file_libro and file_informe:
    if st.button(f"➕ Procesar y Agregar a la lista"):
        try:
            df_lib = pd.read_excel(file_libro)
            df_inf_raw = pd.read_excel(file_informe)
            
            # --- FILTRO DE DÍAS (Solo conteos N°) ---
            id_cols = ['Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres']
            conteos = [c for c in df_lib.columns if re.search(r'^[Nn][°º\.\s]', str(c))]
            df_lib_clean = df_lib[[c for c in id_cols + conteos if c in df_lib.columns]].copy()
            
            # Limpieza de cualquier columna de texto de días que no sea N°
            drop_c = [c for c in df_lib_clean.columns if ('dia' in str(c).lower() or 'día' in str(c).lower()) and not re.search(r'^[Nn][°º\.\s]', str(c))]
            df_lib_clean = df_lib_clean.drop(columns=drop_c)
            
            # --- PROCESAR INFORME (DINERO CON (H) / (D)) ---
            df_parsed = parse_informe(df_inf_raw)
            
            # Aplicamos sufijos (H) y (D) directamente en el concepto
            df_parsed['Concepto_Final'] = df_parsed.apply(
                lambda x: f"{x['Concepto']} (H)" if x['Seccion'] == 'HABERES' else f"{x['Concepto']} (D)", axis=1
            )
            
            df_pivot = df_parsed.pivot_table(index='Rut', columns='Concepto_Final', values='Monto', aggfunc='sum').reset_index()

            # --- UNIÓN ---
            rut_c = [c for c in df_lib_clean.columns if 'Rut' in str(c)][0]
            df_lib_clean['rut_key'] = df_lib_clean[rut_c].apply(clean_rut)
            df_merged = pd.merge(df_lib_clean, df_pivot, left_on='rut_key', right_on='Rut', how='left').drop(columns=['rut_key', 'Rut'])
            
            # Metadatos
            df_merged.insert(0, 'Año', anio_sel)
            df_merged.insert(0, 'Mes', mes_sel)
            df_merged.insert(0, 'Empresa', rut_unico_excel)
            df_merged['Carga_Origen'] = nombre_carga # Para control interno de la app

            st.session_state['datos_acumulados'].append(df_merged)
            st.success(f"✅ Datos de {nombre_carga} cargados con éxito.")
            
        except Exception as e:
            st.error(f"Error: {e}")

# --- DESCARGA FINAL ---
if st.session_state['datos_acumulados']:
    st.divider()
    st.subheader("📋 Resumen de cargas")
    resumen = [{"Carga": d['Carga_Origen'].iloc[0], "Empleados": len(d)} for d in st.session_state['datos_acumulados']]
    st.table(pd.DataFrame(resumen))

    if st.button("🚀 GENERAR EXCEL FINAL CONSOLIDADO"):
        df_total = pd.concat(st.session_state['datos_acumulados'], ignore_index=True)
        
        # Eliminar columna auxiliar y duplicados técnicos
        if 'Carga_Origen' in df_total.columns:
            df_total = df_total.drop(columns=['Carga_Origen'])
        df_total = df_total.loc[:, ~df_total.columns.duplicated()]

        # Orden de columnas
        ids = ['Empresa', 'Mes', 'Año', 'Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres']
        dias_n = [c for c in df_total.columns if re.search(r'^[Nn][°º\.\s]', str(c))]
        
        # Separamos haberes (H) y descuentos (D) para que los (H) salgan primero
        haberes_final = sorted([c for c in df_total.columns if '(H)' in c])
        descuentos_final = sorted([c for c in df_total.columns if '(D)' in c])
        
        resto = sorted([c for c in df_total.columns if c not in ids and c not in dias_n and c not in haberes_final and c not in descuentos_final])
        
        df_total = df_total[ids + dias_n + haberes_final + descuentos_final + resto]
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_total.to_excel(writer, index=False, sheet_name='Consolidado')
        
        nombre_xls = f"Consolidado_{mes_sel}_{anio_sel}.xlsx"
        st.download_button(label=f"📥 Descargar {nombre_xls}", data=output.getvalue(), file_name=nombre_xls)
        st.dataframe(df_total.head(10))
