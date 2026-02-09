import streamlit as st
import pandas as pd
import io
import unicodedata
import re

st.set_page_config(page_title="Consolidador Remuneraciones RUT Único", layout="wide")

# Inicializar la lista de datos si no existe
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
st.title("📊 Consolidador: Identificador de Carga + RUT Único")

with st.sidebar:
    st.header("1. Datos Globales")
    # Este RUT es el que saldrá en la columna 'Empresa' para todos en el Excel
    rut_unico_excel = st.text_input("RUT Empresa para el Excel", value="76.455.680-1")
    
    st.divider()
    st.header("2. Periodo")
    mes_sel = st.selectbox("Mes", ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"])
    anio_sel = st.number_input("Año", 2025)
    
    st.divider()
    if st.button("🗑️ Reiniciar Todo"):
        st.session_state['datos_acumulados'] = []
        st.rerun()

# Nombre interno para que el usuario sepa qué carga está haciendo
st.subheader("Carga Individual")
identificador_carga = st.selectbox("Identificar esta carga como:", ["INGEMARS", "ENAP", "Otra..."])
if identificador_carga == "Otra...":
    nombre_carga = st.text_input("Escriba el nombre de la empresa/carga")
else:
    nombre_carga = identificador_carga

col1, col2 = st.columns(2)
with col1:
    file_libro = st.file_uploader(f"Libro Remuneraciones - {nombre_carga}", type=["xlsx"])
with col2:
    file_informe = st.file_uploader(f"Informe Haberes/Descuentos - {nombre_carga}", type=["xlsx"])

if file_libro and file_informe:
    if st.button(f"➕ Agregar datos de {nombre_carga} a la lista"):
        try:
            df_lib = pd.read_excel(file_libro)
            df_inf_raw = pd.read_excel(file_informe)
            
            # --- FILTRO DE DÍAS (Solo N°) ---
            id_cols = ['Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres']
            conteos = [c for c in df_lib.columns if re.search(r'^[Nn][°º\.\s]', str(c))]
            df_lib_clean = df_lib[[c for c in id_cols + conteos if c in df_lib.columns]].copy()
            
            # Limpieza de duplicados de texto de días
            drop_c = [c for c in df_lib_clean.columns if ('dia' in str(c).lower() or 'día' in str(c).lower()) and not re.search(r'^[Nn][°º\.\s]', str(c))]
            df_lib_clean = df_lib_clean.drop(columns=drop_c)
            
            # --- PROCESAR INFORME (DINERO) ---
            df_parsed = parse_informe(df_inf_raw)
            df_pivot = df_parsed.pivot_table(index='Rut', columns='Concepto', values='Monto', aggfunc='sum').reset_index()
            
            # Renombrar para claridad
            rens = {c: f"{c} (Monto)" for c in df_pivot.columns if c != 'Rut' and ('dia' in c.lower() or 'día' in c.lower())}
            df_pivot = df_pivot.rename(columns=rens)

            # --- UNIÓN ---
            rut_c = [c for c in df_lib_clean.columns if 'Rut' in str(c)][0]
            df_lib_clean['rut_key'] = df_lib_clean[rut_c].apply(clean_rut)
            df_merged = pd.merge(df_lib_clean, df_pivot, left_on='rut_key', right_on='Rut', how='left').drop(columns=['rut_key', 'Rut'])
            
            # Aquí es donde guardamos el nombre_carga para la tabla de la APP, 
            # pero el rut_unico_excel para la columna Empresa
            df_merged.insert(0, 'Año', anio_sel)
            df_merged.insert(0, 'Mes', mes_sel)
            df_merged.insert(0, 'Empresa', rut_unico_excel) # Siempre usamos el RUT de la barra lateral
            df_merged['Carga_Origen'] = nombre_carga # Columna auxiliar para control interno

            st.session_state['datos_acumulados'].append(df_merged)
            st.success(f"✅ Datos de {nombre_carga} listos. Se mostrarán con el RUT {rut_unico_excel}")
            
        except Exception as e:
            st.error(f"Error: {e}")

# --- SECCIÓN FINAL ---
if st.session_state['datos_acumulados']:
    st.divider()
    st.subheader("📋 Resumen de cargas en memoria")
    resumen = []
    for d in st.session_state['datos_acumulados']:
        resumen.append({"Carga": d['Carga_Origen'].iloc[0], "RUT en Excel": d['Empresa'].iloc[0], "Empleados": len(d)})
    st.table(pd.DataFrame(resumen))

    if st.button("🚀 GENERAR EXCEL FINAL"):
        df_total = pd.concat(st.session_state['datos_acumulados'], ignore_index=True)
        
        # Eliminamos la columna auxiliar antes de exportar
        if 'Carga_Origen' in df_total.columns:
            df_total = df_total.drop(columns=['Carga_Origen'])
            
        df_total = df_total.loc[:, ~df_total.columns.duplicated()]

        # Limpieza final de columnas de texto de días
        final_d = [c for c in df_total.columns if ('dia' in str(c).lower() or 'día' in str(c).lower()) 
                   and not re.search(r'^[Nn][°º\.\s]', str(c)) 
                   and '(Monto)' not in str(c)]
        df_total = df_total.drop(columns=final_d)

        # Orden de columnas
        ids = ['Empresa', 'Mes', 'Año', 'Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres']
        dias_n = [c for c in df_total.columns if re.search(r'^[Nn][°º\.\s]', str(c))]
        otros = sorted([c for c in df_total.columns if c not in ids and c not in dias_n])
        
        df_total = df_total[ids + dias_n + otros]
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_total.to_excel(writer, index=False, sheet_name='Consolidado')
        
        nombre_xls = f"Consolidado_{mes_sel}_{anio_sel}.xlsx"
        st.download_button(label=f"📥 Descargar {nombre_xls}", data=output.getvalue(), file_name=nombre_xls)
        st.dataframe(df_total.head(10))
