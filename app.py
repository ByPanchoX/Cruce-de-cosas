import streamlit as st
import pandas as pd
import io
import unicodedata
import re

st.set_page_config(page_title="Consolidador Remuneraciones RUT", layout="wide")

# Inicializar la lista de datos si no existe
if 'datos_acumulados' not in st.session_state:
    st.session_state['datos_acumulados'] = []

def clean_rut(rut):
    if pd.isna(rut): return None
    rut = str(rut).upper()
    return "".join(filter(lambda x: x.isdigit() or x == 'K', rut))

def parse_informe(df):
    """Extrae conceptos del informe de haberes y descuentos"""
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
st.title("📊 Consolidador de Remuneraciones por RUT")

with st.sidebar:
    st.header("Configuración")
    # Cambio solicitado: Ahora se ingresa el RUT de la empresa
    rut_empresa = st.text_input("RUT de la Empresa (Ej: 76.455.680-1)", placeholder="12.345.678-9")
    
    mes_sel = st.selectbox("Mes", ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"])
    anio_sel = st.number_input("Año", 2025)
    
    st.divider()
    if st.button("🗑️ Reiniciar Todo"):
        st.session_state['datos_acumulados'] = []
        st.rerun()

col1, col2 = st.columns(2)
with col1:
    file_libro = st.file_uploader("1. Libro de Remuneraciones (.xlsx)", type=["xlsx"])
with col2:
    file_informe = st.file_uploader("2. Informe Haberes y Descuentos (.xlsx)", type=["xlsx"])

if file_libro and file_informe:
    if not rut_empresa:
        st.warning("⚠️ Por favor, ingresa el RUT de la empresa en la barra lateral.")
    else:
        if st.button(f"➕ Agregar Empresa RUT: {rut_empresa}"):
            try:
                df_lib = pd.read_excel(file_libro)
                df_inf_raw = pd.read_excel(file_informe)
                
                # --- FILTRO RADICAL DEL LIBRO ---
                id_cols = ['Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres']
                
                # Buscamos columnas de conteo con regex (N°, Nº, N. o N )
                conteos = [c for c in df_lib.columns if re.search(r'^[Nn][°º\.\s]', str(c))]
                
                # Reconstruimos el libro SOLO con identificación y conteos de días
                df_lib_clean = df_lib[[c for c in id_cols + conteos if c in df_lib.columns]].copy()
                
                # Borramos cualquier columna que diga 'dia' pero no tenga el símbolo de número
                intrusa_cols = [c for c in df_lib_clean.columns if ('dia' in str(c).lower() or 'día' in str(c).lower()) and not re.search(r'^[Nn][°º\.\s]', str(c))]
                df_lib_clean = df_lib_clean.drop(columns=intrusa_cols)
                
                # --- PROCESAR INFORME (DINERO) ---
                df_parsed = parse_informe(df_inf_raw)
                df_pivot = df_parsed.pivot_table(index='Rut', columns='Concepto', values='Monto', aggfunc='sum').reset_index()
                
                # Renombrar para evitar choques
                renames = {}
                for col in df_pivot.columns:
                    if col == 'Rut': continue
                    if 'dia' in col.lower() or 'día' in col.lower():
                        renames[col] = f"{col} (Monto)"
                df_pivot = df_pivot.rename(columns=renames)

                # --- UNIÓN ---
                rut_col = [c for c in df_lib_clean.columns if 'Rut' in str(c)][0]
                df_lib_clean['rut_key'] = df_lib_clean[rut_col].apply(clean_rut)
                
                df_merged = pd.merge(df_lib_clean, df_pivot, left_on='rut_key', right_on='Rut', how='left').drop(columns=['rut_key', 'Rut'])
                
                # Insertar RUT Empresa como columna 'Empresa'
                df_merged.insert(0, 'Año', anio_sel)
                df_merged.insert(0, 'Mes', mes_sel)
                df_merged.insert(0, 'Empresa', rut_empresa)
                
                st.session_state['datos_acumulados'].append(df_merged)
                st.success(f"✅ Empresa {rut_empresa} agregada correctamente.")
                
            except Exception as e:
                st.error(f"Error: {e}")

# --- DESCARGA ---
if st.session_state['datos_acumulados']:
    st.divider()
    if st.button("🚀 GENERAR EXCEL FINAL SIN DUPLICADOS"):
        df_total = pd.concat(st.session_state['datos_acumulados'], ignore_index=True)
        
        # Eliminar cualquier duplicado técnico de columnas
        df_total = df_total.loc[:, ~df_total.columns.duplicated()]

        # Limpieza final de columnas de texto de días que no son conteo ni monto del informe
        final_drop = [c for c in df_total.columns if ('dia' in str(c).lower() or 'día' in str(c).lower()) 
                      and not re.search(r'^[Nn][°º\.\s]', str(c)) 
                      and '(Monto)' not in str(c)]
        df_total = df_total.drop(columns=final_drop)

        # ORDEN DE COLUMNAS
        ids = ['Empresa', 'Mes', 'Año', 'Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres']
        dias_n = [c for c in df_total.columns if re.search(r'^[Nn][°º\.\s]', str(c))]
        otros = [c for c in df_total.columns if c not in ids and c not in dias_n]
        
        df_total = df_total[ids + dias_n + sorted(otros)]
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_total.to_excel(writer, index=False, sheet_name='Consolidado')
        
        nombre_xls = f"Consolidado_{mes_sel}_{anio_sel}.xlsx"
        st.download_button(label=f"📥 Descargar {nombre_xls}", data=output.getvalue(), file_name=nombre_xls)
        st.dataframe(df_total.head(10))
