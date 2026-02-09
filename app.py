import streamlit as st
import pandas as pd
import io
import unicodedata
import re

st.set_page_config(page_title="Consolidador Remuneraciones Pro", layout="wide")

if 'datos_acumulados' not in st.session_state:
    st.session_state['datos_acumulados'] = []

def normalizar(texto):
    """Limpia tildes, espacios y mayúsculas para comparar."""
    if pd.isna(texto): return ""
    texto = str(texto).lower().strip()
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    texto = re.sub(r'[^a-z0-9]', '', texto)
    return texto

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
st.title("📊 Consolidador de Remuneraciones (Versión Final)")

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
            
            # --- LIMPIEZA DEL LIBRO (SOLO CONTEO DE DÍAS) ---
            # Solo permitimos columnas de identificación y las que EMPIEZAN con N°
            identificacion = ['Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres']
            cols_conteo = [c for c in df_lib.columns if str(c).strip().startswith('N°')]
            
            # Filtramos estrictamente el libro
            df_lib_clean = df_lib[[c for c in identificacion + cols_conteo if c in df_lib.columns]].copy()
            
            # --- PROCESAR INFORME (DINERO) ---
            df_parsed = parse_informe(df_inf_raw)
            df_pivot = df_parsed.pivot_table(index='Rut', columns='Concepto', values='Monto', aggfunc='sum').reset_index()
            
            # Renombrar conceptos del informe que puedan chocar con los días del libro
            renames = {}
            for col in df_pivot.columns:
                if col == 'Rut': continue
                # Si el nombre del bono/descuento se parece a "Dias" o "Días", le ponemos (Monto)
                if 'dia' in col.lower() or 'día' in col.lower():
                    renames[col] = f"{col} (Monto)"
            df_pivot = df_pivot.rename(columns=renames)

            # --- UNIÓN ---
            rut_col_lib = [c for c in df_lib_clean.columns if 'Rut' in str(c)][0]
            df_lib_clean['rut_key'] = df_lib_clean[rut_col_lib].apply(clean_rut)
            
            df_merged = pd.merge(df_lib_clean, df_pivot, left_on='rut_key', right_on='Rut', how='left').drop(columns=['rut_key', 'Rut'])
            
            # Metadatos
            df_merged.insert(0, 'Año', anio_sel); df_merged.insert(0, 'Mes', mes_sel); df_merged.insert(0, 'Empresa', empresa_final)
            
            st.session_state['datos_acumulados'].append(df_merged)
            st.success(f"✅ {empresa_final} agregada sin duplicados.")
            
        except Exception as e:
            st.error(f"Error: {e}")

# --- DESCARGA ---
if st.session_state['datos_acumulados']:
    st.divider()
    if st.button("🚀 GENERAR EXCEL FINAL"):
        # Unimos todo y eliminamos cualquier columna duplicada que haya quedado por error
        df_total = pd.concat(st.session_state['datos_acumulados'], ignore_index=True)
        
        # Eliminar duplicados de columnas (por si acaso quedaron nombres idénticos)
        df_total = df_total.loc[:, ~df_total.columns.duplicated()]

        # Ordenar columnas
        ids = ['Empresa', 'Mes', 'Año', 'Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres']
        dias = [c for c in df_total.columns if str(c).startswith('N°')]
        # El resto son haberes y descuentos
        resto = [c for c in df_total.columns if c not in ids and c not in dias]
        
        df_total = df_total[ids + dias + sorted(resto)]
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_total.to_excel(writer, index=False, sheet_name='Consolidado')
        
        nombre_xls = f"Consolidado_{mes_sel}_{anio_sel}.xlsx"
        st.download_button(label=f"📥 Descargar {nombre_xls}", data=output.getvalue(), file_name=nombre_xls)
        st.dataframe(df_total.head(10))
