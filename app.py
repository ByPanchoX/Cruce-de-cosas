import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Consolidador Remuneraciones - Solo Días", layout="wide")

# Inicializar estado de la sesión
if 'datos_acumulados' not in st.session_state:
    st.session_state['datos_acumulados'] = []

def clean_rut(rut):
    if pd.isna(rut): return None
    rut = str(rut).upper()
    return "".join(filter(lambda x: x.isdigit() or x == 'K', rut))

def parse_informe_con_secciones(df):
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
st.title("📊 Consolidador: Identificación + Días + Desglose Detallado")

with st.sidebar:
    st.header("1. Configuración")
    opcion_empresa = st.selectbox("Empresa", ["INGEMARS", "ENAP", "Otra..."])
    empresa_nombre = st.text_input("Nombre de la Empresa") if opcion_empresa == "Otra..." else opcion_empresa
    
    mes_sel = st.selectbox("Mes", ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"])
    anio_sel = st.number_input("Año", 2025)
    
    if st.button("🗑️ Limpiar Todo"):
        st.session_state['datos_acumulados'] = []
        st.rerun()

st.info("💡 Este programa filtrará el Libro de Remuneraciones para dejar solo las columnas de Días y lo cruzará con el Informe detallado.")

col1, col2 = st.columns(2)
with col1:
    file_libro = st.file_uploader("1. Libro de Remuneraciones (.xlsx)", type=["xlsx"])
with col2:
    file_informe = st.file_uploader("2. Informe Haberes y Descuentos (.xlsx)", type=["xlsx"])

if file_libro and file_informe:
    if st.button(f"➕ Procesar {empresa_nombre}"):
        try:
            df_libro_raw = pd.read_excel(file_libro)
            df_informe_raw = pd.read_excel(file_informe)
            
            # --- FILTRAR LIBRO PARA DEJAR SOLO IDENTIFICACIÓN Y DÍAS ---
            cols_identificacion = ['Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres']
            # Buscamos columnas que tengan la palabra "Dia" o "Día"
            cols_dias = [c for c in df_libro_raw.columns if 'dia' in str(c).lower() or 'día' in str(c).lower()]
            
            # Mantener solo las columnas deseadas del libro
            columnas_a_mantener = [c for c in (cols_identificacion + cols_dias) if c in df_libro_raw.columns]
            df_libro_filtrado = df_libro_raw[columnas_a_mantener].copy()
            
            # --- PROCESAR INFORME DETALLADO ---
            df_parsed = parse_informe_con_secciones(df_informe_raw)
            haberes_list = df_parsed[df_parsed['Seccion'] == 'HABERES']['Concepto'].unique().tolist()
            descuentos_list = df_parsed[df_parsed['Seccion'] == 'DESCUENTOS']['Concepto'].unique().tolist()
            
            df_pivot = df_parsed.pivot_table(index='Rut', columns='Concepto', values='Monto', aggfunc='sum').reset_index()
            
            # --- CRUZAR DATOS ---
            df_libro_filtrado['rut_join'] = df_libro_filtrado['Rut Trabajador'].apply(clean_rut)
            df_final = pd.merge(df_libro_filtrado, df_pivot, left_on='rut_join', right_on='Rut', how='left').drop(columns=['rut_join', 'Rut'])
            
            # Agregar metadatos
            df_final.insert(0, 'Año', anio_sel)
            df_final.insert(0, 'Mes', mes_sel)
            df_final.insert(0, 'Empresa', empresa_nombre)
            
            # Guardar en sesión
            st.session_state['datos_acumulados'].append({
                'df': df_final,
                'haberes': haberes_list,
                'descuentos': descuentos_list
            })
            st.success(f"Agregado: {empresa_nombre} ({mes_sel})")
            
        except Exception as e:
            st.error(f"Error: {e}")

# --- DESCARGA FINAL ---
if st.session_state['datos_acumulados']:
    st.divider()
    if st.button("🚀 GENERAR EXCEL FINAL CONSOLIDADO"):
        lista_dfs = [item['df'] for item in st.session_state['datos_acumulados']]
        df_total = pd.concat(lista_dfs, ignore_index=True)
        
        # Identificar todas las columnas de haberes y descuentos para el orden
        todos_haberes = []
        todos_descuentos = []
        for item in st.session_state['datos_acumulados']:
            todos_haberes.extend(item['haberes'])
            todos_descuentos.extend(item['descuentos'])
        
        # Ordenar columnas: ID -> Días -> Haberes -> Descuentos -> Resto
        cols_base = ['Empresa', 'Mes', 'Año', 'Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres']
        cols_dias = [c for c in df_total.columns if 'dia' in str(c).lower() or 'día' in str(c).lower()]
        
        # Limpiar listas de haberes/descuentos de duplicados y asegurar que existan en el df final
        hab_final = [c for c in sorted(list(set(todos_haberes))) if c in df_total.columns and c not in cols_base]
        des_final = [c for c in sorted(list(set(todos_descuentos))) if c in df_total.columns and c not in cols_base]
        
        otros = [c for c in df_total.columns if c not in cols_base and c not in cols_dias and c not in hab_final and c not in des_final]
        
        orden_final = []
        for c in (cols_base + cols_dias + hab_final + des_final + otros):
            if c not in orden_final: orden_final.append(c)
            
        df_total = df_total[orden_final]
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_total.to_excel(writer, index=False, sheet_name='Consolidado')
        
        st.download_button(label="📥 Descargar Excel Final", data=output.getvalue(), file_name="Consolidado_Remuneraciones.xlsx")
        st.dataframe(df_total)
