import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Consolidador Remuneraciones", layout="wide")

# Inicializar la lista de datos si no existe
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
st.title("📊 Consolidador: Identificación + Días + Desglose")

with st.sidebar:
    st.header("Configuración")
    empresa_opcion = st.selectbox("Empresa", ["INGEMARS", "ENAP", "Otra..."])
    empresa_final = st.text_input("Nombre Personalizado") if empresa_opcion == "Otra..." else empresa_opcion
    
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
    if st.button(f"➕ Agregar {empresa_final} a la lista"):
        try:
            df_libro_raw = pd.read_excel(file_libro)
            df_informe_raw = pd.read_excel(file_informe)
            
            # --- FILTRO DE DÍAS ---
            cols_id = ['Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres']
            cols_dias = [c for c in df_libro_raw.columns if 'dia' in str(c).lower() or 'día' in str(c).lower()]
            columnas_finales_libro = [c for c in (cols_id + cols_dias) if c in df_libro_raw.columns]
            df_libro_solo_dias = df_libro_raw[columnas_finales_libro].copy()
            
            # --- PROCESAR INFORME ---
            df_parsed = parse_informe_con_secciones(df_informe_raw)
            hab_list = df_parsed[df_parsed['Seccion'] == 'HABERES']['Concepto'].unique().tolist()
            des_list = df_parsed[df_parsed['Seccion'] == 'DESCUENTOS']['Concepto'].unique().tolist()
            
            df_pivot = df_parsed.pivot_table(index='Rut', columns='Concepto', values='Monto', aggfunc='sum').reset_index()
            
            # --- UNIÓN ---
            # Identificar columna RUT en el libro
            col_rut_libro = [c for c in df_libro_solo_dias.columns if 'Rut' in str(c)][0]
            df_libro_solo_dias['rut_key'] = df_libro_solo_dias[col_rut_libro].apply(clean_rut)
            
            df_merged = pd.merge(df_libro_solo_dias, df_pivot, left_on='rut_key', right_on='Rut', how='left').drop(columns=['rut_key', 'Rut'])
            
            # Insertar metadatos
            df_merged.insert(0, 'Año', anio_sel)
            df_merged.insert(0, 'Mes', mes_sel)
            df_merged.insert(0, 'Empresa', empresa_final)
            
            # GUARDAR COMO DICCIONARIO
            st.session_state['datos_acumulados'].append({
                'df': df_merged,
                'haberes': hab_list,
                'descuentos': des_list
            })
            st.success(f"✅ Agregado: {empresa_final} ({mes_sel})")
            
        except Exception as e:
            st.error(f"Error al procesar: {e}")

# --- BOTÓN FINAL DE DESCARGA ---
if st.session_state['datos_acumulados']:
    st.divider()
    st.subheader("Empresas listas para descargar:")
    resumen = [{"Empresa": i['df']['Empresa'].iloc[0], "Mes": i['df']['Mes'].iloc[0]} for i in st.session_state['datos_acumulados']]
    st.table(resumen)

    if st.button("🚀 GENERAR EXCEL FINAL CONSOLIDADO"):
        # Extraer solo los DataFrames de la lista de diccionarios
        lista_solo_dfs = [item['df'] for item in st.session_state['datos_acumulados']]
        df_total = pd.concat(lista_solo_dfs, ignore_index=True)
        
        # Ordenar columnas
        ids = ['Empresa', 'Mes', 'Año', 'Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres']
        dias = [c for c in df_total.columns if 'dia' in str(c).lower() or 'día' in str(c).lower()]
        
        # Recolectar nombres de haberes y descuentos
        h_cols = []
        d_cols = []
        for item in st.session_state['datos_acumulados']:
            h_cols.extend(item['haberes'])
            d_cols.extend(item['descuentos'])
        
        h_final = [c for c in sorted(list(set(h_cols))) if c in df_total.columns and c not in ids]
        d_final = [c for c in sorted(list(set(d_cols))) if c in df_total.columns and c not in ids]
        
        resto = [c for c in df_total.columns if c not in (ids + dias + h_final + d_final)]
        
        orden_final = []
        for c in (ids + dias + h_final + d_final + resto):
            if c not in orden_final: orden_final.append(c)
            
        df_total = df_total[orden_final]
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_total.to_excel(writer, index=False, sheet_name='Consolidado')
        
        st.download_button(label="📥 Descargar Excel Final", data=output.getvalue(), file_name="Consolidado_Final.xlsx")
