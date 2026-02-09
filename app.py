import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Consolidador Multi-Empresa", layout="wide")

# Inicializar el estado de la sesión para guardar datos acumulados
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
st.title("📊 Consolidador Multi-Empresa e Inter-Anual")

with st.sidebar:
    st.header("1. Datos del Período Actual")
    empresa_nombre = st.text_input("Nombre de la Empresa", value="INGEMARS")
    mes_sel = st.selectbox("Mes", ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"])
    anio_sel = st.number_input("Año", 2025)
    
    st.divider()
    if st.button("🗑️ Borrar lista y empezar de cero"):
        st.session_state['datos_acumulados'] = []
        st.rerun()

# Área de carga
col1, col2 = st.columns(2)
with col1:
    file_libro = st.file_uploader(f"Libro Remuneraciones - {empresa_nombre}", type=["xlsx"])
with col2:
    file_informe = st.file_uploader(f"Informe Haberes/Descuentos - {empresa_nombre}", type=["xlsx"])

if file_libro and file_informe:
    if st.button(f"➕ Procesar y Agregar {empresa_nombre} a la lista"):
        try:
            df_libro = pd.read_excel(file_libro)
            df_informe_raw = pd.read_excel(file_informe)
            
            # Procesar Informe
            df_parsed = parse_informe_con_secciones(df_informe_raw)
            haberes_det = df_parsed[df_parsed['Seccion'] == 'HABERES']['Concepto'].unique().tolist()
            descuentos_det = df_parsed[df_parsed['Seccion'] == 'DESCUENTOS']['Concepto'].unique().tolist()
            
            # Pivotar y renombrar duplicados
            df_pivot = df_parsed.pivot_table(index='Rut', columns='Concepto', values='Monto', aggfunc='sum').reset_index()
            for col in df_pivot.columns:
                if col != 'Rut' and col in df_libro.columns:
                    df_pivot = df_pivot.rename(columns={col: f"{col} (Detalle)"})
            
            # Unir
            col_rut_libro = [c for c in df_libro.columns if 'Rut' in str(c)][0]
            df_libro['rut_join'] = df_libro[col_rut_libro].apply(clean_rut)
            df_res = pd.merge(df_libro, df_pivot, left_on='rut_join', right_on='Rut', how='left').drop(columns=['rut_join', 'Rut'])
            
            # Identificadores
            df_res.insert(0, 'Año', anio_sel)
            df_res.insert(0, 'Mes', mes_sel)
            df_res.insert(0, 'Empresa', empresa_nombre)
            
            # Guardar en la lista global
            st.session_state['datos_acumulados'].append(df_res)
            st.success(f"✅ {empresa_nombre} ({mes_sel}) agregada correctamente.")
            
        except Exception as e:
            st.error(f"Error: {e}")

# --- SECCIÓN DE CONSOLIDACIÓN FINAL ---
if st.session_state['datos_acumulados']:
    st.divider()
    st.subheader("📋 Empresas en la lista actual")
    
    # Mostrar resumen de lo que hay en la lista
    resumen = []
    for d in st.session_state['datos_acumulados']:
        resumen.append({"Empresa": d['Empresa'].iloc[0], "Mes": d['Mes'].iloc[0], "Empleados": len(d)})
    st.table(pd.DataFrame(resumen))

    if st.button("🚀 Generar EXCEL FINAL (Todas las empresas juntas)"):
        # Concatenar todo
        df_final_total = pd.concat(st.session_state['datos_acumulados'], ignore_index=True)
        
        # Ordenar columnas (Lógica de Haberes -> Descuentos)
        cols_id = ['Empresa', 'Mes', 'Año', 'Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres', 'Sueldo Base']
        haberes_cols = [c for c in df_final_total.columns if any(k in str(c).lower() for k in ['habere', 'bono', 'gratif', 'movili', 'colaci', 'asig']) and c not in cols_id]
        descuentos_cols = [c for c in df_final_total.columns if any(k in str(c).lower() for k in ['previ', 'salud', 'seguro', 'apv', 'impuesto', 'descuento', 'antici', 'ahorro', 'prestam']) and c not in cols_id]
        resto = [c for c in df_final_total.columns if c not in cols_id and c not in haberes_cols and c not in descuentos_cols]
        
        df_final_total = df_final_total[cols_id + haberes_cols + descuentos_cols + resto]

        # Descarga
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final_total.to_excel(writer, index=False, sheet_name='Consolidado_Total')
        
        st.download_button(
            label="📥 Descargar CONSOLIDADO TOTAL (.xlsx)",
            data=output.getvalue(),
            file_name="Remuneraciones_Consolidado_Total.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.dataframe(df_final_total)
