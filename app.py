import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Consolidador Multi-Empresa", layout="wide")

# Inicializar el estado de la sesión para guardar datos acumulados de los 12 meses/empresas
if 'datos_acumulados' not in st.session_state:
    st.session_state['datos_acumulados'] = []

def clean_rut(rut):
    if pd.isna(rut): return None
    rut = str(rut).upper()
    return "".join(filter(lambda x: x.isdigit() or x == 'K', rut))

def parse_informe_con_secciones(df):
    """Extrae conceptos clasificándolos en HABERES o DESCUENTOS"""
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
    st.header("1. Selección de Empresa")
    
    # Selector de empresa para evitar escribir el nombre
    opcion_empresa = st.selectbox("Seleccionar Empresa", ["INGEMARS", "ENAP", "Otra..."])
    
    if opcion_empresa == "Otra...":
        empresa_nombre = st.text_input("Escriba el nombre de la nueva empresa")
    else:
        empresa_nombre = opcion_empresa
        
    st.divider()
    st.header("2. Período")
    mes_sel = st.selectbox("Mes", ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"])
    anio_sel = st.number_input("Año", 2025)
    
    st.divider()
    if st.button("🗑️ Borrar lista y empezar de cero"):
        st.session_state['datos_acumulados'] = []
        st.rerun()

# Área de carga de archivos
st.subheader(f"Carga de archivos para: {empresa_nombre} ({mes_sel})")
col1, col2 = st.columns(2)
with col1:
    file_libro = st.file_uploader(f"1. Libro Remuneraciones (.xlsx)", type=["xlsx"])
with col2:
    file_informe = st.file_uploader(f"2. Informe Haberes/Descuentos (.xlsx)", type=["xlsx"])

if file_libro and file_informe:
    if st.button(f"➕ Procesar y Agregar a la lista"):
        try:
            df_libro = pd.read_excel(file_libro)
            df_informe_raw = pd.read_excel(file_informe)
            
            # Procesar Informe de desgloses
            df_parsed = parse_informe_con_secciones(df_informe_raw)
            
            # Pivotar y manejar nombres duplicados
            df_pivot = df_parsed.pivot_table(index='Rut', columns='Concepto', values='Monto', aggfunc='sum').reset_index()
            for col in df_pivot.columns:
                if col != 'Rut' and col in df_libro.columns:
                    df_pivot = df_pivot.rename(columns={col: f"{col} (Detalle)"})
            
            # Cruzar con el libro principal
            col_rut_libro = [c for c in df_libro.columns if 'Rut' in str(c)][0]
            df_libro['rut_join'] = df_libro[col_rut_libro].apply(clean_rut)
            df_res = pd.merge(df_libro, df_pivot, left_on='rut_join', right_on='Rut', how='left').drop(columns=['rut_join', 'Rut'])
            
            # Insertar columnas de periodo al principio
            df_res.insert(0, 'Año', anio_sel)
            df_res.insert(0, 'Mes', mes_sel)
            df_res.insert(0, 'Empresa', empresa_nombre)
            
            # Guardar en la lista acumulada
            st.session_state['datos_acumulados'].append(df_res)
            st.success(f"✅ {empresa_nombre} - {mes_sel} agregada a la lista.")
            
        except Exception as e:
            st.error(f"Error al procesar: {e}")

# --- SECCIÓN DE CONSOLIDACIÓN FINAL ---
if st.session_state['datos_acumulados']:
    st.divider()
    st.subheader("📋 Empresas y Meses cargados en esta sesión")
    
    # Resumen de carga
    resumen = []
    for d in st.session_state['datos_acumulados']:
        resumen.append({
            "Empresa": d['Empresa'].iloc[0], 
            "Mes": d['Mes'].iloc[0], 
            "Año": d['Año'].iloc[0],
            "N° Trabajadores": len(d)
        })
    st.table(pd.DataFrame(resumen))

    if st.button("🚀 GENERAR EXCEL FINAL CON TODO"):
        # Unir todos los DataFrames
        df_final_total = pd.concat(st.session_state['datos_acumulados'], ignore_index=True)
        
        # Lógica de Ordenamiento: Datos ID -> Haberes -> Descuentos -> Totales
        cols_id = ['Empresa', 'Mes', 'Año', 'Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres', 'Sueldo Base']
        
        # Agrupar columnas por tipo para el orden final
        haberes_cols = [c for c in df_final_total.columns if any(k in str(c).lower() for k in ['habere', 'bono', 'gratif', 'movili', 'colaci', 'asig', 'sobretiempo', 'viatico', 'incentivo']) and c not in cols_id]
        descuentos_cols = [c for c in df_final_total.columns if any(k in str(c).lower() for k in ['previ', 'salud', 'seguro', 'apv', 'impuesto', 'descuento', 'antici', 'ahorro', 'prestam', 'ccaf', 'isapre', 'fonasa', 'inp', 'pension']) and c not in cols_id]
        resto = [c for c in df_final_total.columns if c not in cols_id and c not in haberes_cols and c not in descuentos_cols]
        
        # Reordenar (eliminando posibles duplicados en la lista de nombres de columnas)
        final_order = []
        for col in (cols_id + haberes_cols + descuentos_cols + resto):
            if col in df_final_total.columns and col not in final_order:
                final_order.append(col)
        
        df_final_total = df_final_total[final_order]

        # Crear el archivo Excel en memoria
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final_total.to_excel(writer, index=False, sheet_name='Consolidado_Total')
        
        st.download_button(
            label="📥 Descargar ARCHIVO CONSOLIDADO TOTAL (.xlsx)",
            data=output.getvalue(),
            file_name=f"Consolidado_Remuneraciones_Final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.write("### Vista previa del archivo final:")
        st.dataframe(df_final_total.head(20))
