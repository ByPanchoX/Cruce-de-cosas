import streamlit as st
import pandas as pd
import io

# Configuración de la página
st.set_page_config(page_title="Consolidador Mensual INGEMARS", layout="wide")

def clean_rut(rut):
    if pd.isna(rut):
        return None
    rut = str(rut).upper()
    cleaned = "".join(filter(lambda x: x.isdigit() or x == 'K', rut))
    return cleaned

def parse_informe_haberes(df):
    """Extrae conceptos de haberes y descuentos del informe detallado"""
    data = []
    current_category = None
    
    for i in range(len(df)):
        row = df.iloc[i]
        val_0 = str(row.iloc[0]) if pd.notna(row.iloc[0]) else None
        val_2 = str(row.iloc[2]) if pd.notna(row.iloc[2]) else None
        val_10 = row.iloc[10]
        
        # Identificar el nombre del Bono o Descuento
        if val_0 and not val_0.startswith('Total') and val_0 not in ['HABERES', 'DESCUENTOS', 'Rut', 'Nombre', 'Razón Social']:
             current_category = val_0
        
        # Extraer monto si hay un RUT válido en la fila
        if val_2 and pd.notna(val_10) and '-' in val_2:
            data.append({
                'Rut': clean_rut(val_2),
                'Concepto': current_category,
                'Monto': val_10
            })
            
    return pd.DataFrame(data)

# --- INTERFAZ DE USUARIO ---
st.title("📊 Gestión de Remuneraciones Anual")
st.markdown("Carga los archivos mensuales para generar el desglose detallado por columnas.")

# Sidebar para metadatos
st.sidebar.header("Parámetros del Período")
razon_social = st.sidebar.text_input("Razón Social", value="INGEMARS")
mes_seleccionado = st.sidebar.selectbox("Mes", 
    ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
     "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"])
anio_seleccionado = st.sidebar.number_input("Año", min_value=2024, max_value=2030, value=2025)

col1, col2 = st.columns(2)
with col1:
    file_libro = st.file_uploader("1. Libro de Remuneraciones (.xlsx)", type=["xlsx"])
with col2:
    file_informe = st.file_uploader("2. Informe Haberes y Descuentos (.xlsx)", type=["xlsx"])

if file_libro and file_informe:
    if st.button(f"Generar Consolidado {mes_seleccionado} {anio_seleccionado}"):
        try:
            # Lectura de datos
            df_libro = pd.read_excel(file_libro)
            df_informe_raw = pd.read_excel(file_informe)
            
            # Procesamiento de desgloses
            df_detalles = parse_informe_haberes(df_informe_raw)
            df_pivot = df_detalles.pivot_table(
                index='Rut', 
                columns='Concepto', 
                values='Monto', 
                aggfunc='sum'
            ).reset_index()
            
            # Limpieza de RUT en el libro principal
            col_rut_libro = [c for c in df_libro.columns if 'Rut' in str(c)][0]
            df_libro['rut_clean'] = df_libro[col_rut_libro].apply(clean_rut)
            
            # Cruce de información
            df_final = pd.merge(df_libro, df_pivot, left_on='rut_clean', right_on='Rut', how='left')
            
            # --- NUEVO: Agregar columnas de periodo ---
            df_final.insert(0, 'Año', anio_seleccionado)
            df_final.insert(0, 'Mes', mes_seleccionado)
            df_final.insert(0, 'Empresa', razon_social)
            
            # Eliminar columnas auxiliares
            df_final = df_final.drop(columns=['rut_clean', 'Rut'])
            
            st.success(f"✅ Procesado con éxito: {len(df_final)} trabajadores encontrados.")

            # Preparar descarga Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name=f'{mes_seleccionado}_{anio_seleccionado}')
            
            st.download_button(
                label=f"📥 Descargar Excel {mes_seleccionado}",
                data=output.getvalue(),
                file_name=f"Remuneraciones_{razon_social}_{mes_seleccionado}_{anio_seleccionado}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.write("### Vista previa de los datos")
            st.dataframe(df_final.head(10))
            
        except Exception as e:
            st.error(f"Se produjo un error: {e}")
