import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="Consolidador de Remuneraciones", layout="wide")

def clean_rut(rut):
    if pd.isna(rut):
        return None
    rut = str(rut).upper()
    # Mantiene solo números y la letra K
    cleaned = "".join(filter(lambda x: x.isdigit() or x == 'K', rut))
    return cleaned

def parse_informe_haberes(df):
    """Extrae cada concepto (bono/descuento) y lo asigna a cada RUT"""
    data = []
    current_category = None
    
    # Identificar las columnas correctas basándose en los datos
    # Según el ejemplo: Col 0 es Categoría, Col 2 es Rut, Col 10 es Monto
    for i, row in df.iterrows():
        val_0 = str(row.iloc[0]) if pd.notna(row.iloc[0]) else None
        val_2 = str(row.iloc[2]) if pd.notna(row.iloc[2]) else None
        val_10 = row.iloc[10]
        
        # Si la columna 0 tiene texto y no es un encabezado o total, es una nueva categoría
        if val_0 and not val_0.startswith('Total') and val_0 not in ['HABERES', 'DESCUENTOS', 'Rut', 'Nombre', 'Razón Social']:
             current_category = val_0
        
        # Si hay un RUT y un monto, guardamos el registro
        if val_2 and pd.notna(val_10) and '-' in val_2:
            data.append({
                'Rut': clean_rut(val_2),
                'Concepto': current_category,
                'Monto': val_10
            })
            
    return pd.DataFrame(data)

st.title("🚀 Consolidador de Remuneraciones")
st.markdown("""
Esta herramienta toma el **Libro de Remuneraciones** y le añade el desglose detallado del **Informe de Haberes y Descuentos**.
""")

# --- Carga de Archivos ---
col1, col2 = st.columns(2)
with col1:
    file_libro = st.file_uploader("Subir Libro de Remuneraciones (CSV)", type=["csv"])
with col2:
    file_informe = st.file_uploader("Subir Informe Haberes y Descuentos (CSV)", type=["csv"])

if file_libro and file_informe:
    if st.button("Generar Libro Consolidado"):
        try:
            # Leer archivos
            df_libro = pd.read_csv(file_libro)
            df_informe_raw = pd.read_csv(file_informe)
            
            # 1. Procesar el informe detallado
            df_detalles = parse_informe_haberes(df_informe_raw)
            
            # 2. Pivotar (convertir filas de bonos en columnas)
            df_pivot = df_detalles.pivot_table(
                index='Rut', 
                columns='Concepto', 
                values='Monto', 
                aggfunc='sum'
            ).reset_index()
            
            # Renombrar columnas del desglose para evitar confusiones
            df_pivot.columns = [f"Detalle - {c}" if c != 'Rut' else c for c in df_pivot.columns]
            
            # 3. Limpiar RUTs en el libro principal para el cruce
            # Buscamos la columna que contenga "Rut"
            col_rut_libro = [c for c in df_libro.columns if 'Rut' in c][0]
            df_libro['rut_clean'] = df_libro[col_rut_libro].apply(clean_rut)
            
            # 4. Cruzar (Merge)
            df_final = pd.merge(df_libro, df_pivot, left_on='rut_clean', right_on='Rut', how='left')
            df_final = df_final.drop(columns=['rut_clean', 'Rut'])
            
            st.success("¡Procesamiento completado!")
            
            # Vista previa
            st.subheader("Vista Previa (Primeras 5 filas)")
            st.dataframe(df_final.head())
            
            # Descarga
            csv = df_final.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label="📥 Descargar Libro Consolidado CSV",
                data=csv,
                file_name="Libro_Remuneraciones_Consolidado.csv",
                mime="text/csv",
            )
            
        except Exception as e:
            st.error(f"Error al procesar: {e}")
else:
    st.info("Por favor, sube ambos archivos para comenzar.")
