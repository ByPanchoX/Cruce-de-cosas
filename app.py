import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Consolidador de Remuneraciones Excel", layout="wide")

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
    
    # En Excel, a veces hay filas vacías al inicio; iteramos por posición
    for i in range(len(df)):
        row = df.iloc[i]
        val_0 = str(row.iloc[0]) if pd.notna(row.iloc[0]) else None
        val_2 = str(row.iloc[2]) if pd.notna(row.iloc[2]) else None
        val_10 = row.iloc[10]
        
        # Identificar cambio de categoría (Bono/Descuento)
        if val_0 and not val_0.startswith('Total') and val_0 not in ['HABERES', 'DESCUENTOS', 'Rut', 'Nombre', 'Razón Social']:
             current_category = val_0
        
        # Si detectamos un RUT (formato 12.345.678-9) y un monto en la columna 10
        if val_2 and pd.notna(val_10) and '-' in val_2:
            data.append({
                'Rut': clean_rut(val_2),
                'Concepto': current_category,
                'Monto': val_10
            })
            
    return pd.DataFrame(data)

st.title("🚀 Consolidador de Remuneraciones (Versión Excel)")

col1, col2 = st.columns(2)
with col1:
    file_libro = st.file_uploader("Subir Libro de Remuneraciones (Excel)", type=["xlsx"])
with col2:
    file_informe = st.file_uploader("Subir Informe Haberes y Descuentos (Excel)", type=["xlsx"])

if file_libro and file_informe:
    if st.button("Generar Libro Consolidado"):
        try:
            # Leer archivos Excel
            df_libro = pd.read_excel(file_libro)
            df_informe_raw = pd.read_excel(file_informe)
            
            # 1. Procesar desglose
            df_detalles = parse_informe_haberes(df_informe_raw)
            
            # 2. Pivotar
            df_pivot = df_detalles.pivot_table(
                index='Rut', 
                columns='Concepto', 
                values='Monto', 
                aggfunc='sum'
            ).reset_index()
            
            # Renombrar columnas para el desglose
            df_pivot.columns = [f"Detalle - {c}" if c != 'Rut' else c for c in df_pivot.columns]
            
            # 3. Limpiar RUTs en libro principal
            col_rut_libro = [c for c in df_libro.columns if 'Rut' in str(c)][0]
            df_libro['rut_clean'] = df_libro[col_rut_libro].apply(clean_rut)
            
            # 4. Cruce de datos
            df_final = pd.merge(df_libro, df_pivot, left_on='rut_clean', right_on='Rut', how='left')
            df_final = df_final.drop(columns=['rut_clean', 'Rut'])
            
            st.success("¡Archivo Excel procesado exitosamente!")
            
            # 5. Preparar descarga en formato Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Consolidado')
            processed_data = output.getvalue()
            
            st.download_button(
                label="📥 Descargar Libro Consolidado (.xlsx)",
                data=processed_data,
                file_name="Libro_Remuneraciones_Consolidado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.dataframe(df_final.head())
            
        except Exception as e:
            st.error(f"Error al procesar los archivos Excel: {e}")
