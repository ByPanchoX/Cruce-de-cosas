import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Consolidador de Remuneraciones", layout="wide")

st.title("📊 Consolidador de Archivos de Remuneraciones")
st.write("Sube los archivos correspondientes a **ENAP** e **INGEMARS** para generar el consolidado final.")

# Función para limpiar el RUT y usarlo como llave de cruce
def limpiar_rut(rut):
    if pd.isna(rut):
        return rut
    return str(rut).replace(".", "").strip().upper()

# Función para procesar y consolidar los datos
def procesar_datos(libros, aportes):
    df_libros = []
    df_aportes = []
    
    # Concatenar todos los libros de remuneraciones
    for file in libros:
        df = pd.read_csv(file) if file.name.endswith('.csv') else pd.read_excel(file)
        if 'Rut Trabajador' in df.columns:
            df['Rut Trabajador'] = df['Rut Trabajador'].apply(limpiar_rut)
        df_libros.append(df)
        
    # Concatenar todos los aportes patronales
    for file in aportes:
        df = pd.read_csv(file) if file.name.endswith('.csv') else pd.read_excel(file)
        if 'Rut Trabajador' in df.columns:
            df['Rut Trabajador'] = df['Rut Trabajador'].apply(limpiar_rut)
        df_aportes.append(df)

    if df_libros and df_aportes:
        consolidado_libros = pd.concat(df_libros, ignore_index=True)
        consolidado_aportes = pd.concat(df_aportes, ignore_index=True)

        # Seleccionar columnas de interés de Aportes para no duplicar columnas al cruzar
        cols_aportes = [
            'Rut Trabajador', 'Tipo de Contrato', 'Aporte Accidentes de Trabajo IPS', 
            'Aporte Accidentes de Trabajo Mutual', 'Expectativa de Vida Seguro Social',
            'Seguro de Cesantía Empleador', 'Seguro de Invalidez y Sobrevivencia Empleador',
            'Adicional de Capitalización Individual AFP'
        ]
        
        # Filtrar solo si las columnas existen en el dataframe
        cols_aportes_existentes = [c for c in cols_aportes if c in consolidado_aportes.columns]
        consolidado_aportes = consolidado_aportes[cols_aportes_existentes]

        # Hacer Merge (Cruce) usando el RUT del trabajador
        df_final = pd.merge(consolidado_libros, consolidado_aportes, on='Rut Trabajador', how='left')
        
        # Eliminar duplicados si es que existen
        df_final = df_final.drop_duplicates(subset=['Rut Trabajador'])
        
        return df_final
    return None

# --- INTERFAZ DE USUARIO ---

col1, col2 = st.columns(2)

with col1:
    st.subheader("Archivos ENAP")
    libro_enap = st.file_uploader("Libro de Remuneraciones ENAP", type=['xlsx', 'csv'], key="l_enap")
    aporte_enap = st.file_uploader("Aporte Patronal ENAP", type=['xlsx', 'csv'], key="a_enap")

with col2:
    st.subheader("Archivos INGEMARS")
    libro_ingemars = st.file_uploader("Libro de Remuneraciones INGEMARS", type=['xlsx', 'csv'], key="l_ing")
    aporte_ingemars = st.file_uploader("Aporte Patronal INGEMARS", type=['xlsx', 'csv'], key="a_ing")


if st.button("Generar Consolidado", type="primary"):
    libros_subidos = [f for f in [libro_enap, libro_ingemars] if f is not None]
    aportes_subidos = [f for f in [aporte_enap, aporte_ingemars] if f is not None]
    
    if not libros_subidos or not aportes_subidos:
        st.warning("⚠️ Debes subir al menos un Libro de Remuneraciones y un Aporte Patronal para consolidar.")
    else:
        with st.spinner("Procesando y cruzando la información..."):
            df_consolidado = procesar_datos(libros_subidos, aportes_subidos)
            
            if df_consolidado is not None:
                st.success("✅ ¡Consolidado generado con éxito!")
                
                # Vista previa de los datos
                st.write(f"Total de registros procesados: **{len(df_consolidado)}**")
                st.dataframe(df_consolidado.head())
                
                # Crear buffer para descargar en Excel
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_consolidado.to_excel(writer, sheet_name='Consolidado', index=False)
                
                # Botón de descarga
                st.download_button(
                    label="📥 Descargar Excel Consolidado",
                    data=buffer,
                    file_name="Consolidado_Final_Remuneraciones.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
