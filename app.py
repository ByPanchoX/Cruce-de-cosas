import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Consolidador de Remuneraciones V2", layout="wide")

st.title("📊 Consolidador de Archivos de Remuneraciones")
st.write("Sube los **3 archivos** de cada empresa para generar el Consolidado Final exacto.")

# Función para limpiar el RUT
def limpiar_rut(rut):
    if pd.isna(rut):
        return rut
    return str(rut).replace(".", "").strip().upper()

# Lector inteligente para el "Informe Haberes y Descuentos"
def procesar_informe_hd(file):
    df = pd.read_csv(file, header=None) if file.name.endswith('.csv') else pd.read_excel(file, header=None)
    records = []
    current_concept = None
    concept_type = None 
    
    for _, row in df.iterrows():
        col0 = str(row[0]).strip()
        col2 = str(row[2]).strip()
        col10 = str(row[10]).strip()
        
        if col0 == "HABERES":
            concept_type = "H"
        elif col0 == "DESCUENTOS":
            concept_type = "D"
            
        if col0 not in ["nan", "HABERES", "DESCUENTOS", "Razón Social", "R.U.T.", "Dirección"]:
            current_concept = col0
            
        if "-" in col2 and len(col2) > 6 and "Página" not in col2:
            rut = limpiar_rut(col2)
            amount = col10.replace(".", "")
            try:
                amount = float(amount)
            except:
                amount = 0
                
            records.append({
                "Rut Trabajador": rut,
                "Concepto": f"{current_concept} ({concept_type})",
                "Monto": amount
            })
            
    if not records:
        return pd.DataFrame(columns=["Rut Trabajador"])
        
    df_records = pd.DataFrame(records)
    # Pivotear para que cada concepto sea una columna
    df_pivot = df_records.pivot_table(index="Rut Trabajador", columns="Concepto", values="Monto", aggfunc='sum').reset_index()
    df_pivot = df_pivot.fillna(0)
    return df_pivot

# Lector normal
def leer_base(file):
    df = pd.read_csv(file) if file.name.endswith('.csv') else pd.read_excel(file)
    if 'Rut Trabajador' in df.columns:
        df['Rut Trabajador'] = df['Rut Trabajador'].apply(limpiar_rut)
    return df

# Procesador central
def generar_consolidado(libros, aportes, informes, rut_empresas):
    df_finales = []
    
    for i in range(len(libros)):
        df_libro = leer_base(libros[i])
        df_aporte = leer_base(aportes[i])
        df_hd = procesar_informe_hd(informes[i])
        
        # 1. Agregar RUT de la empresa, Mes y Año según Aporte Patronal
        df_libro.insert(0, 'RUT EMPRESA', rut_empresas[i])
        df_libro.insert(1, 'MES', df_aporte['Mes'].iloc[0] if 'Mes' in df_aporte else '12')
        df_libro.insert(2, 'AÑO', df_aporte['Ano'].iloc[0] if 'Ano' in df_aporte else '2025')
        
        # Si no existe "Tipo de Contrato" en el libro, lo traemos del aporte
        if 'Tipo de Contrato' not in df_libro.columns and 'Tipo de Contrato' in df_aporte.columns:
            df_libro = pd.merge(df_libro, df_aporte[['Rut Trabajador', 'Tipo de Contrato']], on='Rut Trabajador', how='left')
        
        # 2. Renombrar Imponible si es necesario
        if 'Imponible' in df_libro.columns:
            df_libro.rename(columns={'Imponible': 'Total Imponible (H)'}, inplace=True)
            
        # 3. Extraer columnas Patronales (P) del Aporte
        cols_aportes = ['Rut Trabajador', 'Aporte Accidentes de Trabajo IPS', 'Aporte Accidentes de Trabajo Mutual', 
                        'Expectativa de Vida Seguro Social', 'Seguro de Cesantía Empleador', 
                        'Seguro de Invalidez y Sobrevivencia Empleador', 'Adicional de Capitalización Individual AFP']
        cols_aportes_existentes = [c for c in cols_aportes if c in df_aporte.columns]
        df_aporte_filtrado = df_aporte[cols_aportes_existentes].copy()
        
        # Añadir (P) al final del nombre de las columnas patronales
        renames_p = {c: f"{c} (P)" for c in cols_aportes_existentes if c != 'Rut Trabajador'}
        df_aporte_filtrado.rename(columns=renames_p, inplace=True)
        
        # 4. Cruce Maestro: Libro + Haberes/Descuentos + Aporte Patronal
        df_cruce = pd.merge(df_libro, df_hd, on='Rut Trabajador', how='left')
        df_cruce = pd.merge(df_cruce, df_aporte_filtrado, on='Rut Trabajador', how='left')
        
        df_finales.append(df_cruce)
        
    # Unir la data de ENAP e INGEMARS
    df_consolidado = pd.concat(df_finales, ignore_index=True)
    df_consolidado = df_consolidado.fillna(0)
    
    # Ordenar las columnas básicas al inicio (Las mismas de tu plantilla original)
    cols_base = ['RUT EMPRESA', 'MES', 'AÑO', 'Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres', 'Cargo', 
                 'Tipo de Contrato', 'N° Dias Trabajados', 'N° Dias Ausentes', 'N° Dias Licencia', 'N° Dias Accidentes de Trabajo', 
                 'N° Dias No Contratado', 'N° Cargas Familiares']
    
    cols_existentes_base = [c for c in cols_base if c in df_consolidado.columns]
    cols_restantes = sorted([c for c in df_consolidado.columns if c not in cols_existentes_base])
    
    return df_consolidado[cols_existentes_base + cols_restantes]


# --- INTERFAZ DE USUARIO ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("Archivos ENAP")
    rut_enap = st.text_input("RUT ENAP", value="76.455.680-1", key="r_enap")
    l_enap = st.file_uploader("Libro de Remuneraciones", type=['xlsx', 'csv'], key="l_enap")
    a_enap = st.file_uploader("Aporte Patronal", type=['xlsx', 'csv'], key="a_enap")
    h_enap = st.file_uploader("Informe Haberes y Descuentos", type=['xlsx', 'csv'], key="h_enap")

with col2:
    st.subheader("Archivos INGEMARS")
    rut_ing = st.text_input("RUT INGEMARS", value="76.455.680-1", key="r_ing")
    l_ing = st.file_uploader("Libro de Remuneraciones", type=['xlsx', 'csv'], key="l_ing")
    a_ing = st.file_uploader("Aporte Patronal", type=['xlsx', 'csv'], key="a_ing")
    h_ing = st.file_uploader("Informe Haberes y Descuentos", type=['xlsx', 'csv'], key="h_ing")

if st.button("🚀 Generar Consolidado Perfecto", type="primary", use_container_width=True):
    if l_enap and a_enap and h_enap and l_ing and a_ing and h_ing:
        with st.spinner("Desglosando Haberes, Descuentos y Aportes..."):
            
            # Pasamos las listas emparejadas
            libros = [l_enap, l_ing]
            aportes = [a_enap, a_ing]
            informes = [h_enap, h_ing]
            ruts = [rut_enap, rut_ing]
            
            df_final = generar_consolidado(libros, aportes, informes, ruts)
            
            st.success("✅ ¡Consolidado generado con éxito en el formato solicitado!")
            
            st.dataframe(df_final.head())
            
            # Preparar descarga
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, sheet_name='Sheet1', index=False)
            
            st.download_button(
                label="📥 Descargar Consolidado_Final_Diciembre.xlsx",
                data=buffer,
                file_name="Consolidado_Final_Diciembre.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.error("⚠️ Faltan archivos. Debes subir el Libro, Aporte e Informe de Haberes de AMBAS empresas.")
