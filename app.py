import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Consolidador Exacto", layout="wide")
st.title("📊 Consolidador de Remuneraciones (Plantilla Exacta)")

# Limpiar RUT
def limpiar_rut(rut):
    if pd.isna(rut): return rut
    return str(rut).replace(".", "").strip().upper()

# Extraer conceptos desde el Informe de Haberes y Descuentos
def procesar_informe_hd(file):
    df = pd.read_csv(file, header=None) if file.name.endswith('.csv') else pd.read_excel(file, header=None)
    records = []
    concept_type = None 
    
    for _, row in df.iterrows():
        col0 = str(row[0]).strip()
        col2 = str(row[2]).strip()
        col1 = str(row[1]).strip()
        col10 = str(row[10]).strip()
        
        if col0 == "HABERES": concept_type = "H"
        elif col0 == "DESCUENTOS": concept_type = "D"
            
        if col0 not in ["", "nan", "HABERES", "DESCUENTOS", "Razón Social", "R.U.T.", "Dirección"] and len(col0) > 2:
            current_concept = col0
            
            rut_candidate = col2 if ("-" in col2 and len(col2) > 6 and "Página" not in col2) else (col1 if "-" in col1 and len(col1) > 6 and "Página" not in col1 else "")
            
            if rut_candidate:
                rut = limpiar_rut(rut_candidate)
                try: amount = float(col10.replace(".", ""))
                except: 
                    try: amount = float(str(row[9]).replace(".", ""))
                    except: amount = 0
                    
                records.append({"Rut Trabajador": rut, "Concepto": f"{current_concept} ({concept_type})", "Monto": amount})
                
    if not records: return pd.DataFrame(columns=["Rut Trabajador"])
    
    df_pivot = pd.DataFrame(records).pivot_table(index="Rut Trabajador", columns="Concepto", values="Monto", aggfunc='sum').reset_index()
    return df_pivot.fillna(0)

# Procesar una empresa completa
def procesar_empresa(rut_empresa, libro_file, aporte_file, hd_file):
    l_reader = pd.read_csv if libro_file.name.endswith('.csv') else pd.read_excel
    df_libro = l_reader(libro_file)
    
    a_reader = pd.read_csv if aporte_file.name.endswith('.csv') else pd.read_excel
    df_aporte = a_reader(aporte_file)
    
    df_hd = procesar_informe_hd(hd_file)
    
    df_libro['Rut Trabajador'] = df_libro['Rut Trabajador'].apply(limpiar_rut)
    df_aporte['Rut Trabajador'] = df_aporte['Rut Trabajador'].apply(limpiar_rut)
    
    cols_base = ['RUT EMPRESA', 'MES', 'AÑO', 'Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres', 'Cargo', 
                 'Tipo de Contrato', 'N° Dias Trabajados', 'N° Dias Ausentes', 'N° Dias Licencia', 'N° Dias Accidentes de Trabajo', 
                 'N° Dias No Contratado', 'N° Cargas Familiares']
    
    df_temp = pd.DataFrame()
    df_temp['RUT EMPRESA'] = [rut_empresa] * len(df_libro)
    df_temp['MES'] = 'Diciembre'
    df_temp['AÑO'] = df_aporte['Ano'].iloc[0] if 'Ano' in df_aporte.columns else 2025
    
    for col in cols_base[3:]:
        if col in df_libro.columns:
            df_temp[col] = df_libro[col]
        elif col in df_aporte.columns:
            mapa = df_aporte.set_index('Rut Trabajador')[col].to_dict()
            df_temp[col] = df_temp['Rut Trabajador'].map(mapa)
        else:
            df_temp[col] = 0
            
    if 'Imponible' in df_libro.columns:
        df_temp['Total Imponible (H)'] = df_libro['Imponible']

    df_merged = pd.merge(df_temp, df_hd, on='Rut Trabajador', how='left')
    
    cols_p = ['Aporte Accidentes de Trabajo IPS', 'Aporte Accidentes de Trabajo Mutual', 'Expectativa de Vida Seguro Social', 
              'Seguro de Cesantía Empleador', 'Seguro de Invalidez y Sobrevivencia Empleador', 'Adicional de Capitalización Individual AFP']
    
    for col in cols_p:
        if col in df_aporte.columns:
            mapa_p = df_aporte.set_index('Rut Trabajador')[col].to_dict()
            df_merged[f"{col} (P)"] = df_merged['Rut Trabajador'].map(mapa_p)
            
    return df_merged

# --- INTERFAZ ---
st.info("Sube tu archivo 'Consolidado_Final_Diciembre' para que el sistema copie sus columnas exactas.")
plantilla = st.file_uploader("📥 Plantilla Original", type=['xlsx', 'csv'], key="plantilla")

st.divider()

col1, col2 = st.columns(2)
with col1:
    st.subheader("Archivos ENAP")
    rut_enap = st.text_input("RUT ENAP", value="76.455.680-1", key="r_enap")
    l_enap = st.file_uploader("Libro ENAP", type=['xlsx', 'csv'], key="l_enap")
    a_enap = st.file_uploader("Aporte ENAP", type=['xlsx', 'csv'], key="a_enap")
    h_enap = st.file_uploader("Inf. Haberes ENAP", type=['xlsx', 'csv'], key="h_enap")

with col2:
    st.subheader("Archivos INGEMARS")
    rut_ing = st.text_input("RUT INGEMARS", value="76.455.680-1", key="r_ing")
    l_ing = st.file_uploader("Libro INGEMARS", type=['xlsx', 'csv'], key="l_ing")
    a_ing = st.file_uploader("Aporte INGEMARS", type=['xlsx', 'csv'], key="a_ing")
    h_ing = st.file_uploader("Inf. Haberes INGEMARS", type=['xlsx', 'csv'], key="h_ing")

if st.button("🚀 Generar Archivo Exacto", type="primary", use_container_width=True):
    if not plantilla:
        st.error("⚠️ Sube tu plantilla primero.")
    elif l_enap and a_enap and h_enap and l_ing and a_ing and h_ing:
        with st.spinner("Procesando y alineando columnas..."):
            
            # 1. LECTOR INTELIGENTE DE PLANTILLA
            p_reader = pd.read_csv if plantilla.name.endswith('.csv') else pd.read_excel
            df_plantilla_raw = p_reader(plantilla, header=None) # Leer sin encabezados primero
            
            # Buscar en qué fila están realmente los encabezados
            fila_encabezados = 0
            for idx, row in df_plantilla_raw.iterrows():
                # Revisar si en esta fila aparece la palabra RUT EMPRESA o RUT TRABAJADOR
                valores_fila = [str(v).strip().upper() for v in row.values if pd.notna(v)]
                if "RUT EMPRESA" in valores_fila or "RUT TRABAJADOR" in valores_fila:
                    fila_encabezados = idx
                    break
            
            # Volver a leer la plantilla sabiendo en qué fila empezar
            df_plantilla = p_reader(plantilla, header=fila_encabezados)
            
            # Filtrar columnas basura ("Unnamed")
            columnas_maestras = [col for col in df_plantilla.columns if "Unnamed" not in str(col)]
            
            # 2. Procesar datos
            df_enap_procesado = procesar_empresa(rut_enap, l_enap, a_enap, h_enap)
            df_ing_procesado = procesar_empresa(rut_ing, l_ing, a_ing, h_ing)
            
            df_consolidado = pd.concat([df_enap_procesado, df_ing_procesado], ignore_index=True)
            
            # 3. Llenar la plantilla
            df_final = pd.DataFrame(columns=columnas_maestras)
            
            for col in columnas_maestras:
                if col in df_consolidado.columns:
                    df_final[col] = df_consolidado[col]
                else:
                    df_final[col] = 0 
                    
            df_final = df_final.fillna(0)
            
            st.success("✅ ¡Listo! Columnas estructuradas idénticamente a la plantilla original.")
            st.dataframe(df_final.head())
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, sheet_name='Sheet1', index=False)
            
            st.download_button(
                label="📥 Descargar Consolidado Exacto",
                data=buffer,
                file_name="Consolidado_Final_Remuneraciones.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("⚠️ Faltan archivos de las empresas.")
