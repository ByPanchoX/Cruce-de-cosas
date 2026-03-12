import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Consolidador Exacto v4", layout="wide")
st.title("📊 Consolidador de Remuneraciones (Formato Planilla Nueva)")

# --- BARRA LATERAL: PARÁMETROS AJUSTABLES ---
st.sidebar.header("⚙️ Parámetros Mensuales")
st.sidebar.write("Ajusta los valores que varían mes a mes:")

razon_social_input = st.sidebar.text_input("Rut Razón Social", value="76455680-1")
razon_social_limpio = razon_social_input.replace(".", "").strip()
if "-" not in razon_social_limpio and len(razon_social_limpio) > 1:
    razon_social_limpio = razon_social_limpio[:-1] + "-" + razon_social_limpio[-1]

año_input = st.sidebar.text_input("Año (aaaa)", value="2025")
mes_input = st.sidebar.text_input("Mes (mm)", value="12")
año_mes_str = f"{año_input}{mes_input.zfill(2)}"

st.sidebar.subheader("Tasas Empleador (%)")
tasa_sis = st.sidebar.number_input("Tasa S.I.S. (%)", value=1.54, step=0.01)
tasa_mutual = st.sidebar.number_input("Tasa Mutual (%)", value=0.93, step=0.01)
tasa_cesantia = st.sidebar.number_input("Tasa Seguro Cesantía Empleador (%)", value=2.40, step=0.01)

# --- FUNCIONES DE PROCESAMIENTO ---
def limpiar_rut(rut):
    if pd.isna(rut): return rut
    return str(rut).replace(".", "").strip().upper()

def procesar_informe_hd(file):
    df = pd.read_csv(file, header=None) if file.name.endswith('.csv') else pd.read_excel(file, header=None)
    records = []
    concept_type = None 
    current_concept = None # Variable global para evitar que se pierdan filas
    
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
        
        if rut_candidate and current_concept:
            rut = limpiar_rut(rut_candidate)
            try: amount = float(col10.replace(".", ""))
            except: 
                try: amount = float(str(row[9]).replace(".", ""))
                except: amount = 0
                
            nombre_concepto = f"{current_concept} ({concept_type})"
            if current_concept.upper() == "AGUINALDO":
                nombre_concepto = "Aguinaldo (h)" if concept_type == "H" else "anticipo D"
            
            records.append({"Rut *": rut, "Concepto": nombre_concepto, "Monto": amount})
                
    if not records: return pd.DataFrame(columns=["Rut *"])
    df_pivot = pd.DataFrame(records).pivot_table(index="Rut *", columns="Concepto", values="Monto", aggfunc='sum').reset_index()
    return df_pivot.fillna(0)

def procesar_empresa(libro_file, aporte_file, hd_file):
    # Leer Libro
    l_reader = pd.read_csv if libro_file.name.endswith('.csv') else pd.read_excel
    df_libro = l_reader(libro_file)
    df_libro['Rut Trabajador'] = df_libro['Rut Trabajador'].apply(limpiar_rut)
    
    # Leer Aporte Patronal
    a_reader = pd.read_csv if aporte_file.name.endswith('.csv') else pd.read_excel
    df_aporte = a_reader(aporte_file)
    df_aporte['Rut Trabajador'] = df_aporte['Rut Trabajador'].apply(limpiar_rut)
    
    # Leer Haberes y Descuentos
    df_hd = procesar_informe_hd(hd_file)
    
    # Mapeo base del Libro
    mapa_libro = {
        'Rut Trabajador': 'Rut *',
        'Salud': 'Total Isapre *',
        'Seguro de Cesantía': 'Seguro Cesantía Trabajador *',
        'Prevision': 'AFP *',
        'Líquido': 'Sueldo Liquido *',
        'Sueldo Base': 'Sueldo Base *',
        'Imponible': 'Remuneración Imponible *',
        'Total Haberes': 'Remuneración Total o Sueldo Bruto *',
        'Impuesto Unico': 'Impuestos *',
        'Haberes No Imponibles': 'Remuneración No Imponible *',
        'N° Dias Trabajados': 'Días Trabajados del Mes *',
        'N° Dias Ausentes': 'Días de Ausentismo *',
        'N° Dias Licencia': 'Días de licencia médica *',
        'Movilizacion': 'Movilización',
        'Colacion': 'Colación',
        'Cargas Familiares': 'Asignación familiar y Maternal'
    }
    
    df_base = df_libro.rename(columns=mapa_libro)
    cols_a_mantener = [col for col in mapa_libro.values() if col in df_base.columns]
    df_base = df_base[cols_a_mantener]
    
    # Mapeo del Aporte Patronal
    mapa_aporte = {
        'Rut Trabajador': 'Rut *',
        'Expectativa de Vida Seguro Social': 'Expectativa de Vida Patronal',
        'Adicional de Capitalización Individual AFP': 'Capitalización Individual AFP Patronal',
        'Cotización Empleador Trabajo Pesado': 'Trabajo Pesado Empleador Reliquidado'
    }
    df_aporte_renamed = df_aporte.rename(columns=mapa_aporte)
    cols_aporte = [col for col in mapa_aporte.values() if col in df_aporte_renamed.columns]
    df_aporte_base = df_aporte_renamed[cols_aporte]
    
    # Unir Libro + Aporte + Haberes/Descuentos
    df_merged = pd.merge(df_base, df_aporte_base, on='Rut *', how='left')
    df_merged = pd.merge(df_merged, df_hd, on='Rut *', how='left')
    
    return df_merged

# --- INTERFAZ PRINCIPAL ---
st.info("Sube tu archivo 'PLANILLA.xlsx' para extraer la estructura exacta solicitada.")
plantilla = st.file_uploader("📥 Nueva Plantilla Maestra", type=['xlsx', 'csv'], key="plantilla")

col1, col2 = st.columns(2)
with col1:
    st.subheader("Archivos ENAP")
    l_enap = st.file_uploader("Libro ENAP", type=['xlsx', 'csv'], key="l_enap")
    a_enap = st.file_uploader("Aporte ENAP", type=['xlsx', 'csv'], key="a_enap")
    h_enap = st.file_uploader("Inf. Haberes ENAP", type=['xlsx', 'csv'], key="h_enap")

with col2:
    st.subheader("Archivos INGEMARS")
    l_ing = st.file_uploader("Libro INGEMARS", type=['xlsx', 'csv'], key="l_ing")
    a_ing = st.file_uploader("Aporte INGEMARS", type=['xlsx', 'csv'], key="a_ing")
    h_ing = st.file_uploader("Inf. Haberes INGEMARS", type=['xlsx', 'csv'], key="h_ing")

if st.button("🚀 Generar Planilla Definitiva", type="primary", use_container_width=True):
    if not plantilla:
        st.error("⚠️ Sube tu PLANILLA.xlsx primero.")
    elif l_enap and a_enap and h_enap and l_ing and a_ing and h_ing:
        with st.spinner("Procesando y cruzando variables..."):
            
            # 1. Extraer columnas maestras
            p_reader = pd.read_csv if plantilla.name.endswith('.csv') else pd.read_excel
            df_plantilla_raw = p_reader(plantilla, header=None)
            
            fila_encabezados = 0
            for idx, row in df_plantilla_raw.iterrows():
                valores_fila = [str(v).strip().upper() for v in row.values if pd.notna(v)]
                if any("RUT RAZÓN SOCIAL" in v or "RUT RAZON SOCIAL" in v for v in valores_fila):
                    fila_encabezados = idx
                    break
                    
            df_plantilla = p_reader(plantilla, header=fila_encabezados)
            columnas_maestras = [col for col in df_plantilla.columns if "Unnamed" not in str(col)]
            
            # 2. Procesar datos
            df_enap = procesar_empresa(l_enap, a_enap, h_enap)
            df_ing = procesar_empresa(l_ing, a_ing, h_ing)
            df_consolidado = pd.concat([df_enap, df_ing], ignore_index=True).fillna(0)
            
            # 3. Aplicar campos ajustables
            df_consolidado['Rut Razón Social *'] = razon_social_limpio
            df_consolidado['Año_Mes (aaaamm) *'] = año_mes_str
            
            if 'Remuneración Imponible *' in df_consolidado.columns:
                df_consolidado['Seguro Invalidez y Supervivencia (SIS) *'] = (df_consolidado['Remuneración Imponible *'] * (tasa_sis / 100)).round(0)
                df_consolidado['Seguro Accidente de Trabajo *'] = (df_consolidado['Remuneración Imponible *'] * (tasa_mutual / 100)).round(0)
                df_consolidado['Seguro Cesantía Empleador *'] = (df_consolidado['Remuneración Imponible *'] * (tasa_cesantia / 100)).round(0)
            
            # 4. Alinear con Planilla
            df_final = pd.DataFrame(columns=columnas_maestras)
            
            for col in columnas_maestras:
                # Búsqueda flexible ignorando mayúsculas/minúsculas y espacios extra
                col_match = next((c for c in df_consolidado.columns if str(c).strip().lower() == str(col).strip().lower()), None)
                
                if col_match:
                    df_final[col] = df_consolidado[col_match]
                else:
                    # Intento de mapear los anticipos según las variantes que suelen aparecer
                    if str(col).strip().lower() in ["descuento anticipo", "anticipo d"] and "Anticipos (D)" in df_consolidado.columns:
                        df_final[col] = df_consolidado["Anticipos (D)"]
                    else:
                        df_final[col] = 0
                    
            df_final = df_final.fillna(0)
            
            st.success("✅ ¡Planilla generada! Datos mapeados y variables calculadas.")
            st.dataframe(df_final.head())
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, sheet_name='Sheet1', index=False)
            
            st.download_button(
                label="📥 Descargar Planilla Final",
                data=buffer,
                file_name="PLANILLA_Consolidada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("⚠️ Faltan archivos por subir. Deben estar los 6 de las empresas.")
