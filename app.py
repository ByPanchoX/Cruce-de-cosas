import streamlit as st
import pandas as pd
import io
import unicodedata
import re

st.set_page_config(page_title="Consolidador Remuneraciones", layout="wide")

if 'datos_acumulados' not in st.session_state:
    st.session_state['datos_acumulados'] = []

def normalizar_texto(texto):
    """Limpia tildes, espacios y mayúsculas para comparar nombres de columnas"""
    if pd.isna(texto): return ""
    texto = str(texto).lower().strip()
    # Quitar tildes
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    # Quitar todo lo que no sea letras o números
    texto = re.sub(r'[^a-z0-9]', '', texto)
    return texto

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
            
            # Creamos un mapa de nombres normalizados del Libro
            nombres_libro_norm = {normalizar_texto(c): c for c in (cols_id + cols_dias) if c in df_libro_raw.columns}
            df_libro_solo_dias = df_libro_raw[list(nombres_libro_norm.values())].copy()
            
            # --- PROCESAR INFORME ---
            df_parsed = parse_informe_con_secciones(df_informe_raw)
            
            # Pivotar el informe
            df_pivot = df_parsed.pivot_table(index='Rut', columns='Concepto', values='Monto', aggfunc='sum').reset_index()
            
            # EVITAR DUPLICADOS POR NORMALIZACIÓN
            nuevos_haberes = []
            nuevos_descuentos = []
            
            columnas_pivot_final = {'Rut': 'Rut'}
            for col in df_pivot.columns:
                if col == 'Rut': continue
                
                norm_col = normalizar_texto(col)
                seccion = df_parsed[df_parsed['Concepto'] == col]['Seccion'].iloc[0]
                
                # Si el nombre normalizado ya existe en el Libro (ej: "diasausentes")
                if norm_col in nombres_libro_norm:
                    nuevo_nombre = f"{col} (Monto)"
                else:
                    nuevo_nombre = col
                
                columnas_pivot_final[col] = nuevo_nombre
                if seccion == 'HABERES': nuevos_haberes.append(nuevo_nombre)
                else: nuevos_descuentos.append(nuevo_nombre)

            df_pivot = df_pivot.rename(columns=columnas_pivot_final)

            # --- UNIÓN ---
            col_rut_libro = [c for c in df_libro_solo_dias.columns if 'Rut' in str(c)][0]
            df_libro_solo_dias['rut_key'] = df_libro_solo_dias[col_rut_libro].apply(clean_rut)
            
            df_merged = pd.merge(df_libro_solo_dias, df_pivot, left_on='rut_key', right_on='Rut', how='left')
            df_merged = df_merged.drop(columns=['rut_key', 'Rut'])
            
            # Metadatos
            df_merged.insert(0, 'Año', anio_sel)
            df_merged.insert(0, 'Mes', mes_sel)
            df_merged.insert(0, 'Empresa', empresa_final)
            
            st.session_state['datos_acumulados'].append({
                'df': df_merged,
                'haberes': nuevos_haberes,
                'descuentos': nuevos_descuentos
            })
            st.success(f"✅ {empresa_final} agregada correctamente.")
            
        except Exception as e:
            st.error(f"Error: {e}")

# --- DESCARGA ---
if st.session_state['datos_acumulados']:
    st.divider()
    if st.button("🚀 GENERAR EXCEL FINAL"):
        df_total = pd.concat([item['df'] for item in st.session_state['datos_acumulados']], ignore_index=True)
        
        # Orden de columnas
        ids = ['Empresa', 'Mes', 'Año', 'Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres']
        dias = [c for c in df_total.columns if ('dia' in str(c).lower() or 'día' in str(c).lower()) and c not in ids]
        
        h_all = []
        d_all = []
        for item in st.session_state['datos_acumulados']:
            h_all.extend(item['haberes'])
            d_all.extend(item['descuentos'])
        
        h_f = [c for c in sorted(list(set(h_all))) if c in df_total.columns and c not in ids]
        d_f = [c for c in sorted(list(set(d_all))) if c in df_total.columns and c not in ids]
        
        orden = []
        for c in (ids + dias + h_f + d_f):
            if c in df_total.columns and c not in orden: orden.append(c)
        
        df_total = df_total[orden + [c for c in df_total.columns if c not in orden]]
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_total.to_excel(writer, index=False, sheet_name='Consolidado')
        
        nombre_xls = f"Consolidado_{mes_sel}_{anio_sel}.xlsx"
        st.download_button(label=f"📥 Descargar {nombre_xls}", data=output.getvalue(), file_name=nombre_xls)
        st.dataframe(df_total.head(10))
