import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Consolidador INGEMARS", layout="wide")

def clean_rut(rut):
    if pd.isna(rut): return None
    rut = str(rut).upper()
    return "".join(filter(lambda x: x.isdigit() or x == 'K', rut))

def parse_informe_con_secciones(df):
    """Extrae conceptos clasificándolos en HABERES o DESCUENTOS"""
    data = []
    current_section = 'HABERES' # Por defecto
    current_category = None
    
    for i in range(len(df)):
        row = df.iloc[i]
        val_0 = str(row.iloc[0]) if pd.notna(row.iloc[0]) else None
        
        # Detectar cambios de sección en el informe
        if val_0 == 'HABERES':
            current_section = 'HABERES'
            continue
        if val_0 == 'DESCUENTOS':
            current_section = 'DESCUENTOS'
            continue
            
        val_2 = str(row.iloc[2]) if pd.notna(row.iloc[2]) else None
        val_10 = row.iloc[10]
        
        # Identificar el nombre del concepto
        if val_0 and not val_0.startswith('Total') and val_0 not in ['Rut', 'Nombre', 'Razón Social', 'R.U.T.', 'Dirección']:
             current_category = val_0
        
        # Guardar dato si hay RUT y Monto
        if val_2 and pd.notna(val_10) and '-' in val_2:
            data.append({
                'Rut': clean_rut(val_2),
                'Concepto': current_category,
                'Monto': val_10,
                'Seccion': current_section
            })
    return pd.DataFrame(data)

st.title("📊 Consolidador de Remuneraciones")
st.sidebar.header("Parámetros")
razon_social = st.sidebar.text_input("Razón Social", "INGEMARS")
mes_sel = st.sidebar.selectbox("Mes", ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"])
anio_sel = st.sidebar.number_input("Año", 2025)

file_libro = st.file_uploader("Subir Libro de Remuneraciones (.xlsx)", type=["xlsx"])
file_informe = st.file_uploader("Subir Informe Haberes y Descuentos (.xlsx)", type=["xlsx"])

if file_libro and file_informe:
    if st.button("Procesar y Ordenar"):
        try:
            df_libro = pd.read_excel(file_libro)
            df_informe_raw = pd.read_excel(file_informe)
            
            # 1. Procesar Informe
            df_parsed = parse_informe_con_secciones(df_informe_raw)
            
            # Listas para ordenar después
            haberes_det = df_parsed[df_parsed['Seccion'] == 'HABERES']['Concepto'].unique().tolist()
            descuentos_det = df_parsed[df_parsed['Seccion'] == 'DESCUENTOS']['Concepto'].unique().tolist()
            
            # 2. Pivotar
            df_pivot = df_parsed.pivot_table(index='Rut', columns='Concepto', values='Monto', aggfunc='sum').reset_index()
            
            # Evitar colisión de nombres (como Sueldo Base)
            for col in df_pivot.columns:
                if col != 'Rut' and col in df_libro.columns:
                    df_pivot = df_pivot.rename(columns={col: f"{col} (Detalle)"})
                    # Actualizar listas de orden
                    if col in haberes_det: haberes_det[haberes_det.index(col)] = f"{col} (Detalle)"
                    if col in descuentos_det: descuentos_det[descuentos_det.index(col)] = f"{col} (Detalle)"

            # 3. Cruzar
            col_rut_libro = [c for c in df_libro.columns if 'Rut' in str(c)][0]
            df_libro['rut_join'] = df_libro[col_rut_libro].apply(clean_rut)
            df_final = pd.merge(df_libro, df_pivot, left_on='rut_join', right_on='Rut', how='left').drop(columns=['rut_join', 'Rut'])

            # 4. ORDENAR COLUMNAS
            # Definir grupos
            cols_id = ['Rut Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres', 'Sueldo Base', 'N° Dias Trabajados']
            
            # Identificar Haberes y Descuentos del LIBRO original por palabras clave
            haberes_libro = [c for c in df_libro.columns if any(k in str(c) for k in ['Gratificacion', 'Haberes', 'Movilizacion', 'Colacion', 'Cargas']) and c not in cols_id]
            descuentos_libro = [c for c in df_libro.columns if any(k in str(c) for k in ['Prevision', 'Salud', 'Seguro', 'APV', 'Impuesto', 'Descuentos', 'Anticipos', 'Ahorro']) and c not in cols_id]
            
            resto = [c for c in df_final.columns if c not in cols_id and c not in haberes_libro and c not in descuentos_libro and c not in haberes_det and c not in descuentos_det]

            # Nuevo orden: ID -> Todos los Haberes -> Todos los Descuentos -> Resto/Totales
            column_order = cols_id + haberes_libro + haberes_det + descuentos_libro + descuentos_det + resto
            
            # Eliminar duplicados manteniendo el orden
            final_columns = []
            for c in column_order:
                if c in df_final.columns and c not in final_columns:
                    final_columns.append(c)
            
            df_final = df_final[final_columns]
            df_final.insert(0, 'Periodo', f"{mes_sel} {anio_sel}")
            df_final.insert(0, 'Empresa', razon_social)

            st.success("¡Consolidado generado con éxito!")
            
            # Descarga
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Consolidado')
            
            st.download_button(
                label="📥 Descargar Excel Ordenado",
                data=output.getvalue(),
                file_name=f"Consolidado_{razon_social}_{mes_sel}_{anio_sel}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.dataframe(df_final.head())
            
        except Exception as e:
            st.error(f"Error: {e}")
