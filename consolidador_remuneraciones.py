import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import re

# Valores UF 2025
UF_2025 = {
    'Enero': 38284.86,
    'Febrero': 38647.94,
    'Marzo': 38800.00,
    'Abril': 39050.00,
    'Mayo': 39200.00,
    'Junio': 39267.07,
    'Julio': 39179.01,
    'Agosto': 39383.07,
    'Septiembre': 39500.00,
    'Octubre': 39597.67,
    'Noviembre': 39643.59,
    'Diciembre': 39727.96
}

MESES_2025 = list(UF_2025.keys())
TOPE_UF = 131.8

def limpiar_rut(rut):
    """Limpia el formato del RUT"""
    if pd.isna(rut):
        return ''
    rut_str = str(rut).strip()
    rut_str = re.sub(r'[.\-]', '', rut_str)
    return rut_str

def calcular_seguro_cesantia_empleador(row, valor_uf):
    """Calcula el seguro de cesantía empleador según tipo de contrato"""
    imponible = row.get('Total Imponible (H)', 0)
    seguro_trabajador = row.get('Seguro de Cesantía', 0)
    
    if pd.isna(imponible) or imponible == 0:
        return 0
    
    tope_pesos = TOPE_UF * valor_uf
    base_calculo = min(imponible, tope_pesos)
    
    # Determinar tipo de contrato
    if pd.notna(seguro_trabajador) and seguro_trabajador > 0:
        tasa = 0.024  # Indefinido
    else:
        tasa = 0.030  # Plazo Fijo/Obra
    
    return round(base_calculo * tasa, 0)

def procesar_libro_remuneraciones(df_libro, rut_empresa):
    """Procesa el libro de remuneraciones extrayendo las columnas necesarias"""
    df_procesado = pd.DataFrame()
    
    # RUT Empresa
    df_procesado['Empresa'] = rut_empresa
    
    # RUT Trabajador
    if 'RUT Trabajador' in df_libro.columns:
        df_procesado['RUT Trabajador'] = df_libro['RUT Trabajador'].apply(limpiar_rut)
    
    # Nombres completos
    apellido_paterno = df_libro.get('Apellido Paterno', '')
    apellido_materno = df_libro.get('Apellido Materno', '')
    nombres = df_libro.get('Nombres', '')
    
    df_procesado['Nombres Completos'] = (
        apellido_paterno.fillna('').astype(str) + ' ' +
        apellido_materno.fillna('').astype(str) + ' ' +
        nombres.fillna('').astype(str)
    ).str.strip()
    
    # Extraer columnas de días (solo las que empiezan con 'N°')
    columnas_dias = [col for col in df_libro.columns if col.startswith('N°')]
    for col in columnas_dias:
        df_procesado[col] = df_libro[col]
    
    # Total Imponible
    if 'Imponible' in df_libro.columns:
        df_procesado['Total Imponible (H)'] = df_libro['Imponible']
    
    # Seguro de Cesantía (para cálculo posterior)
    if 'Seguro de Cesantía' in df_libro.columns:
        df_procesado['Seguro de Cesantía'] = df_libro['Seguro de Cesantía']
    else:
        df_procesado['Seguro de Cesantía'] = 0
    
    return df_procesado

def procesar_haberes_descuentos(df_informe):
    """Procesa el informe de haberes y descuentos"""
    df_procesado = pd.DataFrame()
    
    # Identificar columnas de haberes y descuentos
    for col in df_informe.columns:
        col_limpia = col.replace('(Monto)', '').strip()
        
        # Clasificar por tipo
        if 'haber' in col.lower() or col_limpia in ['Sueldo Base', 'Gratificación', 'Bono', 'Horas Extras']:
            nueva_col = f"{col_limpia} (H)"
        elif 'descuento' in col.lower() or col_limpia in ['AFP', 'Salud', 'Impuesto']:
            nueva_col = f"{col_limpia} (D)"
        else:
            # Intentar clasificar por contenido
            if df_informe[col].sum() > 0:
                nueva_col = f"{col_limpia} (H)"
            else:
                nueva_col = f"{col_limpia} (D)"
        
        df_procesado[nueva_col] = df_informe[col]
    
    return df_procesado

def consolidar_datos(datos_empresas, mes, año):
    """Consolida los datos de todas las empresas"""
    if not datos_empresas:
        return None
    
    df_consolidado = pd.concat(datos_empresas, ignore_index=True)
    
    # Agregar Mes y Año
    df_consolidado.insert(1, 'Mes', mes)
    df_consolidado.insert(2, 'Año', año)
    
    # Ordenar columnas
    columnas_fijas = ['Empresa', 'Mes', 'Año', 'RUT Trabajador', 'Nombres Completos']
    columnas_dias = [col for col in df_consolidado.columns if col.startswith('N°')]
    columnas_haberes = [col for col in df_consolidado.columns if col.endswith('(H)')]
    columnas_descuentos = [col for col in df_consolidado.columns if col.endswith('(D)')]
    
    # Remover columna temporal de Seguro de Cesantía si existe
    if 'Seguro de Cesantía' in df_consolidado.columns:
        df_consolidado = df_consolidado.drop(columns=['Seguro de Cesantía'])
    
    columnas_otras = [col for col in df_consolidado.columns 
                     if col not in columnas_fijas + columnas_dias + columnas_haberes + columnas_descuentos]
    
    orden_final = columnas_fijas + columnas_dias + columnas_haberes + columnas_descuentos + columnas_otras
    orden_final = [col for col in orden_final if col in df_consolidado.columns]
    
    df_consolidado = df_consolidado[orden_final]
    
    return df_consolidado

def generar_excel(df, nombre_archivo):
    """Genera un archivo Excel con formato profesional"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Consolidado"
    
    # Encabezados
    for col_idx, col_name in enumerate(df.columns, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.value = col_name
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
        cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Datos
    for row_idx, row_data in enumerate(df.values, start=2):
        for col_idx, value in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.value = value
            cell.alignment = Alignment(vertical='center')
    
    # Ajustar anchos
    for col_idx, col_name in enumerate(df.columns, start=1):
        max_length = max(len(str(col_name)), 12)
        ws.column_dimensions[chr(64 + col_idx) if col_idx <= 26 else f"A{chr(64 + col_idx - 26)}"].width = max_length + 2
    
    # Guardar en BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    
    return output

# Configuración de la página
st.set_page_config(
    page_title="Consolidador de Remuneraciones 2025",
    page_icon="📊",
    layout="wide"
)

# Inicializar session_state
if 'datos_empresas' not in st.session_state:
    st.session_state.datos_empresas = []
if 'empresas_agregadas' not in st.session_state:
    st.session_state.empresas_agregadas = []

# Título
st.title("📊 Consolidador de Remuneraciones 2025")
st.markdown("---")

# Barra lateral
with st.sidebar:
    st.header("⚙️ Configuración")
    
    rut_empresa = st.text_input("RUT Empresa para el Excel", placeholder="12345678-9")
    mes_proceso = st.selectbox("Mes de Proceso", MESES_2025)
    
    st.markdown("---")
    st.markdown("### Empresas agregadas:")
    if st.session_state.empresas_agregadas:
        for empresa in st.session_state.empresas_agregadas:
            st.success(f"✓ {empresa}")
    else:
        st.info("Ninguna empresa agregada aún")
    
    if st.button("🔄 Limpiar Todo", type="secondary"):
        st.session_state.datos_empresas = []
        st.session_state.empresas_agregadas = []
        st.rerun()

# Contenido principal
col1, col2 = st.columns(2)

with col1:
    st.subheader("📄 Libro de Remuneraciones")
    archivo_libro = st.file_uploader(
        "Cargar Libro de Remuneraciones",
        type=['xlsx', 'xls'],
        key='libro'
    )

with col2:
    st.subheader("📋 Informe de Haberes y Descuentos")
    archivo_informe = st.file_uploader(
        "Cargar Informe de Haberes y Descuentos",
        type=['xlsx', 'xls'],
        key='informe'
    )

st.markdown("---")

# Botón para agregar empresa
col_btn1, col_btn2, col_btn3 = st.columns([1, 1, 2])

with col_btn1:
    if st.button("➕ Agregar Empresa", type="primary", disabled=not (archivo_libro and archivo_informe and rut_empresa)):
        try:
            # Leer archivos
            df_libro = pd.read_excel(archivo_libro)
            df_informe = pd.read_excel(archivo_informe)
            
            # Procesar datos
            df_libro_procesado = procesar_libro_remuneraciones(df_libro, rut_empresa)
            df_informe_procesado = procesar_haberes_descuentos(df_informe)
            
            # Combinar por RUT
            df_combinado = df_libro_procesado.copy()
            
            # Agregar columnas del informe
            for col in df_informe_procesado.columns:
                df_combinado[col] = df_informe_procesado[col]
            
            # Calcular Seguro Cesantía Empleador
            valor_uf = UF_2025[mes_proceso]
            df_combinado['Seguro Cesantía Empleador (D)'] = df_combinado.apply(
                lambda row: calcular_seguro_cesantia_empleador(row, valor_uf),
                axis=1
            )
            
            # Agregar a session_state
            st.session_state.datos_empresas.append(df_combinado)
            
            # Determinar nombre de empresa
            if rut_empresa.startswith('96'):
                nombre_empresa = "INGEMARS"
            elif rut_empresa.startswith('90'):
                nombre_empresa = "ENAP"
            else:
                nombre_empresa = f"Empresa {rut_empresa[:8]}"
            
            st.session_state.empresas_agregadas.append(nombre_empresa)
            
            st.success(f"✅ {nombre_empresa} agregada exitosamente!")
            st.rerun()
            
        except Exception as e:
            st.error(f"❌ Error al procesar los archivos: {str(e)}")

with col_btn2:
    if st.button("📥 Generar Excel", type="primary", disabled=len(st.session_state.datos_empresas) == 0):
        try:
            # Consolidar datos
            df_final = consolidar_datos(st.session_state.datos_empresas, mes_proceso, 2025)
            
            if df_final is not None:
                # Generar Excel
                nombre_archivo = f"Consolidado_{mes_proceso}_2025.xlsx"
                excel_buffer = generar_excel(df_final, nombre_archivo)
                
                st.download_button(
                    label="⬇️ Descargar Excel Consolidado",
                    data=excel_buffer,
                    file_name=nombre_archivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                st.success(f"✅ Excel generado: {nombre_archivo}")
                
                # Mostrar preview
                with st.expander("👁️ Vista Previa de Datos"):
                    st.dataframe(df_final.head(10), use_container_width=True)
                    st.info(f"Total de registros: {len(df_final)}")
            
        except Exception as e:
            st.error(f"❌ Error al generar el Excel: {str(e)}")

# Información adicional
with st.expander("ℹ️ Información del Sistema"):
    st.markdown("""
    ### Cómo usar:
    1. **Configure** el RUT de la empresa y el mes de proceso en la barra lateral
    2. **Cargue** el Libro de Remuneraciones y el Informe de Haberes y Descuentos
    3. **Presione** "Agregar Empresa" para procesar y guardar los datos
    4. **Repita** los pasos 1-3 para agregar más empresas al mismo mes
    5. **Genere** el Excel consolidado con todas las empresas
    
    ### Cálculo Seguro Cesantía Empleador:
    - **Contrato Indefinido**: 2.4% (si trabajador tiene Seguro Cesantía > 0)
    - **Contrato Plazo Fijo/Obra**: 3.0% (si trabajador tiene Seguro Cesantía = 0)
    - **Tope**: 131.8 UF del mes correspondiente
    
    ### Valores UF 2025:
    """)
    
    col_uf1, col_uf2, col_uf3 = st.columns(3)
    with col_uf1:
        for mes in MESES_2025[:4]:
            st.write(f"**{mes}**: ${UF_2025[mes]:,.2f}")
    with col_uf2:
        for mes in MESES_2025[4:8]:
            st.write(f"**{mes}**: ${UF_2025[mes]:,.2f}")
    with col_uf3:
        for mes in MESES_2025[8:]:
            st.write(f"**{mes}**: ${UF_2025[mes]:,.2f}")
