import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io

# Configuración de la página
st.set_page_config(
    page_title="Consolidador de Remuneraciones Chile 2025",
    page_icon="📊",
    layout="wide"
)

# Valores UF último día de cada mes 2025
VALORES_UF_2025 = {
    'Enero': 38284.86,
    'Febrero': 38647.94,
    'Marzo': 38800.00,
    'Abril': 38950.00,
    'Mayo': 39100.00,
    'Junio': 39250.00,
    'Julio': 39400.00,
    'Agosto': 39550.00,
    'Septiembre': 39700.00,
    'Octubre': 39850.00,
    'Noviembre': 40000.00,
    'Diciembre': 40150.00
}

# Topes AFC en UF para 2025
TOPES_AFC_UF = {
    'Enero': 131.8,
    'Febrero': 131.9,
    'Marzo': 131.9,
    'Abril': 131.9,
    'Mayo': 131.9,
    'Junio': 131.9,
    'Julio': 131.9,
    'Agosto': 131.9,
    'Septiembre': 131.9,
    'Octubre': 131.9,
    'Noviembre': 131.9,
    'Diciembre': 131.9
}

def inicializar_session_state():
    """Inicializa las variables de sesión"""
    if 'datos_acumulados' not in st.session_state:
        st.session_state.datos_acumulados = []
    if 'empresas_cargadas' not in st.session_state:
        st.session_state.empresas_cargadas = []

def reiniciar_memoria():
    """Limpia todos los datos acumulados"""
    st.session_state.datos_acumulados = []
    st.session_state.empresas_cargadas = []
    st.success("✅ Memoria reiniciada correctamente")

def limpiar_rut(rut):
    """Limpia y formatea el RUT"""
    if pd.isna(rut):
        return ""
    rut_str = str(rut).strip()
    # Eliminar puntos y guiones
    rut_limpio = rut_str.replace('.', '').replace('-', '')
    # Separar dígito verificador
    if len(rut_limpio) > 1:
        return f"{rut_limpio[:-1]}-{rut_limpio[-1]}"
    return rut_str

def calcular_tope_afc_pesos(mes):
    """Calcula el tope AFC en pesos para el mes dado"""
    tope_uf = TOPES_AFC_UF[mes]
    valor_uf = VALORES_UF_2025[mes]
    return tope_uf * valor_uf

def detectar_tipo_contrato(descuento_cesantia):
    """
    Detecta el tipo de contrato basado en el descuento de cesantía
    Retorna la tasa de empleador correspondiente
    """
    if pd.isna(descuento_cesantia) or descuento_cesantia == 0:
        # Plazo Fijo/Obra
        return 3.0
    else:
        # Indefinido
        return 2.4

def calcular_seguro_cesantia_empleador(row, mes):
    """
    Calcula el Seguro de Cesantía Empleador según normativa chilena 2025
    """
    # Obtener el total imponible
    total_imponible = row.get('Total Imponible (H)', 0)
    if pd.isna(total_imponible):
        total_imponible = 0
    
    # Obtener el descuento de cesantía del trabajador
    descuento_cesantia = row.get('Seguro de Cesantía (D)', 0)
    if pd.isna(descuento_cesantia):
        descuento_cesantia = 0
    
    # Detectar tipo de contrato y obtener tasa
    tasa_empleador = detectar_tipo_contrato(descuento_cesantia)
    
    # Calcular tope AFC en pesos
    tope_afc_pesos = calcular_tope_afc_pesos(mes)
    
    # Aplicar tope
    base_calculo = min(total_imponible, tope_afc_pesos)
    
    # Calcular seguro cesantía empleador
    seguro_cesantia = base_calculo * (tasa_empleador / 100)
    
    return round(seguro_cesantia, 0)

def procesar_libro_remuneraciones(df):
    """Procesa el Libro de Remuneraciones extrayendo datos básicos y días"""
    df_procesado = pd.DataFrame()
    
    # Extraer datos personales - buscar con variaciones
    columnas_personales = {
        'Rut Trabajador': 'RUT Trabajador',
        'Apellido Paterno': 'Apellido Paterno',
        'Apellido Materno': 'Apellido Materno',
        'Nombres': 'Nombres'
    }
    
    for col_original, col_nueva in columnas_personales.items():
        if col_original in df.columns:
            df_procesado[col_nueva] = df[col_original]
        else:
            # Buscar variaciones del nombre de columna (case insensitive)
            cols_encontradas = [c for c in df.columns if col_original.lower() in c.lower()]
            if cols_encontradas:
                df_procesado[col_nueva] = df[cols_encontradas[0]]
            else:
                df_procesado[col_nueva] = ""
    
    # Limpiar RUTs
    df_procesado['RUT Trabajador'] = df_procesado['RUT Trabajador'].apply(limpiar_rut)
    
    # Extraer columnas de días (solo las que empiezan ESTRICTAMENTE con 'N°')
    columnas_dias = [col for col in df.columns if col.strip().startswith('N°')]
    for col in columnas_dias:
        df_procesado[col] = df[col]
    
    # Extraer Total Imponible - buscar varias opciones
    columna_imponible = None
    for opcion in ['Imponible', 'Total Imponible', 'Sueldo Imponible', 'Base Imponible']:
        if opcion in df.columns:
            columna_imponible = opcion
            break
    
    if columna_imponible:
        df_procesado['Total Imponible (H)'] = pd.to_numeric(df[columna_imponible], errors='coerce').fillna(0)
    else:
        df_procesado['Total Imponible (H)'] = 0
    
    # Extraer Seguro de Cesantía (descuento del trabajador)
    if 'Seguro de Cesantía' in df.columns:
        df_procesado['Seguro de Cesantía (D)'] = pd.to_numeric(df['Seguro de Cesantía'], errors='coerce').fillna(0)
    else:
        cols_cesantia = [c for c in df.columns if 'cesant' in c.lower() and 'seguro' in c.lower()]
        if cols_cesantia:
            df_procesado['Seguro de Cesantía (D)'] = pd.to_numeric(df[cols_cesantia[0]], errors='coerce').fillna(0)
        else:
            df_procesado['Seguro de Cesantía (D)'] = 0
    
    return df_procesado

def procesar_informe_haberes_descuentos(df_original):
    """
    Procesa el Informe de Haberes y Descuentos con formato especial
    El archivo tiene estructura: secciones de conceptos seguidas de trabajadores
    """
    # Leer el archivo sin encabezado para procesarlo manualmente
    df = df_original.copy()
    
    # Diccionario para almacenar datos por RUT
    datos_por_rut = {}
    
    # Variables de control
    concepto_actual = None
    es_haber = True  # Empieza con haberes
    
    # Procesar fila por fila
    for idx, row in df.iterrows():
        row_values = row.values
        
        # Detectar cambio a DESCUENTOS
        primera_col = str(row_values[0]).strip().upper() if pd.notna(row_values[0]) else ""
        if 'DESCUENTO' in primera_col:
            es_haber = False
            continue
        
        # Detectar nuevo concepto (fila con nombre en primera columna)
        if pd.notna(row_values[0]) and row_values[0] != '' and str(row_values[0]).strip() != '':
            # Es un concepto si no es un RUT
            texto = str(row_values[0]).strip()
            if not any(char.isdigit() for char in texto[:3]):  # No empieza con números
                concepto_actual = texto
                continue
        
        # Detectar fila de trabajador (tiene RUT en columna 2)
        rut_col = row_values[2] if len(row_values) > 2 else None
        monto_col = row_values[10] if len(row_values) > 10 else None
        
        if pd.notna(rut_col) and pd.notna(monto_col):
            rut_str = str(rut_col).strip()
            # Verificar que parece un RUT
            if '-' in rut_str or (len(rut_str) >= 8 and any(c.isdigit() for c in rut_str)):
                rut_limpio = limpiar_rut(rut_str)
                
                # Inicializar diccionario para este RUT si no existe
                if rut_limpio not in datos_por_rut:
                    datos_por_rut[rut_limpio] = {}
                
                # Agregar el concepto y monto
                if concepto_actual:
                    sufijo = "(H)" if es_haber else "(D)"
                    nombre_columna = f"{concepto_actual} {sufijo}"
                    
                    try:
                        monto = float(monto_col)
                        # Si ya existe el concepto, sumar (por si hay duplicados)
                        if nombre_columna in datos_por_rut[rut_limpio]:
                            datos_por_rut[rut_limpio][nombre_columna] += monto
                        else:
                            datos_por_rut[rut_limpio][nombre_columna] = monto
                    except:
                        pass
    
    # Convertir el diccionario a DataFrame
    if not datos_por_rut:
        # Retornar DataFrame vacío con columna RUT
        return pd.DataFrame(columns=['RUT Trabajador'])
    
    df_resultado = pd.DataFrame.from_dict(datos_por_rut, orient='index')
    df_resultado.reset_index(inplace=True)
    df_resultado.rename(columns={'index': 'RUT Trabajador'}, inplace=True)
    
    # Rellenar NaN con 0
    df_resultado = df_resultado.fillna(0)
    
    return df_resultado

def consolidar_datos(df_libro, df_informe, empresa, mes, rut_empresa):
    """Consolida los datos del libro y el informe"""
    
    # Merge por RUT
    df_consolidado = pd.merge(
        df_libro,
        df_informe,
        on='RUT Trabajador',
        how='left'
    )
    
    # Agregar columnas de metadata al inicio
    df_consolidado.insert(0, 'Empresa', empresa)
    df_consolidado.insert(1, 'Mes', mes)
    df_consolidado.insert(2, 'Año', 2025)
    df_consolidado.insert(3, 'RUT Empresa', rut_empresa)
    
    # Calcular Seguro de Cesantía Empleador
    df_consolidado['Seguro de Cesantía Empleador (D)'] = df_consolidado.apply(
        lambda row: calcular_seguro_cesantia_empleador(row, mes),
        axis=1
    )
    
    # Ordenar columnas: Metadata | Datos Personales | Días | Haberes | Descuentos
    columnas_ordenadas = ['Empresa', 'Mes', 'Año', 'RUT Empresa', 'RUT Trabajador', 
                          'Apellido Paterno', 'Apellido Materno', 'Nombres']
    
    # Agregar columnas de días (N°)
    columnas_dias = [col for col in df_consolidado.columns if col.startswith('N°')]
    columnas_ordenadas.extend(sorted(columnas_dias))
    
    # Agregar haberes (H)
    columnas_haberes = [col for col in df_consolidado.columns if col.endswith('(H)')]
    columnas_ordenadas.extend(sorted(columnas_haberes))
    
    # Agregar descuentos (D)
    columnas_descuentos = [col for col in df_consolidado.columns if col.endswith('(D)')]
    columnas_ordenadas.extend(sorted(columnas_descuentos))
    
    # Reordenar
    df_consolidado = df_consolidado[columnas_ordenadas]
    
    return df_consolidado

def generar_excel(datos_acumulados):
    """Genera el archivo Excel consolidado"""
    if not datos_acumulados:
        return None
    
    # Concatenar todos los dataframes
    df_final = pd.concat(datos_acumulados, ignore_index=True)
    
    # Crear archivo Excel en memoria
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_final.to_excel(writer, index=False, sheet_name='Consolidado')
        
        # Ajustar ancho de columnas
        worksheet = writer.sheets['Consolidado']
        for idx, col in enumerate(df_final.columns):
            max_length = max(
                df_final[col].astype(str).apply(len).max(),
                len(col)
            )
            worksheet.column_dimensions[chr(65 + idx)].width = min(max_length + 2, 50)
    
    output.seek(0)
    return output

# Interfaz principal
def main():
    st.title("📊 Consolidador de Remuneraciones Chile 2025")
    st.markdown("---")
    
    # Inicializar session state
    inicializar_session_state()
    
    # Barra lateral
    with st.sidebar:
        st.header("⚙️ Configuración")
        
        # Selector de mes
        meses = list(VALORES_UF_2025.keys())
        mes_seleccionado = st.selectbox(
            "Mes de Proceso",
            meses,
            index=0
        )
        
        # RUT de empresa
        rut_empresa = st.text_input(
            "RUT Único de Empresa",
            value="76.455.680-1"
        )
        
        st.markdown("---")
        
        # Botón de reinicio
        if st.button("🔄 Reiniciar Memoria", type="secondary", use_container_width=True):
            reiniciar_memoria()
        
        # Información del mes
        st.markdown("---")
        st.info(f"""
        **Información del Mes:**
        - UF: ${VALORES_UF_2025[mes_seleccionado]:,.2f}
        - Tope AFC: {TOPES_AFC_UF[mes_seleccionado]} UF
        - Tope AFC: ${calcular_tope_afc_pesos(mes_seleccionado):,.0f}
        """)
        
        # Mostrar empresas cargadas
        if st.session_state.empresas_cargadas:
            st.markdown("---")
            st.success(f"✅ Empresas cargadas: {len(st.session_state.empresas_cargadas)}")
            for empresa in st.session_state.empresas_cargadas:
                st.write(f"- {empresa}")
    
    # Área principal
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📄 Libro de Remuneraciones")
        libro_file = st.file_uploader(
            "Cargar archivo Excel",
            type=['xlsx', 'xls'],
            key='libro'
        )
    
    with col2:
        st.subheader("📋 Informe de Haberes y Descuentos")
        informe_file = st.file_uploader(
            "Cargar archivo Excel",
            type=['xlsx', 'xls'],
            key='informe'
        )
    
    # Nombre de empresa
    empresa_nombre = st.text_input(
        "Nombre de la Empresa",
        value="INGEMARS",
        help="Ejemplo: INGEMARS o ENAP"
    )
    
    # Botón de procesamiento
    if st.button("➕ Procesar y Agregar", type="primary", use_container_width=True):
        if libro_file and informe_file and empresa_nombre:
            with st.spinner("Procesando archivos..."):
                try:
                    # Leer archivos
                    df_libro = pd.read_excel(libro_file)
                    # Leer informe SIN encabezado (header=None) para procesar manualmente
                    df_informe = pd.read_excel(informe_file, header=None)
                    
                    # Procesar
                    df_libro_proc = procesar_libro_remuneraciones(df_libro)
                    df_informe_proc = procesar_informe_haberes_descuentos(df_informe)
                    
                    # Consolidar
                    df_consolidado = consolidar_datos(
                        df_libro_proc,
                        df_informe_proc,
                        empresa_nombre,
                        mes_seleccionado,
                        rut_empresa
                    )
                    
                    # Agregar a acumulados
                    st.session_state.datos_acumulados.append(df_consolidado)
                    st.session_state.empresas_cargadas.append(empresa_nombre)
                    
                    st.success(f"✅ {empresa_nombre} procesada correctamente ({len(df_consolidado)} registros)")
                    
                    # Mostrar preview
                    with st.expander("👁️ Vista previa de los datos procesados"):
                        st.dataframe(df_consolidado.head(10))
                    
                except Exception as e:
                    st.error(f"❌ Error al procesar archivos: {str(e)}")
        else:
            st.warning("⚠️ Por favor, carga ambos archivos y especifica el nombre de la empresa")
    
    # Sección de generación de Excel
    if st.session_state.datos_acumulados:
        st.markdown("---")
        st.subheader("📥 Generar Excel Consolidado")
        
        total_registros = sum(len(df) for df in st.session_state.datos_acumulados)
        st.info(f"Total de registros acumulados: {total_registros}")
        
        if st.button("📊 Generar Excel", type="primary", use_container_width=True):
            with st.spinner("Generando archivo Excel..."):
                try:
                    excel_file = generar_excel(st.session_state.datos_acumulados)
                    
                    if excel_file:
                        nombre_archivo = f"Consolidado_{mes_seleccionado}_2025.xlsx"
                        
                        st.download_button(
                            label="⬇️ Descargar Excel Consolidado",
                            data=excel_file,
                            file_name=nombre_archivo,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                        
                        st.success(f"✅ Archivo generado: {nombre_archivo}")
                    
                except Exception as e:
                    st.error(f"❌ Error al generar Excel: {str(e)}")

if __name__ == "__main__":
    main()
