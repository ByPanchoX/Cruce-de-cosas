import streamlit as st
import pandas as pd
import io
import re

# Configuración de la App
st.set_page_config(page_title="INGEMARS - Carga Talana Pro", page_icon="📈", layout="wide")

st.title("📈 Convertidor Histórico INGEMARS")
st.markdown("Configura las tasas y la razón social para procesar tus libros de remuneraciones.")

# --- BARRA LATERAL: CONFIGURACIÓN EDITABLE ---
st.sidebar.header("⚙️ Configuración de Empresa")

# Campo para Razón Social (editable)
rs_input = st.sidebar.text_input("RUT Razón Social (Sin puntos y con guion)", value="76455680-1")
# Limpieza de seguridad para la Razón Social
rut_empresa = rs_input.replace(".", "").strip()

st.sidebar.subheader("Tasa Mutual")
tasa_mutual = st.sidebar.number_input("Mutual INGEMARS (%)", value=1.27, step=0.01) / 100

st.sidebar.subheader("Tasas SIS Mensuales (%)")
tasas_sis_input = {}
meses_nombres = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
]
# Valores por defecto proporcionados por Francisco
defaults_sis = [1.38, 1.38, 1.38, 1.78, 1.78, 1.78, 1.88, 1.88, 1.88, 1.49, 1.49, 1.49]

for i, mes in enumerate(meses_nombres):
    tasas_sis_input[i+1] = st.sidebar.number_input(f"SIS {mes}", value=defaults_sis[i], step=0.01) / 100

def limpieza_extrema(val):
    if isinstance(val, str):
        val = re.sub(r'[^\w\s\d.,\-_@\(\)/:;?¿!¡+*#=]', '', val)
        val = "".join(c for c in val if c.isprintable())
        return val.strip()
    return val

# --- INTERFAZ DE CARGA ---
col1, col2 = st.columns(2)
with col1:
    template_file = st.file_uploader("1. Sube la Plantilla de Talana (Excel o CSV)", type=["xlsx", "csv"])
with col2:
    libros_files = st.file_uploader("2. Sube los Libros de Remuneraciones", type=["xlsx", "csv"], accept_multiple_files=True)

if template_file and libros_files:
    try:
        if template_file.name.endswith('.csv'):
            df_struct = pd.read_csv(template_file, skiprows=2, sep=None, engine='python', encoding='latin-1', on_bad_lines='skip')
        else:
            df_struct = pd.read_excel(template_file, skiprows=2)
            
        cols_oficiales = [c.strip() for c in df_struct.columns.tolist() if "Unnamed" not in c]
        lista_meses = []

        with st.form("procesar_ingemars"):
            dict_p = {}
            for f in libros_files:
                c1, c2 = st.columns([3, 1])
                c1.write(f"📄 {f.name}")
                default_p = "202501"
                match = re.search(r'(\d{2})-(\d{4})', f.name)
                if match: default_p = f"{match.group(2)}{match.group(1)}"
                dict_p[f.name] = c2.text_input("Periodo (AAAAMM)", value=default_p, key=f.name)
            
            submit = st.form_submit_button("🚀 Procesar y Consolidar")

        if submit:
            for f in libros_files:
                periodo = dict_p[f.name]
                mes_num = int(str(periodo)[4:])
                tasa_sis_actual = tasas_sis_input.get(mes_num, 0.0149)
                
                df_l = pd.read_csv(f, sep=None, engine='python', encoding='latin-1') if f.name.endswith('.csv') else pd.read_excel(f)
                df_l.columns = [c.strip() for c in df_l.columns]
                
                df_mes = pd.DataFrame(0, index=df_l.index, columns=cols_oficiales)
                
                # Mapeo base con RUT de empresa editable
                df_mes['Rut *'] = df_l['Rut Trabajador'].astype(str).str.replace(".", "")
                df_mes['Rut Razón Social *'] = rut_empresa
                df_mes['Año_Mes (aaaamm) *'] = periodo
                
                map_cols = {
                    'Sueldo Base *': 'Sueldo Base',
                    'Remuneración Imponible *': 'Imponible',
                    'Remuneración Total o Sueldo Bruto *': 'Total Haberes',
                    'Sueldo Liquido *': 'Líquido',
                    'AFP *': 'Prevision',
                    'Seguro Cesantía Trabajador *': 'Seguro de Cesantía',
                    'Días Trabajados del Mes *': 'N° Dias Trabajados',
                    'Mensual *': 'Gratificacion'
                }

                for tal, lib in map_cols.items():
                    if lib in df_l.columns:
                        df_mes[tal] = pd.to_numeric(df_l[lib], errors='coerce').fillna(0)

                df_mes['Total Isapre *'] = pd.to_numeric(df_l['Salud'], errors='coerce').fillna(0) + pd.to_numeric(df_l.get('Diferencia Plan de Salud', 0), errors='coerce').fillna(0)
                
                # Aportes Patronales
                imp = df_mes['Remuneración Imponible *']
                df_mes['Seguro Invalidez y Supervivencia (SIS) *'] = (imp * tasa_sis_actual).round(0)
                df_mes['Seguro Accidente de Trabajo *'] = (imp * tasa_mutual).round(0)
                
                # Lógica Seguro Cesantía (2.4% vs 3.0%)
                df_mes['Seguro Cesantía Empleador *'] = df_mes['Seguro Cesantía Trabajador *'].apply(lambda x: 0.024 if x > 0 else 0.030)
                df_mes['Seguro Cesantía Empleador *'] = (imp * df_mes['Seguro Cesantía Empleador *']).round(0)

                lista_meses.append(df_mes)

            df_final = pd.concat(lista_meses, ignore_index=True)
            df_final = df_final[cols_oficiales].applymap(limpieza_extrema)
            
            for c in df_final.columns:
                if '*' in c and 'Rut' not in c and 'Año_Mes' not in c:
                    df_final[c] = pd.to_numeric(df_final[c]).astype(int)

            st.success(f"✅ ¡Procesado para {rut_empresa}!")
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False)
            
            st.download_button("⬇️ Descargar Carga Histórica", output.getvalue(), "Carga_Masiva_INGEMARS.xlsx")

    except Exception as e:
        st.error(f"Error: {e}")
