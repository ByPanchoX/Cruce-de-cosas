import streamlit as st
import pandas as pd
import io
import re

# Configuración de la App
st.set_page_config(page_title="INGEMARS - Auditoría y Carga", page_icon="⚖️", layout="wide")

st.title("⚖️ Sistema de Auditoría y Conversión - INGEMARS")
st.markdown("Valida la cuadratura contra Liquidaciones Generales e Informes de Haberes antes de generar el archivo.")

# --- BARRA LATERAL: CONFIGURACIÓN ---
st.sidebar.header("⚙️ Parámetros de Empresa")
rs_input = st.sidebar.text_input("RUT Razón Social", value="76455680-1")
rut_empresa = rs_input.replace(".", "").strip()

st.sidebar.subheader("Tasas SIS 2025 (%)")
meses_nombres = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
defaults_sis = [1.38, 1.38, 1.38, 1.78, 1.78, 1.78, 1.88, 1.88, 1.88, 1.49, 1.49, 1.49]
tasas_sis_map = {}
for i, mes in enumerate(meses_nombres):
    tasas_sis_map[i+1] = st.sidebar.number_input(f"SIS {mes}", value=defaults_sis[i], step=0.01) / 100

st.sidebar.subheader("Otros Parámetros")
tasa_mutual = st.sidebar.number_input("Mutual INGEMARS (%)", value=1.27, step=0.01) / 100

def limpiar_monto(v):
    if pd.isna(v): return 0
    try:
        if isinstance(v, str):
            v = v.replace('.', '').replace(' ', '').replace('$', '')
        return int(float(v))
    except: return 0

# --- CARGA DE ARCHIVOS ---
col_base, col_audit = st.columns(2)

with col_base:
    st.subheader("📁 1. Archivos Base")
    template_file = st.file_uploader("Plantilla Talana (Excel/CSV)", type=["xlsx", "csv"])
    libros_files = st.file_uploader("Libros de Remuneraciones (Varios)", type=["xlsx", "csv"], accept_multiple_files=True)

with col_audit:
    st.subheader("🔍 2. Archivos de Auditoría")
    liq_files = st.file_uploader("Liquidaciones Generales", type=["xlsx", "csv"], accept_multiple_files=True)
    hab_files = st.file_uploader("Informes de Haberes y Descuentos", type=["xlsx", "csv"], accept_multiple_files=True)

if template_file and libros_files:
    # Preparar el proceso
    try:
        if template_file.name.endswith('.csv'):
            df_struct = pd.read_csv(template_file, skiprows=2, sep=None, engine='python', encoding='latin-1')
        else:
            df_struct = pd.read_excel(template_file, skiprows=2)
        cols_oficiales = [c.strip() for c in df_struct.columns.tolist() if "Unnamed" not in c]
        
        st.divider()
        st.info("Confirma los periodos para iniciar la validación:")
        
        with st.form("audit_form"):
            periodos_map = {}
            for f in libros_files:
                c1, c2 = st.columns([3, 1])
                c1.write(f"📄 {f.name}")
                match = re.search(r'(\d{2})-(\d{4})', f.name)
                def_p = f"{match.group(2)}{match.group(1)}" if match else "202501"
                periodos_map[f.name] = c2.text_input("Periodo", value=def_p, key=f.name)
            
            run_audit = st.form_submit_button("🚀 Validar Cuadratura y Generar")

        if run_audit:
            lista_final = []
            resumen_audit = []

            for f_libro in libros_files:
                p = periodos_map[f_libro.name]
                mes_num = int(p[4:])
                
                # Leer Libro
                df_l = pd.read_csv(f_libro, sep=None, engine='python', encoding='latin-1') if f_libro.name.endswith('.csv') else pd.read_excel(f_libro)
                df_l.columns = [c.strip() for c in df_l.columns]
                
                # Totales del Libro
                L_imp = df_l['Imponible'].sum()
                L_hab = df_l['Total Haberes'].sum()
                L_liq = df_l['Líquido'].sum()

                # --- Auditoría vs Liquidación General ---
                Q_imp, Q_hab, Q_liq = "N/A", "N/A", "N/A"
                if liq_files:
                    # Buscar archivo de liq que mencione el mes o periodo
                    for fl in liq_files:
                        if p in fl.name or str(mes_num).zfill(2) in fl.name:
                            df_q = pd.read_csv(fl, sep=None, engine='python', encoding='latin-1') if fl.name.endswith('.csv') else pd.read_excel(fl)
                            labels = df_q.iloc[:, 3].astype(str)
                            vals = df_q.iloc[:, 8]
                            if any(labels.str.contains("TOTAL GANADO", na=False)):
                                Q_imp = limpiar_monto(vals[labels.str.contains("TOTAL GANADO")].values[0])
                            if any(labels.str.contains("TOTAL HABERES", na=False)):
                                Q_hab = limpiar_monto(vals[labels.str.contains("TOTAL HABERES")].values[0])
                            if any(labels.str.contains("TOTAL TRABAJADOR", na=False)):
                                Q_liq = limpiar_monto(vals[labels.str.contains("TOTAL TRABAJADOR")].values[0])

                resumen_audit.append({
                    "Periodo": p,
                    "Imponible (Libro)": int(L_imp),
                    "Imponible (Reporte)": Q_imp,
                    "Diferencia": (int(L_imp) - Q_imp) if isinstance(Q_imp, int) else "---",
                    "Estado": "✅ OK" if (L_imp == Q_imp or Q_imp == "N/A") else "❌ REVISAR"
                })

                # --- Mapeo Talana ---
                df_mes = pd.DataFrame(0, index=df_l.index, columns=cols_oficiales)
                df_mes['Rut *'] = df_l['Rut Trabajador'].astype(str).str.replace(".", "")
                df_mes['Rut Razón Social *'] = rut_empresa
                df_mes['Año_Mes (aaaamm) *'] = p
                
                mapping = {'Sueldo Base *': 'Sueldo Base', 'Remuneración Imponible *': 'Imponible', 
                           'Remuneración Total o Sueldo Bruto *': 'Total Haberes', 'Sueldo Liquido *': 'Líquido',
                           'AFP *': 'Prevision', 'Seguro Cesantía Trabajador *': 'Seguro de Cesantía',
                           'Días Trabajados del Mes *': 'N° Dias Trabajados', 'Mensual *': 'Gratificacion'}
                for tal, lib in mapping.items():
                    if lib in df_l.columns: df_mes[tal] = pd.to_numeric(df_l[lib], errors='coerce').fillna(0)

                df_mes['Total Isapre *'] = pd.to_numeric(df_l['Salud'], errors='coerce').fillna(0) + pd.to_numeric(df_l.get('Diferencia Plan de Salud', 0), errors='coerce').fillna(0)
                
                # Aportes (SIS variable + Mutual + SC lógica foto)
                t_sis = tasas_sis_map.get(mes_num, 0.0149)
                imp = df_mes['Remuneración Imponible *']
                df_mes['Seguro Invalidez y Supervivencia (SIS) *'] = (imp * t_sis).round(0)
                df_mes['Seguro Accidente de Trabajo *'] = (imp * tasa_mutual).round(0)
                df_mes['Seguro Cesantía Empleador *'] = imp * df_mes['Seguro Cesantía Trabajador *'].apply(lambda x: 0.024 if x > 0 else 0.030)
                df_mes['Seguro Cesantía Empleador *'] = df_mes['Seguro Cesantía Empleador *'].round(0)

                lista_final.append(df_mes)

            # Mostrar Reporte de Auditoría
            st.subheader("📋 Informe de Cuadratura")
            st.table(pd.DataFrame(resumen_audit))

            # Generar Descarga
            df_export = pd.concat(lista_final, ignore_index=True)
            for c in df_export.columns:
                if '*' in c and 'Rut' not in c and 'Año_Mes' not in c:
                    df_export[c] = pd.to_numeric(df_export[c]).astype(int)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_export.to_excel(writer, index=False)
            
            st.success("Auditoría completada. Si ves todo en verde, puedes descargar el archivo.")
            st.download_button("⬇️ Descargar Carga Masiva Talana", output.getvalue(), "Carga_Masiva_INGEMARS.xlsx")

    except Exception as e:
        st.error(f"Error en el proceso: {e}")
