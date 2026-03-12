# ============================================================
#  COMBINADOR DE REMUNERACIONES — Google Colab  v4 FINAL
#  Funciona con plantilla VACÍA (solo encabezados).
#  Crea las filas desde cero usando los RUTs del consolidado.
# ============================================================
# USO:
#  1. Pega este script en una celda de Colab
#  2. Ejecuta:  main()
#  3. Sube la Plantilla Maestra (vacía con solo encabezados)
#  4. Sube el Consolidado de Remuneraciones
#  5. El archivo combinado se descarga automáticamente
# ============================================================

import re, io, numpy as np
import pandas as pd
from openpyxl import load_workbook
from pathlib import Path

try:
    from google.colab import files
    IN_COLAB = True
except ImportError:
    IN_COLAB = False

OUTPUT_FILE = "plantilla_combinada.xlsx"

# ── Helpers ──────────────────────────────────────────────────

def limpiar_rut(v):
    if pd.isna(v): return ""
    return re.sub(r"[\.\-\s]", "", str(v)).upper()

def to_num(v):
    if pd.isna(v) or str(v).strip() in ("", "nan", "None"): return 0.0
    try: return float(str(v).replace(",", ".").replace(" ", ""))
    except: return 0.0

def detectar_header_consolidado(path_o_bytes, es_csv=False):
    """Detecta fila con 'Rut' en consolidado."""
    raw = (pd.read_csv(path_o_bytes, header=None, dtype=str) if es_csv
           else pd.read_excel(path_o_bytes, header=None, dtype=str, sheet_name=0))
    for idx, row in raw.iterrows():
        if any(isinstance(c, str) and "rut" in c.lower() for c in row.values):
            df = raw.iloc[idx:].reset_index(drop=True)
            df.columns = df.iloc[0]
            return df.iloc[1:].reset_index(drop=True), idx
    raise ValueError("No se encontró columna 'Rut' en consolidado.")

# ── Subida de archivos ────────────────────────────────────────

def subir():
    if IN_COLAB:
        print("📂 Sube la PLANTILLA MAESTRA (.xlsx):")
        up = files.upload()
        nombre_p = list(up.keys())[0]
        bytes_p  = io.BytesIO(up[nombre_p])

        print("\n📂 Sube el CONSOLIDADO DE REMUNERACIONES (.xlsx o .csv):")
        up = files.upload()
        nombre_c = list(up.keys())[0]
        bytes_c  = io.BytesIO(up[nombre_c])
    else:
        nombre_p = "plantilla_maestra.xlsx"
        nombre_c = "consolidado.xlsx"
        bytes_p  = nombre_p
        bytes_c  = nombre_c
    return bytes_p, nombre_p, bytes_c, nombre_c

# ── MAPEO DE COLUMNAS ────────────────────────────────────────
#
# Estructura: (col_PLANTILLA, [cols_CONSOLIDADO], operacion)
#   "first"  → primer valor no vacío de la lista
#   "sum"    → suma todos los valores
#   "yyyymm" → año + mes → formato AAAAMM
#

MAPEO_MANUAL = [
    ("Rut Razón Social *",                       ["RUT EMPRESA"],                                        "first"),
    ("Rut *",                                    ["Rut Trabajador"],                                     "first"),
    ("Año_Mes (aaaamm) *",                       ["AÑO", "MES"],                                         "yyyymm"),
    
    # ── Leyes Sociales ────────────────────────────────────────
    ("Total Isapre *",                            ["Cotización Institución de Salud (D)",
                                                   "Cotización Voluntaria a Isapre (D)"],                "sum"),
    ("Seguro Cesantía Trabajador *",              ["Descuento Seguro Cesantia (D)"],                     "first"),
    ("AFP *",                                     ["Descuento Cotización AFP (D)"],                      "first"),
    ("Impuestos *",                               ["Impuesto Unico (D)"],                                "first"),
    ("Seguro Invalidez y Supervivencia (SIS) *",  ["Seguro de Invalidez y Sobrevivencia Empleador (P)"], "first"),
    ("Seguro Accidente de Trabajo *",             ["Aporte Accidentes de Trabajo Mutual (P)"],           "first"),
    ("Seguro Cesantía Empleador *",               ["Seguro de Cesantía Empleador (P)"],                  "first"),
    ("Expectativa de Vida",                       ["Expectativa de Vida Seguro Social (P)"],             "first"),
    ("Capitalización Individual AFP",             ["Adicional de Capitalización Individual AFP (P)"],    "first"),
    
    # ── Días ──────────────────────────────────────────────────
    ("Días Trabajados del Mes *",                 ["N° Dias Trabajados"],                                "first"),
    ("Días de Ausentismo *",                      ["N° Dias Ausentes"],                                  "first"),
    ("Días de licencia médica *",                 ["N° Dias Licencia"],                                  "first"),
    
    # ── Remuneraciones ────────────────────────────────────────
    ("Sueldo Base *",                             ["Sueldo Base (H)"],                                   "first"),
    ("Mensual *",                                 ["Gratificación Legal (H)"],                           "first"),
    ("Remuneración Imponible *",                  ["Total Imponible (H)"],                               "first"),
    
    # ── Haberes ───────────────────────────────────────────────
    ("Movilización",                              ["Movilización (H)"],                                  "first"),
    ("Colación",                                  ["Colación (H)"],                                      "first"),
    ("Asignación familiar y Maternal",            ["Asignacion Familiar (H)"],                           "first"),
    ("Aguinaldo (h)",                             ["AGUINALDO (H)", "Aguinaldo (H)"],                    "sum"),
    
    # ── Descuentos ────────────────────────────────────────────
    ("Descuento Anticipo",                        ["Anticipos (D)"],                                     "first"),
    ("Descuento Crédito Personal CCAF",           ["Créditos Personales CCAF (D)"],                      "first"),
    ("anticipo D",                                ["ANTICIPO AGUINALDO (D)", "Anticipo Aguinaldo (D)"],  "sum"),
]

# Mapeo automático: columnas con mismo nombre
MAPEO_AUTO = [
    ("ASIGNACION DE VIATICO (H)",            "ASIGNACION DE VIATICO (H)"),
    ("ASIGNACION VIATICO MEU (H)",           "ASIGNACION VIATICO MEU (H)"),
    ("Asignación Capacitación (H)",          "Asignación Capacitación (H)"),
    ("Asignación Navidad (H)",               "Asignación Navidad (H)"),
    ("Asignación de Cargo (H)",              "Asignación de Cargo (H)"),
    ("Asignación de Telefono (H)",           "Asignación de Telefono (H)"),
    ("BONO ADMINISTRACION (H)",              "BONO ADMINISTRACION (H)"),
    ("BONO PMI (H)",                         "BONO PMI (H)"),
    ("BONO RECONOCIMIENTO (H)",              "BONO RECONOCIMIENTO (H)"),
    ("Bono Calidad (H)",                     "Bono Calidad (H)"),
    ("Bono Extraordinario (H)",              "Bono Extraordinario (H)"),
    ("Bono Producción (H)",                  "Bono Producción (H)"),
    ("Bono Reconocimiento (H)",              "Bono Reconocimiento (H)"),
    ("Bono Vacaciones (H)",                  "Bono Vacaciones (H)"),
    ("Bono años de servicio (H)",            "Bono años de servicio (H)"),
    ("Bono de Faena (H)",                    "Bono de Faena (H)"),
    ("Bono de Productividad (H)",            "Bono de Productividad (H)"),
    ("Bono por Asistencia (H)",              "Bono por Asistencia (H)"),
    ("Compensacion dia feriado (H)",         "Compensacion dia feriado (H)"),
    ("DIAS PARO (H)",                        "DIAS PARO (H)"),
    ("Gastos Reembolsables (H)",             "Gastos Reembolsables (H)"),
    ("HORAS EXTRAS DEL MES (H)",             "HORAS EXTRAS DEL MES (H)"),
    ("ICTP (H)",                             "ICTP (H)"),
    ("Incentivo de Desempeño (H)",           "Incentivo de Desempeño (H)"),
    ("Internet (H)",                         "Internet (H)"),
    ("Llamado de Emergencia (H)",            "Llamado de Emergencia (H)"),
    ("MOVILIZACION ENAP (H)",                "MOVILIZACION ENAP (H)"),
    ("Reconocimiento Mentoria (H)",          "Reconocimiento Mentoria (H)"),
    ("Sobretiempo (H)",                      "Sobretiempo (H)"),
    ("Sobretiempo Festivo (H)",              "Sobretiempo Festivo (H)"),
    ("Viatico Alojamiento (H)",              "Viatico Alojamiento (H)"),
    ("Visita Adicional (H)",                 "Visita Adicional (H)"),
    ("Otros Descuentos",                     "Otros Descuentos (D)"),
    ("Anticipo Asignación Capacitación (D)", "Anticipo Asignación Capacitación (D)"),
    ("Anticipo Asignación Navidad (D)",      "Anticipo Asignación Navidad (D)"),
    ("Anticipo Bono Calidad (D)",            "Anticipo Bono Calidad (D)"),
    ("Anticipo Bono Producción (D)",         "Anticipo Bono Producción (D)"),
    ("Anticipo Bono Vacaciones (D)",         "Anticipo Bono Vacaciones (D)"),
    ("Anticipo bono años de servicio (D)",   "Anticipo bono años de servicio (D)"),
    ("Anticipos (D)",                        "Anticipos (D)"),
    ("Crédito Fonasa (D)",                   "Crédito Fonasa (D)"),
    ("Descuento cuota Sindicato (D)",        "Descuento cuota Sindicato (D)"),
    ("Descuentos por Leasing (D)",           "Descuentos por Leasing (D)"),
    ("Gastos Particulares (D)",              "Gastos Particulares (D)"),
    ("Gastos Reembolsables (D)",             "Gastos Reembolsables (D)"),
    ("Prestamo Empresa (D)",                 "Prestamo Empresa (D)"),
    ("Retencion Pension de Alimentos (D)",   "Retencion Pension de Alimentos (D)"),
    ("Seguro Vida Camara (D)",               "Seguro Vida Camara (D)"),
    ("Seguro de vida CCAF (D)",              "Seguro de vida CCAF (D)"),
]

# Diccionario meses
MESES = {
    "enero": "01", "febrero": "02", "marzo": "03", "abril": "04",
    "mayo": "05", "junio": "06", "julio": "07", "agosto": "08",
    "septiembre": "09", "octubre": "10", "noviembre": "11", "diciembre": "12"
}

# ── Motor de combinación ──────────────────────────────────────

def combinar(bytes_p, nombre_p, bytes_c, nombre_c):
    # 1. Leer estructura de plantilla vacía
    print("\n🔍 Leyendo plantilla maestra…")
    wb_plantilla = load_workbook(bytes_p)
    ws_plantilla = wb_plantilla.active
    
    columnas_plantilla = []
    fila_header = 3  # Fila 3 = encabezados
    for i in range(1, ws_plantilla.max_column + 1):
        h = ws_plantilla.cell(fila_header, i).value
        if h:
            columnas_plantilla.append(h)
    
    print(f"   {len(columnas_plantilla)} columnas en plantilla")
    
    # 2. Leer consolidado
    print("🔍 Leyendo consolidado…")
    es_csv = nombre_c.lower().endswith(".csv")
    df_c, _ = detectar_header_consolidado(bytes_c, es_csv)
    print(f"   {len(df_c)} personas en consolidado")
    
    # 3. Crear DataFrame vacío con estructura plantilla
    df_resultado = pd.DataFrame(columns=columnas_plantilla)
    
    # 4. Procesar cada persona del consolidado
    print("\n⚙️  Procesando filas…")
    for idx, row_c in df_c.iterrows():
        nueva_fila = {}
        
        # Mapeo manual
        for col_dest, cols_src, op in MAPEO_MANUAL:
            if op == "yyyymm":
                año = str(row_c.get("AÑO", "")).strip()
                mes = str(row_c.get("MES", "")).strip().lower()
                mes_num = MESES.get(mes, "")
                nueva_fila[col_dest] = f"{año}{mes_num}" if año and mes_num else None
            elif op == "sum":
                total = sum(to_num(row_c.get(c)) for c in cols_src if c in df_c.columns)
                nueva_fila[col_dest] = str(int(round(total))) if total != 0 else None
            else:  # "first"
                val = None
                for c in cols_src:
                    if c in df_c.columns and not pd.isna(row_c[c]) and str(row_c[c]).strip() not in ("", "nan"):
                        val = row_c[c]
                        break
                nueva_fila[col_dest] = val
        
        # Mapeo automático
        for col_dest, col_src in MAPEO_AUTO:
            if col_src in df_c.columns:
                val = row_c[col_src]
                nueva_fila[col_dest] = val if not pd.isna(val) and str(val).strip() not in ("", "nan") else None
        
        df_resultado = pd.concat([df_resultado, pd.DataFrame([nueva_fila])], ignore_index=True)
    
    print(f"✅ {len(df_resultado)} filas creadas")
    return df_resultado, columnas_plantilla, fila_header

# ── Guardar preservando instrucciones ─────────────────────────

def guardar(df_resultado, bytes_p, columnas_plantilla, fila_header):
    print("\n💾 Generando archivo…")
    if isinstance(bytes_p, io.BytesIO):
        bytes_p.seek(0)
    
    wb = load_workbook(bytes_p)
    ws = wb.active
    
    # Borrar filas antiguas (si hubiera)
    if ws.max_row > fila_header:
        ws.delete_rows(fila_header + 1, ws.max_row - fila_header)
    
    # Escribir datos
    for ri, (_, row_data) in enumerate(df_resultado.iterrows(), 1):
        for ci, col in enumerate(columnas_plantilla, 1):
            val = row_data.get(col)
            cell_val = None if (pd.isna(val) or str(val).strip() in ("nan", "None", "")) else val
            ws.cell(row=fila_header + ri, column=ci, value=cell_val)
    
    wb.save(OUTPUT_FILE)
    print(f"✅ {OUTPUT_FILE}")

# ── Main ──────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("  COMBINADOR DE REMUNERACIONES  v4 FINAL")
    print("=" * 60)
    
    bytes_p, nombre_p, bytes_c, nombre_c = subir()
    
    df_resultado, columnas_plantilla, fila_header = combinar(bytes_p, nombre_p, bytes_c, nombre_c)
    
    if isinstance(bytes_p, io.BytesIO):
        bytes_p.seek(0)
    
    guardar(df_resultado, bytes_p, columnas_plantilla, fila_header)
    
    print("\n📊 Vista previa (primeras 3 personas):")
    cols_ver = ["Rut *", "Año_Mes (aaaamm) *", "Total Isapre *", "AFP *",
                "Sueldo Base *", "Bono Calidad (H)", "Sobretiempo (H)"]
    cols_ok = [c for c in cols_ver if c in df_resultado.columns]
    print(df_resultado[cols_ok].head(3).to_string(index=False))
    
    if IN_COLAB:
        print("\n⬇  Descargando…")
        files.download(OUTPUT_FILE)
    else:
        print(f"\nArchivo: {Path(OUTPUT_FILE).resolve()}")
    
    print("\n🎉 ¡Proceso completado!")


if __name__ == "__main__":
    main()
