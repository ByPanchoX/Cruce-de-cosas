# Consolidador de Remuneraciones 2025

Aplicación web desarrollada en Streamlit para consolidar archivos de remuneraciones del año 2025, permitiendo procesar múltiples empresas en un mismo mes.

## 🚀 Instalación

### Requisitos previos
- Python 3.8 o superior
- pip (gestor de paquetes de Python)

### Pasos de instalación

1. **Instalar dependencias**
```bash
pip install -r requirements.txt
```

2. **Ejecutar la aplicación**
```bash
streamlit run consolidador_remuneraciones.py
```

La aplicación se abrirá automáticamente en tu navegador en `http://localhost:8501`

## 📋 Características

### Funcionalidades principales

1. **Carga de múltiples empresas**: Permite agregar datos de INGEMARS y ENAP (u otras empresas) para un mismo mes sin perder información.

2. **Procesamiento automático**: 
   - Extrae RUT, nombres completos y días trabajados del Libro de Remuneraciones
   - Filtra automáticamente columnas de días (solo las que comienzan con 'N°')
   - Clasifica haberes y descuentos con sufijos (H) y (D)

3. **Cálculo de Seguro Cesantía Empleador**:
   - **Contrato Indefinido**: 2.4% (detectado automáticamente si el trabajador tiene monto > 0 en 'Seguro de Cesantía')
   - **Contrato Plazo Fijo/Obra**: 3.0% (si el monto es 0 o no existe)
   - Aplica tope de 131.8 UF según valores mensuales de 2025

4. **Generación de Excel consolidado**: Archivo con formato profesional y estructura ordenada.

## 📖 Guía de uso

### Paso 1: Configuración inicial
En la barra lateral:
- Ingresa el **RUT Empresa** (formato: 12345678-9)
- Selecciona el **Mes de Proceso** (Enero a Diciembre 2025)

### Paso 2: Cargar archivos
- **Libro de Remuneraciones**: Debe contener columnas como 'RUT Trabajador', 'Apellido Paterno', 'Apellido Materno', 'Nombres', 'Imponible', y columnas de días que comiencen con 'N°'
- **Informe de Haberes y Descuentos**: Debe contener los conceptos desglosados

### Paso 3: Agregar empresa
Presiona el botón **"Agregar Empresa"**. Los datos se guardarán en memoria.

### Paso 4: Agregar más empresas (opcional)
Para agregar una segunda empresa al mismo mes:
1. Cambia el RUT Empresa
2. Carga los nuevos archivos
3. Presiona nuevamente "Agregar Empresa"

### Paso 5: Generar consolidado
Presiona **"Generar Excel"** para crear el archivo consolidado con todas las empresas del mes.

## 📊 Estructura del archivo de salida

El Excel generado contiene las siguientes columnas en orden:

1. **Empresa**: RUT de la empresa
2. **Mes**: Mes de proceso
3. **Año**: 2025
4. **RUT Trabajador**: RUT del trabajador (limpio, sin puntos ni guiones)
5. **Nombres Completos**: Apellido Paterno + Apellido Materno + Nombres
6. **Columnas de Días**: Solo las que comienzan con 'N°' (ej: N° Días Trabajados)
7. **Haberes (H)**: Todos los conceptos de haberes con sufijo (H)
8. **Descuentos (D)**: Todos los conceptos de descuentos con sufijo (D), incluyendo 'Seguro Cesantía Empleador (D)'

## 💡 Valores UF 2025

El sistema incluye los valores de UF para todo 2025:

| Mes | Valor UF |
|-----|----------|
| Enero | $38,284.86 |
| Febrero | $38,647.94 |
| Marzo | $38,800.00 |
| Abril | $39,050.00 |
| Mayo | $39,200.00 |
| Junio | $39,267.07 |
| Julio | $39,179.01 |
| Agosto | $39,383.07 |
| Septiembre | $39,500.00 |
| Octubre | $39,597.67 |
| Noviembre | $39,643.59 |
| Diciembre | $39,727.96 |

## 🔧 Lógica de procesamiento

### Limpieza de datos
- **Columnas de días**: Solo se incluyen las que comienzan exactamente con 'N°'
- **RUT**: Se eliminan puntos y guiones para unificación
- **Conceptos**: Se eliminan sufijos como '(Monto)' antes de clasificar

### Clasificación de conceptos
El sistema clasifica automáticamente cada concepto como Haber (H) o Descuento (D) basándose en:
1. Palabras clave en el nombre de la columna
2. Valores típicos de haberes/descuentos
3. Suma total (positivo = Haber, negativo = Descuento)

### Cálculo Seguro Cesantía Empleador
```
Base de cálculo = MIN(Total Imponible, Tope en pesos)
Tope en pesos = 131.8 UF × Valor UF del mes

Si Seguro Cesantía trabajador > 0:
    Seguro Empleador = Base × 2.4%  (Indefinido)
Sino:
    Seguro Empleador = Base × 3.0%  (Plazo Fijo/Obra)
```

## ⚠️ Consideraciones importantes

1. **Formato de archivos**: Los archivos Excel deben tener la estructura estándar de libros de remuneraciones chilenos
2. **Columnas requeridas en Libro de Remuneraciones**:
   - RUT Trabajador
   - Apellido Paterno, Apellido Materno, Nombres
   - Imponible
   - Seguro de Cesantía (para determinar tipo de contrato)
   - Columnas de días con formato 'N° ...'

3. **Orden de carga**: Puedes cargar las empresas en cualquier orden
4. **Reinicio**: Usa el botón "Limpiar Todo" en la barra lateral para empezar de nuevo

## 📝 Ejemplo de uso

```
1. RUT Empresa: 96123456-7 (INGEMARS)
   Mes: Enero
   → Cargar archivos de INGEMARS
   → Agregar Empresa

2. RUT Empresa: 90987654-3 (ENAP)
   Mes: Enero
   → Cargar archivos de ENAP
   → Agregar Empresa

3. Generar Excel → Descarga "Consolidado_Enero_2025.xlsx"
```

## 🐛 Solución de problemas

### La aplicación no inicia
- Verifica que todas las dependencias estén instaladas
- Asegúrate de tener Python 3.8 o superior

### Error al cargar archivos
- Verifica que los archivos sean .xlsx o .xls
- Revisa que las columnas requeridas existan en los archivos

### El cálculo del Seguro Cesantía es incorrecto
- Verifica que la columna 'Seguro de Cesantía' exista en el Libro de Remuneraciones
- Confirma que los valores de 'Imponible' sean correctos

## 📧 Soporte

Para reportar problemas o sugerencias, por favor documenta:
1. Versión de Python utilizada
2. Mensaje de error completo (si aplica)
3. Descripción del problema
4. Pasos para reproducir el error

---

**Versión**: 1.0  
**Última actualización**: Febrero 2025
