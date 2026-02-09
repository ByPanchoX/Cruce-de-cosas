# 🚀 Guía de Inicio Rápido

## Instalación en 3 pasos

### 1. Instalar dependencias
```bash
pip install -r requirements.txt
```

### 2. Ejecutar la aplicación
```bash
streamlit run consolidador_remuneraciones.py
```

### 3. Abrir en el navegador
La aplicación se abrirá automáticamente en: `http://localhost:8501`

---

## Prueba con archivos de ejemplo

Usa los archivos de ejemplo incluidos para probar la aplicación:

### Ejemplo 1: INGEMARS
1. **RUT Empresa**: `96123456-7`
2. **Mes**: Enero
3. **Libro de Remuneraciones**: `ejemplo_libro_ingemars.xlsx`
4. **Informe de Haberes y Descuentos**: `ejemplo_informe_ingemars.xlsx`
5. Presiona **"Agregar Empresa"**

### Ejemplo 2: ENAP
1. **RUT Empresa**: `90987654-3`
2. **Mes**: Enero (mismo mes que INGEMARS)
3. **Libro de Remuneraciones**: `ejemplo_libro_enap.xlsx`
4. **Informe de Haberes y Descuentos**: `ejemplo_informe_enap.xlsx`
5. Presiona **"Agregar Empresa"**

### Generar consolidado
Presiona **"Generar Excel"** para descargar `Consolidado_Enero_2025.xlsx`

---

## Resultado esperado

El archivo consolidado contendrá:
- **8 trabajadores** en total (5 de INGEMARS + 3 de ENAP)
- Columnas organizadas: Empresa | Mes | Año | RUT | Nombres | Días | Haberes | Descuentos
- **Seguro Cesantía Empleador** calculado automáticamente:
  - 2.4% para contratos indefinidos
  - 3.0% para contratos plazo fijo/obra

---

## Estructura esperada de tus archivos

### Libro de Remuneraciones
Debe incluir estas columnas (los nombres deben ser exactos):
- `RUT Trabajador`
- `Apellido Paterno`
- `Apellido Materno`
- `Nombres`
- `N° Dias Trabajados` (o cualquier columna que empiece con 'N°')
- `Imponible`
- `Seguro de Cesantía`

### Informe de Haberes y Descuentos
Puede incluir cualquier concepto, por ejemplo:
- `Sueldo Base (Monto)`
- `Gratificación (Monto)`
- `Bono Asistencia (Monto)`
- `AFP (Monto)`
- `Salud (Monto)`
- `Impuesto Único (Monto)`

**Nota**: El sistema eliminará automáticamente el sufijo `(Monto)` y agregará `(H)` para haberes o `(D)` para descuentos.

---

## Flujo de trabajo típico

```
┌─────────────────────────────────────────┐
│  1. Configurar RUT y Mes                │
│     (Barra lateral)                     │
└──────────────┬──────────────────────────┘
               │
               ▼
┌─────────────────────────────────────────┐
│  2. Cargar archivos Excel               │
│     • Libro de Remuneraciones           │
│     • Informe de Haberes y Descuentos   │
└──────────────┬──────────────────────────┘
               │
               ▼
┌─────────────────────────────────────────┐
│  3. Agregar Empresa                     │
│     (Los datos se guardan en memoria)   │
└──────────────┬──────────────────────────┘
               │
               ▼
┌─────────────────────────────────────────┐
│  4. ¿Agregar otra empresa?              │
│     Sí → Volver al paso 1               │
│     No → Continuar                      │
└──────────────┬──────────────────────────┘
               │
               ▼
┌─────────────────────────────────────────┐
│  5. Generar Excel Consolidado           │
│     Descargar archivo resultante        │
└─────────────────────────────────────────┘
```

---

## Consejos útiles

### ✅ Mejores prácticas
- Verifica que los RUT estén correctamente formateados
- Asegúrate de que las columnas requeridas existan
- Revisa la vista previa antes de descargar

### ⚠️ Errores comunes
- **"Error al procesar archivos"**: Verifica nombres de columnas
- **Cálculos incorrectos**: Confirma que 'Imponible' y 'Seguro de Cesantía' existan
- **Empresas no aparecen**: Usa el botón "Agregar Empresa" después de cargar archivos

### 🔄 Reiniciar
Si necesitas empezar de nuevo:
1. Usa el botón **"Limpiar Todo"** en la barra lateral
2. O simplemente recarga la página en el navegador

---

## Soporte técnico

### Verificar instalación
```bash
python --version  # Debe ser 3.8 o superior
pip list | grep streamlit  # Verificar que streamlit esté instalado
```

### Reinstalar dependencias
```bash
pip install --upgrade -r requirements.txt
```

### Ver logs de error
Si la aplicación falla, los errores aparecerán:
- En la terminal donde ejecutaste `streamlit run`
- En la interfaz web de Streamlit

---

**¡Listo para comenzar! 🎉**

Ejecuta `streamlit run consolidador_remuneraciones.py` y comienza a consolidar tus remuneraciones.
