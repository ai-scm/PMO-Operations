# Prompt: Análisis de Carga de Project Managers – Blend360 Colombia
# Fecha: Abril 2026
# Generado por: Claude Code (Anthropic) + PMO Operations

## Descripción del Proceso
Este prompt y script generan automáticamente los siguientes entregables:
1. Informe Word de carga por PM con datos reales y gráficas
2. Consolidado Excel con evolución y detalle de carga marzo-abril
3. Propuesta Word de manejo eficiente de asignación de proyectos

## Archivos de Entrada (1. Input)
- maestro.proyectos.xlsx: Maestro de proyectos activos con PM asignado
- Horas marzo abril.xlsx: Reporte OpenAir con horas registradas (42,542 registros)
- Plantilla word Blend.docx: Plantilla corporativa para documentos Word

## Lógica del Análisis
- Capacidad estándar: 44 horas/semana por PM
- Marzo 2026: 4.4 semanas efectivas → 193.6 horas/PM
- PMs identificados: Oscar Barragan, Juan Bernal, David Cortes, Miguel Garcia,
  Kelly Carbonell, Daniel Sebastian Vargas, Diana Castro, Diana Rojas, Indira Duarte
- Categorías de horas: Facturable | Interno/Administrativo | Vacaciones/Ausencia
- Proyectos colombianos: prefijo "CO -" en nombre del proyecto + clientes colombianos

## Archivos Generados (3. Output)
- Consolidado_Carga_PM_Blend360.xlsx: 4 hojas (Resumen, Detalle, Evolución, Proyectos)
- Informe_Carga_PMs_Blend360.docx: Informe detallado con gráficas por PM
- Propuesta_Gestion_Asignacion_PM_Blend360.docx: Propuesta de modelo de asignación
- charts/: Gráficas PNG generadas con matplotlib

## Para Ejecutar
1. Asegurarse de tener Python 3.x con: openpyxl, matplotlib, python-docx, pandas
2. Abrir terminal en la carpeta PMO-Operations
3. Ejecutar: py gen_outputs.py
   O desde la carpeta 2. Process: py ../gen_outputs.py

## Actualización Mensual
Para actualizar con nuevos datos:
1. Reemplazar "Horas marzo abril.xlsx" con el nuevo export de OpenAir
2. Actualizar maestro.proyectos.xlsx si hay cambios en asignaciones
3. Ejecutar el script nuevamente
4. Los archivos de Output se sobrescriben automáticamente
