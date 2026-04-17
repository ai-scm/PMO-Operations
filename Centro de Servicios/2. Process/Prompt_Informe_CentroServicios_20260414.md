# Prompt: Informe de Seguimiento Centro de Servicios

**Fecha:** 14 de abril de 2026  
**Fuentes:**
- `1. Input/[P0089] Daily Centro de Servicios.docx` — Transcripción reunión 10 abril 2026
- `1. Input/2026-04-14T15_07_59.058Z BLEND Team - Engineering Colombia - P 0321 BLD Soporte - Soporte Interno.xlsx` — Tablero de tareas ClickUp
- `1. Input/Plantilla word Blend.docx` — Plantilla corporativa

**Objetivo:**
Cruzar la información del Daily (reunión del 10/04/2026) con el tablero de tareas ClickUp para generar un informe profesional que incluya:
1. Introducción — Contexto de la reunión y alcance del reporte
2. Desarrollo — Análisis detallado tarea por tarea: modificaciones de fechas, responsables, prioridades y planes de acción
3. Compromisos — Acuerdos explícitos tomados en la reunión
4. Pendientes — Tareas con riesgo de incumplimiento o sin fecha definida
5. Comentarios — Observaciones generales de gestión
6. Cronograma de trabajo — Tabla con tareas, responsables, fechas y prioridades de abril

**Formato:**
- Plantilla: Plantilla word Blend.docx (conservar encabezado con logo)
- Títulos: Montserrat 16 pt, negrita
- Subtítulos: Montserrat 14 pt, negrita
- Párrafos: Montserrat 10 pt, sin negrita, color negro, justificado

**Cronograma de trabajo:**
- Visual, estilo PM, organizado por semanas (descendente: más reciente primero)
- 3 semanas: 28–30 abr, 21–25 abr, 14–17 abr
- Columnas: #, Fecha, ID, Tarea / Actividad, Responsable, Prioridad, Iniciativa
- Color de fila según prioridad (fondo claro):
  - URGENTE: `FFCCCC` | ALTA: `FFE5CC` | NORMAL: `CCE0FF` | BAJA: `D9F0CC`
- **Mapa de calor en columna Prioridad** (fondo sólido + texto blanco):
  - URGENTE: `C00000` | ALTA: `BF5900` | NORMAL: `1F3864` | BAJA: `375623`
- Encabezado de semana: azul oscuro `1F3864`, texto blanco
- Encabezado de columnas: azul `2E4057`, texto blanco
- Celda de fecha resaltada en `D0D8E8` cuando cambia
- Resumen de carga por responsable al final del cronograma (tabla 3 columnas)

**Nomenclatura de responsables (nombres completos, sin abreviaturas):**
| Nombre completo       | Abreviaturas a evitar               |
|-----------------------|-------------------------------------|
| Johan Espino          | Espino                              |
| Johan Velandia        | Velandia                            |
| Luis H. Bernal M.     | Bernal, L. H. Bernal M., Luis Hernando Bernal |
| Jose Florez           | Florez                              |

**Scripts de generación:**
- `generar_informe_cs.py` — Informe de seguimiento completo
- `generar_resumen_cs.py` — Resumen ejecutivo

**Salidas:**
- `3. Output/Informe_Seguimiento_CentroServicios_20260414.docx`
- `3. Output/Resumen_Ejecutivo_CentroServicios_20260414.docx`
