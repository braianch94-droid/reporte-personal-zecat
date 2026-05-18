# Reporte de Personal - Articulos Promocionales SA

Sistema automatico de reportes de asistencia y horas extras para el personal de deposito y despacho.

## Contenido

| Archivo | Descripcion |
|---|---|
| `GenerarHTML.ps1` | Genera el reporte interactivo HTML principal |
| `GenerarReporte.ps1` | Genera el reporte de asistencia en Excel |
| `GenerarReporteBreaks.ps1` | Genera el reporte de descansos en Excel |
| `Actualizar.bat` | Ejecuta todo y abre el reporte automaticamente |

## Uso

1. Colocar el archivo `Asistencia.xlsx` (exportado desde n8n) en esta carpeta
2. Hacer doble clic en **`Actualizar.bat`**
3. El reporte `Reporte_Personal.html` se genera y se abre en el navegador

## Reporte HTML - Pestanas

- **Asistencia** - Presentes, ausentes, feriados y dias de descanso por persona
- **Descansos y Almuerzo** - Duracion de breaks, desvios y promedios
- **Novedades e Inconsistencias** - Atrasos, salidas anticipadas, ausencias reales
- **Horas Extras** - Horas por encima de la jornada (50% Lun-Vie, 50%/100% Sabado)

## Reglas de negocio

- Jornada laboral: **Lunes a Jueves 9 horas**, **Viernes 8 horas**
- Viernes: salida 1 hora antes del turno habitual
- Horas extras semana: al **50%**
- Horas extras sabado hasta 13:00: al **50%** / despues de 13:00: al **100%**
- Umbral minimo para hora extra: **30 minutos**

## Requisitos

- Windows con PowerShell 5.1
- Microsoft Excel instalado (para leer el .xlsx via COM)

## Estructura fuente (Asistencia.xlsx)

El archivo debe tener la siguiente estructura por fila:

| Col | Campo | Descripcion |
|---|---|---|
| A | Apellidos | Apellido del empleado |
| B | Nombre | Nombre del empleado |
| D | Grupo | Sector (ej: DEPOSITO-Z) |
| E | Fecha | Formato "Lun 04-05-2026" |
| G | Turno | Horario o "Descanso" |
| H | E1 | Entrada |
| L | S1 | Salida al break |
| P | E2 | Retorno del break |
| T | S2 | Salida final |
| AA | HT | Horas trabajadas |
