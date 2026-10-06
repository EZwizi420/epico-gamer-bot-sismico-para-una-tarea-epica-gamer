## Hotfix 3.0.2 — 2026-10-05
- Corregido `NameError: timezone is not defined` en **Actividad**.
- Se conserva la normalización UTC añadida en 3.0.1.

## Hotfix 3.0.1 — 2026-10-05
- Corregido `10062 Unknown interaction` al abrir **Investigación**: la interacción se confirma antes de generar el panel.
- Corregido **Actividad** cuando CSN mezcla timestamps con y sin zona horaria; ahora se normalizan a UTC antes de ordenar.

# Changelog — Edición Definitiva

## QoL v3 / Definitiva
- Mapa de predicciones vs. eventos CSN.
- Evolución del ranking mediante snapshots desde esta versión.
- Modo investigación integrado al panel.
- Accesos visuales a análisis, comparación y actividad.

## QoL v2
- Filtros del explorador por estado.
- Perfil de IA y comparador IA vs. IA.
- Análisis de error temporal, magnitud y distancia.
- Estadísticas descriptivas y cumplimiento por criterio.
- Centro de actividad.

## QoL v1
- Explorador interactivo de predicciones por IA/grupo.
- `/ver_real` interactivo y navegador CSN.
- Ranking reorganizado por aciertos, casi acertados y no acertados.
- Separación visual de casi acertadas y pendientes.

## Base histórica
- Sincronización CSN, importación Excel, evaluación y alertas.
- Caducidad/evaluación de predicciones.
- Score de proximidad, historial, integridad, backups y dashboard.
- Exportación Excel e informe PDF.

## 3.0.3 Hotfix — Gemini ± hour windows
- Added support for Gemini date/time values such as `03/10/2026 08:30 ± 4 h`.
- The central time is now expanded symmetrically (example: 08:30 ± 4 h = 04:30–12:30).
- Windows that cross midnight correctly update both the start/end date and time.
- Also accepts `+/-` and `+-`, plus `h`, `hr(s)` and `hora(s)` variants.
