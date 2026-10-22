## Hotfix 3.0.9 — 2026-10-22
- El botón **Ranking** confirma la interacción de Discord antes de ejecutar el cálculo; el cálculo se realiza en un hilo para evitar `10062 Unknown interaction`.
- Añadido `/reconstruir_correlaciones`: sincroniza las predicciones del Excel, importa catálogos históricos CSN para ventanas de hasta 31 días y reconstruye el historial sin borrar registros ni enviar alertas masivas.
- El mapa Leaflet utiliza una capa oscura CARTO y distingue visualmente aciertos, casi aciertos, pendientes y no acertados.
- El dashboard clasifica correctamente `CASI ACERTADA`, la incluye en estadísticas y ranking, y `/health` informa el estado de la base de datos y sus conteos.
- `EXCEL_PATH` relativo ahora busca también el archivo Excel junto a `bot.py` si Render inicia desde otro directorio.

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

## v3.0.7 - Evaluación CSN explicable
- `/ver_real` y el explorador ya no ocultan el mejor evento cuando una predicción termina NO ACERTADA.
- Prioriza eventos dentro de la ventana temporal y muestra qué criterio falló: tiempo, magnitud o distancia.
- Confirma que los eventos de `/importar_historico` se guardan en `observed_events` y son leídos por el mismo `events()` usado por la evaluación.
