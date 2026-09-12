# Bot Sísmico Discord v1.0.1

Incluye:
- Alertas automáticas CSN 24/7.
- 🚨 coincidencia 3/3 con @everyone.
- 🟡 casi coincidencia cuando cumple exactamente 2/3 filtros (sin @everyone).
- `/ranking` por IA.
- `/historial` de correlaciones automáticas.
- congelado/integridad de predicciones con huella SHA-256.
- `/integridad` para detectar modificaciones posteriores.
- backup SQLite diario y `/backup` manual.
- persistencia Railway usando `DB_PATH=/data/bot_sismico.db`.

## Actualización en Railway
Reemplaza en GitHub los archivos del proyecto por los de v1.0 y haz commit.
Railway hará el redeploy automáticamente.

Mantén estas variables:
- `DISCORD_TOKEN`
- `EXCEL_PATH=Proyecto_Tabla_de_datos_con_coordenadas.xlsx`
- `CSN_CHECK_MINUTES=5`
- `DB_PATH=/data/bot_sismico.db`

Después del redeploy:
1. `/importar_todos`
2. `/integridad`
3. `/estado_alertas`
4. `/ranking`

## Nota de integridad
La primera importación de un código crea su huella. Si luego ese mismo código cambia
en la base, `/integridad` lo marca como MODIFICADA. Esto ayuda a auditar el experimento;
no convierte las correlaciones en predicciones sísmicas científicamente validadas.


## Corrección v1.0.1
Corrige el SyntaxError de construcción del mensaje de alertas y muestra correctamente los checks 2/3.


# v1.1 — Calculadoras

Nuevos comandos:

- `/calcular expresion:` — operaciones normales (+, -, *, /, //, %, ** y paréntesis).
- `/distancia` — distancia Haversine entre dos coordenadas usando la misma función del motor.
- `/comparar_magnitud` — magnitud predicha ± margen contra magnitud CSN.
- `/comparar_tiempo` — diferencia entre dos fechas/horas y margen permitido.
- `/comparar` — evaluación manual conjunta de tiempo + magnitud + ubicación, con resultado 3/3, 2/3 o fallo.
- `/radio` — consulta el radio automático que usaría el motor según magnitud y tipo de zona.

La calculadora normal NO usa `eval`; analiza únicamente operaciones matemáticas permitidas.

## v1.1.1
Corrección de arranque: se añadió `import ast`, requerido por la calculadora segura.


# v1.2 — Interfaz visual

Novedades:
- `/panel`: centro de control con botones.
- Embeds limpios para estado, ranking, historial y calculadoras.
- Centro de cálculo interactivo con botones:
  - Normal
  - Raíz
  - Distancia
  - Magnitud
  - Tiempo
  - Comparación 3/3
- `/raiz` para raíces directas.
- `/ayuda` organizado por categorías.
- Botón para actualizar CSN desde el panel.
- Las respuestas de herramientas del panel son privadas (ephemeral) para no llenar el canal.
- Todos los comandos anteriores se conservan.

Actualización Railway:
reemplaza `bot.py` en GitHub y haz Commit. No cambies `/data`, `worker-volume`
ni las variables de entorno.


## v1.2.2
Corrección UI: el botón Raíz ya no usa `√` como emoji de Discord; ahora usa un emoji Unicode válido.


# v1.3 — Centro de Control

El comando `/panel` ahora incluye:
- 🎯 Predicciones: resumen, ranking, historial e integridad.
- 🌎 Sismos CSN: últimos eventos, actualización manual y estado.
- 📥 Importar: acceso seguro a la rutina existente `/importar_todos`.
- 📊 Resultados.
- 🔔 Alertas.
- 🧮 Calculadora.
- 🩺 Diagnóstico del sistema.
- ⚙️ Estado.
- 📑 Historial.
- ❓ Ayuda.

La v1.3 mantiene las funciones y comandos anteriores. La interfaz no cambia
el esquema SQLite, `DB_PATH`, el volumen `/data` ni el motor de correlaciones.


# v1.3.1 — CSN Sync Fix
- El monitor ya no recorta a 30 eventos: procesa completos hoy y ayer.
- Usa `America/Santiago` para decidir cuál es el día actual del catálogo CSN.
- `events()` ordena de más reciente a más antiguo, corrigiendo “Últimos sismos”.
- `/csn_diagnostico` compara catálogo web vs SQLite y muestra eventos faltantes.
- `/actualizar_csn` importa todos los eventos recientes disponibles y conserva deduplicación por `source_id`.

# v1.4 — Sismologia Lab

Añade un dashboard web al mismo servicio Railway del bot.

- `/` Dashboard web responsive.
- `/api/dashboard` API JSON de solo lectura.
- `/health` comprobación del servidor web.
- Mapa: sismos CSN + ubicaciones de predicciones.
- Ranking por IA, métricas, últimos sismos y correlaciones.
- Refresco automático cada 60 segundos.
- Lee la misma base SQLite de `DB_PATH`; no crea una segunda base ni modifica el esquema.

## Railway
Reemplaza `bot.py`, `csn.py`, `requirements.txt` y añade `dashboard.py`.
Railway debe mantener `DB_PATH=/data/bot_sismico.db` y el volumen existente.
Después del deploy, en Railway crea/genera un dominio público para el servicio worker
(Networking / Public Networking). Ese dominio abrirá Sismologia Lab.

# v1.5 — Inspector + mapa avanzado + evaluador único
Discord y dashboard ahora usan `evaluator.py` como única fuente de verdad. El dashboard ya no limita la evaluación a 300 eventos; usa el mismo universo CSN que `/analizar_grupo`. Incluye mapa centrado en Chile, filtros por IA/estado, radios, líneas al sismo coincidente e inspector 3/3. `Coincidencias 3/3` refleja los aciertos del evaluador; el historial de alertas se muestra aparte.

## v1.5.1 — Feed CSN en canal independiente

Nuevos comandos:
- `/canal_sismos`: configura el canal donde se publican TODOS los sismos nuevos del CSN.
- `/estado_canal_sismos`: muestra el canal configurado.
- `/desactivar_canal_sismos`: desactiva ese feed.

El feed general NO usa @everyone. El canal `/canal_alertas` continúa separado para
correlaciones/casi-correlaciones con predicciones. Los eventos ya publicados quedan
registrados en SQLite para evitar duplicados tras reinicios.


## v1.5.2 - Informes PDF

Nuevo comando `/informe_pdf`.
- `/informe_pdf ia:TODAS` genera el informe general.
- `/informe_pdf ia:GROK` genera un informe de una sola IA.
- Incluye resumen, ranking, detalle de predicciones e inspector de coincidencias 3/3.
- Usa el mismo `evaluate_real` compartido por Discord y el dashboard.
- Requiere `reportlab`, añadido a `requirements.txt`.


## v1.5.3 — Chat IA general

Nuevos comandos:
- `/chat pregunta:...` — conversación general con IA.
- `/chat_nuevo` — reinicia el contexto reciente de ese usuario.
- Botón `💬 Chat IA` dentro de `/panel`.

Variables nuevas en Railway:
- `AI_API_KEY` — obligatoria para activar el chat.
- `AI_MODEL` — opcional; por defecto `gpt-5.6`.
- `AI_API_URL` — opcional; por defecto `https://api.openai.com/v1/responses`.
- `AI_MAX_OUTPUT_TOKENS` — opcional; por defecto 700.
- `AI_HISTORY_TURNS` — opcional; por defecto 6.

La clave NO debe escribirse en bot.py ni subirse a GitHub.
El chat conserva solo unas pocas intervenciones en memoria RAM y se reinicia cuando Railway reinicia.
El asistente recibe un resumen de solo lectura del proyecto sísmico para responder preguntas sobre
ranking, resultados y eventos guardados, pero también puede conversar de temas generales.

## v1.5.3.1 — Dependencies Fix
Corrige el crash de Railway `ModuleNotFoundError: No module named 'reportlab'`.
`requirements.txt` incluye ReportLab y las dependencias principales usadas por el proyecto.

## v1.5.4 — CSN Fresh Sync Fix
- El lector diario usa la tabla oficial del catálogo CSN como fuente primaria.
- Añade headers no-cache y query única para reducir respuestas web antiguas.
- IDs estables basados en hora/coordenadas para evitar duplicados.
- `/actualizar_csn` fuerza lectura fresca y publica nuevos eventos en el feed.
- `/csn_diagnostico` ya no considera sincronizado solo porque el parser y SQLite coincidan:
  también compara el último timestamp real leído contra el último guardado.

## v1.5.4.1 — CSN Compatibility Fix
Corrige el crash por `fetch_historical_events` faltante y restaura la interfaz que espera bot.py:
`fetch_recent_events()` vuelve a devolver `(eventos, errores)`. También mantiene compatibilidad
con `/importar_historico`, `save_csn_event()` y el feed de todos los sismos.

## v1.5.5 — CSN Multi-Day Chile/UTC
El monitor ya no decide un único "hoy". Consulta un margen de catálogos alrededor de
las fechas actuales de Chile y UTC (incluyendo 12/09 cuando UTC ya cruzó de día),
fusiona los eventos por ID y elige el timestamp más reciente.
`/csn_diagnostico` muestra exactamente cuántos eventos obtuvo de cada fecha consultada.


## v1.5.6 — Grok/xAI para /chat

`/chat` usa Grok mediante la API de xAI como proveedor principal.

Variables de Railway:
- `XAI_API_KEY` — clave de xAI.
- `XAI_MODEL` — por defecto `grok-4-fast`.
- `XAI_API_URL` — opcional, por defecto `https://api.x.ai/v1/chat/completions`.

Si `AI_API_KEY` sigue configurada, OpenAI queda como fallback opcional. Si no quieres
que se intente OpenAI, elimina `AI_API_KEY` de Railway.

La API key debe guardarse únicamente como variable secreta de Railway, nunca en GitHub.


## v1.5.7 — Sin IA
Se eliminó por completo el chat de IA:
- sin `/chat`
- sin `/chat_nuevo`
- sin botón `Chat IA`
- sin claves OpenAI/xAI
- sin llamadas a APIs de modelos

Se conservan las funciones sísmicas, CSN multi-día Chile/UTC, feed de sismos,
alertas de predicciones, dashboard, PDF, calculadora e importación/evaluación.


## v1.6 — Score de proximidad
Se conserva el resultado oficial estricto 3/3 (tiempo + magnitud + ubicación).
Se añade `/score codigo:GROK_03` con un índice descriptivo 0–100:
- Ubicación: 40%
- Magnitud: 30%
- Tiempo: 30%

El score NO convierte un fallo en acierto. Sirve para ordenar y medir qué tan cerca estuvo
una predicción. Para evitar resultados engañosos, el informe muestra también cada componente,
la distancia real y el sismo utilizado. El score no demuestra capacidad predictiva ni
significancia estadística por sí solo.


## v1.6.1 — Estado del sistema
Nuevo comando `/estado_sistema`:
- uptime del proceso
- RAM máxima usada
- carga CPU 1/5/15 min
- tamaño de SQLite
- número de predicciones
- número de sismos guardados
- último sismo guardado
- comprobación básica de SQLite/monitor

No añade dependencias nuevas y no hace consultas extra al CSN al abrir el panel.

## v1.6.2 — Estado sistema time-name fix
Corrige el crash `AttributeError: type object 'datetime.time' has no attribute 'time'`.
El uptime usa ahora `time.monotonic()` importado como `pytime`, evitando la colisión
con `datetime.time`.

## v1.6.3 — Score distance fix
Corrige `/score`: `proximity_score()` llamaba por error a `haversine_km`, pero
el evaluador existente define `haversine`. Se añadió una prueba funcional del score
además de la compilación.

## v1.6.5 — Score temporal NULL fix
Corrige `/score` cuando `daily_time_start` o `daily_time_end` vienen NULL desde SQLite.
En ese caso el score interpreta la predicción como día completo (00:00:00–23:59:59),
evitando valores inválidos como `2026-09-11TNone`.

## v1.6.6 — Score contract + fail-safe
Corrige el contrato real de `evaluate_real()`: devuelve 2 valores, no 4.
`/score` ahora hace defer inmediato y captura excepciones para que Discord muestre
un error útil en vez de `La aplicación no respondió`.

## v1.6.7 — Score import fix
Corrige `NameError: best_proximity is not defined`: `bot.py` ahora importa
explícitamente `best_proximity` desde evaluator.py. Se verifica además que la función
exista antes de empaquetar.
