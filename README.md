# SismoBot — Edición Definitiva

Bot de Discord para registrar predicciones sísmicas, sincronizar eventos reales del Centro Sismológico Nacional de Chile (CSN), evaluar correlaciones y comparar resultados entre grupos/IA.

> Esta edición consolida la versión funcional histórica del proyecto con las mejoras finales de interfaz y análisis. Está pensada como copia maestra para GitHub y despliegue en Railway.

## Funciones principales

- Sincronización periódica con eventos del CSN.
- Importación de predicciones desde Excel por grupo/IA.
- Evaluación por ventana temporal, magnitud y distancia geográfica.
- Estados: acertada, casi acertada, no acertada y pendiente.
- Alertas automáticas de correlaciones y feed de sismos.
- Ranking priorizado por: más aciertos → más casi acertados → menos no acertados.
- Explorador interactivo de predicciones con filtros y selector de IA.
- Explorador interactivo de sismos reales (`/ver_real`).
- Perfil y comparación de IA, análisis de errores y centro de actividad.
- Mapa predicción vs. evento real (`/mapa`).
- Evolución del ranking (`/evolucion`) a partir de esta edición.
- Modo investigación (`/investigacion`).
- Exportación de resultados a Excel e informe PDF.
- Backups y diagnósticos de almacenamiento/CSN/sistema.
- Dashboard web incluido.

## Archivos

- `bot.py` — bot principal, comandos, paneles, ranking y QoL.
- `evaluator.py` — motor de evaluación y proximidad.
- `csn.py` — obtención y normalización de eventos CSN.
- `dashboard.py` — dashboard web.
- `Proyecto_Tabla_de_datos_con_coordenadas.xlsx` — plantilla/base de predicciones incluida en la distribución.
- `requirements.txt` — dependencias Python.
- `.env.example` — ejemplo de variables de entorno, sin secretos.
- `railway.toml` y `Procfile` — configuración de despliegue.

## Variables de entorno

Copia `.env.example` o configura estas variables directamente en Railway:

- `DISCORD_TOKEN` — token del bot. **Nunca lo subas a GitHub.**
- `EXCEL_PATH` — ruta del Excel. Por defecto `Proyecto_Tabla_de_datos_con_coordenadas.xlsx`.
- `DB_PATH` — ruta de SQLite. En Railway se recomienda `/data/bot_sismico.db` usando un volumen persistente.
- `CSN_CHECK_MINUTES` — intervalo de sincronización CSN; por defecto 5 minutos.

## Instalación local

```bash
python -m venv .venv
# Windows: .venv\\Scripts\\activate
# Linux/macOS: source .venv/bin/activate
pip install -r requirements.txt
```

Configura las variables de entorno y ejecuta:

```bash
python bot.py
```

## Railway

1. Sube este repositorio a GitHub.
2. Crea/conecta un proyecto Railway al repositorio.
3. Configura `DISCORD_TOKEN` y las demás variables necesarias.
4. Usa un volumen persistente montado en `/data` para conservar SQLite y backups.
5. Despliega. `railway.toml`/`Procfile` contienen la configuración base del proyecto.

**No reemplaces una base SQLite existente si estás actualizando un despliegue con historial.** El bot crea las estructuras nuevas que necesita sin requerir borrar el historial anterior.

## Comandos destacados

- `/panel` — centro de control visual.
- `/investigacion` — vista enfocada en resultados del proyecto.
- `/ranking` — ranking de IA/grupos.
- `/ver_real` — navegador interactivo de eventos reales para una predicción.
- `/mapa` — mapa de predicciones y eventos CSN; acepta IA/grupo opcional.
- `/evolucion` — evolución del ranking registrada desde esta versión.
- `/analizar_grupo` / `/analizar_todos` — evaluación de resultados.
- `/exportar_excel` — exportación automática de resultados.
- `/informe_pdf` — informe PDF.
- `/score` — proximidad de una predicción a su mejor evento.
- `/estado_sistema`, `/estado_csn`, `/csn_diagnostico` — diagnóstico.
- `/backup` — copia manual de SQLite.
- `/ayuda` — guía incorporada en Discord.

## Datos y seguridad

`.gitignore` excluye `.env`, bases SQLite y archivos compilados. No publiques tokens ni una copia de tu base de producción. Los datos persistentes de Railway deben mantenerse en el volumen y no dentro del repositorio.

## Nota metodológica

SismoBot sirve para registrar y evaluar estimaciones experimentales frente a eventos observados. No convierte esas estimaciones en predicciones sísmicas científicamente fiables ni debe utilizarse como sistema oficial de alerta de emergencias.

## Versión

**SismoBot — Edición Definitiva (QoL v3)**  
Consolidada: 5 de octubre de 2026.


## Hotfix 3.0.9: ranking, correlaciones e interfaz Leaflet

1. Despliega los archivos actualizados conservando las variables de entorno y el disco persistente.
2. En Discord, ejecuta `/reconstruir_correlaciones`. El comando sincroniza las predicciones desde el Excel y, si las fechas abarcan como máximo 31 días, descarga esos catálogos históricos del CSN y reconstruye el historial de coincidencias sin eliminar registros.
3. Si el rango supera 31 días, importa los catálogos con `/importar_historico` en bloques de hasta 31 días y vuelve a ejecutar `/reconstruir_correlaciones`.
4. El mapa Leaflet usa una capa oscura. El endpoint `/health` informa los conteos de predicciones, eventos CSN y registros de correlación para ayudar a diagnosticar la base de datos activa.

Nota: una correlación solo se registra si un evento CSN cumple al menos 2 de los 3 filtros actuales (tiempo, magnitud, ubicación). El código no fuerza coincidencias si los datos no cumplen los criterios.


## Render: bot y mapa en servicios separados (sin compartir el disco)

El Background Worker conserva la base original en su volumen `/data`. El Web Service del mapa mantiene una copia SQLite sincronizada por HTTPS cada 60 segundos; **no comparte ni monta el disco del bot**.

### Servicio original (Background Worker)

- Conserva su disco actual montado en `/data` y `DB_PATH=/data/bot_sismico.db`.
- Añade `MAP_SYNC_URL=https://TU-MAPA.onrender.com` (URL base del Web Service del mapa, sin `/api/sync-db`).
- Añade `MAP_SYNC_TOKEN` con un secreto aleatorio largo. No lo publiques en GitHub.

### Nuevo servicio del mapa (Web Service)

- Build Command: `pip install -r requirements.txt`
- Start Command: `gunicorn --bind 0.0.0.0:$PORT dashboard:app`
- Añade un disco persistente montado en `/data`.
- Variables: `DB_PATH=/data/bot_sismico.db` y el mismo `MAP_SYNC_TOKEN` del Background Worker.
- No configures `DISCORD_TOKEN` en el servicio del mapa.

El bot crea una instantánea consistente con la API de backup de SQLite y la envía al endpoint autenticado `/api/sync-db`. El mapa valida la integridad y las tablas requeridas antes de reemplazar atómicamente su copia. El endpoint no permite consultas ni escrituras SQL arbitrarias.

Tras desplegar ambos servicios, revisa `https://TU-MAPA.onrender.com/health`; debería devolver `ok: true` y contadores de predicciones, sismos y registros históricos. La primera sincronización se intenta al arrancar el bot y se repite cada 60 segundos.
