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
