# Bot Sísmico Discord v0.9 — preparado para hosting 24/7

Mantiene las funciones de v0.8.1:
- CSN automático
- histórico
- grupos reales
- `/ver_real`
- alertas automáticas
- `@everyone`

## Variables de entorno
Nunca escribas el token dentro de `bot.py`.

Configura en el hosting:
- `DISCORD_TOKEN` = token secreto del bot
- `EXCEL_PATH` = `Proyecto_Tabla_de_datos_con_coordenadas.xlsx`
- `DB_PATH` = `/data/bot_sismico.db` si montas un volumen persistente en `/data`
- `CSN_CHECK_MINUTES` = `5`

## Persistencia
SQLite necesita almacenamiento persistente. Si el hosting usa un sistema de archivos
efímero, monta un volumen en `/data` y configura `DB_PATH=/data/bot_sismico.db`.
Sin volumen, un redeploy/reinicio podría borrar la base SQLite.

El Excel está incluido en el proyecto. Para cambiarlo en producción, actualiza el archivo
del proyecto y vuelve a desplegar; luego ejecuta `/importar_todos`.

## Arranque
El proceso de producción es:
`python bot.py`

`railway.toml` y `Procfile` ya están incluidos.

## Después del despliegue
En Discord:
1. `/importar_todos`
2. `/canal_alertas`
3. `/estado_alertas`
4. opcional: `/importar_historico ...`

Mientras el servicio cloud esté activo, tu PC puede estar apagado.
