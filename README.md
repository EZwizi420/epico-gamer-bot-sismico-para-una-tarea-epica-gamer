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
