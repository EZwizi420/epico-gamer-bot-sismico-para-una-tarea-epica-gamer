import os, re, math, sqlite3
from datetime import datetime, date, time, timedelta
from pathlib import Path

import discord
from discord import app_commands
from discord.ext import commands, tasks
from openpyxl import load_workbook

from csn import fetch_recent_events, fetch_historical_events

TOKEN = os.getenv("DISCORD_TOKEN")
EXCEL_PATH = os.getenv("EXCEL_PATH", "Proyecto_Tabla_de_datos_con_coordenadas.xlsx")
DB_PATH = os.getenv("DB_PATH", "bot_sismico.db")
CHECK_MINUTES = int(os.getenv("CSN_CHECK_MINUTES", "5"))
DEFAULT_MARGIN_HOURS = 2

def connect():
    con=sqlite3.connect(DB_PATH); con.row_factory=sqlite3.Row; return con

def init_db():
    with connect() as con:
        con.executescript("""
        CREATE TABLE IF NOT EXISTS predictions(
          id INTEGER PRIMARY KEY AUTOINCREMENT, code TEXT UNIQUE NOT NULL,
          group_name TEXT NOT NULL, predicted_at TEXT NOT NULL, place TEXT NOT NULL,
          magnitude REAL NOT NULL, mag_margin REAL NOT NULL,
          latitude REAL NOT NULL, longitude REAL NOT NULL, zone_type TEXT, geo_note TEXT
        );
        CREATE TABLE IF NOT EXISTS real_predictions(
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          code TEXT UNIQUE NOT NULL,
          group_name TEXT NOT NULL,
          prediction_no INTEGER NOT NULL,
          date_start TEXT NOT NULL,
          date_end TEXT NOT NULL,
          daily_time_start TEXT,
          daily_time_end TEXT,
          place TEXT NOT NULL,
          mag_min REAL NOT NULL,
          mag_max REAL NOT NULL,
          latitude REAL NOT NULL,
          longitude REAL NOT NULL,
          zone_type TEXT,
          geo_note TEXT,
          raw_date TEXT,
          raw_magnitude TEXT
        );
        CREATE TABLE IF NOT EXISTS bot_settings(
          key TEXT PRIMARY KEY,
          value TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS sent_alerts(
          event_source_id TEXT NOT NULL,
          prediction_code TEXT NOT NULL,
          sent_at TEXT NOT NULL,
          PRIMARY KEY(event_source_id, prediction_code)
        );
        CREATE TABLE IF NOT EXISTS prediction_snapshots(
          code TEXT PRIMARY KEY,
          group_name TEXT NOT NULL,
          frozen_at TEXT NOT NULL,
          fingerprint TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS correlation_history(
          event_source_id TEXT NOT NULL,
          prediction_code TEXT NOT NULL,
          match_level TEXT NOT NULL,
          detected_at TEXT NOT NULL,
          event_time TEXT NOT NULL,
          event_mag REAL NOT NULL,
          event_place TEXT,
          distance_km REAL NOT NULL,
          source_url TEXT,
          PRIMARY KEY(event_source_id,prediction_code,match_level)
        );
        CREATE TABLE IF NOT EXISTS observed_events(
          id INTEGER PRIMARY KEY AUTOINCREMENT, source TEXT NOT NULL, source_id TEXT NOT NULL,
          occurred_at TEXT NOT NULL, place TEXT, magnitude REAL NOT NULL,
          latitude REAL NOT NULL, longitude REAL NOT NULL, depth_km REAL, source_url TEXT,
          UNIQUE(source,source_id)
        );
        """)

def num(v):
    if v is None:return None
    m=re.search(r"-?\d+(?:[.,]\d+)?",str(v).replace("M","").replace("±",""))
    return float(m.group().replace(",",".")) if m else None

def combine_excel(d,t):
    if isinstance(d,datetime):d=d.date()
    elif not isinstance(d,date):
        d=next((datetime.strptime(str(d).strip(),f).date() for f in ("%d/%m/%Y","%d-%m-%Y","%Y-%m-%d")
                if _can_parse(str(d).strip(),f)),None)
    if isinstance(t,datetime):t=t.time()
    elif not isinstance(t,time):
        t=next((datetime.strptime(str(t).strip(),f).time() for f in ("%H:%M","%H:%M:%S")
                if _can_parse(str(t).strip(),f)),None)
    if not isinstance(d,date) or not isinstance(t,time):raise ValueError("Fecha/hora no reconocida")
    return datetime.combine(d,t).replace(second=0,microsecond=0)

def _can_parse(s,f):
    try: datetime.strptime(s,f); return True
    except ValueError:return False

def haversine(a,b,c,d):
    R=6371.0088;p1,p2=math.radians(a),math.radians(c)
    dp=math.radians(c-a);dl=math.radians(d-b)
    x=math.sin(dp/2)**2+math.cos(p1)*math.cos(p2)*math.sin(dl/2)**2
    return R*2*math.atan2(math.sqrt(x),math.sqrt(1-x))

def radius(mag,zone=""):
    if mag<3.5:r=20
    elif mag<4:r=25
    elif mag<4.5:r=35
    elif mag<5:r=50
    elif mag<5.5:r=70
    else:r=100
    z=(zone or "").lower()
    if "franja" in z:r=max(r,75)
    if "macro" in z or "región amplia" in z:r=max(r,100)
    return r

def import_excel():
    p=Path(EXCEL_PATH)
    if not p.exists():return 0,[f"No encontré {EXCEL_PATH}"]
    wb=load_workbook(p,data_only=True); ws=wb["PRUEBA"]
    n=0;errs=[]
    with connect() as con:
        for row in range(3,ws.max_row+1):
            code=ws.cell(row,3).value
            if not code:continue
            try:
                dt=combine_excel(ws.cell(row,4).value,ws.cell(row,5).value)
                vals=[num(ws.cell(row,c).value) for c in (7,8,9,10)]
                mag,margin,lat,lon=vals
                if None in vals:raise ValueError("faltan magnitud/margen/coordenadas")
                con.execute("""INSERT INTO predictions
                (code,group_name,predicted_at,place,magnitude,mag_margin,latitude,longitude,zone_type,geo_note)
                VALUES(?,'PRUEBA',?,?,?,?,?,?,?,?)
                ON CONFLICT(code) DO UPDATE SET predicted_at=excluded.predicted_at,place=excluded.place,
                magnitude=excluded.magnitude,mag_margin=excluded.mag_margin,latitude=excluded.latitude,
                longitude=excluded.longitude,zone_type=excluded.zone_type,geo_note=excluded.geo_note""",
                (str(code),dt.isoformat(),str(ws.cell(row,6).value or ""),mag,margin,lat,lon,
                 str(ws.cell(row,11).value or ""),str(ws.cell(row,12).value or "")))
                n+=1
            except Exception as e:errs.append(f"fila {row}: {e}")
    return n,errs

def save_csn_event(e):
    with connect() as con:
        before=con.total_changes
        con.execute("""INSERT OR IGNORE INTO observed_events
        (source,source_id,occurred_at,place,magnitude,latitude,longitude,depth_km,source_url)
        VALUES(?,?,?,?,?,?,?,?,?)""",
        ("CSN",e["source_id"],e["occurred_at"].isoformat(),e["place"],e["magnitude"],
         e["latitude"],e["longitude"],e.get("depth_km"),e.get("url")))
        return con.total_changes>before

def update_csn(limit=30):
    events,errors=fetch_recent_events(limit)
    new=0
    for e in events:new+=1 if save_csn_event(e) else 0
    return len(events),new,errors

def preds():
    with connect() as con:return con.execute("SELECT * FROM predictions WHERE group_name='PRUEBA' ORDER BY predicted_at").fetchall()
def events():
    with connect() as con:return con.execute("SELECT * FROM observed_events WHERE source='CSN' ORDER BY occurred_at").fetchall()

def evaluate(p,es,hours):
    target=datetime.fromisoformat(p["predicted_at"]);start=target-timedelta(hours=hours);end=target+timedelta(hours=hours)
    lo=p["magnitude"]-p["mag_margin"];hi=p["magnitude"]+p["mag_margin"];rad=radius(p["magnitude"],p["zone_type"])
    matches=[];cands=[]
    for e in es:
        when=datetime.fromisoformat(e["occurred_at"]);dist=haversine(p["latitude"],p["longitude"],e["latitude"],e["longitude"])
        x={"e":e,"t":start<=when<=end,"m":lo<=e["magnitude"]<=hi,"g":dist<=rad,
           "dist":dist,"mins":abs((when-target).total_seconds())/60,"dmag":abs(e["magnitude"]-p["magnitude"])}
        cands.append(x)
        if x["t"] and x["m"] and x["g"]:matches.append(x)
    matches.sort(key=lambda x:(x["mins"],x["dist"],x["dmag"]))
    now=datetime.now()
    status="ACERTADA" if matches else ("PENDIENTE" if now<=end else "NO ACERTADA")
    return status,(matches[0] if matches else None),rad,cands,start,end,lo,hi


# ---------- PREDICCIONES REALES POR GRUPO ----------

REAL_GROUPS = [
    "GROK", "DEEPSEEK", "CHATGPT", "CLAUDE", "PERPLEXITY",
    "META AI", "MISTRAL AI", "GEMINI", "COPILOT", "QWEN"
]

MONTHS_ES = {
    "enero":1,"febrero":2,"marzo":3,"abril":4,"mayo":5,"junio":6,
    "julio":7,"agosto":8,"septiembre":9,"setiembre":9,"octubre":10,
    "noviembre":11,"diciembre":12,"sept":9,"sep":9
}

def parse_mag_range(raw):
    """Devuelve min/max sin inventar margen."""
    s = str(raw or "").strip().lower().replace("m","").replace(",",".")
    nums = [float(x) for x in re.findall(r"\d+(?:\.\d+)?", s)]
    if not nums:
        raise ValueError("magnitud vacía/no reconocida")
    if len(nums) == 1:
        return nums[0], nums[0]
    return min(nums[0], nums[1]), max(nums[0], nums[1])

def parse_real_date_window(raw, year_default=2026):
    """Interpreta formatos existentes en el Excel real.

    Retorna date_start, date_end, daily_time_start, daily_time_end.
    Las horas, si existen, son bandas diarias dentro del rango de fechas.
    """
    s = str(raw or "").strip().lower()
    if not s:
        raise ValueError("fecha vacía")

    # Hora opcional: 00:00-06:00, 18:00-00:00, etc.
    tm = re.search(r"(\d{1,2}:\d{2})\s*[-–]\s*(\d{1,2}:\d{2})", s)
    t_start = tm.group(1) if tm else None
    t_end = tm.group(2) if tm else None

    # Formato 10 - 13/09/2026
    m = re.search(r"(\d{1,2})\s*[-–]\s*(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(20\d{2})", s)
    if m:
        d1,d2,mo,yr = map(int,m.groups())
        return date(yr,mo,d1), date(yr,mo,d2), t_start, t_end

    # Formato 10-13 de septiembre / 11-13 septiembre / 11-13 sept
    m = re.search(r"(\d{1,2})\s*[-–]\s*(\d{1,2})(?:\s+de)?\s+([a-záéíóú]+)", s)
    if m:
        d1,d2,mon = m.groups()
        mo = MONTHS_ES.get(mon)
        if not mo:
            raise ValueError(f"mes no reconocido: {mon}")
        return date(year_default,mo,int(d1)), date(year_default,mo,int(d2)), t_start, t_end

    # Formato explícito DD/MM/YYYY - DD/MM/YYYY si se agrega después.
    dates = re.findall(r"(\d{1,2})/(\d{1,2})/(20\d{2})", s)
    if len(dates) >= 2:
        a,b = dates[0],dates[1]
        return date(int(a[2]),int(a[1]),int(a[0])), date(int(b[2]),int(b[1]),int(b[0])), t_start,t_end

    raise ValueError(f"rango de fecha no reconocido: {raw}")

def import_real_group(group_name):
    group_name = group_name.strip().upper()
    actual = next((x for x in REAL_GROUPS if x.upper() == group_name), None)
    if not actual:
        return 0, 0, [f"Grupo desconocido: {group_name}"]

    p = Path(EXCEL_PATH)
    if not p.exists():
        return 0, 0, [f"No encontré {EXCEL_PATH}"]

    wb = load_workbook(p, data_only=True)
    if actual not in wb.sheetnames:
        return 0, 0, [f"No existe la hoja {actual}"]

    ws = wb[actual]
    imported, skipped, errors = 0, 0, []

    with connect() as con:
        for row in range(3, ws.max_row + 1):
            pred_no = ws.cell(row,3).value
            raw_date = ws.cell(row,4).value
            place = ws.cell(row,5).value
            raw_mag = ws.cell(row,6).value
            lat = num(ws.cell(row,8).value)
            lon = num(ws.cell(row,9).value)
            zone = str(ws.cell(row,10).value or "")
            note = str(ws.cell(row,11).value or "")

            # Filas plantilla/vacías no son errores.
            if not raw_date and not place and not raw_mag:
                skipped += 1
                continue

            try:
                if pred_no is None:
                    raise ValueError("número de predicción vacío")
                if not place:
                    raise ValueError("lugar vacío")
                if lat is None or lon is None:
                    raise ValueError("faltan coordenadas")
                d1,d2,t1,t2 = parse_real_date_window(raw_date)
                m1,m2 = parse_mag_range(raw_mag)
                n = int(float(pred_no))
                code = f"{actual.replace(' ','_')}_{n:02d}"

                con.execute("""INSERT INTO real_predictions
                (code,group_name,prediction_no,date_start,date_end,daily_time_start,daily_time_end,
                 place,mag_min,mag_max,latitude,longitude,zone_type,geo_note,raw_date,raw_magnitude)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                ON CONFLICT(code) DO UPDATE SET
                  group_name=excluded.group_name,prediction_no=excluded.prediction_no,
                  date_start=excluded.date_start,date_end=excluded.date_end,
                  daily_time_start=excluded.daily_time_start,daily_time_end=excluded.daily_time_end,
                  place=excluded.place,mag_min=excluded.mag_min,mag_max=excluded.mag_max,
                  latitude=excluded.latitude,longitude=excluded.longitude,
                  zone_type=excluded.zone_type,geo_note=excluded.geo_note,
                  raw_date=excluded.raw_date,raw_magnitude=excluded.raw_magnitude""",
                (code,actual,n,d1.isoformat(),d2.isoformat(),t1,t2,str(place),m1,m2,
                 lat,lon,zone,note,str(raw_date),str(raw_mag)))
                imported += 1
            except Exception as exc:
                errors.append(f"fila {row}: {exc}")

    freeze_predictions()
    return imported, skipped, errors

def real_time_ok(event_dt, p):
    d1 = date.fromisoformat(p["date_start"])
    d2 = date.fromisoformat(p["date_end"])
    if not (d1 <= event_dt.date() <= d2):
        return False

    t1s, t2s = p["daily_time_start"], p["daily_time_end"]
    if not t1s or not t2s:
        return True

    t1 = datetime.strptime(t1s,"%H:%M").time()
    t2 = datetime.strptime(t2s,"%H:%M").time()
    et = event_dt.time()

    # 18:00-00:00 se interpreta hasta fin del día.
    if t2 == time(0,0):
        return et >= t1
    if t1 <= t2:
        return t1 <= et <= t2
    # Soporta bandas que crucen medianoche.
    return et >= t1 or et <= t2

def real_window_finished(p):
    d2 = date.fromisoformat(p["date_end"])
    t2s = p["daily_time_end"]
    if t2s:
        t2 = datetime.strptime(t2s,"%H:%M").time()
        if t2 == time(0,0):
            end_dt = datetime.combine(d2, time(23,59,59))
        else:
            end_dt = datetime.combine(d2,t2)
    else:
        end_dt = datetime.combine(d2,time(23,59,59))
    return datetime.now() > end_dt

def evaluate_real(p, es):
    rad = radius((p["mag_min"] + p["mag_max"]) / 2, p["zone_type"])
    matches, candidates = [], []

    for e in es:
        when = datetime.fromisoformat(e["occurred_at"])
        dist = haversine(p["latitude"],p["longitude"],e["latitude"],e["longitude"])
        t_ok = real_time_ok(when,p)
        m_ok = p["mag_min"] <= e["magnitude"] <= p["mag_max"]
        g_ok = dist <= rad
        x = {"e":e,"t":t_ok,"m":m_ok,"g":g_ok,"dist":dist}
        candidates.append(x)
        if t_ok and m_ok and g_ok:
            matches.append(x)

    # Mejor coincidencia: menor distancia; luego magnitud más central.
    center = (p["mag_min"] + p["mag_max"]) / 2
    matches.sort(key=lambda x:(x["dist"], abs(x["e"]["magnitude"]-center)))

    status = "ACERTADA" if matches else ("NO ACERTADA" if real_window_finished(p) else "PENDIENTE")
    return status, (matches[0] if matches else None), rad, candidates

def get_real_group(group_name=None):
    with connect() as con:
        if group_name:
            return con.execute(
                "SELECT * FROM real_predictions WHERE UPPER(group_name)=UPPER(?) ORDER BY prediction_no",
                (group_name,)
            ).fetchall()
        return con.execute(
            "SELECT * FROM real_predictions ORDER BY group_name,prediction_no"
        ).fetchall()




# ---------- V1.0: INTEGRIDAD, ESTADÍSTICAS E HISTORIAL ----------

def prediction_fingerprint(p):
    fields = (
        p["code"], p["group_name"], p["prediction_no"], p["date_start"], p["date_end"],
        p["daily_time_start"], p["daily_time_end"], p["place"], p["mag_min"], p["mag_max"],
        p["latitude"], p["longitude"], p["zone_type"], p["geo_note"]
    )
    return hashlib.sha256(repr(fields).encode("utf-8")).hexdigest()

def freeze_predictions():
    """Registra por primera vez cada predicción. No altera la predicción."""
    now=datetime.now().isoformat()
    added=0
    with connect() as con:
        for p in con.execute("SELECT * FROM real_predictions").fetchall():
            fp=prediction_fingerprint(p)
            row=con.execute("SELECT fingerprint FROM prediction_snapshots WHERE code=?",(p["code"],)).fetchone()
            if row is None:
                con.execute("""INSERT INTO prediction_snapshots(code,group_name,frozen_at,fingerprint)
                               VALUES(?,?,?,?)""",(p["code"],p["group_name"],now,fp))
                added+=1
    return added

def integrity_status(p):
    with connect() as con:
        row=con.execute("SELECT * FROM prediction_snapshots WHERE code=?",(p["code"],)).fetchone()
    if not row:
        return "NO_CONGELADA", None
    return ("OK" if row["fingerprint"]==prediction_fingerprint(p) else "MODIFICADA"), row["frozen_at"]

def save_correlation_history(e,p,level,dist):
    with connect() as con:
        con.execute("""INSERT OR IGNORE INTO correlation_history
          (event_source_id,prediction_code,match_level,detected_at,event_time,event_mag,event_place,distance_km,source_url)
          VALUES(?,?,?,?,?,?,?,?,?)""",
          (str(e["source_id"]),p["code"],level,datetime.now().isoformat(),e["occurred_at"],
           e["magnitude"],e["place"],dist,e["source_url"]))

def backup_database():
    src=Path(DB_PATH)
    if not src.exists():
        return None
    backup_dir=src.parent/"backups"
    backup_dir.mkdir(parents=True,exist_ok=True)
    out=backup_dir/f"bot_sismico_{datetime.now():%Y%m%d_%H%M%S}.db"
    # SQLite online backup: safe while the bot is running.
    source=sqlite3.connect(src)
    target=sqlite3.connect(out)
    try:
        source.backup(target)
    finally:
        target.close(); source.close()
    # Keep latest 14 local backups.
    files=sorted(backup_dir.glob("bot_sismico_*.db"), key=lambda x:x.stat().st_mtime, reverse=True)
    for old in files[14:]:
        try: old.unlink()
        except: pass
    return out


# ---------- ALERTAS AUTOMÁTICAS ----------

def set_setting(key, value):
    with connect() as con:
        con.execute(
            """INSERT INTO bot_settings(key,value) VALUES(?,?)
               ON CONFLICT(key) DO UPDATE SET value=excluded.value""",
            (key, str(value))
        )

def get_setting(key):
    with connect() as con:
        row=con.execute("SELECT value FROM bot_settings WHERE key=?",(key,)).fetchone()
    return row["value"] if row else None

def alert_was_sent(event_id, prediction_code):
    with connect() as con:
        row=con.execute(
            "SELECT 1 FROM sent_alerts WHERE event_source_id=? AND prediction_code=?",
            (str(event_id), prediction_code)
        ).fetchone()
    return row is not None

def mark_alert_sent(event_id, prediction_code):
    with connect() as con:
        con.execute(
            "INSERT OR IGNORE INTO sent_alerts(event_source_id,prediction_code,sent_at) VALUES(?,?,?)",
            (str(event_id), prediction_code, datetime.now().isoformat())
        )

def event_matches_real_prediction(e, p):
    when=datetime.fromisoformat(e["occurred_at"])
    rad=radius((p["mag_min"]+p["mag_max"])/2,p["zone_type"])
    dist=haversine(p["latitude"],p["longitude"],e["latitude"],e["longitude"])
    return {
        "time": real_time_ok(when,p),
        "mag": p["mag_min"] <= e["magnitude"] <= p["mag_max"],
        "geo": dist <= rad,
        "dist": dist,
        "radius": rad,
    }

async def send_alerts_for_new_events(new_events):
    channel_id=get_setting("alert_channel_id")
    if not channel_id or not new_events:
        return 0

    try:
        channel=bot.get_channel(int(channel_id)) or await bot.fetch_channel(int(channel_id))
    except Exception as exc:
        print("No pude abrir el canal de alertas:",exc)
        return 0

    predictions=get_real_group()
    sent=0

    for e in new_events:
        when=datetime.fromisoformat(e["occurred_at"])
        for p in predictions:
            check=event_matches_real_prediction(e,p)

            passed=sum((check["time"],check["mag"],check["geo"]))
            if passed < 2:
                continue

            level="COINCIDENCIA" if passed==3 else "CASI"
            dedupe_code=p["code"] if level=="COINCIDENCIA" else p["code"]+"#CASI"
            if alert_was_sent(e["source_id"],dedupe_code):
                continue

            pred_time=""
            if p["daily_time_start"]:
                pred_time=f"\n🕒 Horario predicho: **{p['daily_time_start']}–{p['daily_time_end']}**"

            msg=(
                ("@everyone\n🚨 **CORRELACIÓN SÍSMICA DETECTADA**\n\n" if level=="COINCIDENCIA"
                 else "🟡 **CASI COINCIDENCIA SÍSMICA (2/3 filtros)**\n\n") +
                "🌐 **Sismo observado por el CSN**\n"
                f"📅 Fecha: **{when:%d/%m/%Y}**\n"
                f"🕒 Hora local: **{when:%H:%M:%S}**\n"
                f"📈 Magnitud: **M{e['magnitude']:.1f}**\n"
                f"📍 Lugar: **{e['place'] or 'Sin referencia'}**\n"
                f"🌎 Coordenadas: **{e['latitude']:.4f}, {e['longitude']:.4f}**\n"
                + (f"⬇️ Profundidad: **{e['depth_km']:.1f} km**\n" if e["depth_km"] is not None else "") +
                "\n🎯 **Predicción relacionada**\n"
                f"🤖 IA/grupo: **{p['group_name']}**\n"
                f"🆔 Predicción: **{p['code']}**\n"
                f"📅 Ventana predicha: **{p['date_start']} → {p['date_end']}**"
                f"{pred_time}\n"
                f"📈 Magnitud predicha: **M{p['mag_min']:.1f}–M{p['mag_max']:.1f}**\n"
                f"📍 Lugar predicho: **{p['place']}**\n"
                f"📏 Distancia a referencia: **{check['dist']:.1f} km** "
                f"(radio permitido {check['radius']} km)\n\n"
                f"Filtros: Tiempo {'✅' if check['time'] else '❌'} · "
                f"Magnitud {'✅' if check['mag'] else '❌'} · "
                f"Ubicación {'✅' if check['geo'] else '❌'}"
            )
            if e["source_url"]:
                msg += f"\n🔗 **Informe oficial CSN:**\n{e['source_url']}"

            try:
                await channel.send(
                    msg[:1950],
                    allowed_mentions=discord.AllowedMentions(
                        everyone=True,
                        users=False,
                        roles=False,
                        replied_user=False
                    )
                )
                mark_alert_sent(e["source_id"],dedupe_code)
                save_correlation_history(e,p,level,check["dist"])
                sent+=1
            except Exception as exc:
                print("Error enviando alerta:",exc)

    return sent


intents=discord.Intents.default()
bot=commands.Bot(command_prefix="!",intents=intents)

@tasks.loop(minutes=CHECK_MINUTES)
async def csn_loop():
    try:
        fetched,errs=fetch_recent_events(30)
        newly_saved=[]
        for e in fetched:
            if save_csn_event(e):
                newly_saved.append(e)

        if newly_saved:
            print(f"CSN: {len(newly_saved)} eventos nuevos ({len(fetched)} leídos)")
            await send_alerts_for_new_events(newly_saved)

        if errs:
            print("CSN avisos:",errs[:3])
    except Exception as e:
        print("Error actualizando CSN:",e)

@csn_loop.before_loop
async def before_csn():await bot.wait_until_ready()


@tasks.loop(hours=24)
async def backup_loop():
    try:
        out=backup_database()
        if out: print("Backup SQLite:",out)
    except Exception as exc:
        print("Error backup SQLite:",exc)

@bot.event
async def on_ready():
    await bot.tree.sync()
    if not csn_loop.is_running():csn_loop.start()
    if not backup_loop.is_running():backup_loop.start()
    freeze_predictions()
    print(f"Conectado: {bot.user}. CSN cada {CHECK_MINUTES} min. Backup diario activo.")

@bot.tree.command(name="importar_pruebas",description="Importa PRUEBA desde el Excel")
async def c_import(i:discord.Interaction):
    n,e=import_excel();await i.response.send_message(f"📥 {n} predicciones importadas."+("\n⚠️ "+"; ".join(e[:3]) if e else ""))

@bot.tree.command(name="actualizar_csn",description="Consulta ahora los últimos eventos reales del CSN")
async def c_csn(i:discord.Interaction):
    await i.response.defer()
    try:
        total,new,errs=update_csn()
        msg=f"🌐 CSN consultado: **{total}** informes leídos, **{new}** nuevos guardados."
        if errs:msg+=f"\n⚠️ {len(errs)} informe(s) no pudieron interpretarse. Primero: `{errs[0][:250]}`"
        await i.followup.send(msg[:1900])
    except Exception as e:await i.followup.send(f"❌ Error consultando CSN: `{e}`")

@bot.tree.command(name="analizar",description="Evalúa PRUEBA usando únicamente eventos reales CSN guardados")
async def c_an(i:discord.Interaction,margen_horas:app_commands.Range[int,1,2]=2):
    ps,es=preds(),events()
    if not ps:await i.response.send_message("Usa `/importar_pruebas`.");return
    if not es:await i.response.send_message("No hay eventos CSN guardados. Usa `/actualizar_csn`.");return
    lines=[f"📊 **CSN REAL · ±{margen_horas} h**"]
    for p in ps:
        st,b,rad,*_=evaluate(p,es,margen_horas)
        icon={"ACERTADA":"🟢","PENDIENTE":"🟡","NO ACERTADA":"🔴"}[st]
        lines.append(f"{icon} **{p['code']} — {st}** · radio {rad} km")
    await i.response.send_message("\n".join(lines)[:1950])

@bot.tree.command(name="ver",description="Detalle de una predicción frente a eventos reales del CSN")
async def c_ver(i:discord.Interaction,codigo:str,margen_horas:app_commands.Range[int,1,2]=2):
    with connect() as con:p=con.execute("SELECT * FROM predictions WHERE UPPER(code)=UPPER(?)",(codigo,)).fetchone()
    if not p:await i.response.send_message("❌ No encontrada.");return
    st,b,rad,cands,start,end,lo,hi=evaluate(p,events(),margen_horas)
    icon={"ACERTADA":"🟢","PENDIENTE":"🟡","NO ACERTADA":"🔴"}[st]
    dt=datetime.fromisoformat(p["predicted_at"])
    text=(f"{icon} **{p['code']} — {st}**\nPredicción: {dt:%d/%m/%Y %H:%M}\n📍 {p['place']} "
          f"({p['latitude']:.4f}, {p['longitude']:.4f})\n📈 M{p['magnitude']:.1f} ±{p['mag_margin']:.1f} "
          f"→ M{lo:.1f}–M{hi:.1f}\n📏 Radio: {rad} km\n⏱️ Ventana: {start:%d/%m %H:%M}–{end:%d/%m %H:%M}")
    if b:
        e=b["e"];when=datetime.fromisoformat(e["occurred_at"])
        text+=(f"\n\n🌐 **Evento CSN #{e['source_id']}**\n{when:%d/%m/%Y %H:%M:%S} · M{e['magnitude']:.1f}\n"
               f"{e['place']}\n📍 {e['latitude']:.4f}, {e['longitude']:.4f}\n"
               f"Distancia **{b['dist']:.1f} km** · Δt **{b['mins']:.0f} min** · ΔM **{b['dmag']:.2f}**\n"
               f"{e['source_url'] or ''}")
    await i.response.send_message(text[:1950])

@bot.tree.command(name="estado_csn",description="Muestra cuántos eventos reales CSN tiene guardados el bot")
async def c_state(i:discord.Interaction):
    with connect() as con:
        row=con.execute("SELECT COUNT(*) n, MAX(occurred_at) last FROM observed_events WHERE source='CSN'").fetchone()
    await i.response.send_message(f"🌐 Eventos CSN guardados: **{row['n']}**\nÚltimo: **{row['last'] or 'ninguno'}**")


def parse_user_date(text):
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(text.strip(), fmt).date()
        except ValueError:
            pass
    raise ValueError("Usa DD/MM/AAAA, por ejemplo 10/09/2026.")

@bot.tree.command(
    name="importar_historico",
    description="Importa eventos históricos del catálogo diario oficial del CSN"
)
@app_commands.describe(
    fecha_inicio="DD/MM/AAAA",
    fecha_fin="DD/MM/AAAA"
)
async def c_history(
    i: discord.Interaction,
    fecha_inicio: str,
    fecha_fin: str
):
    await i.response.defer()
    try:
        start = parse_user_date(fecha_inicio)
        end = parse_user_date(fecha_fin)

        if end < start:
            await i.followup.send("❌ La fecha final no puede ser anterior a la inicial.")
            return

        # Protección para evitar descargas accidentales enormes desde Discord.
        days = (end - start).days + 1
        if days > 31:
            await i.followup.send(
                "❌ Por seguridad, importa como máximo 31 días por comando."
            )
            return

        historical, errs = fetch_historical_events(start, end)

        new = 0
        for event in historical:
            if save_csn_event(event):
                new += 1

        msg = (
            f"📚 **HISTÓRICO CSN IMPORTADO**\n"
            f"Período: **{start:%d/%m/%Y} → {end:%d/%m/%Y}**\n"
            f"Eventos interpretados: **{len(historical)}**\n"
            f"Nuevos guardados: **{new}**\n"
            f"Ya existentes: **{len(historical) - new}**"
        )

        if errs:
            msg += (
                f"\n⚠️ Avisos/errores: **{len(errs)}**\n"
                f"Primero: `{errs[0][:300]}`"
            )

        await i.followup.send(msg[:1900])

    except Exception as exc:
        await i.followup.send(f"❌ Error importando histórico CSN: `{exc}`")

@bot.tree.command(
    name="ultimos_csn",
    description="Muestra los últimos eventos CSN guardados"
)
async def c_last_csn(i: discord.Interaction):
    with connect() as con:
        rows = con.execute("""
            SELECT source_id, occurred_at, place, magnitude, depth_km
            FROM observed_events
            WHERE source='CSN'
            ORDER BY occurred_at DESC
            LIMIT 10
        """).fetchall()

    if not rows:
        await i.response.send_message(
            "No hay eventos CSN guardados. Usa `/actualizar_csn` o `/importar_historico`."
        )
        return

    lines = ["🌐 **ÚLTIMOS EVENTOS CSN GUARDADOS**"]
    for e in rows:
        when = datetime.fromisoformat(e["occurred_at"])
        depth = f" · {e['depth_km']:.1f} km prof." if e["depth_km"] is not None else ""
        lines.append(
            f"**#{e['source_id']}** · {when:%d/%m/%Y %H:%M:%S} · "
            f"M{e['magnitude']:.1f}{depth}\n{e['place'] or 'Sin referencia'}"
        )

    await i.response.send_message("\n".join(lines)[:1950])



@bot.tree.command(name="importar_grupo",description="Importa una hoja/IA real del Excel")
@app_commands.describe(grupo="Ej.: CHATGPT, GROK, CLAUDE, COPILOT")
async def c_import_group(i:discord.Interaction,grupo:str):
    n,skipped,errs=import_real_group(grupo)
    msg=f"📥 **{grupo.upper()}**: {n} predicciones importadas."
    if skipped: msg+=f"\nFilas vacías omitidas: {skipped}"
    if errs: msg+=f"\n⚠️ Errores: {len(errs)} · Primero: `{errs[0][:300]}`"
    await i.response.send_message(msg[:1900])

@bot.tree.command(name="importar_todos",description="Importa todas las IA con datos completos del Excel")
async def c_import_all(i:discord.Interaction):
    await i.response.defer()
    lines=["📚 **IMPORTACIÓN DE GRUPOS REALES**"]
    total=0
    for group in REAL_GROUPS:
        n,skipped,errs=import_real_group(group)
        total+=n
        if n:
            lines.append(f"✅ **{group}**: {n}")
        elif errs:
            lines.append(f"⚠️ **{group}**: 0 · {errs[0][:100]}")
        else:
            lines.append(f"➖ **{group}**: sin datos")
    lines.append(f"\nTotal importado/actualizado: **{total}**")
    await i.followup.send("\n".join(lines)[:1950])

@bot.tree.command(name="analizar_grupo",description="Evalúa una IA contra los eventos reales CSN")
@app_commands.describe(grupo="Ej.: CHATGPT, GROK, CLAUDE, COPILOT")
async def c_an_group(i:discord.Interaction,grupo:str):
    ps=get_real_group(grupo)
    es=events()
    if not ps:
        await i.response.send_message(f"No hay predicciones importadas para **{grupo}**.")
        return
    if not es:
        await i.response.send_message("No hay eventos CSN. Usa `/importar_historico`.")
        return

    lines=[f"📊 **{grupo.upper()} · CSN REAL**"]
    hits=0; finished=0
    for p in ps:
        st,b,rad,_=evaluate_real(p,es)
        icon={"ACERTADA":"🟢","PENDIENTE":"🟡","NO ACERTADA":"🔴"}[st]
        hits += st=="ACERTADA"
        finished += st!="PENDIENTE"
        lines.append(f"{icon} **{p['code']} — {st}** · M{p['mag_min']:.1f}–{p['mag_max']:.1f} · {rad} km")
    if finished:
        lines.append(f"\nResultado cerrado hasta ahora: **{hits}/{finished} acertadas ({hits/finished*100:.1f}%)**")
    pending=len(ps)-finished
    if pending: lines.append(f"Pendientes: **{pending}**")
    await i.response.send_message("\n".join(lines)[:1950])

@bot.tree.command(name="analizar_todos",description="Compara todos los grupos reales importados")
async def c_an_all(i:discord.Interaction):
    ps=get_real_group()
    es=events()
    if not ps:
        await i.response.send_message("No hay grupos reales importados. Usa `/importar_todos`.")
        return
    if not es:
        await i.response.send_message("No hay eventos CSN. Usa `/importar_historico`.")
        return

    groups={}
    for p in ps: groups.setdefault(p["group_name"],[]).append(p)
    lines=["🏁 **COMPARACIÓN DE IA · CSN REAL**"]
    for group,gps in groups.items():
        hits=misses=pending=0
        for p in gps:
            st,_,_,_=evaluate_real(p,es)
            if st=="ACERTADA": hits+=1
            elif st=="NO ACERTADA": misses+=1
            else: pending+=1
        closed=hits+misses
        score=f"{hits}/{closed} ({hits/closed*100:.1f}%)" if closed else "sin resultados cerrados"
        lines.append(f"**{group}** → {score} · 🟡 {pending} pendientes")
    await i.response.send_message("\n".join(lines)[:1950])

@bot.tree.command(name="ver_real",description="Detalle y enlace CSN de una predicción real")
async def c_ver_real(i:discord.Interaction,codigo:str):
    with connect() as con:
        p=con.execute(
            "SELECT * FROM real_predictions WHERE UPPER(code)=UPPER(?)",
            (codigo,)
        ).fetchone()

    if not p:
        await i.response.send_message("❌ Predicción real no encontrada.")
        return

    es = events()
    st,b,rad,cands = evaluate_real(p,es)
    icon={"ACERTADA":"🟢","PENDIENTE":"🟡","NO ACERTADA":"🔴"}[st]

    timeband=""
    if p["daily_time_start"]:
        timeband=f" · horario diario {p['daily_time_start']}–{p['daily_time_end']}"

    text=(
        f"{icon} **{p['code']} — {st}**\n"
        f"Grupo: **{p['group_name']}**\n"
        f"📅 {p['date_start']} → {p['date_end']}{timeband}\n"
        f"📍 {p['place']} ({p['latitude']:.4f}, {p['longitude']:.4f})\n"
        f"📈 M{p['mag_min']:.1f}–M{p['mag_max']:.1f}\n"
        f"📏 Radio usado: **{rad} km**"
    )

    chosen = b
    chosen_title = "🌐 **SISMO COINCIDENTE DEL CSN**"

    if chosen is None and cands:
        # Elegir un candidato explicable:
        # 1) preferir eventos dentro de la ventana temporal;
        # 2) penalizar distancia fuera del radio y magnitud fuera del rango.
        center=(p["mag_min"]+p["mag_max"])/2
        temporal=[x for x in cands if x["t"]]
        pool=temporal if temporal else cands

        def score(x):
            e=x["e"]
            distance_penalty=x["dist"]/max(rad,1)
            mag_half=max((p["mag_max"]-p["mag_min"])/2,0.25)
            mag_penalty=abs(e["magnitude"]-center)/mag_half
            # Si no está en tiempo y no había candidatos temporales, penalizarlo.
            time_penalty=0 if x["t"] else 5
            return time_penalty + distance_penalty + mag_penalty

        chosen=min(pool,key=score)
        chosen_title="🌐 **SISMO CSN MÁS APROXIMADO**"

    if chosen:
        e=chosen["e"]
        when=datetime.fromisoformat(e["occurred_at"])
        text+=(
            f"\n\n{chosen_title}\n"
            f"📅 {when:%d/%m/%Y %H:%M:%S}\n"
            f"📈 M{e['magnitude']:.1f}\n"
            f"📍 {e['place'] or 'Sin referencia geográfica'}\n"
            f"🌎 {e['latitude']:.4f}, {e['longitude']:.4f}\n"
            f"📏 Distancia: **{chosen['dist']:.1f} km**\n\n"
            f"Tiempo: {'✅' if chosen['t'] else '❌'}\n"
            f"Magnitud: {'✅' if chosen['m'] else '❌'}\n"
            f"Ubicación: {'✅' if chosen['g'] else '❌'}"
        )
        if e["source_url"]:
            text+=f"\n\n🔗 **Informe oficial CSN:**\n{e['source_url']}"
    else:
        text+="\n\nNo hay eventos CSN guardados para comparar."

    await i.response.send_message(text[:1950])



@bot.tree.command(name="canal_alertas",description="Usa este canal para las alertas automáticas de correlaciones")
async def c_alert_channel(i:discord.Interaction):
    set_setting("alert_channel_id", i.channel_id)
    await i.response.send_message(
        f"🔔 Alertas activadas en <#{i.channel_id}>.\n"
        "Cuando llegue un sismo NUEVO del CSN que cumpla tiempo + magnitud + ubicación "
        "de una predicción real importada, enviaré aquí toda la información."
    )

@bot.tree.command(name="estado_alertas",description="Muestra dónde están configuradas las alertas")
async def c_alert_state(i:discord.Interaction):
    channel_id=get_setting("alert_channel_id")
    if channel_id:
        await i.response.send_message(
            f"🔔 Canal de alertas: <#{channel_id}>\n"
            f"Consulta CSN: cada **{CHECK_MINUTES} minutos**.\n"
            "Criterio: tiempo ✅ + magnitud ✅ + ubicación/radio ✅."
        )
    else:
        await i.response.send_message(
            "🔕 No hay canal configurado. Ejecuta `/canal_alertas` en el canal deseado."
        )

@bot.tree.command(name="desactivar_alertas",description="Desactiva las notificaciones automáticas")
async def c_alert_off(i:discord.Interaction):
    with connect() as con:
        con.execute("DELETE FROM bot_settings WHERE key='alert_channel_id'")
    await i.response.send_message("🔕 Alertas automáticas desactivadas.")



@bot.tree.command(name="ranking",description="Ranking de aciertos cerrados por IA")
async def c_ranking(i:discord.Interaction):
    ps=get_real_group()
    es=events()
    groups={}
    for p in ps: groups.setdefault(p["group_name"],[]).append(p)
    rows=[]
    for group,gps in groups.items():
        hits=misses=pending=0
        for p in gps:
            st,_,_,_=evaluate_real(p,es)
            if st=="ACERTADA": hits+=1
            elif st=="NO ACERTADA": misses+=1
            else: pending+=1
        closed=hits+misses
        pct=(100*hits/closed) if closed else None
        rows.append((pct if pct is not None else -1,group,hits,closed,pending))
    rows.sort(reverse=True)
    lines=["🏆 **RANKING DE PREDICCIONES**"]
    for pct,g,h,c,pend in rows:
        score=f"{h}/{c} ({pct:.1f}%)" if c else "sin resultados cerrados"
        lines.append(f"**{g}** → {score} · 🟡 {pend}")
    await i.response.send_message("\n".join(lines)[:1950])

@bot.tree.command(name="historial",description="Últimas correlaciones detectadas")
async def c_history(i:discord.Interaction):
    with connect() as con:
        rows=con.execute("""SELECT * FROM correlation_history
                            ORDER BY detected_at DESC LIMIT 15""").fetchall()
    if not rows:
        await i.response.send_message("📭 Todavía no hay correlaciones automáticas guardadas.")
        return
    lines=["📑 **HISTORIAL DE CORRELACIONES**"]
    for r in rows:
        icon="🚨" if r["match_level"]=="COINCIDENCIA" else "🟡"
        when=datetime.fromisoformat(r["event_time"])
        lines.append(f"{icon} **{r['prediction_code']}** ↔ CSN #{r['event_source_id']} · "
                     f"{when:%d/%m %H:%M} · M{r['event_mag']:.1f} · {r['distance_km']:.1f} km")
    await i.response.send_message("\n".join(lines)[:1950])

@bot.tree.command(name="integridad",description="Comprueba si las predicciones cambiaron después de registrarse")
async def c_integrity(i:discord.Interaction):
    ps=get_real_group()
    ok=changed=unfrozen=0
    changed_codes=[]
    for p in ps:
        st,frozen=integrity_status(p)
        if st=="OK": ok+=1
        elif st=="MODIFICADA":
            changed+=1; changed_codes.append(p["code"])
        else: unfrozen+=1
    msg=(f"🔒 **INTEGRIDAD DE PREDICCIONES**\n"
         f"✅ Sin cambios: **{ok}**\n⚠️ Modificadas tras registro: **{changed}**\n"
         f"❔ No congeladas: **{unfrozen}**")
    if changed_codes: msg+="\n\nModificadas: "+", ".join(changed_codes[:20])
    await i.response.send_message(msg[:1950])

@bot.tree.command(name="backup",description="Crea ahora una copia de seguridad de SQLite")
async def c_backup(i:discord.Interaction):
    try:
        out=backup_database()
        await i.response.send_message(
            f"💾 Backup creado: `{out.name if out else 'sin base de datos'}`\n"
            "También se crea automáticamente una copia cada 24 horas."
        )
    except Exception as exc:
        await i.response.send_message(f"❌ No pude crear el backup: `{exc}`")


init_db()
if __name__=="__main__":
    if not TOKEN:raise SystemExit("Configura DISCORD_TOKEN.")
    bot.run(TOKEN)
