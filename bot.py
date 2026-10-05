import ast
import os, re, math, sqlite3, hashlib, threading, tempfile, resource, json, struct, zlib
import time as pytime
from io import BytesIO
from datetime import datetime, date, time, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

import discord
import requests
from discord import app_commands
from discord.ext import commands, tasks
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak

from csn import fetch_recent_events, fetch_recent_events_debug, fetch_historical_events
from dashboard import run_dashboard
from evaluator import evaluate_real as shared_evaluate_real, best_proximity, real_window_finished as shared_real_window_finished, real_time_ok as shared_real_time_ok, _window_bounds as shared_window_bounds

TOKEN = os.getenv("DISCORD_TOKEN")
EXCEL_PATH = os.getenv("EXCEL_PATH", "Proyecto_Tabla_de_datos_con_coordenadas.xlsx")
DB_PATH = os.getenv("DB_PATH", "bot_sismico.db")
CHECK_MINUTES = int(os.getenv("CSN_CHECK_MINUTES", "5"))
DEFAULT_MARGIN_HOURS = 2
BOT_STARTED_AT = pytime.monotonic()

def parse_datetime(value):
    if isinstance(value, datetime):
        return value
    if isinstance(value, str):
        return datetime.fromisoformat(value)
    raise TypeError(f"fecha/hora inválida: {type(value).__name__}")

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
        CREATE TABLE IF NOT EXISTS csn_feed_sent(
          event_source_id TEXT PRIMARY KEY,
          sent_at TEXT NOT NULL
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
        CREATE TABLE IF NOT EXISTS ranking_snapshots(
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          captured_at TEXT NOT NULL,
          ranking_json TEXT NOT NULL
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

def update_csn(limit=None):
    events,errors=fetch_recent_events(limit)
    new=0
    for e in events:new+=1 if save_csn_event(e) else 0
    return len(events),new,errors

def preds():
    with connect() as con:return con.execute("SELECT * FROM predictions WHERE group_name='PRUEBA' ORDER BY predicted_at").fetchall()
def events():
    with connect() as con:return con.execute("SELECT * FROM observed_events WHERE source='CSN' ORDER BY occurred_at DESC").fetchall()

def evaluate(p,es,hours):
    target=datetime.fromisoformat(p["predicted_at"]);start=target-timedelta(hours=hours);end=target+timedelta(hours=hours)
    lo=p["magnitude"]-p["mag_margin"];hi=p["magnitude"]+p["mag_margin"];rad=radius(p["magnitude"],p["zone_type"])
    matches=[];cands=[]
    for e in es:
        when=parse_datetime(e["occurred_at"]);dist=haversine(p["latitude"],p["longitude"],e["latitude"],e["longitude"])
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
    """Return dates and optional absolute start/end times (Chile local).

    A single timestamp is a point, not an invented multi-hour window.
    UTC timestamps are converted using the timezone for the event date.
    """
    from datetime import timezone
    if isinstance(raw, datetime):
        raw = raw.strftime("%d/%m/%Y %H:%M")
    if isinstance(raw, date):
        return raw, raw, None, None
    s = str(raw or "").strip().lower().replace("–", "-").replace("—", "-")
    if not s: raise ValueError("fecha vacía")
    def make(d, m, y=None): return date(int(y or year_default), int(m), int(d))
    def check(a,b,t1=None,t2=None):
        if b < a: raise ValueError("rango de fechas invertido")
        return a,b,t1,t2
    # Dos fechas completas con horas y guion (MISTRAL AI).
    # Se comprueba ANTES de la fecha individual para no truncar la ventana.
    m = re.search(
        r"(\d{1,2})/(\d{1,2})/(20\d{2})\s+(\d{1,2}:\d{2})\s*-\s*"
        r"(\d{1,2})/(\d{1,2})/(20\d{2})\s+(\d{1,2}:\d{2})",
        s,
    )
    if m:
        d1,mo1,y1,t1,d2,mo2,y2,t2 = m.groups()
        start_date, end_date = make(d1,mo1,y1), make(d2,mo2,y2)
        start_dt = datetime.combine(start_date, datetime.strptime(t1, "%H:%M").time())
        end_dt = datetime.combine(end_date, datetime.strptime(t2, "%H:%M").time())
        if end_dt < start_dt:
            raise ValueError("rango de fecha/hora invertido")
        return check(start_date, end_date, t1, t2)
    # Explicit start/end dates, each with its own clock time (META AI).
    m = re.search(r"(\d{1,2})/(\d{1,2})(?:/(20\d{2}))?\s+(\d{1,2}:\d{2})\s+a\s+(\d{1,2})/(\d{1,2})(?:/(20\d{2}))?\s+(\d{1,2}:\d{2})",s)
    if m:
        d1,mo1,y1,t1,d2,mo2,y2,t2=m.groups()
        return check(make(d1,mo1,y1),make(d2,mo2,y2),t1,t2)
    # One full date with UTC time and stated duration (GEMINI).
    m=re.search(r"(\d{1,2})/(\d{1,2})/(20\d{2})\s+(\d{1,2}:\d{2})\s+utc\s*\(\s*(\d+)\s*h\s*\)",s)
    if m:
        d,mo,y,clock,hours=m.groups()
        start=datetime.combine(make(d,mo,y),datetime.strptime(clock,"%H:%M").time(),tzinfo=timezone.utc).astimezone(ZoneInfo("America/Santiago"))
        end=(datetime.combine(make(d,mo,y),datetime.strptime(clock,"%H:%M").time(),tzinfo=timezone.utc)+timedelta(hours=int(hours))).astimezone(ZoneInfo("America/Santiago"))
        return check(start.date(),end.date(),start.strftime("%H:%M"),end.strftime("%H:%M"))
    # 23-24/09/2026; 10 - 13/09/2026.
    m=re.search(r"(\d{1,2})\s*-\s*(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(20\d{2})",s)
    if m:
        d1,d2,mo,y=m.groups();return check(make(d1,mo,y),make(d2,mo,y))
    # Full date and optional time or clock interval (PERPLEXITY / MISTRAL).
    m=re.search(r"(\d{1,2})/(\d{1,2})/(20\d{2})(?:\s*,?\s*(\d{1,2}:\d{2})(?:\s*-\s*(\d{1,2}:\d{2}))?)?",s)
    if m:
        d,mo,y,t1,t2=m.groups(); day=make(d,mo,y)
        return check(day,day,t1,t2 or t1)
    # 11-13 de septiembre / 21-24 de septiembre de 2026.
    m=re.search(r"(\d{1,2})\s*-\s*(\d{1,2})(?:\s+de)?\s+([a-záéíóú]+)(?:\s+de\s+(20\d{2}))?",s)
    if m:
        d1,d2,mon,y=m.groups();mo=MONTHS_ES.get(mon)
        if not mo:raise ValueError(f"mes no reconocido: {mon}")
        return check(make(d1,mo,y),make(d2,mo,y))
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

def _window_bounds(p):
    return shared_window_bounds(p)

def real_time_ok(event_dt, p):
    return shared_real_time_ok(event_dt, p)

def real_window_finished(p):
    return shared_real_window_finished(p)

def prediction_alert_active(p, now=None):
    return not shared_real_window_finished(p, now)

def evaluate_real(p, es):
    # Fuente única de verdad compartida con el dashboard web.
    return shared_evaluate_real(p, es)

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


def storage_report():
    src = Path(DB_PATH)
    backup_dir = src.parent / "backups"
    backups = sorted(backup_dir.glob("bot_sismico_*.db")) if backup_dir.is_dir() else []
    db_files = [src, Path(str(src)+"-wal"), Path(str(src)+"-shm")]
    return {"db_bytes": sum(f.stat().st_size for f in db_files if f.is_file()),
            "backup_bytes": sum(f.stat().st_size for f in backups),
            "backup_count": len(backups), "backup_dir": str(backup_dir)}

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
    when=parse_datetime(e["occurred_at"])
    rad=radius((p["mag_min"]+p["mag_max"])/2,p["zone_type"])
    dist=haversine(p["latitude"],p["longitude"],e["latitude"],e["longitude"])
    return {
        "time": real_time_ok(when,p),
        "mag": p["mag_min"] <= e["magnitude"] <= p["mag_max"],
        "geo": dist <= rad,
        "dist": dist,
        "radius": rad,
    }


def csn_feed_was_sent(event_id):
    with connect() as con:
        return con.execute(
            "SELECT 1 FROM csn_feed_sent WHERE event_source_id=?",
            (str(event_id),)
        ).fetchone() is not None

def mark_csn_feed_sent(event_id):
    with connect() as con:
        con.execute(
            "INSERT OR IGNORE INTO csn_feed_sent(event_source_id,sent_at) VALUES(?,?)",
            (str(event_id), datetime.now().isoformat())
        )

async def send_csn_feed_for_new_events(new_events):
    """Publica TODOS los sismos nuevos del CSN en un canal independiente."""
    channel_id=get_setting("csn_feed_channel_id")
    if not channel_id or not new_events:
        return 0
    try:
        channel=bot.get_channel(int(channel_id)) or await bot.fetch_channel(int(channel_id))
    except Exception as exc:
        print("No pude abrir el canal feed CSN:",exc)
        return 0

    sent=0
    for e in sorted(new_events, key=lambda x: parse_datetime(x["occurred_at"])):
        if csn_feed_was_sent(e["source_id"]):
            continue
        when=parse_datetime(e["occurred_at"])
        embed=discord.Embed(
            title=f"🌎 Nuevo sismo CSN · M{e['magnitude']:.1f}",
            description=e["place"] or "Sin referencia geográfica",
            color=discord.Color.blurple()
        )
        embed.add_field(name="📅 Fecha",value=f"{when:%d/%m/%Y}",inline=True)
        embed.add_field(name="🕒 Hora local",value=f"{when:%H:%M:%S}",inline=True)
        embed.add_field(name="📈 Magnitud",value=f"M{e['magnitude']:.1f}",inline=True)
        embed.add_field(name="🌐 Coordenadas",value=f"{e['latitude']:.4f}, {e['longitude']:.4f}",inline=True)
        embed.add_field(
            name="⬇️ Profundidad",
            value=(f"{e['depth_km']:.1f} km" if e["depth_km"] is not None else "Sin dato"),
            inline=True
        )
        event_url = e.get("source_url") or e.get("url")
        if event_url:
            embed.add_field(name="🔗 Informe oficial",value=f"[Abrir en CSN]({event_url})",inline=False)
        embed.set_footer(text=f"CSN ID: {e['source_id']} · Feed automático, sin @everyone")
        try:
            await channel.send(
                embed=embed,
                allowed_mentions=discord.AllowedMentions.none()
            )
            mark_csn_feed_sent(e["source_id"])
            sent+=1
        except Exception as exc:
            print("Error enviando sismo al feed CSN:",exc)
    return sent

async def send_alerts_for_new_events(new_events):
    channel_id=get_setting("alert_channel_id")
    if not channel_id or not new_events:
        return 0

    try:
        channel=bot.get_channel(int(channel_id)) or await bot.fetch_channel(int(channel_id))
    except Exception as exc:
        print("No pude abrir el canal de alertas:",exc)
        return 0

    predictions=[p for p in get_real_group() if prediction_alert_active(p)]
    sent=0

    for e in new_events:
        when=parse_datetime(e["occurred_at"])
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

def _volume_inventory(root=Path("/data"), limit=12):
    """Inventario solo lectura del volumen; no sigue enlaces simbólicos."""
    root = Path(root)
    if not root.is_dir():
        return {"error": f"No existe o no es accesible: {root}"}
    files = []
    dirs = []
    errors = []
    total = 0
    count = 0
    # os.walk no sigue directorios enlazados por defecto. No se inspeccionan
    # rutas externas al volumen ni se muestran contenidos de archivos.
    for base, subdirs, filenames in os.walk(root, followlinks=False, onerror=lambda exc: errors.append(str(exc))):
        subdirs[:] = [name for name in subdirs if not (Path(base) / name).is_symlink()]
        folder_bytes = 0
        for name in filenames:
            item = Path(base) / name
            try:
                if item.is_symlink() or not item.is_file():
                    continue
                size = item.stat().st_size
            except OSError as exc:
                errors.append(f"{item}: {exc}")
                continue
            total += size
            folder_bytes += size
            count += 1
            files.append((size, str(item.relative_to(root))))
        dirs.append((folder_bytes, str(Path(base).relative_to(root))))
    files.sort(key=lambda row: row[0], reverse=True)
    dirs.sort(key=lambda row: row[0], reverse=True)
    return {"total": total, "count": count, "files": files[:limit],
            "dirs": dirs[:limit], "errors": errors[:3]}



def _volume_filesystem_diagnostic(root=Path("/data")):
    """Diagnóstico de solo lectura: bloques, statvfs y ficheros borrados abiertos."""
    root = Path(root)
    if not root.is_dir():
        return {"error": f"No existe o no es accesible: {root}"}
    result = {"errors": []}
    try:
        st = os.statvfs(root)
        result["capacity"] = st.f_blocks * st.f_frsize
        result["used"] = (st.f_blocks - st.f_bfree) * st.f_frsize
        result["available"] = st.f_bavail * st.f_frsize
    except OSError as exc:
        result["errors"].append(f"statvfs: {type(exc).__name__}: {exc}")
    # st_size es el tamaño lógico; st_blocks contabiliza bloques realmente asignados.
    allocated = 0
    logical = 0
    nfiles = 0
    for base, subdirs, names in os.walk(root, followlinks=False,
                                        onerror=lambda exc: result["errors"].append(str(exc))):
        subdirs[:] = [n for n in subdirs if not (Path(base) / n).is_symlink()]
        for name in names:
            path = Path(base) / name
            try:
                if path.is_symlink() or not path.is_file():
                    continue
                st = path.stat()
                logical += st.st_size
                allocated += getattr(st, "st_blocks", 0) * 512
                nfiles += 1
            except OSError as exc:
                result["errors"].append(f"stat: {type(exc).__name__}: {exc}")
    result.update(logical=logical, allocated=allocated, files=nfiles)
    # Solo examina descriptores abiertos del proceso del bot: no inspecciona
    # otros procesos ni lee contenidos. Puede no detectar borrados abiertos
    # por procesos distintos o ficheros que ya no figuran en /proc.
    deleted_bytes = 0
    deleted_count = 0
    fd_dir = Path("/proc/self/fd")
    if fd_dir.is_dir():
        try:
            for fd in fd_dir.iterdir():
                try:
                    target = os.readlink(fd)
                    if not target.endswith(" (deleted)"):
                        continue
                    # /proc/self/fd/<n> puede señalar sockets, pipes, etc.
                    st = fd.stat()
                    if not __import__('stat').S_ISREG(st.st_mode):
                        continue
                    deleted_count += 1
                    deleted_bytes += getattr(st, "st_blocks", 0) * 512
                except (OSError, ValueError):
                    continue
        except OSError as exc:
            result["errors"].append(f"/proc/self/fd: {type(exc).__name__}: {exc}")
    else:
        result["errors"].append("No se puede consultar /proc/self/fd")
    result["deleted_count"] = deleted_count
    result["deleted_bytes"] = deleted_bytes
    return result


@bot.tree.command(name="almacenamiento",description="Diagnostica SQLite, backups y archivos grandes de /data sin borrar nada")
@app_commands.guild_only()
@app_commands.default_permissions(administrator=True)
async def c_almacenamiento(i:discord.Interaction):
    if not i.user.guild_permissions.administrator:
        await i.response.send_message("Solo administradores pueden consultar este informe.",ephemeral=True)
        return
    await i.response.defer(ephemeral=True)
    try:
        def report():
            size = storage_report()
            with connect() as con:
                page_size = con.execute('PRAGMA page_size').fetchone()[0]
                free = con.execute('PRAGMA freelist_count').fetchone()[0]
            return size, free * page_size, _volume_inventory(), _volume_filesystem_diagnostic()
        size, free_bytes, inventory, fs = await __import__('asyncio').to_thread(report)
        lines = ["💾 **Almacenamiento del bot (solo lectura)**",
                 f"SQLite (+ WAL/SHM): **{_human_bytes(size['db_bytes'])}**",
                 f"Backups: **{size['backup_count']}** · **{_human_bytes(size['backup_bytes'])}**",
                 f"Reutilizable dentro de SQLite: **{_human_bytes(free_bytes)}**"]
        if 'error' in inventory:
            lines.append(f"⚠️ {inventory['error']}")
        else:
            lines += [f"📁 **/data: {_human_bytes(inventory['total'])} en {inventory['count']} archivos**",
                      "**Archivos más grandes:**"]
            lines += [f"`/{name}` — **{_human_bytes(n)}**" for n,name in inventory['files'][:10]]
            lines.append("**Carpetas con archivos más grandes (sin subcarpetas):**")
            lines += [f"`/{name}` — **{_human_bytes(n)}**" for n,name in inventory['dirs'][:5] if n]
            if inventory['errors']:
                lines.append(f"⚠️ {len(inventory['errors'])} errores de lectura; total posiblemente incompleto.")
            lines.append("El total suma archivos visibles; puede diferir de la métrica de Railway.")
        lines.append("**Diagnóstico del sistema de archivos (/data):**")
        if 'error' in fs:
            lines.append(f"⚠️ {fs['error']}")
        else:
            if 'used' in fs:
                lines.append(f"Uso según sistema (statvfs): **{_human_bytes(fs['used'])}** · Capacidad: **{_human_bytes(fs['capacity'])}** · Disponible: **{_human_bytes(fs['available'])}**")
            lines.append(f"Archivos visibles: **{fs['files']}** · Tamaño lógico: **{_human_bytes(fs['logical'])}** · Bloques asignados: **{_human_bytes(fs['allocated'])}**")
            lines.append(f"Archivos borrados aún abiertos por este bot: **{fs['deleted_count']}** · Bloques: **{_human_bytes(fs['deleted_bytes'])}**")
            if 'used' in fs:
                difference = max(0, fs['used'] - fs['allocated'])
                lines.append(f"Diferencia sistema − bloques visibles: **{_human_bytes(difference)}** (no identifica por sí sola la causa)")
            if fs['errors']:
                lines.append(f"⚠️ {len(fs['errors'])} avisos de inspección; el diagnóstico puede ser incompleto.")
            lines.append("El conteo de archivos borrados solo revisa este proceso. La métrica de Railway puede diferir por el montaje, otros procesos o retrasos de actualización.")
        message = "\n".join(lines)
        # Dividir por líneas para evitar el límite de 2000 caracteres de Discord.
        parts=[]; current=""
        for line in message.splitlines():
            if len(current)+len(line)+1 > 1900:
                parts.append(current);current=""
            current += ("\n" if current else "") + line[:1800]
        if current: parts.append(current)
        for part in parts:
            await i.followup.send(part,ephemeral=True)
    except Exception as exc:
        await i.followup.send(f"❌ Error: `{type(exc).__name__}: {str(exc)[:300]}`",ephemeral=True)

@tasks.loop(minutes=CHECK_MINUTES)
async def csn_loop():
    try:
        fetched,errs=fetch_recent_events()
        newly_saved=[]
        for e in fetched:
            if save_csn_event(e):
                newly_saved.append(e)

        if newly_saved:
            print(f"CSN: {len(newly_saved)} eventos nuevos ({len(fetched)} leídos)")
            feed_sent=await send_csn_feed_for_new_events(newly_saved)
            alerts_sent=await send_alerts_for_new_events(newly_saved)
            if feed_sent or alerts_sent:
                print(f"Discord: feed CSN={feed_sent}, alertas correlación={alerts_sent}")
            save_ranking_snapshot()

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

@bot.tree.command(name="actualizar_csn",description="Fuerza una sincronización fresca con el catálogo oficial CSN")
async def c_update_csn(i:discord.Interaction):
    await i.response.defer(thinking=True, ephemeral=True)
    try:
        fetched, errs = fetch_recent_events(limit=None)
        newly_saved=[]
        for e in fetched:
            if save_csn_event(e):
                newly_saved.append(e)
        feed_sent=await send_csn_feed_for_new_events(newly_saved) if newly_saved else 0
        alerts_sent=await send_alerts_for_new_events(newly_saved) if newly_saved else 0
        newest=fetched[0] if fetched else None
        newest_txt="sin eventos"
        if newest:
            dt=parse_datetime(newest["occurred_at"])
            newest_txt=f"{dt:%d/%m/%Y %H:%M:%S} · M{newest['magnitude']:.1f} · {newest['place']}"
        await i.followup.send(
            f"🌐 **CSN actualizado**\n"
            f"Informes frescos leídos: **{len(fetched)}**\n"
            f"Nuevos guardados: **{len(newly_saved)}**\n"
            f"Publicados en #sismos: **{feed_sent}**\n"
            f"Alertas de predicción: **{alerts_sent}**\n"
            f"Último evento leído: **{newest_txt}**" +
            (f"\n⚠️ Avisos CSN: **{len(errs)}** · `{errs[0][:250]}`" if errs else ""),
            ephemeral=True
        )
    except Exception as exc:
        await i.followup.send(f"❌ Error actualizando CSN: `{type(exc).__name__}: {str(exc)[:900]}`",ephemeral=True)


@bot.tree.command(name="analizar",description="Evalúa PRUEBA usando únicamente eventos reales CSN guardados")
async def c_an(i:discord.Interaction,margen_horas:app_commands.Range[int,1,2]=2):
    ps,es=preds(),events()
    if not ps:await i.response.send_message("Usa `/importar_pruebas`.");return
    if not es:await i.response.send_message("No hay eventos CSN guardados. Usa `/actualizar_csn`.");return
    lines=[f"📊 **CSN REAL · ±{margen_horas} h**"]
    for p in ps:
        st,b,rad,*_=evaluate(p,es,margen_horas)
        icon={"ACERTADA":"🟢","PENDIENTE":"🟡","NO ACERTADA":"🔴","CASI ACERTADA":"🟠"}[st]
        lines.append(f"{icon} **{p['code']} — {st}** · radio {rad} km")
    await i.response.send_message("\n".join(lines)[:1950])

@bot.tree.command(name="ver",description="Detalle de una predicción frente a eventos reales del CSN")
async def c_ver(i:discord.Interaction,codigo:str,margen_horas:app_commands.Range[int,1,2]=2):
    with connect() as con:p=con.execute("SELECT * FROM predictions WHERE UPPER(code)=UPPER(?)",(codigo,)).fetchone()
    if not p:await i.response.send_message("❌ No encontrada.");return
    st,b,rad,cands,start,end,lo,hi=evaluate(p,events(),margen_horas)
    icon={"ACERTADA":"🟢","PENDIENTE":"🟡","NO ACERTADA":"🔴","CASI ACERTADA":"🟠"}[st]
    dt=datetime.fromisoformat(p["predicted_at"])
    text=(f"{icon} **{p['code']} — {st}**\nPredicción: {dt:%d/%m/%Y %H:%M}\n📍 {p['place']} "
          f"({p['latitude']:.4f}, {p['longitude']:.4f})\n📈 M{p['magnitude']:.1f} ±{p['mag_margin']:.1f} "
          f"→ M{lo:.1f}–M{hi:.1f}\n📏 Radio: {rad} km\n⏱️ Ventana: {start:%d/%m %H:%M}–{end:%d/%m %H:%M}")
    if b:
        e=b["e"];when=parse_datetime(e["occurred_at"])
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


@bot.tree.command(name="csn_diagnostico",description="Diagnostica catálogos CSN Chile/UTC y SQLite")
async def c_csn_diag(i:discord.Interaction):
    await i.response.defer(thinking=True, ephemeral=True)
    try:
        d=fetch_recent_events_debug()
        fetched=d["events"]
        with connect() as con:
            db_ids={r[0] for r in con.execute("SELECT source_id FROM observed_events").fetchall()}
        missing=[e for e in fetched if e["source_id"] not in db_ids]
        db_events=events()
        newest_web=d["newest"]
        newest_db=db_events[0] if db_events else None
        same_latest=False
        if newest_web and newest_db:
            same_latest=(parse_datetime(newest_web["occurred_at"]) == parse_datetime(newest_db["occurred_at"]))
        elif newest_web is None and newest_db is None:
            same_latest=True
        state="🟢 Sincronizado" if not missing and same_latest else "🔴 DESACTUALIZADO"

        def fmt(e):
            if not e: return "—"
            dt=parse_datetime(e["occurred_at"])
            return f"{dt:%d/%m/%Y %H:%M:%S} · M{e['magnitude']:.1f}\n{e['place']}"

        cat="\n".join(
            f"`{day}` → **{d['counts'].get(day,0)}**"
            for day in sorted(d["counts"])
        )
        embed=discord.Embed(title="🛰️ Diagnóstico CSN multi-día",color=(discord.Color.green() if state.startswith("🟢") else discord.Color.red()))
        embed.add_field(name="📆 Catálogos consultados",value=cat or "—",inline=False)
        embed.add_field(name="Eventos únicos",value=str(d["total_count"]),inline=True)
        embed.add_field(name="Faltantes SQLite",value=str(len(missing)),inline=True)
        embed.add_field(name="Estado",value=state,inline=True)
        embed.add_field(name="🌐 Último CSN leído",value=fmt(newest_web),inline=False)
        embed.add_field(name="💾 Último guardado",value=fmt(newest_db),inline=False)
        if missing:
            sample="\n".join(f"• {parse_datetime(e['occurred_at']):%d/%m %H:%M:%S} M{e['magnitude']:.1f} {e['place']}" for e in missing[:8])
            embed.add_field(name="⚠️ Faltantes",value=sample[:1000],inline=False)
        if d["errors"]:
            embed.add_field(name="Errores de consulta",value="\n".join(d["errors"][:6])[:1000],inline=False)
        await i.followup.send(embed=embed,ephemeral=True)
    except Exception as exc:
        await i.followup.send(f"❌ Diagnóstico falló: `{type(exc).__name__}: {str(exc)[:900]}`",ephemeral=True)


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
        when = parse_datetime(e["occurred_at"])
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

    counts={"ACERTADA":0,"CASI ACERTADA":0,"NO ACERTADA":0,"PENDIENTE":0}
    lines=[f"📊 **{grupo.upper()} · CSN REAL**"]
    icons={"ACERTADA":"🟢","CASI ACERTADA":"🟠","NO ACERTADA":"🔴","PENDIENTE":"🟡"}
    for p in ps:
        st,b,rad,_=evaluate_real(p,es)
        counts[st]+=1
        lines.append(f"{icons[st]} **{p['code']} — {st}** · M{p['mag_min']:.1f}–{p['mag_max']:.1f} · {rad} km")
    closed=len(ps)-counts["PENDIENTE"]
    lines.append(f"\n🟢 Acertadas: **{counts['ACERTADA']}** · 🟠 Casi acertadas: **{counts['CASI ACERTADA']}** · 🔴 No acertadas: **{counts['NO ACERTADA']}** · 🟡 Pendientes: **{counts['PENDIENTE']}**")
    if closed:
        lines.append(f"Aciertos sobre resultados cerrados: **{counts['ACERTADA']}/{closed} ({counts['ACERTADA']/closed*100:.1f}%)**")
    # Discord limita los mensajes a 2000 caracteres; enviar todas las predicciones.
    await i.response.defer()
    chunk=""
    for line in lines:
        if len(chunk)+len(line)+1>1900:
            await i.followup.send(chunk)
            chunk=""
        chunk+=("\n" if chunk else "")+line
    if chunk:
        await i.followup.send(chunk)

def build_results_excel(predictions, observed):
    """Exporta una instantánea; no modifica las predicciones ni la base de datos."""
    wb = Workbook()
    summary = wb.active
    summary.title = "RESUMEN"
    columns = ["Grupo", "Código", "Estado", "Fecha inicio", "Fecha fin", "Hora inicio", "Hora fin",
               "Lugar predicho", "Magnitud mín.", "Magnitud máx.", "Latitud pred.", "Longitud pred.",
               "Radio (km)", "ID sismo CSN", "Fecha sismo", "Lugar sismo", "Magnitud sismo",
               "Distancia (km)", "Tiempo válido", "Ubicación válida", "Magnitud válida", "Enlace oficial CSN"]
    categories = ("ACERTADA", "CASI ACERTADA", "NO ACERTADA", "PENDIENTE")
    names = {"ACERTADA": "ACERTADAS", "CASI ACERTADA": "CASI ACERTADAS",
             "NO ACERTADA": "NO ACERTADAS", "PENDIENTE": "PENDIENTES"}
    tabs = {status: wb.create_sheet(names[status]) for status in categories}
    for ws in tabs.values():
        ws.append(columns)
    counts = {}
    for p in predictions:
        status, match, rad, _ = evaluate_real(p, observed)
        group = p["group_name"]
        if group not in counts:
            counts[group] = {key: 0 for key in categories}
        counts[group][status] += 1
        e = match["e"] if match else None
        row = [group, p["code"], status, p["date_start"], p["date_end"],
               p["daily_time_start"], p["daily_time_end"], p["place"],
               p["mag_min"], p["mag_max"], p["latitude"], p["longitude"], rad,
               e["source_id"] if e else None, str(e["occurred_at"]) if e else None,
               e["place"] if e else None, e["magnitude"] if e else None,
               round(match["dist"], 2) if match else None,
               ("SÍ" if match["t"] else "NO") if match is not None else None,
               ("SÍ" if match["g"] else "NO") if match is not None else None,
               ("SÍ" if match["m"] else "NO") if match is not None else None,
               e["source_url"] if e else None]
        tabs[status].append(row)
    summary.append(["Grupo", "Acertadas", "Casi acertadas", "No acertadas", "Pendientes", "Total"])
    for group in sorted(counts):
        c = counts[group]
        summary.append([group, c["ACERTADA"], c["CASI ACERTADA"], c["NO ACERTADA"],
                        c["PENDIENTE"], sum(c.values())])
    summary.append(["TOTAL"] + [sum(c[key] for c in counts.values()) for key in categories] + [len(predictions)])
    summary.append([])
    summary.append(["Clasificación recalculada al exportar con los eventos CSN guardados en el bot."])
    summary.append(["Una predicción pendiente no se considera no acertada hasta cerrar su ventana."])
    for ws in [summary, *tabs.values()]:
        ws.freeze_panes = "D2" if ws is not summary else "B2"
        ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{max(1, ws.max_row if ws is not summary else len(counts)+2)}"
        for cell in ws[1]:
            cell.font = Font(color="FFFFFF", bold=True)
            cell.fill = PatternFill("solid", fgColor="17365D")
            cell.alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[1].height = 32
        for idx in range(1, ws.max_column+1):
            width = 17
            if ws is not summary:
                width = {2:20, 3:19, 8:32, 16:32, 22:47}.get(idx, 17)
            ws.column_dimensions[get_column_letter(idx)].width = width
        if ws is not summary:
            for row in ws.iter_rows(min_row=2):
                if row[21].value:
                    row[21].hyperlink = row[21].value
                    row[21].font = Font(color="0563C1", underline="single")
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output, counts


@bot.tree.command(name="exportar_excel", description="Descarga un Excel de acertadas, casi acertadas, no acertadas y pendientes")
@app_commands.describe(grupo="Opcional: nombre de la IA; vacío para incluir todos los grupos")
async def c_export_excel(i:discord.Interaction, grupo:str=None):
    await i.response.defer(thinking=True)
    try:
        ps = get_real_group(grupo) if grupo else get_real_group()
        if not ps:
            await i.followup.send("No hay predicciones importadas para ese grupo." if grupo else "No hay predicciones importadas. Usa `/importar_todos`.")
            return
        output, counts = await __import__('asyncio').to_thread(build_results_excel, ps, events())
        if output.getbuffer().nbytes > 8 * 1024 * 1024:
            await i.followup.send("El Excel supera 8 MB. Prueba exportar un grupo específico con `/exportar_excel grupo:`.")
            return
        filename = "resultados_csn_" + (re.sub(r"[^a-zA-Z0-9_-]", "_", grupo.upper()) if grupo else "TODOS") + ".xlsx"
        await i.followup.send("📊 Excel generado con las categorías actuales y un resumen por grupo.",
                              file=discord.File(output, filename=filename))
    except Exception as exc:
        await i.followup.send(f"❌ No se pudo generar el Excel: `{type(exc).__name__}: {str(exc)[:300]}`")


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
        hits=almost=misses=pending=0
        for p in gps:
            st,_,_,_=evaluate_real(p,es)
            if st=="ACERTADA": hits+=1
            elif st=="CASI ACERTADA": almost+=1
            elif st=="NO ACERTADA": misses+=1
            else: pending+=1
        closed=hits+almost+misses
        score=f"{hits}/{closed} ({hits/closed*100:.1f}%)" if closed else "sin resultados cerrados"
        lines.append(f"**{group}** → {score} · 🟠 {almost} casi · 🔴 {misses} no acertadas · 🟡 {pending} pendientes")
    await i.response.send_message("\n".join(lines)[:1950])

def _real_candidate_score(p, x, rad):
    e=x["e"]
    center=(p["mag_min"]+p["mag_max"])/2
    distance_penalty=x["dist"]/max(rad,1)
    mag_half=max((p["mag_max"]-p["mag_min"])/2,0.25)
    mag_penalty=abs(e["magnitude"]-center)/mag_half
    time_penalty=0 if x["t"] else 5
    return time_penalty + distance_penalty + mag_penalty

def _real_browser_data(p):
    es=events()
    st,b,rad,cands=evaluate_real(p,es)
    # Mantener primero los eventos temporalmente compatibles; dentro de cada grupo,
    # ordenar por proximidad combinada para que el navegador empiece por lo relevante.
    cands=sorted(cands,key=lambda x:(0 if x["t"] else 1,_real_candidate_score(p,x,rad)))
    initial=0
    if b is not None and cands:
        source_id=b["e"]["source_id"]
        initial=next((n for n,x in enumerate(cands) if x["e"]["source_id"]==source_id),0)
    return st,rad,cands,initial

def build_real_browser_embed(p, st, rad, cands, index=0):
    colors={"ACERTADA":discord.Color.green(),"CASI ACERTADA":discord.Color.gold(),
            "NO ACERTADA":discord.Color.red(),"PENDIENTE":discord.Color.blurple()}
    icons={"ACERTADA":"✅","CASI ACERTADA":"🟠","NO ACERTADA":"❌","PENDIENTE":"🟡"}
    e=ui_embed(f"{icons.get(st,'🎯')} {p['code']} · {st}",
               f"**{p['group_name']}** · Predicción #{p['prediction_no']}",colors.get(st,discord.Color.blurple()))
    timeband=f"\nHorario diario: **{p['daily_time_start']}–{p['daily_time_end']}**" if p["daily_time_start"] else ""
    e.add_field(name="🎯 Predicción",value=(
        f"📅 **{p['date_start']} → {p['date_end']}**{timeband}\n"
        f"📍 {p['place']}\n📈 **M{p['mag_min']:.1f}–M{p['mag_max']:.1f}** · 📏 **{rad} km**"),inline=False)
    if not cands:
        e.add_field(name="🌎 Comparación CSN",value="No hay sismos CSN guardados para comparar.",inline=False)
        e.set_footer(text="Sismología Lab · /ver_real")
        return e
    index=max(0,min(index,len(cands)-1)); x=cands[index]; ev=x["e"]
    try: when=parse_datetime(ev["occurred_at"]).strftime("%d/%m/%Y %H:%M:%S")
    except Exception: when=str(ev["occurred_at"])
    checks=f"🕒 {'✅' if x['t'] else '❌'}  ·  📈 {'✅' if x['m'] else '❌'}  ·  📍 {'✅' if x['g'] else '❌'}"
    e.add_field(name=f"🌎 Sismo {index+1}/{len(cands)} · M{ev['magnitude']:.1f}",value=(
        f"📅 {when}\n📍 {ev['place'] or 'Sin referencia geográfica'}\n"
        f"🌐 `{ev['latitude']:.4f}, {ev['longitude']:.4f}`\n📏 **{x['dist']:.1f} km** de la predicción\n\n{checks}"),inline=False)
    if ev["source_url"]:
        e.add_field(name="🔗 CSN",value=f"[Abrir informe oficial]({ev['source_url']})",inline=False)
    e.set_footer(text="◀️/▶️ recorre los sismos · ⭐ vuelve al más relevante")
    return e

class RealPredictionBrowser(discord.ui.View):
    def __init__(self,p,st,rad,cands,index=0):
        super().__init__(timeout=900)
        self.p=p; self.st=st; self.rad=rad; self.cands=cands; self.index=index
        self._sync()
    def _sync(self):
        disabled=len(self.cands)<=1
        self.previous.disabled=disabled
        self.next.disabled=disabled
        self.best.disabled=not self.cands
    async def _render(self,interaction):
        self._sync()
        await interaction.response.edit_message(embed=build_real_browser_embed(self.p,self.st,self.rad,self.cands,self.index),view=self)
    @discord.ui.button(label="Anterior",emoji="◀️",style=discord.ButtonStyle.secondary,row=0)
    async def previous(self,interaction,button):
        if self.cands:self.index=(self.index-1)%len(self.cands)
        await self._render(interaction)
    @discord.ui.button(label="Siguiente",emoji="▶️",style=discord.ButtonStyle.primary,row=0)
    async def next(self,interaction,button):
        if self.cands:self.index=(self.index+1)%len(self.cands)
        await self._render(interaction)
    @discord.ui.button(label="Más relevante",emoji="⭐",style=discord.ButtonStyle.success,row=0)
    async def best(self,interaction,button):
        if self.cands:self.index=min(range(len(self.cands)),key=lambda n:_real_candidate_score(self.p,self.cands[n],self.rad))
        await self._render(interaction)
    @discord.ui.button(label="Cerrar",emoji="✖️",style=discord.ButtonStyle.secondary,row=1)
    async def close(self,interaction,button):
        for child in self.children: child.disabled=True
        await interaction.response.edit_message(view=self)

@bot.tree.command(name="ver_real",description="Explora una predicción real y sus sismos CSN en un menú interactivo")
async def c_ver_real(i:discord.Interaction,codigo:str):
    with connect() as con:
        p=con.execute("SELECT * FROM real_predictions WHERE UPPER(code)=UPPER(?)",(codigo,)).fetchone()
    if not p:
        await i.response.send_message("❌ Predicción real no encontrada.",ephemeral=True)
        return
    st,rad,cands,initial=_real_browser_data(p)
    await i.response.send_message(embed=build_real_browser_embed(p,st,rad,cands,initial),
                                  view=RealPredictionBrowser(p,st,rad,cands,initial),ephemeral=True)




@bot.tree.command(name="canal_sismos",description="Usa este canal para publicar TODOS los sismos nuevos del CSN")
async def c_csn_feed_channel(i:discord.Interaction):
    set_setting("csn_feed_channel_id", i.channel_id)
    await i.response.send_message(
        f"🌎 Feed CSN activado en <#{i.channel_id}>.\n"
        f"Publicaré aquí **todos los sismos nuevos detectados por el monitor** cada {CHECK_MINUTES} minutos.\n"
        "Este feed no usa `@everyone`; las correlaciones importantes siguen yendo al canal de alertas."
    )

@bot.tree.command(name="estado_canal_sismos",description="Muestra dónde está configurado el feed de todos los sismos CSN")
async def c_csn_feed_state(i:discord.Interaction):
    channel_id=get_setting("csn_feed_channel_id")
    if channel_id:
        await i.response.send_message(
            f"🌎 Canal de todos los sismos: <#{channel_id}>\n"
            f"📡 Consulta CSN: cada **{CHECK_MINUTES} minutos**.\n"
            "🔕 Sin `@everyone`."
        )
    else:
        await i.response.send_message(
            "🌎 No hay canal de feed CSN configurado. Ejecuta `/canal_sismos` dentro del canal que quieras usar."
        )

@bot.tree.command(name="desactivar_canal_sismos",description="Desactiva el feed de todos los sismos del CSN")
async def c_csn_feed_off(i:discord.Interaction):
    with connect() as con:
        con.execute("DELETE FROM bot_settings WHERE key='csn_feed_channel_id'")
    await i.response.send_message("🔕 Feed de todos los sismos CSN desactivado.")

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



@bot.tree.command(name="ranking",description="Ranking por aciertos, casi acertados y no acertados")
async def c_ranking(i:discord.Interaction):
    await i.response.send_message(embed=build_ranking_embed())

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


# ---------- V1.5.2: INFORMES PDF ----------

def _pdf_safe(value):
    if value is None:
        return "-"
    return str(value).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

def build_experiment_pdf(output_path, group_name=None):
    """Genera un informe PDF usando el mismo evaluador oficial que Discord/dashboard."""
    ps=get_real_group(group_name)
    es=events()
    styles=getSampleStyleSheet()
    title=ParagraphStyle("ReportTitle", parent=styles["Title"], alignment=TA_CENTER, fontSize=20, leading=24, spaceAfter=8)
    sub=ParagraphStyle("ReportSub", parent=styles["Normal"], alignment=TA_CENTER, fontSize=9, leading=12, textColor=colors.HexColor("#555555"), spaceAfter=14)
    h2=ParagraphStyle("ReportH2", parent=styles["Heading2"], fontSize=13, leading=16, spaceBefore=8, spaceAfter=7)
    body=ParagraphStyle("ReportBody", parent=styles["BodyText"], fontSize=8.5, leading=11)
    small=ParagraphStyle("ReportSmall", parent=styles["BodyText"], fontSize=7.5, leading=9)

    doc=SimpleDocTemplate(str(output_path), pagesize=landscape(A4), rightMargin=12*mm, leftMargin=12*mm, topMargin=12*mm, bottomMargin=12*mm,
                          title="Sismologia Lab - Informe experimental")
    story=[]
    scope=(group_name.upper() if group_name else "TODAS LAS IA")
    story.append(Paragraph("SISMOLOGIA LAB - INFORME EXPERIMENTAL", title))
    story.append(Paragraph(f"Alcance: <b>{_pdf_safe(scope)}</b> | Generado: {datetime.now():%d/%m/%Y %H:%M} (Chile)", sub))
    story.append(Paragraph("Este documento evalua estimaciones experimentales contra eventos registrados por el CSN. No constituye un sistema cientificamente validado de prediccion de terremotos.", body))
    story.append(Spacer(1,5*mm))

    results=[]
    hit=miss=pending=almost=0
    for pred in ps:
        st,b,rad,cands=evaluate_real(pred,es)
        results.append((pred,st,b,rad,cands))
        if st=="ACERTADA": hit+=1
        elif st=="NO ACERTADA": miss+=1
        elif st=="CASI ACERTADA": almost+=1
        else: pending+=1
    closed=hit+miss+almost
    pct=(100.0*hit/closed) if closed else 0.0

    summary=[["Predicciones","Acertadas","Casi acertadas","No acertadas","Pendientes","Precision cerrada","Eventos CSN"],
             [str(len(ps)),str(hit),str(almost),str(miss),str(pending),f"{pct:.1f}%" if closed else "-",str(len(es))]]
    t=Table(summary, colWidths=[34*mm,34*mm,39*mm,39*mm,34*mm,43*mm,34*mm])
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#1f2937")),("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("ALIGN",(0,0),(-1,-1),"CENTER"),
        ("GRID",(0,0),(-1,-1),0.4,colors.HexColor("#bbbbbb")),("BOTTOMPADDING",(0,0),(-1,-1),6),("TOPPADDING",(0,0),(-1,-1),6)
    ]))
    story.append(t)
    story.append(Spacer(1,5*mm))

    # Ranking uses exactly the same evaluation results.
    story.append(Paragraph("Ranking por IA",h2))
    groups={}
    for pred,st,b,rad,cands in results:
        g=pred["group_name"]
        groups.setdefault(g,{"h":0,"a":0,"m":0,"p":0})
        if st=="ACERTADA": groups[g]["h"]+=1
        elif st=="NO ACERTADA": groups[g]["m"]+=1
        elif st=="CASI ACERTADA": groups[g]["a"]+=1
        else: groups[g]["p"]+=1
    rank=[]
    for g,v in groups.items():
        c=v["h"]+v["a"]+v["m"]
        score=(100*v["h"]/c) if c else None
        rank.append((score if score is not None else -1,g,v))
    rank.sort(reverse=True)
    rank_data=[["IA","Aciertos","Casi","Fallos","Pendientes","Precision"]]
    for score,g,v in rank:
        rank_data.append([g,str(v["h"]),str(v["a"]),str(v["m"]),str(v["p"]),f"{score:.1f}%" if score>=0 else "-"])
    if len(rank_data)==1: rank_data.append(["-","0","0","0","0","-"])
    rt=Table(rank_data, colWidths=[55*mm,30*mm,30*mm,30*mm,35*mm,35*mm], repeatRows=1)
    rt.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#374151")),("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("ALIGN",(1,1),(-1,-1),"CENTER"),
        ("GRID",(0,0),(-1,-1),0.35,colors.HexColor("#c7c7c7")),("FONTSIZE",(0,0),(-1,-1),8),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white,colors.HexColor("#f6f7f8")])
    ]))
    story.append(rt)
    story.append(PageBreak())

    story.append(Paragraph("Detalle de predicciones",h2))
    detail=[["Codigo","IA","Ventana","Lugar","Magnitud","Radio","Estado","Sismo CSN relacionado"]]
    for pred,st,b,rad,cands in results:
        timeband=""
        if pred["daily_time_start"] and pred["daily_time_end"]:
            timeband=f" {pred['daily_time_start']}-{pred['daily_time_end']}"
        window=f"{pred['date_start']} a {pred['date_end']}{timeband}"
        related="-"
        if b:
            e=b["e"]; when=parse_datetime(e["occurred_at"])
            related=f"{when:%d/%m %H:%M} | M{e['magnitude']:.1f} | {b['dist']:.1f} km | {e['place']}"
        detail.append([
            Paragraph(_pdf_safe(pred["code"]),small), Paragraph(_pdf_safe(pred["group_name"]),small),
            Paragraph(_pdf_safe(window),small), Paragraph(_pdf_safe(pred["place"]),small),
            f"M{pred['mag_min']:.1f}-{pred['mag_max']:.1f}", f"{rad:.0f} km", st,
            Paragraph(_pdf_safe(related),small)
        ])
    if len(detail)==1:
        detail.append(["-","-","-","No hay predicciones para este filtro","-","-","-","-"])
    dt=Table(detail, colWidths=[25*mm,22*mm,39*mm,55*mm,27*mm,20*mm,27*mm,67*mm], repeatRows=1)
    dt.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#111827")),("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("VALIGN",(0,0),(-1,-1),"TOP"),
        ("GRID",(0,0),(-1,-1),0.3,colors.HexColor("#c7c7c7")),("FONTSIZE",(0,0),(-1,-1),7.3),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white,colors.HexColor("#f8f8f8")]),
        ("TEXTCOLOR",(6,1),(6,-1),colors.HexColor("#111111"))
    ]))
    story.append(dt)

    # Inspector section: one block per confirmed hit.
    hits=[x for x in results if x[1] in ("ACERTADA", "CASI ACERTADA") and x[2]]
    if hits:
        story.append(PageBreak())
        story.append(Paragraph("Inspector de coincidencias acertadas y casi acertadas",h2))
        inspect=[["Prediccion","Sismo CSN","Distancia","Magnitud","Tiempo","Ubicacion","Enlace"]]
        for pred,st,b,rad,cands in hits:
            e=b["e"]; when=parse_datetime(e["occurred_at"])
            link=e["source_url"] or "-"
            inspect.append([
                pred["code"], Paragraph(_pdf_safe(f"{when:%d/%m/%Y %H:%M} - M{e['magnitude']:.1f} - {e['place']}"),small),
                f"{b['dist']:.1f}/{rad:.0f} km", "OK", "OK", "OK", Paragraph(_pdf_safe(link),small)
            ])
        it=Table(inspect,colWidths=[30*mm,85*mm,35*mm,25*mm,25*mm,25*mm,60*mm],repeatRows=1)
        it.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#065f46")),("TEXTCOLOR",(0,0),(-1,0),colors.white),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("VALIGN",(0,0),(-1,-1),"TOP"),
            ("GRID",(0,0),(-1,-1),0.35,colors.HexColor("#b7b7b7")),("FONTSIZE",(0,0),(-1,-1),7.5)
        ]))
        story.append(it)

    def footer(canvas,doc):
        canvas.saveState(); canvas.setFont("Helvetica",7); canvas.setFillColor(colors.HexColor("#666666"))
        canvas.drawString(12*mm,6*mm,"Sismologia Lab - evaluacion experimental basada en eventos CSN guardados")
        canvas.drawRightString(landscape(A4)[0]-12*mm,6*mm,f"Pagina {doc.page}")
        canvas.restoreState()
    doc.build(story,onFirstPage=footer,onLaterPages=footer)
    return {"predictions":len(ps),"hits":hit,"misses":miss,"pending":pending,"events":len(es)}

@bot.tree.command(name="informe_pdf",description="Genera un PDF con resultados del experimento")
@app_commands.describe(ia="IA/grupo a incluir, o TODAS")
async def c_informe_pdf(i:discord.Interaction, ia:str="TODAS"):
    await i.response.defer(thinking=True)
    requested=(ia or "TODAS").strip()
    group=None if requested.upper() in {"TODAS","TODO","ALL","*"} else requested.upper()
    if group and not get_real_group(group):
        await i.followup.send(f"❌ No encontré predicciones para `{group}`. Usa `TODAS` o un grupo importado.")
        return
    safe_name=re.sub(r"[^A-Za-z0-9_-]+","_",group or "TODAS")
    path=Path(tempfile.gettempdir())/f"Sismologia_Lab_{safe_name}_{datetime.now():%Y%m%d_%H%M%S}.pdf"
    try:
        stats=build_experiment_pdf(path,group)
        msg=(f"📄 **Informe PDF generado** - {group or 'TODAS LAS IA'}\n"
             f"🎯 {stats['predictions']} predicciones · ✅ {stats['hits']} aciertos · "
             f"❌ {stats['misses']} fallos · 🟡 {stats['pending']} pendientes")
        await i.followup.send(msg,file=discord.File(str(path),filename=path.name))
    except Exception as exc:
        print("Error generando informe PDF:",repr(exc))
        await i.followup.send(f"❌ No pude generar el PDF: `{exc}`")
    finally:
        try: path.unlink(missing_ok=True)
        except Exception: pass

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



# ---------- V1.1: CALCULADORAS ----------

_ALLOWED_BINOPS = {
    ast.Add: lambda a,b: a+b,
    ast.Sub: lambda a,b: a-b,
    ast.Mult: lambda a,b: a*b,
    ast.Div: lambda a,b: a/b,
    ast.FloorDiv: lambda a,b: a//b,
    ast.Mod: lambda a,b: a%b,
    ast.Pow: lambda a,b: a**b,
}
_ALLOWED_UNARY = {
    ast.UAdd: lambda a:+a,
    ast.USub: lambda a:-a,
}

def safe_calculate(expression):
    expression=str(expression).strip()
    if len(expression) > 100:
        raise ValueError("expresión demasiado larga")
    tree=ast.parse(expression, mode="eval")

    def ev(node):
        if isinstance(node, ast.Expression):
            return ev(node.body)
        if isinstance(node, ast.Constant) and type(node.value) in (int,float):
            return node.value
        if isinstance(node, ast.BinOp) and type(node.op) in _ALLOWED_BINOPS:
            a,b=ev(node.left),ev(node.right)
            # Evitar cálculos gigantes accidentales.
            if isinstance(node.op, ast.Pow) and abs(b) > 12:
                raise ValueError("exponente demasiado grande")
            value=_ALLOWED_BINOPS[type(node.op)](a,b)
            if abs(value) > 1e15:
                raise ValueError("resultado demasiado grande")
            return value
        if isinstance(node, ast.UnaryOp) and type(node.op) in _ALLOWED_UNARY:
            return _ALLOWED_UNARY[type(node.op)](ev(node.operand))
        raise ValueError("solo se permiten números, paréntesis y + - * / // % **")

    return ev(tree)

@bot.tree.command(name="calcular",description="Calculadora normal segura")
@app_commands.describe(expresion="Ej.: (25*3)+10 o 2**8")
async def c_calcular(i:discord.Interaction, expresion:str):
    try:
        result=safe_calculate(expresion)
        if isinstance(result,float):
            shown=f"{result:.10g}"
        else:
            shown=str(result)
        await i.response.send_message(
            f"🧮 **CALCULADORA**\n`{expresion}`\n= **{shown}**"
        )
    except ZeroDivisionError:
        await i.response.send_message("❌ No se puede dividir por cero.")
    except Exception as exc:
        await i.response.send_message(f"❌ Expresión no válida: `{exc}`")

@bot.tree.command(name="distancia",description="Calcula distancia geográfica entre dos coordenadas")
@app_commands.describe(
    lat1="Latitud del punto 1", lon1="Longitud del punto 1",
    lat2="Latitud del punto 2", lon2="Longitud del punto 2"
)
async def c_distancia(i:discord.Interaction,lat1:float,lon1:float,lat2:float,lon2:float):
    if not (-90<=lat1<=90 and -90<=lat2<=90 and -180<=lon1<=180 and -180<=lon2<=180):
        await i.response.send_message("❌ Coordenadas fuera de rango.")
        return
    km=haversine(lat1,lon1,lat2,lon2)
    await i.response.send_message(
        "🌎 **DISTANCIA GEOGRÁFICA**\n"
        f"Punto 1: `{lat1:.5f}, {lon1:.5f}`\n"
        f"Punto 2: `{lat2:.5f}, {lon2:.5f}`\n"
        f"📏 Distancia: **{km:.2f} km**\n\n"
        "Usa la misma fórmula geográfica que el motor del bot."
    )

@bot.tree.command(name="comparar_magnitud",description="Compara una magnitud observada con una predicción y margen")
@app_commands.describe(
    magnitud_predicha="Magnitud central predicha",
    margen="Margen permitido, ej. 0.3",
    magnitud_real="Magnitud observada por CSN"
)
async def c_comp_mag(i:discord.Interaction,magnitud_predicha:float,margen:float,magnitud_real:float):
    margen=abs(margen)
    lo=magnitud_predicha-margen
    hi=magnitud_predicha+margen
    diff=abs(magnitud_real-magnitud_predicha)
    ok=lo<=magnitud_real<=hi
    await i.response.send_message(
        "📈 **COMPARACIÓN DE MAGNITUD**\n"
        f"Predicción: **M{magnitud_predicha:.2f} ±{margen:.2f}**\n"
        f"Rango: **M{lo:.2f}–M{hi:.2f}**\n"
        f"CSN: **M{magnitud_real:.2f}**\n"
        f"Diferencia absoluta: **{diff:.2f}**\n"
        f"Resultado: {'✅ DENTRO DEL MARGEN' if ok else '❌ FUERA DEL MARGEN'}"
    )

@bot.tree.command(name="comparar_tiempo",description="Calcula diferencia temporal entre predicción y sismo")
@app_commands.describe(
    fecha_predicha="DD/MM/AAAA HH:MM, ej. 11/09/2026 18:30",
    fecha_real="DD/MM/AAAA HH:MM, ej. 11/09/2026 19:45",
    margen_horas="Margen permitido en horas"
)
async def c_comp_time(i:discord.Interaction,fecha_predicha:str,fecha_real:str,margen_horas:float):
    try:
        a=datetime.strptime(fecha_predicha.strip(),"%d/%m/%Y %H:%M")
        b=datetime.strptime(fecha_real.strip(),"%d/%m/%Y %H:%M")
    except ValueError:
        await i.response.send_message("❌ Usa el formato `DD/MM/AAAA HH:MM`.")
        return
    delta=abs((b-a).total_seconds())/3600
    ok=delta<=abs(margen_horas)
    await i.response.send_message(
        "🕒 **COMPARACIÓN TEMPORAL**\n"
        f"Predicción: **{a:%d/%m/%Y %H:%M}**\n"
        f"Sismo: **{b:%d/%m/%Y %H:%M}**\n"
        f"Diferencia: **{delta:.2f} h** ({delta*60:.0f} min)\n"
        f"Margen permitido: **±{abs(margen_horas):.2f} h**\n"
        f"Resultado: {'✅ DENTRO DEL MARGEN' if ok else '❌ FUERA DEL MARGEN'}"
    )

@bot.tree.command(name="comparar",description="Calculadora rápida: magnitud + distancia + tiempo")
@app_commands.describe(
    magnitud_predicha="Magnitud central predicha",
    margen_magnitud="Margen ± de magnitud",
    magnitud_real="Magnitud CSN",
    distancia_km="Distancia observada en km",
    radio_km="Radio máximo permitido",
    diferencia_horas="Diferencia absoluta de tiempo en horas",
    margen_horas="Margen temporal permitido en horas"
)
async def c_comparar(
    i:discord.Interaction,
    magnitud_predicha:float,
    margen_magnitud:float,
    magnitud_real:float,
    distancia_km:float,
    radio_km:float,
    diferencia_horas:float,
    margen_horas:float
):
    mm=abs(margen_magnitud)
    mh=abs(margen_horas)
    lo,hi=magnitud_predicha-mm,magnitud_predicha+mm
    mag_ok=lo<=magnitud_real<=hi
    geo_ok=distancia_km<=abs(radio_km)
    time_ok=abs(diferencia_horas)<=mh
    passed=sum((mag_ok,geo_ok,time_ok))
    result="✅ COINCIDENCIA 3/3" if passed==3 else ("🟡 CASI COINCIDENCIA 2/3" if passed==2 else f"❌ NO COINCIDE ({passed}/3)")
    await i.response.send_message(
        "🧮 **COMPARACIÓN COMPLETA**\n\n"
        f"📈 Magnitud: M{magnitud_real:.2f} vs M{lo:.2f}–M{hi:.2f} → {'✅' if mag_ok else '❌'}\n"
        f"📏 Ubicación: {distancia_km:.2f} km / máximo {abs(radio_km):.2f} km → {'✅' if geo_ok else '❌'}\n"
        f"🕒 Tiempo: {abs(diferencia_horas):.2f} h / máximo {mh:.2f} h → {'✅' if time_ok else '❌'}\n\n"
        f"**{result}**"
    )

@bot.tree.command(name="radio",description="Muestra el radio automático que usaría el motor")
@app_commands.describe(
    magnitud="Magnitud central de la predicción",
    tipo_zona="Ej.: costera, interior, cordillera"
)
async def c_radio(i:discord.Interaction,magnitud:float,tipo_zona:str=""):
    r=radius(magnitud,tipo_zona)
    await i.response.send_message(
        "📐 **RADIO AUTOMÁTICO DEL MOTOR**\n"
        f"Magnitud de referencia: **M{magnitud:.2f}**\n"
        f"Tipo de zona: **{tipo_zona or 'sin especificar'}**\n"
        f"Radio calculado: **{r} km**"
    )





# ---------- V1.2: INTERFAZ VISUAL ----------

def ui_embed(title, description="", color=discord.Color.blurple()):
    return discord.Embed(title=title, description=description, color=color)

def fmt_bool(v):
    return "✅ Cumple" if v else "❌ No cumple"

def build_status_embed():
    channel_id=get_setting("alert_channel_id")
    with connect() as con:
        n_events=con.execute("SELECT COUNT(*) FROM observed_events").fetchone()[0]
        n_preds=con.execute("SELECT COUNT(*) FROM real_predictions").fetchone()[0]
        n_corr=con.execute("SELECT COUNT(*) FROM correlation_history").fetchone()[0]
    e=ui_embed("⚙️ Estado del Bot Sísmico","Resumen rápido del sistema.")
    e.add_field(name="🌐 CSN",value=f"Consulta cada **{CHECK_MINUTES} min**",inline=True)
    e.add_field(name="🎯 Predicciones",value=f"**{n_preds}** guardadas",inline=True)
    e.add_field(name="🌎 Eventos CSN",value=f"**{n_events}** guardados",inline=True)
    e.add_field(name="📑 Correlaciones",value=f"**{n_corr}** registradas",inline=True)
    e.add_field(name="🔔 Canal de alertas",value=(f"<#{channel_id}>" if channel_id else "No configurado"),inline=True)
    e.add_field(name="💾 Persistencia",value=("`"+DB_PATH+"`"),inline=True)
    e.set_footer(text="CSN + predicciones • v1.2")
    return e

def ranking_rows():
    ps=get_real_group(); es=events(); groups={}
    for p in ps: groups.setdefault(p["group_name"],[]).append(p)
    rows=[]
    for group,gps in groups.items():
        hits=almost=misses=pending=0
        for p in gps:
            st,_,_,_=evaluate_real(p,es)
            if st=="ACERTADA": hits+=1
            elif st=="CASI ACERTADA": almost+=1
            elif st=="NO ACERTADA": misses+=1
            else: pending+=1
        # Prioridad solicitada: más aciertos > más casi > menos no acertados.
        rows.append((group,hits,almost,misses,pending))
    rows.sort(key=lambda r:(-r[1],-r[2],r[3],r[4],r[0].lower()))
    return rows

def save_ranking_snapshot(force=False):
    """Guarda una foto del ranking solo cuando cambia. El historial empieza al instalar esta versión."""
    rows=ranking_rows()
    payload=json.dumps(rows,ensure_ascii=False,separators=(",",":"))
    with connect() as con:
        last=con.execute("SELECT ranking_json FROM ranking_snapshots ORDER BY id DESC LIMIT 1").fetchone()
        if force or last is None or last["ranking_json"]!=payload:
            con.execute("INSERT INTO ranking_snapshots(captured_at,ranking_json) VALUES(?,?)",(datetime.now().isoformat(),payload))
            return True
    return False

def _ranking_history():
    with connect() as con:
        return con.execute("SELECT * FROM ranking_snapshots ORDER BY id ASC").fetchall()

def build_evolution_embed():
    hist=_ranking_history(); e=ui_embed("📈 Evolución del ranking","Historial de posiciones guardado desde la instalación de esta actualización.")
    if not hist:
        e.description += "\n\nAún no hay snapshots. Se guardará uno cuando cambien los resultados del experimento."
        return e
    first=json.loads(hist[0]["ranking_json"]); last=json.loads(hist[-1]["ranking_json"])
    groups=[]
    for r in last:
        if r[0] not in groups: groups.append(r[0])
    for g in groups[:10]:
        seq=[]
        for row in hist:
            data=json.loads(row["ranking_json"]); pos=next((i+1 for i,x in enumerate(data) if x[0]==g),None)
            if pos is not None and (not seq or seq[-1]!=pos): seq.append(pos)
        current=next((i+1 for i,x in enumerate(last) if x[0]==g),None)
        initial=next((i+1 for i,x in enumerate(first) if x[0]==g),None)
        delta=(initial-current) if initial and current else 0
        trend=(f"▲ {delta}" if delta>0 else f"▼ {abs(delta)}" if delta<0 else "• 0")
        path=" → ".join(f"#{x}" for x in seq[-8:]) or f"#{current}"
        e.add_field(name=f"{'🏆 ' if current==1 else ''}{g}",value=f"{path}\nCambio: **{trend}** · actual **#{current}**",inline=False)
    e.set_footer(text=f"{len(hist)} snapshots · no reconstruye posiciones anteriores a esta versión")
    return e

# PNG mínimo con stdlib: evita añadir dependencias a Railway.
def _png_chunk(kind,data):
    return struct.pack('>I',len(data))+kind+data+struct.pack('>I',zlib.crc32(kind+data)&0xffffffff)

def _encode_png_rgb(w,h,pixels):
    raw=b''.join(b'\x00'+bytes(pixels[y*w*3:(y+1)*w*3]) for y in range(h))
    return b'\x89PNG\r\n\x1a\n'+_png_chunk(b'IHDR',struct.pack('>IIBBBBB',w,h,8,2,0,0,0))+_png_chunk(b'IDAT',zlib.compress(raw,9))+_png_chunk(b'IEND',b'')

def _dot(px,w,h,x,y,r,color):
    for yy in range(max(0,y-r),min(h,y+r+1)):
        for xx in range(max(0,x-r),min(w,x+r+1)):
            if (xx-x)**2+(yy-y)**2<=r*r:
                i=(yy*w+xx)*3; px[i:i+3]=color

def _line(px,w,h,x0,y0,x1,y1,color):
    dx=abs(x1-x0); sx=1 if x0<x1 else -1; dy=-abs(y1-y0); sy=1 if y0<y1 else -1; err=dx+dy
    while True:
        if 0<=x0<w and 0<=y0<h:
            i=(y0*w+x0)*3; px[i:i+3]=color
        if x0==x1 and y0==y1: break
        e2=2*err
        if e2>=dy: err+=dy; x0+=sx
        if e2<=dx: err+=dx; y0+=sy

def build_prediction_map(group=None):
    ps=get_real_group(group) if group else get_real_group(); es=events()
    pairs=[]
    for pred in ps:
        ev,score=best_proximity(pred,es)
        if ev is not None and score is not None: pairs.append((pred,ev,score))
    if not pairs: return None,0
    # Marco geográfico Chile + margen oceánico; se expande si los datos lo requieren.
    lats=[float(p['latitude']) for p,_,_ in pairs]+[float(e['latitude']) for _,e,_ in pairs]
    lons=[float(p['longitude']) for p,_,_ in pairs]+[float(e['longitude']) for _,e,_ in pairs]
    lat_min=min(-57,min(lats)-1); lat_max=max(-17,max(lats)+1); lon_min=min(-77,min(lons)-1); lon_max=max(-66,max(lons)+1)
    w,h=900,1200; margin=55; px=bytearray([248,249,250])*(w*h)
    def xy(lat,lon):
        x=margin+int((lon-lon_min)/(lon_max-lon_min)*(w-2*margin)); y=margin+int((lat_max-lat)/(lat_max-lat_min)*(h-2*margin)); return x,y
    # rejilla geográfica
    for lat in range(-55,-19,5):
        x0,y=xy(lat,lon_min); x1,_=xy(lat,lon_max); _line(px,w,h,x0,y,x1,y,(220,224,228))
    for lon in range(-75,-65,2):
        x,y0=xy(lat_max,lon); _,y1=xy(lat_min,lon); _line(px,w,h,x,y0,x,y1,(220,224,228))
    # silueta aproximada del eje chileno para orientación visual
    chile=[(-17.5,-69.5),(-23,-70.4),(-30,-71.4),(-35,-72.5),(-41,-73.7),(-46,-74.5),(-52,-73.5),(-56,-68.5)]
    for a,b in zip(chile,chile[1:]):
        x0,y0=xy(*a); x1,y1=xy(*b); _line(px,w,h,x0,y0,x1,y1,(70,80,90))
    for pred,ev,score in pairs:
        xp,yp=xy(float(pred['latitude']),float(pred['longitude'])); xe,ye=xy(float(ev['latitude']),float(ev['longitude']))
        _line(px,w,h,xp,yp,xe,ye,(160,160,160)); _dot(px,w,h,xp,yp,7,(220,55,55)); _dot(px,w,h,xe,ye,6,(45,105,210))
    return _encode_png_rgb(w,h,px),len(pairs)

def build_research_embed():
    rows=ranking_rows(); ps=get_real_group(); es=events(); total=len(ps); finalized=0; hit=almost=miss=pending=0
    for p in ps:
        st,_,_,_=evaluate_real(p,es)
        if st=='ACERTADA': hit+=1; finalized+=1
        elif st=='CASI ACERTADA': almost+=1; finalized+=1
        elif st=='NO ACERTADA': miss+=1; finalized+=1
        else: pending+=1
    e=ui_embed("🧪 Modo investigación","Vista centrada en resultados comparables del experimento, sin información técnica del bot.")
    e.add_field(name="📦 Muestra",value=f"**{total}** predicciones · **{finalized}** finalizadas · **{pending}** pendientes",inline=False)
    e.add_field(name="🎯 Resultados globales",value=f"✅ **{hit}** · 🟠 **{almost}** · ❌ **{miss}**",inline=False)
    if rows:
        top=rows[0]; e.add_field(name="🏆 Líder actual",value=f"**{top[0]}** · ✅ {top[1]} · 🟠 {top[2]} · ❌ {top[3]}",inline=False)
    profiles=[]
    for g,_,_,_,_ in rows[:5]:
        a=_group_analysis(g); n=a['evaluated'] or 1
        profiles.append(f"**{g}** — tiempo {100*a['time_ok']/n:.0f}% · magnitud {100*a['mag_ok']/n:.0f}% · ubicación {100*a['geo_ok']/n:.0f}%")
    if profiles:e.add_field(name="🔬 Cumplimiento por criterio",value="\n".join(profiles),inline=False)
    e.set_footer(text="Usa los botones para profundizar: ranking, comparar, mapa y evolución.")
    return e

def build_ranking_embed():
    rows=ranking_rows()
    e=ui_embed("🏆 Ranking por IA","Prioridad: **✅ aciertos → 🟠 casi acertados → ❌ menos no acertados**. Las pendientes no deciden el puesto.")
    if not rows:
        e.description="Todavía no hay predicciones importadas."
        return e
    medals=["🥇","🥈","🥉"]
    for idx,(g,h,a,m,pending) in enumerate(rows[:10]):
        icon=medals[idx] if idx<3 else f"**#{idx+1}**"
        closed=h+a+m
        exact_pct=(100*h/closed) if closed else 0
        e.add_field(name=f"{icon} {g}",value=(
            f"✅ **{h}** aciertos  ·  🟠 **{a}** casi  ·  ❌ **{m}** no acertados\n"
            f"🟡 {pending} pendientes  ·  🎯 {exact_pct:.1f}% acierto exacto"),inline=False)
    e.set_footer(text="Desempate: aciertos > casi > menos fallos")
    return e



def _status_icon(st):
    return {"ACERTADA":"✅","CASI ACERTADA":"🟠","NO ACERTADA":"❌","PENDIENTE":"🟡"}.get(st,"🎯")

def _group_analysis(group):
    ps=get_real_group(group); es=events()
    out={"total":len(ps),"hits":0,"almost":0,"misses":0,"pending":0,"finalized":0,
         "distances":[],"mag_errors":[],"time_errors":[],"time_ok":0,"mag_ok":0,"geo_ok":0,"evaluated":0}
    for pred in ps:
        st,b,rad,cands=evaluate_real(pred,es)
        if st=="ACERTADA": out["hits"]+=1
        elif st=="CASI ACERTADA": out["almost"]+=1
        elif st=="NO ACERTADA": out["misses"]+=1
        else: out["pending"]+=1
        if st!="PENDIENTE": out["finalized"]+=1
        ev,score=best_proximity(pred,es)
        if ev is not None and score is not None:
            out["evaluated"]+=1; out["distances"].append(float(score["distance_km"])); out["time_errors"].append(float(score["time_error_hours"]))
            mag=float(ev["magnitude"]); lo=float(pred["mag_min"]); hi=float(pred["mag_max"])
            out["mag_errors"].append(0.0 if lo<=mag<=hi else min(abs(mag-lo),abs(mag-hi)))
            when=parse_datetime(ev["occurred_at"]); out["time_ok"]+=int(real_time_ok(when,pred)); out["mag_ok"]+=int(lo<=mag<=hi)
            out["geo_ok"]+=int(float(score["distance_km"])<=float(score["radius_km"]))
    return out

def build_ai_profile_embed(group):
    a=_group_analysis(group); rows=ranking_rows(); pos=next((i+1 for i,r in enumerate(rows) if r[0]==group),None)
    closed=a["finalized"]; exact=100*a["hits"]/closed if closed else 0
    e=ui_embed(f"🤖 Perfil · {group}",f"{'🏆 Puesto **#'+str(pos)+'**' if pos else 'Sin puesto todavía'} · muestra: **{closed} finalizadas** de {a['total']}")
    e.add_field(name="📊 Resultados",value=f"✅ **{a['hits']}** · 🟠 **{a['almost']}** · ❌ **{a['misses']}** · 🟡 **{a['pending']}**\n🎯 Acierto exacto: **{exact:.1f}%**",inline=False)
    if a["evaluated"]:
        avg=lambda xs: sum(xs)/len(xs) if xs else 0
        n=a["evaluated"]
        e.add_field(name="🔬 Precisión descriptiva",value=(f"📍 Distancia media al mejor candidato: **{avg(a['distances']):.1f} km**\n"
            f"📈 Error medio fuera del rango: **{avg(a['mag_errors']):.2f} M**\n🕒 Error temporal medio fuera de ventana: **{avg(a['time_errors']):.2f} h**"),inline=False)
        e.add_field(name="🧩 Cumplimiento del mejor candidato",value=f"🕒 Tiempo **{100*a['time_ok']/n:.1f}%** · 📈 Magnitud **{100*a['mag_ok']/n:.1f}%** · 📍 Ubicación **{100*a['geo_ok']/n:.1f}%**",inline=False)
    e.set_footer(text="Las métricas descriptivas usan el mejor candidato CSN disponible para cada predicción.")
    return e

def build_comparison_embed(a_name,b_name):
    a=_group_analysis(a_name); b=_group_analysis(b_name)
    e=ui_embed(f"⚔️ {a_name} vs {b_name}","Comparación directa. El ranking oficial sigue priorizando aciertos → casi → menos fallos.")
    def line(x):
        closed=x['finalized']; pct=100*x['hits']/closed if closed else 0
        return f"✅ **{x['hits']}** · 🟠 **{x['almost']}** · ❌ **{x['misses']}** · 🟡 {x['pending']}\n🎯 {pct:.1f}% exacto · 🧪 **{closed}** finalizadas"
    e.add_field(name=f"🤖 {a_name}",value=line(a),inline=True); e.add_field(name=f"🤖 {b_name}",value=line(b),inline=True)
    def av(x,k): return sum(x[k])/len(x[k]) if x[k] else 0
    if a['evaluated'] and b['evaluated']:
        e.add_field(name="🔬 Errores medios",value=(f"📍 Distancia: **{av(a,'distances'):.1f}** vs **{av(b,'distances'):.1f} km**\n"
            f"📈 Magnitud: **{av(a,'mag_errors'):.2f}** vs **{av(b,'mag_errors'):.2f} M**\n"
            f"🕒 Tiempo: **{av(a,'time_errors'):.2f}** vs **{av(b,'time_errors'):.2f} h**"),inline=False)
    return e

def build_error_embed(pred):
    es=events(); st,b,rad,cands=evaluate_real(pred,es); ev,score=best_proximity(pred,es)
    e=ui_embed(f"🎯 Qué tan cerca estuvo · {pred['code']}",f"{_status_icon(st)} Resultado actual: **{st}**")
    if ev is None or score is None:
        e.description += "\n\nTodavía no hay un evento CSN disponible para comparar."; return e
    mag=float(ev['magnitude']); lo=float(pred['mag_min']); hi=float(pred['mag_max']); mag_out=0 if lo<=mag<=hi else min(abs(mag-lo),abs(mag-hi))
    dist=float(score['distance_km']); radius_km=float(score['radius_km']); geo_out=max(0,dist-radius_km); time_out=float(score['time_error_hours'])
    e.add_field(name="📍 Ubicación",value=f"Distancia **{dist:.1f} km** / radio **{radius_km:.1f} km**\n{'✅ Dentro del radio' if geo_out==0 else f'❌ {geo_out:.1f} km fuera del radio'}",inline=False)
    e.add_field(name="📈 Magnitud",value=f"CSN **M{mag:.1f}** / esperado **M{lo:.1f}–M{hi:.1f}**\n{'✅ Dentro del rango' if mag_out==0 else f'❌ {mag_out:.2f} M fuera del rango'}",inline=False)
    e.add_field(name="🕒 Tiempo",value=('✅ Dentro de la ventana' if time_out==0 else f'❌ **{time_out:.2f} h** fuera de la ventana'),inline=False)
    e.add_field(name="📊 Proximidad",value=f"Score descriptivo **{score['score']:.1f}%** · espacial {score['spatial_pct']:.1f}% · magnitud {score['magnitude_pct']:.1f}% · temporal {score['temporal_pct']:.1f}%",inline=False)
    return e

def build_activity_embed():
    ps=get_real_group(); es=events(); finished=[]; active=[]
    for pred in ps:
        st,b,rad,c=evaluate_real(pred,es)
        (active if st=='PENDIENTE' else finished).append((pred,st,b))
    e=ui_embed("📡 Centro de actividad","Lo más importante que está ocurriendo en el experimento.")
    e.add_field(name="🎯 Predicciones",value=f"🟡 **{len(active)}** activas · 🧾 **{len(finished)}** finalizadas",inline=False)
    if es:
        x=es[0]; e.add_field(name="🌎 Último sismo CSN",value=f"**M{x['magnitude']:.1f}** · {x['place'] or 'Sin referencia'}\n🕒 {str(x['occurred_at']).replace('T',' ')[:19]}",inline=False)
    if active:
        upcoming=sorted(active,key=lambda z:_window_bounds(z[0])[1])[:3]
        e.add_field(name="⏳ Próximas en finalizar",value="\n".join(f"• **{x[0]['code']}** · {_window_bounds(x[0])[1]:%d/%m %H:%M}" for x in upcoming),inline=False)
    recent=[x for x in finished if x[1] in ('ACERTADA','CASI ACERTADA') and x[2] is not None]
    if recent:
        recent.sort(key=lambda z:parse_datetime(z[2]['e']['occurred_at']),reverse=True); pred,st,b=recent[0]
        e.add_field(name="✨ Último resultado destacado",value=f"{_status_icon(st)} **{pred['code']}** · {st}\n🌎 M{b['e']['magnitude']:.1f} · {b['e']['place'] or 'Sin referencia'}",inline=False)
    return e

def build_history_embed():
    with connect() as con:
        rows=con.execute("""SELECT * FROM correlation_history
                            ORDER BY detected_at DESC LIMIT 8""").fetchall()
    e=ui_embed("📑 Historial de correlaciones","Últimas coincidencias registradas automáticamente.")
    if not rows:
        e.description="Todavía no hay correlaciones guardadas."
        return e
    for r in rows:
        icon="🚨" if r["match_level"]=="COINCIDENCIA" else "🟡"
        when=datetime.fromisoformat(r["event_time"])
        e.add_field(
            name=f"{icon} {r['prediction_code']} ↔ CSN #{r['event_source_id']}",
            value=f"📅 {when:%d/%m/%Y %H:%M}  •  📈 M{r['event_mag']:.1f}  •  📏 {r['distance_km']:.1f} km",
            inline=False
        )
    return e

def build_help_embed():
    e=ui_embed(
        "❓ Centro de ayuda",
        "Usa `/panel` como menú principal. Los comandos antiguos siguen funcionando."
    )
    e.add_field(name="🎯 Predicciones",value="`/ver_real` · `/analizar_grupo` · `/analizar_todos` · `/importar_todos`",inline=False)
    e.add_field(name="🌐 CSN",value="`/actualizar_csn` · `/importar_historico` · `/ultimos_csn`",inline=False)
    e.add_field(name="🧮 Calculadora",value="Desde `/panel` → **Calculadora**, o `/calcular`, `/distancia`, `/comparar`, `/radio`.",inline=False)
    e.add_field(name="📊 Resultados",value="`/ranking` · `/historial` · `/integridad`",inline=False)
    e.add_field(name="🔔 Alertas",value="`/canal_alertas` · `/estado_alertas` · `/desactivar_alertas`",inline=False)
    e.set_footer(text="Los botones del panel no modifican predicciones por sí solos.")
    return e

class BasicCalcModal(discord.ui.Modal, title="🧮 Calculadora"):
    expression=discord.ui.TextInput(
        label="Expresión",
        placeholder="Ej.: 25**(1/2)  o  (25*3)+10",
        max_length=100
    )
    async def on_submit(self, interaction):
        try:
            result=safe_calculate(str(self.expression))
            shown=f"{result:.10g}" if isinstance(result,float) else str(result)
            e=ui_embed("🧮 Resultado",f"`{self.expression}`\n\n### = **{shown}**",discord.Color.green())
            await interaction.response.send_message(embed=e,ephemeral=True)
        except Exception as exc:
            await interaction.response.send_message(f"❌ Expresión no válida: `{exc}`",ephemeral=True)

class RootModal(discord.ui.Modal, title="√ Calculadora de raíces"):
    number=discord.ui.TextInput(label="Número",placeholder="Ej.: 25")
    index=discord.ui.TextInput(label="Índice de la raíz",placeholder="2 = cuadrada, 3 = cúbica",default="2")
    async def on_submit(self, interaction):
        try:
            x=float(str(self.number).replace(",","."))
            n=float(str(self.index).replace(",","."))
            if n==0: raise ValueError("el índice no puede ser 0")
            if x<0 and abs(n-round(n))<1e-9 and int(round(n))%2==1:
                result=-((-x)**(1/n))
            elif x<0:
                raise ValueError("esta calculadora usa resultados reales; esa raíz no es real")
            else:
                result=x**(1/n)
            e=ui_embed("√ Resultado",f"Raíz **{n:g}** de **{x:g}**\n\n### = **{result:.10g}**",discord.Color.green())
            await interaction.response.send_message(embed=e,ephemeral=True)
        except Exception as exc:
            await interaction.response.send_message(f"❌ No pude calcularla: `{exc}`",ephemeral=True)

class DistanceModal(discord.ui.Modal, title="🌎 Distancia entre coordenadas"):
    lat1=discord.ui.TextInput(label="Latitud punto 1",placeholder="-33.05")
    lon1=discord.ui.TextInput(label="Longitud punto 1",placeholder="-71.62")
    lat2=discord.ui.TextInput(label="Latitud punto 2",placeholder="-33.18")
    lon2=discord.ui.TextInput(label="Longitud punto 2",placeholder="-71.74")
    async def on_submit(self, interaction):
        try:
            a,b,c,d=[float(str(x).replace(",",".")) for x in (self.lat1,self.lon1,self.lat2,self.lon2)]
            if not (-90<=a<=90 and -90<=c<=90 and -180<=b<=180 and -180<=d<=180):
                raise ValueError("coordenadas fuera de rango")
            km=haversine(a,b,c,d)
            e=ui_embed("🌎 Distancia geográfica",f"📍 `{a:.5f}, {b:.5f}`\n📍 `{c:.5f}, {d:.5f}`\n\n### 📏 **{km:.2f} km**")
            await interaction.response.send_message(embed=e,ephemeral=True)
        except Exception as exc:
            await interaction.response.send_message(f"❌ Datos no válidos: `{exc}`",ephemeral=True)

class MagnitudeModal(discord.ui.Modal, title="📈 Comparar magnitud"):
    predicted=discord.ui.TextInput(label="Magnitud predicha",placeholder="4.2")
    margin=discord.ui.TextInput(label="Margen ±",placeholder="0.3")
    observed=discord.ui.TextInput(label="Magnitud CSN",placeholder="4.4")
    async def on_submit(self, interaction):
        try:
            pred=float(str(self.predicted).replace(",","."))
            margin=abs(float(str(self.margin).replace(",",".")))
            obs=float(str(self.observed).replace(",","."))
            lo,hi=pred-margin,pred+margin
            ok=lo<=obs<=hi
            e=ui_embed("📈 Comparación de magnitud",color=(discord.Color.green() if ok else discord.Color.red()))
            e.add_field(name="Predicción",value=f"M{pred:.2f} ±{margin:.2f}\n`M{lo:.2f}–M{hi:.2f}`",inline=True)
            e.add_field(name="CSN",value=f"**M{obs:.2f}**",inline=True)
            e.add_field(name="Diferencia",value=f"**{abs(obs-pred):.2f}**",inline=True)
            e.add_field(name="Resultado",value=fmt_bool(ok),inline=False)
            await interaction.response.send_message(embed=e,ephemeral=True)
        except Exception as exc:
            await interaction.response.send_message(f"❌ Datos no válidos: `{exc}`",ephemeral=True)

class TimeModal(discord.ui.Modal, title="🕒 Comparar tiempo"):
    predicted=discord.ui.TextInput(label="Fecha/hora predicha",placeholder="11/09/2026 18:30")
    observed=discord.ui.TextInput(label="Fecha/hora del sismo",placeholder="11/09/2026 19:45")
    margin=discord.ui.TextInput(label="Margen permitido (horas)",placeholder="2")
    async def on_submit(self, interaction):
        try:
            a=datetime.strptime(str(self.predicted).strip(),"%d/%m/%Y %H:%M")
            b=datetime.strptime(str(self.observed).strip(),"%d/%m/%Y %H:%M")
            m=abs(float(str(self.margin).replace(",",".")))
            delta=abs((b-a).total_seconds())/3600
            ok=delta<=m
            e=ui_embed("🕒 Comparación temporal",color=(discord.Color.green() if ok else discord.Color.red()))
            e.add_field(name="Diferencia",value=f"**{delta:.2f} h**\n{delta*60:.0f} min",inline=True)
            e.add_field(name="Margen",value=f"± **{m:.2f} h**",inline=True)
            e.add_field(name="Resultado",value=fmt_bool(ok),inline=False)
            await interaction.response.send_message(embed=e,ephemeral=True)
        except Exception:
            await interaction.response.send_message("❌ Usa fechas con formato `DD/MM/AAAA HH:MM` y un margen numérico.",ephemeral=True)

class FullCompareModal(discord.ui.Modal, title="🎯 Comparación completa"):
    magnitudes=discord.ui.TextInput(label="Magnitud: predicha,margen,real",placeholder="4.2,0.3,4.4")
    geography=discord.ui.TextInput(label="Ubicación: distancia,radio (km)",placeholder="18,25")
    timing=discord.ui.TextInput(label="Tiempo: diferencia,margen (horas)",placeholder="1.5,2")
    async def on_submit(self, interaction):
        try:
            pred,mm,obs=[float(x.strip().replace(",",".")) for x in str(self.magnitudes).replace(";",",").split(",")]
            dist,rad=[float(x.strip().replace(",",".")) for x in str(self.geography).replace(";",",").split(",")]
            dh,mh=[float(x.strip().replace(",",".")) for x in str(self.timing).replace(";",",").split(",")]
            mm,mh,rad=abs(mm),abs(mh),abs(rad)
            mag_ok=pred-mm<=obs<=pred+mm
            geo_ok=dist<=rad
            time_ok=abs(dh)<=mh
            passed=sum((mag_ok,geo_ok,time_ok))
            title="✅ Coincidencia 3/3" if passed==3 else ("🟡 Casi coincidencia 2/3" if passed==2 else f"❌ No coincide ({passed}/3)")
            color=discord.Color.green() if passed==3 else (discord.Color.gold() if passed==2 else discord.Color.red())
            e=ui_embed(title,color=color)
            e.add_field(name="📈 Magnitud",value=f"{fmt_bool(mag_ok)}\nM{obs:.2f} vs M{pred-mm:.2f}–M{pred+mm:.2f}",inline=False)
            e.add_field(name="📍 Ubicación",value=f"{fmt_bool(geo_ok)}\n{dist:.2f} / {rad:.2f} km",inline=False)
            e.add_field(name="🕒 Tiempo",value=f"{fmt_bool(time_ok)}\n{abs(dh):.2f} / {mh:.2f} h",inline=False)
            await interaction.response.send_message(embed=e,ephemeral=True)
        except Exception:
            await interaction.response.send_message(
                "❌ Formato incorrecto. Ejemplos:\n"
                "Magnitud: `4.2,0.3,4.4`\nUbicación: `18,25`\nTiempo: `1.5,2`",
                ephemeral=True
            )

class CalculatorView(discord.ui.View):
    def __init__(self):
        super().__init__(timeout=900)

    @discord.ui.button(label="Normal",emoji="🧮",style=discord.ButtonStyle.primary,row=0)
    async def normal(self,interaction,button):
        await interaction.response.send_modal(BasicCalcModal())

    @discord.ui.button(label="Raíz",style=discord.ButtonStyle.primary,row=0)
    async def root(self,interaction,button):
        await interaction.response.send_modal(RootModal())

    @discord.ui.button(label="Distancia",emoji="🌎",style=discord.ButtonStyle.secondary,row=0)
    async def distance(self,interaction,button):
        await interaction.response.send_modal(DistanceModal())

    @discord.ui.button(label="Magnitud",emoji="📈",style=discord.ButtonStyle.secondary,row=1)
    async def magnitude(self,interaction,button):
        await interaction.response.send_modal(MagnitudeModal())

    @discord.ui.button(label="Tiempo",emoji="🕒",style=discord.ButtonStyle.secondary,row=1)
    async def timing(self,interaction,button):
        await interaction.response.send_modal(TimeModal())

    @discord.ui.button(label="Comparación 3/3",emoji="🎯",style=discord.ButtonStyle.success,row=1)
    async def compare(self,interaction,button):
        await interaction.response.send_modal(FullCompareModal())

def _prediction_groups():
    ps=get_real_group()
    groups={}
    for p in ps:
        groups.setdefault(p["group_name"],[]).append(p)
    return groups

def _prediction_status_details(p):
    es=events()
    st,b,rad,cands=evaluate_real(p,es)
    return st,b,rad,cands

def build_prediction_explorer_embed(group, predictions, index=0):
    if not predictions:
        return ui_embed("🎯 Explorador de predicciones","No hay predicciones importadas para mostrar.")
    index=max(0,min(index,len(predictions)-1)); p=predictions[index]
    st,b,rad,cands=_prediction_status_details(p)
    colors={"ACERTADA":discord.Color.green(),"CASI ACERTADA":discord.Color.gold(),
            "NO ACERTADA":discord.Color.red(),"PENDIENTE":discord.Color.blurple()}
    icons={"ACERTADA":"✅","CASI ACERTADA":"🟠","NO ACERTADA":"❌","PENDIENTE":"🟡"}
    e=ui_embed(f"🔮 {group} · Predicción {index+1}/{len(predictions)}",
               f"**{p['code']}**\n{icons.get(st,'🎯')} **{st}**",colors.get(st,discord.Color.blurple()))
    schedule=f"📅 **{p['date_start']} → {p['date_end']}**"
    if p["daily_time_start"]:
        schedule+=f"\n🕐 **{p['daily_time_start']} → {p['daily_time_end']}**"
    e.add_field(name="🗓️ Ventana temporal",value=schedule,inline=False)
    e.add_field(name="📍 Ubicación predicha",value=(
        f"**{p['place']}**\n🌐 `{p['latitude']:.4f}, {p['longitude']:.4f}`\n"
        f"🧭 {p['zone_type'] or 'Zona no especificada'} · 📏 radio **{rad} km**"),inline=False)
    e.add_field(name="📈 Magnitud",value=f"**M{p['mag_min']:.1f} – M{p['mag_max']:.1f}**",inline=True)
    if p["geo_note"]:
        note=str(p["geo_note"])
        e.add_field(name="📝 Nota geográfica",value=note[:1000],inline=False)
    if b is not None:
        x=b; ev=x["e"]
        try: when=parse_datetime(ev["occurred_at"]).strftime("%d/%m/%Y %H:%M:%S")
        except Exception: when=str(ev["occurred_at"])
        checks=f"🕒 {'✅' if x['t'] else '❌'}  ·  📈 {'✅' if x['m'] else '❌'}  ·  📍 {'✅' if x['g'] else '❌'}"
        result=(f"🌎 **M{ev['magnitude']:.1f}** · {when}\n📍 {ev['place'] or 'Sin referencia'}\n"
                f"📏 **{x['dist']:.1f} km** de la ubicación predicha\n{checks}")
        if ev["source_url"]: result+=f"\n🔗 [Informe oficial CSN]({ev['source_url']})"
        e.add_field(name="🌎 Sismo más relevante",value=result,inline=False)
    elif st=="PENDIENTE":
        e.add_field(name="⏳ Evaluación",value="La ventana todavía está abierta. El resultado final aún no se cierra.",inline=False)
    else:
        e.add_field(name="🌎 Evaluación",value="No se encontró un sismo candidato para mostrar como resultado.",inline=False)
    e.set_footer(text=f"IA: {group} · ◀️/▶️ cambia predicción · menú desplegable cambia IA")
    return e

class PredictionGroupSelect(discord.ui.Select):
    def __init__(self,browser,groups):
        self.browser=browser
        options=[discord.SelectOption(label=name[:100],value=name,description=f"{len(ps)} predicciones",default=(name==browser.group)) for name,ps in list(groups.items())[:25]]
        super().__init__(placeholder="🤖 Cambiar IA / grupo",min_values=1,max_values=1,options=options,row=2)
    async def callback(self,interaction):
        self.browser.group=self.values[0]; self.browser.index=0; self.browser.apply_filter(); self.browser.rebuild_components(); await self.browser.render(interaction)

class PredictionFilterSelect(discord.ui.Select):
    def __init__(self,browser):
        self.browser=browser
        opts=[("TODAS","Todas","📚"),("ACERTADA","Acertadas","✅"),("CASI ACERTADA","Casi acertadas","🟠"),("NO ACERTADA","No acertadas","❌"),("PENDIENTE","Pendientes","🟡")]
        super().__init__(placeholder="🔍 Filtrar por estado",min_values=1,max_values=1,options=[discord.SelectOption(label=l,value=v,emoji=em,default=(browser.filter_status==v)) for v,l,em in opts],row=3)
    async def callback(self,interaction):
        self.browser.filter_status=self.values[0]; self.browser.index=0; self.browser.apply_filter(); self.browser.rebuild_components(); await self.browser.render(interaction)

class PredictionExplorer(discord.ui.View):
    def __init__(self,group=None,index=0):
        super().__init__(timeout=900); self.groups=_prediction_groups(); self.group=group if group in self.groups else next(iter(self.groups),None); self.filter_status="TODAS"; self.predictions=[]; self.index=index; self.apply_filter(); self.rebuild_components()
    def apply_filter(self):
        base=self.groups.get(self.group,[])
        if self.filter_status=="TODAS": self.predictions=list(base)
        else: self.predictions=[p for p in base if evaluate_real(p,events())[0]==self.filter_status]
        self.index=max(0,min(self.index,len(self.predictions)-1)) if self.predictions else 0
    def rebuild_components(self):
        for child in list(self.children):
            if isinstance(child,(PredictionGroupSelect,PredictionFilterSelect)): self.remove_item(child)
        if self.groups:self.add_item(PredictionGroupSelect(self,self.groups)); self.add_item(PredictionFilterSelect(self)); self.sync_buttons()
    def sync_buttons(self):
        disabled=len(self.predictions)<=1; self.previous.disabled=disabled; self.next.disabled=disabled; self.details.disabled=not self.predictions; self.error.disabled=not self.predictions
    async def render(self,interaction):
        self.sync_buttons(); emb=build_prediction_explorer_embed(self.group,self.predictions,self.index)
        if self.filter_status!="TODAS": emb.description=(emb.description or "")+f"\n🔍 Filtro: **{self.filter_status}**"
        await interaction.response.edit_message(embed=emb,view=self)
    @discord.ui.button(label="Anterior",emoji="◀️",style=discord.ButtonStyle.secondary,row=0)
    async def previous(self,interaction,button):
        if self.predictions:self.index=(self.index-1)%len(self.predictions)
        await self.render(interaction)
    @discord.ui.button(label="Siguiente",emoji="▶️",style=discord.ButtonStyle.primary,row=0)
    async def next(self,interaction,button):
        if self.predictions:self.index=(self.index+1)%len(self.predictions)
        await self.render(interaction)
    @discord.ui.button(label="Qué tan cerca",emoji="🎯",style=discord.ButtonStyle.secondary,row=0)
    async def error(self,interaction,button):
        await interaction.response.send_message(embed=build_error_embed(self.predictions[self.index]),ephemeral=True)
    @discord.ui.button(label="Ver sismos",emoji="🌎",style=discord.ButtonStyle.success,row=0)
    async def details(self,interaction,button):
        if not self.predictions: await interaction.response.send_message("No hay predicciones para explorar.",ephemeral=True); return
        p=self.predictions[self.index]; st,rad,cands,initial=_real_browser_data(p)
        await interaction.response.send_message(embed=build_real_browser_embed(p,st,rad,cands,initial),view=RealPredictionBrowser(p,st,rad,cands,initial),ephemeral=True)
    @discord.ui.button(label="Perfil IA",emoji="🤖",style=discord.ButtonStyle.secondary,row=1)
    async def profile(self,interaction,button): await interaction.response.send_message(embed=build_ai_profile_embed(self.group),ephemeral=True)
    @discord.ui.button(label="Cerrar",emoji="✖️",style=discord.ButtonStyle.secondary,row=1)
    async def close(self,interaction,button):
        for child in self.children: child.disabled=True
        await interaction.response.edit_message(view=self)

class CompareAISelect(discord.ui.Select):
    def __init__(self,view,which,groups):
        self.owner=view; self.which=which; current=view.a if which=='a' else view.b
        super().__init__(placeholder=("🤖 IA A" if which=='a' else "🤖 IA B"),options=[discord.SelectOption(label=g,value=g,default=(g==current)) for g in groups[:25]],row=(0 if which=='a' else 1))
    async def callback(self,interaction):
        if self.which=='a': self.owner.a=self.values[0]
        else:self.owner.b=self.values[0]
        await interaction.response.edit_message(embed=build_comparison_embed(self.owner.a,self.owner.b),view=self.owner)

class AICompareView(discord.ui.View):
    def __init__(self):
        super().__init__(timeout=900); groups=list(_prediction_groups()); self.a=groups[0]; self.b=groups[1] if len(groups)>1 else groups[0]; self.add_item(CompareAISelect(self,'a',groups)); self.add_item(CompareAISelect(self,'b',groups))

class PredictionsPanel(discord.ui.View):
    def __init__(self):
        super().__init__(timeout=900)

    @discord.ui.button(label="Explorar", emoji="🔮", style=discord.ButtonStyle.success, row=0)
    async def explore(self, interaction, button):
        groups=_prediction_groups()
        if not groups:
            await interaction.response.send_message("No hay predicciones importadas. Usa `/importar_todos`.",ephemeral=True)
            return
        view=PredictionExplorer()
        await interaction.response.send_message(embed=build_prediction_explorer_embed(view.group,view.predictions,view.index),view=view,ephemeral=True)

    @discord.ui.button(label="Resumen", emoji="🎯", style=discord.ButtonStyle.primary, row=0)
    async def summary(self, interaction, button):
        ps=get_real_group()
        es=events()
        hit=almost=miss=pending=0
        for p in ps:
            st,_,_,_=evaluate_real(p,es)
            if st=="ACERTADA": hit+=1
            elif st=="CASI ACERTADA": almost+=1
            elif st=="NO ACERTADA": miss+=1
            else: pending+=1
        e=ui_embed("🎯 Predicciones")
        e.add_field(name="Total",value=f"**{len(ps)}**",inline=True)
        e.add_field(name="✅ Acertadas",value=f"**{hit}**",inline=True)
        e.add_field(name="🟠 Casi acertadas",value=f"**{almost}**",inline=True)
        e.add_field(name="❌ No acertadas",value=f"**{miss}**",inline=True)
        e.add_field(name="🟡 Pendientes",value=f"**{pending}**",inline=True)
        await interaction.response.send_message(embed=e,ephemeral=True)

    @discord.ui.button(label="Ranking", emoji="🏆", style=discord.ButtonStyle.secondary, row=0)
    async def ranking(self, interaction, button):
        await interaction.response.send_message(embed=build_ranking_embed(),ephemeral=True)

    @discord.ui.button(label="Historial", emoji="📑", style=discord.ButtonStyle.secondary, row=0)
    async def history(self, interaction, button):
        await interaction.response.send_message(embed=build_history_embed(),ephemeral=True)

    @discord.ui.button(label="Comparar IAs", emoji="⚔️", style=discord.ButtonStyle.success, row=1)
    async def compare_ai(self, interaction, button):
        groups=list(_prediction_groups())
        if len(groups)<2: await interaction.response.send_message("Necesito al menos dos IAs/grupos importados.",ephemeral=True); return
        view=AICompareView(); await interaction.response.send_message(embed=build_comparison_embed(view.a,view.b),view=view,ephemeral=True)

    @discord.ui.button(label="Actividad", emoji="📡", style=discord.ButtonStyle.secondary, row=1)
    async def activity(self, interaction, button): await interaction.response.send_message(embed=build_activity_embed(),ephemeral=True)

    @discord.ui.button(label="Evolución", emoji="📈", style=discord.ButtonStyle.secondary, row=2)
    async def evolution(self, interaction, button):
        save_ranking_snapshot(); await interaction.response.send_message(embed=build_evolution_embed(),ephemeral=True)

    @discord.ui.button(label="Mapa", emoji="🗺️", style=discord.ButtonStyle.secondary, row=2)
    async def map_view(self, interaction, button):
        await interaction.response.defer(ephemeral=True); data,n=build_prediction_map()
        if not data: await interaction.followup.send("No hay suficientes datos para el mapa.",ephemeral=True); return
        e=ui_embed("🗺️ Predicción vs realidad",f"**{n}** pares · 🔴 predicción · 🔵 CSN"); e.set_image(url="attachment://mapa_predicciones.png")
        await interaction.followup.send(embed=e,file=discord.File(BytesIO(data),filename="mapa_predicciones.png"),ephemeral=True)

    @discord.ui.button(label="Integridad", emoji="🔒", style=discord.ButtonStyle.secondary, row=1)
    async def integrity(self, interaction, button):
        ps=get_real_group()
        ok=changed=unfrozen=0
        for p in ps:
            st,_=integrity_status(p)
            if st=="OK": ok+=1
            elif st=="MODIFICADA": changed+=1
            else: unfrozen+=1
        e=ui_embed("🔒 Integridad de predicciones")
        e.add_field(name="✅ Sin cambios",value=str(ok),inline=True)
        e.add_field(name="⚠️ Modificadas",value=str(changed),inline=True)
        e.add_field(name="🧊 Sin congelar",value=str(unfrozen),inline=True)
        await interaction.response.send_message(embed=e,ephemeral=True)

def build_csn_browser_embed(es,index=0):
    e=ui_embed("🌎 Explorador CSN","Navega por los sismos guardados sin llenar el canal de mensajes.")
    if not es:
        e.description="No hay eventos CSN guardados todavía."
        return e
    index=max(0,min(index,len(es)-1)); x=es[index]
    try: when=parse_datetime(x["occurred_at"]).strftime("%d/%m/%Y %H:%M:%S")
    except Exception: when=str(x["occurred_at"])
    e.add_field(name=f"🌎 Sismo {index+1}/{len(es)} · M{x['magnitude']:.1f}",value=(
        f"📅 **{when}**\n📍 {x['place'] or 'Sin referencia geográfica'}\n"
        f"🌐 `{x['latitude']:.4f}, {x['longitude']:.4f}`\n"
        f"⬇️ Profundidad: **{x['depth_km']:.1f} km**" if x['depth_km'] is not None else
        f"📅 **{when}**\n📍 {x['place'] or 'Sin referencia geográfica'}\n🌐 `{x['latitude']:.4f}, {x['longitude']:.4f}`"),inline=False)
    if x["source_url"]: e.add_field(name="🔗 Fuente",value=f"[Abrir informe oficial CSN]({x['source_url']})",inline=False)
    e.set_footer(text="◀️/▶️ para navegar · 🔄 para volver al más reciente")
    return e

class CSNBrowser(discord.ui.View):
    def __init__(self,es,index=0):
        super().__init__(timeout=900); self.es=es; self.index=index
        disabled=len(es)<=1; self.previous.disabled=disabled; self.next.disabled=disabled
    async def _render(self,interaction):
        await interaction.response.edit_message(embed=build_csn_browser_embed(self.es,self.index),view=self)
    @discord.ui.button(label="Anterior",emoji="◀️",style=discord.ButtonStyle.secondary)
    async def previous(self,interaction,button):
        if self.es:self.index=(self.index-1)%len(self.es)
        await self._render(interaction)
    @discord.ui.button(label="Siguiente",emoji="▶️",style=discord.ButtonStyle.primary)
    async def next(self,interaction,button):
        if self.es:self.index=(self.index+1)%len(self.es)
        await self._render(interaction)
    @discord.ui.button(label="Más reciente",emoji="🔄",style=discord.ButtonStyle.success)
    async def newest(self,interaction,button):
        self.index=0; await self._render(interaction)

class CSNPanel(discord.ui.View):
    def __init__(self):
        super().__init__(timeout=900)

    @discord.ui.button(label="Explorar sismos", emoji="🌎", style=discord.ButtonStyle.primary, row=0)
    async def latest(self, interaction, button):
        es=events()
        await interaction.response.send_message(embed=build_csn_browser_embed(es,0),view=CSNBrowser(es,0),ephemeral=True)

    @discord.ui.button(label="Actualizar ahora", emoji="🔄", style=discord.ButtonStyle.success, row=0)
    async def update_now(self, interaction, button):
        await interaction.response.defer(ephemeral=True)
        try:
            total,new,errs=update_csn()
            e=ui_embed("🌐 CSN actualizado",color=discord.Color.green())
            e.add_field(name="Leídos",value=str(total),inline=True)
            e.add_field(name="Nuevos",value=str(new),inline=True)
            e.add_field(name="Avisos",value=str(len(errs)),inline=True)
            await interaction.followup.send(embed=e,ephemeral=True)
        except Exception as exc:
            await interaction.followup.send(f"❌ Error consultando CSN: `{exc}`",ephemeral=True)

    @discord.ui.button(label="Estado CSN", emoji="📡", style=discord.ButtonStyle.secondary, row=0)
    async def csn_status(self, interaction, button):
        with connect() as con:
            count=con.execute("SELECT COUNT(*) FROM observed_events").fetchone()[0]
        e=ui_embed("📡 Estado CSN")
        e.add_field(name="Monitor",value=f"🟢 Cada **{CHECK_MINUTES} min**",inline=True)
        e.add_field(name="Eventos guardados",value=f"**{count}**",inline=True)
        await interaction.response.send_message(embed=e,ephemeral=True)

class AlertsPanel(discord.ui.View):
    def __init__(self):
        super().__init__(timeout=900)

    @discord.ui.button(label="Estado", emoji="🔔", style=discord.ButtonStyle.primary, row=0)
    async def status(self, interaction, button):
        channel_id=get_setting("alert_channel_id")
        e=ui_embed("🔔 Alertas")
        e.add_field(name="Canal",value=(f"<#{channel_id}>" if channel_id else "❌ No configurado"),inline=False)
        e.add_field(name="Coincidencia 3/3",value="🚨 Notificación prioritaria",inline=True)
        e.add_field(name="Casi 2/3",value="🟡 Aviso normal",inline=True)
        await interaction.response.send_message(embed=e,ephemeral=True)

class DiagnosticsPanel(discord.ui.View):
    def __init__(self):
        super().__init__(timeout=900)

    @discord.ui.button(label="Ejecutar diagnóstico", emoji="🩺", style=discord.ButtonStyle.success)
    async def run_diag(self, interaction, button):
        checks=[]
        try:
            with connect() as con:
                con.execute("SELECT 1").fetchone()
            checks.append("✅ SQLite")
        except Exception as exc:
            checks.append(f"❌ SQLite: {exc}")
        checks.append("✅ Discord conectado" if bot.is_ready() else "⚠️ Discord iniciando")
        checks.append(f"{'✅' if Path(DB_PATH).parent.exists() else '❌'} Ruta de datos: `{Path(DB_PATH).parent}`")
        try:
            with connect() as con:
                n=con.execute("SELECT COUNT(*) FROM observed_events").fetchone()[0]
            checks.append(f"✅ Eventos CSN en base: **{n}**")
        except Exception as exc:
            checks.append(f"❌ Eventos CSN: {exc}")
        e=ui_embed("🩺 Diagnóstico del sistema","\n".join(checks))
        await interaction.response.send_message(embed=e,ephemeral=True)

class ImportConfirmView(discord.ui.View):
    def __init__(self):
        super().__init__(timeout=60)

    @discord.ui.button(label="Sí, importar", emoji="✅", style=discord.ButtonStyle.success)
    async def confirm(self, interaction, button):
        e=ui_embed(
            "📥 Importación",
            "Para conservar exactamente la lógica y validaciones actuales, ejecuta **`/importar_todos`**.\n\n"
            "El panel no duplica esa rutina: así evitamos que una actualización de interfaz cambie accidentalmente tus datos.",
            discord.Color.green()
        )
        await interaction.response.edit_message(embed=e,view=None)

    @discord.ui.button(label="Cancelar", emoji="❌", style=discord.ButtonStyle.secondary)
    async def cancel(self, interaction, button):
        await interaction.response.edit_message(
            embed=ui_embed("Importación cancelada","No se modificó ninguna predicción."),
            view=None
        )

class MainPanel(discord.ui.View):
    def __init__(self):
        super().__init__(timeout=900)

    @discord.ui.button(label="Predicciones", emoji="🎯", style=discord.ButtonStyle.primary, row=0)
    async def predictions(self,interaction,button):
        await interaction.response.send_message(
            embed=ui_embed("🎯 Centro de predicciones","Explora predicciones por IA, revisa resultados, ranking, historial e integridad."),
            view=PredictionsPanel(),ephemeral=True)

    @discord.ui.button(label="Sismos CSN", emoji="🌎", style=discord.ButtonStyle.primary, row=0)
    async def csn(self,interaction,button):
        await interaction.response.send_message(
            embed=ui_embed("🌎 Centro CSN","Consulta los datos guardados o actualiza el catálogo."),
            view=CSNPanel(),ephemeral=True)

    @discord.ui.button(label="Importar", emoji="📥", style=discord.ButtonStyle.secondary, row=0)
    async def importing(self,interaction,button):
        await interaction.response.send_message(
            embed=ui_embed("📥 Importar predicciones","¿Quieres abrir la importación del Excel?"),
            view=ImportConfirmView(),ephemeral=True)

    @discord.ui.button(label="Resultados", emoji="📊", style=discord.ButtonStyle.secondary, row=0)
    async def results(self,interaction,button):
        await interaction.response.send_message(embed=build_ranking_embed(),ephemeral=True)

    @discord.ui.button(label="Alertas", emoji="🔔", style=discord.ButtonStyle.secondary, row=0)
    async def alerts(self,interaction,button):
        await interaction.response.send_message(
            embed=ui_embed("🔔 Centro de alertas","Consulta la configuración actual de notificaciones."),
            view=AlertsPanel(),ephemeral=True)

    @discord.ui.button(label="Calculadora", emoji="🧮", style=discord.ButtonStyle.success, row=1)
    async def calculator(self,interaction,button):
        await interaction.response.send_message(
            embed=ui_embed("🧮 Centro de cálculo","Normal, raíz, distancia, magnitud, tiempo y comparación 3/3."),
            view=CalculatorView(),ephemeral=True)

    @discord.ui.button(label="Diagnóstico", emoji="🩺", style=discord.ButtonStyle.secondary, row=1)
    async def diagnostics(self,interaction,button):
        await interaction.response.send_message(
            embed=ui_embed("🩺 Diagnóstico","Comprueba rápidamente los componentes principales."),
            view=DiagnosticsPanel(),ephemeral=True)

    @discord.ui.button(label="Estado", emoji="⚙️", style=discord.ButtonStyle.secondary, row=1)
    async def status(self,interaction,button):
        await interaction.response.send_message(embed=build_status_embed(),ephemeral=True)

    @discord.ui.button(label="Historial", emoji="📑", style=discord.ButtonStyle.secondary, row=1)
    async def history(self,interaction,button):
        await interaction.response.send_message(embed=build_history_embed(),ephemeral=True)

    @discord.ui.button(label="Actividad", emoji="📡", style=discord.ButtonStyle.success, row=2)
    async def activity(self,interaction,button):
        await interaction.response.send_message(embed=build_activity_embed(),ephemeral=True)

    @discord.ui.button(label="Investigación", emoji="🧪", style=discord.ButtonStyle.success, row=2)
    async def research(self,interaction,button):
        save_ranking_snapshot(); await interaction.response.send_message(embed=build_research_embed(),view=ResearchView(),ephemeral=True)

    @discord.ui.button(label="Ayuda", emoji="❓", style=discord.ButtonStyle.secondary, row=1)
    async def help(self,interaction,button):
        await interaction.response.send_message(embed=build_help_embed(),ephemeral=True)


@bot.tree.command(name="mapa",description="Mapa de predicciones frente a sus sismos CSN más cercanos")
@app_commands.describe(ia="IA/grupo opcional; vacío = todas")
async def c_mapa(i:discord.Interaction,ia:str=""):
    await i.response.defer(ephemeral=True)
    group=ia.strip() or None
    if group and group not in _prediction_groups():
        await i.followup.send(f"❌ No encontré el grupo/IA `{group}`.",ephemeral=True); return
    data,n=build_prediction_map(group)
    if not data:
        await i.followup.send("Todavía no hay suficientes predicciones y eventos CSN para construir el mapa.",ephemeral=True); return
    title=f"🗺️ Mapa · {group}" if group else "🗺️ Mapa · todas las IAs"
    e=ui_embed(title,f"**{n}** predicciones comparadas con su mejor candidato CSN.\n🔴 Predicción · 🔵 Sismo CSN · línea = separación geográfica")
    e.set_image(url="attachment://mapa_predicciones.png")
    await i.followup.send(embed=e,file=discord.File(BytesIO(data),filename="mapa_predicciones.png"),ephemeral=True)

@bot.tree.command(name="evolucion",description="Muestra cómo ha cambiado el ranking desde esta actualización")
async def c_evolucion(i:discord.Interaction):
    save_ranking_snapshot(); await i.response.send_message(embed=build_evolution_embed(),ephemeral=True)

class ResearchView(discord.ui.View):
    def __init__(self): super().__init__(timeout=900)
    @discord.ui.button(label="Ranking",emoji="🏆",style=discord.ButtonStyle.primary,row=0)
    async def ranking(self,interaction,button): await interaction.response.send_message(embed=build_ranking_embed(),ephemeral=True)
    @discord.ui.button(label="Comparar IAs",emoji="⚔️",style=discord.ButtonStyle.success,row=0)
    async def compare(self,interaction,button):
        if len(_prediction_groups())<2: await interaction.response.send_message("Necesito al menos dos IAs/grupos.",ephemeral=True); return
        v=AICompareView(); await interaction.response.send_message(embed=build_comparison_embed(v.a,v.b),view=v,ephemeral=True)
    @discord.ui.button(label="Mapa",emoji="🗺️",style=discord.ButtonStyle.secondary,row=0)
    async def map(self,interaction,button):
        await interaction.response.defer(ephemeral=True); data,n=build_prediction_map()
        if not data: await interaction.followup.send("No hay suficientes datos para el mapa.",ephemeral=True); return
        e=ui_embed("🗺️ Predicción vs realidad",f"**{n}** pares · 🔴 predicción · 🔵 CSN"); e.set_image(url="attachment://mapa_predicciones.png")
        await interaction.followup.send(embed=e,file=discord.File(BytesIO(data),filename="mapa_predicciones.png"),ephemeral=True)
    @discord.ui.button(label="Evolución",emoji="📈",style=discord.ButtonStyle.secondary,row=1)
    async def evolution(self,interaction,button): save_ranking_snapshot(); await interaction.response.send_message(embed=build_evolution_embed(),ephemeral=True)
    @discord.ui.button(label="Actividad",emoji="📡",style=discord.ButtonStyle.secondary,row=1)
    async def activity(self,interaction,button): await interaction.response.send_message(embed=build_activity_embed(),ephemeral=True)

@bot.tree.command(name="investigacion",description="Abre la vista de resultados para el proyecto de investigación")
async def c_investigacion(i:discord.Interaction):
    save_ranking_snapshot(); await i.response.send_message(embed=build_research_embed(),view=ResearchView(),ephemeral=True)

@bot.tree.command(name="panel",description="Abre el centro de control visual del Bot Sísmico")
async def c_panel(i:discord.Interaction):
    ps=get_real_group()
    with connect() as con:
        n_events=con.execute("SELECT COUNT(*) FROM observed_events").fetchone()[0]
        n_corr=con.execute("SELECT COUNT(*) FROM correlation_history").fetchone()[0]
    channel_id=get_setting("alert_channel_id")
    e=ui_embed(
        "🌎 BOT SÍSMICO · CENTRO DE CONTROL",
        "Panel principal para revisar el experimento sin memorizar todos los comandos."
    )
    e.add_field(name="🟢 Sistema",value="Online",inline=True)
    e.add_field(name="🎯 Predicciones",value=f"**{len(ps)}**",inline=True)
    e.add_field(name="🌎 Sismos CSN",value=f"**{n_events}**",inline=True)
    e.add_field(name="📑 Correlaciones",value=f"**{n_corr}**",inline=True)
    e.add_field(name="🔔 Alertas",value=("Configuradas" if channel_id else "Sin canal"),inline=True)
    e.add_field(name="📡 Monitor",value=f"Cada **{CHECK_MINUTES} min**",inline=True)
    e.set_footer(text="UI/QoL Update • Sismologia Lab")
    await i.response.send_message(embed=e,view=MainPanel())

@bot.tree.command(name="ayuda",description="Muestra una guía limpia de los comandos del bot")
async def c_ayuda(i:discord.Interaction):
    await i.response.send_message(embed=build_help_embed(),ephemeral=True)

@bot.tree.command(name="raiz",description="Calcula una raíz de forma directa")
@app_commands.describe(numero="Número",indice="2=cuadrada, 3=cúbica, etc.")
async def c_raiz(i:discord.Interaction,numero:float,indice:float=2.0):
    if indice==0:
        await i.response.send_message("❌ El índice no puede ser 0.",ephemeral=True)
        return
    try:
        if numero<0 and abs(indice-round(indice))<1e-9 and int(round(indice))%2==1:
            result=-((-numero)**(1/indice))
        elif numero<0:
            raise ValueError("esa raíz no tiene resultado real")
        else:
            result=numero**(1/indice)
        e=ui_embed("√ Resultado",f"Raíz **{indice:g}** de **{numero:g}**\n\n### = **{result:.10g}**",discord.Color.green())
        await i.response.send_message(embed=e)
    except Exception as exc:
        await i.response.send_message(f"❌ No pude calcularla: `{exc}`",ephemeral=True)


init_db()
if __name__=="__main__":
    if not TOKEN:raise SystemExit("Configura DISCORD_TOKEN.")
    threading.Thread(target=run_dashboard, daemon=True, name="sismologia-lab-web").start()
    print("Sismologia Lab web iniciado en PORT", os.getenv("PORT","8080"))
    
@bot.tree.command(name="score",description="Muestra el porcentaje de proximidad de una predicción al mejor sismo")
@app_commands.describe(codigo="Código de predicción, por ejemplo GROK_03")
async def c_score(i:discord.Interaction,codigo:str):
    await i.response.defer(ephemeral=True)
    try:
        ps=[p for p in get_real_group() if str(p["code"]).upper()==codigo.upper()]
        if not ps:
            await i.followup.send(f"❌ No encontré `{codigo}`.",ephemeral=True)
            return
        p=ps[0]
        es=events()
        status, matched, strict_radius, candidates = evaluate_real(p,es)
        candidate, sc=best_proximity(p,es,DEFAULT_MARGIN_HOURS)
        if not sc:
            await i.followup.send("No hay sismos guardados para calcular el score.",ephemeral=True)
            return
        strict="✅ ACERTADA" if status=="ACERTADA" else ("🟠 CASI ACERTADA" if status=="CASI ACERTADA" else ("❌ NO ACERTADA" if status=="NO ACERTADA" else "🟡 PENDIENTE"))
        e=candidate
        dt=parse_datetime(e["occurred_at"])
        emb=discord.Embed(title=f"🎯 Score · {p['code']}",description=f"**{sc['score']:.1f}%** de proximidad\nResultado (3/3 o score ≥95%): **{strict}**",color=discord.Color.blurple())
        emb.add_field(name="📍 Ubicación · 40%",value=f"**{sc['spatial_pct']:.1f}%**\nDistancia: {sc['distance_km']:.2f} km / radio estricto {strict_radius:.1f} km",inline=True)
        emb.add_field(name="📈 Magnitud · 30%",value=f"**{sc['magnitude_pct']:.1f}%**\nReal: M{float(e['magnitude']):.1f} · Pred.: {float(p['mag_min']):.1f}–{float(p['mag_max']):.1f}",inline=True)
        emb.add_field(name="⏱️ Tiempo · 30%",value=f"**{sc['temporal_pct']:.1f}%**\nEvento: {dt:%d/%m/%Y %H:%M:%S}",inline=True)
        emb.add_field(name="🌎 Sismo más próximo al criterio",value=e["place"],inline=False)
        if e["source_url"]:
            emb.add_field(name="🔗 Informe oficial del sismo",value=f"[Ver sismo en CSN]({e['source_url']})",inline=False)
        else:
            emb.add_field(name="🔗 Informe oficial del sismo",value="Este sismo no tiene un enlace CSN guardado.",inline=False)
        emb.set_footer(text="Resultado: coincidencia 3/3 o score global ≥95%. Evento 3/3 priorizado cuando existe.")
        await i.followup.send(embed=emb)
    
    except Exception as exc:
        print(f"ERROR /score {codigo}: {type(exc).__name__}: {exc}")
        await i.followup.send(f"❌ No pude calcular el score de `{codigo}`. Error: `{type(exc).__name__}: {str(exc)[:300]}`", ephemeral=True)

# ---------- v1.6.1: ESTADO DEL SISTEMA ----------

def _human_bytes(n):
    n=float(n or 0)
    for unit in ("B","KB","MB","GB","TB"):
        if n < 1024 or unit=="TB":
            return f"{n:.1f} {unit}"
        n /= 1024

def _uptime_text():
    sec=max(0,int(pytime.monotonic()-BOT_STARTED_AT))
    days,sec=divmod(sec,86400)
    hours,sec=divmod(sec,3600)
    mins,_=divmod(sec,60)
    if days: return f"{days}d {hours}h {mins}m"
    if hours: return f"{hours}h {mins}m"
    return f"{mins}m"

def _system_snapshot():
    # ru_maxrss is KB on Linux (Railway).
    rss_kb=resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    ram_bytes=rss_kb*1024

    db_size=0
    try:
        db_size=os.path.getsize(DB_PATH)
    except Exception:
        pass

    try:
        ps_count=len(get_real_group())
    except Exception:
        ps_count=0
    try:
        ev_count=len(events())
    except Exception:
        ev_count=0

    # Load average is more reliable/cheap than inventing a CPU percentage without psutil.
    try:
        load1,load5,load15=os.getloadavg()
        cpu_txt=f"Load: {load1:.2f} / {load5:.2f} / {load15:.2f}"
    except Exception:
        cpu_txt="No disponible"

    # Basic CSN freshness check from the DB, without making an extra web request.
    newest=None
    try:
        evs=events()
        newest=evs[0] if evs else None
    except Exception:
        newest=None

    latest_txt="Sin eventos"
    monitor_state="🟡 Sin datos"
    if newest:
        try:
            dt=parse_datetime(newest["occurred_at"])
            latest_txt=f"{dt:%d/%m/%Y %H:%M:%S} · M{float(newest['magnitude']):.1f}"
            monitor_state="🟢 Con datos"
        except Exception:
            latest_txt=str(newest.get("occurred_at","—"))

    return {
        "ram": _human_bytes(ram_bytes),
        "db": _human_bytes(db_size),
        "predictions": ps_count,
        "events": ev_count,
        "cpu": cpu_txt,
        "uptime": _uptime_text(),
        "latest": latest_txt,
        "monitor": monitor_state,
    }

@bot.tree.command(name="estado_sistema",description="Muestra salud, RAM, base de datos y estado general del bot")
async def c_system_status(i:discord.Interaction):
    snap=_system_snapshot()

    # Conservative health label: process alive + DB accessible is enough for green;
    # CSN freshness remains separately visible instead of claiming web sync.
    general="🟢 SALUDABLE"
    try:
        with connect() as con:
            con.execute("SELECT 1").fetchone()
    except Exception:
        general="🔴 PROBLEMA CON SQLITE"

    embed=discord.Embed(
        title="🖥️ Estado del sistema",
        description=f"Estado general: **{general}**",
        color=discord.Color.green() if general.startswith("🟢") else discord.Color.red()
    )
    embed.add_field(name="⏱️ Uptime",value=snap["uptime"],inline=True)
    embed.add_field(name="🧠 RAM máx.",value=snap["ram"],inline=True)
    embed.add_field(name="⚙️ CPU",value=snap["cpu"],inline=True)
    embed.add_field(name="💾 SQLite",value=snap["db"],inline=True)
    embed.add_field(name="🎯 Predicciones",value=str(snap["predictions"]),inline=True)
    embed.add_field(name="🌎 Sismos guardados",value=str(snap["events"]),inline=True)
    embed.add_field(name="📡 Monitor CSN",value=snap["monitor"],inline=True)
    embed.add_field(name="🌐 Dashboard",value="🟢 Proceso integrado" if 'dashboard' else "—",inline=True)
    embed.add_field(name="🕒 Último sismo guardado",value=snap["latest"],inline=False)
    embed.set_footer(text="CPU se muestra como carga del sistema (1/5/15 min), no como porcentaje.")
    await i.response.send_message(embed=embed,ephemeral=True)


bot.run(TOKEN)
