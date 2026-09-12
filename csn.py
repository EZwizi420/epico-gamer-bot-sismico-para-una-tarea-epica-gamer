import re
import time
import requests
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from urllib.parse import urljoin

BASE = "https://www.sismologia.cl"
TZ = ZoneInfo("America/Santiago")

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/152.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "es-CL,es;q=0.9,en;q=0.7",
    "Cache-Control": "no-cache",
    "Pragma": "no-cache",
}

def _daily_url(day):
    return f"{BASE}/sismicidad/catalogo/{day:%Y/%m}/{day:%Y%m%d}.html"

def _fresh_get(url, timeout=25):
    # Query string + no-cache headers prevent stale CDN/proxy responses.
    sep = "&" if "?" in url else "?"
    fresh_url = f"{url}{sep}_bot_ts={time.time_ns()}"
    r = requests.get(fresh_url, headers=HEADERS, timeout=timeout)
    r.raise_for_status()
    return r

def _num(txt):
    m = re.search(r"-?\d+(?:[.,]\d+)?", txt or "")
    return float(m.group(0).replace(",", ".")) if m else None

def _parse_datetime(txt):
    m = re.search(r"(\d{4}-\d{2}-\d{2})\s+(\d{2}:\d{2}:\d{2})", txt or "")
    if not m:
        return None
    return datetime.strptime(f"{m.group(1)} {m.group(2)}", "%Y-%m-%d %H:%M:%S").replace(tzinfo=TZ)

def fetch_day(day):
    url = _daily_url(day)
    html = _fresh_get(url).text
    soup = BeautifulSoup(html, "html.parser")
    events = []

    # The official daily catalogue is a table. Parse rows instead of relying on
    # report-detail pages, which may lag or be cached independently.
    for tr in soup.find_all("tr"):
        cells = tr.find_all("td")
        if len(cells) < 5:
            continue

        local_cell = cells[0]
        local_text = " ".join(local_cell.stripped_strings)
        occurred = _parse_datetime(local_text)
        if not occurred:
            continue

        # Place is the text in the first cell after the date/time link/text.
        place = local_text
        place = re.sub(r"^\s*\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}\s*", "", place).strip()

        coord_text = " ".join(cells[2].stripped_strings)
        nums = re.findall(r"-?\d+(?:[.,]\d+)?", coord_text)
        if len(nums) < 2:
            continue
        lat = float(nums[0].replace(",", "."))
        lon = float(nums[1].replace(",", "."))

        depth = _num(" ".join(cells[3].stripped_strings))
        mag = _num(" ".join(cells[4].stripped_strings))
        if mag is None:
            continue

        a = local_cell.find("a", href=True)
        source_url = urljoin(BASE, a["href"]) if a else url

        # Stable ID from the actual event fields. This avoids duplicates even if
        # the CSN changes the report URL.
        source_id = f"CSN:{occurred:%Y%m%d%H%M%S}:{lat:.4f}:{lon:.4f}"

        events.append({
            "source_id": source_id,
            "occurred_at": occurred,
            "place": place or "Sin referencia",
            "latitude": lat,
            "longitude": lon,
            "depth_km": depth,
            "magnitude": mag,
            "source_url": source_url,
            "url": source_url,
        })

    # Newest first and unique by source_id.
    unique = {e["source_id"]: e for e in events}
    return sorted(unique.values(), key=lambda e: e["occurred_at"], reverse=True)

def fetch_historical_events(start_day, end_day):
    """Compatibilidad con /importar_historico: devuelve (eventos, errores)."""
    if end_day < start_day:
        raise ValueError("La fecha final no puede ser anterior a la inicial.")
    events, errors = [], []
    day = start_day
    while day <= end_day:
        try:
            events.extend(fetch_day(day))
        except requests.HTTPError as exc:
            status = getattr(exc.response, "status_code", None)
            if status == 404:
                errors.append(f"Sin catálogo publicado: {day:%d/%m/%Y}")
            else:
                errors.append(f"{day:%d/%m/%Y}: HTTP {status or '?'}")
        except Exception as exc:
            errors.append(f"{day:%d/%m/%Y}: {type(exc).__name__}: {exc}")
        day += timedelta(days=1)

    unique = {e["source_id"]: e for e in events}
    ordered = sorted(unique.values(), key=lambda e: e["occurred_at"], reverse=True)
    return ordered, errors

def _candidate_catalog_days():
    """
    Consulta un margen alrededor de las fechas de Chile y UTC.
    Evita perder eventos al cruzar medianoche UTC/local o por publicación tardía.
    """
    now_cl = datetime.now(TZ)
    now_utc = datetime.now(ZoneInfo("UTC"))
    bases = {now_cl.date(), now_utc.date()}
    days = set()
    for base in bases:
        for offset in (-2, -1, 0, 1):
            days.add(base + timedelta(days=offset))
    return sorted(days)

def fetch_recent_events(limit=None):
    """Lee varios catálogos alrededor de la fecha Chile/UTC y devuelve (eventos, errores)."""
    merged={}
    errors=[]
    for day in _candidate_catalog_days():
        try:
            for e in fetch_day(day):
                merged[e["source_id"]]=e
        except requests.HTTPError as exc:
            status=getattr(exc.response,"status_code",None)
            # A page for tomorrow may legitimately not exist yet.
            if status not in (404,):
                errors.append(f"{day:%d/%m/%Y}: HTTP {status or '?'}")
        except Exception as exc:
            errors.append(f"{day:%d/%m/%Y}: {type(exc).__name__}: {exc}")

    events=sorted(merged.values(), key=lambda e:e["occurred_at"], reverse=True)
    return (events[:limit] if limit is not None else events), errors

def fetch_recent_events_debug():
    days=_candidate_catalog_days()
    counts={}
    merged={}
    errors=[]
    for day in days:
        try:
            evs=fetch_day(day)
            counts[day.isoformat()]=len(evs)
            for e in evs:
                merged[e["source_id"]]=e
        except requests.HTTPError as exc:
            status=getattr(exc.response,"status_code",None)
            counts[day.isoformat()]=0
            if status not in (404,):
                errors.append(f"{day:%d/%m/%Y}: HTTP {status or '?'}")
        except Exception as exc:
            counts[day.isoformat()]=0
            errors.append(f"{day:%d/%m/%Y}: {type(exc).__name__}: {exc}")

    events=sorted(merged.values(), key=lambda e:e["occurred_at"], reverse=True)
    newest=events[0] if events else None
    now=datetime.now(TZ)
    age=None
    if newest:
        age=max(0,(now-newest["occurred_at"]).total_seconds()/60)

    return {
        "catalog_days": days,
        "counts": counts,
        "total_count": len(events),
        "newest": newest,
        "newest_age_minutes": age,
        "events": events,
        "errors": errors,
    }
