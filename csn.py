import re
from datetime import datetime, timedelta
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup

BASE_URL = "https://www.sismologia.cl/"
HEADERS = {"User-Agent": "BotSismicoEducativo/0.6 (+Discord; lectura respetuosa CSN)"}
REPORT_RE = re.compile(r"/sismicidad/informes/\d{4}/\d{2}/(\d+)\.html", re.I)

def _num(text):
    if text is None:
        return None
    m = re.search(r"-?\d+(?:[.,]\d+)?", str(text))
    return float(m.group().replace(",", ".")) if m else None

def catalog_url_for_day(day):
    return f"{BASE_URL}sismicidad/catalogo/{day:%Y}/{day:%m}/{day:%Y%m%d}.html"

def _parse_catalog_html(html, page_url):
    """Lee directamente las filas del catálogo diario del CSN."""
    soup = BeautifulSoup(html, "html.parser")
    events, errors = [], []

    for tr in soup.find_all("tr"):
        cells = tr.find_all(["td", "th"])
        if len(cells) < 5:
            continue

        # La tabla visible tiene:
        # Fecha Local/Lugar | Fecha UTC | Latitud/Longitud | Profundidad | Magnitud
        c0 = " ".join(cells[0].stripped_strings)
        c1 = " ".join(cells[1].stripped_strings)
        c2 = " ".join(cells[2].stripped_strings)
        c3 = " ".join(cells[3].stripped_strings)
        c4 = " ".join(cells[4].stripped_strings)

        dtm = re.search(r"(20\d{2}-\d{2}-\d{2})\s+(\d{2}:\d{2}:\d{2})", c0)
        if not dtm:
            continue  # encabezado u otra fila

        try:
            local_dt = datetime.strptime(
                f"{dtm.group(1)} {dtm.group(2)}", "%Y-%m-%d %H:%M:%S"
            )

            # Lugar = texto restante de la primera celda.
            place = re.sub(
                r"20\d{2}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}", "", c0, count=1
            ).strip()

            coords = re.findall(r"-?\d+(?:[.,]\d+)?", c2)
            if len(coords) < 2:
                raise ValueError(f"coordenadas no reconocidas: {c2!r}")
            lat = float(coords[0].replace(",", "."))
            lon = float(coords[1].replace(",", "."))

            depth = _num(c3)
            mag = _num(c4)
            if mag is None:
                raise ValueError(f"magnitud no reconocida: {c4!r}")

            a = cells[0].find("a", href=True)
            report_url = urljoin(page_url, a["href"]) if a else page_url
            mid = REPORT_RE.search(report_url)

            # Preferimos ID oficial del informe. Si no existe, generamos una clave
            # estable con fecha/hora+coordenadas para evitar duplicados.
            source_id = (
                mid.group(1)
                if mid
                else f"{local_dt:%Y%m%d%H%M%S}_{lat:.3f}_{lon:.3f}"
            )

            utc_dt = None
            um = re.search(r"(20\d{2}-\d{2}-\d{2})\s+(\d{2}:\d{2}:\d{2})", c1)
            if um:
                utc_dt = datetime.strptime(
                    f"{um.group(1)} {um.group(2)}", "%Y-%m-%d %H:%M:%S"
                )

            events.append({
                "source": "CSN",
                "source_id": source_id,
                # El motor compara con el Excel usando HORA LOCAL.
                "occurred_at": local_dt,
                "occurred_at_utc": utc_dt,
                "place": place,
                "magnitude": mag,
                "latitude": lat,
                "longitude": lon,
                "depth_km": depth,
                "url": report_url,
            })
        except Exception as exc:
            errors.append(f"Fila {c0[:45]!r}: {exc}")

    return events, errors

def fetch_catalog_day(day):
    url = catalog_url_for_day(day)
    r = requests.get(url, headers=HEADERS, timeout=20)
    if r.status_code == 404:
        return [], [f"Sin catálogo publicado: {day:%d/%m/%Y}"]
    r.raise_for_status()
    return _parse_catalog_html(r.text, url)

def fetch_historical_events(start_day, end_day):
    if end_day < start_day:
        raise ValueError("La fecha final no puede ser anterior a la inicial.")
    events, errors = [], []
    day = start_day
    while day <= end_day:
        try:
            ev, er = fetch_catalog_day(day)
            events.extend(ev)
            errors.extend(er)
        except Exception as exc:
            errors.append(f"{day:%d/%m/%Y}: {exc}")
        day += timedelta(days=1)
    return events, errors

def fetch_recent_events(limit=30):
    """Para tiempo real lee el catálogo de hoy y, por seguridad, el de ayer."""
    today = datetime.now().date()
    events, errors = fetch_historical_events(today - timedelta(days=1), today)
    events.sort(key=lambda e: e["occurred_at"], reverse=True)
    return events[:limit], errors
