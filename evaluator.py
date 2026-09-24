import math
from datetime import datetime, date, time

def parse_datetime(value):
    if isinstance(value, datetime): return value
    if isinstance(value, str): return datetime.fromisoformat(value)
    raise TypeError(f"fecha/hora inválida: {type(value).__name__}")

def haversine(a,b,c,d):
    R=6371.0088; p1,p2=math.radians(a),math.radians(c)
    dp=math.radians(c-a); dl=math.radians(d-b)
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

def _window_bounds(p):
    start=datetime.combine(date.fromisoformat(p["date_start"]),time.min)
    end=datetime.combine(date.fromisoformat(p["date_end"]),time.max)
    if p["daily_time_start"]:
        start=datetime.combine(start.date(),datetime.strptime(p["daily_time_start"],"%H:%M").time())
    if p["daily_time_end"]:
        end=datetime.combine(end.date(),datetime.strptime(p["daily_time_end"],"%H:%M").time())
    return start,end

def real_time_ok(event_dt,p):
    start,end=_window_bounds(p)
    return start <= event_dt.replace(tzinfo=None) <= end

def real_window_finished(p,now=None):
    now=now or datetime.now()
    return now.replace(tzinfo=None)>_window_bounds(p)[1]


def evaluate_real(p,es,now=None):
    rad=radius((p["mag_min"]+p["mag_max"])/2,p["zone_type"]); matches=[]; candidates=[]
    center=(p["mag_min"]+p["mag_max"])/2
    for e in es:
        when=parse_datetime(e["occurred_at"]); dist=haversine(p["latitude"],p["longitude"],e["latitude"],e["longitude"])
        t_ok=real_time_ok(when,p); m_ok=p["mag_min"]<=e["magnitude"]<=p["mag_max"]; g_ok=dist<=rad
        x={"e":e,"t":t_ok,"m":m_ok,"g":g_ok,"dist":dist}; candidates.append(x)
        if t_ok and m_ok and g_ok:matches.append(x)
    matches.sort(key=lambda x:(x["dist"],abs(x["e"]["magnitude"]-center)))
    # Regla experimental: coincidencia estricta 3/3 O score global >= 95%.
    # Recalcular en cada consulta permite actualizar predicciones antes fallidas
    # sin modificar ni eliminar registros históricos de la base de datos.
    if matches:
        return "ACERTADA", matches[0], rad, candidates
    candidate, score = best_proximity(p, es)
    if score is not None and score["score"] >= 95.0:
        for x in candidates:
            if x["e"]["source_id"] == candidate["source_id"]:
                return "ACERTADA", x, rad, candidates
    # CASI: tiempo y ubicacion cumplen sus margenes, sin exigir magnitud ni score.
    near = [x for x in candidates if x["t"] and x["g"]]
    if near:
        near.sort(key=lambda x: (x["dist"], abs(x["e"]["magnitude"] - center)))
        return "CASI ACERTADA", near[0], rad, candidates
    status = "NO ACERTADA" if real_window_finished(p, now) else "PENDIENTE"
    return status, None, rad, candidates



def _field(row, *names, default=None):
    """Read first available key from dict/sqlite3.Row without relying on .get()."""
    for name in names:
        try:
            return row[name]
        except (KeyError, IndexError, TypeError):
            pass
    return default

# ---------- v1.6: SCORE DE PROXIMIDAD ----------
# Experimental result: a score >= 95% also qualifies as ACERTADA.

def _clamp01(x):
    return max(0.0, min(1.0, float(x)))

def _range_closeness(value, low, high, tolerance):
    """100% inside range; decays linearly outside until tolerance is exhausted."""
    if value is None:
        return 0.0
    low=float(low); high=float(high); value=float(value)
    if low <= value <= high:
        return 1.0
    error = low-value if value < low else value-high
    tol=max(float(tolerance or 0), 1e-9)
    return _clamp01(1.0-error/tol)

def proximity_score(prediction, event, default_margin_hours=2):
    """
    Descriptive 0-100 score. Does not alter the strict 3/3 evaluator.
    Accepts the project's sqlite3.Row prediction schema.
    """
    plat=float(_field(prediction,"lat","latitude"))
    plon=float(_field(prediction,"lon","longitude"))
    elat=float(_field(event,"lat","latitude"))
    elon=float(_field(event,"lon","longitude"))
    distance=haversine(plat,plon,elat,elon)

    mag=float(_field(event,"mag","magnitude"))
    mmin=float(_field(prediction,"mag_min","magnitude_min"))
    mmax=float(_field(prediction,"mag_max","magnitude_max"))

    # Use the exact same radius rule as the official strict evaluator.
    strict_radius=max(float(radius((mmin+mmax)/2, _field(prediction,"zone_type"))),1e-9)
    spatial=_clamp01(1.0-distance/(2.0*strict_radius))
    mag_width=max(mmax-mmin,0.1)
    magnitude=_range_closeness(mag,mmin,mmax,mag_width)

    edt=parse_datetime(_field(event,"occurred_at","datetime","date"))
    margin_h=float(_field(prediction,"margin_hours",default=default_margin_hours) or default_margin_hours)

    # Actual project predictions use date_start/date_end + a daily time window.
    ds=_field(prediction,"date_start")
    de=_field(prediction,"date_end")
    ts=_field(prediction,"daily_time_start")
    te=_field(prediction,"daily_time_end")
    # Some imported predictions have SQL NULL for daily times.
    # Treat missing times as the whole day instead of constructing "...TNone".
    ts = "00:00:00" if ts is None or str(ts).strip() in ("", "None", "nan") else str(ts).strip()
    te = "23:59:59" if te is None or str(te).strip() in ("", "None", "nan") else str(te).strip()
    if ds and de:
        pstart=parse_datetime(f"{ds}T{ts}")
        pend=parse_datetime(f"{de}T{te}")
    else:
        pstart=parse_datetime(_field(prediction,"start_at","start"))
        pend=parse_datetime(_field(prediction,"end_at","end"))

    # Normalize naive project window to the event timezone when needed.
    if edt.tzinfo is not None and pstart.tzinfo is None:
        pstart=pstart.replace(tzinfo=edt.tzinfo)
        pend=pend.replace(tzinfo=edt.tzinfo)

    if pstart <= edt <= pend:
        temporal=1.0
        time_error_h=0.0
    else:
        time_error_h=(pstart-edt).total_seconds()/3600 if edt<pstart else (edt-pend).total_seconds()/3600
        temporal=_clamp01(1.0-time_error_h/max(margin_h,1e-9))

    total=100.0*(0.40*spatial+0.30*magnitude+0.30*temporal)
    return {
        "score": round(total,1),
        "spatial_pct": round(spatial*100,1),
        "magnitude_pct": round(magnitude*100,1),
        "temporal_pct": round(temporal*100,1),
        "distance_km": round(distance,2),
        "radius_km": round(strict_radius,2),
        "time_error_hours": round(time_error_h,2),
    }

def best_proximity(prediction, event_list, default_margin_hours=2):
    if not event_list:
        return None, None
    ranked=[(proximity_score(prediction,e,default_margin_hours),e) for e in event_list]
    ranked.sort(key=lambda x:x[0]["score"],reverse=True)
    return ranked[0][1], ranked[0][0]
