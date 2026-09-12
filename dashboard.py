from flask import Flask, jsonify, render_template_string
import os, sqlite3, math
from datetime import datetime

app = Flask(__name__)
DB_PATH = os.getenv('DB_PATH', 'bot_sismico.db')

def db():
    con=sqlite3.connect(DB_PATH); con.row_factory=sqlite3.Row; return con

def rows(sql,args=()):
    with db() as con: return [dict(r) for r in con.execute(sql,args).fetchall()]

def scalar(sql,args=()):
    with db() as con:
        r=con.execute(sql,args).fetchone(); return r[0] if r else 0

def hav(lat1,lon1,lat2,lon2):
    R=6371.0088; p1,p2=math.radians(lat1),math.radians(lat2)
    dp=math.radians(lat2-lat1); dl=math.radians(lon2-lon1)
    a=math.sin(dp/2)**2+math.cos(p1)*math.cos(p2)*math.sin(dl/2)**2
    return 2*R*math.asin(math.sqrt(a))

def radius(mag, zone):
    z=(zone or '').lower()
    if 'macro' in z: return 80.0
    if 'región amplia' in z or 'region amplia' in z: return 60.0
    if mag < 3.5: return 20.0
    if mag < 4.5: return 30.0
    return 40.0

def evaluate(p, es):
    matches=[]; center=(p['mag_min']+p['mag_max'])/2
    rad=radius(center,p.get('zone_type'))
    ds=datetime.fromisoformat(p['date_start']); de=datetime.fromisoformat(p['date_end'])
    for e in es:
        t=datetime.fromisoformat(e['occurred_at'])
        time_ok=ds.date() <= t.date() <= de.date()
        if p.get('daily_time_start') and p.get('daily_time_end'):
            try:
                a=datetime.strptime(p['daily_time_start'],'%H:%M:%S').time(); b=datetime.strptime(p['daily_time_end'],'%H:%M:%S').time()
                time_ok=time_ok and a <= t.time() <= b
            except Exception: pass
        d=hav(p['latitude'],p['longitude'],e['latitude'],e['longitude'])
        if time_ok and p['mag_min'] <= e['magnitude'] <= p['mag_max'] and d <= rad:
            matches.append((d,abs(e['magnitude']-center),e))
    if matches: return 'ACERTADA', sorted(matches,key=lambda x:(x[0],x[1]))[0][2]
    return ('NO ACERTADA' if datetime.now().date()>de.date() else 'PENDIENTE'), None

@app.get('/api/dashboard')
def api_dashboard():
    es=rows("SELECT * FROM observed_events WHERE source='CSN' ORDER BY occurred_at DESC LIMIT 300")
    ps=rows("SELECT * FROM real_predictions ORDER BY group_name,prediction_no")
    results=[]; groups={}
    for p in ps:
        st,match=evaluate(p,es); results.append({'code':p['code'],'group':p['group_name'],'status':st})
        g=groups.setdefault(p['group_name'],{'group':p['group_name'],'hits':0,'misses':0,'pending':0})
        g['hits' if st=='ACERTADA' else 'misses' if st=='NO ACERTADA' else 'pending']+=1
    for g in groups.values():
        closed=g['hits']+g['misses']; g['accuracy']=round(100*g['hits']/closed,1) if closed else None
    corr=rows("SELECT * FROM correlation_history ORDER BY detected_at DESC LIMIT 20")
    return jsonify({
      'stats':{'predictions':len(ps),'events':scalar("SELECT COUNT(*) FROM observed_events WHERE source='CSN'"),'correlations':scalar("SELECT COUNT(*) FROM correlation_history"),'hits':sum(x['status']=='ACERTADA' for x in results),'pending':sum(x['status']=='PENDIENTE' for x in results)},
      'events':es[:40], 'predictions':ps, 'ranking':sorted(groups.values(),key=lambda x:(x['accuracy'] if x['accuracy'] is not None else -1),reverse=True), 'correlations':corr
    })

HTML='''<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Sismologia Lab</title>
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"><style>
:root{--bg:#07111f;--card:#0e1c2e;--line:#20344d;--text:#eaf2ff;--muted:#8ea4bd;--cyan:#42d9ff;--green:#56e39f;--yellow:#ffd166;--red:#ff667a}*{box-sizing:border-box}body{margin:0;background:linear-gradient(180deg,#06101d,#091626);color:var(--text);font:14px system-ui,Segoe UI,sans-serif}.wrap{max-width:1400px;margin:auto;padding:24px}.top{display:flex;justify-content:space-between;align-items:end;gap:16px;margin-bottom:20px}.brand h1{margin:0;font-size:28px}.brand p{color:var(--muted);margin:6px 0}.live{color:var(--green);font-weight:700}.grid{display:grid;grid-template-columns:repeat(5,1fr);gap:12px}.card{background:rgba(14,28,46,.94);border:1px solid var(--line);border-radius:16px;padding:16px;box-shadow:0 12px 35px #0003}.metric b{font-size:28px;display:block;margin-top:8px}.metric span,.muted{color:var(--muted)}.main{display:grid;grid-template-columns:1.65fr 1fr;gap:14px;margin-top:14px}#map{height:520px;border-radius:12px}.section-title{font-size:16px;font-weight:800;margin:0 0 12px}.table{width:100%;border-collapse:collapse}.table td,.table th{padding:9px;border-bottom:1px solid var(--line);text-align:left}.table th{color:var(--muted);font-size:12px}.pill{padding:3px 8px;border-radius:999px;background:#172b42}.scroll{max-height:430px;overflow:auto}.corr{padding:10px 0;border-bottom:1px solid var(--line)}.foot{color:var(--muted);font-size:12px;margin:18px 2px}.tabs{display:flex;gap:8px;margin-bottom:12px}.tabs button{background:#15283e;color:var(--text);border:1px solid var(--line);border-radius:9px;padding:8px 11px;cursor:pointer}@media(max-width:900px){.grid{grid-template-columns:repeat(2,1fr)}.main{grid-template-columns:1fr}.top{align-items:start;flex-direction:column}}
</style></head><body><div class="wrap"><div class="top"><div class="brand"><h1>🌎 SISMOLOGIA LAB</h1><p>Dashboard experimental · CSN + evaluación de predicciones</p></div><div class="live">● SISTEMA EN LÍNEA</div></div>
<div class="grid"><div class="card metric"><span>🎯 Predicciones</span><b id="pred">—</b></div><div class="card metric"><span>🌎 Sismos CSN</span><b id="events">—</b></div><div class="card metric"><span>✅ Aciertos</span><b id="hits">—</b></div><div class="card metric"><span>🟡 Pendientes</span><b id="pending">—</b></div><div class="card metric"><span>🚨 Correlaciones</span><b id="corrs">—</b></div></div>
<div class="main"><div class="card"><h2 class="section-title">🗺️ Mapa de Chile</h2><div class="muted" style="margin-bottom:10px">Rojo: sismos CSN · Azul: predicciones</div><div id="map"></div></div><div class="card"><h2 class="section-title">🏆 Ranking por IA</h2><div class="scroll"><table class="table"><thead><tr><th>IA</th><th>✅</th><th>❌</th><th>🟡</th><th>%</th></tr></thead><tbody id="rank"></tbody></table></div></div></div>
<div class="main"><div class="card"><h2 class="section-title">🌎 Últimos sismos CSN</h2><div class="scroll"><table class="table"><thead><tr><th>Fecha</th><th>Lugar</th><th>Mag.</th><th>Prof.</th></tr></thead><tbody id="quake"></tbody></table></div></div><div class="card"><h2 class="section-title">🚨 Correlaciones recientes</h2><div id="corr" class="scroll"></div></div></div>
<div class="foot">⚠️ Proyecto experimental de evaluación. Las coincidencias estadísticas mostradas aquí no constituyen una capacidad científicamente validada para predecir terremotos.</div></div>
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script><script>
const map=L.map('map').setView([-33.4,-70.7],4);L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',{attribution:'© OpenStreetMap'}).addTo(map);let layers=[];
function esc(x){return String(x??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}function fmt(s){try{return new Date(s).toLocaleString('es-CL')}catch{return s}}
async function load(){let d=await fetch('/api/dashboard',{cache:'no-store'}).then(r=>r.json());for(let k of ['predictions','events','hits','pending','correlations'])document.getElementById(k==='predictions'?'pred':k==='correlations'?'corrs':k).textContent=d.stats[k];
document.getElementById('rank').innerHTML=d.ranking.map((x,i)=>`<tr><td>${i<3?['🥇','🥈','🥉'][i]:'▫️'} ${esc(x.group)}</td><td>${x.hits}</td><td>${x.misses}</td><td>${x.pending}</td><td><b>${x.accuracy==null?'—':x.accuracy+'%'}</b></td></tr>`).join('');
document.getElementById('quake').innerHTML=d.events.slice(0,25).map(x=>`<tr><td>${fmt(x.occurred_at)}</td><td>${esc(x.place)}</td><td><b>M${Number(x.magnitude).toFixed(1)}</b></td><td>${x.depth_km==null?'—':x.depth_km+' km'}</td></tr>`).join('');
document.getElementById('corr').innerHTML=d.correlations.length?d.correlations.map(x=>`<div class="corr"><b>${esc(x.prediction_code)}</b> ↔ <span class="pill">M${Number(x.event_mag).toFixed(1)}</span><br><span class="muted">${fmt(x.event_time)} · ${Number(x.distance_km).toFixed(1)} km · ${esc(x.event_place)}</span></div>`).join(''):'<span class="muted">Aún no hay correlaciones.</span>';
layers.forEach(x=>map.removeLayer(x));layers=[];d.events.slice(0,80).forEach(x=>{let m=L.circleMarker([x.latitude,x.longitude],{radius:5,color:'#ff667a',fillOpacity:.8}).bindPopup(`<b>CSN M${x.magnitude}</b><br>${esc(x.place)}<br>${fmt(x.occurred_at)}`);m.addTo(map);layers.push(m)});d.predictions.forEach(x=>{let m=L.circleMarker([x.latitude,x.longitude],{radius:6,color:'#42d9ff',fillOpacity:.7}).bindPopup(`<b>${esc(x.group_name)} · ${esc(x.code)}</b><br>${esc(x.place)}<br>M${x.mag_min}–M${x.mag_max}`);m.addTo(map);layers.push(m)});}
load().catch(console.error);setInterval(()=>load().catch(console.error),60000);
</script></body></html>'''

@app.get('/')
def home(): return render_template_string(HTML)

@app.get('/health')
def health(): return {'ok':True,'db':DB_PATH}

def run_dashboard():
    port=int(os.getenv('PORT','8080'))
    app.run(host='0.0.0.0',port=port,debug=False,use_reloader=False)
