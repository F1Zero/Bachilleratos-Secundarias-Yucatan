# -*- coding: utf-8 -*-
"""Cruza los archivos de 4.5 km, 2 km y el catálogo SIC y genera la página
'Bachilleratos y Secundarias.html' (un solo archivo autónomo).

Uso:  py build_datos.py
"""
import os, json, math, base64, collections, time, urllib.request
import openpyxl
from pyxlsb import open_workbook

AQUI = os.path.dirname(os.path.abspath(__file__))
F45 = r"C:\Users\SEP\Documents\4.5KM-MEDIA SUPERIOR Y SECUNDARIASv2.xlsx"
F2 = r"F:\Descargas\Presentación 31ES INEGI Media Superior Escuelas Secundaria Vertical.xlsx"
FCAT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "datos", "a_ctba_sic_v2-20261103v2.xlsb")  # copia fija del catálogo SIC (no se sube a GitHub)
LIBS = r"C:\Users\SEP\Documents\Proyecto_SIG_Portable\Mapa de Escuelas V2\_libs"
LOGO = r"C:\Users\SEP\Documents\Proyecto_SIG_Portable\Mapa de Escuelas V2\logo.png"
F27 = r"F:\Descargas\Asignacion_Secundarias_Promocion_Bachillerato_2027-v3.xlsx"
RUTAS = os.path.join(AQUI, "rutas_osrm.json")
TEMPLATE = os.path.join(AQUI, "_template.html")
OUT = os.path.join(AQUI, "Bachilleratos y Secundarias.html")

TURNOS = {"1": "Matutino", "2": "Vespertino", "3": "Nocturno", "4": "Discontinuo", "5": "Mixto"}


def s(v):
    v = "" if v is None else str(v).strip()
    return "" if v.startswith("#") and v.upper() in ("#N/A", "#N/D", "#VALUE!", "#¡VALOR!", "#REF!") else v


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def dist_km(a, b, c, d):
    r = 6371.0088
    p1, p2 = math.radians(a), math.radians(c)
    dp, dl = p2 - p1, math.radians(d - b)
    h = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * r * math.asin(math.sqrt(h))


# ---------------- Catálogo SIC ----------------
cat, cat10 = {}, {}
with open_workbook(FCAT) as wb:
    with wb.get_sheet(1) as sh:
        it = sh.rows()
        hdr = [c.v for c in next(it)]
        for row in it:
            d = dict(zip(hdr, [c.v for c in row]))
            cct = s(d.get("CLAVECCT"))
            if not cct:
                continue
            cat[(cct, s(d.get("TURNO"))[:1])] = d
            cat10.setdefault(cct, d)


def buscar_cat(cct11):
    return cat.get((cct11[:10], cct11[10:])) or cat10.get(cct11[:10])


def domicilio(d):
    if not d:
        return ""
    partes = [s(d.get("DOMICILIO"))]
    n = s(d.get("NUMEXT"))
    if n and n != "0":
        partes[0] += " #" + n
    elif s(d.get("ALFANUMEXT")) and s(d.get("ALFANUMEXT")) not in ("SN", "NA"):
        partes[0] += " #" + s(d.get("ALFANUMEXT"))
    e1, e2 = s(d.get("ENTRECALLE")), s(d.get("YCALLE"))
    if e1 and "NINGUNO" not in e1:
        partes.append("entre " + e1 + ("" if not e2 or "NINGUNO" in e2 else " y " + e2))
    return ", ".join(p for p in partes if p)


# ---------------- Bachilleratos (hoja MEDIA SUPERIOR) ----------------
ms, ms_idx = [], {}


def nuevo_ms(cct, nombre, lat, lon, carpeta):
    partes = [p.strip() for p in s(carpeta).split("/")]
    sost = partes[0] if partes else ""
    sub = partes[3] if len(partes) > 3 else ""
    d = buscar_cat(cct)
    if not lat and d:  # bachillerato que no está en la lista maestra: coordenadas del catálogo
        lat, lon = num(d.get("LATITUD")) or 0, num(d.get("LONGITUD")) or 0
    o = {
        "c": cct, "n": nombre, "la": round(lat, 6), "lo": round(lon, 6),
        "so": sost, "sb": sub, "t": TURNOS.get(cct[10:], ""),
        "mu": s(d.get("N_MUNICIPIO")) if d else "", "lc": s(d.get("N_LOCALIDAD")) if d else "",
        "as": s(d.get("ASENTAMIEN")) if d else "", "cp": s(d.get("CODPOST")) if d else "",
        "do": domicilio(d), "te": s(d.get("TELEFONO")) if d else "",
        "fi": s(d.get("SOSTENIMIE")) if d else "",
        "ma": None, "dc": None, "h45": 0, "s45": [], "s2": [],
    }
    ms_idx[cct] = len(ms)
    ms.append(o)
    return o


wb = openpyxl.load_workbook(F45, read_only=True, data_only=True)
for r in wb["MEDIA SUPERIOR"].iter_rows(min_row=2, values_only=True):
    if r[0] and s(r[0]) not in ms_idx:
        nuevo_ms(s(r[0]), s(r[5]), num(r[1]), num(r[2]), r[8])

# ---------------- Secundarias ----------------
sec, sec_idx = [], {}


def get_sec(cct, nombre, lat, lon, tipo, sost, extra=None):
    if cct in sec_idx:
        o = sec[sec_idx[cct]]
        if extra:  # la info del archivo de 2 km (asentamiento INEGI) tiene prioridad
            o.update({k: v for k, v in extra.items() if v})
        return sec_idx[cct]
    d = buscar_cat(cct)
    o = {
        "c": cct, "n": nombre, "la": round(lat, 6), "lo": round(lon, 6), "ti": tipo, "so": sost,
        "t": TURNOS.get(cct[10:], ""),
        "mu": s(d.get("N_MUNICIPIO")) if d else "", "lc": s(d.get("N_LOCALIDAD")) if d else "",
        "as": s(d.get("ASENTAMIEN")) if d else "", "cp": s(d.get("CODPOST")) if d else "",
        "do": domicilio(d),
    }
    if extra:
        o.update({k: v for k, v in extra.items() if v})
    sec_idx[cct] = len(sec)
    sec.append(o)
    return sec_idx[cct]


# 4.5 km: hoja 1
for r in wb.worksheets[0].iter_rows(min_row=4, values_only=True):
    if not r[0] or not r[9]:
        continue
    mcct = s(r[0])
    m = ms[ms_idx[mcct]] if mcct in ms_idx else nuevo_ms(mcct, s(r[4]), num(r[1]), num(r[2]), r[7])
    carp = [p.strip() for p in s(r[17]).split("/")]
    sost = carp[0] if carp else ""
    tipo = carp[3] if len(carp) > 3 else ""
    if sost == "CONAFE":
        tipo = "CONAFE"
    i = get_sec(s(r[9]), s(r[14]), num(r[10]), num(r[11]), tipo, sost)
    m["h45"] = 1  # este bachillerato está en el análisis de 4.5 km
    m["s45"].append([i, round(dist_km(m["la"], m["lo"], sec[i]["la"], sec[i]["lo"]), 2)])

# 2 km: archivo nuevo (cada fila = secundaria + bachillerato a <= 2 km)
wb2 = openpyxl.load_workbook(F2, read_only=True, data_only=True)
for r in wb2.worksheets[0].iter_rows(min_row=5, values_only=True):
    if not r[0] or not r[18]:
        continue
    mcct = s(r[18])
    if mcct not in ms_idx:
        sost = s(r[19])
        nuevo_ms(mcct, "", 0, 0, f"{sost}/MEDIA SUPERIOR/BACHILLERATO/{s(r[20])}")
    m = ms[ms_idx[mcct]]
    if not m["mu"]:
        m["mu"], m["lc"] = s(r[21]), s(r[22])
    if m["ma"] is None:
        m["ma"], m["dc"] = r[23], r[24]
    extra = {"mu": s(r[4]), "lc": s(r[6]), "as": s(r[8]), "cp": s(r[11]), "cv": s(r[1])}
    i = get_sec(s(r[0]), s(r[2]), num(r[12]), num(r[13]), s(r[10]), s(r[9]), extra)
    m["s2"].append([i, round(dist_km(m["la"], m["lo"], sec[i]["la"], sec[i]["lo"]), 2)])

for m in ms:
    if not m["n"]:
        d = buscar_cat(m["c"])
        m["n"] = s(d.get("NOMBRECT")) if d else m["c"]
    if not m["la"]:
        d = buscar_cat(m["c"])
        if d:
            m["la"], m["lo"] = num(d.get("LATITUD")) or 0, num(d.get("LONGITUD")) or 0
    # Bachilleratos que no están en la hoja 1 (privados/autónomos): se calcula el radio de 4.5 km
    # con las mismas secundarias de los archivos. Validado: este cálculo reproduce exactamente
    # la hoja 1 en los 370 bachilleratos públicos.
    if not m["h45"] and m["la"]:
        m["h45"] = 2
        m["s45"] = [[i, round(dist_km(m["la"], m["lo"], x["la"], x["lo"]), 2)]
                    for i, x in enumerate(sec) if dist_km(m["la"], m["lo"], x["la"], x["lo"]) <= 4.5]
    # Bachilleratos que no vienen en el archivo de 2 km (p. ej. Bachilleratos Nacionales
    # «Margarita Maza»): se calcula con las mismas secundarias. Es aproximado: el archivo usa
    # coordenadas propias y no coincide exacto con este cálculo.
    if not m["s2"] and m["la"]:
        calc = [[i, round(dist_km(m["la"], m["lo"], x["la"], x["lo"]), 2)]
                for i, x in enumerate(sec) if dist_km(m["la"], m["lo"], x["la"], x["lo"]) <= 2]
        if calc:
            m["s2"], m["h2"] = calc, 2
    m["mu"] = m["mu"] or "SIN DATO"
    m["lc"] = m["lc"] or "SIN DATO"
    m["s45"].sort(key=lambda x: x[1])
    m["s2"].sort(key=lambda x: x[1])

# ---------------- Planteles: solo los 353 de la lista 2027 ----------------
# Del archivo de asignación 2027 SOLO se toma qué planteles se usan (su lista de claves CCT).
# Todos los datos de cada plantel salen de los archivos originales (4.5 km, 2 km) y del catálogo SIC.
# Zona: regla por municipio (Zona Metropolitana = Mérida, Kanasín, Umán, Conkal y Ucú).
ZM_MUN = {"MÉRIDA", "KANASÍN", "UMÁN", "CONKAL", "UCÚ"}
wb27 = openpyxl.load_workbook(F27, read_only=True, data_only=True)
claves27 = [s(r[0]) for r in wb27["Vista por Bachillerato"].iter_rows(min_row=5, values_only=True) if r[0]]
por10 = collections.defaultdict(list)
for m in ms:
    por10[m["c"][:10]].append(m)


def radio(m, km):
    return [[i, round(dist_km(m["la"], m["lo"], x["la"], x["lo"]), 2)]
            for i, x in enumerate(sec) if dist_km(m["la"], m["lo"], x["la"], x["lo"]) <= km]


ms27 = []
for c10 in claves27:
    grupo = sorted(por10.get(c10, []), key=lambda g: g["c"])
    if grupo:  # uno o varios turnos del mismo plantel: se unen
        o = dict(grupo[0])
        d2, d45 = {}, {}
        for g in grupo:
            d2.update({i: k for i, k in g["s2"]})
            d45.update({i: k for i, k in g["s45"]})
        o["s2"] = sorted(([i, k] for i, k in d2.items()), key=lambda x: x[1])
        o["s45"] = sorted(([i, k] for i, k in d45.items()), key=lambda x: x[1])
        mats = [g["ma"] for g in grupo if g["ma"] is not None]
        docs = [g["dc"] for g in grupo if g["dc"] is not None]
        o["ma"] = sum(mats) if mats else None
        o["dc"] = sum(docs) if docs else None
        o["h45"] = 1 if any(g["h45"] == 1 for g in grupo) else 2
        con2 = [g for g in grupo if g["s2"]]
        o.pop("h2", None)
        if con2 and all(g.get("h2") == 2 for g in con2):
            o["h2"] = 2
        ts = list(dict.fromkeys(g["t"] for g in grupo if g["t"]))
        o["t"] = " y ".join(ts)
    else:  # plantel que no viene en los archivos de 2 y 4.5 km: datos del catálogo SIC y radios calculados
        d = cat10.get(c10) or {}
        o = {"c": c10, "n": s(d.get("NOMBRECT")), "sb": s(d.get("SOSTENIMIE")) or "SIN DATO",
             "t": TURNOS.get(s(d.get("TURNO"))[:1], ""), "mu": s(d.get("N_MUNICIPIO")), "lc": s(d.get("N_LOCALIDAD")),
             "la": round(num(d.get("LATITUD")) or 0, 6), "lo": round(num(d.get("LONGITUD")) or 0, 6),
             "as": s(d.get("ASENTAMIEN")), "cp": s(d.get("CODPOST")), "do": domicilio(d) if d else "",
             "te": s(d.get("TELEFONO")), "fi": s(d.get("SOSTENIMIE")), "ma": None, "dc": None}
        o["s2"], o["s45"] = radio(o, 2), radio(o, 4.5)
        o["h2"], o["h45"] = 2, 2
    o["c"] = c10
    o["zo"] = "Zona Metropolitana" if o["mu"] in ZM_MUN else "Interior del Estado"
    o["so"] = "PÚBLICO"
    o.pop("es", None)
    ms27.append(o)
print("Planteles de la lista 2027:", len(ms27), "| sin datos en los archivos de 2 y 4.5 km:",
      sum(1 for c in claves27 if c not in por10), "| zonas:", dict(collections.Counter(m["zo"] for m in ms27)))
ms = ms27

# ---------------- Solo secundarias activas del catálogo SIC ----------------
# Se quitan las secundarias de los archivos que no están activas en el SIC (NIVEL 04, STATUS 1) y se agregan
# las activas que no vienen en los archivos; a estas se les calcula la distancia a los 353 planteles (campo f="SIC").
activas = {}
for d in cat.values():
    if s(d.get("NIVEL")) == "04" and s(d.get("STATUS")) == "1":
        activas.setdefault(s(d["CLAVECCT"]), d)
bajas = {i for i, x in enumerate(sec) if x["c"][:10] not in activas}
for m in ms:
    m["s2"] = [p for p in m["s2"] if p[0] not in bajas]
    m["s45"] = [p for p in m["s45"] if p[0] not in bajas]


def tipo_sec(c10):
    if c10[2] == "K":
        return "CONAFE"
    return {"TV": "TELESECUNDARIA", "ST": "TECNICA", "SN": "PARA TRABAJADORES"}.get(c10[3:5], "GENERAL")


en_arch = {x["c"][:10] for x in sec}
n_nuevas = 0
for c10, d in sorted(activas.items()):
    la, lo = num(d.get("LATITUD")), num(d.get("LONGITUD"))
    if c10 in en_arch or not la:
        continue
    so = "CONAFE" if c10[2] == "K" else "PRIVADO" if c10[2] == "P" else "PÚBLICO"
    i = get_sec(c10 + s(d.get("TURNO"))[:1], s(d.get("NOMBRECT")), la, lo, tipo_sec(c10), so)
    sec[i]["f"] = "SIC"
    for m in ms:
        if m["la"]:
            k = round(dist_km(m["la"], m["lo"], la, lo), 2)
            if k <= 4.5:
                m["s45"].append([i, k])
            if k <= 2:
                m["s2"].append([i, k])
    n_nuevas += 1
for m in ms:
    m["s2"].sort(key=lambda x: x[1])
    m["s45"].sort(key=lambda x: x[1])
print(f"Secundarias activas en el SIC: {len(activas)} | quitadas por no estar activas: {len(bajas)} | "
      f"agregadas del SIC (no venían en los archivos): {n_nuevas}")

# ---------------- Cobertura total ----------------
# Toda secundaria activa debe tener bachillerato: si no tiene ningún plantel a 4.5 km o menos,
# se asigna a su plantel más cercano en línea recta (lista "sx").
cubiertas = {i for m in ms for i, _ in m["s2"] + m["s45"]}
for m in ms:
    m["sx"] = []
n_sin = 0
for i, x in enumerate(sec):
    if i in cubiertas or i in bajas or not x["la"]:
        continue
    m, dk = min(((m, dist_km(m["la"], m["lo"], x["la"], x["lo"])) for m in ms if m["la"]), key=lambda z: z[1])
    m["sx"].append([i, round(dk, 2)])
    n_sin += 1
print(f"Secundarias (CCT y turno): {len(sec) - len(bajas)} | con plantel a 4.5 km o menos: {len(cubiertas)} | "
      f"asignadas al plantel más cercano (fuera de 4.5 km): {n_sin}")

# Reindexar secundarias (todas quedan asignadas)
usadas = sorted({i for m in ms for i, _ in m["s2"] + m["s45"] + m["sx"]})
nuevo_idx = {v: k for k, v in enumerate(usadas)}
sec = [sec[i] for i in usadas]
for m in ms:
    for key in ("s2", "s45", "sx"):
        m[key] = sorted(([nuevo_idx[i], k] for i, k in m[key]), key=lambda z: z[1])

# ---------------- Distancia por ruta (OSRM / OpenStreetMap, en auto) ----------------
# Se guarda en rutas_osrm.json para no volver a consultar el servicio.
try:
    with open(RUTAS, encoding="utf-8") as f:
        cache = json.load(f)
except (OSError, ValueError):
    cache = {}
pend = 0
for k, m in enumerate(ms):
    destinos = sorted({i for i, _ in m["s2"] + m["s45"] + m["sx"]})
    ya = cache.setdefault(m["c"], {})
    faltan = [i for i in destinos if sec[i]["c"] not in ya]
    for a in range(0, len(faltan), 90):
        lote = faltan[a:a + 90]
        coords = ";".join([f"{m['lo']},{m['la']}"] + [f"{sec[i]['lo']},{sec[i]['la']}" for i in lote])
        url = f"https://router.project-osrm.org/table/v1/driving/{coords}?sources=0&annotations=distance"
        for intento in range(4):
            try:
                req = urllib.request.Request(url, headers={"User-Agent": "bachilleratos-secundarias/1.0"})
                with urllib.request.urlopen(req, timeout=60) as resp:
                    res = json.load(resp)
                break
            except Exception as e:  # reintento con espera
                print("  reintento", m["c"], e)
                time.sleep(5 * (intento + 1))
        else:
            continue
        fila = (res.get("distances") or [[]])[0][1:]
        for i, dm in zip(lote, fila):
            ya[sec[i]["c"]] = None if dm is None else round(dm / 1000, 2)
        pend += 1
        time.sleep(1.1)  # uso moderado del servidor público
    if pend and pend % 25 == 0:
        with open(RUTAS, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False)
with open(RUTAS, "w", encoding="utf-8") as f:
    json.dump(cache, f, ensure_ascii=False)
print("Consultas OSRM nuevas:", pend)
for m in ms:
    ru = cache.get(m["c"], {})
    m["s2"] = [[i, k, ru.get(sec[i]["c"])] for i, k in m["s2"]]
    m["s45"] = [[i, k, ru.get(sec[i]["c"])] for i, k in m["s45"]]
    m["sx"] = [[i, k, ru.get(sec[i]["c"])] for i, k in m["sx"]]

print("Bachilleratos:", len(ms), "| con sec 4.5 km:", sum(1 for m in ms if m["s45"]),
      "| con sec 2 km:", sum(1 for m in ms if m["s2"]))
print("Secundarias distintas:", len(sec))
print("Sin municipio:", sum(1 for m in ms if m["mu"] == "SIN DATO"))
print("Sostenimiento:", collections.Counter(m["so"] for m in ms))

# ---------------- Generar HTML ----------------
def leer(p):
    with open(p, encoding="utf-8") as f:
        return f.read()


html = leer(TEMPLATE)
css = leer(os.path.join(LIBS, "leaflet.css"))
js = leer(os.path.join(LIBS, "leaflet.js"))
with open(LOGO, "rb") as f:
    logo = "data:image/png;base64," + base64.b64encode(f.read()).decode()
data = json.dumps({"ms": ms, "sec": sec}, ensure_ascii=False, separators=(",", ":"))
html = (html.replace("/*__LEAFLET_CSS__*/", css)
            .replace("/*__LEAFLET_JS__*/", js)
            .replace("__LOGO__", logo)
            .replace("/*__DATA__*/", "window.DATA=" + data + ";"))
with open(OUT, "w", encoding="utf-8") as f:
    f.write(html)
with open(os.path.join(AQUI, "index.html"), "w", encoding="utf-8") as f:  # copia para GitHub Pages
    f.write(html)
print("OK ->", OUT, round(os.path.getsize(OUT) / 1024), "KB")
