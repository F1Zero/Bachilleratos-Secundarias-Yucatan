# -*- coding: utf-8 -*-
"""Cruza los archivos de 4.5 km, 2 km y el catálogo SIC y genera la página
'Bachilleratos y Secundarias.html' (un solo archivo autónomo).

Uso:  py build_datos.py
"""
import os, json, math, base64, collections
import openpyxl
from pyxlsb import open_workbook

AQUI = os.path.dirname(os.path.abspath(__file__))
F45 = r"C:\Users\SEP\Documents\4.5KM-MEDIA SUPERIOR Y SECUNDARIASv2.xlsx"
F2 = r"F:\Descargas\Presentación 31ES INEGI Media Superior Escuelas Secundaria Vertical.xlsx"
FCAT = (r"C:\Users\SEP\AppData\Roaming\Microsoft\Excel\a_ctba_sic_v2-20261103v2312914373413578445"
        r"\a_ctba_sic_v2-20261103v2((Autorecovered-312914031905818769)).xlsb")
LIBS = r"C:\Users\SEP\Documents\Proyecto_SIG_Portable\Mapa de Escuelas V2\_libs"
LOGO = r"C:\Users\SEP\Documents\Proyecto_SIG_Portable\Mapa de Escuelas V2\logo.png"
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
