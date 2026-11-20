# -*- coding: utf-8 -*-
"""Genera "Secundarias sin bachillerato considerado.xlsx".

Secundarias ACTIVAS del catálogo SIC (nivel 04, status 1) que no aparecen en los archivos de 2 km ni de
4.5 km, y por lo tanto no quedaron asignadas a ninguno de los 353 planteles. Para cada una se indica su
plantel más cercano (línea recta y ruta OSRM) y cuántos planteles tiene a 2 km y a 4.5 km.

Uso:  py excel_sin_bachillerato.py   (requiere Excel + pywin32 para recalcular)
"""
import os, json, time, collections, urllib.request
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import CellIsRule, DataBarRule
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList

AQUI = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(AQUI, "Secundarias sin bachillerato considerado.xlsx")
RUTAS = os.path.join(AQUI, "rutas_osrm.json")

# ---------------- Datos: el mismo cruce de build_datos.py hasta la cobertura ----------------
src = open(os.path.join(AQUI, "build_datos.py"), encoding="utf-8").read()
src = src[:src.index("# Reindexar secundarias")]
G = {"__file__": os.path.join(AQUI, "build_datos.py")}
exec(compile(src, "build_datos.py", "exec"), G)
SEC, MS, CAT, s, num, dist_km, domicilio = G["sec"], G["ms"], G["cat"], G["s"], G["num"], G["dist_km"], G["domicilio"]
TURNOS = G["TURNOS"]

en_archivos = {x["c"][:10] for x in SEC if x.get("f") != "SIC"}   # build_datos ya agrega las del SIC (f="SIC")
sic = [d for d in CAT.values() if s(d.get("NIVEL")) == "04"]
activas = [d for d in sic if s(d.get("STATUS")) == "1"]
activas10 = {s(d["CLAVECCT"]) for d in activas}
fuera = [d for d in activas if s(d["CLAVECCT"]) not in en_archivos]
bajas = [x for x in SEC if x["c"][:10] not in activas10 and x.get("f") != "SIC"]   # en los archivos pero sin status activo en el SIC

TIPO = {"ES": "Secundaria general", "ST": "Secundaria técnica", "TV": "Telesecundaria", "SN": "Secundaria para trabajadores"}


def tipo(cct):
    if cct[2] == "K":
        return "Secundaria comunitaria (CONAFE)"
    return TIPO.get(cct[3:5], s(cct[2:5]))


def sost(d, cct):
    return "CONAFE" if cct[2] == "K" else s(d.get("SOSTENIMIE"))


filas = []
for d in fuera:
    cct = s(d["CLAVECCT"])
    la, lo = num(d.get("LATITUD")) or 0, num(d.get("LONGITUD")) or 0
    ds = sorted(((dist_km(m["la"], m["lo"], la, lo), m) for m in MS if m["la"]), key=lambda z: z[0])
    k, m = ds[0]
    filas.append({"c": cct, "n": s(d.get("NOMBRECT")), "t": TURNOS.get(s(d.get("TURNO"))[:1], ""),
                  "ti": tipo(cct), "so": sost(d, cct), "mu": s(d.get("N_MUNICIPIO")), "lc": s(d.get("N_LOCALIDAD")),
                  "as": s(d.get("ASENTAMIEN")), "cp": s(d.get("CODPOST")), "do": domicilio(d),
                  "la": round(la, 6), "lo": round(lo, 6), "m": m, "k": round(k, 2),
                  "n2": sum(1 for z, _ in ds if z <= 2), "n45": sum(1 for z, _ in ds if z <= 4.5)})
filas.sort(key=lambda f: (f["mu"], f["lc"], f["k"], f["c"]))
print(f"SIC secundarias activas: {len(activas)} | en archivos: {len(activas) - len(fuera)} | sin bachillerato: {len(fuera)} | "
      f"con plantel a 4.5 km: {sum(1 for f in filas if f['n45'])} | bajas en archivos: {len(bajas)}")

# ---------------- Ruta OSRM al plantel más cercano (caché compartida con la página) ----------------
try:
    cache = json.load(open(RUTAS, encoding="utf-8"))
except (OSError, ValueError):
    cache = {}
por_plantel = collections.defaultdict(list)
for f in filas:
    if f["c"] not in cache.get(f["m"]["c"], {}):
        por_plantel[f["m"]["c"]].append(f)
for mc, lote in por_plantel.items():
    m = lote[0]["m"]
    coords = ";".join([f"{m['lo']},{m['la']}"] + [f"{f['lo']},{f['la']}" for f in lote])
    url = f"https://router.project-osrm.org/table/v1/driving/{coords}?sources=0&annotations=distance"
    for intento in range(4):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "bachilleratos-secundarias/1.0"})
            res = json.load(urllib.request.urlopen(req, timeout=60))
            break
        except Exception as e:
            print("  reintento", mc, e)
            time.sleep(5 * (intento + 1))
    else:
        continue
    for f, dm in zip(lote, (res.get("distances") or [[]])[0][1:]):
        cache.setdefault(mc, {})[f["c"]] = None if dm is None else round(dm / 1000, 2)
    time.sleep(1.1)
json.dump(cache, open(RUTAS, "w", encoding="utf-8"), ensure_ascii=False)
for f in filas:
    f["r"] = cache.get(f["m"]["c"], {}).get(f["c"])
print("Consultas OSRM nuevas:", len(por_plantel))

# ---------------- Estilo (igual que el resumen) ----------------
GUINDA, GUINDA_OSC, DORADO, DORADO_T, ROJO = "AB2240", "4A0D1E", "B08D3C", "8C6A24", "B91C1C"
GRIS, LINEA, ROSA, CREMA, ROJO_S = "6B7280", "E5E7EB", "F8E8EC", "F8F1E2", "FDE2E2"
FN = "Calibri"
fnt = lambda **k: Font(name=FN, **k)
fill = lambda c: PatternFill("solid", fgColor=c)
thin = Side(style="thin", color=LINEA)
caja = Border(left=thin, right=thin, top=thin, bottom=thin)
CENTRO = Alignment(horizontal="center", vertical="center", wrap_text=True)
IZQ = Alignment(horizontal="left", vertical="center")
DER = Alignment(horizontal="right", vertical="center")

wb = openpyxl.Workbook()
wb.remove(wb.active)


def encabezado(ws, titulo, sub, ancho):
    ws.sheet_view.showGridLines = False
    ws["A1"], ws["A2"] = titulo, sub
    ws["A1"].font = fnt(size=18, bold=True, color=GUINDA_OSC)
    ws["A2"].font = fnt(size=11, color=GRIS)
    ws.row_dimensions[1].height = 28
    for c in range(1, ancho + 1):
        ws.cell(row=3, column=c).fill = fill(GUINDA)
    ws.row_dimensions[3].height = 4


def hdr_cell(c, txt, color=GUINDA_OSC):
    c.value = txt
    c.font = fnt(bold=True, color="FFFFFF")
    c.fill = fill(color)
    c.alignment = CENTRO
    c.border = caja


def tabla(ws, nombre, hdrs, datos, anchos, fila0=5, formatos=None):
    for j, t in enumerate(hdrs, 1):
        hdr_cell(ws.cell(row=fila0 - 1, column=j), t)
    ws.row_dimensions[fila0 - 1].height = 45
    for i, fila in enumerate(datos):
        for j, v in enumerate(fila, 1):
            c = ws.cell(row=fila0 + i, column=j, value=v)
            c.font, c.border, c.alignment = fnt(), caja, IZQ
            if formatos and j in formatos:
                c.number_format, c.alignment = formatos[j], DER
    for j, w in enumerate(anchos, 1):
        ws.column_dimensions[L(j)].width = w
    ult = fila0 + len(datos) - 1
    t = Table(displayName=nombre, ref=f"A{fila0 - 1}:{L(len(hdrs))}{ult}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws.add_table(t)
    ws.freeze_panes = ws.cell(row=fila0, column=1)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{fila0 - 1}:{fila0 - 1}"
    return ult


SIT_45 = "Asignada a sus planteles a 4.5 km"
SIT_X = "Asignada al plantel más cercano (> 4.5 km)"

# ===================== Hoja: Secundarias sin bachillerato =====================
SL = "Secundarias sin bachillerato"
wsL = wb.create_sheet(SL)
hL = ["CCT Secundaria", "Nombre Secundaria", "Turno", "Tipo", "Sostenimiento", "Municipio", "Localidad",
      "Asentamiento", "CP", "Domicilio", "Situación", "Planteles a 2 km", "Planteles a 4.5 km",
      "CCT Bachillerato más cercano", "Plantel más cercano", "Subsistema", "Municipio plantel", "Zona plantel",
      "Distancia lineal (km)", "Distancia ruta (km)", "Rango de distancia", "Mismo municipio", "Ver ruta",
      "Lat. Secundaria", "Lon. Secundaria", "Lat. Bachillerato", "Lon. Bachillerato"]
r0 = 5
datos = []
for i, f in enumerate(filas):
    r, m = r0 + i, f["m"]
    datos.append([f["c"], f["n"], f["t"], f["ti"], f["so"], f["mu"], f["lc"], f["as"], f["cp"], f["do"],
                  f'=IF(M{r}>0,"{SIT_45}","{SIT_X}")', f["n2"], f["n45"],
                  m["c"], m["n"], m["sb"], m["mu"], m["zo"], f["k"], f["r"],
                  f'=IF(S{r}<=2,"≤ 2 km",IF(S{r}<=4.5,"2 a 4.5 km",IF(S{r}<=10,"4.5 a 10 km",IF(S{r}<=20,"10 a 20 km","Más de 20 km"))))',
                  f'=IF(F{r}=Q{r},"SÍ","NO")',
                  f'=HYPERLINK("https://www.google.com/maps/dir/?api=1&origin="&X{r}&","&Y{r}&"&destination="&Z{r}&","&AA{r}&"&travelmode=driving","Ver ruta")',
                  f["la"], f["lo"], m["la"], m["lo"]])
encabezado(wsL, "Secundarias que no estaban consideradas (ya asignadas)",
           "Secundarias activas del catálogo SIC que no aparecen en los archivos de 2 km ni de 4.5 km. Ya se agregaron a la página, al resumen y a la presentación: "
           "con plantel a 4.5 km o menos se asignan a esos planteles; sin plantel a 4.5 km, a su plantel más cercano. Orden: municipio, localidad y distancia.", len(hL))
uL = tabla(wsL, "tSinBach", hL, datos,
           [13, 34, 11, 26, 16, 16, 22, 22, 8, 34, 30, 9, 9, 13, 34, 20, 16, 18, 10, 10, 13, 9, 10, 10, 10, 10, 10],
           formatos={12: "0", 13: "0", 19: "0.00", 20: "0.00", 24: "0.000000", 25: "0.000000", 26: "0.000000", 27: "0.000000"})
for r in range(r0, uL + 1):
    wsL.cell(row=r, column=23).font = fnt(color=DORADO_T, underline="single")
wsL.conditional_formatting.add(f"K{r0}:K{uL}", CellIsRule(operator="equal", formula=[f'"{SIT_45}"'], font=Font(name=FN, color=DORADO_T, bold=True), fill=fill(CREMA)))
wsL.conditional_formatting.add(f"K{r0}:K{uL}", CellIsRule(operator="equal", formula=[f'"{SIT_X}"'], font=Font(name=FN, color=ROJO, bold=True), fill=fill(ROJO_S)))
wsL.conditional_formatting.add(f"S{r0}:S{uL}", DataBarRule(start_type="num", start_value=0, end_type="max", color=DORADO))
for col in ("X", "Y", "Z", "AA"):
    wsL.column_dimensions[col].outlineLevel = 1
RL = lambda col: f"'{SL}'!${col}${r0}:${col}${uL}"

# ===================== Hoja: Por municipio =====================
SM = "Por municipio"
wsM = wb.create_sheet(SM)
muns = sorted({f["mu"] for f in filas})
hM = ["Municipio", "Secundarias sin bachillerato", "Comunitarias CONAFE", "Telesecundarias", "Generales y técnicas",
      "Con plantel a 4.5 km", "Sin plantel a 4.5 km", "Distancia lineal prom. al más cercano (km)", "Distancia ruta prom. (km)",
      "Máxima distancia lineal (km)"]
datosM = []
for i, mu in enumerate(muns):
    r = r0 + i
    datosM.append([mu, f'=COUNTIF({RL("F")},A{r})', f'=COUNTIFS({RL("F")},A{r},{RL("D")},"Secundaria comunitaria (CONAFE)")',
                   f'=COUNTIFS({RL("F")},A{r},{RL("D")},"Telesecundaria")',
                   f'=B{r}-C{r}-D{r}', f'=COUNTIFS({RL("F")},A{r},{RL("M")},">0")', f'=COUNTIFS({RL("F")},A{r},{RL("M")},0)',
                   f'=AVERAGEIFS({RL("S")},{RL("F")},A{r})', f'=IFERROR(AVERAGEIFS({RL("T")},{RL("F")},A{r}),"")',
                   f'=_xlfn.MAXIFS({RL("S")},{RL("F")},A{r})'])
encabezado(wsM, "Secundarias sin bachillerato por municipio", f"{len(muns)} municipios. Conteos con fórmulas a partir de la hoja «{SL}».", len(hM))
uM = tabla(wsM, "tMunicipios", hM, datosM, [22, 12, 12, 12, 12, 12, 12, 14, 12, 12], formatos={2: "0", 3: "0", 4: "0", 5: "0", 6: "0", 7: "0", 8: "0.00", 9: "0.00", 10: "0.00"})
rt = uM + 1
wsM.cell(row=rt, column=1, value="Total").font = fnt(bold=True, color="FFFFFF")
for j in range(1, len(hM) + 1):
    c = wsM.cell(row=rt, column=j)
    c.fill, c.border = fill(GUINDA_OSC), caja
    if j > 1:
        col = L(j)
        c.value = (f"=SUM({col}{r0}:{col}{uM})" if j <= 7 else
                   f'=AVERAGE({RL("S")})' if j == 8 else f'=AVERAGE({RL("T")})' if j == 9 else f'=MAX({RL("S")})')
        c.font, c.number_format, c.alignment = fnt(bold=True, color="FFFFFF"), ("0" if j <= 7 else "0.00"), DER
wsM.conditional_formatting.add(f"B{r0}:B{uM}", DataBarRule(start_type="num", start_value=0, end_type="max", color=GUINDA))

# ===================== Hoja: Resumen (primera) =====================
wsR = wb.create_sheet("Resumen", 0)
encabezado(wsR, "Secundarias que no estaban consideradas",
           "Secundarias activas del catálogo SIC que no venían en los archivos de 2 km y 4.5 km · ya asignadas a los 353 planteles 2027", 12)
for j, w in enumerate([30, 13, 13, 13, 13, 13, 3, 30, 13, 13, 13, 13], 1):
    wsR.column_dimensions[L(j)].width = w

kpis = [("Secundarias activas en el SIC", f'={len(activas)}'),
        ("Venían en los archivos de 2 y 4.5 km", f'=B6-B8'),
        ("No venían (agregadas y asignadas)", f'=ROWS({RL("A")})'),
        ("…asignadas a planteles a 4.5 km", f'=COUNTIF({RL("M")},">0")'),
        ("…asignadas al más cercano (> 4.5 km)", f'=COUNTIF({RL("M")},0)')]
hdr_cell(wsR.cell(row=5, column=1), "Indicador")
hdr_cell(wsR.cell(row=5, column=2), "Secundarias")
hdr_cell(wsR.cell(row=5, column=3), "%")
for i, (t, fml) in enumerate(kpis):
    r = 6 + i
    wsR.cell(row=r, column=1, value=t)
    wsR.cell(row=r, column=2, value=fml).number_format = "#,##0"
    wsR.cell(row=r, column=3, value=f"=B{r}/$B$6").number_format = "0.0%"
    for j in (1, 2, 3):
        c = wsR.cell(row=r, column=j)
        c.border = caja
        c.font = fnt(bold=(i == 2), color=(ROJO if i == 2 else "000000"), size=(13 if i == 2 else 11))
        if j > 1:
            c.alignment = DER
wsR.cell(row=6, column=2).font = fnt()
wsR.cell(row=6, column=2).number_format = "#,##0"

# Por tipo
r = 13
wsR.cell(row=r - 1, column=1, value="1 · Por tipo de secundaria").font = fnt(size=13, bold=True, color=GUINDA_OSC)
for j, t in enumerate(["Tipo", "Total", "Con plantel a 4.5 km", "Sin plantel a 4.5 km", "Lineal prom. (km)", "Ruta prom. (km)"], 1):
    hdr_cell(wsR.cell(row=r, column=j), t)
tipos = sorted({f["ti"] for f in filas}, key=lambda t: -sum(1 for f in filas if f["ti"] == t))
for i, t in enumerate(tipos):
    rr = r + 1 + i
    vals = [t, f'=COUNTIF({RL("D")},A{rr})', f'=COUNTIFS({RL("D")},A{rr},{RL("M")},">0")', f'=COUNTIFS({RL("D")},A{rr},{RL("M")},0)',
            f'=AVERAGEIFS({RL("S")},{RL("D")},A{rr})', f'=IFERROR(AVERAGEIFS({RL("T")},{RL("D")},A{rr}),"")']
    for j, v in enumerate(vals, 1):
        c = wsR.cell(row=rr, column=j, value=v)
        c.border, c.font = caja, fnt()
        if j > 1:
            c.number_format, c.alignment = ("0" if j <= 4 else "0.00"), DER
rt_tipo = r + 1 + len(tipos)
for j in range(1, 7):
    c = wsR.cell(row=rt_tipo, column=j)
    c.fill, c.border, c.font = fill(GUINDA_OSC), caja, fnt(bold=True, color="FFFFFF")
    c.value = "Total" if j == 1 else (f"=SUM({L(j)}{r + 1}:{L(j)}{rt_tipo - 1})" if j <= 4 else
                                      f'=AVERAGE({RL("S")})' if j == 5 else f'=AVERAGE({RL("T")})')
    if j > 1:
        c.number_format, c.alignment = ("0" if j <= 4 else "0.00"), DER

# Por rango de distancia al plantel más cercano
r = 13
wsR.cell(row=r - 1, column=8, value="2 · Distancia al plantel más cercano").font = fnt(size=13, bold=True, color=GUINDA_OSC)
for j, t in enumerate(["Rango (lineal)", "Secundarias", "%"], 8):
    hdr_cell(wsR.cell(row=r, column=j), t)
rangos = ["≤ 2 km", "2 a 4.5 km", "4.5 a 10 km", "10 a 20 km", "Más de 20 km"]
for i, t in enumerate(rangos):
    rr = r + 1 + i
    for j, v in enumerate([t, f'=COUNTIF({RL("U")},H{rr})', f"=I{rr}/$B$8"], 8):
        c = wsR.cell(row=rr, column=j, value=v)
        c.border, c.font = caja, fnt()
        if j > 8:
            c.number_format, c.alignment = ("0" if j == 9 else "0.0%"), DER
ch = BarChart()
ch.type, ch.style = "bar", 10
ch.title = "Secundarias por distancia al plantel más cercano"
ch.add_data(Reference(wsR, min_col=9, min_row=r, max_row=r + len(rangos)), titles_from_data=True)
ch.set_categories(Reference(wsR, min_col=8, min_row=r + 1, max_row=r + len(rangos)))
ch.legend = None
ch.y_axis.delete = ch.x_axis.delete = False
ch.y_axis.majorGridlines = None
ch.x_axis.scaling.orientation = "maxMin"
ch.roundedCorners = False
ch.title.overlay = False
ch.series[0].graphicalProperties.solidFill = DORADO
ch.series[0].graphicalProperties.line.solidFill = DORADO
ch.dataLabels = DataLabelList()
ch.dataLabels.showVal = True
ch.dataLabels.showSerName = ch.dataLabels.showCatName = ch.dataLabels.showLegendKey = False
ch.width, ch.height = 15, 7
wsR.add_chart(ch, f"H{r + len(rangos) + 3}")

# Nota de lectura
rn = r + len(rangos) + 3 + 16   # debajo de la gráfica (el texto se desborda hacia la derecha)
notas = [
    "Cómo leerlo",
    f"• Solo se consideran secundarias activas del SIC. Las que no venían en los archivos de 2 km y 4.5 km (la mayoría comunitarias de CONAFE) ya se agregaron.",
    f"• Con plantel a 4.5 km o menos: se asignan a todos esos planteles, igual que las demás secundarias.",
    f"• Sin plantel a 4.5 km: se asignan a su plantel más cercano en línea recta (columna «Plantel más cercano»).",
    f"• Detalle en la hoja «{SL}»; resumen por municipio en «{SM}».",
]
for i, t in enumerate(notas):
    c = wsR.cell(row=rn + i, column=1, value=t)
    c.font = fnt(bold=(i == 0), size=(12 if i == 0 else 11), color=(GUINDA_OSC if i == 0 else "000000"))
wsR.sheet_view.zoomScale = 100

# ===================== Hoja: Notas y método =====================
wsN = wb.create_sheet("Notas y método")
encabezado(wsN, "Notas y método", "Fuentes y criterio", 6)
wsN.column_dimensions["A"].width = 130
txt = [
    ("Universo", True),
    (f"Secundarias activas del catálogo SIC (NIVEL = 04 Secundaria, STATUS = 1): {len(activas)}.", False),
    ("Se comparan por clave CCT con las secundarias de los archivos de 2 km («Presentación 31ES INEGI Media Superior Escuelas Secundaria "
     "Vertical») y de 4.5 km («4.5KM-MEDIA SUPERIOR Y SECUNDARIASv2», hoja 1). Las que no aparecen en ninguno de los dos quedan en esta lista.", False),
    ("Plantel más cercano", True),
    ("Entre los 353 planteles del proceso de ingreso 2027 (misma lista y datos que la página y el resumen). Distancia lineal: haversine con las "
     "coordenadas del SIC. Distancia por ruta: camino más corto en auto con OSRM (OpenStreetMap).", False),
    ("Planteles a 2 km / 4.5 km: cuántos de los 353 planteles hay a esa distancia en línea recta.", False),
    ("Tipo", True),
    ("Por la clave CCT: K = comunitaria CONAFE; ES = general; ST = técnica; TV = telesecundaria; SN = para trabajadores.", False),
    ("Secundarias en los archivos que no están activas en el SIC", True),
    (f"{len(bajas)} secundarias de los archivos no tienen status activo en el SIC (siguen asignadas en el resumen): "
     + ", ".join(f"{x['c'][:10]} {x['n']}" for x in bajas) + ".", False),
    ("Departamento de Estudios Territoriales · SEGEY", True),
]
for i, (t, b) in enumerate(txt):
    c = wsN.cell(row=5 + i, column=1, value=t)
    c.font = fnt(bold=b, size=(12 if b else 11), color=(GUINDA_OSC if b else "000000"))
    c.alignment = Alignment(wrap_text=True, vertical="top")

for ws in wb.worksheets:
    ws.sheet_properties.tabColor = GUINDA if ws.title != "Notas y método" else DORADO
wb.save(OUT)
print("openpyxl OK ->", OUT)

# ---------------- Recalcular y guardar con Excel ----------------
import win32com.client as win32
xl = win32.DispatchEx("Excel.Application")
xl.Visible, xl.DisplayAlerts = False, False
try:
    wbx = xl.Workbooks.Open(OUT)
    xl.CalculateFull()
    wbx.Worksheets(1).Activate()
    wbx.Save()
    wbx.Close(False)
finally:
    xl.Quit()
print("Excel OK")
