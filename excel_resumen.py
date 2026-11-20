# -*- coding: utf-8 -*-
"""Genera "Resumen Bachilleratos y Secundarias 2027.xlsx".

Datos: los de la página (index.html), que salen de los archivos de 2 km y 4.5 km, del catálogo SIC
y de las rutas OSRM. Del archivo de asignación 2027 solo se usa la lista de planteles (en build_datos.py).

1) openpyxl: hojas, datos, fórmulas, tablas, formato condicional y gráficas.
2) Excel (COM): tabla dinámica RESUMEN con segmentaciones, recálculo y guardado.

Uso:  py build_datos.py  y luego  py excel_resumen.py
"""
import os, json, math, datetime, collections
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import DataBarRule, CellIsRule
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.comments import Comment

AQUI = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(AQUI, "Resumen Bachilleratos y Secundarias 2027.xlsx")
URL_PAGINA = "https://f1zero.github.io/Bachilleratos-Secundarias-Yucatan/"
R2, R45, RX = "≤ 2 km", "2 a 4.5 km", "Más cercano (> 4.5 km)"   # valores de la columna Radio

# ---------------- Datos ----------------
html = open(os.path.join(AQUI, "index.html"), encoding="utf-8").read()
i0 = html.index("window.DATA=") + 12
D = json.loads(html[i0:html.index(";</script>", i0)])
MS, SEC = D["ms"], D["sec"]
ZM = "Zona Metropolitana"

# Secundarias asignadas a cada plantel: a 2 km, de 2 a 4.5 km y, por cobertura, las que no tienen plantel a 4.5 km
for m in MS:
    d = {}
    for s, k, r in m["s45"]:
        d[s] = [s, k, r, R45]
    for s, k, r in m["s2"]:
        d[s] = [s, k, r, R2]
    for s, k, r in m.get("sx", []):
        d[s] = [s, k, r, RX]
    m["_pares"] = sorted(d.values(), key=lambda x: (x[1], x[2] if x[2] is not None else 1e9, SEC[x[0]]["c"]))
orden_ms = sorted(MS, key=lambda m: (m["mu"], m["zo"], m["sb"], m["c"]))
pares = [(m, SEC[s], k, r, rad) for m in orden_ms for s, k, r, rad in m["_pares"]]

# ---------------- Estilo ----------------
GUINDA, GUINDA_OSC, DORADO, DORADO_T, ROJO = "AB2240", "4A0D1E", "B08D3C", "8C6A24", "B91C1C"
GRIS, LINEA, SUAVE, ROSA, CREMA, ROJO_S = "6B7280", "E5E7EB", "F7F5F6", "F8E8EC", "F8F1E2", "FDE2E2"
F = "Calibri"
fnt = lambda **k: Font(name=F, **k)
fill = lambda c: PatternFill("solid", fgColor=c)
thin = Side(style="thin", color=LINEA)
caja = Border(left=thin, right=thin, top=thin, bottom=thin)
CENTRO = Alignment(horizontal="center", vertical="center", wrap_text=True)
IZQ = Alignment(horizontal="left", vertical="center", wrap_text=False)
IZQW = Alignment(horizontal="left", vertical="top", wrap_text=True)

wb = openpyxl.Workbook()
wb.remove(wb.active)


def encabezado(ws, titulo, subtitulo, ancho):
    ws.sheet_view.showGridLines = False
    ws["A1"] = titulo
    ws["A1"].font = fnt(size=18, bold=True, color=GUINDA_OSC)
    ws["A2"] = subtitulo
    ws["A2"].font = fnt(size=11, color=GRIS)
    ws.row_dimensions[1].height = 28
    for c in range(1, ancho + 1):
        ws.cell(row=3, column=c).fill = fill(GUINDA)
    ws.row_dimensions[3].height = 4


def tabla(ws, nombre, hdrs, filas, anchos, fila0=5, formatos=None, wrap=None, alto_hdr=50):
    """Encabezado en fila0-1, datos desde fila0 y una Tabla de Excel."""
    h = fila0 - 1
    for j, t in enumerate(hdrs, 1):
        c = ws.cell(row=h, column=j, value=t)
        c.font = fnt(bold=True, color="FFFFFF", size=11)
        c.fill = fill(GUINDA_OSC)
        c.alignment = CENTRO
        c.border = caja
    ws.row_dimensions[h].height = alto_hdr
    for i, fila in enumerate(filas):
        r = fila0 + i
        for j, v in enumerate(fila, 1):
            c = ws.cell(row=r, column=j, value=v)
            c.font = fnt(size=11)
            c.border = caja
            c.alignment = IZQW if (wrap and j in wrap) else IZQ
            if formatos and j in formatos:
                c.number_format = formatos[j]
                c.alignment = Alignment(horizontal="right", vertical="center")
    for j, w in enumerate(anchos, 1):
        ws.column_dimensions[L(j)].width = w
    ultimo = fila0 + len(filas) - 1
    t = Table(displayName=nombre, ref=f"A{h}:{L(len(hdrs))}{ultimo}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws.add_table(t)
    ws.freeze_panes = ws.cell(row=fila0, column=1)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{h}:{h}"
    return ultimo


def gmaps_ruta(fila, la1, lo1, la2, lo2):
    return (f'=HYPERLINK("https://www.google.com/maps/dir/?api=1&origin="&{la1}{fila}&","&{lo1}{fila}'
            f'&"&destination="&{la2}{fila}&","&{lo2}{fila}&"&travelmode=driving","Ver ruta")')


def gmaps_punto(fila, la, lo):
    return f'=HYPERLINK("https://www.google.com/maps/search/?api=1&query="&{la}{fila}&","&{lo}{fila},"Ver en mapa")'


def liga(ws, col, f0, f1):
    for r in range(f0, f1 + 1):
        ws.cell(row=r, column=col).font = fnt(size=11, color=DORADO_T, underline="single")


r0 = 5

# ===================== Hoja: Detalle =====================
SD = "Detalle plantel-secundaria"
wsD = wb.create_sheet(SD)
hdrD = ["CCT Bachillerato", "Subsistema", "Plantel Bachillerato", "Turno Bachillerato", "Municipio Bachillerato",
        "Localidad Bachillerato", "Zona Criterio", "Ranking Cercanía", "Radio", "Distancia Lineal (km)",
        "Distancia Ruta (km)", "Factor Ruta / Lineal", "Mismo Municipio", "CCT Secundaria", "Nombre Secundaria",
        "Turno Secundaria", "Tipo Secundaria", "Sostenimiento Secundaria", "Municipio Secundaria",
        "Localidad Secundaria", "Asentamiento Secundaria", "CP Secundaria", "Fuente",
        "Ver ruta", "Lat. Bachillerato", "Lon. Bachillerato", "Lat. Secundaria", "Lon. Secundaria"]
nD = len(pares)
rngA, rngJ = f"$A${r0}:$A${r0 + nD - 1}", f"$J${r0}:$J${r0 + nD - 1}"
filasD = []
for i, (m, x, k, ru, rad) in enumerate(pares):
    r = r0 + i
    if rad == RX:
        fuente = "Cobertura: plantel más cercano (sin plantel a 4.5 km)"
    elif rad == R2:
        fuente = "Calculado" if m.get("h2") == 2 else "Archivo de 2 km"
    else:
        fuente = "Calculado" if m.get("h45") == 2 else "Archivo de 4.5 km (hoja 1)"
    filasD.append([m["c"], m["sb"], m["n"], m["t"], m["mu"], m["lc"], m["zo"],
                   f'=COUNTIFS({rngA},A{r},{rngJ},"<"&J{r})+1', rad, k, ru,
                   f'=IF(AND(ISNUMBER(K{r}),J{r}>0.05),K{r}/J{r},"")', f'=IF(E{r}=S{r},"SÍ","NO")',
                   x["c"], x["n"], x.get("t", ""), x.get("ti", ""), x.get("so", ""), x.get("mu", ""), x.get("lc", ""),
                   x.get("as", ""), x.get("cp", ""), fuente, gmaps_ruta(r, "Y", "Z", "AA", "AB"),
                   m["la"], m["lo"], x["la"], x["lo"]])
encabezado(wsD, "Detalle plantel – secundaria",
           "Una fila por par: secundarias a 4.5 km o menos en línea recta (archivos de 2 km y 4.5 km) y, para que toda secundaria tenga bachillerato, "
           "las que no tienen plantel a 4.5 km asignadas al más cercano. Orden: municipio, zona, subsistema, plantel y distancia lineal de menor a mayor.", len(hdrD))
ultD = tabla(wsD, "tDetalle", hdrD, filasD,
             [12, 18, 34, 13, 16, 20, 18, 9, 13, 10, 10, 9, 9, 13, 34, 12, 16, 14, 16, 20, 22, 8, 22, 10, 10, 10, 10, 10],
             formatos={8: "0", 10: "0.00", 11: "0.00", 12: "0.00", 25: "0.000000", 26: "0.000000", 27: "0.000000", 28: "0.000000"})
liga(wsD, 24, r0, ultD)
for val, col, bg in ((R2, GUINDA, ROSA), (R45, DORADO_T, CREMA), (RX, ROJO, ROJO_S)):
    wsD.conditional_formatting.add(f"I{r0}:I{ultD}", CellIsRule(operator="equal", formula=[f'"{val}"'], font=Font(name=F, color=col, bold=True), fill=fill(bg)))
wsD.conditional_formatting.add(f"L{r0}:L{ultD}", CellIsRule(operator="greaterThanOrEqual", formula=["2"], font=Font(name=F, color=ROJO, bold=True)))
for col in ("Y", "Z", "AA", "AB"):
    wsD.column_dimensions[col].outlineLevel = 1
R = lambda col: f"'{SD}'!${col}${r0}:${col}${ultD}"
NO_X = f'"<>{RX}"'   # criterio: excluir los pares de cobertura (> 4.5 km)

# ===================== Hoja: Vista por Plantel =====================
SV = "Vista por Plantel"
wsV = wb.create_sheet(SV)
hdrV = ["CCT Bachillerato", "Subsistema", "Plantel Bachillerato", "Turno", "Municipio", "Localidad", "Zona Criterio",
        "Matrícula", "Docentes", "Secundarias a 2 km", "Secundarias a 4.5 km", "Más cercana · lineal (km)",
        "Más cercana · ruta (km)", "Lineal prom. 4.5 km", "Ruta prom. 4.5 km", "Asignadas por cobertura (> 4.5 km)",
        "Total secundarias asignadas", "Fuente de los radios", "Secundarias asignadas (CCT · nombre · lineal / ruta; ▲ = cobertura > 4.5 km)",
        "Ver en mapa", "Latitud", "Longitud"]
filasV = []
for i, m in enumerate(orden_ms):
    r = r0 + i
    txt = "\n".join(f"{j + 1}. {SEC[s]['c']} · {SEC[s]['n']} · {k:.2f} / {('–' if ru is None else f'{ru:.2f}')} km" + (" ▲" if rad == RX else "")
                    for j, (s, k, ru, rad) in enumerate(m["_pares"]))
    fuente = "Calculado" if (m.get("h2") == 2 or m.get("h45") == 2) else "Archivos de 2 y 4.5 km"
    filasV.append([m["c"], m["sb"], m["n"], m["t"], m["mu"], m["lc"], m["zo"], m["ma"], m["dc"],
                   f'=COUNTIFS({R("A")},A{r},{R("I")},"{R2}")', f'=COUNTIFS({R("A")},A{r},{R("I")},{NO_X})',
                   f'=IFERROR(_xlfn.MINIFS({R("J")},{R("A")},A{r}),"")', f'=IFERROR(_xlfn.MINIFS({R("K")},{R("A")},A{r}),"")',
                   f'=IFERROR(AVERAGEIFS({R("J")},{R("A")},A{r},{R("I")},{NO_X}),"")', f'=IFERROR(AVERAGEIFS({R("K")},{R("A")},A{r},{R("I")},{NO_X}),"")',
                   f'=COUNTIFS({R("A")},A{r},{R("I")},"{RX}")', f'=COUNTIF({R("A")},A{r})',
                   fuente, txt, gmaps_punto(r, "U", "V"), m["la"], m["lo"]])
encabezado(wsV, "Vista por plantel de bachillerato",
           f"{len(MS)} planteles del proceso de ingreso 2027. Conteos y distancias con fórmulas a partir del detalle. Las secundarias a 4.5 km incluyen las de 2 km.", len(hdrV))
ultV = tabla(wsV, "tPlanteles", hdrV, filasV,
             [12, 18, 34, 14, 15, 20, 18, 10, 9, 10, 10, 11, 11, 10, 10, 12, 11, 20, 70, 11, 10, 10],
             formatos={8: "#,##0", 9: "#,##0", 10: "0", 11: "0", 12: "0.00", 13: "0.00", 14: "0.00", 15: "0.00", 16: "0", 17: "0", 21: "0.000000", 22: "0.000000"},
             wrap={19})
liga(wsV, 20, r0, ultV)
wsV.conditional_formatting.add(f"J{r0}:J{ultV}", DataBarRule(start_type="num", start_value=0, end_type="max", color=GUINDA))
wsV.conditional_formatting.add(f"K{r0}:K{ultV}", DataBarRule(start_type="num", start_value=0, end_type="max", color=DORADO))
wsV.conditional_formatting.add(f"J{r0}:J{ultV}", CellIsRule(operator="equal", formula=["0"], fill=fill(ROJO_S), font=Font(name=F, color=ROJO, bold=True)))
wsV.conditional_formatting.add(f"P{r0}:P{ultV}", CellIsRule(operator="greaterThan", formula=["0"], fill=fill(ROJO_S), font=Font(name=F, color=ROJO, bold=True)))
wsV.cell(row=4, column=19).comment = Comment("Todas las secundarias asignadas al plantel, de menor a mayor distancia lineal: las de 4.5 km o menos y, con ▲, "
                                             "las que no tienen ningún plantel a 4.5 km y se asignan a este por ser el más cercano.", "Estudios Territoriales")
RV = lambda col: f"'{SV}'!${col}${r0}:${col}${ultV}"

# ===================== Hoja: Cobertura Secundarias =====================
SC = "Cobertura Secundarias"
wsC = wb.create_sheet(SC)
mas_cerca = {}
for m, x, k, ru, rad in pares:
    if x["c"] not in mas_cerca or k < mas_cerca[x["c"]][1]:
        mas_cerca[x["c"]] = (m, k, ru)
hdrC = ["CCT Secundaria", "Nombre Secundaria", "Turno", "Tipo", "Sostenimiento", "Municipio", "Localidad",
        "Asentamiento", "CP", "Planteles a 2 km", "Planteles a 4.5 km", "Planteles asignados (total)", "Plantel más cercano (CCT)",
        "Plantel más cercano", "Más cercano · lineal (km)", "Más cercano · ruta (km)", "Ver en mapa", "Latitud", "Longitud"]
filasC = []
for i, x in enumerate(sorted(SEC, key=lambda x: (x.get("mu", ""), x.get("lc", ""), x["c"]))):
    r = r0 + i
    mc = mas_cerca.get(x["c"])
    filasC.append([x["c"], x["n"], x.get("t", ""), x.get("ti", ""), x.get("so", ""), x.get("mu", ""), x.get("lc", ""),
                   x.get("as", ""), x.get("cp", ""),
                   f'=COUNTIFS({R("N")},A{r},{R("I")},"{R2}")', f'=COUNTIFS({R("N")},A{r},{R("I")},{NO_X})', f'=COUNTIF({R("N")},A{r})',
                   mc[0]["c"] if mc else "", mc[0]["n"] if mc else "",
                   f'=IFERROR(_xlfn.MINIFS({R("J")},{R("N")},A{r}),"")',
                   f'=IFERROR(INDEX({R("K")},MATCH(1,INDEX(({R("N")}=A{r})*({R("J")}=O{r}),0),0)),"")',
                   gmaps_punto(r, "R", "S"), x["la"], x["lo"]])
encabezado(wsC, "Cobertura de secundarias",
           f"{len(SEC)} secundarias de los archivos de 2 km y 4.5 km (públicas, particulares y CONAFE). Todas tienen al menos un plantel asignado.", len(hdrC))
ultC = tabla(wsC, "tSecundarias", hdrC, filasC, [13, 34, 12, 16, 13, 16, 20, 22, 8, 10, 10, 11, 13, 32, 11, 11, 11, 10, 10],
             formatos={10: "0", 11: "0", 12: "0", 15: "0.00", 16: "0.00", 18: "0.000000", 19: "0.000000"})
liga(wsC, 17, r0, ultC)
wsC.conditional_formatting.add(f"J{r0}:J{ultC}", DataBarRule(start_type="num", start_value=0, end_type="max", color=GUINDA))
wsC.conditional_formatting.add(f"K{r0}:K{ultC}", DataBarRule(start_type="num", start_value=0, end_type="max", color=DORADO))
wsC.conditional_formatting.add(f"K{r0}:K{ultC}", CellIsRule(operator="equal", formula=["0"], fill=fill(ROJO_S), font=Font(name=F, color=ROJO, bold=True)))
RC = lambda col: f"'{SC}'!${col}${r0}:${col}${ultC}"

# ===================== Hoja: Por Municipio =====================
SM = "Por Municipio"
wsM = wb.create_sheet(SM)
munis = sorted({(m["mu"], m["zo"]) for m in MS})
hdrM = ["Municipio", "Zona", "Planteles", "Matrícula", "Pares a 2 km", "Pares a 4.5 km", "Prom. secundarias a 2 km",
        "Prom. secundarias a 4.5 km", "Ruta prom. a la más cercana (km)", "Secundarias del municipio con plantel a 2 km",
        "Secundarias del municipio", "Asignadas por cobertura (> 4.5 km)", "Ruta prom. de los pares (km)"]
filasM = []
for i, (mu, zo) in enumerate(munis):
    r = r0 + i
    filasM.append([mu, zo, f'=COUNTIF({RV("E")},A{r})', f'=SUMIF({RV("E")},A{r},{RV("H")})',
                   f'=COUNTIFS({R("E")},A{r},{R("I")},"{R2}")', f'=COUNTIFS({R("E")},A{r},{R("I")},{NO_X})',
                   f'=IF(C{r}>0,E{r}/C{r},"")', f'=IF(C{r}>0,F{r}/C{r},"")',
                   f'=IFERROR(AVERAGEIFS({RV("M")},{RV("E")},A{r}),"")',
                   f'=COUNTIFS({RC("F")},A{r},{RC("J")},">0")', f'=COUNTIF({RC("F")},A{r})',
                   f'=COUNTIFS({R("E")},A{r},{R("I")},"{RX}")', f'=IFERROR(AVERAGEIF({R("E")},A{r},{R("K")}),"")'])
encabezado(wsM, "Resumen por municipio", "Municipios con planteles del proceso de ingreso 2027. Todo se calcula con fórmulas.", len(hdrM))
ultM = tabla(wsM, "tMunicipios", hdrM, filasM, [22, 18, 10, 11, 10, 10, 12, 12, 13, 15, 12, 13, 12],
             formatos={3: "0", 4: '#,##0;-#,##0;"s/d"', 5: "0", 6: "0", 7: "0.0", 8: "0.0", 9: "0.00", 10: "0", 11: "0", 12: "0", 13: "0.00"})
wsM.conditional_formatting.add(f"C{r0}:C{ultM}", DataBarRule(start_type="num", start_value=0, end_type="max", color=GUINDA_OSC))

# ===================== Hoja: Tablero =====================
ST = "Tablero"
wsT = wb.create_sheet(ST, 0)
wsT.sheet_view.showGridLines = False
for col, w in zip("ABCDEFGHIJKLMN", [34, 12, 12, 12, 12, 13, 13, 13, 3, 36, 12, 12, 12, 16]):
    wsT.column_dimensions[col].width = w
wsT["A1"] = "Bachilleratos y secundarias cercanas — Proceso de ingreso a bachillerato 2027"
wsT["A1"].font = fnt(size=20, bold=True, color=GUINDA_OSC)
wsT["A2"] = (f"Planteles del proceso de ingreso 2027 y secundarias a 2 km y 4.5 km, con distancia en línea recta y por ruta real. "
             f"Departamento de Estudios Territoriales · SEGEY · {datetime.date.today():%d/%m/%Y}")
wsT["A2"].font = fnt(size=11, color=GRIS)
for c in range(1, 15):
    wsT.cell(row=3, column=c).fill = fill(GUINDA)
wsT.row_dimensions[3].height = 4

kpis = [
    ("Planteles", f"=COUNTA({RV('A')})", "#,##0", GUINDA_OSC),
    ("Matrícula", f"=SUM({RV('H')})", "#,##0", GUINDA_OSC),
    ("Secundarias con plantel a 2 km", f'=COUNTIF({RC("J")},">0")', "#,##0", GUINDA),
    ("Secundarias con plantel a 4.5 km", f'=COUNTIF({RC("K")},">0")', "#,##0", DORADO_T),
    ("Ruta prom. a la más cercana", f"=AVERAGE({RV('M')})", '0.00 "km"', GUINDA_OSC),
    ("Factor ruta / lineal", f"=SUM({R('K')})/SUM({R('J')})", '0.00"×"', GUINDA_OSC),
    ("Pares asignados", f"=COUNTA({R('A')})", "#,##0", GUINDA_OSC),
]
bloques = [("A", "A"), ("B", "C"), ("D", "E"), ("F", "G"), ("J", "J"), ("K", "M"), ("N", "N")]
for (lab, f, nf, col), (c1, c2) in zip(kpis, bloques):
    if c1 != c2:
        wsT.merge_cells(f"{c1}5:{c2}5")
        wsT.merge_cells(f"{c1}6:{c2}6")
    a = wsT[f"{c1}5"]
    a.value = lab
    a.font = fnt(size=10, bold=True, color=GRIS)
    a.alignment = CENTRO
    v = wsT[f"{c1}6"]
    v.value = f
    v.number_format = nf
    v.font = fnt(size=22, bold=True, color=col)
    v.alignment = CENTRO
    for cc in range(openpyxl.utils.column_index_from_string(c1), openpyxl.utils.column_index_from_string(c2) + 1):
        for rr in (5, 6):
            wsT.cell(row=rr, column=cc).fill = fill(SUAVE)
            wsT.cell(row=rr, column=cc).border = caja
wsT.row_dimensions[5].height = 30
wsT.row_dimensions[6].height = 38


def seccion(fila, texto, col=1):
    c = wsT.cell(row=fila, column=col, value=texto)
    c.font = fnt(size=13, bold=True, color=GUINDA_OSC)
    wsT.row_dimensions[fila].height = 22


def cab(fila, textos, col0=1):
    for j, t in enumerate(textos):
        c = wsT.cell(row=fila, column=col0 + j, value=t)
        c.font = fnt(bold=True, color="FFFFFF", size=10)
        c.fill = fill(GUINDA_OSC)
        c.alignment = CENTRO
        c.border = caja
    wsT.row_dimensions[fila].height = 42


def celda(fila, col, v, nf=None, bold=False, color="262626", bg=None):
    c = wsT.cell(row=fila, column=col, value=v)
    c.font = fnt(size=11, bold=bold, color=color)
    c.border = caja
    if nf:
        c.number_format = nf
        c.alignment = Alignment(horizontal="right", vertical="center")
    if bg:
        c.fill = fill(bg)
    return c


# 1. Por zona
f = 8
seccion(f, "1. Resumen por zona")
cab(f + 1, ["Zona", "Planteles", "Matrícula", "Prom. sec. a 2 km", "Prom. sec. a 4.5 km", "Ruta prom. más cercana (km)",
            "Lineal prom. 4.5 km (km)", "Ruta prom. 4.5 km (km)"])
zonas = [ZM, "Interior del Estado"]
for j, z in enumerate(zonas):
    r = f + 2 + j
    celda(r, 1, z, bold=True)
    celda(r, 2, f'=COUNTIF({RV("G")},A{r})', "0")
    celda(r, 3, f'=SUMIF({RV("G")},A{r},{RV("H")})', '#,##0;-#,##0;"s/d"')
    celda(r, 4, f'=AVERAGEIF({RV("G")},A{r},{RV("J")})', "0.0", color=GUINDA, bold=True)
    celda(r, 5, f'=AVERAGEIF({RV("G")},A{r},{RV("K")})', "0.0", color=DORADO_T, bold=True)
    celda(r, 6, f'=AVERAGEIF({RV("G")},A{r},{RV("M")})', "0.00")
    celda(r, 7, f'=AVERAGEIFS({R("J")},{R("G")},A{r},{R("I")},{NO_X})', "0.00")
    celda(r, 8, f'=AVERAGEIFS({R("K")},{R("G")},A{r},{R("I")},{NO_X})', "0.00")
rt = f + 4
celda(rt, 1, "TOTAL", bold=True, bg=SUAVE)
celda(rt, 2, f"=SUM(B{f + 2}:B{f + 3})", "0", bold=True, bg=SUAVE)
celda(rt, 3, f"=SUM(C{f + 2}:C{f + 3})", "#,##0", bold=True, bg=SUAVE)
celda(rt, 4, f"=AVERAGE({RV('J')})", "0.0", bold=True, color=GUINDA, bg=SUAVE)
celda(rt, 5, f"=AVERAGE({RV('K')})", "0.0", bold=True, color=DORADO_T, bg=SUAVE)
celda(rt, 6, f"=AVERAGE({RV('M')})", "0.00", bold=True, bg=SUAVE)
celda(rt, 7, f'=AVERAGEIF({R("I")},{NO_X},{R("J")})', "0.00", bold=True, bg=SUAVE)
celda(rt, 8, f'=AVERAGEIF({R("I")},{NO_X},{R("K")})', "0.00", bold=True, bg=SUAVE)

# 2. Línea recta contra ruta real
f = rt + 2
seccion(f, "2. Línea recta contra ruta real")
cab(f + 1, ["Radio (en línea recta)", "Pares", "Lineal prom. (km)", "Ruta prom. (km)", "Pares que por ruta superan el radio", "% que supera", "Factor ruta / lineal"])
for j, (rad, lim) in enumerate([(R2, 2), (R45, 4.5)]):
    r = f + 2 + j
    celda(r, 1, rad, bold=True, color=GUINDA if lim == 2 else DORADO_T)
    celda(r, 2, f'=COUNTIF({R("I")},A{r})', "#,##0")
    celda(r, 3, f'=AVERAGEIF({R("I")},A{r},{R("J")})', "0.00")
    celda(r, 4, f'=AVERAGEIF({R("I")},A{r},{R("K")})', "0.00")
    celda(r, 5, f'=COUNTIFS({R("I")},A{r},{R("K")},">{lim}")', "#,##0")
    celda(r, 6, f"=IF(B{r}>0,E{r}/B{r},0)", "0.0%")
    celda(r, 7, f'=SUMIF({R("I")},A{r},{R("K")})/SUMIF({R("I")},A{r},{R("J")})', '0.00"×"')
fila_radios = (f + 2, f + 3)

# 3. Por subsistema
f = f + 5
seccion(f, "3. Resumen por subsistema")
cab(f + 1, ["Subsistema", "Planteles", "Matrícula", "Prom. sec. a 2 km", "Prom. sec. a 4.5 km", "Ruta prom. más cercana (km)",
            "Secundarias asignadas", "Ruta prom. de los pares (km)"])
subs = sorted({m["sb"] for m in MS}, key=lambda s: (-sum(1 for m in MS if m["sb"] == s), s))
for j, sb in enumerate(subs):
    r = f + 2 + j
    celda(r, 1, sb, bold=True)
    celda(r, 2, f'=COUNTIF({RV("B")},A{r})', "0")
    celda(r, 3, f'=SUMIF({RV("B")},A{r},{RV("H")})', '#,##0;-#,##0;"s/d"')
    celda(r, 4, f'=AVERAGEIF({RV("B")},A{r},{RV("J")})', "0.0", color=GUINDA, bold=True)
    celda(r, 5, f'=AVERAGEIF({RV("B")},A{r},{RV("K")})', "0.0", color=DORADO_T, bold=True)
    celda(r, 6, f'=AVERAGEIF({RV("B")},A{r},{RV("M")})', "0.00")
    celda(r, 7, f'=COUNTIF({R("B")},A{r})', "#,##0")
    celda(r, 8, f'=IFERROR(AVERAGEIF({R("B")},A{r},{R("K")}),"")', "0.00")
fila_subs = (f + 2, f + 1 + len(subs))
rt = f + 2 + len(subs)
celda(rt, 1, "TOTAL", bold=True, bg=SUAVE)
celda(rt, 2, f"=SUM(B{fila_subs[0]}:B{fila_subs[1]})", "0", bold=True, bg=SUAVE)
celda(rt, 3, f"=SUM(C{fila_subs[0]}:C{fila_subs[1]})", "#,##0", bold=True, bg=SUAVE)
celda(rt, 4, f"=AVERAGE({RV('J')})", "0.0", bold=True, color=GUINDA, bg=SUAVE)
celda(rt, 5, f"=AVERAGE({RV('K')})", "0.0", bold=True, color=DORADO_T, bg=SUAVE)
celda(rt, 6, f"=AVERAGE({RV('M')})", "0.00", bold=True, bg=SUAVE)
celda(rt, 7, f"=SUM(G{fila_subs[0]}:G{fila_subs[1]})", "#,##0", bold=True, bg=SUAVE)
celda(rt, 8, f"=AVERAGE({R('K')})", "0.00", bold=True, bg=SUAVE)

# 4. Cobertura de secundarias (columnas J–N)
fa = 8
seccion(fa, "4. Cobertura de secundarias", col=10)
cab(fa + 1, ["Concepto", "Secundarias", "% del total"], col0=10)
wsT.merge_cells(start_row=fa + 1, start_column=12, end_row=fa + 1, end_column=13)
filas4 = [("Secundarias en los archivos", f"=COUNTA({RC('A')})"),
          ("Con plantel a 2 km o menos", f'=COUNTIF({RC("J")},">0")'),
          ("Con plantel a 4.5 km o menos", f'=COUNTIF({RC("K")},">0")'),
          ("Sin plantel a 4.5 km (al más cercano)", f'=COUNTIF({R("I")},"{RX}")'),
          ("Sin bachillerato asignado", f'=COUNTIF({RC("L")},0)')]
for j, (lab, fo) in enumerate(filas4):
    r = fa + 2 + j
    celda(r, 10, lab, bold=True, color=ROJO if j == 3 else "262626")
    celda(r, 11, fo, "#,##0", bold=True, color=ROJO if j == 3 else "262626")
    wsT.merge_cells(start_row=r, start_column=12, end_row=r, end_column=13)
    celda(r, 12, f"=K{r}/$K${fa + 2}", "0.0%")
r = fa + 7
celda(r, 10, "Distancia máx. de asignación por cobertura", bold=True)
wsT.merge_cells(start_row=r, start_column=10, end_row=r, end_column=10)
celda(r, 11, f'=_xlfn.MAXIFS({R("J")},{R("I")},"{RX}")', '0.00 "km"', bold=True, color=ROJO)
r += 2
wsT.merge_cells(start_row=r, start_column=10, end_row=r + 3, end_column=14)
c = wsT.cell(row=r, column=10, value=("Criterio: cada plantel tiene todas las secundarias a 4.5 km o menos en línea recta (sin tope). Toda secundaria debe tener "
                                         "bachillerato: si no tiene ningún plantel a 4.5 km, se asigna al más cercano en línea recta. "
                                         "La ruta real (OSRM / OpenStreetMap) se reporta aparte; la liga «Ver ruta» abre cada trayecto en Google Maps."))
c.font = fnt(size=10, italic=True, color=GRIS)
c.alignment = IZQW
r += 5
c = wsT.cell(row=r, column=10, value="Página interactiva (mapa y filtros)")
c.font = fnt(size=11, bold=True, color=GUINDA_OSC)
c = wsT.cell(row=r + 1, column=10, value=f'=HYPERLINK("{URL_PAGINA}","{URL_PAGINA.replace("https://", "")}")')
c.font = fnt(size=11, color=DORADO_T, underline="single")

# Gráficas
ch = BarChart()
ch.type = "bar"
ch.title = "Planteles por subsistema"
ch.add_data(Reference(wsT, min_col=2, min_row=fila_subs[0] - 1, max_row=fila_subs[1]), titles_from_data=True)
ch.set_categories(Reference(wsT, min_col=1, min_row=fila_subs[0], max_row=fila_subs[1]))
ch.x_axis.scaling.orientation = "maxMin"
ch.series[0].graphicalProperties.solidFill = GUINDA_OSC
ch.series[0].graphicalProperties.line.solidFill = GUINDA_OSC
ch.dataLabels = DataLabelList()
ch.dataLabels.showVal = True
ch.dataLabels.showSerName = ch.dataLabels.showCatName = ch.dataLabels.showLegendKey = False
ch.legend = None
ch.height, ch.width = 9.5, 13
ch.roundedCorners = False
ch.x_axis.delete = ch.y_axis.delete = False
ch.title.overlay = False
ch.y_axis.majorGridlines = None
wsT.add_chart(ch, f"J{r + 3}")

ch2 = BarChart()
ch2.type = "col"
ch2.title = "Línea recta contra ruta (km promedio)"
ch2.add_data(Reference(wsT, min_col=3, max_col=4, min_row=fila_radios[0] - 1, max_row=fila_radios[1]), titles_from_data=True)
ch2.set_categories(Reference(wsT, min_col=1, min_row=fila_radios[0], max_row=fila_radios[1]))
for sr, colr in zip(ch2.series, ("9CA3AF", GUINDA_OSC)):
    sr.graphicalProperties.solidFill = colr
    sr.graphicalProperties.line.solidFill = colr
ch2.dataLabels = DataLabelList()
ch2.dataLabels.showVal = True
ch2.dataLabels.showSerName = ch2.dataLabels.showCatName = ch2.dataLabels.showLegendKey = False
ch2.height, ch2.width = 7.5, 13
ch2.roundedCorners = False
ch2.x_axis.delete = ch2.y_axis.delete = False
ch2.title.overlay = False
ch2.legend.position = "b"
ch2.legend.overlay = False
ch2.y_axis.majorGridlines = None
ch2.y_axis.numFmt = "0.0"
wsT.add_chart(ch2, f"J{r + 23}")
wsT.freeze_panes = "A4"
wsT.page_setup.orientation = "landscape"
wsT.page_setup.fitToWidth = 1
wsT.page_setup.fitToHeight = 0
wsT.sheet_properties.pageSetUpPr.fitToPage = True

# ===================== Hoja: RESUMEN (la tabla dinámica la crea Excel) =====================
SR = "RESUMEN"
wsR = wb.create_sheet(SR, 1)
wsR.sheet_view.showGridLines = False
wsR["A1"] = "Tabla dinámica: planteles y secundarias asignadas"
wsR["A1"].font = fnt(size=18, bold=True, color=GUINDA_OSC)
wsR["A2"] = ("Filas: Municipio › Zona › Subsistema › CCT y Plantel › Localidad › Secundaria (de menor a mayor distancia lineal). "
             "Valores: distancia lineal y por ruta. Usa las segmentaciones para filtrar.")
wsR["A2"].font = fnt(size=11, color=GRIS)
for c in range(1, 13):
    wsR.cell(row=3, column=c).fill = fill(GUINDA)
wsR.row_dimensions[3].height = 4

# ===================== Hoja: Notas y método =====================
SN = "Notas y método"
wsN = wb.create_sheet(SN)
wsN.sheet_view.showGridLines = False
wsN.column_dimensions["A"].width = 30
wsN.column_dimensions["B"].width = 110
wsN["A1"] = "Notas, fuentes y método"
wsN["A1"].font = fnt(size=18, bold=True, color=GUINDA_OSC)
for c in range(1, 3):
    wsN.cell(row=3, column=c).fill = fill(GUINDA)
wsN.row_dimensions[3].height = 4
calc = [m for m in MS if m.get("h2") == 2 or m.get("h45") == 2]
n_x = sum(1 for p in pares if p[4] == RX)
notas = [
    ("Propósito", "Resumen de los planteles de bachillerato del proceso de ingreso 2027 y de las secundarias que tienen a 2 km y a 4.5 km, "
                  "con la distancia en línea recta y la distancia real por carretera, para planear la promoción del bachillerato en secundarias."),
    ("Planteles", f"Los {len(MS)} planteles de la lista del proceso de ingreso 2027 (de ese archivo solo se toma qué planteles se usan). "
                  "Nombre, subsistema, turno, municipio y localidad salen del catálogo SIC y de los archivos de 2 y 4.5 km; la matrícula y los docentes, del archivo de 2 km."),
    ("Zona", "Zona Metropolitana = municipios de Mérida, Kanasín, Umán, Conkal y Ucú; el resto, Interior del Estado."),
    ("Secundarias a 2 km", "«Presentación 31ES INEGI Media Superior Escuelas Secundaria Vertical»: secundarias a 2 km o menos de cada plantel, con asentamiento INEGI."),
    ("Secundarias a 4.5 km", "«4.5KM-MEDIA SUPERIOR Y SECUNDARIASv2.xlsx», hoja 1. Las de 4.5 km incluyen las de 2 km."),
    ("Criterio de asignación", f"Máximo 4.5 km en línea recta, sin tope por plantel. Toda secundaria de los archivos de 2 y 4.5 km debe tener bachillerato: "
                               f"las {n_x} que no tienen ningún plantel a 4.5 km se asignan a su plantel más cercano en línea recta "
                               f"(Radio = «{RX}»; ▲ en «Vista por Plantel»)."),
    ("Planteles calculados", f"{len(calc)} planteles no vienen en el archivo de 2 km o de 4.5 km; sus secundarias se calcularon con las mismas secundarias y la distancia "
                             "en línea recta (aproximado): " + "; ".join(f"{m['c']} {m['n']}" for m in calc) + "."),
    ("Ubicación", "Catálogo SIC (a_ctba_sic_v2): municipio, localidad, asentamiento, código postal, domicilio y coordenadas."),
    ("Distancia lineal", "Distancia en línea recta (fórmula del gran círculo) entre las coordenadas del plantel y de la secundaria."),
    ("Distancia por ruta", "Camino más corto en automóvil por carreteras y calles, calculado con OSRM (Open Source Routing Machine) sobre datos de OpenStreetMap, "
                           "noviembre de 2026. No considera tráfico. La liga «Ver ruta» abre el mismo trayecto en Google Maps."),
    ("Conteo de secundarias", "Las secundarias se cuentan por clave de centro de trabajo y turno, igual que en los archivos de origen."),
    ("Radio «≤ 2 km»", "Viene del archivo de 2 km. Algunas secundarias aparecen con 2.01–2.10 km lineales porque ese archivo usa coordenadas propias."),
    ("Cómo actualizar", "En la carpeta Bachilleratos_Secundarias_Radio: py build_datos.py → py excel_resumen.py. Las rutas ya consultadas se guardan en rutas_osrm.json."),
    ("Hojas", "Tablero (indicadores y gráficas) · RESUMEN (tabla dinámica con segmentaciones) · Detalle plantel-secundaria (un renglón por par) · "
              "Vista por Plantel · Cobertura Secundarias · Por Municipio · Notas y método."),
]
for j, (a, b) in enumerate(notas):
    r = 5 + j
    ca = wsN.cell(row=r, column=1, value=a)
    ca.font = fnt(size=11, bold=True, color=GUINDA_OSC)
    ca.alignment = IZQW
    cb = wsN.cell(row=r, column=2, value=b)
    cb.font = fnt(size=11)
    cb.alignment = IZQW
    wsN.row_dimensions[r].height = 15 * max(1, math.ceil(len(b) / 120)) + 6

orden = [ST, SR, SD, SV, SC, SM, SN]
wb._sheets = [wb[n] for n in orden]
for n, ccol in {ST: GUINDA, SR: GUINDA, SD: DORADO, SV: DORADO, SC: DORADO, SM: DORADO, SN: GRIS}.items():
    wb[n].sheet_properties.tabColor = ccol
wb.properties.creator = "Departamento de Estudios Territoriales · SEGEY"
wb.properties.title = "Bachilleratos y secundarias cercanas 2027"
wb.save(OUT)
print("openpyxl OK ->", OUT, "| pares:", nD, f"(cobertura: {n_x})", "| planteles:", len(filasV), "| secundarias:", len(filasC))

# ===================== Excel: tabla dinámica, segmentaciones y recálculo =====================
import win32com.client as win32
import pythoncom

wbx = None
xl = win32.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
try:
    wbx = xl.Workbooks.Open(OUT)
    bgr = lambda h: int(h[4:6] + h[2:4] + h[0:2], 16)
    tc = wbx.Theme.ThemeColorScheme      # acento 1 = guinda, acento 2 = dorado
    tc.Colors(5).RGB = bgr(GUINDA)
    tc.Colors(6).RGB = bgr(DORADO)
    ws = wbx.Worksheets(SR)
    pc = wbx.PivotCaches().Create(SourceType=1, SourceData="tDetalle")
    pt = pc.CreatePivotTable(TableDestination=ws.Range("A16"), TableName="ptResumen")
    pt.RowAxisLayout(1)  # tabular
    campos = ["Municipio Bachillerato", "Zona Criterio", "Subsistema", "CCT Bachillerato", "Plantel Bachillerato",
              "Localidad Bachillerato", "CCT Secundaria", "Nombre Secundaria", "Municipio Secundaria", "Localidad Secundaria"]
    for k, nom in enumerate(campos, 1):
        pf = pt.PivotFields(nom)
        pf.Orientation = 1
        pf.Position = k
    for nom in campos:
        try:
            pt.PivotFields(nom).Subtotals = tuple([False] * 12)
        except Exception as e:
            print("subtotales:", nom, e)
    d1 = pt.AddDataField(pt.PivotFields("Distancia Lineal (km)"), "Lineal (km)", -4157)
    d1.NumberFormat = "0.00"
    d2 = pt.AddDataField(pt.PivotFields("Distancia Ruta (km)"), "Ruta (km)", -4157)
    d2.NumberFormat = "0.00"
    pt.PivotFields("CCT Secundaria").AutoSort(1, "Lineal (km)")
    pt.ColumnGrand = False
    pt.RowGrand = False
    pt.TableStyle2 = "PivotStyleLight16"
    pt.HasAutoFormat = False
    for col, w in zip("ABCDEFGHIJKL", [18, 18, 20, 12, 30, 20, 13, 32, 16, 20, 10, 10]):
        ws.Columns(col).ColumnWidth = w     # antes de colocar las segmentaciones (se mueven con las celdas)
    top = ws.Range("A4").Top + 6
    for campo, cap in [("Zona Criterio", "Zona"), ("Subsistema", "Subsistema"), ("Radio", "Radio"), ("Municipio Bachillerato", "Municipio")]:
        sc = wbx.SlicerCaches.Add2(pt, campo)
        sl = sc.Slicers.Add(ws, pythoncom.Missing, "seg" + cap, cap, top, 4, 190, 165)
        sl.Style = "SlicerStyleLight1"
    x_seg = 4
    for k in range(1, ws.Shapes.Count + 1):
        shp = ws.Shapes.Item(k)
        shp.Top, shp.Width, shp.Height = top, 190, 165
        shp.Left = x_seg
        x_seg = shp.Left + shp.Width + 10
    xl.CalculateFull()
    wbx.Worksheets(ST).Activate()
    wbx.Save()
    print("Excel OK: tabla dinámica con", pt.TableRange1.Rows.Count, "renglones y", wbx.SlicerCaches.Count, "segmentaciones")
    wbx.Close(SaveChanges=False)
    wbx = None
finally:
    try:
        if wbx is not None:
            wbx.Close(SaveChanges=False)
    except Exception:
        pass
    try:
        xl.Quit()
    except Exception:
        pass
