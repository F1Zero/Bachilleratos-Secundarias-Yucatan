# Bachilleratos y secundarias cercanas · Yucatán

Página interactiva para consultar los planteles de bachillerato del proceso de ingreso 2027 en Yucatán y las escuelas secundarias que tienen a **2 km** y a **4.5 km**, con distancia en línea recta y **por ruta**.

**Ver la página:** https://f1zero.github.io/Bachilleratos-Secundarias-Yucatan/

## Qué hace

- Filtra por municipio, localidad, zona (Metropolitana / Interior), subsistema o búsqueda libre (CCT, nombre, asentamiento, especialidad).
- Muestra un resumen: número de planteles, matrícula y secundarias distintas a 2 km y a 4.5 km.
- Al elegir un plantel muestra su ficha (incluye especialidades) y dos listas de secundarias (a 2 km y a 4.5 km), con asentamiento, CP, localidad, municipio, distancia lineal y distancia por ruta.
- Mapa con los radios de 2 km y 4.5 km y las secundarias alrededor.
- Descarga en CSV del resumen y de las secundarias de cada plantel.

## Fuentes

- **Planteles:** «Asignación Secundarias Promoción Bachillerato 2027», hoja «Vista por Bachillerato» (353 planteles). Solo se incluyen estos planteles; nombre, subsistema, turno, municipio, localidad y zona vienen de esa lista.
- **2 km:** «Presentación 31ES INEGI Media Superior Escuelas Secundaria Vertical».
- **4.5 km:** «4.5KM-MEDIA SUPERIOR Y SECUNDARIASv2», hoja 1.
- Los planteles que no vienen en el archivo de 2 km o de 4.5 km (Bachilleratos Nacionales «Margarita Maza», Bachillerato Tecnológico Metropolitano y UTR del Sur) se calcularon con las mismas secundarias y la distancia en línea recta (aproximado; marcados con \*).
- **Ubicación (municipio, localidad, asentamiento, domicilio):** catálogo SIC.
- **Distancia por ruta:** camino más corto en auto por carreteras y calles, calculado con [OSRM](https://project-osrm.org/) sobre datos de OpenStreetMap. Se guarda en `rutas_osrm.json` para no volver a consultar el servicio.

## Regenerar

```
py build_datos.py
```

Genera `index.html` (y `Bachilleratos y Secundarias.html`) a partir de `_template.html` y los archivos de Excel (rutas al inicio del script). Requiere `openpyxl` y `pyxlsb`. Solo consulta OSRM para los pares plantel–secundaria que no estén en `rutas_osrm.json`.

Departamento de Estudios Territoriales · SEGEY
