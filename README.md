# Bachilleratos y secundarias cercanas · Yucatán

Página interactiva para consultar las escuelas de media superior de Yucatán y las escuelas secundarias que tienen a **2 km** y a **4.5 km** de distancia.

**Ver la página:** https://f1zero.github.io/Bachilleratos-Secundarias-Yucatan/

## Qué hace

- Filtra por municipio, localidad, sostenimiento, subsistema o búsqueda libre (CCT, nombre, asentamiento).
- Muestra un resumen: número de bachilleratos, matrícula y secundarias distintas a 2 km y a 4.5 km.
- Al elegir un bachillerato muestra su ficha y dos listas de secundarias (a 2 km y a 4.5 km), con asentamiento, CP, localidad, municipio y distancia en línea recta.
- Mapa con los radios de 2 km y 4.5 km y las secundarias alrededor.
- Descarga en CSV del resumen y de las secundarias de cada bachillerato.

## Fuentes

- **2 km:** «Presentación 31ES INEGI Media Superior Escuelas Secundaria Vertical». Los Bachilleratos Nacionales «Margarita Maza», que no vienen en ese archivo, se calcularon con las mismas secundarias y la distancia en línea recta (aproximado; marcados con \*).
- **4.5 km:** «4.5KM-MEDIA SUPERIOR Y SECUNDARIASv2», hoja 1 (bachilleratos públicos). Para los bachilleratos privados y autónomos, que no vienen en esa hoja, el radio de 4.5 km se calculó con las mismas secundarias (marcados con \*). Ese cálculo reproduce exactamente la hoja 1 en los 370 bachilleratos públicos.
- **Ubicación (municipio, localidad, asentamiento, domicilio):** catálogo SIC.

## Regenerar

```
py build_datos.py
```

Genera `index.html` (y `Bachilleratos y Secundarias.html`) a partir de `_template.html` y los archivos de Excel (rutas al inicio del script). Requiere `openpyxl` y `pyxlsb`.

Departamento de Estudios Territoriales · SEGEY
