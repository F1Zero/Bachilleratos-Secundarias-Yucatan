# Bachilleratos y secundarias cercanas · Yucatán

Página interactiva para consultar los planteles de bachillerato del proceso de ingreso 2027 en Yucatán y las escuelas secundarias que tienen a **2 km** y a **4.5 km**, con distancia en línea recta y **por ruta**.

**Ver la página:** https://f1zero.github.io/Bachilleratos-Secundarias-Yucatan/

## Qué hace

- Filtra por municipio, localidad, zona (Metropolitana / Interior), subsistema o búsqueda libre (CCT, nombre, asentamiento).
- Muestra un resumen: número de planteles, matrícula y secundarias distintas a 2 km y a 4.5 km.
- Al elegir un plantel muestra su ficha y dos listas de secundarias (a 2 km y a 4.5 km), con asentamiento, CP, localidad, municipio, distancia lineal y distancia por ruta.
- Mapa con los radios de 2 km y 4.5 km y las secundarias alrededor.
- Descarga en CSV del resumen y de las secundarias de cada plantel.

## Criterio

- Máximo **4.5 km** en línea recta, sin tope de secundarias por plantel.
- **Solo secundarias activas** del catálogo SIC (nivel secundaria, status activo). Las activas que no vienen en los archivos de 2 y 4.5 km se agregan con distancia en línea recta calculada (marcadas con «SIC» en la página).
- **Toda secundaria tiene bachillerato:** las secundarias que no tienen ningún plantel a 4.5 km se asignan a su plantel más cercano en línea recta (aparecen en la ficha de ese plantel).

## Fuentes

- **Planteles:** los 353 planteles de la lista del proceso de ingreso 2027 (de esa lista solo se toma qué planteles se usan).
- **Datos de los planteles** (nombre, subsistema, turno, municipio, localidad, ubicación): catálogo SIC y archivos de 2 y 4.5 km. Matrícula y docentes: archivo de 2 km.
- **Zona:** Zona Metropolitana = municipios de Mérida, Kanasín, Umán, Conkal y Ucú; el resto, Interior del Estado.
- **2 km:** «Presentación 31ES INEGI Media Superior Escuelas Secundaria Vertical».
- **4.5 km:** «4.5KM-MEDIA SUPERIOR Y SECUNDARIASv2», hoja 1.
- Los planteles que no vienen en el archivo de 2 km o de 4.5 km se calcularon con las mismas secundarias y la distancia en línea recta (aproximado; marcados con \*).
- **Distancia por ruta:** camino más corto en auto por carreteras y calles, calculado con [OSRM](https://project-osrm.org/) sobre datos de OpenStreetMap. Se guarda en `rutas_osrm.json` para no volver a consultar el servicio.

## Regenerar

```
py build_datos.py
```

Genera `index.html` (y `Bachilleratos y Secundarias.html`) a partir de `_template.html` y los archivos de Excel (rutas al inicio del script). Requiere `openpyxl` y `pyxlsb`. Solo consulta OSRM para los pares plantel–secundaria que no estén en `rutas_osrm.json`.

El resumen en Excel se genera con `py excel_resumen.py` (requiere Excel y `pywin32`).

Departamento de Estudios Territoriales · SEGEY
