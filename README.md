# Reporte de Excesos de Velocidad

Página web para generar el reporte mensual de excesos de velocidad de **Transportes Jirena S.A.C.**, a partir de los reportes del GPS (o de la plantilla JIR-FO-OP-21 ya llenada).

## Qué hace

1. **Extraer:** lee uno o varios Excel y detecta sola la hoja y las columnas (placa, conductor, fecha, hora, ubicación o coordenadas, velocidad, límite).
2. **Transformar:**
   - Unifica las coordenadas, que llegan en dos formatos: `-16.43, -72.29` y pegadas como `-12.62-76.67`. También acepta enlaces de Google Maps.
   - Corrige las horas guardadas como fecha completa.
   - Unifica los nombres de conductores (espacios de más, comas).
   - Quita los registros duplicados entre archivos.
   - Calcula el exceso y la falta: leve 1–5 km/h, moderado 6–10 km/h, grave más de 10 km/h. Los rangos se pueden ajustar.
3. **Tramos:** agrupa los excesos cercanos entre sí (radio de 1.5 km, DBSCAN) y nombra cada tramo con la localidad más cercana (datos de GeoNames). Puedes renombrar cada tramo; el nombre se guarda en el navegador y se puede exportar o importar como archivo.
4. **Reporte:** indicadores, mapa de tramos críticos, ranking de tramos, gráficos por falta, hora, día, conductor y unidad, y el detalle de cada exceso.
5. **Descarga:** un Excel con tres hojas: `Plantilla exceso` (formato JIR-FO-OP-21), `Estadisticas` y `Tramos`.

Los archivos se procesan **en el navegador**: no se suben a ningún servidor.

## Publicar con GitHub Pages

1. Sube estos archivos a la raíz del repositorio: `index.html`, `app.js`, `etl.js`, `styles.css`, `lugares-pe.js`, `logo-jirena.png`, `.nojekyll`.
2. En GitHub, ve a **Settings → Pages**.
3. En "Source" elige **Deploy from a branch**, rama `main`, carpeta `/ (root)`, y guarda.
4. En uno o dos minutos la página queda en `https://<tu-usuario>.github.io/<repositorio>/`.

> **Importante:** no subas los Excel con datos reales al repositorio. Tienen nombres de conductores y, si el repositorio es público, cualquiera podría verlos. El archivo `.gitignore` ya los excluye.

## Probar en tu computadora

```bash
python -m http.server 8000
# abrir http://localhost:8000
```

## Archivos

| Archivo | Contenido |
|---|---|
| `index.html` | Estructura de la página |
| `styles.css` | Estilos |
| `etl.js` | Lectura, limpieza, clasificación y agrupación en tramos |
| `app.js` | Mapa, gráficos, tablas y descarga del Excel |
| `lugares-pe.js` | 1,682 localidades del Perú para nombrar los tramos ([GeoNames](https://www.geonames.org/), CC BY 4.0) |

Librerías (por CDN): [SheetJS](https://sheetjs.com/) para leer y escribir Excel, [Leaflet](https://leafletjs.com/) con OpenStreetMap para el mapa y [Chart.js](https://www.chartjs.org/) para los gráficos.
