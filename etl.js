/* ETL de excesos de velocidad: lectura, limpieza, clasificación y agrupación en tramos.
   Funciona en el navegador (window.ETL) y en Node (module.exports) para pruebas. */
(function (root) {
  'use strict';

  const norm = s => String(s == null ? '' : s)
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toUpperCase().replace(/\s+/g, ' ').trim();

  // ---------- Detección de columnas ----------
  const COLS = {
    fecha:     h => /^FECHA\b/.test(h) && !/HORA/.test(h) || h === 'FECHA',
    fechahora: h => /FECHA.*HORA|HORA.*FECHA/.test(h),
    placa:     h => /PLACA|UNIDAD|VEHICULO|PATENTE/.test(h),
    conductor: h => /CONDUCTOR|CHOFER|OPERADOR/.test(h),
    hora:      h => /^HORA\b/.test(h) && !/FECHA/.test(h),
    ubicacion: h => /UBICACION|COORDENADA|POSICION/.test(h),
    gmaps:     h => /GOOGLE|MAPS|ENLACE|LINK/.test(h),
    lat:       h => /^LAT(ITUD)?\b/.test(h),
    lon:       h => /^LON(GITUD)?\b|^LNG\b/.test(h),
    velocidad: h => /VELOCIDAD|\bVEL\b|KM\/H/.test(h) && !/LIMITE|EXCESO|PERMITIDA/.test(h),
    limite:    h => /LIMITE|PERMITIDA|MAXIMA PERMITIDA/.test(h),
  };

  function findHeader(rows) {
    for (let i = 0; i < Math.min(rows.length, 25); i++) {
      const h = (rows[i] || []).map(norm);
      const has = k => h.some(x => x && COLS[k](x));
      if (has('placa') && has('velocidad') && (has('ubicacion') || has('lat') || has('gmaps'))) {
        const idx = {};
        h.forEach((x, j) => {
          if (!x) return;
          for (const k in COLS) if (idx[k] == null && COLS[k](x)) { idx[k] = j; break; }
        });
        // "VELOCIDAD MAXIMA" de un reporte = velocidad realizada; preferir "REALIZADA" si existe
        const real = h.findIndex(x => /VELOCIDAD (REALIZADA|REGISTRADA|ALCANZADA)/.test(x));
        if (real >= 0) idx.velocidad = real;
        return { row: i, idx };
      }
    }
    return null;
  }

  // ---------- Conversión de valores ----------
  const EPOCH = Date.UTC(1899, 11, 30);
  const pad = n => String(n).padStart(2, '0');
  const isoFromUTC = ms => { const d = new Date(ms); return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`; };

  function parseFecha(v) {
    if (v == null || v === '') return null;
    if (typeof v === 'number') return v >= 1 ? isoFromUTC(EPOCH + Math.floor(v) * 864e5) : null;
    if (v instanceof Date) return `${v.getFullYear()}-${pad(v.getMonth() + 1)}-${pad(v.getDate())}`;
    const s = String(v).trim();
    let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return `${m[1]}-${pad(+m[2])}-${pad(+m[3])}`;
    m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
    if (m) { const y = m[3].length === 2 ? '20' + m[3] : m[3]; return `${y}-${pad(+m[2])}-${pad(+m[1])}`; }
    return null;
  }

  function parseHora(v) {
    if (v == null || v === '') return null;
    if (typeof v === 'number') {
      let f = v - Math.floor(v); let s = Math.round(f * 86400) % 86400;
      return `${pad(Math.floor(s / 3600))}:${pad(Math.floor(s / 60) % 60)}:${pad(s % 60)}`;
    }
    if (v instanceof Date) return `${pad(v.getHours())}:${pad(v.getMinutes())}:${pad(v.getSeconds())}`;
    const m = String(v).match(/(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp]\.?\s?[Mm]\.?)?/);
    if (!m) return null;
    let h = +m[1]; const ap = m[4] && m[4][0].toUpperCase();
    if (ap === 'P' && h < 12) h += 12; if (ap === 'A' && h === 12) h = 0;
    return `${pad(h)}:${m[2]}:${m[3] || '00'}`;
  }

  // Acepta "-16.43, -72.29", "-12.6265-76.6702" (pegado), enlaces de Google Maps, etc.
  function parseCoord(v) {
    if (v == null) return null;
    const s = String(v);
    const nums = s.match(/-?\d+(?:\.\d+)?/g);
    if (!nums || nums.length < 2) return null;
    let a = parseFloat(nums[0]), b = parseFloat(nums[1]);
    if (Math.abs(a) > Math.abs(b) && Math.abs(a) > 30) [a, b] = [b, a]; // lon, lat invertidos
    if (!isFinite(a) || !isFinite(b) || Math.abs(a) > 90 || Math.abs(b) > 180 || (a === 0 && b === 0)) return null;
    return { lat: a, lon: b, fmt: /,|;|\s/.test(s.replace(/^.*q=/, '').trim()) ? 'separado' : 'pegado' };
  }

  const num = v => { if (typeof v === 'number') return v; const m = String(v == null ? '' : v).replace(',', '.').match(/-?\d+(\.\d+)?/); return m ? parseFloat(m[0]) : NaN; };
  const cleanName = s => norm(String(s == null ? '' : s).replace(/,/g, ' ')).replace(/^-+$/, '');

  function guessOrigen(fileName, sheetName) {
    const s = norm(fileName + ' ' + sheetName);
    if (/GOLD/.test(s)) return 'GoldCar';
    if (/TRACK/.test(s)) return 'Tracklog';
    return 'Plantilla';
  }

  /** Extrae los registros de una hoja (matriz de filas). */
  function extractSheet(rows, meta) {
    const hd = findHeader(rows);
    if (!hd) return null;
    const { idx } = hd; const out = []; const issues = { coord: 0, hora: 0, sinVel: 0 };
    for (let i = hd.row + 1; i < rows.length; i++) {
      const r = rows[i] || [];
      const g = k => idx[k] == null ? null : r[idx[k]];
      const placa = cleanName(g('placa'));
      if (!placa) continue;
      let c = null;
      if (idx.lat != null && idx.lon != null) c = parseCoord(`${g('lat')},${g('lon')}`);
      if (!c && idx.ubicacion != null) c = parseCoord(g('ubicacion'));
      if (!c && idx.gmaps != null) c = parseCoord(g('gmaps'));
      if (!c) { issues.coord++; continue; }
      const vel = num(g('velocidad'));
      if (!isFinite(vel)) { issues.sinVel++; continue; }
      const fv = g('fecha') != null ? g('fecha') : g('fechahora');
      const fecha = parseFecha(fv);
      const hora = parseHora(idx.hora != null ? g('hora') : g('fechahora')) || (typeof fv === 'number' && fv % 1 ? parseHora(fv) : null);
      if (!hora) issues.hora++;
      const lim = num(g('limite'));
      out.push({
        fecha, hora: hora || '00:00:00', placa, cond: cleanName(g('conductor')) || 'SIN CONDUCTOR',
        lat: c.lat, lon: c.lon, fmt: c.fmt, vel: Math.round(vel), lim: isFinite(lim) && lim > 0 ? Math.round(lim) : null,
        origen: meta.origen, archivo: meta.file, hoja: meta.sheet,
      });
    }
    return { rows: out, issues, header: hd.row + 1 };
  }

  /** Calcula exceso y falta; descarta registros sin exceso. */
  function classify(rows, cfg) {
    const out = [];
    for (const r of rows) {
      const lim = r.lim || cfg.limite;
      const exc = r.vel - lim;
      if (exc <= 0) continue;
      out.push(Object.assign({}, r, { limUsado: lim, exc, falta: exc <= cfg.leve ? 'LEVE' : exc <= cfg.moderado ? 'MODERADO' : 'GRAVE', h: +r.hora.slice(0, 2) }));
    }
    return out;
  }

  /** Quita duplicados (misma placa, fecha, hora y velocidad) que aparecen en dos archivos. */
  function dedupe(rows) {
    const seen = new Set(); const out = [];
    for (const r of rows) { const k = `${r.placa}|${r.fecha}|${r.hora}|${r.vel}`; if (seen.has(k)) continue; seen.add(k); out.push(r); }
    return { rows: out, removed: rows.length - out.length };
  }

  // ---------- Geografía ----------
  const R = 6371;
  function hav(a, b, c, d) {
    const r = Math.PI / 180, dl = (c - a) * r, dn = (d - b) * r;
    const x = Math.sin(dl / 2) ** 2 + Math.cos(a * r) * Math.cos(c * r) * Math.sin(dn / 2) ** 2;
    return 2 * R * Math.asin(Math.sqrt(x));
  }

  /** DBSCAN con índice de rejilla. Devuelve etiqueta de grupo por punto (-1 = aislado). */
  function dbscan(pts, epsKm, minPts) {
    const n = pts.length, lab = new Int32Array(n).fill(-2);
    const cs = epsKm / 111; const grid = new Map();
    const key = (i, j) => i + ':' + j;
    pts.forEach((p, k) => { const kk = key(Math.floor(p.lat / cs), Math.floor(p.lon / cs)); (grid.get(kk) || grid.set(kk, []).get(kk)).push(k); });
    const nb = k => {
      const p = pts[k], gi = Math.floor(p.lat / cs), gj = Math.floor(p.lon / cs), res = [];
      for (let a = -1; a <= 1; a++) for (let b = -2; b <= 2; b++) {
        const cell = grid.get(key(gi + a, gj + b)); if (!cell) continue;
        for (const q of cell) if (hav(p.lat, p.lon, pts[q].lat, pts[q].lon) <= epsKm) res.push(q);
      }
      return res;
    };
    let c = 0;
    for (let k = 0; k < n; k++) {
      if (lab[k] !== -2) continue;
      const N = nb(k);
      if (N.length < minPts) { lab[k] = -1; continue; }
      lab[k] = c; const seed = N.slice();
      for (let s = 0; s < seed.length; s++) {
        const q = seed[s];
        if (lab[q] === -1) lab[q] = c;
        if (lab[q] !== -2) continue;
        lab[q] = c; const N2 = nb(q);
        if (N2.length >= minPts) for (const z of N2) if (lab[z] === -2 || lab[z] === -1) seed.push(z);
      }
      c++;
    }
    return lab;
  }

  function nearestPlace(lat, lon, places) {
    let best = null, bd = Infinity;
    for (const p of places) { const d = hav(lat, lon, p[0], p[1]); if (d < bd) { bd = d; best = p; } }
    return best ? { lugar: best[2], prov: best[3], reg: best[4], dist: bd } : { lugar: '—', prov: '', reg: '', dist: 0 };
  }

  const tramoKey = (lat, lon) => `${lat.toFixed(2)},${lon.toFixed(2)}`;

  /** Agrupa los excesos en tramos y arma la ficha de cada uno. */
  function buildTramos(rows, cfg, places) {
    const lab = dbscan(rows, cfg.eps, cfg.minPts);
    const groups = new Map();
    rows.forEach((r, i) => { r.tramo = null; if (lab[i] >= 0) (groups.get(lab[i]) || groups.set(lab[i], []).get(lab[i])).push(r); });
    const tramos = [];
    for (const pts of groups.values()) {
      const lat = pts.reduce((s, p) => s + p.lat, 0) / pts.length, lon = pts.reduce((s, p) => s + p.lon, 0) / pts.length;
      let ext = 0; // extensión aproximada: distancia entre los dos puntos extremos sobre el eje más largo
      const byLat = pts.slice().sort((a, b) => a.lat - b.lat), byLon = pts.slice().sort((a, b) => a.lon - b.lon);
      ext = Math.max(hav(byLat[0].lat, byLat[0].lon, byLat.at(-1).lat, byLat.at(-1).lon), hav(byLon[0].lat, byLon[0].lon, byLon.at(-1).lat, byLon.at(-1).lon));
      const pl = nearestPlace(lat, lon, places || []);
      const t = {
        lat, lon, key: tramoKey(lat, lon), n: pts.length, ext,
        grave: pts.filter(p => p.falta === 'GRAVE').length,
        moderado: pts.filter(p => p.falta === 'MODERADO').length,
        leve: pts.filter(p => p.falta === 'LEVE').length,
        vmax: Math.max(...pts.map(p => p.vel)),
        excProm: pts.reduce((s, p) => s + p.exc, 0) / pts.length,
        placas: new Set(pts.map(p => p.placa)).size, conds: new Set(pts.map(p => p.cond)).size,
        ...pl,
      };
      tramos.push(t); pts.forEach(p => { p.tramo = t; });
    }
    tramos.sort((a, b) => b.n - a.n || b.grave - a.grave);
    tramos.forEach((t, i) => { t.rank = i + 1; });
    return { tramos, aislados: rows.filter(r => !r.tramo).length };
  }

  const api = { norm, findHeader, parseFecha, parseHora, parseCoord, extractSheet, classify, dedupe, dbscan, buildTramos, nearestPlace, hav, guessOrigen, tramoKey };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.ETL = api;
})(this);
