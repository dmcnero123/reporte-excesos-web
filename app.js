/* Interfaz: navegación, carga de archivos, filtros, mapa, gráficos y exportación. */
(function () {
  'use strict';
  const $ = id => document.getElementById(id);
  const $$ = s => [...document.querySelectorAll(s)];
  const fmtN = n => Number(n).toLocaleString('es-PE');
  const pct = (a, b) => b ? Math.round(100 * a / b) + '%' : '0%';
  const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const title = s => String(s).toLowerCase().replace(/(^|\s)\S/g, m => m.toUpperCase());
  const F = ['LEVE', 'MODERADO', 'GRAVE'];
  const COLOR = { LEVE: '#7F97B2', MODERADO: '#E08A1E', GRAVE: '#C0392B' };
  const CAT_KEY = 'excesos.catalogoTramos.v1';
  const gmaps = (la, lo) => `https://www.google.com/maps?q=${la.toFixed(6)},${lo.toFixed(6)}`;

  const state = { sources: [], places: [], all: [], view: [], tramos: [], charts: {}, map: null, layers: {}, markers: new Map(), tab: 'resumen', on: { pts: true, tr: true }, evoF: '', evoOpen: null };

  // ---------- Catálogo de nombres ----------
  const loadCat = () => { try { return JSON.parse(localStorage.getItem(CAT_KEY) || '{}'); } catch (e) { return {}; } };
  const saveCat = c => { try { localStorage.setItem(CAT_KEY, JSON.stringify(c)); } catch (e) { /* sin almacenamiento */ } };
  let catalogo = loadCat();
  const nombreTramo = t => catalogo[t.key] || t.lugar;

  state.places = window.LUGARES_PE || [];

  // ---------- Navegación ----------
  const PAGES = { excesos: ['Monitoreo', 'Excesos de velocidad'], conduccion: ['Monitoreo', 'Conducción continua'], estatus: ['Monitoreo', 'Estatus de unidades'], odometro: ['Monitoreo', 'Odómetro'], fuentes: ['Administración', 'Fuentes de datos'], catalogo: ['Administración', 'Catálogo de tramos'] };
  function route() {
    const p = PAGES[location.hash.slice(1)] ? location.hash.slice(1) : 'excesos';
    $$('.view').forEach(v => { v.hidden = v.dataset.page !== p; });
    $$('.nav-i').forEach(a => a.classList.toggle('on', a.dataset.view === p));
    $('crumbSec').textContent = PAGES[p][0]; $('crumbPage').textContent = PAGES[p][1];
    closeMenu(); mapVisible(p === 'excesos' && state.tab === 'mapa');
  }
  window.addEventListener('hashchange', route);
  const closeMenu = () => { $('side').classList.remove('open'); $('scrim').classList.remove('open'); };
  $('btnMenu').addEventListener('click', () => { $('side').classList.add('open'); $('scrim').classList.add('open'); });
  $('scrim').addEventListener('click', closeMenu);

  function setTab(t) {
    state.tab = t;
    $$('.tab').forEach(b => { b.classList.toggle('on', b.dataset.tab === t); b.setAttribute('aria-selected', b.dataset.tab === t); });
    $$('.tabp').forEach(p => { p.hidden = p.dataset.panel !== t; });
    if (t === 'mapa') { renderMap(); } else { mapVisible(false); if (state.view.length) renderTab(t); }
  }
  $$('.tab').forEach(b => b.addEventListener('click', () => setTab(b.dataset.tab)));
  document.addEventListener('click', e => { const g = e.target.closest('[data-goto]'); if (g) setTab(g.dataset.goto); });

  // ---------- Carga ----------
  const drop = $('drop');
  ['dragenter', 'dragover'].forEach(e => drop.addEventListener(e, ev => { ev.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(e => drop.addEventListener(e, ev => { ev.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', ev => handleFiles([...ev.dataTransfer.files]));
  ['file', 'file2'].forEach(id => $(id).addEventListener('change', ev => { handleFiles([...ev.target.files]); ev.target.value = ''; }));

  async function handleFiles(files) {
    if (!window.XLSX) return msg('No se pudo cargar el lector de Excel. Revisa la conexión a internet y recarga la página.', true);
    const skipped = [];
    for (const f of files) {
      try {
        const wb = XLSX.read(await f.arrayBuffer(), { type: 'array', cellDates: false });
        let found = false;
        for (const sn of wb.SheetNames) {
          const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: true, defval: null, blankrows: false });
          const origen = ETL.guessOrigen(f.name, sn);
          const res = ETL.extractSheet(rows, { file: f.name, sheet: sn, origen });
          if (!res || !res.rows.length) continue;
          found = true;
          state.sources.push({ id: Math.random().toString(36).slice(2), file: f.name, sheet: sn, origen, on: true, ...res });
        }
        if (!found) skipped.push(f.name);
      } catch (e) { console.error(e); skipped.push(f.name); }
    }
    renderSources();
    msg(skipped.length ? `No se encontró una tabla de excesos en: ${skipped.join(', ')}. La hoja debe tener columnas de placa, velocidad y ubicación.` : '', !!skipped.length);
    if (skipped.length && !state.sources.length) location.hash = '#fuentes';
    rebuild();
  }
  function msg(t, err) { const m = $('msg'); m.textContent = t; m.classList.toggle('err', !!err); const b = $('banner'); b.textContent = t; b.hidden = !t; }

  function renderSources() {
    const n = state.sources.length;
    $('navSrc').hidden = !n; $('navSrc').textContent = n;
    if (!n) { $('sources').innerHTML = '<p class="muted pad">Todavía no has cargado reportes. Usa el botón “Cargar reportes”.</p>'; return; }
    $('sources').innerHTML = `<table class="tbl src"><thead><tr><th>Activa</th><th>Archivo</th><th>Hoja</th><th>Periodo</th><th class="r">Registros</th><th>Coordenadas</th><th></th></tr></thead><tbody>${state.sources.map(s => {
      const sep = s.rows.filter(r => r.fmt === 'separado').length, peg = s.rows.length - sep;
      const fe = s.rows.map(r => r.fecha).filter(Boolean).sort();
      return `<tr><td><input type="checkbox" data-on="${s.id}" ${s.on ? 'checked' : ''} aria-label="Usar hoja ${esc(s.sheet)}" style="height:auto"></td>
      <td><b>${esc(s.file)}</b></td><td>${esc(s.sheet)}</td>
      <td>${fe.length ? `${fmtDate(fe[0])} – ${fmtDate(fe.at(-1))}` : '—'}</td><td class="r">${fmtN(s.rows.length)}</td>
      <td class="small">${sep ? fmtN(sep) + ' separadas' : ''}${sep && peg ? ' · ' : ''}${peg ? fmtN(peg) + ' pegadas' : ''}${s.issues.coord ? ` · <span style="color:var(--grave)">${s.issues.coord} inválidas</span>` : ''}</td>
      <td><button class="btn" data-del="${s.id}">Quitar</button></td></tr>`;
    }).join('')}</tbody></table>`;
  }
  $('sources').addEventListener('change', e => {
    const s = state.sources.find(x => x.id === (e.target.dataset.on || e.target.dataset.origen)); if (!s) return;
    if (e.target.dataset.on) s.on = e.target.checked;
    if (e.target.dataset.origen) { s.origen = e.target.value; s.rows.forEach(r => { r.origen = s.origen; }); }
    rebuild();
  });
  $('sources').addEventListener('click', e => { const id = e.target.dataset.del; if (!id) return; state.sources = state.sources.filter(s => s.id !== id); renderSources(); rebuild(); });

  // ---------- Procesamiento ----------
  const cfg = () => ({
    limite: +$('cLim').value || 80, leve: +$('cLeve').value || 5, moderado: Math.max(+$('cMod').value || 10, (+$('cLeve').value || 5) + 1),
    eps: Math.max(+$('cEps').value || 1.5, 0.2), minPts: Math.max(+$('cMin').value || 8, 2),
  });
  const MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
  const fmtDate = s => s ? `${s.slice(8)}/${s.slice(5, 7)}/${s.slice(0, 4)}` : '';

  let lastRange = '';
  function rebuild() {
    const dd = ETL.dedupe(state.sources.filter(s => s.on).flatMap(s => s.rows));
    state.all = ETL.classify(dd.rows, cfg()); state.dupes = dd.removed;
    const has = state.all.length > 0;
    $('report').hidden = !has; $('empty').hidden = has; $('btnPlantilla').disabled = !has;
    const st = $('dataStatus');
    st.classList.toggle('ok', has);
    st.textContent = has ? `${state.sources.filter(s => s.on).length} hojas · ${fmtN(state.all.length)} eventos` : 'Sin datos cargados';
    if (!has) { $('periodoTxt').textContent = 'Carga los reportes de GPS para generar el informe.'; return; }
    const fechas = state.all.map(r => r.fecha).filter(Boolean).sort();
    const rango = fechas[0] + '|' + fechas.at(-1);
    if (rango !== lastRange) {
      lastRange = rango; const ult = fechas.at(-1).slice(0, 7);
      $('fDesde').value = fechas.find(f => f.startsWith(ult)) || fechas[0]; $('fHasta').value = fechas.at(-1);
      $('fDesde').min = $('fHasta').min = fechas[0]; $('fDesde').max = $('fHasta').max = fechas.at(-1);
    }
    fillSelect('fPlaca', [...new Set(state.all.map(r => r.placa))].sort(), v => v);
    fillSelect('fCond', [...new Set(state.all.map(r => r.cond))].sort(), title);
    applyFilters();
  }
  function fillSelect(id, vals, lab) {
    const el = $(id), cur = el.value, first = el.options[0].outerHTML;
    el.innerHTML = first + vals.map(v => `<option value="${esc(v)}">${esc(lab(v))}</option>`).join('');
    if (vals.includes(cur)) el.value = cur;
  }
  ['fDesde', 'fHasta', 'fPlaca', 'fCond', 'fFalta'].forEach(id => $(id).addEventListener('change', applyFilters));
  ['cLim', 'cLeve', 'cMod', 'cEps', 'cMin'].forEach(id => $(id).addEventListener('change', rebuild));

  function applyFilters() {
    const d = $('fDesde').value, h = $('fHasta').value, p = $('fPlaca').value, c = $('fCond').value, f = $('fFalta').value;
    state.view = state.all.filter(r => (!d || !r.fecha || r.fecha >= d) && (!h || !r.fecha || r.fecha <= h) && (!p || r.placa === p) && (!c || r.cond === c) && (!f || r.falta === f));
    const tr = ETL.buildTramos(state.view, cfg(), state.places);
    state.tramos = tr.tramos; state.aislados = tr.aislados;
    // texto del periodo
    if (d && h) {
      const [y1, m1] = d.split('-'), [y2, m2] = h.split('-');
      const mesTxt = (y1 === y2 && m1 === m2 && d.endsWith('-01')) ? `${title(MESES[+m1 - 1])} ${y1}` : `${fmtDate(d)} al ${fmtDate(h)}`;
      $('periodoTxt').textContent = `Periodo: ${mesTxt} · ${fmtN(state.view.length)} eventos de ${new Set(state.view.map(r => r.placa)).size} unidades`;
    }
    renderTramos();
    if (state.tab === 'mapa') renderMap(); else { state.mapDirty = true; renderTab(state.tab); }
  }

  function renderTab(t) {
    if (t === 'resumen') renderResumen();
    if (t === 'conductores') renderConductores();
    if (t === 'detalle') renderDetalle();
    if (t === 'evolucion') renderEvolucion();
  }

  // ---------- Gráficos ----------
  function chart(id, config) {
    if (!window.Chart) return;
    if (state.charts[id]) state.charts[id].destroy();
    Chart.defaults.font.family = "Inter, 'Segoe UI', Arial, sans-serif"; Chart.defaults.font.size = 12; Chart.defaults.color = '#6B7280';
    state.charts[id] = new Chart($(id), config);
  }
  const grid = { color: '#EEF1F5' };
  const opts = (extra = {}) => Object.assign({ responsive: true, maintainAspectRatio: false, animation: { duration: 250 }, plugins: { legend: { display: false }, tooltip: { backgroundColor: '#111827', padding: 10, cornerRadius: 6 } }, scales: { x: { grid: { display: false }, border: { color: '#E2E6EC' } }, y: { beginAtZero: true, grid, border: { display: false }, ticks: { precision: 0 } } } }, extra);
  const count = (arr, f) => { const m = new Map(); for (const r of arr) m.set(f(r), (m.get(f(r)) || 0) + 1); return m; };
  const legendBottom = { display: true, position: 'bottom', align: 'start', labels: { boxWidth: 10, boxHeight: 10, usePointStyle: true, pointStyle: 'rectRounded', padding: 14 } };

  function renderResumen() {
    const v = state.view;
    const dias = [...new Set(v.map(r => r.fecha).filter(Boolean))].sort();
    const byF = new Map(F.map(f => [f, count(v.filter(r => r.falta === f), r => r.fecha)]));
    chart('chDia', { type: 'bar', data: { labels: dias.map(d => d.slice(8) + '/' + d.slice(5, 7)), datasets: F.map(f => ({ label: title(f), data: dias.map(d => byF.get(f).get(d) || 0), backgroundColor: COLOR[f], borderRadius: 2, maxBarThickness: 22 })) },
      options: opts({ plugins: { legend: legendBottom, tooltip: { mode: 'index', backgroundColor: '#111827', padding: 10 } }, scales: { x: { stacked: true, grid: { display: false } }, y: { stacked: true, beginAtZero: true, grid, border: { display: false } } } }) });
    const ch = count(v, r => r.h), hrs = [...Array(24).keys()], mx = Math.max(1, ...hrs.map(h => ch.get(h) || 0));
    chart('chHora', { type: 'bar', data: { labels: hrs.map(h => String(h).padStart(2, '0')), datasets: [{ label: 'Eventos', data: hrs.map(h => ch.get(h) || 0), backgroundColor: hrs.map(h => (ch.get(h) || 0) >= mx * 0.75 ? '#6E9418' : '#CBD3C0'), borderRadius: 2 }] },
      options: opts({ plugins: { legend: { display: false }, tooltip: { callbacks: { title: i => `${i[0].label}:00 – ${i[0].label}:59` } } } }) });
    const cf = count(v, r => r.falta);
    const rng = { LEVE: `1 a ${cfg().leve} km/h`, MODERADO: `${cfg().leve + 1} a ${cfg().moderado} km/h`, GRAVE: `más de ${cfg().moderado} km/h` };
    $('sevList').innerHTML = F.slice().reverse().map(f => { const n = cf.get(f) || 0; return `<div class="sev-row"><b>${title(f)} <span>· ${rng[f]}</span></b><span>${fmtN(n)} · ${pct(n, v.length)}</span><div class="sev-bar"><i style="width:${v.length ? 100 * n / v.length : 0}%;background:${COLOR[f]}"></i></div></div>`; }).join('') +
      `<p class="sev-note">Hora pico: ${String(hrs.reduce((a, b) => (ch.get(b) || 0) > (ch.get(a) || 0) ? b : a, 0)).padStart(2, '0')}:00 h · ${state.dupes ? fmtN(state.dupes) + ' duplicados entre archivos eliminados' : 'sin duplicados entre archivos'}</p>`;
    const T = state.tramos.slice(0, 6);
    $('tTop5').innerHTML = T.length ? `<thead><tr><th></th><th>Tramo</th><th class="r">Eventos</th><th class="r">Graves</th><th class="r">Máx.</th></tr></thead><tbody>${T.map(t => `<tr><td><span class="rank">${t.rank}</span></td><td>${esc(nombreTramo(t))}<br><span class="muted small">${esc(t.prov)}${t.reg ? ', ' + esc(t.reg) : ''}</span></td><td class="r"><b>${t.n}</b></td><td class="r">${pct(t.grave, t.n)}</td><td class="r">${t.vmax} km/h</td></tr>`).join('')}</tbody>` : '<tbody><tr><td class="muted">Sin tramos con los filtros actuales.</td></tr></tbody>';
  }

  function stackedTop(id, key, n, lab) {
    const v = state.view;
    const top = [...count(v, key).entries()].sort((a, b) => b[1] - a[1]).slice(0, n).map(x => x[0]);
    chart(id, { type: 'bar', data: { labels: top.map(t => { const s = lab(t); return s.length > 28 ? s.slice(0, 27) + '…' : s; }), datasets: F.map(f => ({ label: title(f), data: top.map(t => v.filter(r => key(r) === t && r.falta === f).length), backgroundColor: COLOR[f], maxBarThickness: 18 })) },
      options: opts({ indexAxis: 'y', plugins: { legend: legendBottom, tooltip: { mode: 'index', backgroundColor: '#111827', padding: 10 } }, scales: { x: { stacked: true, beginAtZero: true, grid, border: { display: false } }, y: { stacked: true, grid: { display: false } } } }) });
  }

  function renderConductores() {
    stackedTop('chCond', r => r.cond, 10, title);
    stackedTop('chPlaca', r => r.placa, 10, s => s);
    const v = state.view, by = new Map();
    for (const r of v) { const o = by.get(r.cond) || { n: 0, L: 0, M: 0, G: 0, vmax: 0, placas: new Map(), tramos: new Map() }; o.n++; o[r.falta[0]]++; o.vmax = Math.max(o.vmax, r.vel); o.placas.set(r.placa, (o.placas.get(r.placa) || 0) + 1); if (r.tramo) o.tramos.set(r.tramo, (o.tramos.get(r.tramo) || 0) + 1); by.set(r.cond, o); }
    const top = (m) => [...m.entries()].sort((a, b) => b[1] - a[1])[0];
    const rows = [...by.entries()].sort((a, b) => b[1].n - a[1].n);
    $('tCond').closest('.table-wrap').classList.add('tall');
    $('tCond').innerHTML = `<thead><tr><th>#</th><th>Conductor</th><th>Unidad principal</th><th class="r">Leve</th><th class="r">Moderado</th><th class="r">Grave</th><th class="r">Total</th><th class="r">Vel. máx.</th><th>Tramo más frecuente</th></tr></thead><tbody>${rows.map(([c, o], i) => {
      const tt = top(o.tramos);
      return `<tr><td><span class="rank">${i + 1}</span></td><td>${esc(title(c))}</td><td>${esc(top(o.placas)[0])}</td><td class="r">${o.L}</td><td class="r">${o.M}</td><td class="r" style="color:var(--grave);font-weight:600">${o.G}</td><td class="r"><b>${o.n}</b></td><td class="r">${o.vmax} km/h</td><td>${tt ? `${tt[0].rank} · ${esc(nombreTramo(tt[0]))}` : '<span class="muted">—</span>'}</td></tr>`;
    }).join('')}</tbody>`;
  }

  // ---------- Mapa ----------
  const sevColor = g => g >= 0.8 ? '#C0392B' : g >= 0.6 ? '#E0702A' : g >= 0.4 ? '#E0A21E' : '#5B7FA8';
  function renderMap() {
    if (!window.L) { $('map').innerHTML = '<p class="muted pad">No se pudo cargar el mapa. Revisa la conexión.</p>'; return; }
    if (!state.map) {
      state.map = L.map('map', { preferCanvas: true, zoomControl: true, attributionControl: true });
      L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/Canvas/World_Light_Gray_Base/MapServer/tile/{z}/{y}/{x}', { maxZoom: 16, attribution: 'Mapa base © Esri' }).addTo(state.map);
      state.map.createPane('labels'); state.map.getPane('labels').style.zIndex = 450; state.map.getPane('labels').style.pointerEvents = 'none';
      L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/Canvas/World_Light_Gray_Reference/MapServer/tile/{z}/{y}/{x}', { maxZoom: 16, pane: 'labels' }).addTo(state.map);
      state.map.attributionControl.setPrefix(false);
      state.layers = { pts: L.layerGroup(), tr: L.layerGroup() };
      $$('.map-wrap .seg').forEach(b => b.addEventListener('click', () => { const k = b.dataset.layer; state.on[k] = !state.on[k]; b.classList.toggle('on', state.on[k]); syncLayers(); }));
    }
    const { pts, tr } = state.layers; pts.clearLayers(); tr.clearLayers(); state.markers.clear();
    const v = state.view;
    const rend = L.canvas({ padding: 0.5 });
    for (const r of v) {
      L.circleMarker([r.lat, r.lon], { renderer: rend, radius: 3.5, weight: 1, color: '#fff', fillOpacity: 0.9, fillColor: COLOR[r.falta] })
        .bindPopup(`<div class="pop-h">${esc(r.placa)} · ${title(r.falta)}</div>${esc(title(r.cond))}<div class="pop-g"><span>Fecha</span><b>${fmtDate(r.fecha)} ${r.hora}</b><span>Velocidad</span><b>${r.vel} km/h</b><span>Límite</span><b>${r.limUsado} km/h</b></div><a href="${gmaps(r.lat, r.lon)}" target="_blank" rel="noopener">Ver en Google Maps</a>`).addTo(pts);
    }
    state.tramos.forEach(t => {
      const s = t.rank <= 10 ? 26 : 18;
      const m = L.marker([t.lat, t.lon], { icon: L.divIcon({ className: '', html: `<div class="tpin" style="width:${s}px;height:${s}px;background:${sevColor(t.grave / t.n)};font-size:${t.rank <= 10 ? 11 : 9}px">${t.rank <= 99 ? t.rank : ''}</div>`, iconSize: [s, s], iconAnchor: [s / 2, s / 2] }), zIndexOffset: 1000 - t.rank, riseOnHover: true })
        .bindPopup(popupTramo(t)).on('click', () => selectTramo(t, false)).addTo(tr);
      state.markers.set(t, m);
    });
    $('mapLegend').innerHTML = `<span><i style="background:#C0392B"></i>Tramo con ≥80% graves</span><span><i style="background:#E0702A"></i>60–79% graves</span><span><i style="background:#5B7FA8"></i>menos de 60%</span>`;
    syncLayers(); renderTList();
    setTimeout(() => { state.map.invalidateSize(); if (v.length) state.map.fitBounds(L.latLngBounds(v.map(r => [r.lat, r.lon])).pad(0.05)); }, 40);
  }
  function mapVisible(on) { if (!state.map) return; if (on) setTimeout(() => { state.map.invalidateSize(); syncLayers(); }, 30); }
  function syncLayers() { for (const k in state.layers) { if (state.on[k]) state.layers[k].addTo(state.map); else state.layers[k].remove(); } }

  function popupTramo(t) {
    return `<div class="pop-h">Tramo ${t.rank} · ${esc(nombreTramo(t))}</div><span class="muted">${esc(t.prov)}${t.reg ? ', ' + esc(t.reg) : ''}</span>
    <div class="pop-g"><span>Eventos</span><b>${t.n}</b><span>Graves</span><b>${t.grave} (${pct(t.grave, t.n)})</b><span>Vel. máxima</span><b>${t.vmax} km/h</b><span>Exceso prom.</span><b>+${t.excProm.toFixed(1)} km/h</b><span>Extensión</span><b>${t.ext.toFixed(1)} km</b><span>Unidades</span><b>${t.placas}</b></div>
    <a href="${gmaps(t.lat, t.lon)}" target="_blank" rel="noopener">Ver en Google Maps</a>`;
  }
  function renderTList() {
    const T = state.tramos;
    $('tCount').textContent = `${T.length} tramos · ${fmtN(state.aislados)} eventos aislados`;
    $('tList').innerHTML = T.map((t, i) => `<li data-i="${i}"><span class="rank">${t.rank}</span><span class="tn" title="${esc(nombreTramo(t))}">${esc(nombreTramo(t))}</span><span class="tm">${esc(t.prov)} · ${pct(t.grave, t.n)} graves · máx ${t.vmax}</span><span class="tc">${t.n}<small>eventos</small></span></li>`).join('') || '<li class="muted">Sin tramos con los filtros actuales.</li>';
  }
  $('tList').addEventListener('click', e => { const li = e.target.closest('li[data-i]'); if (li) selectTramo(state.tramos[+li.dataset.i], true); });
  function selectTramo(t, fly) {
    $$('#tList li').forEach(li => li.classList.toggle('on', state.tramos[+li.dataset.i] === t));
    const li = $('tList').querySelector('li.on'); if (li && !fly) li.scrollIntoView({ block: 'nearest' });
    if (fly) { state.map.flyTo([t.lat, t.lon], 12, { duration: 0.6 }); const m = state.markers.get(t); if (m) setTimeout(() => m.openPopup(), 650); }
  }


  // ---------- Evolución mensual ----------
  const CH_KEY = 'excesos.charlas.v1';
  let charlas = (() => { try { return JSON.parse(localStorage.getItem(CH_KEY) || '{}'); } catch (e) { return {}; } })();
  const saveCh = () => { try { localStorage.setItem(CH_KEY, JSON.stringify(charlas)); } catch (e) { /* sin almacenamiento */ } };
  const mesTxt = m => `${title(MESES[+m.slice(5) - 1])} ${m.slice(0, 4)}`;
  const mesCorto = m => `${title(MESES[+m.slice(5) - 1]).slice(0, 3)} ${m.slice(2, 4)}`;
  // Estado de un cambio: menos de ±10% se considera sin cambio
  function estado(a, b) {
    if (!a && !b) return { k: 'eq', t: 'Sin eventos' };
    if (!a) return { k: 'new', t: 'Nuevo' };
    if (!b) return { k: 'down', t: 'Sin faltas' };
    const p = (b - a) / a;
    return p <= -0.1 ? { k: 'down', t: 'Mejoró' } : p >= 0.1 ? { k: 'up', t: 'Empeoró' } : { k: 'eq', t: 'Se mantiene' };
  }
  function delta(a, b) {
    const d = b - a; if (!d) return '<span class="muted">0</span>';
    const p = a ? ` (${d > 0 ? '+' : ''}${Math.round(100 * d / a)}%)` : '';
    return `<span class="${d > 0 ? 'up' : 'down'}">${d > 0 ? '▲ +' : '▼ '}${d}${p}</span>`;
  }
  const pillE = e => `<span class="pill e-${e.k}">${e.t}</span>`;

  function baseRows() {
    const p = $('fPlaca').value, c = $('fCond').value, f = $('fFalta').value;
    return state.all.filter(r => r.fecha && (!p || r.placa === p) && (!c || r.cond === c) && (!f || r.falta === f));
  }

  function renderEvolucion() {
    const base = baseRows();
    const meses = [...new Set(base.map(r => r.fecha.slice(0, 7)))].sort();
    const ok = meses.length >= 2;
    $('evoEmpty').hidden = ok; $('evoBody').hidden = !ok; $('evoSum').hidden = !ok;
    for (const [id, def] of [['mA', meses.at(-2)], ['mB', meses.at(-1)]]) {
      const el = $(id), cur = el.value;
      el.innerHTML = meses.map(m => `<option value="${m}">${mesTxt(m)}</option>`).join('');
      el.value = meses.includes(cur) ? cur : (def || '');
    }
    if (!ok) return;
    const A = $('mA').value, B = $('mB').value;
    const rows = base.filter(r => r.fecha.startsWith(A) || r.fecha.startsWith(B)).map(r => ({ ...r, mes: r.fecha.startsWith(A) ? 'a' : 'b' }));
    const T = ETL.buildTramos(rows, cfg(), state.places).tramos; // tramos comunes a los dos meses
    const ra = rows.filter(r => r.mes === 'a'), rb = rows.filter(r => r.mes === 'b');

    // por conductor
    const by = new Map();
    for (const r of rows) {
      const o = by.get(r.cond) || { a: 0, b: 0, ga: 0, gb: 0, placas: new Map(), tr: new Map() };
      o[r.mes]++; if (r.falta === 'GRAVE') o['g' + r.mes]++;
      o.placas.set(r.placa, (o.placas.get(r.placa) || 0) + 1);
      if (r.tramo) { const t = o.tr.get(r.tramo) || { a: 0, b: 0 }; t[r.mes]++; o.tr.set(r.tramo, t); }
      by.set(r.cond, o);
    }
    const conds = [...by.entries()].map(([c, o]) => ({ c, ...o, e: estado(o.a, o.b), d: o.b - o.a }));
    const n = k => conds.filter(x => x.e.k === k).length;
    $('evoSum').innerHTML = [
      ['Eventos', `${fmtN(ra.length)} → ${fmtN(rb.length)}`, delta(ra.length, rb.length)],
      ['Graves', `${fmtN(ra.filter(r => r.falta === 'GRAVE').length)} → ${fmtN(rb.filter(r => r.falta === 'GRAVE').length)}`, delta(ra.filter(r => r.falta === 'GRAVE').length, rb.filter(r => r.falta === 'GRAVE').length)],
      ['Conductores que mejoraron', n('down'), ''], ['Conductores que empeoraron', n('up'), ''], ['Con eventos por primera vez', n('new'), ''],
    ].map(([l, v, d]) => `<div class="es"><p>${l}</p><b>${v}</b>${d ? `<span>${d}</span>` : ''}</div>`).join('');

    const list = conds.filter(x => !state.evoF || x.e.k === state.evoF).sort((x, y) => y.d - x.d || y.b - x.b);
    $('tEvoCond').innerHTML = `<thead><tr><th>Conductor</th><th>Unidad</th><th class="r">${mesCorto(A)}</th><th class="r">${mesCorto(B)}</th><th class="r">Variación</th><th class="r">Graves</th><th>Estado</th><th>Charla</th></tr></thead><tbody>${list.map(x => {
      const open = state.evoOpen === x.c;
      const main = `<tr class="evo-row${open ? ' open' : ''}" data-c="${esc(x.c)}"><td><span class="caret">${open ? '▾' : '▸'}</span>${esc(title(x.c))}</td><td>${esc([...x.placas.entries()].sort((p, q) => q[1] - p[1])[0][0])}</td>
        <td class="r">${x.a}</td><td class="r"><b>${x.b}</b></td><td class="r">${delta(x.a, x.b)}</td><td class="r">${x.ga} → ${x.gb}</td><td>${pillE(x.e)}</td><td>${charlas[x.c] ? fmtDate(charlas[x.c]) : '<span class="muted">—</span>'}</td></tr>`;
      return main + (open ? `<tr class="evo-det"><td colspan="8">${detalleConductor(x, base, A, B)}</td></tr>` : '');
    }).join('')}</tbody>`;

    // por tramo
    const tt = T.map(t => { const a = ra.filter(r => r.tramo === t).length, b = rb.filter(r => r.tramo === t).length; return { t, a, b, e: estado(a, b), cb: new Set(rb.filter(r => r.tramo === t).map(r => r.cond)).size }; })
      .sort((x, y) => (y.b - y.a) - (x.b - x.a) || y.b - x.b);
    $('tEvoTramo').innerHTML = `<thead><tr><th>Tramo</th><th>Provincia · región</th><th class="r">${mesCorto(A)}</th><th class="r">${mesCorto(B)}</th><th class="r">Variación</th><th>Estado</th><th class="r">Conductores (${mesCorto(B)})</th></tr></thead><tbody>${tt.map(x => `
      <tr><td><a href="${gmaps(x.t.lat, x.t.lon)}" target="_blank" rel="noopener">${esc(nombreTramo(x.t))}</a></td><td>${esc(x.t.prov)}${x.t.reg ? ' · ' + esc(x.t.reg) : ''}</td><td class="r">${x.a}</td><td class="r"><b>${x.b}</b></td><td class="r">${delta(x.a, x.b)}</td><td>${pillE(x.e)}</td><td class="r">${x.cb}</td></tr>`).join('')}</tbody>`;
  }

  function detalleConductor(x, base, A, B) {
    const trs = [...x.tr.entries()].sort((p, q) => Math.max(q[1].a, q[1].b) - Math.max(p[1].a, p[1].b)).slice(0, 8);
    const tabla = trs.length ? `<table class="tbl mini"><thead><tr><th>Tramo</th><th class="r">${mesCorto(A)}</th><th class="r">${mesCorto(B)}</th><th class="r">Variación</th></tr></thead><tbody>${trs.map(([t, v]) => `<tr><td>${esc(nombreTramo(t))}</td><td class="r">${v.a}</td><td class="r"><b>${v.b}</b></td><td class="r">${delta(v.a, v.b)}</td></tr>`).join('')}</tbody></table>` : '<p class="muted small">Sin eventos dentro de tramos.</p>';
    const f = charlas[x.c];
    let efecto = '<p class="muted small">Registra la fecha para comparar los 30 días antes y después de la charla.</p>';
    if (f) {
      const t0 = Date.parse(f), dia = 864e5;
      const mine = base.filter(r => r.cond === x.c), t = r => Date.parse(r.fecha);
      const antes = mine.filter(r => t(r) >= t0 - 30 * dia && t(r) < t0).length;
      const desp = mine.filter(r => t(r) > t0 && t(r) <= t0 + 30 * dia).length;
      const fin = Math.max(...state.all.map(r => Date.parse(r.fecha) || 0));
      const dias = Math.min(30, Math.floor((fin - t0) / dia));
      efecto = `<div class="charla-res"><div><p>30 días antes</p><b>${antes}</b></div><div><p>${dias < 30 ? `Después (${Math.max(dias, 0)} días con datos)` : '30 días después'}</p><b>${desp}</b></div><div><p>Resultado</p>${pillE(estado(antes, desp))}</div></div>${dias < 30 ? '<p class="muted small">Aún no hay 30 días de datos después de la charla; el resultado puede cambiar al cargar el siguiente mes.</p>' : ''}`;
    }
    return `<div class="evo-grid"><div><h4>Tramos donde cometió faltas</h4>${tabla}</div>
      <div><h4>Charla de seguridad</h4><label class="ch-l">Fecha de la charla <input type="date" class="ch-in" data-c="${esc(x.c)}" value="${f || ''}"></label>${f ? `<button class="link ch-del" data-c="${esc(x.c)}">Quitar fecha</button>` : ''}${efecto}</div></div>`;
  }

  ['mA', 'mB'].forEach(id => $(id).addEventListener('change', renderEvolucion));
  $('evoFilt').addEventListener('click', e => {
    const b = e.target.closest('.seg'); if (!b) return;
    state.evoF = b.dataset.f; $$('#evoFilt .seg').forEach(s => s.classList.toggle('on', s === b)); renderEvolucion();
  });
  $('tEvoCond').addEventListener('click', e => {
    if (e.target.closest('.evo-det')) {
      const del = e.target.closest('.ch-del'); if (del) { delete charlas[del.dataset.c]; saveCh(); renderEvolucion(); }
      return;
    }
    const tr = e.target.closest('.evo-row'); if (!tr) return;
    state.evoOpen = state.evoOpen === tr.dataset.c ? null : tr.dataset.c; renderEvolucion();
  });
  $('tEvoCond').addEventListener('change', e => {
    const i = e.target.closest('.ch-in'); if (!i) return;
    if (i.value) charlas[i.dataset.c] = i.value; else delete charlas[i.dataset.c];
    saveCh(); renderEvolucion();
  });

  // ---------- Catálogo ----------
  function renderTramos() {
    const T = state.tramos;
    $('tTramos').innerHTML = T.length ? `<thead><tr><th>#</th><th>Nombre del tramo</th><th>Localidad más cercana</th><th>Provincia · región</th><th>Coordenada central</th><th class="r">Eventos</th><th class="r">Graves</th><th class="r">Vel. máx.</th><th class="r">Extensión</th></tr></thead><tbody>${T.map(t => `
      <tr><td><span class="rank">${t.rank}</span></td>
      <td><input class="tname" data-key="${t.key}" value="${esc(catalogo[t.key] || '')}" placeholder="${esc(t.lugar)}" aria-label="Nombre del tramo ${t.rank}"></td>
      <td>${esc(t.lugar)} <span class="muted small">(${t.dist.toFixed(0)} km)</span></td><td>${esc(t.prov)}${t.reg ? ' · ' + esc(t.reg) : ''}</td>
      <td><a href="${gmaps(t.lat, t.lon)}" target="_blank" rel="noopener">${t.lat.toFixed(5)}, ${t.lon.toFixed(5)}</a></td>
      <td class="r"><b>${t.n}</b></td><td class="r"><span class="meter"><i style="width:${Math.round(100 * t.grave / t.n)}%"></i></span>${pct(t.grave, t.n)}</td>
      <td class="r">${t.vmax} km/h</td><td class="r">${t.ext.toFixed(1)} km</td></tr>`).join('')}</tbody>` : '<tbody><tr><td class="muted pad">Carga reportes para ver los tramos.</td></tr></tbody>';
  }
  $('tTramos').addEventListener('change', e => {
    if (!e.target.classList.contains('tname')) return;
    const k = e.target.dataset.key, v = e.target.value.trim();
    if (v) catalogo[k] = v; else delete catalogo[k];
    saveCat(catalogo); state.mapDirty = true;
  });
  $('btnCatExport').addEventListener('click', () => download(new Blob([JSON.stringify(catalogo, null, 2)], { type: 'application/json' }), 'catalogo-de-tramos.json'));
  $('catImport').addEventListener('change', async e => {
    const f = e.target.files[0]; if (!f) return;
    try { catalogo = Object.assign(catalogo, JSON.parse(await f.text())); saveCat(catalogo); applyFilters(); } catch (err) { alert('El archivo del catálogo no es válido.'); }
    e.target.value = '';
  });

  // ---------- Detalle ----------
  const LIMIT = 400;
  $('q').addEventListener('input', renderDetalle);
  function renderDetalle() {
    const q = ETL.norm($('q').value);
    let rows = state.view.slice().sort((a, b) => (b.fecha || '').localeCompare(a.fecha || '') || b.hora.localeCompare(a.hora));
    if (q) rows = rows.filter(r => ETL.norm(`${r.placa} ${r.cond} ${r.tramo ? nombreTramo(r.tramo) : ''}`).includes(q));
    $('tDetalle').innerHTML = `<thead><tr><th>Fecha</th><th>Hora</th><th>Unidad</th><th>Conductor</th><th>Tramo</th><th>Ubicación</th><th class="r">Velocidad</th><th class="r">Límite</th><th class="r">Exceso</th><th>Falta</th><th>Archivo</th></tr></thead><tbody>${rows.slice(0, LIMIT).map(r => `
      <tr><td>${fmtDate(r.fecha)}</td><td>${r.hora}</td><td>${esc(r.placa)}</td><td>${esc(title(r.cond))}</td><td>${r.tramo ? `${r.tramo.rank} · ${esc(nombreTramo(r.tramo))}` : '<span class="muted">Aislado</span>'}</td>
      <td><a href="${gmaps(r.lat, r.lon)}" target="_blank" rel="noopener">${r.lat.toFixed(5)}, ${r.lon.toFixed(5)}</a></td>
      <td class="r">${r.vel} km/h</td><td class="r">${r.limUsado}</td><td class="r">+${r.exc}</td><td><span class="pill ${r.falta.toLowerCase()}">${title(r.falta)}</span></td><td class="muted small">${esc(r.archivo)}</td></tr>`).join('')}</tbody>`;
    $('detNota').textContent = rows.length > LIMIT ? `Mostrando ${LIMIT} de ${fmtN(rows.length)} eventos. El listado completo se incluye en el Excel exportado.` : `${fmtN(rows.length)} eventos.`;
  }

  // ---------- Exportación ----------
  function download(blob, name) { const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500); }
  $('btnPlantilla').addEventListener('click', () => {
    const v = state.view.slice().sort((a, b) => (b.fecha || '').localeCompare(a.fecha || '') || a.placa.localeCompare(b.placa) || a.hora.localeCompare(b.hora));
    const serial = s => { if (!s) return ''; const [y, m, d] = s.split('-').map(Number); return { t: 'n', v: (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 864e5, z: 'dd/mm/yyyy' }; };
    const hoy = new Date(), f2 = n => String(n).padStart(2, '0');
    const aoa = [
      ['', '', 'REPORTE DIARIO DE EXCESO DE VELOCIDAD', '', '', '', '', `Código: JIR-FO-OP-21\nVersión: 01\nFecha: ${f2(hoy.getDate())}/${f2(hoy.getMonth() + 1)}/${hoy.getFullYear()}`, '', ''],
      ['FECHA', 'PLACA', 'CONDUCTOR', 'HORA', 'UBICACIÓN', 'VELOCIDAD REALIZADA', 'LIMITE DE VELOCIDAD', 'EXCESO DE VELOCIDAD', 'FALTA', 'TRAMO'],
      ...v.map(r => [serial(r.fecha), r.placa, r.cond, r.hora, `${r.lat.toFixed(6)}, ${r.lon.toFixed(6)}`, r.vel, r.limUsado, r.exc, r.falta, r.tramo ? `${r.tramo.rank} · ${nombreTramo(r.tramo)}` : '']),
    ];
    const ws1 = XLSX.utils.aoa_to_sheet(aoa);
    ws1['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }, { s: { r: 0, c: 2 }, e: { r: 0, c: 6 } }, { s: { r: 0, c: 7 }, e: { r: 0, c: 9 } }];
    ws1['!cols'] = [12, 11, 36, 10, 26, 12, 12, 12, 11, 32].map(w => ({ wch: w }));
    ws1['!autofilter'] = { ref: `A2:J${aoa.length}` };
    const conds = [...count(v, r => r.cond).entries()].sort((a, b) => b[1] - a[1]);
    const est = [['REPORTE DE EXCESOS DE VELOCIDAD · ESTADÍSTICAS'], [], ['CONDUCTORES', ...F, 'TOTAL'],
      ...conds.map(([c, n]) => [c, ...F.map(f => v.filter(r => r.cond === c && r.falta === f).length), n]),
      ['TOTAL', ...F.map(f => v.filter(r => r.falta === f).length), v.length], [], ['FALTA', 'PORCENTAJE', 'CANTIDAD'],
      ...F.map(f => { const n = v.filter(r => r.falta === f).length; return [f, v.length ? n / v.length : 0, n]; }), ['TOTAL', 1, v.length]];
    const ws2 = XLSX.utils.aoa_to_sheet(est); ws2['!cols'] = [{ wch: 38 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 10 }];
    const tram = [['#', 'NOMBRE DEL TRAMO', 'LOCALIDAD MÁS CERCANA', 'PROVINCIA', 'REGIÓN', 'LATITUD', 'LONGITUD', 'GOOGLE MAPS', 'EVENTOS', 'GRAVES', 'MODERADOS', 'LEVES', 'VEL. MÁX.', 'EXCESO PROM.', 'EXTENSIÓN (KM)', 'UNIDADES', 'CONDUCTORES'],
      ...state.tramos.map(t => [t.rank, nombreTramo(t), t.lugar, t.prov, t.reg, +t.lat.toFixed(6), +t.lon.toFixed(6), gmaps(t.lat, t.lon), t.n, t.grave, t.moderado, t.leve, t.vmax, +t.excProm.toFixed(1), +t.ext.toFixed(1), t.placas, t.conds])];
    const ws3 = XLSX.utils.aoa_to_sheet(tram); ws3['!cols'] = [4, 30, 22, 16, 14, 11, 11, 46, 9, 9, 11, 8, 10, 12, 14, 10, 12].map(w => ({ wch: w }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws1, 'Plantilla exceso');
    XLSX.utils.book_append_sheet(wb, ws2, 'Estadisticas');
    XLSX.utils.book_append_sheet(wb, ws3, 'Tramos');
    const d = $('fDesde').value || 'reporte', h = $('fHasta').value || '';
    XLSX.writeFile(wb, `REPORTE_EXCESO_DE_VELOCIDAD_${d}${h && h !== d ? '_a_' + h : ''}.xlsx`, { compression: true });
  });

  const faltan = [!window.XLSX && 'lector de Excel', !window.L && 'mapa', !window.Chart && 'gráficos'].filter(Boolean);
  if (faltan.length) msg(`No se pudieron cargar: ${faltan.join(', ')}. La página necesita internet; si usas Brave, desactiva los escudos (ícono del león) para esta página y recarga.`, true);
  route(); renderSources(); renderTramos();
})();