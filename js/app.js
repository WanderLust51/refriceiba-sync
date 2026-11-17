/* Interfaz de la página. La lógica está en sync.js. */
(function () {
  'use strict';

  const MAX_CAMBIOS_SIN_CONFIRMAR = 50;   // más cambios que esto pide confirmación explícita
  const MAX_MINUTOS_PLANILLA = 60;        // planilla de Cashea más vieja que esto: advertencia
  const DIAS_SIN_COMPLETO = 7;            // sugerir un inventario completo cada tanto
  const STORE_KEY = 'refriceiba-sync/cierres';
  const DOCS_KEY = 'refriceiba-sync/documentos';
  const MODE_KEY = 'refriceiba-sync/modo';
  const MAX_CIERRES = 14;                 // cierres guardados (uno por día)
  const MAX_DOCS = 400;                   // documentos aplicados que se recuerdan

  if (window.pdfjsLib) window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'vendor/pdf.worker.min.js';

  const $ = (id) => document.getElementById(id);
  const state = {
    mode: 'completo',
    valery: null, valeryFile: null,          // modo completo
    docs: [],                                // modo movimientos: [{ id, file, tipo, parsed, error, aplicado }]
    movRes: null, baseMov: null,
    cashea: null, casheaFile: null,
    cmp: null, filter: 'cambios', base: null, ventas: null, mov: null,
  };

  /* ---------- utilidades ---------- */
  const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const fmt = (n) => (n === null || n === undefined || Number.isNaN(n) ? '—' : Number(n).toLocaleString('es-VE', { maximumFractionDigits: 2 }));
  const pl = (n, uno, varios) => `${fmt(n)} ${n === 1 ? uno : varios}`;
  const fmtDate = (t) => new Date(t).toLocaleString('es-VE', { weekday: 'short', day: 'numeric', month: 'numeric', hour: '2-digit', minute: '2-digit' });
  const minutesAgo = (t) => Math.round((Date.now() - t) / 60000);
  const ago = (t) => {
    const m = minutesAgo(t);
    if (m < 1) return 'hace menos de un minuto';
    if (m < 60) return `hace ${m} min`;
    const h = Math.floor(m / 60);
    if (h < 24) return `hace ${h} h ${m % 60} min`;
    return `el ${new Date(t).toLocaleDateString('es-VE')}`;
  };
  const stamp = () => {
    const d = new Date(), p = (n) => String(n).padStart(2, '0');
    return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}`;
  };
  const dayKey = (t) => new Date(t).toDateString();
  const store = {
    get(k, def) { try { const v = JSON.parse(localStorage.getItem(k)); return v ?? def; } catch (e) { return def; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch (e) { return false; } },
  };
  async function sha(bytes) {
    try {
      const h = await crypto.subtle.digest('SHA-256', bytes);
      return [...new Uint8Array(h)].slice(0, 12).map((b) => b.toString(16).padStart(2, '0')).join('');
    } catch (e) {   // sin crypto.subtle (página abierta como archivo local): hash simple
      let h = 2166136261; for (const b of new Uint8Array(bytes)) { h ^= b; h = Math.imul(h, 16777619); } return (h >>> 0).toString(16);
    }
  }
  // Fecha del reporte de Valery: el .xls no trae fecha adentro; se usa la hora en que se exportó el archivo.
  const valeryTime = () => (state.valeryFile ? state.valeryFile.lastModified : Date.now());

  /* ---------- cierres guardados en este navegador ----------
     [{ t, guardado, modo, archivo, valery: { código: existencia }, cashea: { SKU: stock subido }, casheaT, docs: [id] }] */
  function loadCierres() {
    const v = store.get(STORE_KEY, []);
    return Array.isArray(v) ? v.map((c) => ({ ...c, valery: c.valery || c.items || {}, modo: c.modo || 'completo' })) : [];
  }
  function saveCierre() {
    if (!state.cmp) return false;
    const movimientos = state.mode === 'movimientos';
    if (!movimientos && !state.valery) return false;
    const t = movimientos ? Date.now() : valeryTime();
    const prevList = loadCierres();
    const ultimoCompleto = movimientos
      ? (state.baseMov && (state.baseMov.ultimoCompleto || (state.baseMov.modo === 'completo' ? state.baseMov.t : null)))
      : t;
    const entry = {
      t, guardado: Date.now(), modo: state.mode, ultimoCompleto,
      archivo: movimientos ? activeDocs().map((d) => d.file.name).join(', ') : (state.valeryFile ? state.valeryFile.name : ''),
      valery: RefriSync.snapshot(movimientos ? state.movRes.valery : state.valery),
      cashea: RefriSync.casheaSnapshot(state.cmp),
      casheaT: state.cashea && state.cashea.generatedAt ? state.cashea.generatedAt.getTime() : null,
      docs: movimientos ? activeDocs().map((d) => d.id) : [],
    };
    const list = prevList.filter((c) => dayKey(c.t) !== dayKey(t));   // uno por día: el último gana
    list.push(entry);
    list.sort((a, b) => a.t - b.t);
    const ok = store.set(STORE_KEY, list.slice(-MAX_CIERRES).map(({ items, ...c }) => c));
    if (movimientos) {
      const docs = store.get(DOCS_KEY, []);
      for (const d of activeDocs()) docs.push({ id: d.id, tipo: d.tipo, nombre: d.file.name, etiqueta: docLabel(d), aplicado: Date.now() });
      store.set(DOCS_KEY, docs.slice(-MAX_DOCS));
    }
    return ok;
  }
  // Modo completo: base de ventas por canal = último cierre de un día ANTERIOR al del reporte de Valery.
  function pickBase() {
    const hoy = dayKey(valeryTime());
    const prev = loadCierres().filter((c) => c.t < valeryTime() && dayKey(c.t) !== hoy);
    if (!prev.length) return null;
    const c = prev[prev.length - 1];
    return { label: `cierre del ${fmtDate(c.t)}`, valery: c.valery, cashea: c.cashea || null, t: c.t, manual: false };
  }
  // Modo movimientos: se parte del último cierre guardado, sea del día que sea.
  function lastCierre() {
    const l = loadCierres().sort((a, b) => (a.guardado || a.t) - (b.guardado || b.t));
    return l.length ? l[l.length - 1] : null;
  }
  function lastCompleto() {
    const l = loadCierres();
    const c = l.filter((x) => x.modo === 'completo').map((x) => x.t);
    const u = l.map((x) => x.ultimoCompleto).filter(Boolean);
    const all = c.concat(u);
    return all.length ? Math.max(...all) : null;
  }
  const appliedDocs = () => new Map(store.get(DOCS_KEY, []).map((d) => [d.id, d]));

  /* ---------- modo ---------- */
  function setMode(m) {
    state.mode = m;
    store.set(MODE_KEY, m);
    document.querySelectorAll('.modes button').forEach((b) => b.setAttribute('aria-checked', String(b.dataset.mode === m)));
    $('dropValery').hidden = m !== 'completo';
    $('dropDocs').hidden = m !== 'movimientos';
    $('docList').hidden = m !== 'movimientos' || !state.docs.length;
    state.base = null;
    if (state.filter === 'mov') state.filter = 'cambios';
    selectTabUI(state.filter);
    renderBaseInfo();
    run();
  }

  function renderBaseInfo() {
    if (state.mode !== 'movimientos') { $('baseInfo').innerHTML = ''; return; }
    const b = lastCierre();
    state.baseMov = b;
    if (!b) {
      $('baseInfo').innerHTML = alertBox('bad', 'Todavía no hay un cierre guardado en esta computadora.',
        'Los movimientos se aplican sobre la última existencia guardada. Haz primero un cierre con <b>Inventario completo</b> (y descarga la planilla); desde el siguiente ya puedes usar movimientos.');
      return;
    }
    const lc = lastCompleto();
    const viejo = lc && (Date.now() - lc) / 86400000 > DIAS_SIN_COMPLETO;
    $('baseInfo').innerHTML = alertBox('info', `Se parte del cierre del ${fmtDate(b.guardado || b.t)} (${b.modo === 'completo' ? 'inventario completo' : 'movimientos'}).`,
      'Carga <b>todas</b> las ventas y facturas de compra desde ese cierre, ni más ni menos. En Valery, pon en la Relación de ventas el rango de fechas desde ese cierre.')
      + (viejo ? alertBox('warn', `El último inventario completo fue ${ago(lc)}.`,
        'Los movimientos acumulan cualquier diferencia (ventas sin facturar, ajustes en Valery). Conviene hacer un cierre con inventario completo al menos una vez por semana.') : '');
  }

  /* ---------- carga de archivos ---------- */
  function setDrop(el, kind, file, msg, extra) {
    el.classList.remove('ok', 'err');
    if (kind) el.classList.add(kind);
    const f = el.querySelector('.file');
    if (f) f.textContent = file ? `${file.name}${extra ? ' · ' + extra : ''}` : '';
    el.querySelector('.state').textContent = msg || '';
  }

  function wireDrop(dropId, inputId, handler, multiple) {
    const drop = $(dropId), input = $(inputId);
    input.addEventListener('change', () => { const fs = [...input.files]; input.value = ''; if (fs.length) handler(multiple ? fs : fs[0]); });
    ['dragenter', 'dragover'].forEach((e) => drop.addEventListener(e, (ev) => { ev.preventDefault(); drop.classList.add('over'); }));
    ['dragleave', 'drop'].forEach((e) => drop.addEventListener(e, () => drop.classList.remove('over')));
    drop.addEventListener('drop', (ev) => { ev.preventDefault(); const fs = [...ev.dataTransfer.files]; if (fs.length) handler(multiple ? fs : fs[0]); });
  }

  async function onValery(file) {
    const drop = $('dropValery');
    setDrop(drop, null, file, 'Leyendo…');
    try {
      const bytes = new Uint8Array(await file.arrayBuffer());
      if (/\.pdf$/i.test(file.name)) throw new Error('Esto es un PDF: las facturas de compra van en "Movimientos del día".');
      if (RefriSync.detectXls(XLSX, bytes) === 'ventas') throw new Error('Esta es la Relación de ventas: va en "Movimientos del día" (arriba).');
      state.valery = RefriSync.parseValery(XLSX, bytes);
      state.valeryFile = file;
      if (state.base && !state.base.manual) state.base = null;
      setDrop(drop, 'ok', file, `✓ ${state.valery.size} productos`, `exportado ${fmtDate(file.lastModified)}`);
    } catch (e) {
      state.valery = null;
      setDrop(drop, 'err', file, '✗ ' + (/\.xlsx$/i.test(file.name) && /Código/.test(e.message) ? 'Esta parece la planilla de Cashea: va en el otro recuadro.' : e.message));
    }
    run();
  }

  async function onCashea(file) {
    const drop = $('dropCashea');
    setDrop(drop, null, file, 'Leyendo…');
    try {
      state.cashea = await RefriSync.loadCashea(JSZip, await file.arrayBuffer());
      state.casheaFile = file;
      const g = state.cashea.generatedAt;
      setDrop(drop, 'ok', file, `✓ ${[...state.cashea.rows.values()].filter((r) => r.B).length} publicaciones`,
        g ? `generada por Cashea ${fmtDate(g.getTime())}` : `modificado ${ago(file.lastModified)}`);
    } catch (e) {
      state.cashea = null;
      setDrop(drop, 'err', file, '✗ ' + (/\.(xls|pdf)$/i.test(file.name) ? 'Este es un reporte de Valery: va en el otro recuadro.' : e.message));
    }
    run();
  }

  async function onDocs(files) {
    const drop = $('dropDocs');
    setDrop(drop, null, null, 'Leyendo…');
    const applied = appliedDocs();
    for (const file of files) {
      const d = { file, tipo: null, parsed: null, error: null, aplicado: null, id: null };
      try {
        const buf = await file.arrayBuffer();
        if (/\.pdf$/i.test(file.name) || new Uint8Array(buf.slice(0, 5)).every((b, i) => b === [37, 80, 68, 70, 45][i])) {
          if (!window.pdfjsLib) throw new Error('No se pudo cargar el lector de PDF.');
          d.parsed = await RefriSync.parseCompra(window.pdfjsLib, buf);
          d.tipo = 'compra';
          d.id = d.parsed.rif && d.parsed.numero ? `compra:${d.parsed.rif}:${d.parsed.numero}` : `compra:${await sha(buf)}`;
        } else {
          const bytes = new Uint8Array(buf);
          let kind = null;
          try { kind = RefriSync.detectXls(XLSX, bytes); } catch (e) { /* no es Excel */ }
          if (kind === 'listado') throw new Error('Este es el Listado de productos (inventario completo): cámbiate a "Inventario completo".');
          if (kind !== 'ventas') throw new Error('No es una Relación de ventas por producto ni una factura de compra en PDF.');
          d.parsed = RefriSync.parseVentas(XLSX, bytes);
          d.tipo = 'ventas';
          d.id = `ventas:${await sha(buf)}`;
        }
        if (state.docs.some((x) => x.id === d.id)) continue;            // ya está en la lista
        d.aplicado = applied.get(d.id) || null;
      } catch (e) { d.error = e.message; }
      state.docs.push(d);
    }
    const ok = state.docs.filter((d) => !d.error).length;
    setDrop(drop, ok ? 'ok' : (state.docs.length ? 'err' : null), null, ok ? `✓ ${pl(ok, 'reporte cargado', 'reportes cargados')}` : '');
    renderDocs();
    run();
  }

  function docLabel(d) {
    if (d.tipo === 'compra') return `Factura de compra ${d.parsed.numero || ''} · ${d.parsed.proveedor || ''} · ${d.parsed.fecha || ''}`.replace(/ · $/, '');
    if (d.tipo === 'ventas') return 'Relación de ventas por producto';
    return d.file.name;
  }

  function renderDocs() {
    const ul = $('docList');
    ul.hidden = state.mode !== 'movimientos' || !state.docs.length;
    ul.innerHTML = state.docs.map((d, i) => {
      let tag, cls = '', det;
      if (d.error) { tag = '<span class="tag error">Error</span>'; cls = 'err'; det = esc(d.error); }
      else if (d.tipo === 'ventas') {
        tag = '<span class="tag ventas">Ventas</span>';
        det = `${pl(d.parsed.lineas.length, 'producto', 'productos')} · ${fmt(d.parsed.unidades)} unid. vendidas`;
      } else {
        tag = '<span class="tag compra">Compra</span>';
        const malas = d.parsed.lineas.filter((l) => !l.cuadra).length;
        det = `${esc(docLabel(d))} · ${pl(d.parsed.lineas.length, 'producto', 'productos')} · ${fmt(d.parsed.unidades)} unid.`
          + (malas ? ` · <b>${pl(malas, 'línea no cuadra', 'líneas no cuadran')}</b>` : d.parsed.cuadraTotal ? ' · ✓ cuadra con el total' : '');
        if (malas || !d.parsed.cuadraTotal) cls = 'warn';
      }
      if (d.aplicado) { cls = 'warn'; det += ` · <b>ya se aplicó el ${esc(fmtDate(d.aplicado.aplicado))}; no se vuelve a aplicar</b>`; }
      return `<li class="${cls}">${tag}<div class="info"><div class="n">${esc(d.file.name)}</div><div class="d">${det}</div></div>
        <button type="button" data-i="${i}" title="Quitar" aria-label="Quitar ${esc(d.file.name)}">✕</button></li>`;
    }).join('');
    ul.querySelectorAll('button[data-i]').forEach((b) => b.addEventListener('click', () => {
      state.docs.splice(+b.dataset.i, 1); renderDocs();
      const ok = state.docs.filter((d) => !d.error).length;
      setDrop($('dropDocs'), ok ? 'ok' : null, null, ok ? `✓ ${pl(ok, 'reporte cargado', 'reportes cargados')}` : '');
      run();
    }));
  }
  const activeDocs = () => state.docs.filter((d) => !d.error && !d.aplicado);
  // Reportes ya aplicados justo en el último cierre: permiten regenerar la planilla (p. ej. si falló la subida).
  const replayDocs = () => {
    const last = lastCierre();
    const ids = new Set((last && last.docs) || []);
    return state.docs.filter((d) => !d.error && d.aplicado && ids.has(d.id));
  };

  /* ---------- análisis ---------- */
  const alertBox = (kind, title, body) => `<div class="alert ${kind}"><b>${title}</b>${body || ''}</div>`;
  const listLis = (arr, f, max = 12) => `<ul>${arr.slice(0, max).map(f).join('')}${arr.length > max ? `<li>… y ${arr.length - max} más (ver el reporte del cierre)</li>` : ''}</ul>`;

  function run() {
    const mov = state.mode === 'movimientos';
    let ready;
    if (mov) {
      state.baseMov = lastCierre();
      ready = !!(state.cashea && state.baseMov && (activeDocs().length || replayDocs().length));
    } else ready = !!(state.valery && state.cashea);
    ['step2', 'step3'].forEach((id) => ($(id).hidden = !ready));
    if (!ready) { state.cmp = null; return; }

    let cmp;
    if (mov) {
      state.movRes = activeDocs().length
        ? RefriSync.applyMovements(state.baseMov.valery, activeDocs().map((d) => d.parsed))
        : RefriSync.replayMovements(state.baseMov.valery, replayDocs().map((d) => d.parsed));
      cmp = RefriSync.compare(state.movRes.valery, state.cashea, { only: state.movRes.touched, sinBase: state.movRes.sinBase });
    } else {
      state.movRes = null;
      cmp = RefriSync.compare(state.valery, state.cashea);
    }
    state.cmp = cmp;
    const s = cmp.resumen;
    $('tCambia').textContent = s.cambia;
    $('tCero').textContent = cmp.aCero.length;
    $('tIgual').textContent = s.igual;
    $('tSin').textContent = s.sinTocar;
    $('tSinMov').textContent = mov ? `${fmt(s.sinMovimiento)} sin movimiento` : '';
    $('tSuben').textContent = cmp.suben.length;
    $('tBajan').textContent = cmp.bajan.length;

    if (mov) { state.base = null; state.ventas = null; state.mov = null; }
    else {
      if (!state.base) state.base = pickBase();
      const b = state.base;
      state.ventas = b && b.cashea ? RefriSync.dailySales(b, state.valery, cmp) : null;
      state.mov = b && !b.cashea ? RefriSync.movement(b.valery, state.valery) : null;
    }
    renderPeriodo();

    const a = [];
    const casheaT = state.cashea.generatedAt ? state.cashea.generatedAt.getTime() : state.casheaFile.lastModified;
    if (minutesAgo(casheaT) > MAX_MINUTOS_PLANILLA)
      a.push(alertBox('warn', `La planilla de Cashea se generó ${ago(casheaT)}.`,
        'Si hubo ventas por Cashea desde entonces, subirla les devolvería el stock anterior. Lo seguro es descargarla de nuevo.'));
    if (!mov && dayKey(valeryTime()) !== dayKey(Date.now()))
      a.push(alertBox('warn', 'El listado de Valery no es de hoy.', `Se exportó ${ago(valeryTime())}. Exporta uno nuevo antes de seguir.`));
    if (mov) movAlerts(a);
    if (s.cambia > MAX_CAMBIOS_SIN_CONFIRMAR)
      a.push(alertBox('bad', `Cambian ${s.cambia} productos, más de lo normal en un día.`,
        'Puede pasar en la primera sincronización. Si no, revisa que el reporte de Valery sea el correcto y que incluya todos los depósitos.'));
    if (s.cambia === 0)
      a.push(alertBox('info', 'Cashea ya está igual que Valery.', mov ? 'Ninguno de los productos con movimiento cambia en Cashea. Igual puedes descargar la planilla para que el cierre quede registrado.' : 'No hace falta subir nada hoy.'));
    if (cmp.revisarSku.length)
      a.push(alertBox('bad', `${cmp.revisarSku.length} SKU de Cashea apuntan a otro producto en Valery`,
        listLis(cmp.revisarSku, (r) => `<li><code>${esc(r.sku)}</code> en Cashea es «${esc(r.titulo)}», pero en Valery ese código es «${esc(r.nombreValery)}».</li>`) + 'No se les cambia el stock hasta que se corrija el SKU en Cashea.'));
    if (cmp.noPublicados.length && (mov || state.ventas)) {
      // productos que se movieron pero no están publicados en Cashea: no generan cambios
      const list = mov ? cmp.noPublicados
        : cmp.noPublicados.filter((v) => state.ventas.byCode.has(v.code) && state.ventas.byCode.get(v.code).neto !== 0);
      if (list.length)
        a.push(alertBox('info', `${pl(list.length, 'producto con movimiento no está', 'productos con movimiento no están')} publicados en Cashea`,
          listLis(list, (v) => `<li><code>${esc(v.code)}</code> ${esc(v.nombre)}</li>`) + 'Sus ventas o compras no generan cambios en Cashea. Si se quieren vender ahí, hay que publicarlos con SKU <code>REF-</code> + código.'));
    }
    if (cmp.duplicados.length) {
      const g = groupDup(cmp.duplicados);
      a.push(alertBox('warn', 'Productos publicados dos veces en Cashea',
        listLis(g, (x) => `<li>${x.map((r) => `<code>${esc(r.sku)}</code>`).join(' y ')} — ${esc(x[0].titulo)}</li>`) + 'Ambas quedan con el mismo stock; conviene pausar una para no vender dos veces la misma unidad.'));
    }
    if (cmp.decimales.length)
      a.push(alertBox('info', `${cmp.decimales.length} producto(s) con decimales o negativos`,
        listLis(cmp.decimales, (r) => `<li><code>${esc(r.sku)}</code> ${esc(r.titulo)}: ${esc(r.motivo)}</li>`)));
    $('alerts').innerHTML = a.join('');

    const needConfirm = s.cambia > MAX_CAMBIOS_SIN_CONFIRMAR;
    $('confirmBox').innerHTML = needConfirm
      ? `<label class="confirm"><input type="checkbox" id="ok"> <span>Revisé la lista: los reportes son los correctos y los ${s.cambia} cambios están bien.</span></label>`
      : '';
    const enable = () => {
      const ok = !needConfirm || $('ok').checked;
      $('dlCashea').disabled = !ok || (s.cambia === 0 && !mov);
      $('dlReport').disabled = false;
    };
    if (needConfirm) $('ok').addEventListener('change', enable);
    enable();
    if (s.cambia === 0 && !mov) saveCierre();   // no hay nada que subir: el cierre igual queda registrado
    $('outName').textContent = `cashea_stock_${stamp()}.xlsx`;
    renderRows();
  }

  function movAlerts(a) {
    const r = state.movRes, b = state.baseMov;
    const docs = activeDocs();
    if (r.replay) {
      a.push(alertBox('info', 'Estos reportes ya se aplicaron en el último cierre.',
        'Se vuelve a generar la planilla con ese resultado, <b>sin descontar ni sumar otra vez</b>. Úsalo si la subida a Cashea falló: descarga una planilla nueva de Cashea y vuelve a cargarla aquí.'));
      return;
    }
    const yaAplicados = state.docs.filter((d) => d.aplicado);
    if (yaAplicados.length)
      a.push(alertBox('warn', `${pl(yaAplicados.length, 'reporte ya se había aplicado', 'reportes ya se habían aplicado')}`,
        listLis(yaAplicados, (d) => `<li>${esc(d.file.name)} — aplicado el ${esc(fmtDate(d.aplicado.aplicado))}</li>`) + 'No se vuelven a descontar ni a sumar.'));
    for (const d of docs.filter((x) => x.tipo === 'compra')) {
      const malas = d.parsed.lineas.filter((l) => !l.cuadra);
      if (malas.length)
        a.push(alertBox('bad', `Factura ${esc(d.parsed.numero)}: ${pl(malas.length, 'línea no se pudo leer bien', 'líneas no se pudieron leer bien')}`,
          listLis(malas, (l) => `<li><code>${esc(l.code)}</code> ${esc(l.nombre)}: cantidad ${fmt(l.ordenada)} × costo ${fmt(l.costo)} no da el subtotal ${fmt(l.subtotal)}</li>`) + 'Esas líneas <b>no se aplican</b>. Ajusta el stock de esos productos a mano en Cashea.'));
      if (!d.parsed.cuadraTotal)
        a.push(alertBox('warn', `Factura ${esc(d.parsed.numero)}: la suma de las líneas (${fmt(d.parsed.sumaSubtotales)}) no da el Total Líneas (${fmt(d.parsed.totalLineas)})`,
          'Puede faltar alguna línea. Revisa la factura contra la lista de movimientos.'));
    }
    if (r.sinBase.size)
      a.push(alertBox('warn', `${pl(r.sinBase.size, 'producto nuevo', 'productos nuevos')} en Valery`,
        listLis([...r.sinBase], ([c, n]) => `<li><code>${esc(c)}</code> ${esc(n)}</li>`) + 'No estaban en el último inventario completo, así que no hay existencia de dónde partir. Se incluyen solos la próxima vez que hagas un cierre con <b>Inventario completo</b>.'));
    const neg = r.detalle.filter((d) => d.despues !== null && d.despues < 0);
    if (neg.length)
      a.push(alertBox('warn', `${pl(neg.length, 'producto quedaría', 'productos quedarían')} en negativo`,
        listLis(neg, (d) => `<li><code>${esc(d.code)}</code> ${esc(d.nombre)}: guardado ${fmt(d.antes)} − vendido ${fmt(d.ventas)}</li>`) + 'Se publica 0. Probablemente la existencia guardada estaba desactualizada: conviene un cierre con inventario completo.'));
    if (b && dayKey(b.guardado || b.t) === dayKey(Date.now()) && docs.some((d) => d.tipo === 'ventas'))
      a.push(alertBox('warn', `Ya hubo un cierre hoy a las ${new Date(b.guardado || b.t).toLocaleTimeString('es-VE', { hour: '2-digit', minute: '2-digit' })}.`,
        'Si la Relación de ventas es del día completo, las ventas anteriores a ese cierre se descontarían dos veces. Usa un reporte que empiece después de ese cierre.'));
    // Cashea vendió más de lo que dice el reporte de ventas → falta facturar algo
    const sold = RefriSync.casheaSold(b && b.cashea, state.cmp);
    const falta = r.detalle.filter((d) => (sold.get(d.code) || 0) > d.ventas);
    const fueraReporte = [...sold].filter(([c]) => !r.touched.has(c));
    if (falta.length || fueraReporte.length) {
      const items = falta.map((d) => `<li><code>${esc(d.code)}</code> ${esc(d.nombre)}: Cashea vendió ${fmt(sold.get(d.code))}, el reporte dice ${fmt(d.ventas)}</li>`)
        .concat(fueraReporte.map(([c, n]) => `<li><code>${esc(c)}</code>: Cashea vendió ${fmt(n)} y no aparece en el reporte</li>`));
      a.push(alertBox('warn', 'Cashea vendió productos que no están (o están de menos) en la Relación de ventas',
        listLis(items, (x) => x) + '¿Falta facturar alguna venta de Cashea en Valery? Mientras no se facture, esos productos se publican con el stock calculado.'));
    }
  }

  function groupDup(list) {
    const m = new Map();
    list.forEach((r) => { if (!m.has(r.codigo)) m.set(r.codigo, []); m.get(r.codigo).push(r); });
    return [...m.values()];
  }

  /* ---------- panel: ventas del día / movimientos ---------- */
  function renderPeriodo() {
    const tab = $('tabMov');
    if (state.mode === 'movimientos') {
      const r = state.movRes;
      tab.hidden = false; tab.textContent = 'Movimientos';
      const sold = RefriSync.casheaSold(state.baseMov && state.baseMov.cashea, state.cmp);
      let porCashea = 0;
      for (const d of r.detalle) porCashea += Math.min(sold.get(d.code) || 0, d.ventas);
      const nV = r.detalle.filter((d) => d.ventas > 0).length, nC = r.detalle.filter((d) => d.compras > 0).length;
      $('movBox').innerHTML = `<div class="mov"><b>Movimientos cargados</b>
        <div class="kpis">
          <div><span class="k down">${fmt(r.totales.ventas)}</span> unid. <b class="inline">vendidas</b> <span class="muted">(${pl(nV, 'producto', 'productos')})</span></div>
          <div><span class="k">${fmt(porCashea)}</span> de ellas por <b class="inline">Cashea</b> <span class="muted">· ${fmt(Math.max(0, r.totales.ventas - porCashea))} tienda y web</span></div>
          <div><span class="k up">${fmt(r.totales.compras)}</span> unid. <b class="inline">compradas</b> <span class="muted">(${pl(nC, 'producto', 'productos')})</span></div>
        </div>
        <div class="muted" style="margin-top:6px">Existencia nueva = existencia guardada en el último cierre − ventas + compras. Solo cambian en Cashea los productos con movimiento.
        <button class="linkish" id="seeMov">Ver detalle</button></div></div>`;
      $('seeMov').addEventListener('click', () => selectTab('mov'));
      return;
    }
    const v = state.ventas, m = state.mov, b = state.base;
    tab.hidden = !(v || m);
    tab.textContent = v ? 'Ventas del día' : 'Movimiento en Valery';
    if (!(v || m) && state.filter === 'mov') state.filter = 'cambios';
    const manual = (txt) => `<label class="linkish">${txt}<input type="file" id="filePrev" accept=".xls,.xlsx"></label>`;

    if (v) {
      const extra = [];
      if (v.faltaFacturar.length)
        extra.push(alertBox('warn', `En ${pl(v.faltaFacturar.length, 'producto', 'productos')} Cashea vendió más de lo que bajó Valery`,
          listLis(v.faltaFacturar, (r) => `<li><code>${esc(r.code)}</code> ${esc(r.nombre)}: Cashea vendió ${fmt(r.cashea)}, Valery bajó ${fmt(Math.max(0, r.neto))}</li>`) + '¿Falta facturar alguna venta de Cashea en Valery, o entró mercancía?'));
      if (v.ajustesCashea.length)
        extra.push(alertBox('info', `En ${pl(v.ajustesCashea.length, 'producto', 'productos')} el stock de Cashea subió desde el último cierre`,
          'Alguien lo cambió a mano en Cashea. Hoy se vuelve a poner igual que Valery.'));
      $('movBox').innerHTML = `<div class="mov"><b>Ventas desde el ${esc(b.label)}</b>
        <div class="kpis">
          <div><span class="k">${fmt(v.totales.cashea)}</span> unid. por <b class="inline">Cashea</b> <span class="muted">(${pl(v.productos.cashea, 'producto', 'productos')})</span></div>
          <div><span class="k">${fmt(v.totales.otros)}</span> unid. por <b class="inline">tienda y web</b> <span class="muted">(${pl(v.productos.otros, 'producto', 'productos')})</span></div>
          <div><span class="k up">${fmt(v.totales.entradas)}</span> unid. de <b class="inline">entradas</b> <span class="muted">(${pl(v.productos.entradas, 'producto', 'productos')})</span></div>
        </div>
        <div class="muted" style="margin-top:6px">Cashea: lo que bajó su stock desde la planilla que se subió en el cierre anterior. Tienda y web: el resto de lo que bajó en Valery.
        ${v.nuevos.length ? ` ${pl(v.nuevos.length, 'código nuevo', 'códigos nuevos')} en Valery.` : ''}
        <button class="linkish" id="seeMov">Ver detalle</button></div></div>${extra.join('')}`;
    } else if (m) {
      $('movBox').innerHTML = `<div class="mov"><b>Movimiento en Valery desde el ${esc(b.label)}</b>
        <div class="row">
          <span class="down">▼ ${pl(m.bajan.length, 'producto bajó', 'productos bajaron')} (${fmt(m.unidadesBajan)} unid.)</span>
          <span class="up">▲ ${pl(m.suben.length, 'producto subió', 'productos subieron')} (${fmt(m.unidadesSuben)} unid.)</span>
          ${m.nuevos.length ? `<span>${pl(m.nuevos.length, 'código nuevo', 'códigos nuevos')}</span>` : ''}
          ${m.quitados.length ? `<span>${pl(m.quitados.length, 'código ya no está', 'códigos ya no están')}</span>` : ''}
        </div>
        <div class="muted" style="margin-top:6px">Sin datos de Cashea de ese día, no se puede separar por canal: es el neto de ventas y entradas.
        <button class="linkish" id="seeMov">Ver detalle</button> · ${manual('comparar con otro reporte')}</div></div>`;
    } else {
      $('movBox').innerHTML = `<div class="mov"><b>Ventas del día</b>
        Todavía no hay un cierre de un día anterior guardado en esta computadora. Al descargar la planilla se guarda el de hoy, y desde el próximo cierre se verá cuánto se vendió por Cashea y por tienda/web.
        Mientras tanto puedes ${manual('cargar el listado de Valery anterior')} para ver el movimiento total.</div>`;
    }
    const see = $('seeMov'); if (see) see.addEventListener('click', () => selectTab('mov'));
    const fp = $('filePrev');
    if (fp) fp.addEventListener('change', async (ev) => {
      const f = ev.target.files[0]; if (!f) return;
      try {
        const prev = RefriSync.parseValery(XLSX, new Uint8Array(await f.arrayBuffer()));
        state.base = { label: `listado ${f.name} (${fmtDate(f.lastModified)})`, valery: RefriSync.snapshot(prev), cashea: null, manual: true };
        run();
      } catch (e) { $('movBox').insertAdjacentHTML('beforeend', alertBox('bad', 'No se pudo leer ese listado: ' + esc(e.message), '')); }
    });
  }

  function selectTabUI(f) {
    document.querySelectorAll('.tabs button').forEach((x) => x.setAttribute('aria-pressed', String(x.dataset.f === f)));
  }
  function selectTab(f) { selectTabUI(f); state.filter = f; renderRows(); }

  /* ---------- tabla ---------- */
  const casheaChange = (r) => {
    if (!r) return '<span class="muted">no publicado</span>';
    if (r.stockNuevo === null) return `<span class="muted">${fmt(r.stockAnterior)} (sin cambio)</span>`;
    if (r.stockNuevo === r.stockAnterior) return `${fmt(r.stockAnterior)}`;
    return `${fmt(r.stockAnterior)}<span class="arrow">→</span><span class="${r.stockNuevo < r.stockAnterior ? 'down' : 'up'}">${fmt(r.stockNuevo)}</span>`;
  };
  const deltaTxt = (d) => (!d ? '' : `<span class="delta ${d < 0 ? 'down' : 'up'}">${d > 0 ? '+' : '−'}${fmt(Math.abs(d))}</span>`);
  const empty = (cols) => `<tr><td colspan="${cols}" class="empty">Nada que mostrar.</td></tr>`;

  function renderRows() {
    const cmp = state.cmp;
    if (!cmp) return;
    const v = state.ventas, mov = state.mov, r0 = state.movRes;
    const q = $('q').value.trim().toLowerCase();
    const match = (a, b) => !q || String(a).toLowerCase().includes(q) || String(b).toLowerCase().includes(q);
    const byCode = new Map();
    cmp.filas.filter((r) => r.codigo && !r.revisarSku).forEach((r) => { if (!byCode.has(r.codigo)) byCode.set(r.codigo, r); });

    if (state.filter === 'mov' && r0) {
      $('thead').innerHTML = `<tr><th>Código</th><th>Producto</th><th class="num">Guardado</th><th class="num">− Vendido</th><th class="num">+ Comprado</th><th class="num">= Nuevo</th><th class="num">Cashea: antes → después</th></tr>`;
      const list = r0.detalle.filter((d) => match(d.code, d.nombre));
      $('rows').innerHTML = list.length ? list.map((d) => `<tr>
        <td class="sku">${esc(d.code)}</td><td>${esc(d.nombre)}</td>
        <td class="num">${d.antes === null ? '<span class="muted">nuevo</span>' : fmt(d.antes)}</td>
        <td class="num">${d.ventas ? `<span class="down">${fmt(d.ventas)}</span>` : '<span class="muted">0</span>'}</td>
        <td class="num">${d.compras ? `<span class="up">${fmt(d.compras)}</span>` : '<span class="muted">0</span>'}</td>
        <td class="num"><b>${d.despues === null ? '—' : fmt(d.despues)}</b></td>
        <td class="num">${d.antes === null ? '<span class="muted">sin existencia guardada</span>' : casheaChange(byCode.get(d.code))}</td></tr>`).join('')
        : empty(7);
      return;
    }
    if (state.filter === 'mov' && v) {
      $('thead').innerHTML = `<tr><th>Código</th><th>Producto (Valery)</th><th class="num">Valery: antes → ahora</th><th class="num">Cashea vendió</th><th class="num">Tienda/web vendió</th><th class="num">Entró</th><th>Nota</th></tr>`;
      const list = v.rows.filter((r) => match(r.code, r.nombre));
      $('rows').innerHTML = list.length ? list.map((r) => `<tr>
        <td class="sku">${esc(r.code)}</td><td>${esc(r.nombre)}</td>
        <td class="num">${fmt(r.antes)}<span class="arrow">→</span>${fmt(r.ahora)}</td>
        <td class="num">${r.cashea ? `<b>${fmt(r.cashea)}</b>` : '<span class="muted">0</span>'}</td>
        <td class="num">${r.otros ? `<b>${fmt(r.otros)}</b>` : '<span class="muted">0</span>'}</td>
        <td class="num">${r.entrada ? `<span class="up">+${fmt(r.entrada)}</span>` : '<span class="muted">—</span>'}</td>
        <td class="muted">${[r.faltaFacturar ? '¿falta facturar?' : '', r.ajusteCashea ? `subido a mano en Cashea (+${fmt(r.ajusteCashea)})` : '', !r.skus.length ? 'no publicado en Cashea' : ''].filter(Boolean).join(' · ')}</td></tr>`).join('')
        : empty(7);
      return;
    }
    if (state.filter === 'mov' && mov) {
      $('thead').innerHTML = `<tr><th>Código</th><th>Producto (Valery)</th><th class="num">Valery: antes → ahora</th><th class="num">Cashea: antes → después</th><th>Movimiento</th></tr>`;
      const list = mov.list.filter((m) => match(m.code, m.nombre));
      $('rows').innerHTML = list.length ? list.map((m) => `<tr>
        <td class="sku">${esc(m.code)}</td><td>${esc(m.nombre)}</td>
        <td class="num">${m.nuevo ? '—' : fmt(m.antes)}<span class="arrow">→</span>${fmt(m.ahora)}</td>
        <td class="num">${casheaChange(byCode.get(m.code))}</td><td>${m.nuevo ? '<span class="delta muted">nuevo</span>' : deltaTxt(m.delta)}</td></tr>`).join('')
        : empty(5);
      return;
    }
    if (state.filter === 'mov') state.filter = 'cambios';

    $('thead').innerHTML = `<tr><th>SKU</th><th>Producto</th><th class="num">Valery</th><th class="num">Cashea: antes → después</th><th>Nota</th></tr>`;
    const lists = { cambios: cmp.cambios, suben: cmp.suben, bajan: cmp.bajan, cero: cmp.aCero, sin: cmp.sinTocar,
                    todo: r0 ? cmp.filas.filter((r) => r.accion !== 'Sin movimiento') : cmp.filas };
    const list = lists[state.filter].filter((r) => match(r.sku, r.titulo));
    if (!list.length) { $('rows').innerHTML = empty(5); return; }
    $('rows').innerHTML = list.map((r) => {
      let d = null;
      if (r.codigo && v && v.byCode.has(r.codigo)) d = -v.byCode.get(r.codigo).neto;
      else if (r.codigo && mov && mov.byCode.has(r.codigo)) d = mov.byCode.get(r.codigo).delta;
      else if (r.codigo && r0) { const x = r0.detalle.find((y) => y.code === r.codigo); if (x && x.antes !== null) d = x.compras - x.ventas; }
      return `<tr class="${r.accion === 'Cambia' && r.stockNuevo === 0 ? 'zero' : ''}">
        <td class="sku">${esc(r.sku)}</td><td>${esc(r.titulo)}</td>
        <td class="num">${fmt(r.existenciaValery)}${deltaTxt(d)}</td><td class="num">${casheaChange(r)}</td>
        <td class="muted">${esc(r.motivo)}</td></tr>`;
    }).join('');
  }

  /* ---------- descargas ---------- */
  function download(bytes, name) {
    const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const a = Object.assign(document.createElement('a'), { href: URL.createObjectURL(blob), download: name });
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
  }

  $('dlCashea').addEventListener('click', async () => {
    const btn = $('dlCashea'), label = btn.innerHTML;
    btn.disabled = true; btn.innerHTML = '<span class="spinner"></span> Generando…';
    try {
      const name = `cashea_stock_${stamp()}.xlsx`;
      download(await RefriSync.buildCasheaFile(state.cashea, state.cmp), name);
      $('outName').textContent = name;
      if (!(state.movRes && state.movRes.replay)) saveCierre();
      if (state.mode === 'movimientos') {          // los reportes quedan marcados como aplicados
        const applied = appliedDocs();
        state.docs.forEach((d) => { if (!d.error && applied.has(d.id)) d.aplicado = applied.get(d.id); });
        renderDocs(); renderBaseInfo();
        $('confirmBox').insertAdjacentHTML('beforeend', alertBox('info', 'Cierre registrado.',
          'Los reportes quedaron marcados como aplicados. Sube la planilla a Cashea ahora; si vuelves a cargar estos reportes, no se aplicarán dos veces.'));
        btn.innerHTML = label; return;
      }
    } catch (e) { $('confirmBox').insertAdjacentHTML('beforeend', alertBox('bad', 'No se pudo generar la planilla.', esc(e.message))); }
    btn.innerHTML = label; btn.disabled = false;
  });
  $('dlReport').addEventListener('click', () =>
    download(RefriSync.buildReport(XLSX, state.cmp, state.mov, state.base && state.base.label, state.ventas, state.movRes), `reporte_cierre_${stamp()}.xlsx`));

  document.querySelectorAll('.tabs button').forEach((b) => b.addEventListener('click', () => selectTab(b.dataset.f)));
  document.querySelectorAll('.modes button').forEach((b) => b.addEventListener('click', () => setMode(b.dataset.mode)));
  $('q').addEventListener('input', () => state.cmp && renderRows());

  wireDrop('dropValery', 'fileValery', onValery);
  wireDrop('dropCashea', 'fileCashea', onCashea);
  wireDrop('dropDocs', 'fileDocs', onDocs, true);
  setMode(store.get(MODE_KEY, 'completo') === 'movimientos' ? 'movimientos' : 'completo');
})();
