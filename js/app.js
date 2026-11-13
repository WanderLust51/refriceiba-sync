/* Interfaz de la página. La lógica está en sync.js. */
(function () {
  'use strict';

  const MAX_CAMBIOS_SIN_CONFIRMAR = 50;   // más cambios que esto pide confirmación explícita
  const MAX_MINUTOS_PLANILLA = 60;        // planilla de Cashea más vieja que esto: advertencia
  const STORE_KEY = 'refriceiba-sync/cierres';
  const MAX_CIERRES = 14;                 // cierres guardados (uno por día)

  const $ = (id) => document.getElementById(id);
  const state = { valery: null, cashea: null, casheaFile: null, valeryFile: null, cmp: null, filter: 'cambios',
                  base: null, ventas: null, mov: null };

  /* ---------- utilidades ---------- */
  const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const fmt = (n) => (n === null || n === undefined ? '—' : Number(n).toLocaleString('es-VE', { maximumFractionDigits: 2 }));
  const pl = (n, uno, varios) => `${fmt(n)} ${n === 1 ? uno : varios}`;
  const fmtDate = (t) => new Date(t).toLocaleString('es-VE', { weekday: 'short', day: 'numeric', month: 'numeric', hour: '2-digit', minute: '2-digit' });
  const fmtTime = (t) => new Date(t).toLocaleTimeString('es-VE', { hour: '2-digit', minute: '2-digit' });
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
  // Fecha del reporte de Valery: el .xls no trae fecha adentro; se usa la hora en que se exportó el archivo.
  const valeryTime = () => (state.valeryFile ? state.valeryFile.lastModified : Date.now());

  /* ---------- cierres guardados en este navegador ----------
     [{ t: fecha del reporte de Valery, guardado, archivo,
        valery: { código: existencia }, cashea: { SKU: stock subido }, casheaT: fecha de la planilla }] */
  function loadCierres() {
    try {
      const v = JSON.parse(localStorage.getItem(STORE_KEY) || '[]');
      return Array.isArray(v) ? v.map((c) => ({ ...c, valery: c.valery || c.items || {} })) : [];
    } catch (e) { return []; }
  }
  function saveCierre() {
    if (!state.valery || !state.cmp) return false;
    try {
      const t = valeryTime();
      const list = loadCierres().filter((c) => dayKey(c.t) !== dayKey(t));   // uno por día: el último gana
      list.push({
        t, guardado: Date.now(), archivo: state.valeryFile ? state.valeryFile.name : '',
        valery: RefriSync.snapshot(state.valery),
        cashea: RefriSync.casheaSnapshot(state.cmp),
        casheaT: state.cashea && state.cashea.generatedAt ? state.cashea.generatedAt.getTime() : null,
      });
      list.sort((a, b) => a.t - b.t);
      localStorage.setItem(STORE_KEY, JSON.stringify(list.slice(-MAX_CIERRES).map(({ items, ...c }) => c)));
      return true;
    } catch (e) { return false; }
  }
  // Base: el último cierre de un día ANTERIOR al del reporte de Valery que se está usando.
  // Así, aunque hoy se haga el cierre varias veces, siempre se compara contra ayer (o el último día con cierre).
  function pickBase() {
    const hoy = dayKey(valeryTime());
    const prev = loadCierres().filter((c) => c.t < valeryTime() && dayKey(c.t) !== hoy);
    if (!prev.length) return null;
    const c = prev[prev.length - 1];
    return { label: `cierre del ${fmtDate(c.t)}`, valery: c.valery, cashea: c.cashea || null, t: c.t, manual: false };
  }

  /* ---------- carga de archivos ---------- */
  function setDrop(el, kind, file, msg, extra) {
    el.classList.remove('ok', 'err');
    if (kind) el.classList.add(kind);
    el.querySelector('.file').textContent = file ? `${file.name}${extra ? ' · ' + extra : ''}` : '';
    el.querySelector('.state').textContent = msg || '';
  }

  function wireDrop(dropId, inputId, handler) {
    const drop = $(dropId), input = $(inputId);
    input.addEventListener('change', () => input.files[0] && handler(input.files[0]));
    ['dragenter', 'dragover'].forEach((e) => drop.addEventListener(e, (ev) => { ev.preventDefault(); drop.classList.add('over'); }));
    ['dragleave', 'drop'].forEach((e) => drop.addEventListener(e, () => drop.classList.remove('over')));
    drop.addEventListener('drop', (ev) => { ev.preventDefault(); const f = ev.dataTransfer.files[0]; if (f) handler(f); });
  }

  async function onValery(file) {
    const drop = $('dropValery');
    setDrop(drop, null, file, 'Leyendo…');
    try {
      state.valery = RefriSync.parseValery(XLSX, new Uint8Array(await file.arrayBuffer()));
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
      setDrop(drop, 'err', file, '✗ ' + (/\.xls$/i.test(file.name) ? 'Este parece el reporte de Valery: va en el otro recuadro.' : e.message));
    }
    run();
  }

  /* ---------- análisis ---------- */
  const alertBox = (kind, title, body) => `<div class="alert ${kind}"><b>${title}</b>${body || ''}</div>`;

  function run() {
    const ready = state.valery && state.cashea;
    ['step2', 'step3'].forEach((id) => ($(id).hidden = !ready));
    if (!ready) return;
    const cmp = (state.cmp = RefriSync.compare(state.valery, state.cashea));
    const s = cmp.resumen;
    $('tCambia').textContent = s.cambia;
    $('tCero').textContent = cmp.aCero.length;
    $('tIgual').textContent = s.igual;
    $('tSin').textContent = s.sinTocar;
    $('tSuben').textContent = cmp.suben.length;
    $('tBajan').textContent = cmp.bajan.length;

    if (!state.base) state.base = pickBase();
    const b = state.base;
    state.ventas = b && b.cashea ? RefriSync.dailySales(b, state.valery, cmp) : null;
    state.mov = b && !b.cashea ? RefriSync.movement(b.valery, state.valery) : null;
    renderPeriodo();

    const a = [];
    const casheaT = state.cashea.generatedAt ? state.cashea.generatedAt.getTime() : state.casheaFile.lastModified;
    if (minutesAgo(casheaT) > MAX_MINUTOS_PLANILLA)
      a.push(alertBox('warn', `La planilla de Cashea se generó ${ago(casheaT)}.`,
        'Si hubo ventas por Cashea desde entonces, subirla les devolvería el stock anterior. Lo seguro es descargarla de nuevo.'));
    if (dayKey(valeryTime()) !== dayKey(Date.now()))
      a.push(alertBox('warn', 'El reporte de Valery no es de hoy.', `Se exportó ${ago(valeryTime())}. Exporta uno nuevo antes de seguir.`));
    if (s.cambia > MAX_CAMBIOS_SIN_CONFIRMAR)
      a.push(alertBox('bad', `Cambian ${s.cambia} productos, más de lo normal en un día.`,
        'Puede pasar en la primera sincronización. Si no, revisa que el reporte de Valery sea el correcto y que incluya todos los depósitos.'));
    if (s.cambia === 0)
      a.push(alertBox('info', 'Cashea ya está igual que Valery.', 'No hace falta subir nada hoy.'));
    if (cmp.revisarSku.length)
      a.push(alertBox('bad', `${cmp.revisarSku.length} SKU de Cashea apuntan a otro producto en Valery`,
        `<ul>${cmp.revisarSku.map((r) => `<li><code>${esc(r.sku)}</code> en Cashea es «${esc(r.titulo)}», pero en Valery ese código es «${esc(r.nombreValery)}».</li>`).join('')}</ul>No se les cambia el stock hasta que se corrija el SKU en Cashea.`));
    if (cmp.duplicados.length)
      a.push(alertBox('warn', 'Productos publicados dos veces en Cashea',
        `<ul>${groupDup(cmp.duplicados).map((g) => `<li>${g.map((r) => `<code>${esc(r.sku)}</code>`).join(' y ')} — ${esc(g[0].titulo)}</li>`).join('')}</ul>Ambas quedan con el mismo stock; conviene pausar una para no vender dos veces la misma unidad.`));
    if (cmp.decimales.length)
      a.push(alertBox('info', `${cmp.decimales.length} producto(s) con decimales o negativos en Valery`,
        `<ul>${cmp.decimales.map((r) => `<li><code>${esc(r.sku)}</code> ${esc(r.titulo)}: ${esc(r.motivo)}</li>`).join('')}</ul>`));
    $('alerts').innerHTML = a.join('');

    const needConfirm = s.cambia > MAX_CAMBIOS_SIN_CONFIRMAR;
    $('confirmBox').innerHTML = needConfirm
      ? `<label class="confirm"><input type="checkbox" id="ok"> <span>Revisé la lista: el reporte de Valery es de hoy y los ${s.cambia} cambios son correctos.</span></label>`
      : '';
    const enable = () => {
      const ok = !needConfirm || $('ok').checked;
      $('dlCashea').disabled = !ok || s.cambia === 0;
      $('dlReport').disabled = false;
    };
    if (needConfirm) $('ok').addEventListener('change', enable);
    enable();
    if (s.cambia === 0) saveCierre();   // no hay nada que subir: el cierre igual queda registrado
    $('outName').textContent = `cashea_stock_${stamp()}.xlsx`;
    renderRows();
  }

  function groupDup(list) {
    const m = new Map();
    list.forEach((r) => { if (!m.has(r.codigo)) m.set(r.codigo, []); m.get(r.codigo).push(r); });
    return [...m.values()];
  }

  /* ---------- panel: ventas del día / movimiento ---------- */
  function renderPeriodo() {
    const v = state.ventas, m = state.mov, b = state.base;
    const tab = $('tabMov');
    tab.hidden = !(v || m);
    tab.textContent = v ? 'Ventas del día' : 'Movimiento en Valery';
    if (!(v || m) && state.filter === 'mov') state.filter = 'cambios';
    const manual = (txt) => `<label class="linkish">${txt}<input type="file" id="filePrev" accept=".xls,.xlsx"></label>`;

    if (v) {
      const extra = [];
      if (v.faltaFacturar.length)
        extra.push(alertBox('warn', `En ${pl(v.faltaFacturar.length, 'producto', 'productos')} Cashea vendió más de lo que bajó Valery`,
          `<ul>${v.faltaFacturar.map((r) => `<li><code>${esc(r.code)}</code> ${esc(r.nombre)}: Cashea vendió ${fmt(r.cashea)}, Valery bajó ${fmt(Math.max(0, r.neto))}</li>`).join('')}</ul>¿Falta facturar alguna venta de Cashea en Valery, o entró mercancía?`));
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
        Todavía no hay un cierre anterior guardado en esta computadora. Al descargar la planilla se guarda el de hoy (existencias de Valery y stock subido a Cashea), y desde el próximo cierre se verá cuánto se vendió por Cashea y por tienda/web.
        Mientras tanto puedes ${manual('cargar el reporte de Valery anterior')} para ver el movimiento total.</div>`;
    }
    const see = $('seeMov'); if (see) see.addEventListener('click', () => selectTab('mov'));
    const fp = $('filePrev');
    if (fp) fp.addEventListener('change', async (ev) => {
      const f = ev.target.files[0]; if (!f) return;
      try {
        const prev = RefriSync.parseValery(XLSX, new Uint8Array(await f.arrayBuffer()));
        state.base = { label: `reporte ${f.name} (${fmtDate(f.lastModified)})`, valery: RefriSync.snapshot(prev), cashea: null, manual: true };
        run();
      } catch (e) { $('movBox').insertAdjacentHTML('beforeend', alertBox('bad', 'No se pudo leer ese reporte: ' + esc(e.message), '')); }
    });
  }

  function selectTab(f) {
    document.querySelectorAll('.tabs button').forEach((x) => x.setAttribute('aria-pressed', String(x.dataset.f === f)));
    state.filter = f; renderRows();
  }

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
    const cmp = state.cmp, v = state.ventas, mov = state.mov;
    const q = $('q').value.trim().toLowerCase();
    const match = (a, b) => !q || String(a).toLowerCase().includes(q) || String(b).toLowerCase().includes(q);
    const byCode = new Map(cmp.filas.filter((r) => r.codigo && !r.revisarSku).map((r) => [r.codigo, r]));

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

    $('thead').innerHTML = `<tr><th>SKU</th><th>Producto</th><th class="num">Valery</th><th class="num">Cashea: antes → después</th><th>Nota</th></tr>`;
    const lists = { cambios: cmp.cambios, suben: cmp.suben, bajan: cmp.bajan, cero: cmp.aCero, sin: cmp.sinTocar, todo: cmp.filas };
    const list = lists[state.filter].filter((r) => match(r.sku, r.titulo));
    if (!list.length) { $('rows').innerHTML = empty(5); return; }
    $('rows').innerHTML = list.map((r) => {
      let d = null;
      if (r.codigo && v && v.byCode.has(r.codigo)) d = -v.byCode.get(r.codigo).neto;
      else if (r.codigo && mov && mov.byCode.has(r.codigo)) d = mov.byCode.get(r.codigo).delta;
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
      saveCierre();
    } catch (e) { $('confirmBox').insertAdjacentHTML('beforeend', alertBox('bad', 'No se pudo generar la planilla.', esc(e.message))); }
    btn.innerHTML = label; btn.disabled = false;
  });
  $('dlReport').addEventListener('click', () =>
    download(RefriSync.buildReport(XLSX, state.cmp, state.mov, state.base && state.base.label, state.ventas), `reporte_cierre_${stamp()}.xlsx`));

  document.querySelectorAll('.tabs button').forEach((b) => b.addEventListener('click', () => selectTab(b.dataset.f)));
  $('q').addEventListener('input', () => state.cmp && renderRows());

  wireDrop('dropValery', 'fileValery', onValery);
  wireDrop('dropCashea', 'fileCashea', onCashea);
})();
