/* Interfaz de la página. La lógica está en sync.js. */
(function () {
  'use strict';

  const MAX_CAMBIOS_SIN_CONFIRMAR = 50;   // más cambios que esto pide confirmación explícita
  const MAX_MINUTOS_PLANILLA = 60;        // planilla de Cashea más vieja que esto: advertencia

  const STORE_KEY = 'refriceiba-sync/cierres';
  const MAX_CIERRES = 14;                 // fotos de Valery guardadas (una por día)

  const $ = (id) => document.getElementById(id);
  const state = { valery: null, cashea: null, casheaFile: null, valeryFile: null, cmp: null, filter: 'cambios',
                  base: null, mov: null };

  /* ---- Cierres guardados en el navegador: [{ t, archivo, items: { código: existencia } }] ---- */
  const dayKey = (t) => new Date(t).toDateString();
  function loadCierres() {
    try { const v = JSON.parse(localStorage.getItem(STORE_KEY) || '[]'); return Array.isArray(v) ? v : []; }
    catch (e) { return []; }
  }
  function saveCierre() {
    if (!state.valery) return false;
    try {
      const now = Date.now();
      const list = loadCierres().filter((c) => dayKey(c.t) !== dayKey(now));   // uno por día: el último gana
      list.push({ t: now, archivo: state.valeryFile ? state.valeryFile.name : '', items: RefriSync.snapshot(state.valery) });
      list.sort((a, b) => a.t - b.t);
      localStorage.setItem(STORE_KEY, JSON.stringify(list.slice(-MAX_CIERRES)));
      return true;
    } catch (e) { return false; }
  }
  // Base de comparación: el último cierre guardado de un día anterior a hoy.
  function pickBase() {
    const today = dayKey(Date.now());
    const prev = loadCierres().filter((c) => dayKey(c.t) !== today && c.t < Date.now());
    return prev.length ? { label: `cierre del ${fmtDate(prev[prev.length - 1].t)}`, items: prev[prev.length - 1].items, manual: false } : null;
  }
  const pl = (n, uno, varios) => `${n} ${n === 1 ? uno : varios}`;
  const fmtDate = (t) => new Date(t).toLocaleString('es-VE', { day: 'numeric', month: 'numeric', hour: '2-digit', minute: '2-digit' });

  const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const fmt = (n) => (n === null || n === undefined ? '—' : Number(n).toLocaleString('es-VE', { maximumFractionDigits: 2 }));
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

  function setDrop(el, kind, file, msg) {
    el.classList.remove('ok', 'err');
    if (kind) el.classList.add(kind);
    el.querySelector('.file').textContent = file ? `${file.name} · modificado ${ago(file.lastModified)}` : '';
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
      setDrop(drop, 'ok', file, `✓ ${state.valery.size} productos`);
    } catch (e) {
      state.valery = null;
      setDrop(drop, 'err', file, '✗ ' + (/\.xlsx$/i.test(file.name) && /Código/.test(e.message) ? e.message + ' Parece que es la planilla de Cashea: va en el otro recuadro.' : e.message));
    }
    run();
  }

  async function onCashea(file) {
    const drop = $('dropCashea');
    setDrop(drop, null, file, 'Leyendo…');
    try {
      state.cashea = await RefriSync.loadCashea(JSZip, await file.arrayBuffer());
      state.casheaFile = file;
      setDrop(drop, 'ok', file, `✓ ${[...state.cashea.rows.values()].filter((r) => r.B).length} publicaciones`);
    } catch (e) {
      state.cashea = null;
      setDrop(drop, 'err', file, '✗ ' + (/\.xls$/i.test(file.name) ? 'Este parece el reporte de Valery: va en el otro recuadro.' : e.message));
    }
    run();
  }

  function alertBox(kind, title, body) {
    return `<div class="alert ${kind}"><b>${title}</b>${body || ''}</div>`;
  }

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

    if (!state.base || !state.base.manual) state.base = pickBase();
    state.mov = state.base ? RefriSync.movement(state.base.items, state.valery) : null;
    renderMov();
    if (s.cambia === 0) saveCierre();   // no hay nada que subir: el cierre igual queda registrado

    const a = [];
    const mins = minutesAgo(state.casheaFile.lastModified);
    if (mins > MAX_MINUTOS_PLANILLA)
      a.push(alertBox('warn', `La planilla de Cashea es de ${ago(state.casheaFile.lastModified)}.`,
        'Si hubo ventas por Cashea desde entonces, subirla les devolvería el stock anterior. Lo seguro es descargarla de nuevo.'));
    const vDate = new Date(state.valeryFile.lastModified).toDateString();
    if (vDate !== new Date().toDateString())
      a.push(alertBox('warn', 'El reporte de Valery no es de hoy.', `Fue modificado ${ago(state.valeryFile.lastModified)}. Exporta uno nuevo antes de seguir.`));
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
    $('outName').textContent = `cashea_stock_${stamp()}.xlsx`;
    renderRows();
  }

  function renderMov() {
    const m = state.mov;
    $('tabMov').hidden = !m;
    if (!m && state.filter === 'mov') state.filter = 'cambios';
    const manual = `<label class="linkish">cargar el reporte de Valery anterior<input type="file" id="filePrev" accept=".xls,.xlsx"></label>`;
    if (!m) {
      $('movBox').innerHTML = `<div class="mov"><b>Movimiento en Valery</b>
        Todavía no hay un cierre anterior guardado en esta computadora. Al descargar la planilla se guarda el de hoy, y desde el próximo cierre se verá cuánto se movió cada producto. También puedes ${manual}.</div>`;
    } else {
      $('movBox').innerHTML = `<div class="mov"><b>Movimiento en Valery desde el ${esc(state.base.label)}</b>
        <div class="row">
          <span class="down">▼ ${pl(m.bajan.length, 'producto bajó', 'productos bajaron')} (${fmt(m.unidadesBajan)} unid.)</span>
          <span class="up">▲ ${pl(m.suben.length, 'producto subió', 'productos subieron')} (${fmt(m.unidadesSuben)} unid.)</span>
          ${m.nuevos.length ? `<span>${pl(m.nuevos.length, 'código nuevo', 'códigos nuevos')}</span>` : ''}
          ${m.quitados.length ? `<span>${pl(m.quitados.length, 'código ya no está', 'códigos ya no están')}</span>` : ''}
        </div>
        <div class="muted" style="margin-top:6px">Incluye todas las ventas facturadas y las entradas de mercancía; no distingue por canal.
        <button class="linkish" id="seeMov">Ver detalle</button> · ${manual.replace('cargar el reporte de Valery anterior', 'comparar con otro reporte')}</div></div>`;
      $('seeMov').addEventListener('click', () => selectTab('mov'));
    }
    $('filePrev').addEventListener('change', async (ev) => {
      const f = ev.target.files[0]; if (!f) return;
      try {
        const prev = RefriSync.parseValery(XLSX, new Uint8Array(await f.arrayBuffer()));
        state.base = { label: `reporte ${f.name} (${fmtDate(f.lastModified)})`, items: RefriSync.snapshot(prev), manual: true };
        run();
      } catch (e) { alertMov('No se pudo leer ese reporte: ' + e.message); }
    });
  }
  function alertMov(msg) { $('movBox').insertAdjacentHTML('beforeend', alertBox('bad', msg, '')); }

  function selectTab(f) {
    document.querySelectorAll('.tabs button').forEach((x) => x.setAttribute('aria-pressed', String(x.dataset.f === f)));
    state.filter = f; renderRows();
  }

  function groupDup(list) {
    const m = new Map();
    list.forEach((r) => { if (!m.has(r.codigo)) m.set(r.codigo, []); m.get(r.codigo).push(r); });
    return [...m.values()];
  }

  const casheaChange = (r) => {
    if (!r) return '<span class="muted">no publicado</span>';
    if (r.stockNuevo === null) return `<span class="muted">${fmt(r.stockAnterior)} (sin cambio)</span>`;
    if (r.stockNuevo === r.stockAnterior) return `${fmt(r.stockAnterior)}`;
    return `${fmt(r.stockAnterior)}<span class="arrow">→</span><span class="${r.stockNuevo < r.stockAnterior ? 'down' : 'up'}">${fmt(r.stockNuevo)}</span>`;
  };
  const deltaTxt = (m) => m.nuevo ? '<span class="delta muted">nuevo</span>'
    : `<span class="delta ${m.delta < 0 ? 'down' : 'up'}">${m.delta > 0 ? '+' : '−'}${fmt(Math.abs(m.delta))}</span>`;

  function renderRows() {
    const cmp = state.cmp, mov = state.mov;
    const q = $('q').value.trim().toLowerCase();
    const match = (a, b) => !q || String(a).toLowerCase().includes(q) || String(b).toLowerCase().includes(q);

    if (state.filter === 'mov' && mov) {
      $('thead').innerHTML = `<tr><th>Código</th><th>Producto (Valery)</th><th class="num">Valery: antes → ahora</th><th class="num">Cashea: antes → después</th><th>Movimiento</th></tr>`;
      const byCode = new Map(cmp.filas.filter((r) => r.codigo).map((r) => [r.codigo, r]));
      const list = mov.list.filter((m) => match(m.code, m.nombre));
      $('rows').innerHTML = list.length ? list.map((m) => `<tr>
        <td class="sku">${esc(m.code)}</td><td>${esc(m.nombre)}</td>
        <td class="num">${m.nuevo ? '—' : fmt(m.antes)}<span class="arrow">→</span>${fmt(m.ahora)}</td>
        <td class="num">${casheaChange(byCode.get(m.code))}</td><td>${deltaTxt(m)}</td></tr>`).join('')
        : `<tr><td colspan="5" class="empty">Nada que mostrar.</td></tr>`;
      return;
    }

    $('thead').innerHTML = `<tr><th>SKU</th><th>Producto</th><th class="num">Valery</th><th class="num">Cashea: antes → después</th><th>Nota</th></tr>`;
    const lists = { cambios: cmp.cambios, suben: cmp.suben, bajan: cmp.bajan, cero: cmp.aCero, sin: cmp.sinTocar, todo: cmp.filas };
    const list = lists[state.filter].filter((r) => match(r.sku, r.titulo));
    if (!list.length) { $('rows').innerHTML = `<tr><td colspan="5" class="empty">Nada que mostrar.</td></tr>`; return; }
    $('rows').innerHTML = list.map((r) => {
      const m = mov && r.codigo ? mov.byCode.get(r.codigo) : null;
      return `<tr class="${r.accion === 'Cambia' && r.stockNuevo === 0 ? 'zero' : ''}">
        <td class="sku">${esc(r.sku)}</td><td>${esc(r.titulo)}</td>
        <td class="num">${fmt(r.existenciaValery)}${m ? deltaTxt(m) : ''}</td><td class="num">${casheaChange(r)}</td>
        <td class="muted">${esc(r.motivo)}</td></tr>`;
    }).join('');
  }

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
    download(RefriSync.buildReport(XLSX, state.cmp, state.mov, state.base && state.base.label), `reporte_cierre_${stamp()}.xlsx`));

  document.querySelectorAll('.tabs button').forEach((b) => b.addEventListener('click', () => selectTab(b.dataset.f)));
  $('q').addEventListener('input', () => state.cmp && renderRows());

  wireDrop('dropValery', 'fileValery', onValery);
  wireDrop('dropCashea', 'fileCashea', onCashea);
})();
