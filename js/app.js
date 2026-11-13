/* Interfaz de la página. La lógica está en sync.js. */
(function () {
  'use strict';

  const MAX_CAMBIOS_SIN_CONFIRMAR = 50;   // más cambios que esto pide confirmación explícita
  const MAX_MINUTOS_PLANILLA = 60;        // planilla de Cashea más vieja que esto: advertencia

  const $ = (id) => document.getElementById(id);
  const state = { valery: null, cashea: null, casheaFile: null, valeryFile: null, cmp: null, filter: 'cambios' };

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

  function groupDup(list) {
    const m = new Map();
    list.forEach((r) => { if (!m.has(r.codigo)) m.set(r.codigo, []); m.get(r.codigo).push(r); });
    return [...m.values()];
  }

  function renderRows() {
    const cmp = state.cmp;
    const lists = { cambios: cmp.cambios, cero: cmp.aCero, sin: cmp.sinTocar, todo: cmp.filas };
    const q = $('q').value.trim().toLowerCase();
    const list = lists[state.filter].filter((r) => !q || r.sku.toLowerCase().includes(q) || r.titulo.toLowerCase().includes(q));
    if (!list.length) { $('rows').innerHTML = `<tr><td colspan="5" class="empty">Nada que mostrar.</td></tr>`; return; }
    $('rows').innerHTML = list.map((r) => {
      let change;
      if (r.stockNuevo === null) change = `<span class="muted">${fmt(r.stockAnterior)} (sin cambio)</span>`;
      else if (r.stockNuevo === r.stockAnterior) change = `${fmt(r.stockAnterior)}`;
      else change = `${fmt(r.stockAnterior)}<span class="arrow">→</span><span class="${r.stockNuevo < r.stockAnterior ? 'down' : 'up'}">${fmt(r.stockNuevo)}</span>`;
      return `<tr class="${r.accion === 'Cambia' && r.stockNuevo === 0 ? 'zero' : ''}">
        <td class="sku">${esc(r.sku)}</td><td>${esc(r.titulo)}</td>
        <td class="num">${fmt(r.existenciaValery)}</td><td class="num">${change}</td>
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
    } catch (e) { $('confirmBox').insertAdjacentHTML('beforeend', alertBox('bad', 'No se pudo generar la planilla.', esc(e.message))); }
    btn.innerHTML = label; btn.disabled = false;
  });
  $('dlReport').addEventListener('click', () => download(RefriSync.buildReport(XLSX, state.cmp), `reporte_cierre_${stamp()}.xlsx`));

  document.querySelectorAll('.tabs button').forEach((b) => b.addEventListener('click', () => {
    document.querySelectorAll('.tabs button').forEach((x) => x.setAttribute('aria-pressed', String(x === b)));
    state.filter = b.dataset.f; renderRows();
  }));
  $('q').addEventListener('input', () => state.cmp && renderRows());

  wireDrop('dropValery', 'fileValery', onValery);
  wireDrop('dropCashea', 'fileCashea', onCashea);
})();
