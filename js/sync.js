/*
 * Lógica de sincronización Valery → Cashea.
 * Funciona igual en el navegador (usa window.XLSX y window.JSZip) y en Node (para las pruebas).
 *
 * Reglas:
 *  - SKU "REF-123" o "123" en Cashea = Código 123 en Valery.
 *  - Stock = parte entera de la existencia de Valery; negativo → 0.
 *  - Solo se modifica la columna Stock (C) de la planilla de Cashea; todo lo demás queda intacto.
 *
 * Dos modos de carga:
 *  - Inventario completo: Listado de Productos de Valery (existencia de todos los productos).
 *  - Movimientos del día: Relación de Ventas por Producto (.xls) y/o facturas de compra (PDF),
 *    aplicados sobre la última existencia guardada:  existencia = guardada − ventas + compras.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.RefriSync = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const FIRST_DATA_ROW = 5;

  function valeryCode(sku) {
    if (sku === null || sku === undefined) return null;
    const m = String(sku).trim().match(/^(?:REF-)?0*(\d+)$/i);
    return m ? m[1] : null;
  }

  // Redondeo a 6 decimales antes de truncar: 2.9999999999 de Valery cuenta como 3.
  function targetStock(existencia) {
    const clean = Math.round(existencia * 1e6) / 1e6;
    return Math.max(0, Math.floor(clean));
  }

  function noteFor(existencia) {
    const notes = [];
    const clean = Math.round(existencia * 1e6) / 1e6;
    if (clean < 0) notes.push('negativo en Valery → 0');
    if (clean % 1 !== 0) notes.push(`decimal ${clean} → ${targetStock(existencia)}`);
    return notes.join('; ');
  }

  const norm = (s) => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').trim().toLowerCase();

  // Palabras de 3+ letras y números de un nombre de producto.
  const STOP = new Set(['para', 'con', 'sin', 'por', 'del', 'los', 'las', 'tipo', 'marca', 'de']);
  const tokens = (s) => (norm(s).match(/[a-z]{3,}|\d+/g) || []).filter((t) => !STOP.has(t));

  // ¿El título de Cashea y el nombre de Valery describen el mismo producto?
  // Basta una palabra en común; "pega" cuenta igual que "pegamento" (prefijo de 4+ letras).
  function sameProduct(a, b) {
    const ta = tokens(a), tb = tokens(b);
    if (!ta.length || !tb.length) return true;
    return ta.some((x) => tb.some((y) => x === y || (x.length >= 4 && y.length >= 4 && (x.startsWith(y) || y.startsWith(x)))));
  }

  /* ---------------- Valery ---------------- */
  function parseValery(XLSX, data) {
    const wb = XLSX.read(data, { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    const hIdx = rows.findIndex((r) => r.some((c) => norm(c) === 'codigo') && r.some((c) => norm(c) === 'existencia actual'));
    if (hIdx < 0) throw new Error('El archivo de Valery no tiene las columnas "Código" y "Existencia Actual". ¿Es el Listado de Productos?');
    const h = rows[hIdx].map(norm);
    const cCod = h.indexOf('codigo'), cNom = h.indexOf('nombre'), cEx = h.indexOf('existencia actual'), cUnd = h.indexOf('und. ppal.');
    if (cNom < 0) throw new Error('El archivo de Valery no tiene la columna "Nombre".');

    const items = new Map();
    const dups = [];
    for (const r of rows.slice(hIdx + 1)) {
      const nombre = r[cNom];
      if (nombre === null || String(nombre).trim() === '') continue; // fila de totales y filas vacías
      const code = String(r[cCod] ?? '').trim().replace(/\.0+$/, '');
      if (!/^\d+$/.test(code)) continue;
      const ex = Number(r[cEx]);
      if (!Number.isFinite(ex)) continue;
      if (items.has(code)) dups.push(code);
      items.set(code, { code, nombre: String(nombre).trim(), unidad: cUnd >= 0 ? r[cUnd] : null, existencia: ex });
    }
    if (dups.length) throw new Error(`Códigos repetidos en Valery: ${dups.join(', ')}`);
    if (items.size === 0) throw new Error('No se encontró ningún producto en el archivo de Valery.');
    return items;
  }

  /* ---------------- Cashea (XML directo) ---------------- */
  const CELL_RE = /(<c r="([A-Z]+)(\d+)"[^>]*?)(\/>|>([\s\S]*?)<\/c>)/g;

  function unescapeXml(s) {
    return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
      .replace(/&#(\d+);/g, (_, d) => String.fromCharCode(+d)).replace(/&amp;/g, '&');
  }

  function parseSharedStrings(xml) {
    if (!xml) return [];
    const out = [];
    const siRe = /<si>([\s\S]*?)<\/si>/g;
    let m;
    while ((m = siRe.exec(xml))) {
      const parts = [...m[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((x) => x[1]);
      out.push(unescapeXml(parts.join('')));
    }
    return out;
  }

  function readCells(xml, shared, minRow) {
    const rows = new Map();
    for (const m of xml.matchAll(CELL_RE)) {
      const col = m[2], row = +m[3], inner = m[5] || '';
      if (row < minRow || !['A', 'B', 'C'].includes(col)) continue;
      const v = inner.match(/<v>([\s\S]*?)<\/v>/);
      let val = null;
      if (v) val = / t="s"/.test(m[1]) ? shared[+v[1]] : unescapeXml(v[1]);
      else if (inner.includes('<is>')) val = unescapeXml(inner.replace(/<[^>]+>/g, ''));
      if (!rows.has(row)) rows.set(row, {});
      rows.get(row)[col] = val;
    }
    return rows;
  }

  async function loadCashea(JSZip, data) {
    let zip;
    try { zip = await JSZip.loadAsync(data); }
    catch (e) { throw new Error('La planilla de Cashea no se pudo abrir. Debe ser el .xlsx que descarga Cashea.'); }
    const sheets = Object.keys(zip.files).filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n));
    if (sheets.length !== 1) throw new Error(`La planilla de Cashea debería tener una sola hoja y tiene ${sheets.length}. ¿Es la "Planilla de precio, cuotas y stock"?`);
    const sheetPath = sheets[0];
    const xml = await zip.file(sheetPath).async('string');
    const ssFile = zip.file('xl/sharedStrings.xml');
    const shared = parseSharedStrings(ssFile ? await ssFile.async('string') : '');
    const header = readCells(xml, shared, 2).get(2) || {};
    if (norm(header.B) !== 'sku' || norm(header.C) !== 'stock')
      throw new Error('La planilla no tiene "SKU" en la columna B y "Stock" en la C. ¿Es la "Planilla de precio, cuotas y stock" de Cashea?');
    // Cashea graba la hora de generación en las entradas del zip: es la fecha real de la planilla,
    // aunque el archivo se renombre o se copie.
    const dates = Object.values(zip.files).map((f) => f.date && f.date.getTime()).filter(Boolean);
    const generatedAt = dates.length ? new Date(Math.min(...dates)) : null;
    return { zip, sheetPath, xml, generatedAt, rows: readCells(xml, shared, FIRST_DATA_ROW) };
  }

  /* ---------------- Cruce ---------------- */
  // opts.only: Set de códigos con movimiento (modo movimientos). Los demás quedan "Sin movimiento".
  // opts.sinBase: Map código → nombre de productos con movimiento pero sin existencia guardada.
  function compare(valery, cashea, opts = {}) {
    const only = opts.only || null, sinBase = opts.sinBase || new Map();
    const result = [];
    const byCode = new Map();
    for (const [row, cells] of [...cashea.rows.entries()].sort((a, b) => a[0] - b[0])) {
      const sku = cells.B;
      if (!sku) continue;
      const old = cells.C === null || cells.C === undefined || cells.C === '' ? null : Math.trunc(Number(cells.C));
      const code = valeryCode(sku);
      const rec = { fila: row, sku, titulo: cells.A || '', codigo: code, stockAnterior: old, existenciaValery: null, stockNuevo: null, accion: '', motivo: '' };
      if (code === null) { rec.accion = only ? 'Sin movimiento' : 'Sin tocar'; rec.motivo = 'SKU sin código de Valery (combo u otro)'; }
      else if (only && sinBase.has(code)) { rec.accion = 'Sin tocar'; rec.motivo = 'Producto sin existencia guardada: hacer un cierre con inventario completo'; }
      else if (only && !only.has(code)) { rec.accion = 'Sin movimiento'; rec.motivo = ''; }
      else if (!valery.has(code)) { rec.accion = 'Sin tocar'; rec.motivo = 'El código no existe en Valery'; }
      else if (!sameProduct(rec.titulo, valery.get(code).nombre)) {
        rec.accion = 'Sin tocar';
        rec.revisarSku = true;
        rec.nombreValery = valery.get(code).nombre;
        rec.motivo = `El código ${code} en Valery es "${valery.get(code).nombre}": revisar el SKU`;
      }
      else {
        const ex = valery.get(code).existencia;
        rec.existenciaValery = ex;
        rec.stockNuevo = targetStock(ex);
        rec.accion = rec.stockNuevo !== old ? 'Cambia' : 'Igual';
        rec.motivo = noteFor(ex);
      }
      if (code) { if (!byCode.has(code)) byCode.set(code, []); byCode.get(code).push(rec); }
      result.push(rec);
    }
    const duplicados = [...byCode.values()].filter((l) => l.length > 1).flat();
    const enCashea = new Set(byCode.keys());
    const noPublicados = [...valery.values()].filter((v) => !enCashea.has(v.code) && (!only || only.has(v.code)));
    const count = (a) => result.filter((r) => r.accion === a).length;
    return {
      filas: result,
      cambios: result.filter((r) => r.accion === 'Cambia'),
      aCero: result.filter((r) => r.accion === 'Cambia' && r.stockNuevo === 0),
      suben: result.filter((r) => r.accion === 'Cambia' && (r.stockAnterior === null || r.stockNuevo > r.stockAnterior)),
      bajan: result.filter((r) => r.accion === 'Cambia' && r.stockAnterior !== null && r.stockNuevo < r.stockAnterior),
      decimales: result.filter((r) => r.motivo.includes('decimal') || r.motivo.includes('negativo')),
      sinTocar: result.filter((r) => r.accion === 'Sin tocar'),
      sinMovimiento: result.filter((r) => r.accion === 'Sin movimiento'),
      revisarSku: result.filter((r) => r.revisarSku),
      duplicados,
      noPublicados,
      resumen: { total: result.length, cambia: count('Cambia'), igual: count('Igual'), sinTocar: count('Sin tocar'),
                 sinMovimiento: count('Sin movimiento'), valery: only ? only.size : valery.size, modo: only ? 'movimientos' : 'completo' },
    };
  }

  /* ---------------- Reportes del día ---------------- */
  const num = (v) => {
    if (typeof v === 'number') return v;
    const t = String(v ?? '').trim();
    if (!t) return NaN;
    // "1.234,56" (Venezuela) → 1234.56 ; "2.6" → 2.6
    return /,\d+$/.test(t) ? Number(t.replace(/\./g, '').replace(',', '.')) : Number(t);
  };

  // ¿Qué reporte de Valery es este .xls? 'listado' | 'ventas' | null
  function detectXls(XLSX, data) {
    const wb = XLSX.read(data, { type: 'array' });
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: null }).slice(0, 15);
    const has = (t) => rows.some((r) => r.some((c) => norm(c) === t));
    if (has('codigo') && has('existencia actual')) return 'listado';
    if (has('codigo') && rows.some((r) => r.some((c) => norm(c).startsWith('cant')))) return 'ventas';
    return null;
  }

  // Valery: Reportes → Ventas → Relación de ventas → Relación de ventas por producto.
  // Columnas: Código, Nombre, …, "Cant (Und. Principal)", … La última fila es de totales (sin nombre).
  function parseVentas(XLSX, data) {
    const wb = XLSX.read(data, { type: 'array' });
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: null });
    const hIdx = rows.findIndex((r) => r.some((c) => norm(c) === 'codigo') && r.some((c) => norm(c).startsWith('cant')));
    if (hIdx < 0) throw new Error('No tiene las columnas "Código" y "Cant". ¿Es la Relación de Ventas por Producto?');
    const h = rows[hIdx].map(norm);
    const cCod = h.indexOf('codigo'), cNom = h.indexOf('nombre');
    const cCant = h.findIndex((c) => c.startsWith('cant'));
    const lineas = new Map();
    for (const r of rows.slice(hIdx + 1)) {
      const nombre = cNom >= 0 ? r[cNom] : 'x';
      if (nombre === null || String(nombre).trim() === '') continue;      // fila de totales
      const code = String(r[cCod] ?? '').trim().replace(/\.0+$/, '');
      if (!/^\d+$/.test(code)) continue;
      const cant = num(r[cCant]);
      if (!Number.isFinite(cant)) continue;
      const prev = lineas.get(code);
      lineas.set(code, { code, nombre: String(nombre).trim(), cant: round6((prev ? prev.cant : 0) + cant) });
    }
    if (!lineas.size) throw new Error('El reporte de ventas no tiene productos.');
    const list = [...lineas.values()];
    return { tipo: 'ventas', lineas: list, unidades: round6(list.reduce((a, l) => a + l.cant, 0)) };
  }

  /*
   * Valery: Reimprimir → Compras → Facturas de compras ("Recepción de factura de compra"), en PDF.
   * Se lee por posición: el texto del PDF se monta (nombres de dos líneas, "440V6,00"), así que se
   * ubica cada código a la izquierda y la cantidad alineada bajo la columna "Ord.".
   * Validación: cantidad × costo = subtotal en cada línea, y la suma de subtotales = "Total Lineas".
   */
  async function parseCompra(pdfjs, data) {
    let doc;
    try { doc = await pdfjs.getDocument({ data: new Uint8Array(data), isEvalSupported: false }).promise; }
    catch (e) { throw new Error('No se pudo abrir el PDF.'); }
    const items = [];
    for (let p = 1; p <= doc.numPages; p++) {
      const tc = await (await doc.getPage(p)).getTextContent();
      for (const it of tc.items) {
        const s = String(it.str || '').trim();
        if (s) items.push({ s, p, x: it.transform[4], y: it.transform[5], r: it.transform[4] + it.width });
      }
    }
    const find = (txt, p) => items.find((i) => (p === undefined || i.p === p) && norm(i.s) === norm(txt));
    if (!items.some((i) => norm(i.s) === 'factura de compra'))
      throw new Error('Este PDF no es una "Factura de compra" de Valery.');

    // Encabezado del documento
    const sameRow = (a, b) => a.p === b.p && Math.abs(a.y - b.y) < 3;
    const rightOf = (label) => { const l = find(label); return l ? items.filter((i) => sameRow(i, l) && i.x > l.x).sort((a, b) => a.x - b.x) : []; };
    const provRow = rightOf('Proveedor:');
    const rif = (provRow.find((i) => /^[JVGEP]-?\d/i.test(i.s)) || {}).s || '';
    const proveedor = (provRow.find((i) => !/^[JVGEP]-?\d/i.test(i.s)) || {}).s || '';
    const fc = find('FACTURA DE COMPRA');
    const numero = fc ? ((items.filter((i) => i.p === fc.p && i.y < fc.y && fc.y - i.y < 20 && /^[\w-]+$/.test(i.s))
      .sort((a, b) => b.y - a.y)[0]) || {}).s || '' : '';
    const fecha = (rightOf('Fecha Emisión:')[0] || {}).s || '';

    const lineas = [];
    let totalLineas = null;
    for (let p = 1; p <= doc.numPages; p++) {
      const hOrd = find('Ord.', p), hNom = find('Nombre', p);
      if (!hOrd || !hNom) continue;
      // Columnas numéricas, alineadas a la derecha: cada número va a la columna cuyo encabezado
      // termina más cerca de donde termina el número.
      const heads = [['ord', 'Ord.'], ['bonif', 'Bonif.'], ['costo', 'Costo Lista'], ['d1', '% Dcto 1'], ['d2', '% Dcto 2'],
        ['sub', 'Sub Total'], ['imp', 'Impuesto'], ['total', 'Total']]
        .map(([k, t]) => ({ k, h: items.find((i) => i.p === p && norm(i.s) === norm(t) && Math.abs(i.y - hOrd.y) < 3) }))
        .filter((c) => c.h);
      const colOf = (i) => heads.reduce((best, c) => (!best || Math.abs(c.h.r - i.r) < Math.abs(best.h.r - i.r) ? c : best), null).k;
      const tl = find('Total Lineas', p);
      if (tl) { const v = items.find((i) => sameRow(i, tl) && i.x > tl.x && /\d/.test(i.s)); if (v) totalLineas = num(v.s); }
      const bottom = tl ? tl.y : -Infinity;
      const isNum = (i) => /^[\d.]+,\d+$/.test(i.s);
      const codes = items.filter((i) => i.p === p && i.y < hOrd.y - 3 && i.y > bottom && i.r < hNom.x - 2 && /^\d+$/.test(i.s))
        .sort((a, b) => b.y - a.y);
      codes.forEach((c, k) => {
        const next = codes[k + 1];
        const top = c.y + 4, low = next ? next.y + 4 : bottom;          // franja de la línea
        const band = items.filter((i) => i.p === p && i.y <= top && i.y > low && i !== c);
        const val = (key) => {
          const it = band.filter((i) => isNum(i) && i.x > hNom.x + 20 && colOf(i) === key)
            .sort((a, b) => Math.abs(a.y - c.y) - Math.abs(b.y - c.y))[0];
          return it ? num(it.s) : NaN;
        };
        const nombre = band.filter((i) => i.x >= hNom.x - 3 && i.x < hOrd.x - 4 && !isNum(i))
          .sort((a, b) => b.y - a.y || a.x - b.x).map((i) => i.s).join(' ');
        const cant = val('ord'), bonif = val('bonif'), costo = val('costo'), subtotal = val('sub');
        const d1 = val('d1') || 0, d2 = val('d2') || 0;
        const esperado = cant * costo * (1 - d1 / 100) * (1 - d2 / 100);
        const cuadra = [cant, costo, subtotal].every(Number.isFinite) && Math.abs(esperado - subtotal) <= Math.max(0.05, subtotal * 0.0005);
        lineas.push({ code: c.s, nombre, cant: Number.isFinite(cant) ? round6(cant + (Number.isFinite(bonif) ? bonif : 0)) : NaN,
                      ordenada: cant, bonif: Number.isFinite(bonif) ? bonif : 0, costo, subtotal, cuadra });
      });
    }
    if (!lineas.length) throw new Error('No se encontraron productos en la factura de compra.');
    const sumSub = round6(lineas.reduce((a, l) => a + (Number.isFinite(l.subtotal) ? l.subtotal : 0), 0));
    const cuadraTotal = totalLineas !== null && Math.abs(sumSub - totalLineas) <= 0.05;
    return { tipo: 'compra', proveedor, rif, numero, fecha, lineas, totalLineas, sumaSubtotales: sumSub, cuadraTotal,
             unidades: round6(lineas.filter((l) => Number.isFinite(l.cant)).reduce((a, l) => a + l.cant, 0)) };
  }

  /*
   * Aplica ventas y compras sobre la última existencia guardada (prev: { código: existencia }).
   * Devuelve un Valery "virtual" con todos los códigos guardados, solo cambiando los que tienen movimiento.
   * Las líneas de compra que no cuadran (cantidad × costo ≠ subtotal) no se aplican.
   */
  function applyMovements(prev, docs) {
    const mov = new Map();
    const get = (code, nombre) => {
      if (!mov.has(code)) mov.set(code, { code, nombre: nombre || '', ventas: 0, compras: 0 });
      const m = mov.get(code); if (nombre && !m.nombre) m.nombre = nombre; return m;
    };
    for (const d of docs) {
      for (const l of d.lineas) {
        if (d.tipo === 'ventas') get(l.code, l.nombre).ventas = round6(get(l.code).ventas + l.cant);
        else if (d.tipo === 'compra' && l.cuadra && Number.isFinite(l.cant)) get(l.code, l.nombre).compras = round6(get(l.code).compras + l.cant);
      }
    }
    const valery = new Map();
    for (const [code, ex] of Object.entries(prev)) valery.set(code, { code, nombre: '', existencia: ex });
    const detalle = [], sinBase = new Map();
    for (const m of mov.values()) {
      if (!(m.code in prev)) { sinBase.set(m.code, m.nombre); detalle.push({ ...m, antes: null, despues: null }); continue; }
      const antes = prev[m.code], despues = round6(antes - m.ventas + m.compras);
      valery.set(m.code, { code: m.code, nombre: m.nombre, existencia: despues });
      detalle.push({ ...m, antes, despues });
    }
    detalle.sort((a, b) => Number(a.code) - Number(b.code));
    return { valery, touched: new Set(mov.keys()), detalle, sinBase,
             totales: { ventas: round6(detalle.reduce((a, d) => a + d.ventas, 0)), compras: round6(detalle.reduce((a, d) => a + d.compras, 0)) } };
  }

  // Reportes que ya se aplicaron en el último cierre: se vuelve a generar la planilla con ese resultado,
  // sin descontar ni sumar otra vez (current ya los incluye).
  function replayMovements(current, docs) {
    const r = applyMovements({}, docs);             // solo para agrupar ventas/compras por código
    const valery = new Map(Object.entries(current).map(([code, ex]) => [code, { code, nombre: '', existencia: ex }]));
    const detalle = [], sinBase = new Map();
    for (const d of r.detalle) {
      if (!(d.code in current)) { sinBase.set(d.code, d.nombre); detalle.push({ ...d, antes: null, despues: null }); continue; }
      valery.set(d.code, { code: d.code, nombre: d.nombre, existencia: current[d.code] });
      detalle.push({ ...d, antes: round6(current[d.code] + d.ventas - d.compras), despues: current[d.code] });
    }
    return { valery, touched: new Set(r.detalle.map((d) => d.code)), detalle, sinBase, totales: r.totales, replay: true };
  }

  /* ---------------- Movimiento en Valery entre dos cierres ---------------- */
  const round6 = (x) => Math.round(x * 1e6) / 1e6;

  // Foto compacta de Valery: { código: existencia } (~10 KB para 600 productos).
  function snapshot(valery) {
    const out = {};
    for (const [code, v] of valery) out[code] = round6(v.existencia);
    return out;
  }

  // prev: foto anterior { código: existencia } (de snapshot()) o un Map de parseValery.
  function movement(prev, valery) {
    const before = prev instanceof Map ? snapshot(prev) : prev;
    const list = [];
    for (const [code, v] of valery) {
      const ahora = round6(v.existencia);
      if (!(code in before)) { list.push({ code, nombre: v.nombre, antes: null, ahora, delta: null, nuevo: true }); continue; }
      const delta = round6(ahora - before[code]);
      if (delta !== 0) list.push({ code, nombre: v.nombre, antes: before[code], ahora, delta, nuevo: false });
    }
    const quitados = Object.keys(before).filter((c) => !valery.has(c));
    const bajan = list.filter((m) => m.delta < 0), suben = list.filter((m) => m.delta > 0);
    const sum = (l) => round6(l.reduce((a, m) => a + Math.abs(m.delta), 0));
    list.sort((a, b) => (a.delta ?? 0) - (b.delta ?? 0));
    return {
      list, bajan, suben, nuevos: list.filter((m) => m.nuevo), quitados,
      unidadesBajan: sum(bajan), unidadesSuben: sum(suben),
      byCode: new Map(list.map((m) => [m.code, m])),
    };
  }

  // Stock con el que queda Cashea después de subir la planilla: { SKU: stock }.
  function casheaSnapshot(comparison) {
    const out = {};
    for (const r of comparison.filas) {
      const v = r.stockNuevo !== null ? r.stockNuevo : r.stockAnterior;
      if (v !== null && v !== undefined) out[r.sku] = v;
    }
    return out;
  }

  /*
   * Ventas por canal desde el cierre anterior.
   *  base.valery : { código: existencia } al cierre anterior
   *  base.cashea : { SKU: stock subido a Cashea } al cierre anterior
   * Cashea vendido   = stock subido ayer − stock que trae hoy la planilla (Cashea descuenta al vender).
   * Valery neto      = existencia ayer − existencia hoy (todas las ventas facturadas, menos entradas).
   * Otros canales    = Valery neto − Cashea vendido (tienda y web), cuando Valery bajó.
   */
  function dailySales(base, valery, comparison) {
    const byCode = new Map();
    const get = (code) => {
      if (!byCode.has(code)) {
        const v = valery.get(code);
        const antes = base.valery[code];
        byCode.set(code, { code, nombre: v ? v.nombre : '', antes: antes === undefined ? null : antes,
          ahora: v ? round6(v.existencia) : null, cashea: 0, skus: [], ajusteCashea: 0 });
      }
      return byCode.get(code);
    };
    const subidoSinBase = [];
    for (const r of comparison.filas) {
      if (!r.codigo || r.revisarSku) continue;
      const prev = base.cashea[r.sku];
      if (prev === undefined || r.stockAnterior === null) { if (prev === undefined) subidoSinBase.push(r.sku); continue; }
      const d = prev - r.stockAnterior;
      const rec = get(r.codigo);
      rec.skus.push(r.sku);
      if (d > 0) rec.cashea += d; else if (d < 0) rec.ajusteCashea += -d;
    }
    for (const [code] of valery) if (code in base.valery) get(code);

    const rows = [];
    for (const rec of byCode.values()) {
      if (rec.antes === null || rec.ahora === null) continue;
      const neto = round6(rec.antes - rec.ahora);          // > 0: bajó en Valery
      const salida = Math.max(0, neto), entrada = Math.max(0, -neto);
      const otros = round6(salida - rec.cashea);
      rec.neto = neto; rec.entrada = entrada;
      rec.otros = Math.max(0, otros);
      rec.faltaFacturar = rec.cashea > salida;              // Cashea vendió más de lo que bajó Valery
      if (neto !== 0 || rec.cashea || rec.ajusteCashea) rows.push(rec);
    }
    rows.sort((a, b) => (b.cashea + b.otros) - (a.cashea + a.otros) || b.entrada - a.entrada);
    const sum = (k) => round6(rows.reduce((a, r) => a + r[k], 0));
    const nuevos = [...valery.values()].filter((v) => !(v.code in base.valery));
    return {
      rows,
      byCode: new Map(rows.map((r) => [r.code, r])),
      totales: { cashea: sum('cashea'), otros: sum('otros'), entradas: sum('entrada'), ajusteCashea: sum('ajusteCashea') },
      productos: { cashea: rows.filter((r) => r.cashea > 0).length, otros: rows.filter((r) => r.otros > 0).length,
                   entradas: rows.filter((r) => r.entrada > 0).length },
      faltaFacturar: rows.filter((r) => r.faltaFacturar),
      ajustesCashea: rows.filter((r) => r.ajusteCashea > 0),
      nuevos,
      quitados: Object.keys(base.valery).filter((c) => !valery.has(c)),
    };
  }

  // Unidades que vendió Cashea por código desde el último stock subido (base: { SKU: stock }).
  function casheaSold(baseCashea, comparison) {
    const out = new Map();
    if (!baseCashea) return out;
    for (const r of comparison.filas) {
      if (!r.codigo || r.revisarSku || r.stockAnterior === null) continue;
      const prev = baseCashea[r.sku];
      if (prev === undefined) continue;
      const d = prev - r.stockAnterior;
      if (d > 0) out.set(r.codigo, (out.get(r.codigo) || 0) + d);
    }
    return out;
  }

  /* ---------------- Planilla de salida ---------------- */
  async function buildCasheaFile(cashea, comparison) {
    const newVals = new Map(comparison.filas.filter((r) => r.stockNuevo !== null).map((r) => [r.fila, r.stockNuevo]));
    const newXml = cashea.xml.replace(CELL_RE, (whole, open, col, row) => {
      if (col !== 'C' || !newVals.has(+row)) return whole;
      const attrs = open.replace(/\s+t="[^"]*"/, '');
      return `${attrs}><v>${newVals.get(+row)}</v></c>`;
    });
    cashea.zip.file(cashea.sheetPath, newXml);
    return cashea.zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }

  function buildReport(XLSX, comparison, mov, baseLabel, ventas, movimientos) {
    const toRows = (list) => list.map((r) => ({
      'SKU': r.sku, 'Título': r.titulo, 'Código Valery': r.codigo, 'Existencia Valery': r.existenciaValery,
      'Stock anterior': r.stockAnterior, 'Stock nuevo': r.stockNuevo, 'Acción': r.accion, 'Motivo': r.motivo, 'Fila': r.fila,
    }));
    const wb = XLSX.utils.book_new();
    const s = comparison.resumen;
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
      ['Productos en la planilla', s.total], ['Cambian', s.cambia], ['Quedan en 0', comparison.aCero.length],
      ['Iguales', s.igual], ['Sin tocar', s.sinTocar], ['Productos en Valery', s.valery],
      ['SKU que no coinciden con Valery', comparison.revisarSku.length],
      ['Publicaciones duplicadas', comparison.duplicados.length], ['En Valery pero no en Cashea', comparison.noPublicados.length],
    ]), 'Resumen');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(toRows(comparison.cambios)), 'Cambios');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(toRows(comparison.sinTocar)), 'Sin tocar');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(toRows(comparison.decimales)), 'Decimales y negativos');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(toRows(comparison.duplicados)), 'Duplicados');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(comparison.noPublicados.map((v) => ({
      'Código': v.code, 'Nombre': v.nombre, 'Existencia': v.existencia }))), 'No publicados en Cashea');
    if (movimientos) {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(movimientos.detalle.map((d) => ({
        'Código': d.code, 'Nombre': d.nombre, 'Existencia guardada': d.antes, 'Ventas': d.ventas, 'Compras': d.compras,
        'Existencia nueva': d.despues, 'Nota': d.antes === null ? 'Sin existencia guardada: hacer inventario completo' : '' }))), 'Movimientos');
    }
    if (ventas) {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(ventas.rows.map((r) => ({
        'Código': r.code, 'Nombre': r.nombre, [`Valery ${baseLabel || 'antes'}`]: r.antes, 'Valery ahora': r.ahora,
        'Vendido por Cashea': r.cashea, 'Vendido tienda/web': r.otros, 'Entrada de mercancía': r.entrada,
        'Nota': [r.faltaFacturar ? 'Cashea vendió más de lo que bajó Valery: ¿falta facturar?' : '',
                 r.ajusteCashea ? `Stock subido a mano en Cashea (+${r.ajusteCashea})` : ''].filter(Boolean).join('; ') }))), 'Ventas del día');
    }
    if (mov) {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(mov.list.map((m) => ({
        'Código': m.code, 'Nombre': m.nombre, [`Valery ${baseLabel || 'antes'}`]: m.antes, 'Valery ahora': m.ahora,
        'Movimiento': m.nuevo ? 'nuevo' : m.delta }))), 'Movimiento Valery');
    }
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(toRows(comparison.filas)), 'Todo');
    return XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
  }

  return { valeryCode, targetStock, sameProduct, snapshot, casheaSnapshot, movement, dailySales, parseValery, loadCashea, compare,
           buildCasheaFile, buildReport, detectXls, parseVentas, parseCompra, applyMovements, replayMovements, casheaSold };
});
