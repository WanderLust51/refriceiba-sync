/*
 * Lógica de sincronización Valery → Cashea.
 * Funciona igual en el navegador (usa window.XLSX y window.JSZip) y en Node (para las pruebas).
 *
 * Reglas:
 *  - SKU "REF-123" o "123" en Cashea = Código 123 en Valery.
 *  - Stock = parte entera de la existencia de Valery; negativo → 0.
 *  - Solo se modifica la columna Stock (C) de la planilla de Cashea; todo lo demás queda intacto.
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
    return { zip, sheetPath, xml, rows: readCells(xml, shared, FIRST_DATA_ROW) };
  }

  /* ---------------- Cruce ---------------- */
  function compare(valery, cashea) {
    const result = [];
    const byCode = new Map();
    for (const [row, cells] of [...cashea.rows.entries()].sort((a, b) => a[0] - b[0])) {
      const sku = cells.B;
      if (!sku) continue;
      const old = cells.C === null || cells.C === undefined || cells.C === '' ? null : Math.trunc(Number(cells.C));
      const code = valeryCode(sku);
      const rec = { fila: row, sku, titulo: cells.A || '', codigo: code, stockAnterior: old, existenciaValery: null, stockNuevo: null, accion: '', motivo: '' };
      if (code === null) { rec.accion = 'Sin tocar'; rec.motivo = 'SKU sin código de Valery (combo u otro)'; }
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
    const noPublicados = [...valery.values()].filter((v) => !enCashea.has(v.code));
    const count = (a) => result.filter((r) => r.accion === a).length;
    return {
      filas: result,
      cambios: result.filter((r) => r.accion === 'Cambia'),
      aCero: result.filter((r) => r.accion === 'Cambia' && r.stockNuevo === 0),
      decimales: result.filter((r) => r.motivo.includes('decimal') || r.motivo.includes('negativo')),
      sinTocar: result.filter((r) => r.accion === 'Sin tocar'),
      revisarSku: result.filter((r) => r.revisarSku),
      duplicados,
      noPublicados,
      resumen: { total: result.length, cambia: count('Cambia'), igual: count('Igual'), sinTocar: count('Sin tocar'), valery: valery.size },
    };
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

  function buildReport(XLSX, comparison) {
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
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(toRows(comparison.filas)), 'Todo');
    return XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
  }

  return { valeryCode, targetStock, sameProduct, parseValery, loadCashea, compare, buildCasheaFile, buildReport };
});
