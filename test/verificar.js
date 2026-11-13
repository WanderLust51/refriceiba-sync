#!/usr/bin/env node
// Uso: node test/verificar.js VALERY.xls PLANILLA_CASHEA.xlsx [SALIDA.xlsx]
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const JSZip = require('jszip');
const S = require(path.join(__dirname, '..', 'js', 'sync.js'));

(async () => {
  const [valeryPath, casheaPath, outPath] = process.argv.slice(2);
  if (!valeryPath || !casheaPath) {
    console.error('Uso: node test/verificar.js VALERY.xls PLANILLA_CASHEA.xlsx [SALIDA.xlsx]');
    process.exit(1);
  }
  const valery = S.parseValery(XLSX, fs.readFileSync(valeryPath));
  const cashea = await S.loadCashea(JSZip, fs.readFileSync(casheaPath));
  const cmp = S.compare(valery, cashea);
  console.log('Resumen:', cmp.resumen);
  console.log('Quedan en 0:', cmp.aCero.length);
  console.log('SKU a revisar:', cmp.revisarSku.map((r) => `${r.sku} (${r.titulo} ≠ ${r.nombreValery})`));
  console.log('Duplicados:', cmp.duplicados.map((r) => r.sku));
  if (outPath) {
    fs.writeFileSync(outPath, await S.buildCasheaFile(cashea, cmp));
    console.log('Planilla escrita en', outPath);
  }
})().catch((e) => { console.error('ERROR:', e.message); process.exit(1); });
