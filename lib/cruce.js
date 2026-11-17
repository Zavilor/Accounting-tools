// Cruce Compras Bis vs Portal IVA Compras (ARCA).
// Recibe los dos Excel como ArrayBuffer/Buffer y devuelve los dos archivos marcados + resumen.
// ExcelJS se usa SOLO para leer. Los archivos marcados se generan editando el XML original (ver marcar-xlsx.js),
// así que todo lo que no es la marca queda exactamente igual. Funciona igual en el navegador y en Node.

import { marcarXlsx } from './marcar-xlsx.js';

const TIPOS = { 1: 'Factura A', 2: 'ND A', 3: 'NC A', 6: 'Factura B', 11: 'Factura C', 51: 'Factura M', 63: 'Liquidación A', 81: 'Tique Factura A' };
const LETRA = { 1: 'A', 2: 'A', 3: 'A', 6: 'B', 11: 'C', 51: 'M', 63: 'A', 81: 'A' };

// ---------- helpers ----------
function valor(v) {
  if (v === null || v === undefined) return null;
  if (v instanceof Date) return v;
  if (typeof v === 'object') {
    if ('result' in v) return valor(v.result);           // fórmula
    if (v.richText) return v.richText.map(t => t.text).join('');
    if ('text' in v) return v.text;                       // hipervínculo
    if (v.error) return null;
  }
  return v;
}
const texto = v => { const x = valor(v); return x === null ? '' : String(x).trim(); };
const digitos = v => texto(v).replace(/\D/g, '');
const num = v => { const x = valor(v); return typeof x === 'number' ? x : (x ? Number(String(x).replace(',', '.')) || 0 : 0); };
const r2 = x => Math.round(x * 100) / 100;
const iguales = (a, b) => Math.abs(a - b) < 0.005;
const fmtCuit = c => (c.length === 11 ? `${c.slice(0, 2)}-${c.slice(2, 10)}-${c.slice(10)}` : c);
const fmtFecha = d => `${String(d.getUTCDate()).padStart(2, '0')}/${String(d.getUTCMonth() + 1).padStart(2, '0')}/${d.getUTCFullYear()}`;
const fmtPlata = x => x.toLocaleString('es-AR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const mismoDia = (a, b) => a.getUTCFullYear() === b.getUTCFullYear() && a.getUTCMonth() === b.getUTCMonth() && a.getUTCDate() === b.getUTCDate();

// "0012-00000257", "00012-00000257", "A00104-00003395", "9-813572.0" -> { pv, nro }
function parsearNumero(s) {
  const m = s.trim().match(/^[A-Za-z]*\s*(\d+)\s*-\s*([\d.]+?)(\.0)?$/);
  if (!m) return { pv: null, nro: null };
  return { pv: parseInt(m[1], 10), nro: parseInt(m[2].replace(/\./g, ''), 10) };
}

// última columna con algún dato en toda la hoja (para no pisar notas sueltas a la derecha)
function ultimaColumnaUsada(ws) {
  let max = 0;
  ws.eachRow({ includeEmpty: false }, row => {
    row.eachCell({ includeEmpty: false }, (cell, c) => { if (valor(cell.value) !== null && texto(cell.value) !== '' && c > max) max = c; });
  });
  return max;
}

function buscarColumna(fila, ...nombres) {
  for (let c = 1; c <= fila.cellCount; c++) {
    const t = texto(fila.getCell(c).value).toLowerCase();
    if (nombres.some(n => t === n.toLowerCase() || t.startsWith(n.toLowerCase()))) return c;
  }
  return null;
}

// ---------- lectura ----------
function leerBis(ws) {
  let filaEnc = null;
  for (let r = 1; r <= Math.min(ws.rowCount, 15); r++) {
    if (texto(ws.getRow(r).getCell(1).value).toLowerCase() === 'fecha') { filaEnc = r; break; }
  }
  if (!filaEnc) throw new Error('Compras Bis: no encontré la fila de encabezados (la que empieza con "Fecha").');
  const enc = ws.getRow(filaEnc);
  const col = {
    fecha: 1,
    tipo: buscarColumna(enc, 'Tipo Fact'),
    num: buscarColumna(enc, 'Fact Nº', 'Fact N'),
    cuit: buscarColumna(enc, 'CUIT'),
    concepto: buscarColumna(enc, 'Concepto'),
    total: buscarColumna(enc, 'Total'),
  };
  for (const [k, v] of Object.entries(col)) if (!v) throw new Error(`Compras Bis: falta la columna "${k}".`);
  col.prov = col.num + 1; // la columna del proveedor no tiene un encabezado fijo
  const ultimaCol = Math.max(ultimaColumnaUsada(ws), col.total);

  const filas = [];
  for (let r = filaEnc + 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const fecha = valor(row.getCell(col.fecha).value);
    if (!(fecha instanceof Date)) continue;
    let total = valor(row.getCell(col.total).value);
    if (typeof total !== 'number') { // sin valor calculado: sumo las columnas de importes
      total = 0;
      for (let c = col.concepto + 1; c < col.total; c++) total += num(row.getCell(c).value);
    }
    const n = texto(row.getCell(col.num).value);
    filas.push({
      fila: r, fecha, tipo: texto(row.getCell(col.tipo).value), num: n, prov: texto(row.getCell(col.prov).value),
      cuit: digitos(row.getCell(col.cuit).value), cuitTxt: texto(row.getCell(col.cuit).value),
      concepto: texto(row.getCell(col.concepto).value), total: r2(total), ...parsearNumero(n), m: null, como: null,
    });
  }
  return { filas, filaEnc, ultimaCol };
}

function leerPortal(ws) {
  const enc = ws.getRow(1);
  const col = {
    fecha: buscarColumna(enc, 'Fecha de Emisión', 'Fecha'),
    tipo: buscarColumna(enc, 'Tipo de Comprobante'),
    pv: buscarColumna(enc, 'Punto de Venta'),
    nro: buscarColumna(enc, 'Número de Comprobante', 'Número Desde'),
    cuit: buscarColumna(enc, 'Nro. Doc. Vendedor', 'Nro. Doc. Emisor'),
    prov: buscarColumna(enc, 'Denominación Vendedor', 'Denominación Emisor'),
    total: buscarColumna(enc, 'Importe Total'),
  };
  for (const [k, v] of Object.entries(col)) if (!v) throw new Error(`Portal IVA: falta la columna "${k}".`);
  const filas = [];
  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const fecha = valor(row.getCell(col.fecha).value);
    if (!fecha) continue;
    filas.push({
      fila: r, fecha: fecha instanceof Date ? fecha : new Date(fecha), tipo: parseInt(texto(row.getCell(col.tipo).value), 10),
      pv: parseInt(digitos(row.getCell(col.pv).value), 10), nro: parseInt(digitos(row.getCell(col.nro).value), 10),
      cuit: digitos(row.getCell(col.cuit).value), prov: texto(row.getCell(col.prov).value),
      total: r2(num(row.getCell(col.total).value)), m: null,
    });
  }
  return { filas, ultimaCol: ultimaColumnaUsada(ws) };
}

// "Mis Comprobantes" de ARCA (comprobantes recibidos): mismo cruce que Bis, pero con su propio formato.
// Sus encabezados pueden venir con la codificación rota ("EmisiÃ³n"), por eso se buscan por prefijo.
function leerArca(ws) {
  const enc = ws.getRow(1);
  const col = {
    fecha: buscarColumna(enc, 'Fecha de Emisi'),
    tipo: buscarColumna(enc, 'Tipo de Comprobante'),
    pv: buscarColumna(enc, 'Punto de Venta'),
    nro: buscarColumna(enc, 'Número Desde', 'NÃºmero Desde', 'Numero Desde'),
    cuit: buscarColumna(enc, 'Nro. Doc. Emisor'),
    prov: buscarColumna(enc, 'Denominaci'),
    total: buscarColumna(enc, 'Imp. Total'),
  };
  for (const [k, v] of Object.entries(col)) if (!v) throw new Error(`Mis Comprobantes: falta la columna "${k}".`);
  const filas = [];
  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const fecha = valor(row.getCell(col.fecha).value);
    if (!fecha) continue;
    const cod = parseInt(digitos(row.getCell(col.tipo).value), 10);
    const pv = parseInt(digitos(row.getCell(col.pv).value), 10);
    const nro = parseInt(digitos(row.getCell(col.nro).value), 10);
    filas.push({
      fila: r, fecha: fecha instanceof Date ? fecha : new Date(fecha), tipo: TIPOS[cod] || String(cod),
      num: `${pv}-${nro}`, prov: texto(row.getCell(col.prov).value),
      cuit: digitos(row.getCell(col.cuit).value), cuitTxt: texto(row.getCell(col.cuit).value),
      concepto: '', total: r2(num(row.getCell(col.total).value)), pv, nro, m: null, como: null,
    });
  }
  return { filas, filaEnc: 1, ultimaCol: Math.max(ultimaColumnaUsada(ws), col.total) };
}

// ---------- cruce ----------
function cruzarFilas(A, B) {
  const unir = (a, b, como) => { a.m = b; b.m = a; a.como = como; };
  const libres = () => B.filter(b => !b.m);
  // 1) CUIT + punto de venta + número
  for (const a of A) {
    const b = libres().find(b => a.cuit === b.cuit && a.pv === b.pv && a.nro === b.nro);
    if (b) unir(a, b, 'CUIT + PV + Número');
  }
  // 2) punto de venta + número + total exacto (CUIT distinto)
  for (const a of A.filter(a => !a.m)) {
    const c = libres().filter(b => a.pv === b.pv && a.nro === b.nro && iguales(a.total, b.total));
    if (c.length === 1) unir(a, c[0], 'PV + Número + Total');
  }
  // 3) CUIT + total exacto (número distinto o vacío)
  for (const a of A.filter(a => !a.m)) {
    const c = libres().filter(b => a.cuit === b.cuit && iguales(a.total, b.total));
    if (c.length === 1) unir(a, c[0], 'CUIT + Total');
  }
}

function diferencias(a, b, lado) {
  const o = [];
  const otro = lado === 'bis' ? 'Portal' : 'Bis';
  if (a.cuit !== b.cuit) o.push(lado === 'bis' ? `CUIT distinto (Portal ${fmtCuit(b.cuit)} ${b.prov})` : `CUIT distinto (Bis ${a.cuitTxt} ${a.prov})`);
  if (a.pv !== b.pv || a.nro !== b.nro) o.push(lado === 'bis' ? `Número distinto (Portal ${b.pv}-${b.nro})` : `Número distinto (Bis ${a.num || 'vacío'})`);
  if (!mismoDia(a.fecha, b.fecha)) o.push(`Fecha distinta (${otro} ${fmtFecha(lado === 'bis' ? b.fecha : a.fecha)})`);
  if (a.tipo.slice(-1) !== LETRA[b.tipo]) o.push(lado === 'bis' ? `Tipo distinto (Portal ${TIPOS[b.tipo] || b.tipo})` : `Tipo distinto (Bis ${a.tipo})`);
  const d = r2(b.total - a.total);
  if (d !== 0) o.push(`Importe distinto: Bis ${fmtPlata(a.total)} / Portal ${fmtPlata(b.total)} (dif ${d > 0 ? '+' : ''}${fmtPlata(d)})`);
  return o;
}

// Notas de crédito del Portal que anulan exactamente una factura del mismo CUIT que tampoco está en Bis.
function paresAnulados(B) {
  const notas = {};
  const sueltas = B.filter(b => !b.m);
  for (const nc of sueltas.filter(b => b.tipo === 3 || b.tipo === 8 || b.tipo === 13)) {
    const f = sueltas.find(b => b !== nc && !notas[b.fila] && b.cuit === nc.cuit && b.total > 0 && iguales(b.total, -nc.total));
    if (f) { notas[f.fila] = `anulada por la NC de la fila ${nc.fila} (neto $0)`; notas[nc.fila] = `anula la factura de la fila ${f.fila} (neto $0)`; }
  }
  return notas;
}

// ---------- entrada principal ----------
export async function cruzar(ExcelJS, JSZip, bufBis, bufPortal, { mes } = {}) {
  const wbA = new ExcelJS.Workbook(); await wbA.xlsx.load(bufBis);
  const wbB = new ExcelJS.Workbook(); await wbB.xlsx.load(bufPortal);
  const wsA = wbA.worksheets[0], wsB = wbB.worksheets[0];
  const bis = (await tipoArchivo(ExcelJS, bufBis)) === 'mis' ? leerArca(wsA) : leerBis(wsA);
  const portal = leerPortal(wsB);
  const A = bis.filas, B = portal.filas;
  cruzarFilas(A, B);

  // período: el mes que más se repite en el Portal (salvo que se indique)
  if (!mes) {
    const cuenta = {};
    for (const b of B) { const k = `${b.fecha.getUTCFullYear()}-${b.fecha.getUTCMonth()}`; cuenta[k] = (cuenta[k] || 0) + 1; }
    mes = Object.entries(cuenta).sort((x, y) => y[1] - x[1])[0][0];
  }
  const [anio, m0] = mes.split('-').map(Number);
  const enPeriodo = d => d.getUTCFullYear() === anio && d.getUTCMonth() === m0;
  const periodo = `${String(m0 + 1).padStart(2, '0')}/${anio}`;

  const detalle = { bisSinPortal: [], portalSinBis: [], conDiferencias: [] };

  // --- marcar Bis
  const cA = bis.ultimaCol + 1;
  const marcasA = new Map();
  for (const a of A) {
    let t;
    if (!a.m) {
      const dup = A.find(x => x !== a && x.m && x.cuit === a.cuit && x.pv === a.pv && x.nro === a.nro);
      if (dup) t = `NO ESTÁ EN PORTAL — DUPLICADO de la fila ${dup.fila}`;
      else if (!enPeriodo(a.fecha)) t = `NO ESTÁ EN PORTAL — fecha fuera del período ${periodo}`;
      else t = 'NO ESTÁ EN PORTAL';
      marcasA.set(a.fila, { color: 'rojo' });
      detalle.bisSinPortal.push({ fila: a.fila, fecha: fmtFecha(a.fecha), proveedor: a.prov, numero: a.num, total: a.total, nota: t.replace(/^NO ESTÁ EN PORTAL( — )?/, '') });
    } else {
      const o = diferencias(a, a.m, 'bis');
      t = `OK — fila ${a.m.fila} del Portal` + (o.length ? ' | ' + o.join(' | ') : '');
      if (o.length) {
        marcasA.set(a.fila, { color: 'amarillo' });
        detalle.conDiferencias.push({ filaBis: a.fila, filaPortal: a.m.fila, proveedor: a.prov, nota: o.join(' | ') });
      }
    }
    marcasA.set(a.fila, { color: marcasA.get(a.fila)?.color || null, texto: t });
  }

  // --- marcar Portal
  const notas = paresAnulados(B);
  const cB = portal.ultimaCol + 1;
  const marcasB = new Map();
  for (const b of B) {
    let t;
    if (!b.m) {
      t = 'NO ESTÁ EN BIS' + (notas[b.fila] ? ` — ${notas[b.fila]}` : '');
      marcasB.set(b.fila, { color: 'rojo' });
      detalle.portalSinBis.push({ fila: b.fila, fecha: fmtFecha(b.fecha), proveedor: b.prov, numero: `${b.pv}-${b.nro}`, tipo: TIPOS[b.tipo] || String(b.tipo), total: b.total, nota: notas[b.fila] || '' });
    } else {
      const o = diferencias(b.m, b, 'portal');
      t = `OK — fila ${b.m.fila} de Bis` + (o.length ? ' | ' + o.join(' | ') : '');
      if (o.length) marcasB.set(b.fila, { color: 'amarillo' });
    }
    marcasB.set(b.fila, { color: marcasB.get(b.fila)?.color || null, texto: t });
  }

  const resumen = {
    periodo,
    totalBis: A.length,
    totalPortal: B.length,
    coincidencias: A.filter(a => a.m).length,
    bisSinPortal: detalle.bisSinPortal.length,
    bisSinPortalFueraDePeriodo: A.filter(a => !a.m && !enPeriodo(a.fecha)).length,
    portalSinBis: detalle.portalSinBis.length,
    conDiferencias: detalle.conDiferencias.length,
  };
  return {
    resumen, detalle,
    pares: A.filter(a => a.m).map(a => [a.fila, a.m.fila, a.como]),
    bis: await marcarXlsx(JSZip, bufBis, marcasA, { colTexto: cA, encabezado: { fila: bis.filaEnc, texto: 'Cruce con Portal' } }),
    portal: await marcarXlsx(JSZip, bufPortal, marcasB, { colTexto: cB, encabezado: { fila: 1, texto: 'Cruce con Bis' } }),
  };
}

// Qué tipo de archivo es, mirando los encabezados:
//  'portal' = Portal IVA Compras (fila 1, tiene "Importe Total")
//  'mis'    = Mis Comprobantes de ARCA (fila 1, tiene "Imp. Total")
//  'bis'    = planilla Compras Bis (encabezado "Fecha" + "Tipo Fact" en las primeras filas)
//  null     = no lo reconoce
export async function tipoArchivo(ExcelJS, buf) {
  const wb = new ExcelJS.Workbook(); await wb.xlsx.load(buf);
  const ws = wb.worksheets[0];
  const enc = ws.getRow(1);
  const a1 = texto(enc.getCell(1).value).toLowerCase();
  if (a1.startsWith('fecha de emisi')) {
    const h = [];
    for (let c = 1; c <= enc.cellCount; c++) h.push(texto(enc.getCell(c).value).toLowerCase());
    if (h.some(x => x.startsWith('importe total'))) return 'portal';
    if (h.some(x => x.startsWith('imp. total'))) return 'mis';
  }
  for (let r = 1; r <= Math.min(ws.rowCount, 15); r++) {
    if (texto(ws.getRow(r).getCell(1).value).toLowerCase() === 'fecha') return 'bis';
  }
  return null;
}
