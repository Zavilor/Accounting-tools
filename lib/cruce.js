// Cruce de dos listas de comprobantes de compras: la planilla "Bis" (o un "Mis Comprobantes" de ARCA)
// contra el "Portal IVA Compras". Recibe los dos Excel como ArrayBuffer/Buffer y devuelve los dos archivos
// marcados + resumen. ExcelJS se usa SOLO para leer. Los archivos marcados se generan editando el XML original
// (ver marcar-xlsx.js), así que todo lo que no es la marca queda exactamente igual. Funciona en el navegador y en Node.
//
// Los encabezados se reconocen por sinónimos (no por posición ni por nombre exacto), así que sirve para
// distintos formatos de exportación. Cada campo tiene una lista de nombres aceptados, en orden de preferencia.

import { marcarXlsx } from './marcar-xlsx.js';

const TIPOS = { 1: 'Factura A', 2: 'ND A', 3: 'NC A', 6: 'Factura B', 7: 'ND B', 8: 'NC B', 11: 'Factura C', 12: 'ND C', 13: 'NC C', 51: 'Factura M', 63: 'Liquidación A', 81: 'Tique Factura A' };
const LETRA = { 1: 'A', 2: 'A', 3: 'A', 6: 'B', 7: 'B', 8: 'B', 11: 'C', 12: 'C', 13: 'C', 51: 'M', 63: 'A', 81: 'A' };

// Nombres aceptados por campo (se comparan normalizados: sin tildes, sin mayúsculas, sin puntos ni símbolos).
const CAMPOS = {
  fecha: ['Fecha de Emisión', 'Fecha'],
  tipo: ['Tipo de Comprobante', 'Tipo Fact', 'Tipo', 'TC', 'Fact', 'Tipo Cbte'],
  pv: ['Punto de Venta', 'Pto', 'PV'],
  numero: ['Número de Comprobante', 'Número Desde', 'Fact Nº', 'Fact N', 'Num', 'Número'],
  cuit: ['Nro. Doc. Vendedor', 'Nro. Doc. Emisor', 'CUIT', 'Cuit'],
  prov: ['Denominación Vendedor', 'Denominación Emisor', 'Prov', 'Proveedor'],
  total: ['Importe Total', 'Imp. Total', 'Total'],
};

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

// Clave de comparación: solo letras y números en minúscula. Así "Número", "NÃºmero" (codificación rota)
// y "numero" dan la misma clave, y "Nro. Doc." y "nro doc" también.
const clave = s => String(s ?? '').toLowerCase().replace(/[^a-z0-9]/g, '');

// "0012-00000257", "00012-00000257", "A00104-00003395", "9-813572.0" -> { pv, nro }
function parsearNumero(s) {
  const m = s.trim().match(/^[A-Za-z]*\s*(\d+)\s*-\s*([\d.]+?)(\.0)?$/);
  if (!m) return { pv: null, nro: null };
  return { pv: parseInt(m[1], 10), nro: parseInt(m[2].replace(/\./g, ''), 10) };
}

// Tipo de comprobante: acepta código ARCA (1, 6, 81…) o texto ("FC A", "Factura B", "NC.A", "A", "TKA ").
// Devuelve { cod, letra, txt }. cod es null si no lo reconoce (la fila igual se cruza por número).
function parsearTipo(raw) {
  const original = String(raw ?? '').trim();
  const t = original.toUpperCase().replace(/\./g, '').replace(/\s+/g, ' ').trim();
  if (!t) return { cod: null, letra: null, txt: '' };
  const n = Number(t);
  if (Number.isInteger(n) && TIPOS[n]) return { cod: n, letra: LETRA[n], txt: TIPOS[n] };
  const letra = (t.match(/([ABCM])$/) || [])[1] || null;
  let base = null;
  if (/^(NC|NCR|NOTA DE CRED)/.test(t)) base = 'NC';
  else if (/^(ND|NOTA DE DEB)/.test(t)) base = 'ND';
  else if (/^(TKA|TIQ)/.test(t)) base = 'TIQ';            // "TKA" se toma como tique factura (A). Confirmar.
  else if (/^(FC|FACTURA)/.test(t) || /^[ABCM]$/.test(t)) base = 'FC';
  const tabla = { FC: { A: 1, B: 6, C: 11 }, ND: { A: 2, B: 7, C: 12 }, NC: { A: 3, B: 8, C: 13 }, TIQ: { A: 81 } };
  const cod = base && letra && tabla[base][letra] ? tabla[base][letra] : null;
  return { cod, letra, txt: cod ? TIPOS[cod] : original };
}

// última columna con algún dato en toda la hoja (para no pisar notas sueltas a la derecha)
function ultimaColumnaUsada(ws) {
  let max = 0;
  ws.eachRow({ includeEmpty: false }, row => {
    row.eachCell({ includeEmpty: false }, (cell, c) => { if (valor(cell.value) !== null && texto(cell.value) !== '' && c > max) max = c; });
  });
  return max;
}

// ---------- reconocer encabezados ----------
// Busca en la fila de encabezados la columna de cada campo, probando los nombres en orden de preferencia.
function mapearEncabezados(enc) {
  const buscar = nombres => {
    for (const n of nombres) {
      const k = clave(n);
      for (let c = 1; c <= enc.cellCount; c++) if (clave(texto(enc.getCell(c).value)) === k) return c;
    }
    return null;
  };
  const col = {};
  for (const campo of Object.keys(CAMPOS)) col[campo] = buscar(CAMPOS[campo]);
  col.concepto = null;
  for (let c = 1; c <= enc.cellCount; c++) if (clave(texto(enc.getCell(c).value)).startsWith('concepto')) { col.concepto = c; break; }
  return col;
}

// Primera fila (dentro de las 15 primeras) que tiene los campos mínimos. null si no hay ninguna.
function buscarEncabezado(ws) {
  for (let r = 1; r <= Math.min(ws.rowCount, 15); r++) {
    const col = mapearEncabezados(ws.getRow(r));
    if (col.fecha && col.total && col.tipo && col.cuit && col.numero) return { fila: r, col };
  }
  return null;
}

// ---------- lectura ----------
// Lee una lista de comprobantes con cualquiera de los formatos reconocidos.
function leerLista(ws) {
  const enc = buscarEncabezado(ws);
  if (!enc) throw new Error('No encontré los encabezados. Necesito, como mínimo: fecha, tipo de comprobante, número, CUIT y total (o sus equivalentes).');
  const { fila: filaEnc, col } = enc;
  if (!col.prov) col.prov = col.numero + 1; // planillas sin encabezado para el proveedor: la columna siguiente al número
  // Si falta el total en una fila, se suman las columnas de importes entre las de identificación y el total.
  const ids = [col.fecha, col.tipo, col.pv, col.numero, col.cuit, col.prov, col.concepto].filter(Boolean);
  const primeraImporte = Math.max(...ids) + 1;
  const ultimaCol = Math.max(ultimaColumnaUsada(ws), col.total);

  const filas = [];
  for (let r = filaEnc + 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const f = valor(row.getCell(col.fecha).value);
    // fecha real, o texto con formato AAAA-MM-DD (exports); cualquier otra cosa son totales o notas: fuera
    const fecha = f instanceof Date ? f : (typeof f === 'string' && /^\d{4}-\d{2}-\d{2}/.test(f) ? new Date(f) : null);
    if (!fecha || isNaN(fecha)) continue;
    // Total: el valor de la celda (número o texto numérico). Solo si la celda está vacía o es un error,
    // se suman las columnas de importes entre las de identificación y el total.
    let total;
    if (texto(row.getCell(col.total).value) !== '') total = num(row.getCell(col.total).value);
    else {
      total = 0;
      for (let c = primeraImporte; c < col.total; c++) total += num(row.getCell(c).value);
    }
    const tipo = parsearTipo(texto(row.getCell(col.tipo).value));
    let pv, nro, numTxt;
    if (col.pv) {
      pv = parseInt(digitos(row.getCell(col.pv).value), 10);
      nro = parseInt(digitos(row.getCell(col.numero).value), 10);
      numTxt = `${pv}-${nro}`;
    } else {
      numTxt = texto(row.getCell(col.numero).value);
      ({ pv, nro } = parsearNumero(numTxt));
    }
    filas.push({
      fila: r, fecha, tipo: tipo.txt, cod: tipo.cod, letra: tipo.letra,
      num: numTxt, pv, nro, prov: texto(row.getCell(col.prov).value),
      cuit: digitos(row.getCell(col.cuit).value), cuitTxt: texto(row.getCell(col.cuit).value),
      total: r2(total), m: null, como: null,
    });
  }
  return { filas, filaEnc, ultimaCol };
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
  if (a.letra && b.letra && a.letra !== b.letra) o.push(lado === 'bis' ? `Tipo distinto (Portal ${b.tipo})` : `Tipo distinto (Bis ${a.tipo})`);
  const d = r2(b.total - a.total);
  if (d !== 0) o.push(`Importe distinto: Bis ${fmtPlata(a.total)} / Portal ${fmtPlata(b.total)} (dif ${d > 0 ? '+' : ''}${fmtPlata(d)})`);
  return o;
}

// Notas de crédito del Portal que anulan exactamente una factura del mismo CUIT que tampoco está en Bis.
function paresAnulados(B) {
  const notas = {};
  const sueltas = B.filter(b => !b.m);
  for (const nc of sueltas.filter(b => b.cod === 3 || b.cod === 8 || b.cod === 13)) {
    const f = sueltas.find(b => b !== nc && !notas[b.fila] && b.cuit === nc.cuit && b.total > 0 && iguales(b.total, -nc.total));
    if (f) { notas[f.fila] = `anulada por la NC de la fila ${nc.fila} (neto $0)`; notas[nc.fila] = `anula la factura de la fila ${f.fila} (neto $0)`; }
  }
  return notas;
}

// ---------- entrada principal ----------
// bufBis = la lista que se marca con "Cruce con Portal" (planilla Bis o Mis Comprobantes).
// bufPortal = el Portal IVA Compras, que es la referencia.
export async function cruzar(ExcelJS, JSZip, bufBis, bufPortal, { mes } = {}) {
  const wbA = new ExcelJS.Workbook(); await wbA.xlsx.load(bufBis);
  const wbB = new ExcelJS.Workbook(); await wbB.xlsx.load(bufPortal);
  const wsA = wbA.worksheets[0], wsB = wbB.worksheets[0];
  const bis = leerLista(wsA), portal = leerLista(wsB);
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
      detalle.portalSinBis.push({ fila: b.fila, fecha: fmtFecha(b.fecha), proveedor: b.prov, numero: `${b.pv}-${b.nro}`, tipo: b.tipo, total: b.total, nota: notas[b.fila] || '' });
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
    tiposSinReconocer: [...new Set([...A, ...B].filter(x => x.cod === null && x.tipo).map(x => x.tipo))],
  };
  return {
    resumen, detalle,
    pares: A.filter(a => a.m).map(a => [a.fila, a.m.fila, a.como]),
    bis: await marcarXlsx(JSZip, bufBis, marcasA, { colTexto: cA, encabezado: { fila: bis.filaEnc, texto: 'Cruce con Portal' } }),
    portal: await marcarXlsx(JSZip, bufPortal, marcasB, { colTexto: cB, encabezado: { fila: portal.filaEnc, texto: 'Cruce con Bis' } }),
  };
}

// ---------- detección de archivos (para la página) ----------
// Mira un archivo y dice si tiene columnas reconocibles y si es una planilla (encabezado "Total" simple,
// como la planilla Compras Bis) o un export de ARCA/Portal (encabezado tipo "Importe Total").
export async function infoArchivo(ExcelJS, buf) {
  const wb = new ExcelJS.Workbook(); await wb.xlsx.load(buf);
  const ws = wb.worksheets[0];
  const e = buscarEncabezado(ws);
  if (!e) return { ok: false, planilla: false };
  const planilla = clave(texto(ws.getRow(e.fila).getCell(e.col.total).value)) === 'total';
  return { ok: true, planilla };
}

// Decide qué archivo es la lista a marcar (Bis) y cuál es el Portal (referencia).
// 1) Si uno es planilla y el otro no, la planilla es Bis.
// 2) Si no, el que tiene "portal" en el nombre es el Portal; el otro es Bis.
// Devuelve { bis, portal } con los mismos objetos que recibe. Lanza error si no se puede decidir.
export function asignarLados(archivos) {
  const lados = {};
  if (archivos.length === 2 && archivos[0].planilla !== archivos[1].planilla) {
    const [a, b] = archivos;
    return a.planilla ? { bis: a, portal: b } : { bis: b, portal: a };
  }
  for (const f of archivos) {
    const lado = !f.planilla && /portal/i.test(f.nombre) ? 'portal' : 'bis';
    if (lados[lado]) throw new Error(`No pude distinguir cuál es cuál: ${lados[lado].nombre} y ${f.nombre} parecen del mismo tipo.`);
    lados[lado] = f;
  }
  return lados;
}
