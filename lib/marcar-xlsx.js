// Marca filas en un .xlsx tocando SOLO lo necesario del XML interno.
// No se reescribe el libro: fechas, fórmulas, valores calculados, formatos y todo lo demás quedan byte por byte
// como en el original. Cambia únicamente:
//   - styles.xml: se agregan 2 rellenos (rojo/amarillo) y las combinaciones de estilo que hagan falta (al final).
//   - la hoja: el atributo de estilo (s) de las celdas pintadas, la columna nueva con el detalle,
//     el ancho de esa columna y el rango <dimension>.

const COLORES = { rojo: 'FFFF9999', amarillo: 'FFFFE699' };

const letraACol = l => [...l].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
const colALetra = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
const escapar = t => String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const attr = (tag, nombre) => { const m = tag.match(new RegExp(`\\s${nombre}="([^"]*)"`)); return m ? m[1] : null; };
function setAttr(tag, nombre, valor) {
  if (new RegExp(`\\s${nombre}="[^"]*"`).test(tag)) return tag.replace(new RegExp(`\\s${nombre}="[^"]*"`), ` ${nombre}="${valor}"`);
  return `${tag} ${nombre}="${valor}"`; // tag viene sin el cierre (> o />)
}

async function rutaPrimeraHoja(zip) {
  const wb = await zip.file('xl/workbook.xml').async('string');
  const rid = attr(wb.match(/<sheet\b[^>]*>/)[0], 'r:id');
  const rels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  const rel = [...rels.matchAll(/<Relationship\b[^>]*>/g)].map(m => m[0]).find(t => attr(t, 'Id') === rid);
  const target = attr(rel, 'Target');
  return target.startsWith('/') ? target.slice(1) : 'xl/' + target;
}

// ---------- estilos ----------
function prepararEstilos(xml) {
  const fillsM = xml.match(/<fills\b[^>]*>([\s\S]*?)<\/fills>/);
  const cantFills = (fillsM[1].match(/<fill\b/g) || []).length;
  const nuevosFills = Object.values(COLORES)
    .map(c => `<fill><patternFill patternType="solid"><fgColor rgb="${c}"/><bgColor indexed="64"/></patternFill></fill>`).join('');
  xml = xml.replace(fillsM[0], fillsM[0].replace(/<\/fills>$/, nuevosFills + '</fills>').replace(/(<fills\b[^>]*\scount=")\d+/, `$1${cantFills + 2}`));
  const fillId = { rojo: cantFills, amarillo: cantFills + 1 };

  const xfsM = xml.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/);
  const xfs = xfsM[1].match(/<xf\b[^>]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>/g);
  const agregados = [];
  const cache = {};
  function estilo(sOriginal, color) {
    const clave = `${sOriginal}|${color}`;
    if (cache[clave] !== undefined) return cache[clave];
    const base = xfs[sOriginal] || xfs[0];
    const apertura = base.match(/^<xf\b[^>]*?(?=\/?>)/)[0];
    let nueva = setAttr(setAttr(apertura, 'fillId', fillId[color]), 'applyFill', '1');
    const xf = base.replace(/^<xf\b[^>]*?(?=\/?>)/, nueva);
    cache[clave] = xfs.length + agregados.length;
    agregados.push(xf);
    return cache[clave];
  }
  function cerrar() {
    return xml.replace(xfsM[0], xfsM[0]
      .replace(/<\/cellXfs>$/, agregados.join('') + '</cellXfs>')
      .replace(/(<cellXfs\b[^>]*\scount=")\d+/, `$1${xfs.length + agregados.length}`));
  }
  return { estilo, cerrar };
}

// ---------- columnas ----------
function ponerAncho(xml, col, ancho) {
  const nueva = `<col min="${col}" max="${col}" width="${ancho}" customWidth="1"/>`;
  const colsM = xml.match(/<cols>([\s\S]*?)<\/cols>/);
  if (!colsM) return xml.replace(/<sheetData\b/, `<cols>${nueva}</cols><sheetData`);
  const piezas = colsM[1].match(/<col\b[^>]*\/>/g) || [];
  const salida = [];
  let puesta = false;
  for (const p of piezas) {
    const min = +attr(p, 'min'), max = +attr(p, 'max');
    if (col < min && !puesta) { salida.push(nueva); puesta = true; }
    if (col >= min && col <= max) {
      // se parte el rango para cambiar solo esta columna; el resto queda igual
      if (min < col) salida.push(setAttr(setAttr(p, 'min', min), 'max', col - 1));
      salida.push(attr(p, 'style') ? `<col min="${col}" max="${col}" width="${ancho}" style="${attr(p, 'style')}" customWidth="1"/>` : nueva);
      if (max > col) salida.push(setAttr(setAttr(p, 'min', col + 1), 'max', max));
      puesta = true;
    } else salida.push(p);
  }
  if (!puesta) salida.push(nueva);
  return xml.replace(colsM[0], `<cols>${salida.join('')}</cols>`);
}

// ---------- filas ----------
function editarFila(filaXml, nroFila, { pintarHasta, color, colTexto, texto, sTexto }, estilos) {
  let apertura, interior;
  const auto = filaXml.match(/^(<row\b[^>]*?)\s*\/>$/);
  if (auto) { apertura = auto[1] + '>'; interior = ''; }
  else { const m = filaXml.match(/^(<row\b[^>]*>)([\s\S]*)<\/row>$/); apertura = m[1]; interior = m[2]; }
  apertura = apertura.replace(/\sspans="[^"]*"/, ''); // spans es opcional; se saca porque la fila ahora tiene una celda más

  const celdas = new Map();
  const resto = interior.replace(/<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g, c => {
    const r = attr(c.match(/^<c\b[^>]*?(?=\/?>)/)[0], 'r');
    if (!r) throw new Error('Formato de Excel no soportado (celdas sin referencia).');
    celdas.set(letraACol(r.replace(/\d+$/, '')), c);
    return '';
  });

  const conEstilo = (c, nroCol) => {
    const tag = c ?? `<c r="${colALetra(nroCol)}${nroFila}"/>`;
    const s = +(attr(tag.match(/^<c\b[^>]*?(?=\/?>)/)[0], 's') || 0);
    return tag.replace(/^<c\b[^>]*?(?=\/?>)/, t => setAttr(t, 's', estilos.estilo(s, color)));
  };
  if (color) for (let c = 1; c <= pintarHasta; c++) if (c !== colTexto) celdas.set(c, conEstilo(celdas.get(c), c));

  if (texto !== undefined) {
    const previa = celdas.get(colTexto);
    let s = previa ? attr(previa.match(/^<c\b[^>]*?(?=\/?>)/)[0], 's') : null;
    if (sTexto !== undefined) s = String(sTexto);
    if (color) s = String(estilos.estilo(+(s || 0), color));
    celdas.set(colTexto, `<c r="${colALetra(colTexto)}${nroFila}"${s ? ` s="${s}"` : ''} t="inlineStr"><is><t xml:space="preserve">${escapar(texto)}</t></is></c>`);
  }
  const ordenadas = [...celdas.keys()].sort((a, b) => a - b).map(k => celdas.get(k)).join('');
  return `${apertura}${ordenadas}${resto}</row>`;
}

/**
 * marcas: Map<nroFila, { color: 'rojo'|'amarillo'|null, texto: string }>
 * opciones: { colTexto, encabezado: { fila, texto }, anchoTexto }
 */
export async function marcarXlsx(JSZip, buf, marcas, { colTexto, encabezado, anchoTexto = 70 }) {
  const zip = await JSZip.loadAsync(buf);
  const rutaHoja = await rutaPrimeraHoja(zip);
  let hoja = await zip.file(rutaHoja).async('string');
  const estilos = prepararEstilos(await zip.file('xl/styles.xml').async('string'));

  // estilo del encabezado nuevo = el de la primera celda de la fila de encabezados
  const filaEncM = hoja.match(new RegExp(`<row\\b[^>]*\\sr="${encabezado.fila}"[^>]*?(?:\\/>|>[\\s\\S]*?<\\/row>)`));
  const primera = filaEncM && filaEncM[0].match(/<c\b[^>]*?(?=\/?>)/);
  const sEnc = primera ? attr(primera[0], 's') : null;

  const todas = new Map(marcas);
  todas.set(encabezado.fila, { color: null, texto: encabezado.texto, sTexto: sEnc ?? undefined });

  hoja = hoja.replace(/<row\b[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g, fila => {
    const nro = +attr(fila.match(/^<row\b[^>]*?(?=\/?>)/)[0], 'r');
    const m = todas.get(nro);
    if (!m) return fila;
    todas.delete(nro);
    return editarFila(fila, nro, { pintarHasta: colTexto, color: m.color, colTexto, texto: m.texto, sTexto: m.sTexto }, estilos);
  });
  if (todas.size) throw new Error(`No encontré las filas ${[...todas.keys()].join(', ')} en la hoja.`);

  hoja = hoja.replace(/<dimension ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"\/>/, (t, c1, f1, c2, f2) =>
    `<dimension ref="${c1}${f1}:${colALetra(Math.max(letraACol(c2), colTexto))}${f2}"/>`);
  hoja = ponerAncho(hoja, colTexto, anchoTexto);

  zip.file(rutaHoja, hoja, { createFolders: false });
  zip.file('xl/styles.xml', estilos.cerrar(), { createFolders: false });
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
