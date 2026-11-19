'use client';

import { useCallback, useRef, useState } from 'react';
import { cruzar, infoArchivo, asignarLados } from '../lib/cruce';

const plata = n => n.toLocaleString('es-AR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

function descargar(buffer, nombre) {
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = nombre; a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
const marcado = nombre => nombre.replace(/\.xlsx$/i, '') + '_MARCADO.xlsx';
const MOSTRAR_FOTO = false; // temporal: foto oculta

export default function Pagina() {
  const [archivos, setArchivos] = useState({ bis: null, portal: null });
  const [estado, setEstado] = useState('inicio'); // inicio | procesando | listo
  const [error, setError] = useState('');
  const [resultado, setResultado] = useState(null);
  const [arrastrando, setArrastrando] = useState(false);
  const input = useRef(null);

  const recibir = useCallback(async lista => {
    setError(''); setResultado(null); setEstado('inicio');
    const xlsx = [...lista].filter(f => /\.xlsx$/i.test(f.name));
    if (!xlsx.length) { setError('Soltá archivos .xlsx. Si vienen en .xls, abrilos en Excel y guardalos como .xlsx.'); return; }
    const ExcelJS = (await import('exceljs')).default;
    const leidos = [];
    for (const f of xlsx.slice(0, 2)) {
      const buf = await f.arrayBuffer();
      let info;
      try {
        info = await infoArchivo(ExcelJS, buf);
      } catch {
        setError(`No pude abrir ${f.name}. Revisá que sea un Excel válido.`); return;
      }
      if (!info.ok) { setError(`No reconozco las columnas de ${f.name}. Necesito fecha, tipo de comprobante, número, CUIT y total.`); return; }
      leidos.push({ nombre: f.name, buf, planilla: info.planilla });
    }
    let lados;
    try {
      lados = asignarLados(leidos);
    } catch (e) { setError(e.message); return; }
    // con dos archivos se reemplazan los dos; con uno, solo la ranura que le toca
    const nuevos = leidos.length === 2 ? { bis: null, portal: null } : { ...archivos };
    for (const [lado, f] of Object.entries(lados)) nuevos[lado] = { nombre: f.nombre, buf: f.buf };
    setArchivos(nuevos);
  }, [archivos]);

  async function procesar() {
    setEstado('procesando'); setError('');
    try {
      const ExcelJS = (await import('exceljs')).default;
      const JSZip = (await import('jszip')).default;
      const r = await cruzar(ExcelJS, JSZip, archivos.bis.buf.slice(0), archivos.portal.buf.slice(0));
      setResultado(r); setEstado('listo');
    } catch (e) {
      setError(e.message || 'El cruce falló. Revisá que los archivos tengan el formato de siempre.');
      setEstado('inicio');
    }
  }

  function reiniciar() {
    setArchivos({ bis: null, portal: null }); setResultado(null); setEstado('inicio'); setError('');
  }

  const listos = archivos.bis && archivos.portal;
  const r = resultado?.resumen;

  return (
    <main>
      <header className="barra">
        <div className="marca">
          <img src="/logo.svg" alt="" width="36" height="36" />
          <span><b>JR</b>-Accounting tools</span>
        </div>
      </header>

      <section className="hero">
        <div>
          <p className="etiqueta">IVA Compras · Conciliación</p>
          <h1>Cruce de compras</h1>
          <p className="bajada">Compará la planilla Compras Bis con el Portal IVA Compras y descargá los dos Excel marcados con lo que no coincide.</p>
          <p className="privado"><IconoCandado /> Todo se procesa en tu computadora. Nada se sube a ningún lado.</p>
        </div>
        {MOSTRAR_FOTO && (
        <figure className="polaroid" tabIndex={0}>
          <img src="/juan-y-maia.jpg" alt="Juan y Maia" width="212" height="236" />
          <figcaption>
            <span className="polaroid__firma">Juan <span aria-hidden="true">♥</span> Maia</span>
            <span className="polaroid__mensaje">juntos, somos<br />mucho más lindos<br /><b>te amo hermosa</b></span>
          </figcaption>
        </figure>
        )}
      </section>

      <section
        className={`zona ${arrastrando ? 'zona--activa' : ''}`}
        onDragOver={e => { e.preventDefault(); setArrastrando(true); }}
        onDragLeave={() => setArrastrando(false)}
        onDrop={e => { e.preventDefault(); setArrastrando(false); recibir(e.dataTransfer.files); }}
      >
        <div className="zona__cabeza">
          <div className="zona__icono"><IconoSubir /></div>
          <div className="zona__textos">
            <p className="zona__texto">Soltá acá los dos Excel</p>
            <p className="zona__ayuda">En cualquier orden, solo archivos .xlsx</p>
          </div>
          <button type="button" className="boton boton--secundario" onClick={() => input.current?.click()}>Elegir archivos</button>
        </div>
        <input ref={input} type="file" accept=".xlsx" multiple hidden onChange={e => { recibir(e.target.files); e.target.value = ''; }} />
        <div className="ranuras">
          <Ranura paso={1} titulo="Compras Bis o Mis Comprobantes" archivo={archivos.bis} />
          <Ranura paso={2} titulo="Portal IVA Compras" archivo={archivos.portal} />
        </div>
      </section>

      {error && <p className="error" role="alert">{error}</p>}

      <div className="acciones">
        <button type="button" className="boton boton--grande" disabled={!listos || estado === 'procesando'} onClick={procesar}>
          {estado === 'procesando' ? 'Cruzando…' : 'Cruzar archivos'}
        </button>
        {(listos || resultado) && <button type="button" className="boton boton--texto" onClick={reiniciar}>Empezar de nuevo</button>}
      </div>

      {r && (
        <section className="resultado" aria-live="polite">
          <h2>Período {r.periodo}</h2>
          <p className="control">
            Bis tiene {r.totalBis} comprobantes y el Portal {r.totalPortal}. Cruzaron {r.coincidencias}.
          </p>

          <dl className="cifras">
            <div className="cifra cifra--rojo"><dt>de Bis no están en el Portal</dt><dd>{r.bisSinPortal}</dd>
              {r.bisSinPortalFueraDePeriodo > 0 && <span>{r.bisSinPortalFueraDePeriodo} con fecha fuera del período</span>}</div>
            <div className="cifra cifra--rojo"><dt>del Portal no están en Bis</dt><dd>{r.portalSinBis}</dd></div>
            <div className="cifra cifra--amarillo"><dt>cruzaron con alguna diferencia</dt><dd>{r.conDiferencias}</dd></div>
          </dl>

          <div className="descargas">
            <button type="button" className="boton" onClick={() => descargar(resultado.bis, marcado(archivos.bis.nombre))}>Descargar Compras Bis marcado</button>
            <button type="button" className="boton" onClick={() => descargar(resultado.portal, marcado(archivos.portal.nombre))}>Descargar Portal marcado</button>
          </div>

          <Tabla
            titulo={`De Bis, no están en el Portal (${r.bisSinPortal})`} color="rojo" abierta
            columnas={['Fila', 'Fecha', 'Proveedor', 'Número', 'Total', 'Nota']}
            filas={resultado.detalle.bisSinPortal.map(x => [x.fila, x.fecha, x.proveedor, x.numero || '—', plata(x.total), x.nota])}
          />
          <Tabla
            titulo={`Del Portal, no están en Bis (${r.portalSinBis})`} color="rojo"
            columnas={['Fila', 'Fecha', 'Proveedor', 'Comprobante', 'Total', 'Nota']}
            filas={resultado.detalle.portalSinBis.map(x => [x.fila, x.fecha, x.proveedor, `${x.tipo} ${x.numero}`, plata(x.total), x.nota])}
          />
          <Tabla
            titulo={`Cruzaron con diferencias (${r.conDiferencias})`} color="amarillo"
            columnas={['Fila Bis', 'Fila Portal', 'Proveedor', 'Diferencia']}
            filas={resultado.detalle.conDiferencias.map(x => [x.filaBis, x.filaPortal, x.proveedor, x.nota])}
          />
        </section>
      )}
    </main>
  );
}

function Ranura({ paso, titulo, archivo }) {
  return (
    <div className={`ranura ${archivo ? 'ranura--ok' : ''}`}>
      <span className="ranura__icono">{archivo ? <IconoCheck /> : <IconoHoja />}</span>
      <span className="ranura__textos">
        <span className="ranura__paso">Archivo {paso}</span>
        <span className="ranura__titulo">{titulo}</span>
        <span className="ranura__archivo" title={archivo?.nombre}>{archivo ? archivo.nombre : 'Falta este archivo'}</span>
      </span>
    </div>
  );
}

const trazo = { fill: 'none', stroke: 'currentColor', strokeWidth: 2, strokeLinecap: 'round', strokeLinejoin: 'round', 'aria-hidden': true };

function IconoSubir() {
  return <svg width="28" height="28" viewBox="0 0 24 24" {...trazo}><path d="M12 15V4M7 9l5-5 5 5" /><path d="M4 15v3a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-3" /></svg>;
}
function IconoHoja() {
  return <svg width="20" height="20" viewBox="0 0 24 24" {...trazo}><rect x="4" y="3" width="16" height="18" rx="2" /><path d="M4 9h16M4 15h16M10 3v18" /></svg>;
}
function IconoCheck() {
  return <svg width="20" height="20" viewBox="0 0 24 24" {...trazo} strokeWidth={2.5}><path d="M5 12.5l4.5 4.5L19 7.5" /></svg>;
}
function IconoCandado() {
  return <svg width="18" height="18" viewBox="0 0 24 24" {...trazo}><rect x="5" y="11" width="14" height="10" rx="2" /><path d="M8 11V8a4 4 0 0 1 8 0v3" /></svg>;
}

function Tabla({ titulo, columnas, filas, color, abierta }) {
  if (!filas.length) return null;
  return (
    <details className={`detalle detalle--${color}`} open={abierta}>
      <summary>{titulo}</summary>
      <div className="tabla">
        <table>
          <thead><tr>{columnas.map(c => <th key={c}>{c}</th>)}</tr></thead>
          <tbody>
            {filas.map((f, i) => (
              <tr key={i}>{f.map((v, j) => <td key={j} className={columnas[j] === 'Total' ? 'num' : ''}>{v}</td>)}</tr>
            ))}
          </tbody>
        </table>
      </div>
    </details>
  );
}
