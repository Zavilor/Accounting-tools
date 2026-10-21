// Corre el cruce desde la terminal, sin la web:
//   npm run probar -- ruta/Compras_Bis.xlsx ruta/Portal_Iva.xlsx
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import fs from 'fs';
import path from 'path';
import { cruzar } from '../lib/cruce.js';

const [bis, portal] = process.argv.slice(2);
if (!bis || !portal) { console.error('Uso: npm run probar -- <Compras Bis.xlsx> <Portal IVA.xlsx>'); process.exit(1); }
const r = await cruzar(ExcelJS, JSZip, fs.readFileSync(bis), fs.readFileSync(portal));
const salida = f => path.join(path.dirname(f), path.basename(f, '.xlsx') + '_MARCADO.xlsx');
fs.writeFileSync(salida(bis), Buffer.from(r.bis));
fs.writeFileSync(salida(portal), Buffer.from(r.portal));
console.log(r.resumen);
console.log('Archivos:', salida(bis), '|', salida(portal));
