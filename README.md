# Cruce de compras

Cruza la planilla **Compras Bis** contra el **Portal IVA Compras** y devuelve los dos Excel marcados:

- **Rojo:** el comprobante no está en el otro archivo.
- **Amarillo:** está en los dos pero con alguna diferencia (CUIT, número, fecha, tipo o importe).
- Se agrega una columna al final con el detalle y la fila del otro archivo con la que cruzó.

Todo corre en el navegador: los archivos no se suben a ningún servidor.

Los archivos originales no se reescriben: se edita solo el XML interno de la hoja y de los estilos para
agregar los colores y la columna nueva. Fechas, fórmulas, valores, formatos y todo lo demás quedan
exactamente como estaban.

## Criterio de cruce

Cada comprobante del Portal se usa una sola vez. Números y CUIT se normalizan (sin guiones, espacios, letras ni ceros a la izquierda).

1. CUIT + punto de venta + número.
2. Si no cruza: punto de venta + número + total exacto (detecta CUIT mal cargado).
3. Si no cruza: CUIT + total exacto, solo si hay un único candidato (detecta número mal cargado o vacío).

Además marca duplicados en Bis, filas fuera del período y facturas del Portal anuladas por una nota de crédito.

El cruce está en `lib/cruce.js` y el marcado de los Excel en `lib/marcar-xlsx.js`.

## Uso local

```bash
npm install
npm run dev                                   # http://localhost:3000
npm run probar -- Compras_Bis.xlsx Portal.xlsx  # mismo cruce desde la terminal
```

## Deploy en Netlify

Conectá el repo en Netlify. El `netlify.toml` ya tiene el build (`npm run build`) y la carpeta a publicar (`out`).
