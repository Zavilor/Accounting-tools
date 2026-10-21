import './globals.css';
import { Caveat, Public_Sans } from 'next/font/google';

const fuente = Public_Sans({ subsets: ['latin'], weight: ['400', '500', '700'], display: 'swap' });
const mano = Caveat({ subsets: ['latin'], weight: ['500', '700'], variable: '--fuente-mano', display: 'swap' });

export const metadata = {
  title: 'Cruce de compras | JR-Accounting tools',
  description: 'Cruce de la planilla Compras Bis contra el Portal IVA Compras',
  icons: { icon: '/logo.svg' },
};

export default function RootLayout({ children }) {
  return (
    <html lang="es">
      <body className={`${fuente.className} ${mano.variable}`}>{children}</body>
    </html>
  );
}
