import type { Metadata, Viewport } from 'next';
import '@fontsource/carlito/400.css';
import '@fontsource/carlito/400-italic.css';
import '@fontsource/carlito/700.css';
import '@fontsource/carlito/700-italic.css';
import './globals.css';

export const metadata: Metadata = {
  title: 'MyExcel',
  description: 'A local-first spreadsheet',
  icons: { icon: '/icon.svg' },
};

export const viewport: Viewport = {
  width: 'device-width',
  initialScale: 1,
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en" suppressHydrationWarning>
      <body>{children}</body>
    </html>
  );
}
