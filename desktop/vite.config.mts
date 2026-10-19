import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig(({ command }) => ({
  plugins: [react(), {
    name: 'local-dev-csp',
    transformIndexHtml(html) {
      return command === 'serve' ? html.replace("script-src 'self'", "script-src 'self' 'nonce-excelrouter-dev'") : html;
    },
  }],
  html: command === 'serve' ? { cspNonce: 'excelrouter-dev' } : undefined,
  base: './',
  server: { host: '127.0.0.1', port: 5173, strictPort: true },
  build: { outDir: 'dist', emptyOutDir: true },
}));
