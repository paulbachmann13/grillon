import { VitePWA } from 'vite-plugin-pwa';
import { defineConfig } from 'vitest/config';

// GitHub Pages sert l'app sous /grillon/ ; en local (dev, Tauri) on reste à la racine.
const base = process.env.GRILLON_BASE ?? '/';

export default defineConfig({
  base,
  build: { target: 'es2022', chunkSizeWarningLimit: 8000 },
  worker: { format: 'es' },
  plugins: [
    VitePWA({
      registerType: 'prompt',
      injectRegister: false,
      includeAssets: ['favicon.ico', 'apple-touch-icon-180x180.png', 'icon.svg'],
      manifest: {
        name: 'Grillon — tableur',
        short_name: 'Grillon',
        description: 'Tableur local et gratuit, compatible Excel (.xlsx), qui fonctionne hors ligne.',
        lang: 'fr',
        theme_color: '#1f7a4d',
        background_color: '#ffffff',
        display: 'standalone',
        orientation: 'any',
        start_url: '.',
        scope: '.',
        icons: [
          { src: 'pwa-64x64.png', sizes: '64x64', type: 'image/png' },
          { src: 'pwa-192x192.png', sizes: '192x192', type: 'image/png' },
          { src: 'pwa-512x512.png', sizes: '512x512', type: 'image/png' },
          { src: 'maskable-icon-512x512.png', sizes: '512x512', type: 'image/png', purpose: 'maskable' },
        ],
        file_handlers: [
          {
            action: '.',
            accept: {
              'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
              'text/csv': ['.csv'],
            },
          },
        ],
      },
      workbox: {
        // Univer + ExcelJS font plusieurs Mo : on les met tous en cache pour le hors-ligne.
        maximumFileSizeToCacheInBytes: 20 * 1024 * 1024,
        globPatterns: ['**/*.{js,css,html,ico,png,svg,woff,woff2,wasm}'],
        navigateFallback: 'index.html',
      },
    }),
  ],
  test: {
    include: ['tests/unit/**/*.test.ts'],
    testTimeout: 120_000,
  },
});
