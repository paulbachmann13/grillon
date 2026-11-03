import { defineConfig, devices } from '@playwright/test';

/**
 * Les tests tournent sur le build de production (service worker compris),
 * servi par `vite preview` sur le même sous-chemin que GitHub Pages.
 */
// Réglable : un autre projet peut occuper 4173 en même temps (GRILLON_E2E_PORT=4180 npm run test:e2e).
const PORT = Number(process.env.GRILLON_E2E_PORT) || 4173;

export default defineConfig({
  testDir: 'tests/e2e',
  timeout: 120_000,
  expect: { timeout: 15_000 },
  fullyParallel: false,
  workers: 1,
  retries: process.env.CI ? 1 : 0,
  reporter: process.env.CI ? [['list'], ['html', { open: 'never' }]] : 'list',
  use: {
    baseURL: `http://localhost:${PORT}/grillon/`,
    locale: 'fr-FR',
    acceptDownloads: true,
    trace: 'retain-on-failure',
  },
  projects: [
    { name: 'chromium', use: { ...devices['Desktop Chrome'] } },
    { name: 'webkit-ipad', use: { ...devices['iPad (gen 7) landscape'] }, testMatch: /smoke\.spec\.ts/ },
  ],
  webServer: {
    command: `npm run build && npm run preview -- --port ${PORT}`,
    env: { GRILLON_BASE: '/grillon/' },
    url: `http://localhost:${PORT}/grillon/`,
    reuseExistingServer: !process.env.CI,
    timeout: 300_000,
  },
});
