import { defineConfig } from '@playwright/test';
export default defineConfig({
  testDir: './e2e',
  timeout: 120000,
  use: { baseURL: 'http://localhost:3100', viewport: { width: 1440, height: 900 } },
  webServer: { command: 'npm run start', url: 'http://localhost:3100', reuseExistingServer: true, timeout: 120000 },
});
