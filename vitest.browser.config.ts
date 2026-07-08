import { defineConfig } from 'vitest/config'

// Playwright-driven tests for the page-aware renderer. These launch headless
// Chromium (needs `npx playwright install chromium` once) and are kept out of
// the fast node unit suite. Run with `npm run test:browser`.
export default defineConfig({
  test: {
    environment: 'node',
    include: ['test/browser/**/*.browser.test.ts'],
    testTimeout: 30_000,
    hookTimeout: 60_000,
  },
})
