import { defineConfig } from 'vitest/config'

export default defineConfig({
  test: {
    environment: 'node',
    include: ['test/**/*.test.ts'],
    // Browser-driven tests (Playwright) run via `npm run test:browser` with
    // their own config; they need a real layout engine, not node/jsdom.
    exclude: ['test/browser/**', 'node_modules/**'],
  },
})
