import { defineConfig } from 'vitest/config'

export default defineConfig({
  define: {
    // tests run against the standalone (non-embed) build surface
    __GENOFFICE_EMBED_ONLY__: JSON.stringify(false),
  },
  test: {
    include: ['tests/**/*.test.ts'],
    environment: 'jsdom',
    testTimeout: 20000,
  },
})
