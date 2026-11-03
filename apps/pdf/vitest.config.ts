import { defineConfig } from 'vitest/config'

export default defineConfig({
  define: {
    // main-process code branches on this (electron-vite injects it in real
    // builds); tests run the standalone closure, so default it to false
    __GENOFFICE_EMBED_ONLY__: JSON.stringify(process.env.GENOFFICE_EMBED_ONLY === '1'),
  },
  test: {
    include: ['tests/**/*.test.ts'],
    environment: 'jsdom',
    testTimeout: 20000,
  },
})
