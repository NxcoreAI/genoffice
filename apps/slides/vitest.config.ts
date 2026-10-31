import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { defineConfig } from 'vitest/config'

const here = dirname(fileURLToPath(import.meta.url))

export default defineConfig({
  // Unit tests exercise the standalone renderer: the embed-only define folds
  // AI entries out of the EverRoom runtime, so keep them on here (the electron
  // builds still fold it from GENOFFICE_EMBED_ONLY).
  define: {
    __GENOFFICE_EMBED_ONLY__: JSON.stringify(false),
  },
  // Pin resolution to this repo's workspace sources (matches tsconfig paths)
  resolve: {
    alias: {
      // Subpath before the bare name: string aliases are prefix replacements
      '@genoffice/pptx-engine/table-grid': resolve(
        here,
        '../../packages/pptx-engine/src/table-grid.ts',
      ),
      '@genoffice/pptx-engine/identity': resolve(
        here,
        '../../packages/pptx-engine/src/identity.ts',
      ),
      '@genoffice/pptx-engine/background-promote': resolve(
        here,
        '../../packages/pptx-engine/src/background-promote.ts',
      ),
      '@genoffice/pptx-engine/custgeom': resolve(
        here,
        '../../packages/pptx-engine/src/custgeom.ts',
      ),
      '@genoffice/pptx-engine': resolve(here, '../../packages/pptx-engine/src/index.ts'),
      '@genoffice/pptx-render/preset-geometry': resolve(
        here,
        '../../packages/pptx-render/src/preset-geometry.ts',
      ),
      '@genoffice/pptx-render': resolve(here, '../../packages/pptx-render/src/index.ts'),
      '@genoffice/docx-engine/metafile': resolve(
        here,
        '../../packages/docx-engine/src/metafile.ts',
      ),
    },
  },
  test: {
    include: ['tests/**/*.test.ts'],
    environment: 'jsdom',
    testTimeout: 20000,
  },
})
