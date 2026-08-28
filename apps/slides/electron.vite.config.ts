import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import react from '@vitejs/plugin-react'
import { defineConfig, externalizeDepsPlugin } from 'electron-vite'

const here = dirname(fileURLToPath(import.meta.url))

// Pin resolution to this repo's workspace sources (matches tsconfig paths;
// avoids bundling stale implementations when node_modules links point elsewhere)
const workspaceAlias = {
  // Subpath before the bare name: string aliases are prefix replacements
  '@genoffice/pptx-engine/table-grid': resolve(
    here,
    '../../packages/pptx-engine/src/table-grid.ts',
  ),
  '@genoffice/pptx-engine/identity': resolve(here, '../../packages/pptx-engine/src/identity.ts'),
  '@genoffice/pptx-engine/custgeom': resolve(here, '../../packages/pptx-engine/src/custgeom.ts'),
  '@genoffice/pptx-engine/background-promote': resolve(
    here,
    '../../packages/pptx-engine/src/background-promote.ts',
  ),
  '@genoffice/pptx-engine': resolve(here, '../../packages/pptx-engine/src/index.ts'),
  '@genoffice/pptx-render/preset-geometry': resolve(
    here,
    '../../packages/pptx-render/src/preset-geometry.ts',
  ),
  '@genoffice/pptx-render': resolve(here, '../../packages/pptx-render/src/index.ts'),
  // Metafile (EMF/WMF) rasterizer shared with the docs engine (renderer-only: needs canvas)
  '@genoffice/docx-engine/metafile': resolve(here, '../../packages/docx-engine/src/metafile.ts'),
}

export default defineConfig({
  // Main process/preload must bundle @genoffice/* sources (they are pulled in as TS
  // source with extensionless relative imports; externalizing them under Node
  // yields ERR_MODULE_NOT_FOUND).
  main: {
    build: {
      rollupOptions: {
        input: {
          // The standalone bootstrap (updater, GenOffice AI/login IPC, project
          // store) stays out of EverRoom embed builds; only the embed entry ships.
          ...(process.env.GENOFFICE_EMBED_ONLY === '1'
            ? {}
            : { index: resolve(here, 'src/main/index.ts') }),
          embed: resolve(here, 'src/main/embed.ts'),
        },
      },
    },
    define: {
      __GENOFFICE_EMBED_ONLY__: JSON.stringify(process.env.GENOFFICE_EMBED_ONLY === '1'),
    },
    resolve: { alias: workspaceAlias },
    // Bundle opentype.js too (the packaged app ships only out/**, so external deps are unresolvable at runtime)
    plugins: [
      externalizeDepsPlugin({
        exclude: [
          '@genoffice/pptx-engine',
          '@genoffice/pptx-render',
          '@genoffice/ai-search',
          '@genoffice/file-parse',
          '@genoffice/electron-utils',
          'opentype.js',
          // EverRoom embed closure: the copied runtime has no node_modules, so
          // every main-process npm dep must be bundled (prepare script asserts).
          'pngjs',
          'acorn',
          'harfbuzzjs',
          'utif2',
        ],
      }),
    ],
  },
  preload: {
    // electron-utils ships raw TS source — must be bundled, not left external
    plugins: [externalizeDepsPlugin({ exclude: ['@genoffice/electron-utils'] })],
  },
  renderer: {
    define: {
      __GENOFFICE_EMBED_ONLY__: JSON.stringify(process.env.GENOFFICE_EMBED_ONLY === '1'),
    },
    resolve: { alias: workspaceAlias },
    plugins: [react()],
    server: {
      port: Number(process.env.SLIDES_DEV_PORT) || 5175,
      strictPort: Boolean(process.env.SLIDES_DEV_PORT),
    },
  },
})
