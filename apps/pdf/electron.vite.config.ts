import { createRequire } from 'node:module'
import { dirname, join, resolve } from 'node:path'
import react from '@vitejs/plugin-react'
import { defineConfig, externalizeDepsPlugin } from 'electron-vite'
import { normalizePath } from 'vite'
import { viteStaticCopy } from 'vite-plugin-static-copy'

// Non-embedded CMaps/standard fonts (e.g. CJK) need pdfjs data dirs, shipped with renderer output
const require = createRequire(import.meta.url)
const pdfjsRoot = dirname(dirname(require.resolve('pdfjs-dist/package.json')))
// vite-plugin-static-copy globs require POSIX separators; join() breaks on Windows
const pdfjsDir = (sub: string) => normalizePath(join(pdfjsRoot, 'pdfjs-dist', sub))

const embedOnly = process.env.GENOFFICE_EMBED_ONLY === '1'

export default defineConfig({
  // @genoffice/i18n ships as TS source; pdf-lib's package only includes out/** — both must be bundled
  main: {
    build: {
      rollupOptions: {
        input: {
          // Embed hosts load embed.js only; the standalone bootstrap must not
          // ship in the EverRoom closure.
          ...(!embedOnly ? { index: resolve(__dirname, 'src/main/index.ts') } : {}),
          embed: resolve(__dirname, 'src/main/embed.ts'),
        },
      },
    },
    define: {
      __GENOFFICE_EMBED_ONLY__: JSON.stringify(embedOnly),
    },
    plugins: [
      externalizeDepsPlugin({
        exclude: [
          '@genoffice/i18n',
          'pdf-lib',
          '@genoffice/electron-utils',
          '@genoffice/font-metrics',
        ],
      }),
    ],
  },
  preload: {
    // i18n and electron-utils ship as TS source — must be bundled, not left external
    plugins: [externalizeDepsPlugin({ exclude: ['@genoffice/i18n', '@genoffice/electron-utils'] })],
  },
  renderer: {
    define: {
      // EverRoom embed build: folds the Genspark AI surfaces out of the
      // renderer (the embed closure ships no AI backend to talk to).
      __GENOFFICE_EMBED_ONLY__: JSON.stringify(embedOnly),
    },
    plugins: [
      react(),
      viteStaticCopy({
        targets: [
          { src: pdfjsDir('cmaps'), dest: 'pdfjs' },
          { src: pdfjsDir('standard_fonts'), dest: 'pdfjs' },
          { src: pdfjsDir('wasm'), dest: 'pdfjs' },
        ],
      }),
    ],
    server: {
      port: Number(process.env.PDF_DEV_PORT) || 5176,
      strictPort: Boolean(process.env.PDF_DEV_PORT),
    },
  },
})
