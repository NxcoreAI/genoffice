import react from '@vitejs/plugin-react'
import { resolve } from 'node:path'
import { defineConfig, externalizeDepsPlugin } from 'electron-vite'

export default defineConfig({
  main: {
    build: {
      rollupOptions: {
        input: {
          ...(!process.env.GENOFFICE_EMBED_ONLY ? { index: resolve(__dirname, 'src/main/index.ts') } : {}),
          embed: resolve(__dirname, 'src/main/embed.ts'),
        },
      },
    },
    // @genoffice/* workspace packages ship TS source (no build step, no
    // compiled entry point) — externalizing them makes Node's ESM loader try
    // to resolve their relative imports at runtime and fail. Bundle those;
    // externalize everything else (Electron, node builtins). zod and jszip
    // are bundled too so the EverRoom embed closure stays self-contained —
    // the copied runtime has no node_modules to resolve externals from.
    plugins: [
      externalizeDepsPlugin({
        exclude: [
          '@genoffice/ai-provider',
          '@genoffice/agent-core',
          '@genoffice/ai-search',
          '@genoffice/docx-engine',
          '@genoffice/file-parse',
          '@genoffice/electron-utils',
          '@genoffice/i18n',
          '@univerjs/core', '@univerjs/data-validation', '@univerjs/engine-render',
          '@univerjs/find-replace', '@univerjs/preset-sheets-conditional-formatting',
          '@univerjs/preset-sheets-core', '@univerjs/preset-sheets-data-validation',
          '@univerjs/preset-sheets-drawing', '@univerjs/preset-sheets-filter',
          '@univerjs/preset-sheets-find-replace', '@univerjs/preset-sheets-note',
          '@univerjs/preset-sheets-sort', '@univerjs/preset-sheets-table',
          '@univerjs/sheets-filter', '@univerjs/themes', 'rxjs', 'fast-xml-parser',
          'zod', 'jszip',
        ],
      }),
    ],
  },
  preload: {
    // Sandboxed preload scripts cannot require arbitrary npm packages at
    // runtime, so the drop-open bridge must be bundled, not externalized.
    plugins: [externalizeDepsPlugin({ exclude: ['@genoffice/electron-utils'] })],
  },
  renderer: {
    plugins: [react()],
    define: {
      // EverRoom embed build: folds the Genspark AI surfaces out of the renderer
      // (the embed closure ships no AI backend to talk to).
      __GENOFFICE_EMBED_ONLY__: JSON.stringify(process.env.GENOFFICE_EMBED_ONLY === '1'),
    },
  },
})
