import { resolve } from 'node:path'
import react from '@vitejs/plugin-react'
import { defineConfig, externalizeDepsPlugin } from 'electron-vite'

// npm hoists some @tiptap packages to the repo root (shared with docs at a
// different version) and nests others under this app — dedupe forces every
// import onto this app's single copy so the bundle never carries two cores.
const TIPTAP_DEDUPE = [
  '@tiptap/core',
  '@tiptap/pm',
  '@tiptap/react',
  '@tiptap/extensions',
  '@tiptap/extension-list',
  '@tiptap/extension-table',
  '@tiptap/extension-image',
  '@tiptap/suggestion',
  '@tiptap/markdown',
  '@tiptap/extension-highlight',
  '@tiptap/extension-code-block',
]

// EverRoom embed build: standalone bootstrap is folded out; the bundled
// Genspark AI surface is replaced by the host's own agent panel (docs parity).
const embedOnly = process.env.GENOFFICE_EMBED_ONLY === '1'

export default defineConfig({
  // @genoffice/i18n and @genoffice/electron-utils ship as TS source — must be bundled
  main: {
    build: {
      rollupOptions: {
        input: {
          ...(!embedOnly ? { index: resolve(__dirname, 'src/main/index.ts') } : {}),
          embed: resolve(__dirname, 'src/main/embed.ts'),
        },
      },
    },
    define: {
      __GENOFFICE_EMBED_ONLY__: JSON.stringify(embedOnly),
    },
    plugins: [externalizeDepsPlugin({ exclude: ['@genoffice/i18n', '@genoffice/electron-utils'] })],
  },
  preload: {
    // same bundling requirement as main (see comment above)
    plugins: [externalizeDepsPlugin({ exclude: ['@genoffice/i18n', '@genoffice/electron-utils'] })],
  },
  renderer: {
    define: {
      __GENOFFICE_EMBED_ONLY__: JSON.stringify(embedOnly),
    },
    plugins: [react()],
    resolve: { dedupe: TIPTAP_DEDUPE },
    server: {
      port: Number(process.env.MARKDOWN_DEV_PORT) || 5177,
      strictPort: Boolean(process.env.MARKDOWN_DEV_PORT),
    },
  },
})
