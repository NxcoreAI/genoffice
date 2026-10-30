// 由 everroom-office-proto scripts/prepare-markdown-runtime.mjs 生成（不进 git）。
import { resolve } from 'node:path'
import react from '@vitejs/plugin-react'
import { defineConfig, externalizeDepsPlugin } from 'electron-vite'

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

const OUT = process.env.GENOFFICE_MARKDOWN_OUT ?? resolve(__dirname, '../../../..', 'build/genoffice-runtime/markdown')

export default defineConfig({
  main: {
    build: {
      outDir: resolve(OUT, 'main'),
      emptyOutDir: true,
      rollupOptions: { input: { embed: resolve(__dirname, 'src/main/markdown-main.ts') } },
    },
    plugins: [externalizeDepsPlugin({ exclude: ['@genoffice/i18n', '@genoffice/electron-utils'] })],
  },
  preload: {
    build: {
      outDir: resolve(OUT, 'preload'),
      emptyOutDir: true,
      rollupOptions: { input: { index: resolve(__dirname, 'src/preload/index.ts') } },
    },
    plugins: [externalizeDepsPlugin({ exclude: ['@genoffice/i18n', '@genoffice/electron-utils'] })],
  },
  renderer: {
    root: resolve(__dirname, 'src/renderer'),
    build: { outDir: resolve(OUT, 'renderer'), emptyOutDir: true },
    plugins: [react()],
    resolve: { dedupe: TIPTAP_DEDUPE },
  },
})
