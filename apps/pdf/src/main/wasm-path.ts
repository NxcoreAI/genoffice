import { existsSync } from 'node:fs'
import { createRequire } from 'node:module'
import { dirname, join } from 'node:path'

/**
 * Runtime wasm assets live in node_modules during dev/tests but the packaged app
 * ships no node_modules (everything is bundled) — electron-builder copies them
 * into Resources/wasm instead (see apps/shell/electron-builder.cjs extraResources).
 * The EverRoom host's genoffice runtime is a standalone directory with no
 * node_modules either; its prepare script stages the same files under
 * <runtime>/wasm (apps/desktop/scripts/prepare-genoffice-runtime.mjs), resolved
 * relative to this bundle: <runtime>/pdf/main/chunks/… → ../../../wasm/<file>.
 */
const packagedPath = (fileName: string) => join(process.resourcesPath, 'wasm', fileName)

const runtimeSiblingPath = (fileName: string) =>
  join(__dirname, '..', '..', '..', 'wasm', fileName)

const req = () => createRequire(import.meta.url)

export function pdfiumWasmPath(): string {
  try {
    return req().resolve('@embedpdf/pdfium/pdfium.wasm')
  } catch {
    const sibling = runtimeSiblingPath('pdfium.wasm')
    if (existsSync(sibling)) return sibling
    return packagedPath('pdfium.wasm')
  }
}

export function hbSubsetWasmPath(): string {
  const r = req()
  try {
    // harfbuzzjs ≤0.10 ships hb-subset.wasm at the package root with no exports map
    return r.resolve('harfbuzzjs/hb-subset.wasm')
  } catch {
    /* fall through */
  }
  try {
    // harfbuzzjs ≥1.x seals subpaths; the wasm sits next to the exported entry point
    const p = join(dirname(r.resolve('harfbuzzjs')), 'harfbuzz-subset.wasm')
    if (existsSync(p)) return p
  } catch {
    /* fall through */
  }
  const sibling = runtimeSiblingPath('hb-subset.wasm')
  if (existsSync(sibling)) return sibling
  return packagedPath('hb-subset.wasm')
}
