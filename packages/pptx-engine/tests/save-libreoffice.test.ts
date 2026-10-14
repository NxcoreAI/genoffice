import { describe, it, expect } from 'vitest'
import { execFileSync } from 'node:child_process'
import { existsSync, mkdtempSync, readFileSync, rmSync, statSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { inflateRawSync } from 'node:zlib'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import { addPicture, openPptx, savePptxToFile } from '../src/index'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

const OFF = { x: 914400, y: 914400, cx: 1828800, cy: 914400 }

// 1x1 red PNG
const PNG_RED = Uint8Array.from(
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=',
    'base64',
  ),
)
// 1x1 blue PNG (distinct bytes)
const PNG_BLUE = Uint8Array.from(
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
    'base64',
  ),
)
// 1x1 JPEG (grayscale)
const JPEG_1PX = Uint8Array.from(
  Buffer.from(
    '/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/wAALCAABAAEBAREA/8QAFAABAAAAAAAAAAAAAAAAAAAACf/EABQQAQAAAAAAAAAAAAAAAAAAAAD/2gAIAQEAAD8AKp//2Q==',
    'base64',
  ),
)

const CRC_TABLE = (() => {
  const table = new Uint32Array(256)
  for (let n = 0; n < 256; n++) {
    let c = n
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
    table[n] = c >>> 0
  }
  return table
})()

function crc32(buf: Uint8Array): number {
  let c = 0xffffffff
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]!) & 0xff]! ^ (c >>> 8)
  return (c ^ 0xffffffff) >>> 0
}

interface LocalEntry {
  name: string
  flags: number
  method: number
  crc: number
  csize: number
  usize: number
}

/**
 * Walk the local file headers sequentially. Relies on real csize values in the
 * local headers — which is exactly the property under test: streamed entries
 * (data descriptors) carry csize=0 here and make the walk fall off the rails.
 */
function walkLocalHeaders(buf: Buffer): { entries: LocalEntry[]; centralDirOffset: number } {
  const entries: LocalEntry[] = []
  let pos = 0
  for (;;) {
    if (pos + 4 > buf.length || buf.readUInt32LE(pos) !== 0x04034b50) break
    const flags = buf.readUInt16LE(pos + 6)
    const method = buf.readUInt16LE(pos + 8)
    const crc = buf.readUInt32LE(pos + 14)
    const csize = buf.readUInt32LE(pos + 18)
    const usize = buf.readUInt32LE(pos + 22)
    const nlen = buf.readUInt16LE(pos + 26)
    const elen = buf.readUInt16LE(pos + 28)
    const name = buf.toString('utf8', pos + 30, pos + 30 + nlen)
    entries.push({ name, flags, method, crc, csize, usize })
    pos += 30 + nlen + elen + csize
  }
  return { entries, centralDirOffset: pos }
}

/** Build a deck with several embedded images on its slides (the AI generation shape). */
async function openMultiImageDeck() {
  const opened = await openPptx(fx('01_standard_business.pptx'))
  const images = [JPEG_1PX, PNG_RED, PNG_BLUE, JPEG_1PX, PNG_RED, PNG_BLUE, JPEG_1PX]
  images.forEach((bytes, i) => {
    const slide = opened.deck.slides[i % opened.deck.slides.length]!
    const added = addPicture(opened, slide, { bytes, ext: i % 3 === 0 ? 'jpg' : 'png', offset: { ...OFF } })
    expect(added, `addPicture #${i}`).not.toBeNull()
  })
  return opened
}

describe('savePptxToFile package layout (LibreOffice compatibility)', () => {
  it('multi-image deck: no data descriptors, real local sizes, media stored', async () => {
    const opened = await openMultiImageDeck()
    const target = join(mkdtempSync(join(tmpdir(), 'save-lo-')), 'out.pptx')
    await savePptxToFile(opened, target)

    const buf = readFileSync(target)
    const { entries, centralDirOffset } = walkLocalHeaders(buf)

    // The sequential walk must land exactly on the central directory — it
    // derails if any entry still relies on a trailing data descriptor.
    expect(
      buf.readUInt32LE(centralDirOffset),
      'sequential walk should end at the central directory signature',
    ).toBe(0x02014b50)

    const media = entries.filter((e) => /^ppt\/media\//.test(e.name))
    expect(media.length).toBeGreaterThanOrEqual(7)

    for (const entry of entries) {
      // Data descriptors (general-purpose flag bit 3, sizes zeroed in the local
      // header) are what LibreOffice chokes on with stored media entries.
      expect(entry.flags & 0x08, `${entry.name}: no data-descriptor flag`).toBe(0)
      if (entry.name.endsWith('/')) continue
      expect(entry.crc, `${entry.name}: crc present`).not.toBe(0)
      // Local header sizes are real: stored entries carry the raw bytes,
      // deflated entries the compressed stream — both non-zero here.
      expect(entry.csize, `${entry.name}: local csize filled in`).toBeGreaterThan(0)
      expect(entry.usize, `${entry.name}: local usize filled in`).toBeGreaterThan(0)
    }

    // Verify each entry's bytes against its local header crc: stored entries
    // hash directly; deflated entries inflate first (the header crc is over the
    // uncompressed bytes).
    let pos = 0
    for (;;) {
      if (buf.readUInt32LE(pos) !== 0x04034b50) break
      const flags = buf.readUInt16LE(pos + 6)
      const method = buf.readUInt16LE(pos + 8)
      const crc = buf.readUInt32LE(pos + 14)
      const csize = buf.readUInt32LE(pos + 18)
      const nlen = buf.readUInt16LE(pos + 26)
      const elen = buf.readUInt16LE(pos + 28)
      const name = buf.toString('utf8', pos + 30, pos + 30 + nlen)
      const dataStart = pos + 30 + nlen + elen
      const payload = buf.subarray(dataStart, dataStart + csize)
      if (!name.endsWith('/') && !(flags & 0x08)) {
        const raw = method === 0 ? payload : inflateRawSync(payload)
        expect(crc32(raw), `${name}: crc matches content`).toBe(crc)
      }
      pos = dataStart + csize
    }

    // Media stays STORED (the optimization is intentional), xml stays deflated.
    for (const entry of media) {
      if (entry.name.endsWith('/')) continue
      expect(entry.method, `${entry.name}: stored`).toBe(0)
      expect(entry.csize).toBe(entry.usize)
    }
    const xml = entries.find((e) => e.name === 'ppt/presentation.xml')!
    expect(xml.method).toBe(8)

    // And the package still reopens with the embedded media present.
    const reopened = await openPptx(buf)
    const mediaPaths = [...reopened.archive.entries.keys()].filter((p) => /^ppt\/media\//.test(p))
    expect(mediaPaths.length).toBeGreaterThanOrEqual(7)
  })
})

describe('savePptxToFile converts in LibreOffice', () => {
  const soffice = ['/Applications/LibreOffice.app/Contents/MacOS/soffice', '/usr/local/bin/soffice', '/opt/homebrew/bin/soffice']
    .find(existsSync)
  const skip = !soffice

  it.skipIf(skip)('multi-image deck loads (regression: streamed stored media)', async () => {
    const opened = await openMultiImageDeck()
    const dir = mkdtempSync(join(tmpdir(), 'save-lo-conv-'))
    const target = join(dir, 'out.pptx')
    await savePptxToFile(opened, target)

    const profile = join(dir, 'profile')
    execFileSync(soffice!, [
      '--headless',
      `-env:UserInstallation=file://${profile}`,
      '--convert-to',
      'pdf',
      '--outdir',
      dir,
      target,
    ], { timeout: 120_000 })

    const pdf = join(dir, 'out.pdf')
    expect(statSync(pdf).size, 'converted pdf exists and is non-empty').toBeGreaterThan(0)
    rmSync(dir, { recursive: true, force: true })
  }, 180_000)
})
