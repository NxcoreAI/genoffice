/**
 * 组件画廊 v2 PPTX 构建器（genoffice 侧，临时脚本、不入库）：
 * 读 /tmp/ppt-gallery/specs.json，逐页 parsePageSpec + buildPagePptx（真机同链）。
 * 图片走本地映射（https://gallery.local/img/<file> → gap-compare-assets/img）。
 */
import { readFileSync, writeFileSync, mkdirSync } from "node:fs";
import { parsePageSpec, buildPagePptx } from "./src/main/page-spec";
import {
  mergeSlideFromPptx,
  openPptx,
  promoteSlideBackground,
  savePptx,
} from "@genoffice/pptx-engine";

const OUT = "/tmp/ppt-gallery";
const IMG_DIR = "/Users/xjwang/Desktop/work/EerRoom/EverRoomMac/materials/ppt-style-experiments/gap-compare-assets/img";
const specs = JSON.parse(readFileSync(`${OUT}/specs.json`, "utf8")) as {
  file: string;
  title: string;
  spec: unknown;
}[];

function imageDims(bytes: Uint8Array): { width: number; height: number } | null {
  const b = Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (b[0] === 0x89 && b[1] === 0x50) {
    return { width: b.readUInt32BE(16), height: b.readUInt32BE(20) };
  }
  if (b[0] === 0xff && b[1] === 0xd8) {
    let off = 2;
    while (off < b.length - 8) {
      if (b[off] !== 0xff) { off++; continue; }
      const marker = b[off + 1];
      const len = b.readUInt16BE(off + 2);
      if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
        return { height: b.readUInt16BE(off + 5), width: b.readUInt16BE(off + 7) };
      }
      off += 2 + len;
    }
  }
  return null;
}

const deps = {
  async fetchImage(url: string) {
    const m = /^https:\/\/gallery\.local\/img\/(.+)$/.exec(url);
    if (!m) return null;
    try {
      const bytes = readFileSync(`${IMG_DIR}/${m[1]}`);
      const ext = m[1].endsWith(".png") ? "png" : "jpg";
      return { bytes: new Uint8Array(bytes), ext };
    } catch {
      return null;
    }
  },
  imageDims,
};

mkdirSync(`${OUT}/pptx`, { recursive: true });
const singleBytes: Uint8Array[] = [];
for (const [index, page] of specs.entries()) {
  const parsed = parsePageSpec(JSON.stringify(page.spec));
  if (!parsed.ok) {
    console.error(`[parse-fail] ${page.file}: ${JSON.stringify(parsed.issues ?? parsed).slice(0, 300)}`);
    continue;
  }
  const built = await buildPagePptx(parsed.spec, deps);
  writeFileSync(`${OUT}/pptx/${String(index + 1).padStart(2, "0")}-${page.file}.pptx`, built.bytes);
  singleBytes.push(built.bytes);
  const imgNote = built.imageFailures.length ? `（图片缺失 ${built.imageFailures.length}）` : "";
  console.log(`[ok] ${page.file} (${(built.bytes.length / 1024).toFixed(0)} KB)${imgNote}`);
}

if (singleBytes.length > 0) {
  const base = await openPptx(singleBytes[0]!);
  for (const one of singleBytes.slice(1)) await mergeSlideFromPptx(base, one);
  for (const slide of base.deck.slides) promoteSlideBackground(slide, base.deck.size);
  writeFileSync(`${OUT}/component-gallery-v2.pptx`, await savePptx(base));
  console.log(`[deck] ${OUT}/component-gallery-v2.pptx (${singleBytes.length} 页)`);
}
