/** Image helpers shared by the editor and converters. Images live in the document as data URIs. */

export interface ImageType {
  mime: 'image/png' | 'image/jpeg' | 'image/gif';
  ext: 'png' | 'jpeg' | 'gif';
}

/** Formats stored as-is: the ones every version of Word can show. Anything else is converted first. */
const BY_BASE64_PREFIX: [string, ImageType][] = [
  ['iVBORw0KGgo', { mime: 'image/png', ext: 'png' }],
  ['/9j/', { mime: 'image/jpeg', ext: 'jpeg' }],
  ['R0lGOD', { mime: 'image/gif', ext: 'gif' }],
];

/** Width of the text column in CSS px (6.5in); new images are never wider. */
export const COLUMN_PX = 624;
/** New images are kept under a page high (9in), leaving room for the line they sit on. */
const MAX_HEIGHT_PX = 800;
export const EMU_PER_PX = 9525;

/** The image type of base64 data, judged by its content (labels and extensions can lie). */
export function imageTypeOf(base64: string): ImageType | null {
  return BY_BASE64_PREFIX.find(([p]) => base64.startsWith(p))?.[1] ?? null;
}

/** Data URI for base64 image data, or null if it isn't a supported image. */
export function dataUriFor(base64: string): string | null {
  const type = imageTypeOf(base64);
  return type ? `data:${type.mime};base64,${base64}` : null;
}

/** The base64 part of a data URI ('' when it isn't one). */
export function base64Of(src: unknown): string {
  const s = String(src ?? '');
  const i = s.indexOf(';base64,');
  return s.startsWith('data:') && i > 0 ? s.slice(i + 8) : '';
}

export function bytesToBase64(bytes: Uint8Array): string {
  const chunks: string[] = [];
  // String.fromCharCode takes its arguments on the stack, so large images go in pieces.
  for (let i = 0; i < bytes.length; i += 0x8000) chunks.push(String.fromCharCode(...bytes.subarray(i, i + 0x8000)));
  return btoa(chunks.join(''));
}

export function base64ToBytes(base64: string): Uint8Array {
  const bin = atob(base64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

/** A pasted or parsed `src` as a canonical data URI whose label matches its bytes; null if unusable. */
export function normalizeSrc(src: string): string | null {
  const b64 = base64Of(src);
  if (!b64) return null;
  try {
    return dataUriFor(btoa(atob(b64)));
  } catch {
    return null;
  }
}

export interface PixelSize {
  width: number;
  height: number;
  /** Pixels per CSS px (2 for a Retina screenshot). */
  scale: number;
}

/** Pixel size read from the file header, so it works without a DOM. */
export function pixelSize(b: Uint8Array): PixelSize | null {
  const u16 = (i: number) => (b[i] << 8) | b[i + 1];
  const u32 = (i: number) => ((b[i] << 24) | (b[i + 1] << 16) | (b[i + 2] << 8) | b[i + 3]) >>> 0;
  if (b[0] === 0x89 && b[1] === 0x50 && b.length > 24) {
    let scale = 1;
    for (let i = 8; i + 12 <= b.length; ) {
      const len = u32(i);
      const type = String.fromCharCode(b[i + 4], b[i + 5], b[i + 6], b[i + 7]);
      // pHYs holds the resolution; macOS marks Retina screenshots as 144 dpi here.
      if (type === 'pHYs' && b[i + 16] === 1) {
        const dpi = u32(i + 8) * 0.0254;
        if (dpi > 110 && dpi < 400) scale = dpi / 72;
      }
      if (type === 'pHYs' || type === 'IDAT') break;
      i += 12 + len;
    }
    return { width: u32(16), height: u32(20), scale };
  }
  if (b[0] === 0x47 && b[1] === 0x49 && b.length > 10) return { width: b[6] | (b[7] << 8), height: b[8] | (b[9] << 8), scale: 1 };
  if (b[0] === 0xff && b[1] === 0xd8) {
    for (let i = 2; i + 9 < b.length; ) {
      if (b[i] !== 0xff) return null;
      const marker = b[i + 1];
      // SOF markers carry the size (C4, C8 and CC are other tables in the same range).
      if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
        return { width: u16(i + 7), height: u16(i + 5), scale: 1 };
      }
      i += 2 + u16(i + 2);
    }
  }
  return null;
}

/** EXIF orientation of a JPEG (1 = upright or unknown). */
export function jpegOrientation(b: Uint8Array): number {
  if (b[0] !== 0xff || b[1] !== 0xd8) return 1;
  for (let i = 2; i + 4 < b.length && b[i] === 0xff; ) {
    const len = (b[i + 2] << 8) | b[i + 3];
    if (b[i + 1] === 0xda) return 1;
    if (b[i + 1] === 0xe1 && b[i + 4] === 0x45 && b[i + 5] === 0x78) {
      const t = i + 10;
      const view = new DataView(b.buffer, b.byteOffset, b.byteLength);
      const le = b[t] === 0x49;
      if (t + 8 > b.length) return 1;
      const ifd = t + view.getUint32(t + 4, le);
      if (ifd + 2 > b.length) return 1;
      const n = view.getUint16(ifd, le);
      for (let k = 0; k < n; k++) {
        const e = ifd + 2 + k * 12;
        if (e + 12 > b.length) return 1;
        if (view.getUint16(e, le) === 0x0112) return view.getUint16(e + 8, le) || 1;
      }
      return 1;
    }
    i += 2 + len;
  }
  return 1;
}

/** On-page size (CSS px) for a newly added image: natural size, shrunk to fit the column and page. */
export function fitSize(p: PixelSize): { width: number; height: number } {
  const w = p.width / p.scale;
  const h = p.height / p.scale;
  const k = Math.min(1, COLUMN_PX / w, MAX_HEIGHT_PX / h);
  return { width: Math.max(1, Math.round(w * k)), height: Math.max(1, Math.round(h * k)) };
}

/** Default on-page size for a data URI. */
export function defaultSize(src: string): { width: number; height: number } {
  // Headers sit at the front, but JPEG metadata can be long, so read a generous slice.
  const p = pixelSize(base64ToBytes(base64Of(src).slice(0, 400_000)));
  return p && p.width && p.height ? fitSize(p) : { width: 300, height: 200 };
}
