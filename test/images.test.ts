import { DOMParser } from '@xmldom/xmldom';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { exportDocx, importDocx } from '../src/formats/docx';
import { fingerprint } from '../src/formats/fingerprint';
import { bytesToBase64, defaultSize, fitSize, jpegOrientation, normalizeSrc, pixelSize } from '../src/formats/images';
import { exportMarkdown, importMarkdown } from '../src/formats/markdown';
import { DEFAULT_META, type PMNode } from '../src/formats/model';
import { exportText } from '../src/formats/text';

const parseXml = (xml: string) => new DOMParser().parseFromString(xml, 'application/xml') as unknown as Document;
const read = (b: Uint8Array) => importDocx(b, parseXml);

/** A PNG header of the given size; `tag` makes the bytes distinct. */
function png(width: number, height: number, tag = 0): string {
  const b = new Uint8Array(60);
  b.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52]);
  new DataView(b.buffer).setUint32(16, width);
  new DataView(b.buffer).setUint32(20, height);
  b[50] = tag;
  return `data:image/png;base64,${bytesToBase64(b)}`;
}
const A = png(400, 200, 1);
const B = png(100, 100, 2);
const text = (t: string): PMNode => ({ type: 'text', text: t });
const img = (src: string, width: number, height: number, extra: Record<string, unknown> = {}): PMNode => ({ type: 'image', attrs: { src, width, height, alt: '', ...extra } });
const p = (...content: PMNode[]): PMNode => ({ type: 'paragraph', content });
const make = (...content: PMNode[]): PMNode => ({ type: 'doc', content });
function images(n: PMNode, out: PMNode[] = []): PMNode[] {
  if (n.type === 'image') out.push(n);
  for (const c of n.content ?? []) images(c, out);
  return out;
}
const media = (z: JSZip) => Object.keys(z.files).filter((f) => /media\/./.test(f));
const bodyOf = async (b: Uint8Array) => (await (await JSZip.loadAsync(b)).file('word/document.xml')!.async('string'));

/** A package whose document body is replaced, e.g. with a picture as Word writes it. */
async function withBody(base: Uint8Array, edit: (xml: string) => string): Promise<Uint8Array> {
  const z = await JSZip.loadAsync(base);
  z.file('word/document.xml', edit(await z.file('word/document.xml')!.async('string')));
  return z.generateAsync({ type: 'uint8array' });
}

describe('images in docx', () => {
  it('round-trips bytes, size, alt text and position among text and lists', async () => {
    const doc = make(
      p(text('before '), img(A, 300, 150, { alt: 'a "chart" & more' }), text(' after')),
      { type: 'bulletList', content: [{ type: 'listItem', content: [p(text('item '), img(B, 50, 50))] }] },
      { type: 'heading', attrs: { level: 1 }, content: [img(A, 10, 5)] },
    );
    const bytes = await exportDocx(doc, DEFAULT_META);
    const back = await read(bytes);
    expect(fingerprint(back.doc)).toEqual(fingerprint(doc));
    expect(images(back.doc).map((n) => n.attrs!.alt)).toEqual(['a "chart" & more', '', '']);
    expect(back.lossy).toEqual([]);
    // The same picture twice is stored once.
    expect(media(await JSZip.loadAsync(bytes)).length).toBe(2);
    expect(await (await JSZip.loadAsync(bytes)).file('[Content_Types].xml')!.async('string')).toContain('Extension="png"');
  });

  it('saving again reuses the picture verbatim instead of adding copies', async () => {
    const first = await exportDocx(make(p(img(A, 300, 150))), DEFAULT_META);
    const back = await read(first);
    const again = await exportDocx(back.doc, back.meta, first);
    expect(media(await JSZip.loadAsync(again))).toEqual(media(await JSZip.loadAsync(first)));
    expect((await bodyOf(again)).match(/<w:drawing[\s\S]*<\/w:drawing>/)![0]).toBe((await bodyOf(first)).match(/<w:drawing[\s\S]*<\/w:drawing>/)![0]);
  });

  it('resizing changes only the size of an existing picture', async () => {
    const first = await withBody(await exportDocx(make(p(img(A, 300, 150))), DEFAULT_META), (x) => x.replace('<pic:spPr>', '<pic:spPr bwMode="auto">'));
    const back = await read(first);
    Object.assign(images(back.doc)[0].attrs!, { width: 200, height: 100 });
    const out = await exportDocx(back.doc, back.meta, first);
    const xml = await bodyOf(out);
    expect(xml).toContain('bwMode="auto"');
    expect(xml.match(/cx="1905000" cy="952500"/g)!.length).toBe(2);
    expect(images((await read(out)).doc)[0].attrs).toMatchObject({ width: 200, height: 100 });
  });

  it('keeps complicated pictures untouched', async () => {
    const base = await exportDocx(make(p(img(A, 300, 150))), DEFAULT_META);
    const edits: [string, (x: string) => string][] = [
      ['floating', (x) => x.replace(/wp:inline/g, 'wp:anchor')],
      ['cropped', (x) => x.replace('<a:stretch>', '<a:srcRect l="10000"/><a:stretch>')],
      ['rotated', (x) => x.replace('<a:xfrm>', '<a:xfrm rot="5400000">')],
      ['effects', (x) => x.replace(/(<a:blip [^>]*)\/>/, '$1><a:grayscl/></a:blip>')],
      ['linked', (x) => x.replace(/(<wp:docPr [^>]*)\/>/, '$1><a:hlinkClick xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" r:id="x"/></wp:docPr>')],
      ['with a footnote', (x) => x.replace('</w:p>', '<w:r><w:footnoteReference w:id="2"/></w:r></w:p>')],
    ];
    for (const [name, edit] of edits) {
      const file = await withBody(base, edit);
      const back = await read(file);
      expect(back.doc.content![0].type, name).toBe('docxRaw');
      expect(back.lossy, name).toEqual([]);
      const para = (x: string) => x.match(/<w:p>[\s\S]*<\/w:p>/)![0];
      const out = await exportDocx(back.doc, back.meta, file);
      expect(para(await bodyOf(out)), name).toBe(para(await bodyOf(file)));
      expect(media(await JSZip.loadAsync(out)).length, name).toBe(1);
    }
  });

  it('writes a picture afresh when its source XML belongs to another file', async () => {
    const other = await read(await exportDocx(make(p(img(B, 40, 40))), DEFAULT_META));
    const home = await exportDocx(make(p(img(A, 300, 150))), DEFAULT_META);
    // Same relationship id, different picture: reusing the XML would show the wrong image.
    const out = await exportDocx(make(p(images(other.doc)[0])), DEFAULT_META, home);
    expect(images((await read(out)).doc).map((n) => n.attrs!.src)).toEqual([B]);
  });

  it('gives duplicates their own object ids', async () => {
    const first = await exportDocx(make(p(img(A, 300, 150))), DEFAULT_META);
    const back = await read(first);
    const n = images(back.doc)[0];
    const out = await exportDocx(make(p(n, n, img(B, 9, 9))), back.meta, first);
    const ids = [...(await bodyOf(out)).matchAll(/<wp:docPr id="(\d+)"/g)].map((m) => m[1]);
    expect(new Set(ids).size).toBe(3);
    expect(images((await read(out)).doc).length).toBe(3);
  });

  it('removes a deleted picture from the file, but not one a header uses', async () => {
    const first = await exportDocx(make(p(img(A, 300, 150)), p(img(B, 40, 40))), DEFAULT_META);
    const z = await JSZip.loadAsync(first);
    z.file('word/_rels/header1.xml.rels', '<Relationships><Relationship Id="r1" Type="x/image" Target="media/quire-image2.png"/></Relationships>');
    const file = await z.generateAsync({ type: 'uint8array' });
    const back = await read(file);
    const out = await JSZip.loadAsync(await exportDocx(make(p(text('gone'))), back.meta, file));
    expect(media(out)).toEqual(['word/media/quire-image2.png']);
    expect(await out.file('word/_rels/document.xml.rels')!.async('string')).not.toContain('media/');
  });

  it('fingerprints notice a lost, changed or resized picture', () => {
    const f = (n: PMNode) => fingerprint(make(p(text('x'), n))).text;
    expect(new Set([f(img(A, 1, 1)), f(img(B, 1, 1)), f(img(A, 2, 1)), f(text('y'))]).size).toBe(4);
    expect(fingerprint(make(p(text('two words'), img(A, 1, 1)))).words).toBe(2);
  });
});

describe('images in other formats', () => {
  it('round-trips through markdown with their size', () => {
    const doc = make(p(text('see '), img(A, 123, 61, { alt: 'a [b] *c*' }), text(' ok')), p(img(B, 40, 40)));
    const back = importMarkdown(exportMarkdown(doc)).doc;
    expect(images(back).map((n) => n.attrs)).toEqual(images(doc).map((n) => n.attrs));
  });

  it('markdown images that point at files keep their alt text, as before', () => {
    expect(importMarkdown('![cat](cat.png)').doc.content![0].content).toMatchObject([{ type: 'text', text: 'cat' }]);
  });

  it('become a note in plain text', () => {
    expect(exportText(make(p(img(A, 1, 1, { alt: 'cat' }))))).toBe('[Image: cat]\n');
  });
});

describe('image helpers', () => {
  it('reads sizes and fits them to the column', () => {
    expect(defaultSize(A)).toEqual({ width: 400, height: 200 });
    expect(fitSize({ width: 2000, height: 1000, scale: 1 })).toEqual({ width: 624, height: 312 });
    expect(fitSize({ width: 800, height: 400, scale: 2 })).toEqual({ width: 400, height: 200 });
    expect(fitSize({ width: 100, height: 4000, scale: 1 }).height).toBe(800);
    expect(pixelSize(new Uint8Array([1, 2, 3]))).toBeNull();
    expect(jpegOrientation(new Uint8Array([0xff, 0xd8, 0xff, 0xda, 0, 0]))).toBe(1);
  });

  it('relabels sources by content and rejects anything else', () => {
    expect(normalizeSrc(A.replace('image/png', 'image/jpeg'))).toBe(A);
    expect(normalizeSrc('data:image/svg+xml;base64,PHN2Zz4=')).toBeNull();
    expect(normalizeSrc('data:image/png;base64,iVBORw0KGgo!!')).toBeNull();
    expect(normalizeSrc('https://example.com/a.png')).toBeNull();
  });
});
