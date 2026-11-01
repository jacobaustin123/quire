import * as fs from 'node:fs';
import { DOMParser } from '@xmldom/xmldom';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { exportDocx, importDocx } from '../src/formats/docx';
import type { DocMeta, PMMark, PMNode } from '../src/formats/model';

/** Optional real-world document to round-trip (e.g. a large personal file kept out of the repo). */
const REAL_DOCX = process.env.QUIRE_TEST_DOCX ?? '';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const META: DocMeta = { baseFont: 'Times New Roman', baseSize: 12, paragraphSpacing: 0, lineSpacing: 1 };

/** Parses XML with xmldom, throwing on any parse error (used as the converter's parseXml). */
function parseXml(xml: string): Document {
  const errors: string[] = [];
  const doc = new DOMParser({
    onError: (level: string, msg: string) => {
      if (level !== 'warning') errors.push(msg);
    },
  }).parseFromString(xml, 'text/xml');
  if (errors.length) throw new Error('XML parse error: ' + errors.join('; '));
  return doc as unknown as Document;
}

interface PackageParts {
  body: string;
  styles?: string;
  numbering?: string;
  theme?: string;
  links?: Record<string, string>;
  sectPr?: string;
  extraRels?: string[];
  files?: Record<string, string | Uint8Array>;
}

/** Builds a minimal .docx from hand-written XML so import can be tested independently of export. */
async function makeDocx(p: PackageParts): Promise<Uint8Array> {
  const zip = new JSZip();
  const overrides = [
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
  ];
  const rels: string[] = [];
  if (p.styles !== undefined) {
    zip.file('word/styles.xml', `<w:styles xmlns:w="${W}">${p.styles}</w:styles>`);
    rels.push(`<Relationship Id="rS" Type="${R}/styles" Target="styles.xml"/>`);
  }
  if (p.numbering !== undefined) {
    zip.file('word/numbering.xml', `<w:numbering xmlns:w="${W}">${p.numbering}</w:numbering>`);
    rels.push(`<Relationship Id="rN" Type="${R}/numbering" Target="numbering.xml"/>`);
  }
  if (p.theme !== undefined) {
    zip.file('word/theme/theme1.xml', p.theme);
    rels.push(`<Relationship Id="rT" Type="${R}/theme" Target="theme/theme1.xml"/>`);
  }
  rels.push(...(p.extraRels ?? []));
  for (const [name, data] of Object.entries(p.files ?? {})) zip.file(name, data);
  for (const [id, href] of Object.entries(p.links ?? {})) {
    rels.push(`<Relationship Id="${id}" Type="${R}/hyperlink" Target="${href.replace(/&/g, '&amp;')}" TargetMode="External"/>`);
  }
  zip.file(
    '[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' + overrides.join('') + '</Types>',
  );
  zip.file(
    '_rels/.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.file(
    'word/_rels/document.xml.rels',
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`,
  );
  zip.file(
    'word/document.xml',
    `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ` +
      `xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ${DRAWING_NS}><w:body>${p.body}${p.sectPr ?? ''}</w:body></w:document>`,
  );
  return zip.generateAsync({ type: 'uint8array' });
}

const DRAWING_NS =
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';

const BASE_STYLES =
  '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/></w:rPr></w:rPrDefault>' +
  '<w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="480" w:lineRule="exact"/></w:pPr></w:pPrDefault></w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:spacing w:line="276"/></w:pPr><w:rPr><w:sz w:val="24"/></w:rPr></w:style>';

const p = (inner: string, pPr = ''): string => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${inner}</w:p>`;
const r = (text: string, rPr = ''): string =>
  `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;

async function load(parts: PackageParts): Promise<PMNode[]> {
  const res = await importDocx(await makeDocx({ styles: BASE_STYLES, ...parts }), parseXml);
  return res.doc.content ?? [];
}

function txt(text: string, ...marks: PMMark[]): PMNode {
  return marks.length ? { type: 'text', text, marks } : { type: 'text', text };
}

/** Canonical form for equality: default attrs dropped, marks sorted, adjacent texts merged. */
function normalize(n: PMNode): PMNode {
  const out: PMNode = { type: n.type };
  if (n.text !== undefined) out.text = n.text;
  if (n.attrs) {
    const a: Record<string, unknown> = {};
    for (const [k, v] of Object.entries(n.attrs)) {
      const isDefault =
        v === null || v === undefined || (k === 'textAlign' && v === 'left') || (k === 'indent' && v === 0) || (k === 'start' && v === 1);
      if (!isDefault) a[k] = v;
    }
    if (Object.keys(a).length) out.attrs = a;
  }
  if (n.marks?.length) {
    out.marks = [...n.marks].sort((x, y) => x.type.localeCompare(y.type)).map((m) => (m.attrs ? { type: m.type, attrs: m.attrs } : { type: m.type }));
  }
  if (n.content?.length) {
    const kids: PMNode[] = [];
    for (const c of n.content.map(normalize)) {
      const prev = kids[kids.length - 1];
      if (prev && prev.type === 'text' && c.type === 'text' && JSON.stringify(prev.marks) === JSON.stringify(c.marks)) {
        prev.text += c.text ?? '';
      } else kids.push(c);
    }
    out.content = kids;
  }
  return out;
}

async function unzipText(bytes: Uint8Array, path: string): Promise<string> {
  const zip = await JSZip.loadAsync(bytes);
  const f = zip.file(path);
  if (!f) throw new Error('missing part ' + path);
  return f.async('string');
}

/** Asserts every XML part in the package parses, and every r:id in document.xml resolves. */
async function assertWellFormed(bytes: Uint8Array): Promise<void> {
  const zip = await JSZip.loadAsync(bytes);
  for (const name of Object.keys(zip.files)) {
    if (!/\.(xml|rels)$/.test(name)) continue;
    const xml = await zip.file(name)!.async('string');
    expect(() => parseXml(xml), name).not.toThrow();
  }
  const doc = await unzipText(bytes, 'word/document.xml');
  const rels = await unzipText(bytes, 'word/_rels/document.xml.rels');
  const ids = new Set([...rels.matchAll(/Id="([^"]+)"/g)].map((m) => m[1]));
  for (const m of doc.matchAll(/r:id="([^"]+)"/g)) expect(ids.has(m[1]), m[1]).toBe(true);
  const numbering = zip.file('word/numbering.xml') ? await unzipText(bytes, 'word/numbering.xml') : '';
  const numIds = new Set([...numbering.matchAll(/<w:num w:numId="(\d+)"/g)].map((m) => m[1]));
  for (const m of doc.matchAll(/<w:numId w:val="(\d+)"/g)) if (m[1] !== '0') expect(numIds.has(m[1]), 'numId ' + m[1]).toBe(true);
}

async function roundTrip(doc: PMNode, meta: DocMeta = META, original?: Uint8Array): Promise<PMNode> {
  const bytes = await exportDocx(doc, meta, original);
  await assertWellFormed(bytes);
  return (await importDocx(bytes, parseXml)).doc;
}

describe('importDocx: runs and marks', () => {
  it('maps basic marks and merges adjacent identical runs', async () => {
    const c = await load({
      body: p(
        r('a', '<w:b/>') + r('b', '<w:b/><w:bCs/>') + r('c', '<w:i/>') + r('d', '<w:strike/>') +
          r('e', '<w:u w:val="single"/>') + r('f', '<w:vertAlign w:val="superscript"/>') + r('g', '<w:vertAlign w:val="subscript"/>'),
      ),
    });
    expect(c).toEqual([
      {
        type: 'paragraph',
        content: [
          txt('ab', { type: 'bold' }),
          txt('c', { type: 'italic' }),
          txt('d', { type: 'strike' }),
          txt('e', { type: 'underline' }),
          txt('f', { type: 'superscript' }),
          txt('g', { type: 'subscript' }),
        ],
      },
    ]);
  });

  it('emits textStyle only for fonts/sizes that differ from the base', async () => {
    const res = await importDocx(
      await makeDocx({
        styles: BASE_STYLES,
        body: p(r('same', '<w:rFonts w:ascii="Times New Roman"/><w:sz w:val="24"/>') + r('big', '<w:sz w:val="28"/>') + r('arial', '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="21"/>')),
      }),
      parseXml,
    );
    expect(res.meta).toEqual({ baseFont: 'Times New Roman', baseSize: 12, paragraphSpacing: 6, lineSpacing: 1.15 });
    expect(res.doc.content![0].content).toEqual([
      txt('same'),
      txt('big', { type: 'textStyle', attrs: { fontFamily: null, fontSize: '14' } }),
      txt('arial', { type: 'textStyle', attrs: { fontFamily: 'Arial', fontSize: '10.5' } }),
    ]);
  });

  it('does not treat <w:u> without w:val as underline; w:val="none" cancels a style underline', async () => {
    const c = await load({
      styles: BASE_STYLES + '<w:style w:type="character" w:styleId="U"><w:name w:val="U"/><w:rPr><w:u w:val="single"/></w:rPr></w:style>',
      body: p(r('a', '<w:u w:color="000000"/>') + r('b', '<w:rStyle w:val="U"/><w:u w:val="none"/>') + r('c', '<w:rStyle w:val="U"/>')),
    });
    expect(c[0].content).toEqual([txt('ab'), txt('c', { type: 'underline' })]);
  });

  it('honors toggle values 0/false/off', async () => {
    const c = await load({
      styles: BASE_STYLES + '<w:style w:type="paragraph" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:b/><w:i/></w:rPr></w:style>',
      body: p(r('a', '<w:b w:val="0"/>') + r('b', '<w:b w:val="false"/><w:i w:val="off"/>') + r('c'), '<w:pStyle w:val="Strong"/>'),
    });
    expect(c[0].content).toEqual([txt('a', { type: 'italic' }), txt('b'), txt('c', { type: 'bold' }, { type: 'italic' })]);
  });

  it('keeps per-paragraph spacing that differs from the document default, and round-trips it', async () => {
    const body = p(r('a')) + p(r('b'), '<w:spacing w:after="0"/>') + p(r('c'), '<w:spacing w:before="240" w:line="360" w:lineRule="auto"/>');
    const c = await load({ styles: BASE_STYLES, body });
    expect(c.map((n) => n.attrs)).toEqual([undefined, { spaceAfter: 0 }, { spaceBefore: 12, lineSpacing: 1.5 }]);
    const doc = { type: 'doc', content: c };
    // Same document defaults as BASE_STYLES (6 pt after, 1.15 lines), so the overrides stay overrides.
    const again = await importDocx(await exportDocx(doc, { ...META, paragraphSpacing: 6, lineSpacing: 1.15 }), parseXml);
    expect(again.doc.content!.map((n) => n.attrs)).toEqual([undefined, { spaceAfter: 0 }, { spaceBefore: 12, lineSpacing: 1.5 }]);
  });

  it('resolves basedOn chains for paragraph and character styles', async () => {
    const styles =
      BASE_STYLES +
      '<w:style w:type="paragraph" w:styleId="Base"><w:name w:val="Base"/><w:pPr><w:jc w:val="center"/></w:pPr><w:rPr><w:sz w:val="32"/></w:rPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Child"><w:name w:val="Child"/><w:basedOn w:val="Base"/><w:pPr><w:ind w:left="1440"/></w:pPr><w:rPr><w:i/></w:rPr></w:style>' +
      '<w:style w:type="character" w:styleId="C1"><w:name w:val="C1"/><w:rPr><w:rFonts w:ascii="Georgia"/></w:rPr></w:style>' +
      '<w:style w:type="character" w:styleId="C2"><w:name w:val="C2"/><w:basedOn w:val="C1"/><w:rPr><w:b/></w:rPr></w:style>';
    const c = await load({ styles, body: p(r('x') + r('y', '<w:rStyle w:val="C2"/>'), '<w:pStyle w:val="Child"/>') });
    expect(c).toEqual([
      {
        type: 'paragraph',
        // Child isn't based on Normal, so it doesn't inherit Normal's 1.15 lines (Word agrees).
        attrs: { textAlign: 'center', indent: 2, lineSpacing: 1 },
        content: [
          txt('x', { type: 'italic' }, { type: 'textStyle', attrs: { fontFamily: null, fontSize: '16' } }),
          txt('y', { type: 'bold' }, { type: 'italic' }, { type: 'textStyle', attrs: { fontFamily: 'Georgia', fontSize: '16' } }),
        ],
      },
    ]);
  });

  it('resolves theme fonts', async () => {
    const theme =
      '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme name="x">' +
      '<a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont>' +
      '</a:fontScheme></a:themeElements></a:theme>';
    const res = await importDocx(
      await makeDocx({
        theme,
        styles:
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>',
        body: p(r('a') + r('b', '<w:rFonts w:asciiTheme="majorHAnsi"/>')),
      }),
      parseXml,
    );
    expect(res.meta).toEqual({ baseFont: 'Calibri', baseSize: 11, paragraphSpacing: 0, lineSpacing: 1 });
    expect(res.doc.content![0].content).toEqual([
      txt('a'),
      txt('b', { type: 'textStyle', attrs: { fontFamily: 'Calibri Light', fontSize: null } }),
    ]);
  });

  it('handles tabs, breaks, hyphens and symbols', async () => {
    const c = await load({
      body: p(
        '<w:r><w:t>a</w:t><w:tab/><w:t>b</w:t><w:br/><w:t>c</w:t><w:noBreakHyphen/><w:t>d</w:t><w:softHyphen/><w:t>e</w:t>' +
          '<w:br w:type="page"/><w:sym w:font="Symbol" w:char="03B1"/><w:lastRenderedPageBreak/></w:r><w:proofErr w:type="spellStart"/>',
      ),
    });
    expect(c[0].content).toEqual([txt('a\tb'), { type: 'hardBreak' }, txt('c-de'), { type: 'hardBreak' }, txt('α')]);
  });

  it('keeps $ as literal text', async () => {
    const c = await load({ body: p(r('It cost $5 and $10, or $$3$$.')) });
    expect(c[0].content).toEqual([txt('It cost $5 and $10, or $$3$$.')]);
  });

  it('keeps field results, drops instructions, links HYPERLINK fields', async () => {
    const field = (instr: string, result: string): string =>
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">' + instr + '</w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' + r(result) + '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
    const c = await load({
      body: p(r('Page ') + field(' PAGE ', '3')) + p(field(' HYPERLINK "https://x.test/" ', 'site')) +
        p('<w:fldSimple w:instr=" DATE "><w:r><w:t>today</w:t></w:r></w:fldSimple>'),
    });
    expect(c).toEqual([
      { type: 'paragraph', content: [txt('Page 3')] },
      { type: 'paragraph', content: [txt('site', { type: 'link', attrs: { href: 'https://x.test/' } })] },
      { type: 'paragraph', content: [txt('today')] },
    ]);
  });

  it('maps w:hyperlink via rels and suppresses Hyperlink-style underline', async () => {
    const c = await load({
      styles: BASE_STYLES + '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:u w:val="single"/></w:rPr></w:style>',
      links: { rId9: 'https://example.com/?a=1&b=2' },
      body: p(
        r('see ') + '<w:hyperlink r:id="rId9"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:b/></w:rPr><w:t>here</w:t></w:r></w:hyperlink>' +
          '<w:hyperlink w:anchor="sec1"><w:r><w:t>sec</w:t></w:r></w:hyperlink>',
      ),
    });
    expect(c[0].content).toEqual([
      txt('see '),
      txt('here', { type: 'link', attrs: { href: 'https://example.com/?a=1&b=2' } }, { type: 'bold' }),
      txt('sec', { type: 'link', attrs: { href: '#sec1' } }),
    ]);
  });

  it('recurses into ins/smartTag/sdt/AlternateContent, skips del/notes, flattens oMath', async () => {
    const c = await load({
      body: p(
        '<w:ins>' + r('a') + '</w:ins><w:del><w:r><w:delText>gone</w:delText></w:r></w:del><w:smartTag>' + r('b') + '</w:smartTag>' +
          '<w:sdt><w:sdtPr/><w:sdtContent>' + r('c') + '</w:sdtContent></w:sdt>' +
          '<w:r><w:footnoteReference w:id="1"/></w:r>' +
          '<mc:AlternateContent><mc:Choice Requires="x">' + r('no') + '</mc:Choice><mc:Fallback>' + r('d') + '</mc:Fallback></mc:AlternateContent>' +
          '<m:oMath><m:r><m:t>x+1</m:t></m:r></m:oMath>',
      ),
    });
    expect(c[0].content).toEqual([txt('abcdx+1')]);
  });
});

describe('importDocx: blocks', () => {
  it('maps alignment, indent, outline and headings', async () => {
    const styles =
      BASE_STYLES +
      '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/></w:style>' +
      '<w:style w:type="paragraph" w:styleId="H5"><w:name w:val="heading 5"/><w:basedOn w:val="Normal"/></w:style>';
    const c = await load({
      styles,
      body:
        p(r('c'), '<w:jc w:val="center"/>') + p(r('r'), '<w:jc w:val="end"/><w:ind w:left="1500"/>') + p(r('j'), '<w:jc w:val="both"/>') +
        p(r('o', '<w:b/>'), '<w:outlineLvl w:val="0"/>') + p(r('h'), '<w:pStyle w:val="Heading2"/>') + p(r('t'), '<w:pStyle w:val="Title"/>') +
        p(r('5'), '<w:pStyle w:val="H5"/>') + p(''),
    });
    expect(c).toEqual([
      { type: 'paragraph', attrs: { textAlign: 'center' }, content: [txt('c')] },
      { type: 'paragraph', attrs: { textAlign: 'right', indent: 2 }, content: [txt('r')] },
      { type: 'paragraph', attrs: { textAlign: 'justify' }, content: [txt('j')] },
      { type: 'paragraph', attrs: { outline: 0 }, content: [txt('o', { type: 'bold' })] },
      { type: 'heading', attrs: { level: 2 }, content: [txt('h')] },
      { type: 'heading', attrs: { level: 1 }, content: [txt('t')] },
      { type: 'heading', attrs: { level: 3 }, content: [txt('5')] },
      { type: 'paragraph' },
    ]);
  });

  it('turns empty bottom-bordered paragraphs into horizontal rules', async () => {
    const c = await load({
      body:
        p('', '<w:pBdr><w:bottom w:val="single" w:sz="6"/></w:pBdr>') + p('', '<w:pBdr><w:bottom w:val="nil"/></w:pBdr>') +
        p(r('text'), '<w:pBdr><w:bottom w:val="single"/></w:pBdr>'),
    });
    expect(c.map((n) => n.type)).toEqual(['horizontalRule', 'paragraph', 'paragraph']);
  });

  it('passes tables, object paragraphs and object sdts through verbatim as docxRaw', async () => {
    const tbl = (rows: string[][]): string =>
      `<w:tbl><w:tblPr/>${rows.map((cells) => `<w:tr>${cells.map((x) => `<w:tc><w:tcPr/>${x}</w:tc>`).join('')}</w:tr>`).join('')}</w:tbl>`;
    const empty = tbl([[p(''), '<w:p/>']]);
    const full = tbl([[p(r('A')) + p(r('A2')), p(r('B'))], [p(r('C')), '<w:p/>']]);
    const nested = tbl([[tbl([[p(r('in'))]]) + '<w:p/>']]);
    const img = p(r('cap ') + '<w:r><w:drawing><wp:inline><a:graphic/></wp:inline></w:drawing></w:r>');
    const alt = p('<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><w:txbxContent><w:p><w:r><w:t>box</w:t></w:r></w:p></w:txbxContent></w:drawing></mc:Choice><mc:Fallback><w:pict><w:p><w:r><w:t>box</w:t></w:r></w:p></w:pict></mc:Fallback></mc:AlternateContent></w:r>');
    const sdt = '<w:sdt><w:sdtPr/><w:sdtContent>' + full + '</w:sdtContent></w:sdt>';
    const plainSdt = '<w:sdt><w:sdtContent>' + p(r('plain')) + '</w:sdtContent></w:sdt>';
    const c = await load({ body: p(r('before')) + empty + '<!-- c -->' + full + nested + img + alt + sdt + plainSdt + p(r('after')) });
    expect(c).toEqual([
      { type: 'paragraph', content: [txt('before')] },
      { type: 'docxRaw', attrs: { xml: empty, kind: 'table', text: '' } },
      { type: 'docxRaw', attrs: { xml: full, kind: 'table', text: 'A A2\tB\nC' } },
      { type: 'docxRaw', attrs: { xml: nested, kind: 'table', text: 'in' } },
      { type: 'docxRaw', attrs: { xml: img, kind: 'image', text: 'cap' } },
      { type: 'docxRaw', attrs: { xml: alt, kind: 'image', text: 'box' } },
      { type: 'docxRaw', attrs: { xml: sdt, kind: 'other', text: 'A A2\tB\nC' } },
      { type: 'paragraph', content: [txt('plain')] },
      { type: 'paragraph', content: [txt('after')] },
    ]);
  });

  const NUMBERING =
    '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/></w:lvl>' +
    '<w:lvl w:ilvl="1"><w:numFmt w:val="bullet"/></w:lvl><w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/></w:lvl></w:abstractNum>' +
    '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl></w:abstractNum>' +
    '<w:abstractNum w:abstractNumId="2"><w:numStyleLink w:val="LinkedList"/></w:abstractNum>' +
    '<w:abstractNum w:abstractNumId="3"><w:styleLink w:val="LinkedList"/><w:lvl w:ilvl="0"><w:start w:val="4"/><w:numFmt w:val="lowerLetter"/></w:lvl></w:abstractNum>' +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>' +
    '<w:num w:numId="3"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="7"/></w:lvlOverride></w:num>' +
    '<w:num w:numId="4"><w:abstractNumId w:val="2"/></w:num><w:num w:numId="5"><w:abstractNumId w:val="3"/></w:num>';
  const LIST_STYLES =
    BASE_STYLES +
    '<w:style w:type="numbering" w:styleId="LinkedList"><w:name w:val="LinkedList"/><w:pPr><w:numPr><w:numId w:val="5"/></w:numPr></w:pPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Bulleted"><w:name w:val="Bulleted"/><w:basedOn w:val="Normal"/><w:pPr><w:numPr><w:numId w:val="2"/></w:numPr><w:ind w:left="720"/></w:pPr></w:style>';
  const li = (text: string, numId: number, ilvl = 0): string =>
    p(r(text), `<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>`);
  const item = (text: string, ...rest: PMNode[]): PMNode => ({ type: 'listItem', content: [{ type: 'paragraph', content: [txt(text)] }, ...rest] });

  it('builds nested mixed lists', async () => {
    const c = await load({
      styles: LIST_STYLES,
      numbering: NUMBERING,
      body: li('1', 1) + li('a', 1, 1) + li('i', 1, 2) + li('b', 1, 1) + li('2', 1) + li('x', 1, 2) + p(r('end')),
    });
    expect(c).toEqual([
      {
        type: 'orderedList',
        content: [
          item('1', { type: 'bulletList', content: [item('a', { type: 'orderedList', content: [item('i')] }), item('b')] }),
          item('2', { type: 'orderedList', content: [item('x')] }),
        ],
      },
      { type: 'paragraph', content: [txt('end')] },
    ]);
  });

  it('keeps the level of a list that starts below level 0', async () => {
    const c = await load({ styles: LIST_STYLES, numbering: NUMBERING, body: p(r('x')) + li('a', 1, 1) + li('i', 1, 2) + li('b', 1, 1) + li('1', 1) });
    expect(c.slice(1)).toEqual([
      { type: 'bulletList', attrs: { level: 1 }, content: [item('a', { type: 'orderedList', content: [item('i')] }), item('b')] },
      { type: 'orderedList', content: [item('1')] },
    ]);
  });

  it('starts a new list on type change at the same level', async () => {
    const c = await load({ styles: LIST_STYLES, numbering: NUMBERING, body: li('1', 1) + li('b', 2) + li('2', 1) });
    expect(c.map((n) => [n.type, n.attrs?.start ?? 1])).toEqual([['orderedList', 1], ['bulletList', 1], ['orderedList', 2]]);
  });

  it('continues numbering across intervening paragraphs and honors startOverride / style numPr / numStyleLink', async () => {
    const c = await load({
      styles: LIST_STYLES,
      numbering: NUMBERING,
      body:
        li('1', 1) + li('2', 1) + p(r('gap')) + li('3', 1) + li('7', 3) + p(r('styled'), '<w:pStyle w:val="Bulleted"/>') +
        p(r('off'), '<w:pStyle w:val="Bulleted"/><w:numPr><w:numId w:val="0"/></w:numPr>') + li('d', 4) + li('e', 4),
    });
    expect(c.map((n) => [n.type, n.attrs?.start ?? 1, n.content?.length])).toEqual([
      ['orderedList', 1, 2],
      ['paragraph', 1, 1],
      ['orderedList', 3, 1],
      ['orderedList', 7, 1],
      ['bulletList', 1, 1],
      ['paragraph', 1, 1],
      ['orderedList', 4, 2],
    ]);
    // List paragraphs never carry an indent attr, even when their style indents.
    expect(c[4].content![0].content![0].attrs).toBeUndefined();
    expect(c[5].attrs).toEqual({ indent: 1 });
  });

  it('turns QuireMath runs into math nodes', async () => {
    const styles = BASE_STYLES + '<w:style w:type="character" w:styleId="QuireMath"><w:name w:val="Quire Math"/></w:style>';
    const m = (t: string): string => r(t, '<w:rStyle w:val="QuireMath"/>');
    const c = await load({ styles, body: p(r('x = ') + m('$a^2$') + r(' ok')) + p(m('$$\\int f$$'), '<w:jc w:val="center"/>') + p(m('$$') + m('x$$')) });
    expect(c).toEqual([
      { type: 'paragraph', content: [txt('x = '), { type: 'mathInline', attrs: { latex: 'a^2' } }, txt(' ok')] },
      { type: 'mathBlock', attrs: { latex: '\\int f' } },
      { type: 'mathBlock', attrs: { latex: 'x' } },
    ]);
  });
});

const RICH_DOC: PMNode = {
  type: 'doc',
  content: [
    { type: 'heading', attrs: { level: 1 }, content: [txt('Title')] },
    { type: 'heading', attrs: { level: 3, textAlign: 'center' }, content: [txt('Sub', { type: 'italic' })] },
    {
      type: 'paragraph',
      attrs: { textAlign: 'justify', indent: 2, outline: 1 },
      content: [
        txt('Plain $5 & <tags> '),
        txt('bold', { type: 'bold' }),
        txt('\tlink', { type: 'link', attrs: { href: 'https://a.test/?q=1&r=2' } }, { type: 'underline' }),
        txt(' again', { type: 'link', attrs: { href: 'https://a.test/?q=1&r=2' } }),
        { type: 'hardBreak' },
        txt('big', { type: 'textStyle', attrs: { fontFamily: 'Arial', fontSize: '14' } }),
        txt('small', { type: 'textStyle', attrs: { fontFamily: null, fontSize: '9.5' } }),
        txt('x', { type: 'superscript' }, { type: 'strike' }),
        txt('2', { type: 'subscript' }),
        { type: 'mathInline', attrs: { latex: 'e^{i\\pi}+1=0' } },
        txt('ctrl\u0001char'),
      ],
    },
    { type: 'horizontalRule' },
    { type: 'mathBlock', attrs: { latex: '\\sum_{n=1}^\\infty \\frac{1}{n^2}' } },
    {
      type: 'orderedList',
      attrs: { start: 3 },
      content: [
        {
          type: 'listItem',
          content: [
            { type: 'paragraph', content: [txt('three')] },
            { type: 'paragraph', content: [txt('continuation')] },
            {
              type: 'bulletList',
              content: [
                { type: 'listItem', content: [{ type: 'paragraph', content: [txt('dot')] }, { type: 'orderedList', content: [{ type: 'listItem', content: [{ type: 'paragraph' }] }] }] },
              ],
            },
            { type: 'orderedList', attrs: { start: 5 }, content: [{ type: 'listItem', content: [{ type: 'paragraph', content: [txt('mismatch')] }] }] },
          ],
        },
        { type: 'listItem', content: [{ type: 'paragraph', attrs: { textAlign: 'right' }, content: [txt('four')] }] },
      ],
    },
    { type: 'bulletList', content: [{ type: 'listItem', content: [{ type: 'paragraph', content: [txt('b1')] }] }] },
    { type: 'docxRaw', attrs: { xml: '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>', kind: 'table', text: 'cell' } },
    { type: 'bulletList', attrs: { level: 1 }, content: [{ type: 'listItem', content: [{ type: 'paragraph', content: [txt('sub')] }] }] },
    { type: 'orderedList', content: [{ type: 'listItem', content: [{ type: 'paragraph', content: [txt('fresh 1')] }] }] },
    { type: 'paragraph' },
  ],
};

function expectedRich(): PMNode {
  const d = normalize(RICH_DOC);
  // The XML-invalid control char is stripped; underline inside a link is not representable.
  const para = d.content![2];
  para.content = para.content!.map((n) => (n.text === 'ctrl\u0001char' ? { ...n, text: 'ctrlchar' } : n));
  const link = para.content!.find((n) => n.text === '\tlink')!;
  link.marks = link.marks!.filter((m) => m.type !== 'underline');
  return normalize(d);
}

describe('exportDocx', () => {
  it('round-trips a rich document through a fresh package', async () => {
    const back = await roundTrip(RICH_DOC);
    expect(normalize(back)).toEqual(expectedRich());
  });

  it('writes meta into docDefaults and creates a complete minimal package', async () => {
    const meta: DocMeta = { baseFont: 'Charter', baseSize: 11, paragraphSpacing: 6, lineSpacing: 1.5 };
    const bytes = await exportDocx(RICH_DOC, meta);
    await assertWellFormed(bytes);
    const zip = await JSZip.loadAsync(bytes);
    for (const part of ['[Content_Types].xml', '_rels/.rels', 'word/document.xml', 'word/styles.xml', 'word/numbering.xml', 'word/settings.xml', 'word/_rels/document.xml.rels', 'docProps/core.xml', 'docProps/app.xml']) {
      expect(zip.file(part), part).not.toBeNull();
    }
    const res = await importDocx(bytes, parseXml);
    expect(res.meta).toEqual(meta);
    const doc = await unzipText(bytes, 'word/document.xml');
    expect(doc).toContain('<w:pgSz w:w="12240" w:h="15840"/>');
    expect(doc).not.toContain('\u0001');
    // pPr children in schema order.
    expect(doc).toMatch(/<w:pPr><w:pStyle w:val="Heading3"\/><w:jc w:val="center"\/><\/w:pPr>/);
    expect(doc).toMatch(/<w:ind w:left="1440"\/><w:jc w:val="both"\/><w:outlineLvl w:val="1"\/>/);
    expect(doc).toContain('<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>');
  });

  it('round-trips into an original package, keeping its parts and sectPr', async () => {
    const original = await makeDocx({
      styles: BASE_STYLES + '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/></w:pPr></w:style>',
      links: { rId5: 'https://old.test/' },
      body: p('<w:hyperlink r:id="rId5"><w:r><w:t>old</w:t></w:r></w:hyperlink>'),
      sectPr: '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>',
    });
    const loaded = await importDocx(original, parseXml);
    const bytes = await exportDocx(RICH_DOC, loaded.meta, loaded.original);
    await assertWellFormed(bytes);
    expect(normalize((await importDocx(bytes, parseXml)).doc)).toEqual(expectedRich());
    const doc = await unzipText(bytes, 'word/document.xml');
    expect(doc).toContain('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body>');
    expect(doc).toContain('xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"');
    expect(doc).toContain('<w:pStyle w:val="ListParagraph"/><w:numPr>');
    const rels = await unzipText(bytes, 'word/_rels/document.xml.rels');
    expect(rels).not.toContain('old.test');
    expect(rels).toContain('https://a.test/?q=1&amp;r=2');
    // Numbering was absent from the original: created, related and registered.
    expect(rels).toMatch(/relationships\/numbering" Target="numbering.xml"/);
    expect(await unzipText(bytes, '[Content_Types].xml')).toContain('PartName="/word/numbering.xml"');
    const styles = await unzipText(bytes, 'word/styles.xml');
    for (const id of ['Heading1', 'Heading2', 'Heading3', 'Hyperlink', 'QuireMath']) expect(styles).toContain(`w:styleId="${id}"`);
  });

  it('round-trips docxRaw blocks byte-identically, keeping their rels and numbering', async () => {
    const IMG_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
    const table =
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tr><w:tc><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>' +
      '<w:hyperlink r:id="rId5"><w:r><w:t>in &amp; table</w:t></w:r></w:hyperlink></w:p></w:tc></w:tr></w:tbl>';
    const image =
      '<w:p><w:r><w:drawing><wp:inline distT="0"><wp:extent cx="10" cy="10"/><wp:docPr id="1" name="P"/><a:graphic>' +
      '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rId6"/>' +
      '</pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
    const original = await makeDocx({
      styles: BASE_STYLES,
      numbering: '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>',
      links: { rId5: 'https://in-table.test/', rId7: 'https://unused.test/' },
      extraRels: [`<Relationship Id="rId6" Type="${IMG_REL}" Target="media/image1.png"/>`],
      files: { 'word/media/image1.png': new Uint8Array([137, 80, 78, 71]) },
      body: p(r('top')) + table + '\n  ' + image + p('<w:hyperlink r:id="rId7"><w:r><w:t>old</w:t></w:r></w:hyperlink>'),
    });
    const loaded = await importDocx(original, parseXml);
    expect(loaded.doc.content!.map((n) => n.type)).toEqual(['paragraph', 'docxRaw', 'docxRaw', 'paragraph']);
    loaded.doc.content!.pop(); // the paragraph linking to unused.test
    loaded.doc.content!.push({ type: 'paragraph', content: [txt('new', { type: 'link', attrs: { href: 'https://new.test/' } })] });
    const bytes = await exportDocx(loaded.doc, loaded.meta, loaded.original);
    await assertWellFormed(bytes);

    const doc = await unzipText(bytes, 'word/document.xml');
    expect(doc).toContain(table + image);
    const rels = await unzipText(bytes, 'word/_rels/document.xml.rels');
    expect(rels).toContain('Id="rId5"');
    expect(rels).toContain('Id="rId6"');
    expect(rels).not.toContain('unused.test');
    const newId = /<Relationship Id="([^"]+)"[^>]*Target="https:\/\/new.test\/"/.exec(rels)![1];
    expect(['rId5', 'rId6', 'rId7']).not.toContain(newId);
    expect(await unzipText(bytes, 'word/numbering.xml')).toContain('<w:num w:numId="1">');
    expect((await JSZip.loadAsync(bytes)).file('word/media/image1.png')).not.toBeNull();
    const again = await importDocx(bytes, parseXml);
    expect(again.doc).toEqual(loaded.doc);
  });

  it('declares namespaces raw XML needs when writing a fresh package', async () => {
    const xml = '<w:p><w:r><w:drawing><wp:inline><a:graphic/></wp:inline></w:drawing></w:r></w:p>';
    const bytes = await exportDocx({ type: 'doc', content: [{ type: 'docxRaw', attrs: { xml, kind: 'image', text: '' } }] }, META);
    await assertWellFormed(bytes);
    const doc = await unzipText(bytes, 'word/document.xml');
    expect(doc).toContain('xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"');
    expect(doc).toContain(xml);
  });

  it('exports level-offset lists at ilvl = level + depth using that level\'s format', async () => {
    const doc: PMNode = {
      type: 'doc',
      content: [
        {
          type: 'bulletList',
          attrs: { level: 1 },
          content: [{ type: 'listItem', content: [{ type: 'paragraph', content: [txt('a')] }, { type: 'bulletList', content: [{ type: 'listItem', content: [{ type: 'paragraph', content: [txt('b')] }] }] }] }],
        },
        // Without a separator Word (and import) would nest the next list under the previous item.
        { type: 'paragraph' },
        { type: 'orderedList', attrs: { level: 2, start: 4 }, content: [{ type: 'listItem', content: [{ type: 'paragraph', content: [txt('iv')] }] }] },
      ],
    };
    const bytes = await exportDocx(doc, META);
    const xml = await unzipText(bytes, 'word/document.xml');
    expect([...xml.matchAll(/<w:ilvl w:val="(\d)"\/>/g)].map((m) => m[1])).toEqual(['1', '2', '2']);
    const numbering = await unzipText(bytes, 'word/numbering.xml');
    expect(numbering).toContain('<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="\u25E6"/>');
    expect(numbering).toContain('<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/>');
    expect(numbering).toContain('<w:lvlOverride w:ilvl="2"><w:startOverride w:val="4"/>');
    expect(normalize(await roundTrip(doc))).toEqual(normalize(doc));
  });

  it('prunes numbering so repeated save cycles do not grow it', async () => {
    const original = await makeDocx({
      styles: BASE_STYLES + '<w:style w:type="numbering" w:styleId="Keep"><w:name w:val="Keep"/><w:pPr><w:numPr><w:numId w:val="2"/></w:numPr></w:pPr></w:style>',
      numbering:
        '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum>' +
        '<w:abstractNum w:abstractNumId="1"><w:styleLink w:val="Keep"/><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl></w:abstractNum>' +
        '<w:abstractNum w:abstractNumId="2"><w:numStyleLink w:val="Keep"/></w:abstractNum>' +
        '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>' +
        '<w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num>',
      body: p(r('x'), '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'),
    });
    const once = await exportDocx(RICH_DOC, META, original);
    const twice = await exportDocx(RICH_DOC, META, once);
    const n1 = await unzipText(once, 'word/numbering.xml');
    const n2 = await unzipText(twice, 'word/numbering.xml');
    expect(n2).toEqual(n1);
    expect(n1).toContain('<w:num w:numId="2">');
    expect(n1).not.toContain('<w:num w:numId="1">');
    expect(n1).not.toContain('<w:numStyleLink');
    expect(n1.indexOf('<w:num ')).toBeGreaterThan(n1.lastIndexOf('</w:abstractNum>'));
  });
});

describe.skipIf(!REAL_DOCX || !fs.existsSync(REAL_DOCX))('real document (QUIRE_TEST_DOCX)', () => {
  it('imports, exports and re-imports identically', async () => {
    const bytes = new Uint8Array(fs.readFileSync(REAL_DOCX));
    let t = performance.now();
    const loaded = await importDocx(bytes, parseXml);
    const importMs = performance.now() - t;

    const counts: Record<string, number> = {};
    let badTextStyle = 0;
    const walk = (n: PMNode): void => {
      counts[n.type] = (counts[n.type] ?? 0) + 1;
      for (const m of n.marks ?? []) {
        counts['mark:' + m.type] = (counts['mark:' + m.type] ?? 0) + 1;
        if (m.type === 'textStyle' && m.attrs?.fontFamily == null && m.attrs?.fontSize == null) badTextStyle++;
      }
      n.content?.forEach(walk);
    };
    walk(loaded.doc);
    expect(badTextStyle).toBe(0);

    t = performance.now();
    const out = await exportDocx(loaded.doc, loaded.meta, loaded.original);
    const exportMs = performance.now() - t;
    await assertWellFormed(out);
    const again = await importDocx(out, parseXml);
    expect(again.meta).toEqual(loaded.meta);
    expect(again.doc).toEqual(loaded.doc);

    // Saving the saved file again must not grow numbering.xml.
    const out2 = await exportDocx(again.doc, again.meta, out);
    expect(await unzipText(out2, 'word/numbering.xml')).toEqual(await unzipText(out, 'word/numbering.xml'));
    console.log(`notebook: import ${importMs.toFixed(0)}ms, export ${exportMs.toFixed(0)}ms`, counts);
  }, 30000);
});
