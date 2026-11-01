/** Quire document model -> DOCX. */

import JSZip from 'jszip';
import type { DocMeta, PMMark, PMNode } from '../model';
import { MATH_STYLE_ID, mainDocumentPath } from './import';
import { KNOWN_NAMESPACES, relIdsIn, undeclaredPrefixes } from './raw';
import { REL_TYPES, dirOf, escapeXml, parseRels, relTypeIs, relsPathFor, resolvePartPath } from './xml';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = {
  document: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  styles: 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
  numbering: 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml',
  settings: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
  core: 'application/vnd.openxmlformats-package.core-properties+xml',
  app: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
};
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const DEFAULT_SECT_PR =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/></w:sectPr>';
const HEADING_SIZES = [40, 32, 26]; // half-points: 20 / 16 / 13pt

/** Style ids the body writer refers to, resolved against the target package. */
interface StyleIds {
  headings: [string, string, string];
  listParagraph: string | null;
  hyperlink: string;
}

/** Allocates abstractNum / num definitions for the lists in the body. */
class NumberingBuilder {
  readonly abstractXml: string[] = [];
  readonly numXml: string[] = [];
  readonly usedNumIds = new Set<string>();
  private readonly bySig = new Map<string, number>();

  constructor(private nextAbstract: number, private nextNum: number) {}

  /** Returns a fresh numId whose level `ilvl` restarts at `start`. */
  num(sig: string, ilvl: number, start: number): string {
    let abs = this.bySig.get(sig);
    if (abs === undefined) {
      abs = this.nextAbstract++;
      this.bySig.set(sig, abs);
      this.abstractXml.push(abstractNumXml(abs, sig));
    }
    const id = String(this.nextNum++);
    this.usedNumIds.add(id);
    this.numXml.push(
      `<w:num w:numId="${id}"><w:abstractNumId w:val="${abs}"/>` +
        `<w:lvlOverride w:ilvl="${ilvl}"><w:startOverride w:val="${start}"/></w:lvlOverride></w:num>`,
    );
    return id;
  }
}

const ORDERED_FMTS = ['decimal', 'lowerLetter', 'lowerRoman'];
const BULLETS = ['•', '◦', '▪'];

function abstractNumXml(id: number, sig: string): string {
  const lvls: string[] = [];
  for (let i = 0; i < 9; i++) {
    const ordered = sig[i] === 'o';
    const fmt = ordered ? ORDERED_FMTS[i % 3] : 'bullet';
    const text = ordered ? `%${i + 1}.` : BULLETS[i % 3];
    lvls.push(
      `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="${text}"/>` +
        `<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720 * (i + 1)}" w:hanging="360"/></w:pPr></w:lvl>`,
    );
  }
  return `<w:abstractNum w:abstractNumId="${id}"><w:multiLevelType w:val="hybridMultilevel"/>${lvls.join('')}</w:abstractNum>`;
}

function isList(n: PMNode): boolean {
  return n.type === 'bulletList' || n.type === 'orderedList';
}

/** Records the first list type seen at each relative depth of a list subtree. */
function observeTypes(list: PMNode, depth: number, out: string[]): void {
  if (depth > 8) return;
  if (out[depth] === undefined) out[depth] = list.type === 'orderedList' ? 'o' : 'b';
  for (const li of list.content ?? []) {
    for (const c of li.content ?? []) if (isList(c)) observeTypes(c, depth + 1, out);
  }
}

/** Builds a 9-level type signature: `prefix` levels, then observed levels, then the last one repeated. */
function signature(list: PMNode, depth: number, prefix: string): string {
  const seen: string[] = [];
  observeTypes(list, depth, seen);
  let sig = prefix.slice(0, depth);
  while (sig.length < depth) sig += 'b';
  let last = seen[depth] ?? 'b';
  for (let i = depth; i < 9; i++) {
    last = seen[i] ?? last;
    sig += last;
  }
  return sig;
}

/** Collects the verbatim XML of every docxRaw block in the document. */
function collectRaw(n: PMNode, out: string[] = []): string[] {
  if (n.type === 'docxRaw') out.push(String(n.attrs?.xml ?? ''));
  else for (const c of n.content ?? []) if (c.type !== 'text') collectRaw(c, out);
  return out;
}

/** Adds xmlns declarations for prefixes raw XML needs but `rootTag` lacks. */
function declareNamespaces(rootTag: string, rawXml: string): string {
  const missing: string[] = [];
  for (const p of undeclaredPrefixes(rawXml)) {
    if (!new RegExp(`\\sxmlns:${p.replace(/\./g, '\\.')}\\s*=`).test(rootTag) && KNOWN_NAMESPACES[p]) {
      missing.push(` xmlns:${p}="${KNOWN_NAMESPACES[p]}"`);
    }
  }
  return missing.length ? rootTag.replace(/^<w:document\b/, '<w:document' + missing.join('')) : rootTag;
}

function listStart(list: PMNode): number {
  const s = Number(list.attrs?.start ?? 1);
  return list.type === 'orderedList' && Number.isFinite(s) && s >= 0 ? Math.floor(s) : 1;
}

/** Serializes block content into `<w:body>` children, tracking links and numbering. */
class BodyWriter {
  readonly parts: string[] = [];
  readonly links = new Map<string, string>();

  constructor(
    private readonly meta: DocMeta,
    private readonly styles: StyleIds,
    private readonly numbering: NumberingBuilder,
    private readonly newRelId: () => string,
  ) {}

  blocks(nodes: PMNode[]): void {
    for (const n of nodes) this.block(n);
  }

  private block(n: PMNode): void {
    switch (n.type) {
      case 'paragraph':
        this.paragraph(this.pPr(null, null, n.attrs), n.content);
        break;
      case 'heading': {
        const lvl = Math.max(1, Math.min(3, Number(n.attrs?.level ?? 1)));
        this.paragraph(this.pPr(this.styles.headings[lvl - 1], null, { ...n.attrs, outline: null }), n.content);
        break;
      }
      case 'bulletList':
      case 'orderedList': {
        const level = Math.max(0, Math.min(8, Math.round(Number(n.attrs?.level ?? 0)) || 0));
        const sig = signature(n, level, '');
        this.list(n, level, this.numbering.num(sig, level, listStart(n)), sig);
        break;
      }
      case 'docxRaw':
        this.parts.push(String(n.attrs?.xml ?? ''));
        break;
      case 'horizontalRule':
        this.parts.push('<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p>');
        break;
      case 'mathBlock':
        this.parts.push(
          `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${mathRun('$$' + String(n.attrs?.latex ?? '') + '$$')}</w:p>`,
        );
        break;
      default:
        // Unknown blocks (e.g. future node types): keep their text rather than dropping it.
        if (n.content?.some((c) => c.type === 'text' || c.type === 'hardBreak')) this.paragraph(this.pPr(null, null, n.attrs), n.content);
        else if (n.content) this.blocks(n.content);
    }
  }

  private list(list: PMNode, depth: number, numId: string, sig: string): void {
    const ilvl = Math.min(depth, 8);
    for (const li of list.content ?? []) {
      let first = true;
      for (const c of li.content ?? []) {
        if (isList(c)) {
          const d = Math.min(depth + 1, 8);
          const type = c.type === 'orderedList' ? 'o' : 'b';
          const start = listStart(c);
          if (sig[d] === type && start === 1) this.list(c, depth + 1, numId, sig);
          else {
            const sub = signature(c, d, sig);
            this.list(c, depth + 1, this.numbering.num(sub, d, start), sub);
          }
        } else if (c.type === 'paragraph' || c.type === 'heading') {
          const attrs = { ...c.attrs, indent: 0 };
          if (first) this.paragraph(this.pPr(this.styles.listParagraph, [ilvl, numId], attrs), c.content);
          else this.paragraph(this.pPr(this.styles.listParagraph, [ilvl, '0'], attrs, 720 * (ilvl + 1)), c.content);
          first = false;
        } else {
          this.block(c);
          first = false;
        }
      }
    }
  }

  private pPr(
    styleId: string | null,
    num: [number, string] | null,
    attrs: Record<string, unknown> | undefined,
    indTwips?: number,
  ): string {
    let s = '';
    if (styleId) s += `<w:pStyle w:val="${escapeXml(styleId)}"/>`;
    if (num) s += `<w:numPr><w:ilvl w:val="${num[0]}"/><w:numId w:val="${num[1]}"/></w:numPr>`;
    const indent = Math.max(0, Math.min(8, Math.round(Number(attrs?.indent ?? 0)) || 0));
    const left = indTwips ?? (indent ? indent * 720 : 0);
    if (left) s += `<w:ind w:left="${left}"/>`;
    const align = attrs?.textAlign;
    if (align === 'center' || align === 'right') s += `<w:jc w:val="${align}"/>`;
    else if (align === 'justify') s += '<w:jc w:val="both"/>';
    const outline = attrs?.outline;
    if (typeof outline === 'number' && outline >= 0 && outline <= 8) s += `<w:outlineLvl w:val="${Math.round(outline)}"/>`;
    return s ? `<w:pPr>${s}</w:pPr>` : '';
  }

  private paragraph(pPr: string, content: PMNode[] | undefined): void {
    this.parts.push('<w:p>', pPr);
    if (content) this.inline(content);
    this.parts.push('</w:p>');
  }

  private inline(content: PMNode[]): void {
    let openHref: string | null = null;
    for (const n of content) {
      const href = n.type === 'text' ? linkHref(n.marks) : null;
      if (href !== openHref) {
        if (openHref !== null) this.parts.push('</w:hyperlink>');
        if (href !== null) this.parts.push(this.hyperlinkOpen(href));
        openHref = href;
      }
      if (n.type === 'text') this.parts.push(this.textRun(n.text ?? '', n.marks, href !== null));
      else if (n.type === 'hardBreak') this.parts.push('<w:r><w:br/></w:r>');
      else if (n.type === 'mathInline') this.parts.push(mathRun('$' + String(n.attrs?.latex ?? '') + '$'));
      else if (n.content) this.inline(n.content);
    }
    if (openHref !== null) this.parts.push('</w:hyperlink>');
  }

  private hyperlinkOpen(href: string): string {
    if (href.startsWith('#')) return `<w:hyperlink w:anchor="${escapeXml(href.slice(1))}" w:history="1">`;
    let id = this.links.get(href);
    if (!id) {
      id = this.newRelId();
      this.links.set(href, id);
    }
    return `<w:hyperlink r:id="${id}" w:history="1">`;
  }

  private textRun(text: string, marks: PMMark[] | undefined, inLink: boolean): string {
    if (!text) return '';
    const rPr = this.rPr(marks, inLink);
    const body = text.indexOf('\t') < 0 ? tElem(text) : text.split('\t').map((t) => (t ? tElem(t) : '')).join('<w:tab/>');
    return `<w:r>${rPr}${body}</w:r>`;
  }

  private rPr(marks: PMMark[] | undefined, inLink: boolean): string {
    let font: string | null = null;
    let size: number | null = null;
    let b = false, i = false, u = false, s = false, va: string | null = null;
    for (const m of marks ?? []) {
      switch (m.type) {
        case 'bold': b = true; break;
        case 'italic': i = true; break;
        case 'underline': u = true; break;
        case 'strike': s = true; break;
        case 'superscript': va = 'superscript'; break;
        case 'subscript': va = 'subscript'; break;
        case 'textStyle': {
          const ff = m.attrs?.fontFamily;
          if (typeof ff === 'string' && ff.trim()) font = firstFamily(ff);
          const fs = parseFloat(String(m.attrs?.fontSize ?? ''));
          if (Number.isFinite(fs) && fs > 0) size = Math.round(fs * 2);
          break;
        }
      }
    }
    let x = '';
    if (inLink) x += `<w:rStyle w:val="${escapeXml(this.styles.hyperlink)}"/>`;
    if (font && font !== this.meta.baseFont) {
      const f = escapeXml(font);
      x += `<w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:cs="${f}"/>`;
    }
    if (b) x += '<w:b/><w:bCs/>';
    if (i) x += '<w:i/><w:iCs/>';
    if (s) x += '<w:strike/>';
    if (size && size !== Math.round(this.meta.baseSize * 2)) x += `<w:sz w:val="${size}"/><w:szCs w:val="${size}"/>`;
    if (u) x += '<w:u w:val="single"/>';
    if (va) x += `<w:vertAlign w:val="${va}"/>`;
    return x ? `<w:rPr>${x}</w:rPr>` : '';
  }
}

function tElem(t: string): string {
  return `<w:t xml:space="preserve">${escapeXml(t)}</w:t>`;
}

function mathRun(text: string): string {
  return `<w:r><w:rPr><w:rStyle w:val="${MATH_STYLE_ID}"/></w:rPr>${tElem(text)}</w:r>`;
}

function linkHref(marks: PMMark[] | undefined): string | null {
  if (!marks) return null;
  for (const m of marks) {
    if (m.type === 'link') {
      const h = m.attrs?.href;
      return typeof h === 'string' && h ? h : null;
    }
  }
  return null;
}

/** Reduces a CSS font-family list ('"Helvetica Neue", Arial') to its first family name. */
function firstFamily(ff: string): string {
  const first = ff.split(',')[0].trim();
  return first.replace(/^["']|["']$/g, '').trim() || ff.trim();
}

/** Scans styles.xml text for (styleId, name) pairs. */
function scanStyles(xml: string): { ids: Set<string>; byName: Map<string, string> } {
  const ids = new Set<string>();
  const byName = new Map<string, string>();
  const re = /<w:style\b[^>]*?w:styleId="([^"]*)"[^>]*>(?:\s*<w:name w:val="([^"]*)")?/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml))) {
    ids.add(m[1]);
    if (m[2]) byName.set(m[2].toLowerCase(), m[1]);
  }
  return { ids, byName };
}

function headingStyleXml(id: string, n: number): string {
  const sz = HEADING_SIZES[n - 1];
  return (
    `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="heading ${n}"/><w:basedOn w:val="Normal"/>` +
    `<w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/>` +
    `<w:outlineLvl w:val="${n - 1}"/></w:pPr><w:rPr><w:b/><w:bCs/><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/></w:rPr></w:style>`
  );
}

const HYPERLINK_STYLE_XML =
  '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:uiPriority w:val="99"/>' +
  '<w:unhideWhenUsed/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>';
const MATH_STYLE_XML =
  `<w:style w:type="character" w:customStyle="1" w:styleId="${MATH_STYLE_ID}"><w:name w:val="Quire Math"/>` +
  '<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr></w:style>';

/** Ensures the styles Quire writes exist; returns the patched XML and the ids to use. */
function patchStyles(xml: string): { xml: string; ids: StyleIds } {
  const { ids, byName } = scanStyles(xml);
  const add: string[] = [];
  const headings = [1, 2, 3].map((n) => {
    const id = byName.get(`heading ${n}`) ?? (ids.has(`Heading${n}`) ? `Heading${n}` : null);
    if (id) return id;
    add.push(headingStyleXml(`Heading${n}`, n));
    return `Heading${n}`;
  }) as [string, string, string];
  let hyperlink = byName.get('hyperlink') ?? (ids.has('Hyperlink') ? 'Hyperlink' : null);
  if (!hyperlink) {
    add.push(HYPERLINK_STYLE_XML);
    hyperlink = 'Hyperlink';
  }
  if (!ids.has(MATH_STYLE_ID)) add.push(MATH_STYLE_XML);
  const listParagraph = ids.has('ListParagraph') ? 'ListParagraph' : (byName.get('list paragraph') ?? null);
  return { xml: add.length ? insertBeforeClose(xml, 'w:styles', add.join('')) : xml, ids: { headings, listParagraph, hyperlink } };
}

/** Inserts `content` just before `</tag>`, expanding a self-closing root if needed. */
function insertBeforeClose(xml: string, tag: string, content: string): string {
  const close = `</${tag}>`;
  const i = xml.lastIndexOf(close);
  if (i >= 0) return xml.slice(0, i) + content + xml.slice(i);
  const m = new RegExp(`<${tag}\\b([^>]*?)\\s*/>`).exec(xml);
  if (!m) throw new Error(`Malformed part: no <${tag}> root`);
  return xml.slice(0, m.index) + `<${tag}${m[1]}>${content}${close}` + xml.slice(m.index + m[0].length);
}

function newStylesXml(meta: DocMeta): string {
  const f = escapeXml(meta.baseFont);
  const sz = Math.round(meta.baseSize * 2);
  const after = Math.round(meta.paragraphSpacing * 20);
  const line = Math.round((meta.lineSpacing || 1) * 240);
  return (
    XML_DECL +
    `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>` +
    `<w:rFonts w:ascii="${f}" w:eastAsia="${f}" w:hAnsi="${f}" w:cs="${f}"/><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>` +
    `<w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>` +
    `<w:pPrDefault><w:pPr><w:spacing w:after="${after}" w:line="${line}" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>` +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
    '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/>' +
    '<w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>' +
    headingStyleXml('Heading1', 1) + headingStyleXml('Heading2', 2) + headingStyleXml('Heading3', 3) +
    HYPERLINK_STYLE_XML + MATH_STYLE_XML +
    '</w:styles>'
  );
}

const EMPTY_NUMBERING = XML_DECL + `<w:numbering xmlns:w="${W_NS}"></w:numbering>`;

/**
 * Drops all num/abstractNum definitions the rewritten body can no longer reference, keeping
 * those used by styles (and other parts) plus style-linked abstractNums, then appends ours.
 */
function rebuildNumbering(xml: string, keepNumIds: Set<string>, build: (maxAbs: number, maxNum: number) => NumberingBuilder): string {
  const keptAbs = new Set<string>();
  const numRe = /<w:num\b[^>]*?w:numId="(\d+)"[^>]*>([\s\S]*?)<\/w:num>/g;
  let m: RegExpExecArray | null;
  while ((m = numRe.exec(xml))) {
    if (!keepNumIds.has(m[1])) continue;
    const a = /<w:abstractNumId w:val="(\d+)"/.exec(m[2]);
    if (a) keptAbs.add(a[1]);
  }
  let maxAbs = -1;
  let maxNum = 0;
  xml = xml.replace(/<w:abstractNum\b[^>]*?w:abstractNumId="(\d+)"[^>]*>([\s\S]*?)<\/w:abstractNum>/g, (all, id: string, inner: string) => {
    // w:styleLink abstracts define numbering styles (which styles.xml may still use); plain
    // numStyleLink abstracts are mere pointers and survive only via a kept num.
    if (keptAbs.has(id) || /<w:styleLink\b/.test(inner)) {
      maxAbs = Math.max(maxAbs, Number(id));
      return all;
    }
    return '';
  });
  xml = xml.replace(numRe, (all, id: string) => {
    if (keepNumIds.has(id)) {
      maxNum = Math.max(maxNum, Number(id));
      return all;
    }
    return '';
  });
  // numIdMacAtCleanup refers to the old numIds; it is optional, so drop it.
  xml = xml.replace(/<w:numIdMacAtCleanup\b[^>]*\/>/g, '');
  const nb = build(maxAbs + 1, maxNum + 1);
  const absXml = nb.abstractXml.join('');
  if (absXml) {
    const lastAbs = xml.lastIndexOf('</w:abstractNum>');
    if (lastAbs >= 0) {
      const at = lastAbs + '</w:abstractNum>'.length;
      xml = xml.slice(0, at) + absXml + xml.slice(at);
    } else {
      const firstNum = xml.search(/<w:num\b/);
      xml = firstNum >= 0 ? xml.slice(0, firstNum) + absXml + xml.slice(firstNum) : insertBeforeClose(xml, 'w:numbering', absXml);
    }
  }
  if (nb.numXml.length) xml = insertBeforeClose(xml, 'w:numbering', nb.numXml.join(''));
  return xml;
}

/** Collects numIds referenced via `<w:numId w:val>` in a part's XML text. */
function numIdsIn(xml: string, out: Set<string>): void {
  const re = /<w:numId w:val="(\d+)"/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml))) if (m[1] !== '0') out.add(m[1]);
}

function addContentTypeOverride(xml: string, partName: string, type: string): string {
  if (xml.includes(`PartName="${partName}"`)) return xml;
  return insertBeforeClose(xml, 'Types', `<Override PartName="${partName}" ContentType="${type}"/>`);
}

function relXml(id: string, type: string, target: string, external: boolean): string {
  return `<Relationship Id="${id}" Type="${type}" Target="${escapeXml(target)}"${external ? ' TargetMode="External"' : ''}/>`;
}

function uniqueIdFactory(taken: Set<string>): () => string {
  let n = 1;
  return () => {
    while (taken.has(`rIdQ${n}`)) n++;
    const id = `rIdQ${n}`;
    taken.add(id);
    return id;
  };
}

/**
 * Converts a Quire document into a .docx package.
 *
 * Args:
 *   doc: The ProseMirror JSON document (`{type: "doc", content: [...]}`).
 *   meta: Base font / size / paragraph spacing. Used for new packages and to decide which runs
 *     need explicit fonts or sizes.
 *   original: The bytes of the .docx this document was loaded from. When given, its package is
 *     reused (styles, theme, headers/footers, settings, page setup) and only the body is replaced.
 *
 * Returns:
 *   The .docx file contents.
 */
export async function exportDocx(doc: PMNode, meta: DocMeta, original?: Uint8Array): Promise<Uint8Array> {
  return original ? exportIntoOriginal(doc, meta, original) : exportFresh(doc, meta);
}

async function exportFresh(doc: PMNode, meta: DocMeta): Promise<Uint8Array> {
  const stylesXml = newStylesXml(meta);
  const { ids } = patchStyles(stylesXml);
  const taken = new Set(['rId1', 'rId2', 'rId3']);
  let nb: NumberingBuilder | null = null;
  const writer = new BodyWriter(meta, ids, (nb = new NumberingBuilder(0, 1)), uniqueIdFactory(taken));
  writer.blocks(doc.content ?? []);
  const numberingXml = rebuildNumbering(EMPTY_NUMBERING, new Set(), () => nb!);

  const rootTag = declareNamespaces(`<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}">`, collectRaw(doc).join(''));
  const docXml =
    XML_DECL + rootTag + '<w:body>' + writer.parts.join('') + DEFAULT_SECT_PR + '</w:body></w:document>';
  const rels = [
    relXml('rId1', REL_TYPES.styles, 'styles.xml', false),
    relXml('rId2', REL_TYPES.numbering, 'numbering.xml', false),
    relXml('rId3', REL_TYPES.settings, 'settings.xml', false),
    ...[...writer.links].map(([href, id]) => relXml(id, REL_TYPES.hyperlink, href, true)),
  ];
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');

  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    XML_DECL +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      `<Override PartName="/word/document.xml" ContentType="${CT.document}"/>` +
      `<Override PartName="/word/styles.xml" ContentType="${CT.styles}"/>` +
      `<Override PartName="/word/numbering.xml" ContentType="${CT.numbering}"/>` +
      `<Override PartName="/word/settings.xml" ContentType="${CT.settings}"/>` +
      `<Override PartName="/docProps/core.xml" ContentType="${CT.core}"/>` +
      `<Override PartName="/docProps/app.xml" ContentType="${CT.app}"/></Types>`,
  );
  zip.file(
    '_rels/.rels',
    XML_DECL +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      relXml('rId1', REL_TYPES.officeDocument, 'word/document.xml', false) +
      relXml('rId2', REL_TYPES.coreProps, 'docProps/core.xml', false) +
      relXml('rId3', REL_TYPES.extendedProps, 'docProps/app.xml', false) +
      '</Relationships>',
  );
  zip.file('word/document.xml', docXml);
  zip.file('word/styles.xml', stylesXml);
  zip.file('word/numbering.xml', numberingXml);
  zip.file(
    'word/settings.xml',
    XML_DECL + `<w:settings xmlns:w="${W_NS}"><w:defaultTabStop w:val="720"/><w:compat>` +
      '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>' +
      '</w:compat></w:settings>',
  );
  zip.file(
    'word/_rels/document.xml.rels',
    XML_DECL + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + rels.join('') + '</Relationships>',
  );
  zip.file(
    'docProps/core.xml',
    XML_DECL +
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
      'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
      'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>` +
      `<dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`,
  );
  zip.file(
    'docProps/app.xml',
    XML_DECL +
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Quire</Application></Properties>',
  );
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

async function exportIntoOriginal(doc: PMNode, meta: DocMeta, original: Uint8Array): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(original);
  const docPath = await mainDocumentPath(zip);
  const docDir = dirOf(docPath);
  const relsPath = relsPathFor(docPath);
  const read = async (p: string): Promise<string | null> => {
    const f = zip.file(p);
    return f ? f.async('string') : null;
  };
  const docXml = await read(docPath);
  if (docXml === null) throw new Error('Original package has no main document part');
  const relsXml = (await read(relsPath)) ?? '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  let ctXml = (await read('[Content_Types].xml')) ?? '';

  // Old hyperlink rels belong to the body we are replacing, so they are dropped and re-created,
  // except those still referenced from passthrough (docxRaw) XML.
  const rawXml = collectRaw(doc).join('');
  const rawRelIds = new Set<string>();
  relIdsIn(rawXml, rawRelIds);
  const rels = parseRels(relsXml);
  const keptRels = rels.filter((r) => !relTypeIs(r.type, 'hyperlink') || rawRelIds.has(r.id));
  const taken = new Set(rels.map((r) => r.id));
  const extraRels: string[] = [];

  let stylesPath: string;
  const stylesRel = keptRels.find((r) => relTypeIs(r.type, 'styles'));
  let stylesXml: string;
  if (stylesRel) {
    stylesPath = resolvePartPath(docDir, stylesRel.target);
    stylesXml = (await read(stylesPath)) ?? newStylesXml(meta);
  } else {
    stylesPath = resolvePartPath(docDir, 'styles.xml');
    stylesXml = newStylesXml(meta);
    const id = uniqueIdFactory(taken)();
    extraRels.push(relXml(id, REL_TYPES.styles, 'styles.xml', false));
    ctXml = addContentTypeOverride(ctXml, '/' + stylesPath, CT.styles);
  }
  const patched = patchStyles(stylesXml);

  let numPath: string;
  let numXml: string;
  const numRel = keptRels.find((r) => relTypeIs(r.type, 'numbering'));
  if (numRel) {
    numPath = resolvePartPath(docDir, numRel.target);
    numXml = (await read(numPath)) ?? EMPTY_NUMBERING;
  } else {
    numPath = resolvePartPath(docDir, 'numbering.xml');
    numXml = EMPTY_NUMBERING;
    const id = uniqueIdFactory(taken)();
    extraRels.push(relXml(id, REL_TYPES.numbering, 'numbering.xml', false));
    ctXml = addContentTypeOverride(ctXml, '/' + numPath, CT.numbering);
  }

  // numIds still referenced outside the body (styles, headers, footers, notes) must survive.
  const keepNumIds = new Set<string>();
  numIdsIn(patched.xml, keepNumIds);
  numIdsIn(rawXml, keepNumIds);
  for (const name of Object.keys(zip.files)) {
    if (name === docPath || name === numPath || !name.startsWith(docDir + '/') || !name.endsWith('.xml')) continue;
    if (/\/(header|footer|footnotes|endnotes|comments)\d*\.xml$/.test(name)) numIdsIn((await read(name)) ?? '', keepNumIds);
  }

  let writer: BodyWriter | null = null;
  const newNumXml = rebuildNumbering(numXml, keepNumIds, (maxAbs, maxNum) => {
    const nb = new NumberingBuilder(maxAbs, maxNum);
    writer = new BodyWriter(meta, patched.ids, nb, uniqueIdFactory(taken));
    writer.blocks(doc.content ?? []);
    return nb;
  });
  const w = writer as BodyWriter | null;
  if (!w) throw new Error('unreachable');

  const bodyOpen = /<w:body\b[^>]*>/.exec(docXml);
  const bodyClose = docXml.lastIndexOf('</w:body>');
  if (!bodyOpen || bodyClose < 0) throw new Error('Original document has no <w:body>');
  let prefix = docXml.slice(0, bodyOpen.index + bodyOpen[0].length);
  const rootStart = prefix.indexOf('<w:document');
  const rootEnd = prefix.indexOf('>', rootStart) + 1;
  if (rootStart >= 0 && rootEnd > 0) {
    // r: is always needed for hyperlinks; raw XML may need more (wp:, a:, pic:, ...).
    const rootTag = declareNamespaces(prefix.slice(rootStart, rootEnd), rawXml + ' r:x=""');
    prefix = prefix.slice(0, rootStart) + rootTag + prefix.slice(rootEnd);
  }
  const oldBody = docXml.slice(bodyOpen.index + bodyOpen[0].length, bodyClose);
  const sectIdx = oldBody.lastIndexOf('<w:sectPr');
  let sectPr = '';
  if (sectIdx >= 0) {
    const tail = oldBody.slice(sectIdx).trim();
    // Only the body-level sectPr (last child of body) — not one inside the last paragraph's pPr.
    if (!tail.includes('</w:p>') && (tail.endsWith('</w:sectPr>') || /^<w:sectPr\b[^>]*\/>$/.test(tail))) sectPr = tail;
  }
  const newDocXml = prefix + w.parts.join('') + sectPr + docXml.slice(bodyClose);

  const newRels =
    relsXml.slice(0, relsXml.indexOf('>', relsXml.indexOf('<Relationships')) + 1).replace(/\/>$/, '>') +
    keptRels.map((r) => r.raw).join('') +
    extraRels.join('') +
    [...w.links].map(([href, id]) => relXml(id, REL_TYPES.hyperlink, href, true)).join('') +
    '</Relationships>';
  const relsHead = relsXml.startsWith('<?xml') ? '' : XML_DECL;

  zip.file(docPath, newDocXml);
  zip.file(relsPath, relsHead + newRels);
  zip.file(stylesPath, patched.xml);
  zip.file(numPath, newNumXml);
  if (ctXml) zip.file('[Content_Types].xml', ctXml);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
