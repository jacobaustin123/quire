/** DOCX -> Quire document model. */

import { lossyFeatures } from './fidelity';
import JSZip from 'jszip';
import type { DocMeta, LoadedDoc, PMMark, PMNode } from '../model';
import {
  Numbering,
  StyleSheet,
  parseParaProps,
  parseRunProps,
  parseTheme,
  type ParaProps,
  type RunProps,
  type ThemeFonts,
} from './styles';
import { BodySlicer, OBJECT_TAGS, containsObject, rawText } from './raw';
import {
  REL_TYPES,
  attr,
  childByName,
  dirOf,
  parseRels,
  relTypeIs,
  relsPathFor,
  resolvePartPath,
  type Relationship,
} from './xml';

type ParseXml = (xml: string) => Document;

const FALLBACK_FONT = 'Times New Roman';
const FALLBACK_SIZE = 20; // Word's built-in default when no w:sz is present anywhere (10pt).
export const MATH_STYLE_ID = 'QuireMath';

interface Ctx {
  styles: StyleSheet;
  theme: ThemeFonts;
  numbering: Numbering | null;
  links: Map<string, string>;
  baseFont: string;
  baseSize: number;
  markCache: Map<string, PMMark[] | undefined>;
  /** Per-numId per-level counters emulating Word's numbering continuation. */
  counters: Map<string, number[]>;
  /** Set while converting a paragraph that holds a drawing / picture / embedded object. */
  sawObject: boolean;
  /** Set when a paragraph may contain nested paragraphs (objects, AlternateContent). */
  sawComplex: boolean;
  /** The document's default paragraph spacing (twips, multiple); paragraphs differing get attrs. */
  baseSpacing: { before: number; after: number; line: number };
}

type Item =
  | { k: 'block'; node: PMNode; empty: boolean }
  | { k: 'li'; ordered: boolean; ilvl: number; numId: string; start: number; para: PMNode }
  | { k: 'cont'; ilvl: number; para: PMNode; fallback: PMNode };

/** Reads a part from the zip, returning null when it does not exist. */
async function readPart(zip: JSZip, path: string): Promise<string | null> {
  const f = zip.file(path);
  return f ? f.async('string') : null;
}

/** Locates the main document part via the package relationships. */
export async function mainDocumentPath(zip: JSZip): Promise<string> {
  const rootRels = await readPart(zip, '_rels/.rels');
  const rel = rootRels ? parseRels(rootRels).find((r) => relTypeIs(r.type, 'officeDocument')) : undefined;
  return rel ? resolvePartPath('', rel.target) : 'word/document.xml';
}

/**
 * Converts a .docx package into the Quire document model.
 *
 * Args:
 *   bytes: The raw .docx file contents.
 *   parseXml: An XML parser returning a DOM Document (native DOMParser or @xmldom/xmldom).
 *
 * Returns:
 *   The document, its base metadata, and `original` (the input bytes, for lossless re-export).
 */
export async function importDocx(
  bytes: Uint8Array,
  parseXml: ParseXml,
): Promise<LoadedDoc & { original: Uint8Array; lossy: string[] }> {
  const zip = await JSZip.loadAsync(bytes);
  const docPath = await mainDocumentPath(zip);
  const docDir = dirOf(docPath);
  const relsXml = await readPart(zip, relsPathFor(docPath));
  const rels: Relationship[] = relsXml ? parseRels(relsXml) : [];
  const partFor = (kind: string, fallback: string): string => {
    const r = rels.find((x) => relTypeIs(x.type, kind));
    return r ? resolvePartPath(docDir, r.target) : fallback;
  };

  const [docXml, stylesXml, numXml, themeXml] = await Promise.all([
    readPart(zip, docPath),
    readPart(zip, partFor('styles', `${docDir}/styles.xml`)),
    readPart(zip, partFor('numbering', `${docDir}/numbering.xml`)),
    readPart(zip, partFor('theme', `${docDir}/theme/theme1.xml`)),
  ]);
  if (docXml === null) throw new Error('Not a Word document: missing main document part');

  const theme = parseTheme(themeXml ? parseXml(themeXml) : null);
  const styles = new StyleSheet(stylesXml ? parseXml(stylesXml) : null, theme);
  const numbering = numXml ? new Numbering(parseXml(numXml), styles) : null;

  const baseRun = { ...styles.defaultRun, ...styles.styleRun(styles.defaultParaStyleId) };
  const basePara = { ...styles.defaultPara, ...styles.stylePara(styles.defaultParaStyleId) };
  const ctx: Ctx = {
    styles,
    theme,
    numbering,
    links: new Map(),
    baseFont: baseRun.font ?? theme.minor ?? FALLBACK_FONT,
    baseSize: baseRun.size ?? FALLBACK_SIZE,
    markCache: new Map(),
    counters: new Map(),
    sawObject: false,
    sawComplex: false,
    baseSpacing: { before: basePara.spacingBefore ?? 0, after: basePara.spacingAfter ?? 0, line: basePara.lineMultiple ?? 1 },
  };
  for (const r of rels) {
    if (r.type === REL_TYPES.hyperlink || relTypeIs(r.type, 'hyperlink')) ctx.links.set(r.id, r.target);
  }

  const meta: DocMeta = {
    baseFont: ctx.baseFont,
    baseSize: ctx.baseSize / 2,
    paragraphSpacing: (basePara.spacingAfter ?? 0) / 20,
    lineSpacing: basePara.lineMultiple ?? 1,
  };

  const root = parseXml(docXml).documentElement;
  const body = childByName(root, 'w:body');
  const items: Item[] = [];
  if (body) collectBody(body, ctx, items, new BodySlicer(docXml));
  const content = buildLists(items);
  const doc: PMNode = { type: 'doc', content: content.length ? content : [{ type: 'paragraph' }] };
  return { doc, meta, original: bytes, lossy: lossyFeatures(docXml, doc) };
}

function rawItem(xml: string, kind: 'table' | 'image' | 'other', el: Element): Item {
  return { k: 'block', node: { type: 'docxRaw', attrs: { xml, kind, text: rawText(el) } }, empty: false };
}

/**
 * Walks the body's direct children. Elements Quire cannot edit (tables, object paragraphs, sdts
 * wrapping either) become verbatim `docxRaw` blocks, sliced from the source text.
 */
function collectBody(body: Element, ctx: Ctx, out: Item[], slicer: BodySlicer): void {
  const kids = body.childNodes;
  for (let i = 0, n = kids.length; i < n; i++) {
    const k = kids[i];
    if (k.nodeType !== 1) continue;
    const el = k as Element;
    const name = el.nodeName;
    if (name === 'w:p') {
      ctx.sawObject = ctx.sawComplex = false;
      const item = convertParagraph(el, ctx);
      const xml = slicer.take(el, !ctx.sawComplex);
      if (ctx.sawObject && xml !== null) {
        out.push(rawItem(xml, 'image', el));
      } else out.push(item);
      continue;
    }
    const isTable = name === 'w:tbl';
    const opaque = isTable || ((name === 'w:sdt' || name === 'w:customXml') && containsObject(el, true));
    const xml = slicer.take(el, false);
    if (opaque && xml !== null) out.push(rawItem(xml, isTable ? 'table' : 'other', el));
    else collectOne(el, ctx, out);
  }
}

/** Walks block-level containers, flattening tables / sdts into a linear item list. */
function collectBlocks(container: Element, ctx: Ctx, out: Item[]): void {
  const kids = container.childNodes;
  for (let i = 0, n = kids.length; i < n; i++) {
    if (kids[i].nodeType === 1) collectOne(kids[i] as Element, ctx, out);
  }
}

function collectOne(el: Element, ctx: Ctx, out: Item[]): void {
  switch (el.nodeName) {
    case 'w:p':
      out.push(convertParagraph(el, ctx));
      break;
    case 'w:tbl': {
      const cell: Item[] = [];
      collectTable(el, ctx, cell);
      // An all-empty table (common as a layout leftover) contributes nothing.
      if (cell.some((it) => it.k !== 'block' || !it.empty)) out.push(...cell);
      break;
    }
    case 'w:sdt': {
      const c = childByName(el, 'w:sdtContent');
      if (c) collectBlocks(c, ctx, out);
      break;
    }
    case 'w:customXml':
    case 'w:ins':
    case 'w:moveTo':
      collectBlocks(el, ctx, out);
      break;
    case 'mc:AlternateContent': {
      const fb = childByName(el, 'mc:Fallback');
      if (fb) collectBlocks(fb, ctx, out);
      break;
    }
  }
}

function collectTable(tbl: Element, ctx: Ctx, out: Item[]): void {
  const kids = tbl.childNodes;
  for (let i = 0, n = kids.length; i < n; i++) {
    const k = kids[i];
    if (k.nodeType !== 1) continue;
    const el = k as Element;
    if (el.nodeName === 'w:tr' || el.nodeName === 'w:sdt' || el.nodeName === 'w:customXml') {
      collectTable(el.nodeName === 'w:sdt' ? (childByName(el, 'w:sdtContent') ?? el) : el, ctx, out);
    } else if (el.nodeName === 'w:tc') {
      collectBlocks(el, ctx, out);
    }
  }
}

const JC_MAP: Record<string, string> = {
  left: 'left',
  start: 'left',
  center: 'center',
  right: 'right',
  end: 'right',
  both: 'justify',
  distribute: 'justify',
};

function convertParagraph(p: Element, ctx: Ctx): Item {
  const pPrEl = childByName(p, 'w:pPr');
  const direct: ParaProps = pPrEl ? parseParaProps(pPrEl) : {};
  const styleId = direct.styleId ?? ctx.styles.defaultParaStyleId;
  const fromStyle = ctx.styles.stylePara(styleId);
  const eff: ParaProps = { ...ctx.styles.defaultPara, ...fromStyle, ...direct };
  const heading = ctx.styles.headingLevel(styleId);

  // Heading styles carry their own size/bold; those are the heading's look, not marks.
  const runBase: RunProps = heading
    ? { ...ctx.styles.defaultRun, ...ctx.styles.styleRun(ctx.styles.defaultParaStyleId) }
    : { ...ctx.styles.defaultRun, ...ctx.styles.styleRun(styleId) };
  const inline = new InlineBuilder();
  convertInline(p, ctx, runBase, inline, null, []);
  const content = inline.finish();

  const align = eff.jc ? JC_MAP[eff.jc] : undefined;
  const attrs: Record<string, unknown> = {};
  if (align && align !== 'left') attrs.textAlign = align;
  // Explicit spacing applies to every kind of paragraph (headings and list items included).
  Object.assign(attrs, spacingAttrs(eff, ctx, heading ? ctx.styles.stylePara(styleId) : undefined));

  if (heading) {
    attrs.level = heading;
    const indent = indentLevel(eff.indLeft);
    if (indent) attrs.indent = indent;
    return block({ type: 'heading', attrs, ...(content.length ? { content } : {}) });
  }

  // Lists: numId "0" explicitly disables numbering.
  const numId = direct.numId ?? fromStyle.numId;
  const ilvl = Math.max(0, Math.min(8, direct.ilvl ?? fromStyle.ilvl ?? 0));
  const lvl = numId && numId !== '0' && ctx.numbering ? ctx.numbering.level(numId, ilvl) : null;
  const outline = eff.outlineLvl !== undefined && eff.outlineLvl >= 0 && eff.outlineLvl <= 8 ? eff.outlineLvl : null;
  if (outline !== null) attrs.outline = outline;

  if (lvl && numId) {
    const para = paragraph(attrs, content);
    return { k: 'li', ordered: lvl.fmt !== 'bullet' && lvl.fmt !== 'none', ilvl, numId, start: bumpCounter(ctx, numId, ilvl, lvl.start), para };
  }

  if (direct.borderBottom && content.every(isBlankInline)) return block({ type: 'horizontalRule' });

  if (content.length === 1 && content[0].type === 'mathInline' && (content[0] as MathNode).display) {
    return block({ type: 'mathBlock', attrs: { latex: content[0].attrs!.latex } });
  }
  for (const c of content) if (c.type === 'mathInline') delete (c as MathNode).display;

  const indent = indentLevel(eff.indLeft);
  const contAttrs = { ...attrs };
  if (indent) attrs.indent = indent;
  const node = paragraph(attrs, content);
  // Quire writes extra paragraphs inside a list item as numId=0 + explicit ilvl.
  if (direct.numId === '0' && direct.hasIlvl) {
    return { k: 'cont', ilvl, para: paragraph(contAttrs, content), fallback: node };
  }
  return block(node);
}

/**
 * Spacing that differs from the document default becomes paragraph attrs (pt / multiple), e.g. the
 * common Word pattern of "8 pt after" in Normal with "0 pt after" set on most paragraphs.
 */
function spacingAttrs(eff: ParaProps, ctx: Ctx, headingStyle?: ParaProps): Record<string, number> {
  const out: Record<string, number> = {};
  const pt = (twips: number) => Math.round((twips / 20) * 2) / 2;
  // A heading's own style spacing is its look (Quire's heading CSS); only direct overrides count.
  const base = headingStyle
    ? {
        before: headingStyle.spacingBefore ?? ctx.baseSpacing.before,
        after: headingStyle.spacingAfter ?? ctx.baseSpacing.after,
        line: headingStyle.lineMultiple ?? ctx.baseSpacing.line,
      }
    : ctx.baseSpacing;
  if (eff.spacingBefore !== undefined && eff.spacingBefore !== base.before) out.spaceBefore = pt(eff.spacingBefore);
  if (eff.spacingAfter !== undefined && eff.spacingAfter !== base.after) out.spaceAfter = pt(eff.spacingAfter);
  if (eff.lineMultiple !== undefined && Math.abs(eff.lineMultiple - base.line) > 0.005) out.lineSpacing = Math.round(eff.lineMultiple * 100) / 100;
  return out;
}

function paragraph(attrs: Record<string, unknown>, content: PMNode[]): PMNode {
  const node: PMNode = { type: 'paragraph' };
  if (hasKeys(attrs)) node.attrs = attrs;
  if (content.length) node.content = content;
  return node;
}

function block(node: PMNode): Item {
  return { k: 'block', node, empty: node.type === 'paragraph' && !node.content };
}

function hasKeys(o: object): boolean {
  for (const _ in o) return true;
  return false;
}

function isBlankInline(n: PMNode): boolean {
  return n.type === 'text' && !/\S/.test(n.text ?? '');
}

function indentLevel(twips: number | undefined): number {
  if (!twips || twips <= 0) return 0;
  return Math.max(0, Math.min(8, Math.round(twips / 720)));
}

/** Advances Word-style list counters and returns this item's number. */
function bumpCounter(ctx: Ctx, numId: string, ilvl: number, start: number): number {
  let c = ctx.counters.get(numId);
  if (!c) ctx.counters.set(numId, (c = []));
  const v = c[ilvl] === undefined ? start : c[ilvl] + 1;
  c[ilvl] = v;
  c.length = ilvl + 1; // a shallower item restarts all deeper levels
  return v;
}

interface StackEntry {
  list: PMNode;
  ilvl: number;
  ordered: boolean;
  numId: string;
}

/** Turns the flat item stream into nested bullet/ordered lists. */
function buildLists(items: Item[]): PMNode[] {
  const out: PMNode[] = [];
  const stack: StackEntry[] = [];
  for (const it of items) {
    if (it.k === 'block') {
      stack.length = 0;
      out.push(it.node);
      continue;
    }
    if (it.k === 'cont') {
      while (stack.length > 1 && stack[stack.length - 1].ilvl > it.ilvl) stack.pop();
      const top = stack[stack.length - 1];
      if (top) lastItem(top.list).content!.push(it.para);
      else out.push(it.fallback);
      continue;
    }
    while (stack.length && stack[stack.length - 1].ilvl > it.ilvl) stack.pop();
    let top: StackEntry | undefined = stack[stack.length - 1];
    if (top && top.ilvl === it.ilvl && (top.ordered !== it.ordered || top.numId !== it.numId)) {
      stack.pop();
      top = stack[stack.length - 1];
    }
    const li: PMNode = { type: 'listItem', content: [it.para] };
    if (top && top.ilvl === it.ilvl) {
      top.list.content!.push(li);
      continue;
    }
    const attrs: Record<string, unknown> = {};
    if (it.ordered && it.start !== 1) attrs.start = it.start;
    // A list that begins below level 0 with no parent item keeps its offset (Word indents it).
    if (!top && it.ilvl > 0) attrs.level = it.ilvl;
    const list: PMNode = { type: it.ordered ? 'orderedList' : 'bulletList', content: [li] };
    if (hasKeys(attrs)) list.attrs = attrs;
    if (top) lastItem(top.list).content!.push(list);
    else out.push(list);
    stack.push({ list, ilvl: it.ilvl, ordered: it.ordered, numId: it.numId });
  }
  return out;
}

function lastItem(list: PMNode): PMNode {
  return list.content![list.content!.length - 1];
}

interface MathNode extends PMNode {
  display?: boolean;
}

/** Accumulates inline content, merging adjacent text runs that share marks. */
class InlineBuilder {
  readonly out: PMNode[] = [];
  private text: string[] = [];
  private key: string | null = null;
  private marks: PMMark[] | undefined;
  private math: string[] | null = null;

  addText(s: string, key: string, marks: PMMark[] | undefined): void {
    if (!s) return;
    this.flushMath();
    if (key !== this.key) {
      this.flushText();
      this.key = key;
      this.marks = marks;
    }
    this.text.push(s);
  }

  addMath(s: string): void {
    this.flushText();
    (this.math ??= []).push(s);
  }

  addNode(node: PMNode): void {
    this.flushText();
    this.flushMath();
    this.out.push(node);
  }

  finish(): PMNode[] {
    this.flushText();
    this.flushMath();
    return this.out;
  }

  private flushText(): void {
    if (this.text.length) {
      let t = this.text.length === 1 ? this.text[0] : this.text.join('');
      if (t.indexOf('\n') >= 0 || t.indexOf('\r') >= 0) t = t.replace(/\r\n?|\n/g, ' ');
      const node: PMNode = { type: 'text', text: t };
      if (this.marks) node.marks = this.marks.map((m) => (m.attrs ? { type: m.type, attrs: { ...m.attrs } } : { type: m.type }));
      this.out.push(node);
      this.text = [];
    }
    this.key = null;
    this.marks = undefined;
  }

  private flushMath(): void {
    if (!this.math) return;
    const raw = this.math.join('');
    this.math = null;
    const display = raw.length >= 4 && raw.startsWith('$$') && raw.endsWith('$$');
    let latex = raw;
    if (display) latex = raw.slice(2, -2);
    else if (raw.length >= 2 && raw.startsWith('$') && raw.endsWith('$')) latex = raw.slice(1, -1);
    const node: MathNode = { type: 'mathInline', attrs: { latex } };
    if (display) node.display = true;
    this.out.push(node);
  }
}

interface FieldFrame {
  phase: 'instr' | 'result';
  instr: string;
  link: string | null;
}

function fieldLink(instr: string): string | null {
  const m = /^\s*HYPERLINK\s+(?:\\[a-z]\s+)*"([^"]+)"(?:\s+\\l\s+"([^"]+)")?/i.exec(instr);
  if (!m) return null;
  return m[2] ? `${m[1]}#${m[2]}` : m[1];
}

/** Recursively converts the inline children of a paragraph (or inline container). */
function convertInline(
  container: Element,
  ctx: Ctx,
  runBase: RunProps,
  out: InlineBuilder,
  link: string | null,
  fields: FieldFrame[],
): void {
  const kids = container.childNodes;
  for (let i = 0, n = kids.length; i < n; i++) {
    const k = kids[i];
    if (k.nodeType !== 1) continue;
    const el = k as Element;
    switch (el.nodeName) {
      case 'w:r':
        convertRun(el, ctx, runBase, out, link, fields);
        break;
      case 'w:hyperlink': {
        const id = attr(el, 'r:id');
        const anchor = attr(el, 'w:anchor');
        let href = id ? (ctx.links.get(id) ?? null) : null;
        if (anchor) href = (href ?? '') + '#' + anchor;
        convertInline(el, ctx, runBase, out, href ?? link, fields);
        break;
      }
      case 'w:fldSimple': {
        const instr = attr(el, 'w:instr') ?? '';
        convertInline(el, ctx, runBase, out, fieldLink(instr) ?? link, fields);
        break;
      }
      case 'w:ins':
      case 'w:moveTo':
      case 'w:smartTag':
      case 'w:customXml':
      case 'w:dir':
      case 'w:bdo':
        convertInline(el, ctx, runBase, out, link, fields);
        break;
      case 'w:sdt': {
        const c = childByName(el, 'w:sdtContent');
        if (c) convertInline(c, ctx, runBase, out, link, fields);
        break;
      }
      case 'mc:AlternateContent': {
        noteAlternate(el, ctx);
        const fb = childByName(el, 'mc:Fallback');
        if (fb) convertInline(fb, ctx, runBase, out, link, fields);
        break;
      }
      case 'm:oMathPara':
      case 'm:oMath': {
        if (!fields.length || fields[fields.length - 1].phase === 'result') {
          const t = mathText(el);
          if (t) {
            const [key, marks] = marksFor({ ...runBase, underline: link !== null ? false : runBase.underline }, link, ctx);
            out.addText(t, key, marks);
          }
        }
        break;
      }
    }
  }
}

function noteAlternate(el: Element, ctx: Ctx): void {
  ctx.sawComplex = true;
  if (containsObject(el, false)) ctx.sawObject = true;
}

function mathText(el: Element): string {
  const parts: string[] = [];
  const walk = (node: Element): void => {
    const kids = node.childNodes;
    for (let i = 0, n = kids.length; i < n; i++) {
      const k = kids[i];
      if (k.nodeType !== 1) continue;
      if (k.nodeName === 'm:t') parts.push(k.textContent ?? '');
      else walk(k as Element);
    }
  };
  walk(el);
  return parts.join('');
}

function convertRun(
  r: Element,
  ctx: Ctx,
  runBase: RunProps,
  out: InlineBuilder,
  link: string | null,
  fields: FieldFrame[],
): void {
  let rPr: Element | null = null;
  let props: RunProps | null = null;
  let key = '';
  let marks: PMMark[] | undefined;
  let math = false;
  let effLink = link;
  const ensureProps = (): void => {
    if (props) return;
    const styleIdEl = rPr && childByName(rPr, 'w:rStyle');
    const rStyle = styleIdEl ? attr(styleIdEl, 'w:val') : null;
    math = rStyle === MATH_STYLE_ID;
    const direct = rPr ? parseRunProps(rPr, ctx.theme) : {};
    props = { ...runBase, ...ctx.styles.styleRun(rStyle), ...direct };
    // Inside links, underline is the link's own look (Hyperlink char styles), not a mark.
    if (effLink !== null) props.underline = false;
    [key, marks] = marksFor(props, effLink, ctx);
  };

  const kids = r.childNodes;
  for (let i = 0, n = kids.length; i < n; i++) {
    const k = kids[i];
    if (k.nodeType !== 1) continue;
    const el = k as Element;
    const name = el.nodeName;
    if (name === 'w:rPr') {
      rPr = el;
      continue;
    }
    if (OBJECT_TAGS.has(name)) {
      ctx.sawObject = ctx.sawComplex = true;
      continue;
    }
    if (name === 'mc:AlternateContent') noteAlternate(el, ctx);
    if (name === 'w:fldChar') {
      const type = attr(el, 'w:fldCharType');
      if (type === 'begin') fields.push({ phase: 'instr', instr: '', link: null });
      else if (type === 'separate' && fields.length) {
        const f = fields[fields.length - 1];
        f.phase = 'result';
        f.link = fieldLink(f.instr);
      } else if (type === 'end') fields.pop();
      continue;
    }
    if (fields.length) {
      const f = fields[fields.length - 1];
      if (f.phase === 'instr') {
        if (name === 'w:instrText') f.instr += el.textContent ?? '';
        continue;
      }
      if (f.link && effLink === link) {
        effLink = f.link;
        props = null;
      }
    }
    switch (name) {
      case 'w:t':
        ensureProps();
        if (math) out.addMath(el.textContent ?? '');
        else out.addText(el.textContent ?? '', key, marks);
        break;
      case 'w:tab':
      case 'w:ptab':
        ensureProps();
        if (math) out.addMath('\t');
        else out.addText('\t', key, marks);
        break;
      case 'w:br':
      case 'w:cr':
        out.addNode({ type: 'hardBreak' });
        break;
      case 'w:noBreakHyphen':
        ensureProps();
        out.addText('-', key, marks);
        break;
      case 'w:sym': {
        const code = attr(el, 'w:char');
        const cp = code ? parseInt(code, 16) : NaN;
        if (Number.isFinite(cp) && cp > 0) {
          ensureProps();
          out.addText(String.fromCodePoint(cp), key, marks);
        }
        break;
      }
      case 'mc:AlternateContent': {
        const fb = childByName(el, 'mc:Fallback');
        // Fallback content inside a run is itself run content (usually a w:pict we skip).
        if (fb) {
          const t = childByName(fb, 'w:t');
          if (t) {
            ensureProps();
            out.addText(t.textContent ?? '', key, marks);
          }
        }
        break;
      }
      // Everything else (drawings, footnote refs, delText, instrText, softHyphen, ...) is dropped.
    }
  }
}

/** Computes the (cached) mark list and merge key for effective run properties. */
function marksFor(p: RunProps, link: string | null, ctx: Ctx): [string, PMMark[] | undefined] {
  const font = p.font ?? ctx.baseFont;
  const size = p.size ?? ctx.baseSize;
  const fontDiff = font !== ctx.baseFont;
  const sizeDiff = size !== ctx.baseSize;
  const key =
    (p.bold ? 'b' : '') +
    (p.italic ? 'i' : '') +
    (p.underline ? 'u' : '') +
    (p.strike ? 's' : '') +
    (p.vertAlign === 'superscript' ? '^' : p.vertAlign === 'subscript' ? '_' : '') +
    '|' + (fontDiff ? font : '') +
    '|' + (sizeDiff ? size : '') +
    '|' + (link ?? '');
  if (ctx.markCache.has(key)) return [key, ctx.markCache.get(key)];
  const marks: PMMark[] = [];
  if (link !== null) marks.push({ type: 'link', attrs: { href: link } });
  if (p.bold) marks.push({ type: 'bold' });
  if (p.italic) marks.push({ type: 'italic' });
  if (p.underline) marks.push({ type: 'underline' });
  if (p.strike) marks.push({ type: 'strike' });
  if (p.vertAlign === 'superscript') marks.push({ type: 'superscript' });
  if (p.vertAlign === 'subscript') marks.push({ type: 'subscript' });
  if (fontDiff || sizeDiff) {
    marks.push({
      type: 'textStyle',
      attrs: { fontFamily: fontDiff ? font : null, fontSize: sizeDiff ? String(size / 2) : null },
    });
  }
  const result = marks.length ? marks : undefined;
  ctx.markCache.set(key, result);
  return [key, result];
}
