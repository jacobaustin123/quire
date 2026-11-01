/** Markdown <-> Quire document model. See docs/SCHEMA.md for the JSON contract. */

import MarkdownIt from 'markdown-it';
import type Token from 'markdown-it/lib/token.mjs';
import type StateInline from 'markdown-it/lib/rules_inline/state_inline.mjs';
import type StateBlock from 'markdown-it/lib/rules_block/state_block.mjs';
import { DEFAULT_META, type LoadedDoc, type PMMark, type PMNode } from './model';

const CODE_FONT = 'Menlo';
const MAX_INDENT = 8;

/** Canonical mark order used for every emitted text node (ProseMirror re-sorts by schema anyway). */
const MARK_ORDER = ['link', 'bold', 'italic', 'strike', 'underline', 'superscript', 'subscript', 'textStyle'];

/** Marks written with emphasis-style delimiters, which cannot touch inner whitespace. */
const DELIMITER_MARKS = new Set(['bold', 'italic', 'strike']);

// ---------------------------------------------------------------------------
// markdown-it setup (math plugin)
// ---------------------------------------------------------------------------

const mdUtils = new MarkdownIt().utils;

function isWhiteSpaceCode(code: number): boolean {
  return mdUtils.isWhiteSpace(code);
}

function isDigitCode(code: number): boolean {
  return code >= 0x30 && code <= 0x39;
}

/** Finds the next unescaped `$$` at or after `from`, or -1. */
function findDoubleDollar(src: string, from: number, max: number): number {
  for (let i = from; i < max - 1; i++) {
    const c = src.charCodeAt(i);
    if (c === 0x5c) {
      i++;
    } else if (c === 0x24 && src.charCodeAt(i + 1) === 0x24) {
      return i;
    }
  }
  return -1;
}

/** Inline `$...$` (pandoc rules) and inline `$$...$$`, both producing `math_inline` tokens. */
function mathInlineRule(state: StateInline, silent: boolean): boolean {
  const src = state.src;
  const start = state.pos;
  const max = state.posMax;
  if (src.charCodeAt(start) !== 0x24) return false;

  if (src.charCodeAt(start + 1) === 0x24) {
    const end = findDoubleDollar(src, start + 2, max);
    const latex = end < 0 ? '' : src.slice(start + 2, end).trim();
    if (!latex) {
      // An unmatched `$$` is literal; consuming both stops the second `$` opening single-dollar math.
      if (!silent) state.pending += '$$';
      state.pos = start + 2;
      return true;
    }
    if (!silent) state.push('math_inline', '', 0).content = latex;
    state.pos = end + 2;
    return true;
  }

  const first = src.charCodeAt(start + 1);
  if (start + 1 >= max || isWhiteSpaceCode(first)) return false;
  for (let i = start + 1; i < max; i++) {
    const c = src.charCodeAt(i);
    if (c === 0x5c) {
      i++;
      continue;
    }
    if (c !== 0x24) continue;
    if (isWhiteSpaceCode(src.charCodeAt(i - 1))) continue;
    if (i + 1 < max && isDigitCode(src.charCodeAt(i + 1))) continue;
    if (!silent) state.push('math_inline', '', 0).content = src.slice(start + 1, i);
    state.pos = i + 1;
    return true;
  }
  return false;
}

/** Block `$$ ... $$`, either on one line or spanning several lines. */
function mathBlockRule(state: StateBlock, startLine: number, endLine: number, silent: boolean): boolean {
  if (state.sCount[startLine] - state.blkIndent >= 4) return false;
  const pos = state.bMarks[startLine] + state.tShift[startLine];
  const max = state.eMarks[startLine];
  const src = state.src;
  if (!src.startsWith('$$', pos)) return false;

  const rest = src.slice(pos + 2, max).trimEnd();
  let latex: string;
  let lastLine = startLine;
  const innerClose = rest.indexOf('$$');
  if (innerClose >= 0) {
    // Only a line that is exactly `$$...$$` is a block; `$$a$$ text` is inline math in a paragraph.
    if (innerClose !== rest.length - 2) return false;
    latex = rest.slice(0, innerClose).trim();
    if (!latex) return false;
  } else {
    let close = -1;
    let closeText = '';
    for (let line = startLine + 1; line < endLine; line++) {
      const lineStart = state.bMarks[line] + state.tShift[line];
      const lineEnd = state.eMarks[line];
      const isBlank = lineStart >= lineEnd;
      if (!isBlank && state.sCount[line] < state.blkIndent) break;
      const text = src.slice(lineStart, lineEnd).trimEnd();
      if (text.endsWith('$$')) {
        close = line;
        closeText = state.getLines(line, line + 1, state.sCount[startLine], false).trimEnd();
        closeText = closeText.slice(0, closeText.length - 2).trimEnd();
        break;
      }
    }
    if (close < 0) return false;
    const lines: string[] = [];
    if (rest.trim()) lines.push(rest.trim());
    if (close > startLine + 1) {
      lines.push(state.getLines(startLine + 1, close, state.sCount[startLine], false));
    }
    if (closeText.trim()) lines.push(closeText);
    latex = lines.join('\n');
    lastLine = close;
  }
  if (silent) return true;
  const token = state.push('math_block', '', 0);
  token.content = latex;
  token.map = [startLine, lastLine + 1];
  state.line = lastLine + 1;
  return true;
}

function mathPlugin(md: MarkdownIt): void {
  md.inline.ruler.before('escape', 'math_inline', mathInlineRule);
  md.block.ruler.before('fence', 'math_block', mathBlockRule, {
    alt: ['paragraph', 'reference', 'blockquote', 'list'],
  });
}

function createParser(): MarkdownIt {
  const md = new MarkdownIt('default', { html: true, linkify: false, typographer: false });
  // Keep hrefs exactly as written (minus markdown escapes) so they round-trip; the default
  // percent-encodes them.
  md.normalizeLink = (url: string): string => url;
  md.use(mathPlugin);
  return md;
}

const md = createParser();

// ---------------------------------------------------------------------------
// Shared inline helpers
// ---------------------------------------------------------------------------

function markRank(mark: PMMark): number {
  const i = MARK_ORDER.indexOf(mark.type);
  return i < 0 ? MARK_ORDER.length : i;
}

function sortMarks(marks: PMMark[]): PMMark[] {
  return [...marks].sort((a, b) => markRank(a) - markRank(b));
}

function markKey(mark: PMMark): string {
  return mark.type === 'link' ? `link:${String(mark.attrs?.href ?? '')}` : mark.type;
}

function marksKey(marks: PMMark[] | undefined): string {
  return (marks ?? []).map(markKey).join('|');
}

/** Appends inline nodes, merging adjacent text with identical marks and dropping empty text. */
class InlineBuilder {
  readonly nodes: PMNode[] = [];

  addText(text: string, marks: PMMark[]): void {
    if (!text) return;
    const sorted = sortMarks(marks);
    const last = this.nodes[this.nodes.length - 1];
    if (last && last.type === 'text' && marksKey(last.marks) === marksKey(sorted)) {
      last.text += text;
      return;
    }
    const node: PMNode = { type: 'text', text };
    if (sorted.length) node.marks = sorted.map((m) => (m.attrs ? { type: m.type, attrs: { ...m.attrs } } : { type: m.type }));
    this.nodes.push(node);
  }

  addNode(node: PMNode): void {
    this.nodes.push(node);
  }
}

// ---------------------------------------------------------------------------
// Import
// ---------------------------------------------------------------------------

const HTML_TAG_RE = /^<(\/?)([A-Za-z][A-Za-z0-9]*)\b[^>]*?(\/?)>$/;

const HTML_MARKS: Record<string, string> = {
  u: 'underline',
  ins: 'underline',
  sup: 'superscript',
  sub: 'subscript',
  strong: 'bold',
  b: 'bold',
  em: 'italic',
  i: 'italic',
  s: 'strike',
  del: 'strike',
  strike: 'strike',
};

/** Active marks while walking an inline token stream. */
class MarkState {
  private marks: PMMark[] = [];

  get current(): PMMark[] {
    return this.marks;
  }

  add(mark: PMMark): void {
    this.marks = [...this.marks.filter((m) => m.type !== mark.type), mark];
  }

  remove(type: string): void {
    this.marks = this.marks.filter((m) => m.type !== type);
  }
}

function convertHtmlInline(content: string, marks: MarkState, out: InlineBuilder): void {
  const match = HTML_TAG_RE.exec(content.trim());
  if (match) {
    const [, closing, rawName] = match;
    const name = rawName.toLowerCase();
    if (name === 'br') {
      out.addNode({ type: 'hardBreak' });
      return;
    }
    const markType = HTML_MARKS[name];
    if (markType) {
      if (closing) marks.remove(markType);
      else marks.add({ type: markType });
      return;
    }
  }
  out.addText(content, marks.current);
}

/** Converts a markdown-it inline token list to PM inline nodes. */
function convertInline(tokens: Token[] | null): PMNode[] {
  const out = new InlineBuilder();
  const marks = new MarkState();
  for (const tok of tokens ?? []) {
    switch (tok.type) {
      case 'text':
      case 'text_special':
        out.addText(tok.content, marks.current);
        break;
      case 'softbreak':
        out.addText(' ', marks.current);
        break;
      case 'hardbreak':
        out.addNode({ type: 'hardBreak' });
        break;
      case 'strong_open':
        marks.add({ type: 'bold' });
        break;
      case 'strong_close':
        marks.remove('bold');
        break;
      case 'em_open':
        marks.add({ type: 'italic' });
        break;
      case 'em_close':
        marks.remove('italic');
        break;
      case 's_open':
        marks.add({ type: 'strike' });
        break;
      case 's_close':
        marks.remove('strike');
        break;
      case 'link_open':
        marks.add({ type: 'link', attrs: { href: tok.attrGet('href') ?? '' } });
        break;
      case 'link_close':
        marks.remove('link');
        break;
      case 'code_inline':
        out.addText(tok.content, [...marks.current, { type: 'textStyle', attrs: { fontFamily: CODE_FONT } }]);
        break;
      case 'image':
        out.addText(tok.content, marks.current);
        break;
      case 'html_inline':
        convertHtmlInline(tok.content, marks, out);
        break;
      case 'math_inline':
        out.addNode({ type: 'mathInline', attrs: { latex: tok.content } });
        break;
      default:
        if (tok.content) out.addText(tok.content, marks.current);
    }
  }
  return out.nodes;
}

function paragraph(content: PMNode[], indent: number): PMNode {
  const node: PMNode = { type: 'paragraph' };
  if (indent > 0) node.attrs = { indent: Math.min(indent, MAX_INDENT) };
  if (content.length) node.content = content;
  return node;
}

function codeParagraphs(code: string, indent: number): PMNode[] {
  const body = code.endsWith('\n') ? code.slice(0, -1) : code;
  return body.split('\n').map((line) => {
    const b = new InlineBuilder();
    b.addText(line, [{ type: 'textStyle', attrs: { fontFamily: CODE_FONT } }]);
    return paragraph(b.nodes, indent);
  });
}

function htmlBlockParagraphs(html: string, indent: number): PMNode[] {
  const trimmed = html.trim();
  // These two forms are what exportMarkdown writes for an empty paragraph / a lone hard break.
  if (/^<p>\s*<\/p>$/i.test(trimmed)) return [paragraph([], indent)];
  if (/^(<br\s*\/?>\s*)+$/i.test(trimmed)) {
    const count = trimmed.match(/<br/gi)?.length ?? 1;
    return [paragraph(Array.from({ length: count }, () => ({ type: 'hardBreak' })), indent)];
  }
  return trimmed.split('\n').map((line) => {
    const b = new InlineBuilder();
    b.addText(line, []);
    return paragraph(b.nodes, indent);
  });
}

/** Coerces arbitrary converted blocks into valid listItem content. */
function listItemContent(blocks: PMNode[]): PMNode[] {
  const out: PMNode[] = [];
  for (const block of blocks) {
    if (block.type === 'paragraph' || block.type === 'bulletList' || block.type === 'orderedList') {
      out.push(block);
    } else if (block.type === 'heading') {
      const indent = Number(block.attrs?.indent ?? 0);
      out.push(paragraph(block.content ?? [], indent));
    } else if (block.type === 'mathBlock') {
      out.push(paragraph([{ type: 'mathInline', attrs: { latex: String(block.attrs?.latex ?? '') } }], 0));
    }
  }
  if (out[0]?.type !== 'paragraph') out.unshift(paragraph([], 0));
  return out;
}

/** Walks a markdown-it block token stream with a shared cursor. */
class BlockConverter {
  private pos = 0;

  constructor(private readonly tokens: Token[]) {}

  /** Converts blocks until a token of type `closeType` (consumed) or end of stream. */
  blocks(closeType: string | null, quoteDepth: number): PMNode[] {
    const out: PMNode[] = [];
    while (this.pos < this.tokens.length) {
      const tok = this.tokens[this.pos++];
      if (tok.type === closeType) break;
      out.push(...this.block(tok, quoteDepth));
    }
    return out;
  }

  private inlineUntil(closeType: string): PMNode[] {
    let content: PMNode[] = [];
    while (this.pos < this.tokens.length) {
      const tok = this.tokens[this.pos++];
      if (tok.type === closeType) break;
      if (tok.type === 'inline') content = content.concat(convertInline(tok.children));
    }
    return mergeInline(content);
  }

  private block(tok: Token, quoteDepth: number): PMNode[] {
    switch (tok.type) {
      case 'paragraph_open':
        return [paragraph(this.inlineUntil('paragraph_close'), quoteDepth)];
      case 'heading_open': {
        const level = Math.min(3, Math.max(1, Number(tok.tag.slice(1)) || 1));
        const content = this.inlineUntil('heading_close');
        const node: PMNode = { type: 'heading', attrs: { level } };
        if (quoteDepth > 0) node.attrs!.indent = Math.min(quoteDepth, MAX_INDENT);
        if (content.length) node.content = content;
        return [node];
      }
      case 'bullet_list_open':
        return [{ type: 'bulletList', content: this.listItems('bullet_list_close', quoteDepth) }];
      case 'ordered_list_open': {
        const start = Number(tok.attrGet('start') ?? 1);
        const node: PMNode = { type: 'orderedList' };
        if (start !== 1) node.attrs = { start };
        node.content = this.listItems('ordered_list_close', quoteDepth);
        return [node];
      }
      case 'blockquote_open':
        return this.blocks('blockquote_close', quoteDepth + 1);
      case 'hr':
        return [{ type: 'horizontalRule' }];
      case 'fence':
      case 'code_block':
        return grouped(codeParagraphs(tok.content, quoteDepth));
      case 'math_block':
        return [{ type: 'mathBlock', attrs: { latex: tok.content } }];
      case 'html_block':
        return grouped(htmlBlockParagraphs(tok.content, quoteDepth));
      case 'table_open':
        return grouped(this.table(quoteDepth));
      default:
        return [];
    }
  }

  private listItems(closeType: string, quoteDepth: number): PMNode[] {
    const items: PMNode[] = [];
    while (this.pos < this.tokens.length) {
      const tok = this.tokens[this.pos++];
      if (tok.type === closeType) break;
      if (tok.type === 'list_item_open') {
        items.push({ type: 'listItem', content: listItemContent(this.blocks('list_item_close', quoteDepth)) });
      }
    }
    return items;
  }

  private table(quoteDepth: number): PMNode[] {
    const rows: PMNode[] = [];
    let cells: PMNode[][] | null = null;
    while (this.pos < this.tokens.length) {
      const tok = this.tokens[this.pos++];
      if (tok.type === 'table_close') break;
      if (tok.type === 'tr_open') cells = [];
      else if (tok.type === 'inline' && cells) cells.push(convertInline(tok.children));
      else if (tok.type === 'tr_close' && cells) {
        const joined: PMNode[] = [];
        cells.forEach((cell, i) => {
          if (i > 0) joined.push({ type: 'text', text: '\t' });
          joined.push(...cell);
        });
        rows.push(paragraph(mergeInline(joined), quoteDepth));
        cells = null;
      }
    }
    return rows;
  }
}

/** Re-merges an inline node list (e.g. after concatenating fragments). */
function mergeInline(nodes: PMNode[]): PMNode[] {
  const b = new InlineBuilder();
  for (const node of nodes) {
    if (node.type === 'text') b.addText(node.text ?? '', node.marks ?? []);
    else b.addNode(node);
  }
  return b.nodes;
}

/**
 * Parses Markdown (CommonMark + GFM tables/strikethrough + `$` math) into a Quire document.
 *
 * Args:
 *   src: The Markdown source.
 *
 * Returns:
 *   The document and default metadata.
 */
export function importMarkdown(src: string): LoadedDoc {
  const tokens = md.parse(src, {});
  const content = spaceBlocks(new BlockConverter(tokens).blocks(null, 0));
  if (!content.length) content.push({ type: 'paragraph' });
  return { doc: { type: 'doc', content }, meta: { ...DEFAULT_META } };
}

// ---------------------------------------------------------------------------
// Canonicalization (what survives a Markdown round trip)
// ---------------------------------------------------------------------------

function cleanMark(mark: PMMark): PMMark | null {
  if (mark.type === 'link') return { type: 'link', attrs: { href: String(mark.attrs?.href ?? '') } };
  if (MARK_ORDER.includes(mark.type) && mark.type !== 'textStyle') return { type: mark.type };
  return null;
}

/** Splits one run of text nodes into characters, strips delimiter marks from edge whitespace, regroups. */
function trimDelimiterWhitespace(run: PMNode[]): PMNode[] {
  const chars: { ch: string; marks: PMMark[] }[] = [];
  for (const node of run) {
    for (const ch of node.text ?? '') chars.push({ ch, marks: node.marks ?? [] });
  }
  const isWs = (i: number): boolean => isWhiteSpaceCode(chars[i].ch.charCodeAt(0));
  for (const type of DELIMITER_MARKS) {
    const has = (i: number): boolean => chars[i].marks.some((m) => m.type === type);
    const strip = (i: number): void => {
      chars[i].marks = chars[i].marks.filter((m) => m.type !== type);
    };
    let i = 0;
    while (i < chars.length) {
      if (!has(i)) {
        i++;
        continue;
      }
      let end = i;
      while (end < chars.length && has(end)) end++;
      let a = i;
      while (a < end && isWs(a)) strip(a++);
      let b = end - 1;
      while (b >= a && isWs(b)) strip(b--);
      i = end;
    }
  }
  const b = new InlineBuilder();
  for (const c of chars) b.addText(c.ch, c.marks);
  return b.nodes;
}

/** Returns the canonical form of inline content as it survives export + import. */
function canonicalInline(nodes: PMNode[]): PMNode[] {
  const out: PMNode[] = [];
  let run: PMNode[] = [];
  const flush = (): void => {
    out.push(...trimDelimiterWhitespace(run));
    run = [];
  };
  for (const node of nodes) {
    if (node.type === 'text') {
      const marks = (node.marks ?? []).map(cleanMark).filter((m): m is PMMark => m !== null);
      const text = (node.text ?? '').replace(/[\r\n]/g, ' ');
      if (text) run.push({ type: 'text', text, marks: sortMarks(marks) });
    } else if (node.type === 'hardBreak') {
      flush();
      out.push({ type: 'hardBreak' });
    } else if (node.type === 'mathInline') {
      const latex = String(node.attrs?.latex ?? '').trim();
      if (!latex) continue;
      flush();
      out.push({ type: 'mathInline', attrs: { latex } });
    }
  }
  flush();
  return out;
}

function canonicalBlock(node: PMNode): PMNode | null {
  switch (node.type) {
    case 'paragraph': {
      const content = canonicalInline(node.content ?? []);
      return content.length ? { type: 'paragraph', content } : { type: 'paragraph' };
    }
    case 'heading': {
      const level = Math.min(3, Math.max(1, Number(node.attrs?.level ?? 1)));
      const content = canonicalInline(node.content ?? []);
      const out: PMNode = { type: 'heading', attrs: { level } };
      if (content.length) out.content = content;
      return out;
    }
    case 'bulletList':
    case 'orderedList': {
      const items = (node.content ?? []).map((item) => ({
        type: 'listItem',
        content: listItemContent((item.content ?? []).map(canonicalBlock).filter((b): b is PMNode => b !== null)),
      }));
      const out: PMNode = { type: node.type, content: items };
      const start = Number(node.attrs?.start ?? 1);
      if (node.type === 'orderedList' && start !== 1) out.attrs = { start };
      return out;
    }
    case 'horizontalRule':
      return { type: 'horizontalRule' };
    case 'mathBlock':
      return { type: 'mathBlock', attrs: { latex: String(node.attrs?.latex ?? '') } };
    default:
      return null;
  }
}

/**
 * Normalizes a document to exactly what `importMarkdown(exportMarkdown(doc))` yields.
 *
 * Lossy steps: paragraph/heading alignment, indent and outline attrs are dropped; textStyle is
 * dropped; link attrs other than href are dropped; bold/italic/strike are removed from whitespace
 * at the edges of each marked run; inline math latex is trimmed (and dropped if empty); marks on
 * hard breaks / math are dropped.
 *
 * Args:
 *   doc: A Quire document.
 *
 * Returns:
 *   The canonical document.
 */
/** Lines of one code block / table / html block: these must not get spacers between them. */
const CONTINUES_GROUP = new WeakSet<PMNode>();

function grouped(nodes: PMNode[]): PMNode[] {
  nodes.slice(1).forEach((n) => CONTINUES_GROUP.add(n));
  return nodes;
}

/**
 * Quire's paragraphs have no space after them, so a markdown blank line (paragraph break) becomes an
 * explicit empty paragraph — the same thing a user types to get a gap. Headings carry their own
 * margins, so no spacer is added next to them. Export drops these spacers again.
 */
export function spaceBlocks(blocks: PMNode[]): PMNode[] {
  const out: PMNode[] = [];
  for (const b of blocks) {
    const prev = out[out.length - 1];
    if (prev && prev.type !== 'heading' && b.type !== 'heading' && !CONTINUES_GROUP.has(b)) out.push({ type: 'paragraph' });
    out.push(b);
  }
  return out;
}

/**
 * Top-level empty paragraphs are visual spacers (common in docx); markdown already separates
 * blocks with blank lines, so they're dropped rather than written as `<p></p>`.
 */
function isBlankParagraph(node: PMNode): boolean {
  return node.type === 'paragraph' && !(node.content ?? []).length;
}

export function normalizeForMarkdown(doc: PMNode): PMNode {
  const content = spaceBlocks(
    (doc.content ?? []).map(canonicalBlock).filter((b): b is PMNode => b !== null && !isBlankParagraph(b)),
  );
  if (!content.length) content.push({ type: 'paragraph' });
  return { type: 'doc', content };
}

// ---------------------------------------------------------------------------
// Export
// ---------------------------------------------------------------------------

interface InlineOptions {
  /** Headings are single-line, so hard breaks become `<br>` and every `#` is escaped. */
  heading: boolean;
  /** Write bold/italic/strike as HTML tags (fallback when delimiter flanking rules would fail). */
  htmlDelimiters: boolean;
}

const TEXT_ESCAPE_RE = /[\\*_[\]<>`~$|]/g;
const ENTITY_LIKE_RE = /&(?=#?[A-Za-z0-9]+;)/g;

function escapeText(text: string, heading: boolean): string {
  let out = text.replace(TEXT_ESCAPE_RE, '\\$&').replace(ENTITY_LIKE_RE, '\\&');
  if (heading) out = out.replace(/#/g, '\\#');
  return out;
}

function escapeHref(href: string): string {
  if (!href || /[\s<>]/.test(href) || /[\x00-\x1f]/.test(href)) {
    return `<${href.replace(/[\\<>&]/g, '\\$&')}>`;
  }
  return href.replace(/[\\()&]/g, '\\$&');
}

function openMark(mark: PMMark, opts: InlineOptions): string {
  switch (mark.type) {
    case 'bold':
      return opts.htmlDelimiters ? '<strong>' : '**';
    case 'italic':
      return opts.htmlDelimiters ? '<em>' : '*';
    case 'strike':
      return opts.htmlDelimiters ? '<s>' : '~~';
    case 'underline':
      return '<u>';
    case 'superscript':
      return '<sup>';
    case 'subscript':
      return '<sub>';
    case 'link':
      return '[';
    default:
      return '';
  }
}

function closeMark(mark: PMMark, opts: InlineOptions): string {
  switch (mark.type) {
    case 'bold':
      return opts.htmlDelimiters ? '</strong>' : '**';
    case 'italic':
      return opts.htmlDelimiters ? '</em>' : '*';
    case 'strike':
      return opts.htmlDelimiters ? '</s>' : '~~';
    case 'underline':
      return '</u>';
    case 'superscript':
      return '</sup>';
    case 'subscript':
      return '</sub>';
    case 'link':
      return `](${escapeHref(String(mark.attrs?.href ?? ''))})`;
    default:
      return '';
  }
}

/** Number of consecutive text nodes starting at `from` that carry `mark`. */
function markExtent(nodes: PMNode[], from: number, mark: PMMark): number {
  const key = markKey(mark);
  let n = 0;
  for (let i = from; i < nodes.length; i++) {
    if (nodes[i].type !== 'text' || !(nodes[i].marks ?? []).some((m) => markKey(m) === key)) break;
    n++;
  }
  return n;
}

function entity(ch: string): string {
  return `&#${ch.codePointAt(0)};`;
}

/**
 * Makes a serialized line safe from block-level interpretation and whitespace stripping.
 * (`:` matters because `:--` after a hard break is a GFM table delimiter row.)
 */
function protectLine(line: string, isLast: boolean): string {
  let out = line.replace(/^[ \t]+/, (ws) => [...ws].map(entity).join(''));
  if (out === line) {
    out = out.replace(/^[#+\-=:]/, '\\$&').replace(/^(\d{1,9})([.)])/, '$1\\$2');
  }
  if (isLast) out = out.replace(/[ \t]+$/, (ws) => [...ws].map(entity).join(''));
  return out;
}

/** Serializes canonical inline nodes using a mark stack, closing and reopening marks as needed. */
function serializeInline(nodes: PMNode[], opts: InlineOptions): string {
  let out = '';
  let stack: PMMark[] = [];
  // pandoc rules: a `$` followed by a digit cannot close math, so guard text right after math.
  let afterMath = false;
  const closeAll = (): void => {
    for (let i = stack.length - 1; i >= 0; i--) out += closeMark(stack[i], opts);
    stack = [];
  };

  nodes.forEach((node, idx) => {
    if (node.type === 'text') {
      const marks = node.marks ?? [];
      const keys = new Set(marks.map(markKey));
      let keep = 0;
      while (keep < stack.length && keys.has(markKey(stack[keep]))) keep++;
      for (let i = stack.length - 1; i >= keep; i--) out += closeMark(stack[i], opts);
      stack = stack.slice(0, keep);
      const open = new Set(stack.map(markKey));
      const toOpen = marks
        .filter((m) => !open.has(markKey(m)))
        .map((m) => ({ m, extent: markExtent(nodes, idx, m) }))
        .sort((a, b) => b.extent - a.extent || markRank(a.m) - markRank(b.m));
      for (const { m } of toOpen) {
        // escapeText never escapes `!`, so a trailing one is raw and would turn `[` into an image.
        if (m.type === 'link' && out.endsWith('!')) out = out.slice(0, -1) + '\\!';
        out += openMark(m, opts);
        stack.push(m);
      }
      let text = escapeText(node.text ?? '', opts.heading);
      if (afterMath && !toOpen.length && /^\d/.test(text)) {
        text = entity(text[0]) + text.slice(1);
      }
      out += text;
      afterMath = false;
      return;
    }
    closeAll();
    afterMath = false;
    if (node.type === 'hardBreak') {
      out += opts.heading || idx === nodes.length - 1 ? '<br>' : '\\\n';
    } else if (node.type === 'mathInline') {
      out += `$${String(node.attrs?.latex ?? '')}$`;
      afterMath = true;
    }
  });
  closeAll();

  const lines = out.split('\n');
  return lines.map((line, i) => protectLine(line, i === lines.length - 1)).join('\n');
}

function inlineJson(nodes: PMNode[]): string {
  return JSON.stringify(mergeInline(nodes));
}

/** Serializes a paragraph/heading body, falling back to HTML emphasis tags if delimiters misparse. */
function serializeTextBlock(content: PMNode[] | undefined, heading: boolean): string {
  const nodes = canonicalInline(content ?? []);
  if (!nodes.length) return heading ? '' : '<p></p>';
  const expected = inlineJson(nodes);
  const primary = serializeInline(nodes, { heading, htmlDelimiters: false });
  const parsed = md.parseInline(primary, {})[0]?.children ?? [];
  if (inlineJson(convertInline(parsed)) === expected) return primary;
  return serializeInline(nodes, { heading, htmlDelimiters: true });
}

/** True if a serialized paragraph would start a CommonMark HTML block (which swallows following lines). */
function startsHtmlBlock(text: string): boolean {
  return text === '<p></p>' || text.startsWith('<br>');
}

function indentLines(text: string, first: string, rest: string): string {
  return text
    .split('\n')
    .map((line, i) => (i === 0 ? first + line : line ? rest + line : ''))
    .join('\n');
}

/** Tracks alternate markers so adjacent sibling lists are not merged into one list. */
interface SiblingState {
  prevType: string | null;
  alt: boolean;
}

function nextListVariant(state: SiblingState, type: string): boolean {
  state.alt = state.prevType === type ? !state.alt : false;
  return state.alt;
}

function serializeList(node: PMNode, alt: boolean): string {
  const ordered = node.type === 'orderedList';
  let n = Number(node.attrs?.start ?? 1);
  const items = (node.content ?? []).map((item) => {
    const marker = ordered ? `${n++}${alt ? ')' : '.'} ` : alt ? '* ' : '- ';
    return serializeListItem(item, marker);
  });
  return items.join('\n');
}

function serializeListItem(item: PMNode, marker: string): string {
  const children = listItemContent(item.content ?? []);
  const siblings: SiblingState = { prevType: null, alt: false };
  let out = '';
  let prevText = '';
  children.forEach((child, i) => {
    const isList = child.type === 'bulletList' || child.type === 'orderedList';
    const alt = nextListVariant(siblings, child.type);
    const text = isList ? serializeList(child, alt) : serializeTextBlock(child.content, false);
    if (i > 0) {
      const prev = children[i - 1];
      // A list may follow its item's paragraph directly, except where CommonMark would treat it
      // as paragraph continuation (ordered start != 1) or an HTML block would swallow it.
      const canInterrupt =
        isList &&
        prev.type === 'paragraph' &&
        !startsHtmlBlock(prevText) &&
        (child.type === 'bulletList' || Number(child.attrs?.start ?? 1) === 1);
      out += canInterrupt ? '\n' : '\n\n';
    }
    out += text;
    prevText = text;
    siblings.prevType = child.type;
  });
  return indentLines(out, marker, ' '.repeat(marker.length));
}

function serializeBlock(node: PMNode, siblings: SiblingState): string | null {
  const alt = nextListVariant(siblings, node.type);
  siblings.prevType = node.type;
  switch (node.type) {
    case 'paragraph':
      return serializeTextBlock(node.content, false);
    case 'heading': {
      const level = Math.min(3, Math.max(1, Number(node.attrs?.level ?? 1)));
      const body = serializeTextBlock(node.content, true);
      return '#'.repeat(level) + (body ? ` ${body}` : '');
    }
    case 'bulletList':
    case 'orderedList':
      return serializeList(node, alt);
    case 'horizontalRule':
      return '---';
    case 'mathBlock': {
      const latex = String(node.attrs?.latex ?? '');
      return latex ? `$$\n${latex}\n$$` : '$$\n$$';
    }
    case 'docxRaw': {
      // Opaque Word content (tables, images): only its plain text survives, as paragraphs.
      const lines = String(node.attrs?.text ?? '').split('\n').map((l) => l.trim()).filter(Boolean);
      return lines.length ? lines.map((l) => escapeText(l.replace(/\t/g, ' '), false)).join('\n\n') : null;
    }
    default:
      return null;
  }
}

/**
 * Serializes a Quire document to Markdown that `importMarkdown` reads back identically
 * (modulo the lossy steps documented on `normalizeForMarkdown`).
 *
 * Args:
 *   doc: The document root (`{type: "doc"}`).
 *
 * Returns:
 *   Markdown text ending in a single newline.
 */
export function exportMarkdown(doc: PMNode): string {
  const siblings: SiblingState = { prevType: null, alt: false };
  const blocks = (doc.content ?? [])
    .filter((node) => !isBlankParagraph(node))
    .map((node) => serializeBlock(node, siblings))
    .filter((b): b is string => b !== null);
  return blocks.join('\n\n') + '\n';
}
