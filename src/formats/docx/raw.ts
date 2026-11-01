/** Verbatim passthrough of body elements Quire cannot edit (tables, drawings, ...). */

import { childByName, findDescendant } from './xml';

/** Element names whose presence makes a paragraph an opaque object paragraph. */
export const OBJECT_TAGS = new Set(['w:drawing', 'w:pict', 'w:object', 'w:control', 'w:embedObject']);

/** True if `el` has any drawing / picture / object / table descendant. */
export function containsObject(el: Element, includeTables: boolean): boolean {
  const kids = el.childNodes;
  for (let i = 0, n = kids.length; i < n; i++) {
    const k = kids[i];
    if (k.nodeType !== 1) continue;
    const name = k.nodeName;
    if (OBJECT_TAGS.has(name) || (includeTables && name === 'w:tbl')) return true;
    if (containsObject(k as Element, includeTables)) return true;
  }
  return false;
}

// Matches comments, CDATA, PIs, or any tag (attribute values may legally contain '>').
const TOKEN_RE = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<(\/?)[^\s>/!?]+(?:[^>"']|"[^"]*"|'[^']*')*?(\/?)>/g;

/**
 * Maps the DOM's body children back to [start, end) offsets in the source document.xml text, so
 * opaque elements can be copied byte-for-byte (namespace prefixes and all) instead of re-serialized.
 * It must be advanced once per body element child, in document order.
 */
export class BodySlicer {
  private pos: number;
  private broken = false;

  constructor(private readonly src: string) {
    const m = /<w:body\b[^>]*>/.exec(src);
    this.pos = m ? m.index + m[0].length : -1;
    if (!m || m[0].endsWith('/>')) this.broken = true;
  }

  /**
   * Advances past the next body child, which must be `el`, and returns its source text.
   *
   * Args:
   *   el: The DOM element expected next in the source.
   *   simple: True when `el` is a w:p known to contain no nested paragraphs, allowing a fast
   *     indexOf scan to its end tag instead of full tokenization.
   *
   * Returns:
   *   The verbatim outer XML, or null if the source could not be aligned with the DOM (after
   *   which every later call also returns null).
   */
  take(el: Element, simple: boolean): string | null {
    if (this.broken) return null;
    const src = this.src;
    const name = el.nodeName;
    let start = src.indexOf('<', this.pos);
    while (start >= 0 && (src.startsWith('<!--', start) || src.startsWith('<?', start))) {
      const close = src.startsWith('<!--', start) ? src.indexOf('-->', start) + 3 : src.indexOf('?>', start) + 2;
      start = close < 3 ? -1 : src.indexOf('<', close);
    }
    const after = start >= 0 ? src.charCodeAt(start + 1 + name.length) : 0;
    // Next char after the name must end it: whitespace, '>' or '/'.
    if (start < 0 || !src.startsWith(name, start + 1) || !(after === 62 || after === 47 || after <= 32)) {
      this.broken = true;
      return null;
    }
    let end: number;
    if (simple) {
      const gt = src.indexOf('>', start);
      if (src.charCodeAt(gt - 1) === 47) end = gt + 1;
      else {
        const close = src.indexOf(`</${name}>`, gt);
        end = close < 0 ? -1 : close + name.length + 3;
      }
    } else {
      end = this.scanElementEnd(start);
    }
    if (end < 0) {
      this.broken = true;
      return null;
    }
    this.pos = end;
    return src.slice(start, end);
  }

  private scanElementEnd(start: number): number {
    TOKEN_RE.lastIndex = start;
    let depth = 0;
    let m: RegExpExecArray | null;
    while ((m = TOKEN_RE.exec(this.src))) {
      if (m[1] === undefined && m[2] === undefined) continue; // comment / CDATA / PI
      if (m[1] === '/') depth--;
      else if (m[2] !== '/') depth++;
      if (depth === 0) return m.index + m[0].length;
      if (depth < 0) return -1;
    }
    return -1;
  }
}

/** Plain text of an opaque element: one line per paragraph / table row, cells tab-separated. */
export function rawText(el: Element): string {
  const lines: string[] = [];
  collectLines(el, lines);
  return lines.map((l) => l.replace(/\s+$/, '')).filter((l) => l.length > 0).join('\n');
}

function collectLines(el: Element, lines: string[]): void {
  switch (el.nodeName) {
    case 'w:p':
      lines.push(inlineText(el));
      return;
    case 'w:tbl':
      for (const row of rowsOf(el)) {
        const cells: string[] = [];
        for (const tc of childrenNamed(row, 'w:tc')) {
          const inner: string[] = [];
          collectLines(tc, inner);
          cells.push(inner.filter((s) => s.trim()).join(' '));
        }
        lines.push(cells.join('\t'));
      }
      return;
    case 'mc:AlternateContent': {
      // Choice and Fallback hold the same content; read only one.
      const pick = childByName(el, 'mc:Choice') ?? childByName(el, 'mc:Fallback');
      if (pick) collectLines(pick, lines);
      return;
    }
  }
  const kids = el.childNodes;
  for (let i = 0, n = kids.length; i < n; i++) if (kids[i].nodeType === 1) collectLines(kids[i] as Element, lines);
}

function rowsOf(tbl: Element): Element[] {
  const out: Element[] = [];
  const kids = tbl.childNodes;
  for (let i = 0, n = kids.length; i < n; i++) {
    const k = kids[i];
    if (k.nodeName === 'w:tr') out.push(k as Element);
    else if (k.nodeName === 'w:sdt' || k.nodeName === 'w:customXml') {
      const c = k.nodeName === 'w:sdt' ? findDescendant(k, 'w:sdtContent') : (k as Element);
      if (c) out.push(...rowsOf(c));
    }
  }
  return out;
}

function childrenNamed(el: Element, name: string): Element[] {
  const out: Element[] = [];
  const kids = el.childNodes;
  for (let i = 0, n = kids.length; i < n; i++) {
    const k = kids[i];
    if (k.nodeName === name) out.push(k as Element);
    else if (k.nodeName === 'w:sdt' || k.nodeName === 'w:customXml') {
      const c = k.nodeName === 'w:sdt' ? findDescendant(k, 'w:sdtContent') : (k as Element);
      if (c) out.push(...childrenNamed(c, name));
    }
  }
  return out;
}

/** Text of a paragraph's runs, including text boxes nested in drawings (read once). */
function inlineText(p: Element): string {
  const parts: string[] = [];
  const walk = (el: Element): void => {
    const kids = el.childNodes;
    for (let i = 0, n = kids.length; i < n; i++) {
      const k = kids[i];
      if (k.nodeType !== 1) continue;
      const name = k.nodeName;
      if (name === 'w:t') parts.push(k.textContent ?? '');
      else if (name === 'w:tab') parts.push('\t');
      else if (name === 'w:br' || name === 'w:cr') parts.push(' ');
      else if (name === 'w:p') {
        // Text-box paragraph inside a drawing.
        if (parts.length && !/\s$/.test(parts[parts.length - 1])) parts.push(' ');
        walk(k as Element);
      } else if (name === 'mc:AlternateContent') {
        const pick = childByName(k, 'mc:Choice') ?? childByName(k, 'mc:Fallback');
        if (pick) walk(pick);
      } else if (name !== 'w:instrText' && name !== 'w:delText' && name !== 'w:rPr' && name !== 'w:pPr') walk(k as Element);
    }
  };
  walk(p);
  return parts.join('');
}

/** Standard OOXML prefixes, used to declare namespaces raw XML relies on in a new root tag. */
export const KNOWN_NAMESPACES: Record<string, string> = {
  wpc: 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
  cx: 'http://schemas.microsoft.com/office/drawing/2014/chartex',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  aink: 'http://schemas.microsoft.com/office/drawing/2016/ink',
  am3d: 'http://schemas.microsoft.com/office/drawing/2017/model3d',
  o: 'urn:schemas-microsoft-com:office:office',
  oel: 'http://schemas.microsoft.com/office/2019/extlst',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
  v: 'urn:schemas-microsoft-com:vml',
  wp14: 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  w10: 'urn:schemas-microsoft-com:office:word',
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  w14: 'http://schemas.microsoft.com/office/word/2010/wordml',
  w15: 'http://schemas.microsoft.com/office/word/2012/wordml',
  w16cex: 'http://schemas.microsoft.com/office/word/2018/wordml/cex',
  w16cid: 'http://schemas.microsoft.com/office/word/2016/wordml/cid',
  w16: 'http://schemas.microsoft.com/office/word/2018/wordml',
  w16du: 'http://schemas.microsoft.com/office/word/2023/wordml/word16du',
  w16sdtdh: 'http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash',
  w16sdtfl: 'http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock',
  w16se: 'http://schemas.microsoft.com/office/word/2015/wordml/symex',
  wpg: 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
  wpi: 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk',
  wne: 'http://schemas.microsoft.com/office/word/2006/wordml',
  wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  dgm: 'http://schemas.openxmlformats.org/drawingml/2006/diagram',
};

/** Prefixes used by `xml` but not declared inside it (they must come from the root tag). */
export function undeclaredPrefixes(xml: string): Set<string> {
  const used = new Set<string>();
  for (const m of xml.matchAll(/<\/?([A-Za-z_][\w.-]*):|\s([A-Za-z_][\w.-]*):[\w.-]+\s*=/g)) used.add(m[1] ?? m[2]);
  for (const m of xml.matchAll(/xmlns:([\w.-]+)\s*=/g)) used.delete(m[1]);
  used.delete('xml');
  used.delete('xmlns');
  return used;
}

/** Relationship ids referenced by `r:*` attributes (r:id, r:embed, r:link, r:pict, ...). */
export function relIdsIn(xml: string, out: Set<string>): void {
  for (const m of xml.matchAll(/\sr:[\w]+="([^"]+)"/g)) out.add(m[1]);
}
