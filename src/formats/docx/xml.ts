/** Small portable DOM / XML string helpers shared by the DOCX importer and exporter. */

const ELEMENT_NODE = 1;

/** Returns the element children of `node` (text, comments, etc. are skipped). */
export function elementChildren(node: Node): Element[] {
  const out: Element[] = [];
  const kids = node.childNodes;
  for (let i = 0, n = kids.length; i < n; i++) {
    const k = kids[i];
    if (k.nodeType === ELEMENT_NODE) out.push(k as Element);
  }
  return out;
}

/** Returns the first element child of `node` with qualified name `name`, or null. */
export function childByName(node: Node, name: string): Element | null {
  const kids = node.childNodes;
  for (let i = 0, n = kids.length; i < n; i++) {
    const k = kids[i];
    if (k.nodeType === ELEMENT_NODE && k.nodeName === name) return k as Element;
  }
  return null;
}

/** Depth-first search for the first descendant element named `name`. */
export function findDescendant(node: Node, name: string): Element | null {
  const kids = node.childNodes;
  for (let i = 0, n = kids.length; i < n; i++) {
    const k = kids[i];
    if (k.nodeType !== ELEMENT_NODE) continue;
    if (k.nodeName === name) return k as Element;
    const hit = findDescendant(k, name);
    if (hit) return hit;
  }
  return null;
}

/** Collects all descendant elements named `name` in document order. */
export function findAllDescendants(node: Node, name: string, out: Element[] = []): Element[] {
  const kids = node.childNodes;
  for (let i = 0, n = kids.length; i < n; i++) {
    const k = kids[i];
    if (k.nodeType !== ELEMENT_NODE) continue;
    if (k.nodeName === name) out.push(k as Element);
    else findAllDescendants(k, name, out);
  }
  return out;
}

/** Reads an attribute, normalizing "missing" (null in DOM, "" in some parsers) to null. */
export function attr(el: Element, name: string): string | null {
  const v = el.getAttribute(name);
  return v === null || v === '' ? null : v;
}

/** Interprets an OOXML on/off property element (`<w:b/>`, `<w:b w:val="0"/>`, ...). */
export function isOn(el: Element): boolean {
  const v = attr(el, 'w:val');
  if (v === null) return true;
  return !(v === '0' || v === 'false' || v === 'off' || v === 'none');
}

/** Parses an integer attribute, returning null when absent or malformed. */
export function intAttr(el: Element | null, name: string): number | null {
  if (!el) return null;
  const v = attr(el, name);
  if (v === null) return null;
  const n = parseInt(v, 10);
  return Number.isFinite(n) ? n : null;
}

// Control chars and unpaired surrogates are illegal in XML 1.0 and make Word refuse the file.
const INVALID_XML_CHARS =
  /[\x00-\x08\x0B\x0C\x0E-\x1F￾￿]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g;
const ESCAPE_RE = /[&<>"]/g;
const ESCAPES: Record<string, string> = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' };

/** Escapes text for use in XML element content or a double-quoted attribute. */
export function escapeXml(s: string): string {
  return s.replace(INVALID_XML_CHARS, '').replace(ESCAPE_RE, (c) => ESCAPES[c]);
}

/** Reverses the standard XML entity escapes (used on regex-extracted attribute values). */
export function unescapeXml(s: string): string {
  if (s.indexOf('&') < 0) return s;
  return s.replace(/&(amp|lt|gt|quot|apos|#x[0-9a-fA-F]+|#\d+);/g, (_m, e: string) => {
    switch (e) {
      case 'amp': return '&';
      case 'lt': return '<';
      case 'gt': return '>';
      case 'quot': return '"';
      case 'apos': return "'";
      default:
        return String.fromCodePoint(e[1] === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10));
    }
  });
}

/** Parses the attributes of a single XML start tag string into a map. */
export function parseTagAttrs(tag: string): Map<string, string> {
  const out = new Map<string, string>();
  const re = /([\w:.-]+)\s*=\s*("([^"]*)"|'([^']*)')/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(tag))) out.set(m[1], unescapeXml(m[3] ?? m[4] ?? ''));
  return out;
}

export interface Relationship {
  id: string;
  type: string;
  target: string;
  targetMode: string | null;
  /** The raw `<Relationship .../>` tag, kept so untouched rels are written back verbatim. */
  raw: string;
}

/** Parses a .rels part with regexes (exporter has no DOM parser available). */
export function parseRels(xml: string): Relationship[] {
  const out: Relationship[] = [];
  const re = /<Relationship\b[^>]*?\/?>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml))) {
    const a = parseTagAttrs(m[0]);
    out.push({
      id: a.get('Id') ?? '',
      type: a.get('Type') ?? '',
      target: a.get('Target') ?? '',
      targetMode: a.get('TargetMode') ?? null,
      raw: m[0],
    });
  }
  return out;
}

/** Resolves a relationship target relative to the directory of its source part. */
export function resolvePartPath(sourceDir: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const parts = (sourceDir ? sourceDir + '/' + target : target).split('/');
  const out: string[] = [];
  for (const p of parts) {
    if (p === '..') out.pop();
    else if (p !== '.' && p !== '') out.push(p);
  }
  return out.join('/');
}

/** Returns the .rels path for a part, e.g. word/document.xml -> word/_rels/document.xml.rels. */
export function relsPathFor(partPath: string): string {
  const i = partPath.lastIndexOf('/');
  return i < 0 ? `_rels/${partPath}.rels` : `${partPath.slice(0, i)}/_rels/${partPath.slice(i + 1)}.rels`;
}

/** Returns the directory portion of a part path ("" for root-level parts). */
export function dirOf(partPath: string): string {
  const i = partPath.lastIndexOf('/');
  return i < 0 ? '' : partPath.slice(0, i);
}

export const REL_TYPES = {
  officeDocument: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  styles: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
  numbering: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering',
  settings: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings',
  theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  hyperlink: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
  coreProps: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
  extendedProps: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
};

/** Matches a relationship type by its last path segment, tolerating strict-namespace variants. */
export function relTypeIs(type: string, kind: string): boolean {
  return type.endsWith('/' + kind);
}
