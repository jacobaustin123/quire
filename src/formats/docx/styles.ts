/** Parsing of styles.xml, theme1.xml and numbering.xml into resolved property lookups. */

import { attr, childByName, elementChildren, findDescendant, intAttr, isOn } from './xml';

/** Run properties relevant to Quire. Undefined fields inherit from the next layer down. */
export interface RunProps {
  bold?: boolean;
  italic?: boolean;
  strike?: boolean;
  underline?: boolean;
  vertAlign?: 'superscript' | 'subscript' | 'baseline';
  font?: string;
  /** Half-points, as stored in `w:sz`. */
  size?: number;
}

/** Paragraph properties relevant to Quire. */
export interface ParaProps {
  styleId?: string;
  jc?: string;
  indLeft?: number;
  outlineLvl?: number;
  numId?: string;
  ilvl?: number;
  /** True when numPr explicitly carried an ilvl (used to detect list continuation paragraphs). */
  hasIlvl?: boolean;
  borderBottom?: boolean;
  spacingAfter?: number;
  /** Line spacing as a multiple of single; exact/atLeast rules count as 1. */
  lineMultiple?: number;
}

export interface ThemeFonts {
  major: string | null;
  minor: string | null;
}

/** Extracts the major/minor latin typefaces from a theme part. */
export function parseTheme(themeDoc: Document | null): ThemeFonts {
  const out: ThemeFonts = { major: null, minor: null };
  if (!themeDoc) return out;
  const scheme = findDescendant(themeDoc.documentElement, 'a:fontScheme');
  if (!scheme) return out;
  for (const which of ['major', 'minor'] as const) {
    const f = childByName(scheme, which === 'major' ? 'a:majorFont' : 'a:minorFont');
    const latin = f && childByName(f, 'a:latin');
    const face = latin && attr(latin, 'typeface');
    if (face) out[which] = face;
  }
  return out;
}

/** Parses a `w:rPr` element. Unknown / irrelevant properties are ignored. */
export function parseRunProps(rPr: Element, theme: ThemeFonts): RunProps {
  const p: RunProps = {};
  const kids = rPr.childNodes;
  for (let i = 0, n = kids.length; i < n; i++) {
    const k = kids[i];
    if (k.nodeType !== 1) continue;
    const el = k as Element;
    switch (el.nodeName) {
      case 'w:b': p.bold = isOn(el); break;
      case 'w:i': p.italic = isOn(el); break;
      case 'w:strike': p.strike = isOn(el); break;
      case 'w:u': {
        // `<w:u w:color="..."/>` without a val only sets an underline color; it is not an underline.
        const v = attr(el, 'w:val');
        if (v !== null) p.underline = v !== 'none' && v !== '0' && v !== 'false';
        break;
      }
      case 'w:vertAlign': {
        const v = attr(el, 'w:val');
        if (v === 'superscript' || v === 'subscript' || v === 'baseline') p.vertAlign = v;
        break;
      }
      case 'w:sz': {
        const v = intAttr(el, 'w:val');
        if (v !== null && v > 0) p.size = v;
        break;
      }
      case 'w:rFonts': {
        const f = fontFromRFonts(el, theme);
        if (f) p.font = f;
        break;
      }
    }
  }
  return p;
}

function fontFromRFonts(el: Element, theme: ThemeFonts): string | null {
  // Theme attributes take precedence over explicit names in Word.
  const t = attr(el, 'w:asciiTheme') ?? attr(el, 'w:hAnsiTheme');
  if (t) {
    const f = t.startsWith('major') ? theme.major : theme.minor;
    if (f) return f;
  }
  return attr(el, 'w:ascii') ?? attr(el, 'w:hAnsi');
}

/** Parses a `w:pPr` element. `w:rPr` (paragraph mark props) and `w:sectPr` are ignored. */
export function parseParaProps(pPr: Element): ParaProps {
  const p: ParaProps = {};
  const kids = pPr.childNodes;
  for (let i = 0, n = kids.length; i < n; i++) {
    const k = kids[i];
    if (k.nodeType !== 1) continue;
    const el = k as Element;
    switch (el.nodeName) {
      case 'w:pStyle': {
        const v = attr(el, 'w:val');
        if (v) p.styleId = v;
        break;
      }
      case 'w:jc': {
        const v = attr(el, 'w:val');
        if (v) p.jc = v;
        break;
      }
      case 'w:ind': {
        const v = intAttr(el, 'w:left') ?? intAttr(el, 'w:start');
        if (v !== null) p.indLeft = v;
        break;
      }
      case 'w:outlineLvl': {
        const v = intAttr(el, 'w:val');
        if (v !== null) p.outlineLvl = v;
        break;
      }
      case 'w:spacing': {
        const v = intAttr(el, 'w:after');
        if (v !== null) p.spacingAfter = v;
        const line = intAttr(el, 'w:line');
        if (line !== null) {
          const rule = attr(el, 'w:lineRule');
          p.lineMultiple = (rule === null || rule === 'auto') && line > 0 ? line / 240 : 1;
        }
        break;
      }
      case 'w:numPr': {
        const ilvl = intAttr(childByName(el, 'w:ilvl'), 'w:val');
        const numIdEl = childByName(el, 'w:numId');
        const numId = numIdEl ? attr(numIdEl, 'w:val') : null;
        if (ilvl !== null) {
          p.ilvl = ilvl;
          p.hasIlvl = true;
        }
        if (numId !== null) p.numId = numId;
        break;
      }
      case 'w:pBdr': {
        const b = childByName(el, 'w:bottom');
        const v = b ? attr(b, 'w:val') : null;
        if (v !== null && v !== 'nil' && v !== 'none') p.borderBottom = true;
        break;
      }
    }
  }
  return p;
}

interface StyleDef {
  id: string;
  name: string;
  type: string;
  basedOn: string | null;
  rPr: RunProps;
  pPr: ParaProps;
}

/** Resolved view of styles.xml (docDefaults + basedOn chains), memoized per style id. */
export class StyleSheet {
  readonly defaultRun: RunProps = {};
  readonly defaultPara: ParaProps = {};
  readonly defaultParaStyleId: string | null = null;
  private readonly defs = new Map<string, StyleDef>();
  private readonly runCache = new Map<string, RunProps>();
  private readonly paraCache = new Map<string, ParaProps>();

  constructor(stylesDoc: Document | null, theme: ThemeFonts) {
    if (!stylesDoc) return;
    for (const el of elementChildren(stylesDoc.documentElement)) {
      if (el.nodeName === 'w:docDefaults') {
        const rPr = findDescendant(el, 'w:rPr');
        if (rPr) this.defaultRun = parseRunProps(rPr, theme);
        const pPrDef = childByName(el, 'w:pPrDefault');
        const pPr = pPrDef && childByName(pPrDef, 'w:pPr');
        if (pPr) this.defaultPara = parseParaProps(pPr);
      } else if (el.nodeName === 'w:style') {
        const id = attr(el, 'w:styleId');
        if (!id) continue;
        const type = attr(el, 'w:type') ?? 'paragraph';
        const nameEl = childByName(el, 'w:name');
        const basedOnEl = childByName(el, 'w:basedOn');
        const rPr = childByName(el, 'w:rPr');
        const pPr = childByName(el, 'w:pPr');
        this.defs.set(id, {
          id,
          type,
          name: (nameEl && attr(nameEl, 'w:val')) ?? id,
          basedOn: basedOnEl ? attr(basedOnEl, 'w:val') : null,
          rPr: rPr ? parseRunProps(rPr, theme) : {},
          pPr: pPr ? parseParaProps(pPr) : {},
        });
        if (type === 'paragraph' && isOn2(attr(el, 'w:default')) && this.defaultParaStyleId === null) {
          this.defaultParaStyleId = id;
        }
      }
    }
    if (this.defaultParaStyleId === null && this.defs.has('Normal')) this.defaultParaStyleId = 'Normal';
  }

  /** Returns the merged rPr of a style's basedOn chain (root first). Excludes docDefaults. */
  styleRun(id: string | null | undefined): RunProps {
    if (!id) return {};
    let hit = this.runCache.get(id);
    if (hit) return hit;
    hit = {};
    for (const d of this.chain(id)) Object.assign(hit, d.rPr);
    this.runCache.set(id, hit);
    return hit;
  }

  /** Returns the merged pPr of a style's basedOn chain. Excludes docDefaults. */
  stylePara(id: string | null | undefined): ParaProps {
    if (!id) return {};
    let hit = this.paraCache.get(id);
    if (hit) return hit;
    hit = {};
    for (const d of this.chain(id)) {
      const { styleId: _ignored, borderBottom: _bb, ...rest } = d.pPr;
      Object.assign(hit, rest);
    }
    this.paraCache.set(id, hit);
    return hit;
  }

  /** Heading level for paragraph styles named "heading N" / "Title" (levels > 3 clamp to 3). */
  headingLevel(id: string | null | undefined): number | null {
    if (!id) return null;
    const d = this.defs.get(id);
    const candidates = d ? [d.name, d.id] : [id];
    for (const c of candidates) {
      const m = /^heading\s*(\d)$/i.exec(c);
      if (m) {
        const lvl = parseInt(m[1], 10);
        return lvl < 1 ? 1 : Math.min(lvl, 3);
      }
      if (/^title$/i.test(c)) return 1;
    }
    return null;
  }

  /** Numbering-style lookup: the numId a `w:type="numbering"` style points at. */
  numIdOfStyle(id: string): string | null {
    return this.defs.get(id)?.pPr.numId ?? null;
  }

  private chain(id: string): StyleDef[] {
    const out: StyleDef[] = [];
    const seen = new Set<string>();
    let cur: string | null = id;
    while (cur && !seen.has(cur)) {
      seen.add(cur);
      const d = this.defs.get(cur);
      if (!d) break;
      out.push(d);
      cur = d.basedOn;
    }
    return out.reverse();
  }
}

function isOn2(v: string | null): boolean {
  return v === '1' || v === 'true' || v === 'on';
}

export interface LevelDef {
  fmt: string;
  start: number;
}

interface AbstractDef {
  levels: Map<number, LevelDef>;
  numStyleLink: string | null;
}

interface NumDef {
  abstractId: string;
  startOverrides: Map<number, number>;
  lvlOverrides: Map<number, LevelDef>;
}

/** Resolved view of numbering.xml: (numId, ilvl) -> format + start. */
export class Numbering {
  private readonly abstracts = new Map<string, AbstractDef>();
  private readonly nums = new Map<string, NumDef>();
  private readonly cache = new Map<string, LevelDef | null>();

  constructor(numDoc: Document, private readonly styles: StyleSheet) {
    for (const el of elementChildren(numDoc.documentElement)) {
      if (el.nodeName === 'w:abstractNum') {
        const id = attr(el, 'w:abstractNumId');
        if (id === null) continue;
        const levels = new Map<number, LevelDef>();
        let link: string | null = null;
        for (const c of elementChildren(el)) {
          if (c.nodeName === 'w:lvl') {
            const l = parseLevel(c);
            if (l) levels.set(l[0], l[1]);
          } else if (c.nodeName === 'w:numStyleLink') {
            link = attr(c, 'w:val');
          }
        }
        this.abstracts.set(id, { levels, numStyleLink: link });
      } else if (el.nodeName === 'w:num') {
        const id = attr(el, 'w:numId');
        const absEl = childByName(el, 'w:abstractNumId');
        const abstractId = absEl ? attr(absEl, 'w:val') : null;
        if (id === null || abstractId === null) continue;
        const def: NumDef = { abstractId, startOverrides: new Map(), lvlOverrides: new Map() };
        for (const o of elementChildren(el)) {
          if (o.nodeName !== 'w:lvlOverride') continue;
          const ilvl = intAttr(o, 'w:ilvl') ?? 0;
          const so = intAttr(childByName(o, 'w:startOverride'), 'w:val');
          if (so !== null) def.startOverrides.set(ilvl, so);
          const lvl = childByName(o, 'w:lvl');
          const l = lvl && parseLevel(lvl);
          if (l) def.lvlOverrides.set(ilvl, l[1]);
        }
        this.nums.set(id, def);
      }
    }
  }

  /** Returns the level definition for a list paragraph, or null when the numId is unknown. */
  level(numId: string, ilvl: number): LevelDef | null {
    const key = numId + ':' + ilvl;
    const cached = this.cache.get(key);
    if (cached !== undefined) return cached;
    const num = this.nums.get(numId);
    let out: LevelDef | null = null;
    if (num) {
      const base = num.lvlOverrides.get(ilvl) ?? this.abstractLevel(num.abstractId, ilvl, 0);
      out = { fmt: base?.fmt ?? 'decimal', start: num.startOverrides.get(ilvl) ?? base?.start ?? 1 };
    }
    this.cache.set(key, out);
    return out;
  }

  private abstractLevel(abstractId: string, ilvl: number, depth: number): LevelDef | null {
    const abs = this.abstracts.get(abstractId);
    if (!abs) return null;
    if (abs.levels.size === 0 && abs.numStyleLink && depth < 5) {
      // numStyleLink -> numbering style -> numId -> the abstractNum that actually holds levels.
      const linkedNum = this.styles.numIdOfStyle(abs.numStyleLink);
      const n = linkedNum ? this.nums.get(linkedNum) : undefined;
      if (n) return n.lvlOverrides.get(ilvl) ?? this.abstractLevel(n.abstractId, ilvl, depth + 1);
      return null;
    }
    return abs.levels.get(ilvl) ?? null;
  }
}

function parseLevel(lvl: Element): [number, LevelDef] | null {
  const ilvl = intAttr(lvl, 'w:ilvl');
  if (ilvl === null) return null;
  const fmtEl = childByName(lvl, 'w:numFmt');
  const fmt = (fmtEl && attr(fmtEl, 'w:val')) ?? 'decimal';
  const start = intAttr(childByName(lvl, 'w:start'), 'w:val') ?? 1;
  return [ilvl, { fmt, start }];
}
