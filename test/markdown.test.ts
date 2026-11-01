import { describe, expect, it } from 'vitest';
import { exportMarkdown, importMarkdown, normalizeForMarkdown, spaceBlocks } from '../src/formats/markdown';
import { DEFAULT_META, type PMMark, type PMNode } from '../src/formats/model';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const MARK_ORDER = ['link', 'bold', 'italic', 'strike', 'underline', 'superscript', 'subscript', 'textStyle'];

/** Drops default attrs and sorts marks so docs built by hand compare structurally. */
function norm(node: PMNode): PMNode {
  const out: PMNode = { type: node.type };
  const attrs: Record<string, unknown> = { ...(node.attrs ?? {}) };
  if (attrs.textAlign === 'left') delete attrs.textAlign;
  if (attrs.indent === 0) delete attrs.indent;
  if (attrs.outline === null) delete attrs.outline;
  if (node.type === 'orderedList' && attrs.start === 1) delete attrs.start;
  if (Object.keys(attrs).length) out.attrs = attrs;
  if (node.text !== undefined) out.text = node.text;
  if (node.marks?.length) {
    out.marks = [...node.marks]
      .sort((a, b) => MARK_ORDER.indexOf(a.type) - MARK_ORDER.indexOf(b.type))
      .map((m) => (m.type === 'link' ? { type: 'link', attrs: { href: m.attrs?.href } } : m));
  }
  if (node.content?.length) out.content = node.content.map(norm);
  return out;
}

function roundTrip(doc: PMNode): PMNode {
  return importMarkdown(exportMarkdown(doc)).doc;
}

/** What a top-level block list looks like after import: spacer paragraphs between blocks. */
function spaced(doc: PMNode): PMNode {
  const blocks = (doc.content ?? []).filter((b) => !(b.type === 'paragraph' && !b.content?.length));
  return { ...doc, content: blocks.length ? spaceBlocks(blocks) : [{ type: 'paragraph' }] };
}

function expectRoundTrip(doc: PMNode): void {
  const md = exportMarkdown(doc);
  expect(importMarkdown(md).doc, `markdown was:\n${md}`).toEqual(norm(spaced(doc)));
}

const t = (text: string, ...marks: (string | PMMark)[]): PMNode => {
  const node: PMNode = { type: 'text', text };
  if (marks.length) node.marks = marks.map((m) => (typeof m === 'string' ? { type: m } : m));
  return node;
};
const link = (href: string): PMMark => ({ type: 'link', attrs: { href } });
const p = (...content: PMNode[]): PMNode => (content.length ? { type: 'paragraph', content } : { type: 'paragraph' });
const h = (level: number, ...content: PMNode[]): PMNode => ({ type: 'heading', attrs: { level }, content });
const br: PMNode = { type: 'hardBreak' };
const mi = (latex: string): PMNode => ({ type: 'mathInline', attrs: { latex } });
const mb = (latex: string): PMNode => ({ type: 'mathBlock', attrs: { latex } });
const hr: PMNode = { type: 'horizontalRule' };
const li = (...content: PMNode[]): PMNode => ({ type: 'listItem', content });
const ul = (...items: PMNode[]): PMNode => ({ type: 'bulletList', content: items });
const ol = (start: number, ...items: PMNode[]): PMNode => ({ type: 'orderedList', attrs: { start }, content: items });
const doc = (...content: PMNode[]): PMNode => ({ type: 'doc', content });

function plainText(node: PMNode): string {
  if (node.type === 'text') return node.text ?? '';
  if (node.type === 'mathInline') return `$${String(node.attrs?.latex ?? '').trim()}$`;
  return (node.content ?? []).map(plainText).join('');
}

// ---------------------------------------------------------------------------
// Hand-built round trips
// ---------------------------------------------------------------------------

describe('markdown round trip (hand-built)', () => {
  const cases: [string, PMNode][] = [
    ['plain paragraphs', doc(p(t('Hello world.')), p(t('Second paragraph.')))],
    ['headings', doc(h(1, t('One')), h(2, t('Two', 'bold'), t(' x')), h(3, t('Three #hash #')), { type: 'heading', attrs: { level: 2 } })],
    ['every mark', doc(p(
      t('b', 'bold'), t(' '), t('i', 'italic'), t(' '), t('s', 'strike'), t(' '), t('u', 'underline'), t(' '),
      t('sup', 'superscript'), t(' '), t('sub', 'subscript'), t(' '), t('link', link('https://example.com/a?b=c&d=e')),
    ))],
    ['bold+italic overlap', doc(p(t('a', 'bold'), t('b', 'bold', 'italic'), t('c', 'italic')))],
    ['italic then bold adjacent', doc(p(t('a', 'italic'), t('b', 'bold')))],
    ['nested marks', doc(p(t('x ', 'bold'), t('y', 'bold', 'italic', 'strike'), t(' z', 'bold')))],
    ['whitespace inside underline', doc(p(t('a'), t(' u ', 'underline'), t('b')))],
    ['links with marks inside', doc(p(t('see '), t('the ', link('http://x.y')), t('docs', link('http://x.y'), 'bold'), t(' now')))],
    ['link with parens and spaces', doc(p(t('a', link('http://x/(y)')), t(' '), t('b', link('file name.md')), t(' '), t('c', link('a<b>'))))],
    ['adjacent links', doc(p(t('a', link('http://a')), t('b', link('http://b'))))],
    ['escaping specials', doc(p(t('*not bold* _not italic_ [not](link) <not html> `not code` ~~no~~ \\ back | pipe')))],
    ['line-start #', doc(p(t('# not a heading')), p(t('#hashtag')))],
    ['line-start list markers', doc(p(t('1. not a list')), p(t('2) nope')), p(t('- dash')), p(t('+ plus')), p(t('> quote')), p(t('---')), p(t('===')))],
    ['currency', doc(p(t('$5 and $6')), p(t('I have $8500 and $20')), p(t('$x$ is not math $$y$$')))],
    ['entities stay literal', doc(p(t('&amp; &#32; AT&T &copy;')))],
    ['leading/trailing whitespace and tabs', doc(p(t('  two spaces')), p(t('\ttab start')), p(t('trailing  ')), p(t('mid\ttab   and spaces')), p(t('    code?')))],
    ['hard breaks', doc(p(t('a'), br, t('b')), p(br, t('starts')), p(t('ends'), br), p(br), p(br, br), p(t('x'), br, br, t('y')))],
    ['hard break with whitespace and line-start specials', doc(p(t('a '), br, t(' b'), br, t('# c'), br, t('1. d'), br, t('- e')))],
    ['line after break looks like a table delimiter row', doc(p(t('a|b'), br, t(':--')), p(t('x'), br, t('|---|')))],
    ['hard break in heading', doc(h(2, t('a'), br, t('b')))],
    ['math', doc(p(t('Euler: '), mi('e^{i\\pi}+1=0'), t('.')), mb('\\int_0^1 x\\,dx'), mb('a\\\\\nb'), p(mi('x'), mi('y')), p(mi('n'), t('5 items')))],
    ['math next to money', doc(p(t('cost $5 for '), mi('x'), t(' and $6')))],
    ['hr', doc(p(t('a')), hr, p(t('b')))],
    ['simple lists', doc(ul(li(p(t('one'))), li(p(t('two')))), ol(1, li(p(t('first'))), li(p(t('second')))))],
    ['ordered start', doc(ol(7, li(p(t('seven'))), li(p(t('eight'))), li(p(t('nine'))), li(p(t('ten')))))],
    ['adjacent lists of the same type', doc(ul(li(p(t('a')))), ul(li(p(t('b')))), ol(1, li(p(t('c')))), ol(4, li(p(t('d')))), ul(li(p(t('e')))))],
    ['nested mixed lists with paragraphs', doc(
      ul(
        li(p(t('outer')), ol(3, li(p(t('inner three')), p(t('second para'))), li(p(t('inner four')), ul(li(p(t('deep')))))), p(t('after inner'))),
        li(p(t('outer 2')), ul(li(p(t('x')))), ul(li(p(t('y'))))),
      ),
      p(t('tail')),
    )],
    ['list item paragraphs + breaks', doc(ol(9, li(p(t('a'), br, t('b')), p(t('c'), br, t('d'))), li(p(t('e')))))],
    ['empty list item paragraphs', doc(ul(li(p()), li(p(), ul(li(p(t('n'))))), li(p(br), p(t('z')))))],
    ['list item text looks like markup', doc(ul(li(p(t('1. x'))), li(p(t('- y'))), li(p(t('  lead'))), li(p(t('# h')))))],
    ['punctuation inside emphasis', doc(p(t('a'), t('(b)', 'bold'), t('c')), p(t('x'), t('.', 'italic'), t('y')), p(t('*', 'bold')), p(t('_', 'italic'), t('_')))],
    ['marks around math and breaks', doc(p(t('a', 'bold'), mi('x'), t('b', 'bold'), br, t('c', 'bold')))],
  ];
  for (const [name, d] of cases) {
    it(name, () => expectRoundTrip(d));
  }

  it('turns blank lines into spacer paragraphs (not next to headings)', () => {
    expect(importMarkdown('# H\n\na\n\nb\n\n- x').doc).toEqual(norm(doc(
      { type: 'heading', attrs: { level: 1 }, content: [t('H')] }, p(t('a')), p(), p(t('b')), p(), ul(li(p(t('x')))),
    )));
  });

  it('drops blank top-level paragraphs (docx-style spacers)', () => {
    const d = doc(p(), p(t('a')), p(), p(), p(t('b')));
    expect(exportMarkdown(d)).toBe('a\n\nb\n');
    expect(roundTrip(d)).toEqual(norm(doc(p(t('a')), p(), p(t('b')))));
  });

  it('escapes a raw ! before a link', () => {
    expectRoundTrip(doc(p(t('wow!'), t('x', link('http://x')))));
  });

  it('moves whitespace outside delimiters', () => {
    const d = doc(p(t('foo ', 'bold'), t('bar'), t(' baz ', 'italic'), t('!')));
    const md = exportMarkdown(d);
    expect(md).toBe('**foo** bar *baz* !\n');
    expect(roundTrip(d)).toEqual(doc(p(t('foo', 'bold'), t(' bar '), t('baz', 'italic'), t(' !'))));
  });

  it('writes the expected surface syntax', () => {
    const d = doc(
      h(1, t('Title')),
      p(t('a', 'bold'), t(' '), t('b', 'italic'), t(' '), t('c', 'strike'), t(' '), t('d', 'underline'), t(' '), t('e', link('http://e'))),
      ul(li(p(t('x')), ol(2, li(p(t('y'))))), li(p(t('z')))),
      ol(1, li(p(t('p'), br, t('q')))),
      hr,
      mb('x^2'),
      p(mi('y')),
    );
    expect(exportMarkdown(d)).toBe(
      [
        '# Title',
        '',
        '**a** *b* ~~c~~ <u>d</u> [e](http://e)',
        '',
        '- x',
        '',
        '  2. y',
        '- z',
        '',
        '1. p\\',
        '   q',
        '',
        '---',
        '',
        '$$',
        'x^2',
        '$$',
        '',
        '$y$',
        '',
      ].join('\n'),
    );
  });

  it('drops paragraph attrs and textStyle', () => {
    const d = doc(
      { type: 'paragraph', attrs: { textAlign: 'center', indent: 2, outline: 1 }, content: [t('x', { type: 'textStyle', attrs: { fontFamily: 'Menlo', fontSize: null } })] },
    );
    expect(roundTrip(d)).toEqual(doc(p(t('x'))));
    expect(normalizeForMarkdown(d)).toEqual(doc(p(t('x'))));
  });
});

// ---------------------------------------------------------------------------
// Import
// ---------------------------------------------------------------------------

describe('markdown import', () => {
  const imp = (s: string): PMNode => importMarkdown(s).doc;
  /** Import with the inter-block spacer paragraphs removed, to check block structure alone. */
  const impFlat = (s: string): PMNode => {
    const d = importMarkdown(s).doc;
    return { ...d, content: (d.content ?? []).filter((b) => !(b.type === 'paragraph' && !b.content?.length)) };
  };

  it('returns default meta and a non-empty doc', () => {
    const loaded = importMarkdown('');
    expect(loaded.meta).toEqual(DEFAULT_META);
    expect(loaded.doc).toEqual(doc(p()));
  });

  it('parses a GitHub README', () => {
    const src = [
      '# Project',
      '',
      '[![build](https://img.shields.io/badge.svg)](https://ci) A *tiny* tool.',
      'Second line of the same paragraph.',
      '',
      '## Install',
      '',
      '```sh',
      'npm install foo',
      '',
      'npm test',
      '```',
      '',
      '#### Deep heading',
      '',
      '* item one',
      '* item `two`',
      '  1. sub',
      '  2. sub2',
      '',
      '> **Note**',
      '> quoted',
      '',
      '| a | b |',
      '|---|---|',
      '| 1 | **2** |',
      '',
      '***',
      '',
      '<div align="center">hi</div>',
      '',
      'Line with trailing spaces  ',
      'next',
    ].join('\n');
    const code = (s: string): PMNode => t(s, { type: 'textStyle', attrs: { fontFamily: 'Menlo' } });
    expect(imp(src)).toEqual(doc(
      h(1, t('Project')),
      p(t('build', link('https://ci')), t(' A '), t('tiny', 'italic'), t(' tool. Second line of the same paragraph.')),
      h(2, t('Install')),
      p(code('npm install foo')),
      p(),
      p(code('npm test')),
      h(3, t('Deep heading')),
      ul(li(p(t('item one'))), li(p(t('item '), code('two')), { type: 'orderedList', content: [li(p(t('sub'))), li(p(t('sub2')))] })),
      p(),
      { type: 'paragraph', attrs: { indent: 1 }, content: [t('Note', 'bold'), t(' quoted')] },
      p(),
      p(t('a\tb')),
      p(t('1\t'), t('2', 'bold')),
      p(),
      hr,
      p(),
      p(t('<div align="center">hi</div>')),
      p(),
      p(t('Line with trailing spaces'), br, t('next')),
    ));
  });

  it('maps html inline tags to marks and keeps others literal', () => {
    expect(imp('H<sub>2</sub>O x<sup>2</sup> <u>under</u> <kbd>K</kbd> a<br>b')).toEqual(doc(p(
      t('H'), t('2', 'subscript'), t('O x'), t('2', 'superscript'), t(' '), t('under', 'underline'), t(' <kbd>K</kbd> a'), br, t('b'),
    )));
  });

  it('strikethrough, softbreaks, nested emphasis', () => {
    expect(imp('~~gone~~ ***both*** a\nb')).toEqual(doc(p(t('gone', 'strike'), t(' '), t('both', 'bold', 'italic'), t(' a b'))));
  });

  it('list items whose first block is not a paragraph', () => {
    expect(imp('- # heading\n-\n- ```\n  c\n  ```')).toEqual(doc(ul(
      li(p(t('heading'))),
      li(p()),
      li(p(t('c', { type: 'textStyle', attrs: { fontFamily: 'Menlo' } }))),
    )));
  });

  it('nested blockquotes increase indent', () => {
    expect(impFlat('> a\n>\n> > b')).toEqual(doc(
      { type: 'paragraph', attrs: { indent: 1 }, content: [t('a')] },
      { type: 'paragraph', attrs: { indent: 2 }, content: [t('b')] },
    ));
  });

  describe('math vs currency', () => {
    const cases: [string, PMNode[]][] = [
      ['I have $8500 and $20', [t('I have $8500 and $20')]],
      ['$5 vs $6', [t('$5 vs $6')]],
      ['between $5 and $10.', [t('between $5 and $10.')]],
      ['$x$', [mi('x')]],
      ['$x^2 + y$ done', [mi('x^2 + y'), t(' done')]],
      ['$ x$ and $x $', [t('$ x$ and $x $')]],
      ['$x$5', [t('$x$5')]],
      ['$x$ 5', [mi('x'), t(' 5')]],
      ['$$', [t('$$')]],
      ['$$ $$', [t('$$ $$')]],
      ['\\$x$', [t('$x$')]],
      ['$a\\$b$', [mi('a\\$b')]],
      ['$\\\\$', [mi('\\\\')]],
      ['text $$\\sum$$ inline', [t('text '), mi('\\sum'), t(' inline')]],
      ['price: $3.50, tax $0.20', [t('price: $3.50, tax $0.20')]],
      ['a $b$c', [t('a '), mi('b'), t('c')]],
      ['`$x$`', [t('$x$', { type: 'textStyle', attrs: { fontFamily: 'Menlo' } })]],
    ];
    for (const [src, content] of cases) {
      it(JSON.stringify(src), () => expect(imp(src)).toEqual(doc(p(...content))));
    }

    it('block math forms', () => {
      expect(imp('$$x$$')).toEqual(doc(mb('x')));
      expect(imp('$$\na\nb\n$$')).toEqual(doc(mb('a\nb')));
      expect(imp('$$ a\nb $$')).toEqual(doc(mb('a\nb')));
      expect(impFlat('para\n$$\nx\n$$\nafter')).toEqual(doc(p(t('para')), mb('x'), p(t('after'))));
      expect(imp('$$a$$ and $$b$$')).toEqual(doc(p(mi('a'), t(' and '), mi('b'))));
      expect(imp('$$\nunclosed')).toEqual(doc(p(t('$$ unclosed'))));
    });
  });
});

// ---------------------------------------------------------------------------
// Fuzz
// ---------------------------------------------------------------------------

function mulberry32(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let x = a;
    x = Math.imul(x ^ (x >>> 15), x | 1);
    x ^= x + Math.imul(x ^ (x >>> 7), x | 61);
    return ((x ^ (x >>> 14)) >>> 0) / 4294967296;
  };
}

const ALPHABET = 'abcxyz019  \t*_#.$\\<>[]()`~-+=|&;!:"\'{}^%@/';
const HREFS = ['http://a.b', 'https://x.y/p?q=1&r=2', 'http://p/(x)', 'with space', 'a\\b', 'mailto:x@y.z'];
const LATEX = ['x', 'x^2', '\\frac{a}{b}', 'a + b', '\\alpha_1', 'n!', '{x}', 'a \\\\ b'];

class Gen {
  constructor(private readonly rand: () => number) {}

  int(lo: number, hi: number): number {
    return lo + Math.floor(this.rand() * (hi - lo + 1));
  }

  pick<T>(xs: readonly T[]): T {
    return xs[this.int(0, xs.length - 1)];
  }

  chance(p: number): boolean {
    return this.rand() < p;
  }

  text(): string {
    let s = '';
    const n = this.int(1, 7);
    for (let i = 0; i < n; i++) s += this.pick([...ALPHABET]);
    return s;
  }

  marks(): PMMark[] {
    const marks: PMMark[] = [];
    for (const type of ['bold', 'italic', 'strike', 'underline']) if (this.chance(0.25)) marks.push({ type });
    if (this.chance(0.15)) marks.push({ type: this.pick(['superscript', 'subscript']) });
    if (this.chance(0.15)) marks.push(link(this.pick(HREFS)));
    return marks;
  }

  inline(): PMNode[] {
    if (this.chance(0.1)) return [];
    const out: PMNode[] = [];
    const n = this.int(1, 6);
    for (let i = 0; i < n; i++) {
      const r = this.rand();
      if (r < 0.1) out.push(br);
      else if (r < 0.2) out.push(mi(this.pick(LATEX)));
      else {
        const marks = this.marks();
        out.push(marks.length ? { type: 'text', text: this.text(), marks } : { type: 'text', text: this.text() });
      }
    }
    return out;
  }

  paragraph(): PMNode {
    const content = this.inline();
    return content.length ? { type: 'paragraph', content } : { type: 'paragraph' };
  }

  list(depth: number): PMNode {
    const items: PMNode[] = [];
    const n = this.int(1, 3);
    for (let i = 0; i < n; i++) {
      const children = [this.paragraph()];
      const extra = this.int(0, 2);
      for (let j = 0; j < extra; j++) {
        children.push(depth < 2 && this.chance(0.5) ? this.list(depth + 1) : this.paragraph());
      }
      items.push(li(...children));
    }
    if (this.chance(0.5)) return ul(...items);
    return this.chance(0.5) ? ol(1, ...items) : ol(this.int(0, 120), ...items);
  }

  block(): PMNode {
    const r = this.rand();
    if (r < 0.4) return this.paragraph();
    if (r < 0.55) {
      const content = this.inline();
      const node: PMNode = { type: 'heading', attrs: { level: this.int(1, 3) } };
      if (content.length) node.content = content;
      return node;
    }
    if (r < 0.85) return this.list(0);
    if (r < 0.92) return hr;
    return mb(this.pick([...LATEX, 'a\nb', '']));
  }

  doc(): PMNode {
    const blocks: PMNode[] = [];
    const n = this.int(1, 5);
    for (let i = 0; i < n; i++) blocks.push(this.block());
    return doc(...blocks);
  }
}

describe('markdown fuzz round trip', () => {
  it('300 random documents survive export + import', () => {
    const gen = new Gen(mulberry32(1234));
    for (let i = 0; i < 300; i++) {
      const d = gen.doc();
      const md = exportMarkdown(d);
      const expected = norm(normalizeForMarkdown(d));
      expect(importMarkdown(md).doc, `doc #${i}\n${JSON.stringify(d)}\nmarkdown:\n${md}`).toEqual(expected);
      // Normalization may only move marks off whitespace, never change the text itself.
      expect(plainText(expected)).toBe(plainText(d));
      // Export of an already-imported doc is a fixed point.
      expect(exportMarkdown(importMarkdown(md).doc)).toBe(md);
    }
  });

  it('2000 random paragraphs survive export + import', () => {
    const gen = new Gen(mulberry32(99));
    for (let i = 0; i < 2000; i++) {
      const d = doc(gen.paragraph());
      const md = exportMarkdown(d);
      expect(importMarkdown(md).doc, `para #${i}\n${JSON.stringify(d)}\nmarkdown:\n${md}`).toEqual(norm(normalizeForMarkdown(d)));
    }
  });
});
