import { describe, expect, it } from 'vitest';
import { exportText, importText } from '../src/formats/text';
import { DEFAULT_META, type PMNode } from '../src/formats/model';

const t = (text: string): PMNode => ({ type: 'text', text });
const p = (...content: PMNode[]): PMNode => (content.length ? { type: 'paragraph', content } : { type: 'paragraph' });
const li = (...content: PMNode[]): PMNode => ({ type: 'listItem', content });
const doc = (...content: PMNode[]): PMNode => ({ type: 'doc', content });
const br: PMNode = { type: 'hardBreak' };

describe('text import', () => {
  it('one paragraph per line, default meta', () => {
    const loaded = importText('a\n\nb\tc');
    expect(loaded.meta).toEqual(DEFAULT_META);
    expect(loaded.doc).toEqual(doc(p(t('a')), p(), p(t('b\tc'))));
  });

  it('handles CRLF / CR and a trailing newline', () => {
    expect(importText('a\r\nb\rc\n').doc).toEqual(doc(p(t('a')), p(t('b')), p(t('c'))));
    expect(importText('a\n\n').doc).toEqual(doc(p(t('a')), p()));
  });

  it('empty input gives one empty paragraph', () => {
    expect(importText('').doc).toEqual(doc(p()));
    expect(importText('\n').doc).toEqual(doc(p()));
  });

  it('never interprets markdown or math', () => {
    expect(importText('# h\n- x\n$x$ **b**').doc).toEqual(doc(p(t('# h')), p(t('- x')), p(t('$x$ **b**'))));
  });
});

describe('text export', () => {
  it('round-trips plain inputs', () => {
    const inputs = ['', 'hello', 'a\nb', 'a\n\n\nb', '\tindented\n  spaces  ', '# not\n- md\n$5 and $6', 'x\n'];
    for (const src of inputs) {
      const out = exportText(importText(src).doc);
      expect(out.endsWith('\n')).toBe(true);
      expect(out.replace(/\n$/, '')).toBe(src.replace(/\n$/, ''));
    }
  });

  it('writes headings, lists, breaks, rules and math', () => {
    const d = doc(
      { type: 'heading', attrs: { level: 1 }, content: [t('Title')] },
      p(t('a'), br, t('b'), { type: 'mathInline', attrs: { latex: 'x^2' } }),
      {
        type: 'bulletList',
        content: [
          li(p(t('one'), br, t('cont')), {
            type: 'orderedList',
            attrs: { start: 3 },
            content: [li(p(t('three'))), li(p(t('four')), p(t('more')))],
          }),
          li(p(t('two'))),
        ],
      },
      { type: 'horizontalRule' },
      { type: 'mathBlock', attrs: { latex: '\\int x' } },
    );
    expect(exportText(d)).toBe(
      [
        'Title',
        'a',
        'b$x^2$',
        '- one',
        '  cont',
        '    3. three',
        '    4. four',
        '       more',
        '- two',
        '---',
        '$$\\int x$$',
        '',
      ].join('\n'),
    );
  });
});
