import { describe, expect, it } from 'vitest';
import { EditorState, TextSelection, type Transaction } from '@tiptap/pm/state';
import { removeBullet } from '../src/editor/extensions';
import { quireSchema } from '../src/editor';

const schema = quireSchema();
const p = (text: string) => ({ type: 'paragraph', content: text ? [{ type: 'text', text }] : [] });
const li = (text: string, ...rest: object[]) => ({ type: 'listItem', content: [p(text), ...rest] });
const ul = (...items: object[]) => ({ type: 'bulletList', content: items });

/** Presses Backspace-at-start in the textblock beginning with `needle` ('' = first empty line). */
function backspaceAt(content: object[], needle: string) {
  const doc = schema.nodeFromJSON({ type: 'doc', content });
  let pos = -1;
  doc.descendants((n, at) => {
    if (pos < 0 && n.isTextblock && (needle ? n.textContent.startsWith(needle) : !n.content.size)) pos = at + 1;
  });
  let state = EditorState.create({ doc, selection: TextSelection.create(doc, pos) });
  let dispatched = 0;
  const handled = removeBullet(state, (tr: Transaction) => {
    dispatched++;
    state = state.apply(tr);
  });
  // Each line, prefixed with one "• " per enclosing list item.
  const lines: string[] = [];
  state.doc.descendants((n, at) => {
    if (!n.isTextblock) return true;
    const $p = state.doc.resolve(at);
    let bullets = 0;
    for (let d = 1; d <= $p.depth; d++) if ($p.node(d).type.name === 'listItem') bullets++;
    lines.push('• '.repeat(bullets) + n.textContent);
    return false;
  });
  const $c = state.selection.$from;
  return { handled, dispatched, lines, cursor: `${$c.parent.textContent.slice(0, $c.parentOffset)}|${$c.parent.textContent.slice($c.parentOffset)}` };
}

describe('backspace at the start of a bullet', () => {
  it('removes the bullet and keeps the cursor at the start of the line', () => {
    const r = backspaceAt([ul(li('one'), li('two'), li('three'))], 'two');
    expect(r.lines).toEqual(['• one', 'two', '• three']);
    expect(r.cursor).toBe('|two');
  });

  it('turns an empty bullet into an empty line', () => {
    const r = backspaceAt([ul(li('one'), li(''))], '');
    expect(r.lines).toEqual(['• one', '']);
    expect(r.cursor).toBe('|');
  });

  it('takes a nested bullet all the way out, in one undoable step', () => {
    const r = backspaceAt([ul(li('one', ul(li('inner'))))], 'inner');
    expect(r.lines).toEqual(['• one', 'inner']);
    expect(r.dispatched).toBe(1);
    expect(r.cursor).toBe('|inner');
  });

  it('leaves lines outside lists alone', () => {
    expect(backspaceAt([p('plain')], 'plain').handled).toBe(false);
  });
});
