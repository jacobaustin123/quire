import { describe, expect, it } from 'vitest';
import { EditorState, TextSelection, type Transaction } from '@tiptap/pm/state';
import { joinOntoListAbove, listMergePlugin, removeBullet } from '../src/editor/extensions';
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

describe('joining a line back onto a list', () => {
  const ol = (...items: object[]) => ({ type: 'orderedList', content: items });

  it('merges the two lists it separated, so numbering continues', () => {
    // 1. one / 2. two / 3. three → Backspace removes the "2." → Backspace again joins it back.
    let state = EditorState.create({ doc: schema.nodeFromJSON({ type: 'doc', content: [ol(li('one'), li('two'), li('three'))] }), plugins: [listMergePlugin()] });
    const dispatch = (tr: Transaction) => { state = state.applyTransaction(tr).state; };
    const at = (text: string) => {
      let pos = -1;
      state.doc.descendants((n, p) => { if (pos < 0 && n.isTextblock && n.textContent === text) pos = p + 1; });
      state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, pos)));
    };
    at('two');
    removeBullet(state, dispatch);
    expect(state.doc.childCount).toBe(3); // list, paragraph, list
    joinOntoListAbove({ state, dispatch } as never);
    expect(state.doc.childCount).toBe(1);
    expect(state.doc.firstChild!.type.name).toBe('orderedList');
    expect(state.doc.firstChild!.childCount).toBe(2);
    expect(state.doc.textContent).toBe('onetwothree');
  });

  it('keeps lists of different kinds apart', () => {
    let state = EditorState.create({ doc: schema.nodeFromJSON({ type: 'doc', content: [ol(li('a')), p('b'), ul(li('c'))] }), plugins: [listMergePlugin()] });
    let pos = -1;
    state.doc.descendants((n, at) => { if (pos < 0 && n.isTextblock && n.textContent === 'b') pos = at + 1; });
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, pos)));
    joinOntoListAbove({ state, dispatch: (tr: Transaction) => { state = state.applyTransaction(tr).state; } } as never);
    expect(state.doc.childCount).toBe(2);
  });
});

describe('touching lists', () => {
  const ol = (...items: object[]) => ({ type: 'orderedList', content: items });
  const make = (content: object[]) => EditorState.create({ doc: schema.nodeFromJSON({ type: 'doc', content }), plugins: [listMergePlugin()] });
  const paraPos = (state: EditorState, text: string) => {
    let pos = -1;
    state.doc.forEach((n, at) => { if (pos < 0 && n.textContent === text && n.isTextblock) pos = at; });
    return pos;
  };

  it('merge when an edit deletes what separated them', () => {
    let state = make([ol(li('hi')), p('x'), ol(li('there'))]);
    const at = paraPos(state, 'x');
    state = state.applyTransaction(state.tr.delete(at, at + 3)).state;
    expect(state.doc.childCount).toBe(1);
    expect(state.doc.firstChild!.childCount).toBe(2);
  });

  it('stay apart when a document is loaded that way', () => {
    let state = make([p('')]);
    const loaded = schema.nodeFromJSON({ type: 'doc', content: [ol(li('hi')), ol(li('there'))] });
    state = state.applyTransaction(state.tr.replaceWith(0, state.doc.content.size, loaded.content).setMeta('preventUpdate', true)).state;
    expect(state.doc.childCount).toBe(2);
  });

  it('stay apart when untouched by the edit', () => {
    let state = make([ol(li('a')), ol(li('b')), p('far away')]);
    const at = paraPos(state, 'far away');
    state = state.applyTransaction(state.tr.insertText('!', at + 1)).state;
    expect(state.doc.childCount).toBe(3);
  });
});
