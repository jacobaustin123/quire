import { describe, expect, it } from 'vitest';
import { EditorState, TextSelection, type Transaction } from '@tiptap/pm/state';
import { joinListItemBackward } from '../src/editor/extensions';
import { quireSchema } from '../src/editor';

const schema = quireSchema();
const p = (text: string) => ({ type: 'paragraph', content: text ? [{ type: 'text', text }] : [] });
const li = (text: string) => ({ type: 'listItem', content: [p(text)] });
const ul = (...items: object[]) => ({ type: 'bulletList', content: items });

/** Runs Backspace-at-start on the textblock containing `needle` (cursor before it). */
function backspaceAt(content: object[], needle: string) {
  const doc = schema.nodeFromJSON({ type: 'doc', content });
  let pos = -1;
  // An empty needle means the first empty line.
  doc.descendants((n, at) => {
    if (pos < 0 && n.isTextblock && (needle ? n.textContent.startsWith(needle) : !n.content.size)) pos = at + 1;
  });
  let state = EditorState.create({ doc, selection: TextSelection.create(doc, pos) });
  const handled = joinListItemBackward({ state, dispatch: (tr: Transaction) => { state = state.apply(tr); } } as never);
  const lines: string[] = [];
  state.doc.descendants((n, _at, parent) => {
    if (n.isTextblock) {
      lines.push((parent?.type.name === 'listItem' ? '• ' : '') + n.textContent);
      return false;
    }
    return true;
  });
  const $c = state.selection.$from;
  return { handled, lines, cursor: `${$c.parent.textContent.slice(0, $c.parentOffset)}|` };
}

describe('backspace at the start of a bullet', () => {
  it('joins the item onto the previous item', () => {
    const r = backspaceAt([ul(li('one'), li('two'), li('three'))], 'two');
    expect(r.lines).toEqual(['• onetwo', '• three']);
    expect(r.cursor).toBe('one|');
  });

  it('removes an empty bullet and lands at the end of the line above', () => {
    const r = backspaceAt([ul(li('one'), li(''))], '');
    expect(r.lines).toEqual(['• one']);
    expect(r.cursor).toBe('one|');
  });

  it('joins the first item onto the paragraph before the list', () => {
    const r = backspaceAt([p('intro'), ul(li('one'), li('two'))], 'one');
    expect(r.lines).toEqual(['introone', '• two']);
    expect(r.cursor).toBe('intro|');
  });

  it('removes a list whose only item is joined away', () => {
    const r = backspaceAt([p('intro'), ul(li('only')), p('after')], 'only');
    expect(r.lines).toEqual(['introonly', 'after']);
  });

  it('leaves the first line of the document alone', () => {
    expect(backspaceAt([ul(li('one'))], 'one').handled).toBe(false);
  });
});
