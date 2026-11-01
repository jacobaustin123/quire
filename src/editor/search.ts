/** In-document find with highlighted matches. */
import { Extension } from '@tiptap/core';
import { Plugin, PluginKey, TextSelection, type EditorState } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import type { Node as PMNode } from '@tiptap/pm/model';
import type { EditorView } from '@tiptap/pm/view';

export interface Match {
  from: number;
  to: number;
}

interface SearchState {
  query: string;
  matches: Match[];
  index: number;
  decorations: DecorationSet;
}

export const searchKey = new PluginKey<SearchState>('quireSearch');

/** Finds case-insensitive occurrences of `query` in every textblock. */
export function findMatches(doc: PMNode, query: string): Match[] {
  const out: Match[] = [];
  if (!query) return out;
  const q = query.toLowerCase();
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true;
    // Leaf inline nodes (math, breaks) are size 1, so a one-char placeholder keeps offsets == positions.
    const text = node.textBetween(0, node.content.size, undefined, '￼').toLowerCase();
    let i = text.indexOf(q);
    while (i !== -1) {
      out.push({ from: pos + 1 + i, to: pos + 1 + i + q.length });
      i = text.indexOf(q, i + q.length);
    }
    return false;
  });
  return out;
}

function build(doc: PMNode, query: string, index: number): SearchState {
  const matches = findMatches(doc, query);
  const idx = matches.length ? Math.max(0, Math.min(index, matches.length - 1)) : -1;
  const decos = matches.map((m, i) => Decoration.inline(m.from, m.to, { class: i === idx ? 'search-hit current' : 'search-hit' }));
  return { query, matches, index: idx, decorations: DecorationSet.create(doc, decos) };
}

const EMPTY: SearchState = { query: '', matches: [], index: -1, decorations: DecorationSet.empty };

/** Returns the current search state. */
export function searchState(state: EditorState): SearchState {
  return searchKey.getState(state) ?? EMPTY;
}

export const Search = Extension.create({
  name: 'quireSearch',

  addProseMirrorPlugins() {
    return [
      new Plugin<SearchState>({
        key: searchKey,
        state: {
          init: () => EMPTY,
          apply(tr, prev) {
            const meta = tr.getMeta(searchKey) as { query?: string; index?: number } | undefined;
            if (meta) return build(tr.doc, meta.query ?? prev.query, meta.index ?? 0);
            if (tr.docChanged && prev.query) return build(tr.doc, prev.query, prev.index);
            return prev;
          },
        },
        props: {
          decorations: (state) => searchKey.getState(state)?.decorations,
        },
      }),
    ];
  },
});

/** Sets the query (empty clears highlights) and selects the first match at or after the cursor. */
export function setQuery(view: EditorView, query: string): SearchState {
  const matches = findMatches(view.state.doc, query);
  const head = view.state.selection.from;
  let idx = matches.findIndex((m) => m.from >= head);
  if (idx < 0) idx = 0;
  view.dispatch(view.state.tr.setMeta(searchKey, { query, index: idx }));
  return searchState(view.state);
}

/** Moves to the next/previous match, selecting it. Returns the new state. */
export function step(view: EditorView, dir: 1 | -1): SearchState {
  const s = searchState(view.state);
  if (!s.matches.length) return s;
  const idx = (s.index + dir + s.matches.length) % s.matches.length;
  const m = s.matches[idx];
  const tr = view.state.tr.setMeta(searchKey, { index: idx }).setSelection(TextSelection.create(view.state.doc, m.from, m.to));
  view.dispatch(tr);
  return searchState(view.state);
}

/** Selects the current match without moving the index. */
export function selectCurrent(view: EditorView): void {
  const s = searchState(view.state);
  if (s.index < 0) return;
  const m = s.matches[s.index];
  view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, m.from, m.to)));
}
