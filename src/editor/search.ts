/**
 * In-document find. All matches are counted, but only those near the viewport are highlighted:
 * decorating tens of thousands of hits (e.g. "a" in a 400-page journal) is what makes search slow.
 */
import { Extension } from '@tiptap/core';
import { Plugin, PluginKey, TextSelection, type EditorState } from '@tiptap/pm/state';
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view';
import type { Node as PMNode } from '@tiptap/pm/model';

export interface Match {
  from: number;
  to: number;
}

interface Range {
  from: number;
  to: number;
}

interface SearchState {
  query: string;
  matches: Match[];
  index: number;
  /** Document range currently highlighted (around the viewport). */
  range: Range;
  decorations: DecorationSet;
}

export const searchKey = new PluginKey<SearchState>('quireSearch');

/** Lower-cased searchable text per textblock; nodes are immutable, so this stays valid until edited. */
const textCache = new WeakMap<PMNode, string>();

function blockText(node: PMNode): string {
  let t = textCache.get(node);
  if (t === undefined) {
    // Leaf inline nodes (math, breaks) are size 1, so a one-char placeholder keeps offsets == positions.
    t = node.textBetween(0, node.content.size, undefined, '￼').toLowerCase();
    textCache.set(node, t);
  }
  return t;
}

/** Finds case-insensitive occurrences of `query` in every textblock, in document order. */
export function findMatches(doc: PMNode, query: string): Match[] {
  const out: Match[] = [];
  if (!query) return out;
  const q = query.toLowerCase();
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true;
    const text = blockText(node);
    let i = text.indexOf(q);
    while (i !== -1) {
      out.push({ from: pos + 1 + i, to: pos + 1 + i + q.length });
      i = text.indexOf(q, i + q.length);
    }
    return false;
  });
  return out;
}

/** Index of the first match ending at or after `pos` (binary search). */
function firstFrom(matches: Match[], pos: number): number {
  let lo = 0;
  let hi = matches.length;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (matches[mid].to < pos) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

function decorate(doc: PMNode, matches: Match[], index: number, range: Range): DecorationSet {
  const decos: Decoration[] = [];
  for (let i = firstFrom(matches, range.from); i < matches.length && matches[i].from <= range.to; i++) {
    decos.push(Decoration.inline(matches[i].from, matches[i].to, { class: i === index ? 'search-hit current' : 'search-hit' }));
  }
  const cur = matches[index];
  if (cur && (cur.to < range.from || cur.from > range.to)) {
    decos.push(Decoration.inline(cur.from, cur.to, { class: 'search-hit current' }));
  }
  return DecorationSet.create(doc, decos);
}

const EMPTY: SearchState = { query: '', matches: [], index: -1, range: { from: 0, to: 0 }, decorations: DecorationSet.empty };

/** Returns the current search state. */
export function searchState(state: EditorState): SearchState {
  return searchKey.getState(state) ?? EMPTY;
}

interface SearchMeta {
  query?: string;
  /** Precomputed matches for `query` (saves a second scan). */
  matches?: Match[];
  index?: number;
  range?: Range;
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
            const meta = tr.getMeta(searchKey) as SearchMeta | undefined;
            if (!meta && !(tr.docChanged && prev.query)) return prev;
            const query = meta?.query ?? prev.query;
            if (!query) return EMPTY;
            const matches = meta?.matches ?? (meta?.query !== undefined || tr.docChanged ? findMatches(tr.doc, query) : prev.matches);
            const range = meta?.range ?? {
              from: tr.mapping.map(prev.range.from),
              to: tr.mapping.map(prev.range.to),
            };
            const wanted = meta?.index ?? prev.index;
            const index = matches.length ? Math.max(0, Math.min(wanted, matches.length - 1)) : -1;
            return { query, matches, index, range, decorations: decorate(tr.doc, matches, index, range) };
          },
        },
        props: {
          decorations: (state) => searchKey.getState(state)?.decorations,
        },
      }),
    ];
  },
});

/** Document range covering the visible part of `scroller`, plus a screen of margin either way. */
export function visibleRange(view: EditorView, scroller: HTMLElement): Range {
  const box = scroller.getBoundingClientRect();
  const left = view.dom.getBoundingClientRect().left + 4;
  const at = (y: number, fallback: number) => view.posAtCoords({ left, top: y })?.pos ?? fallback;
  const size = view.state.doc.content.size;
  return {
    from: at(Math.max(box.top + 1, box.top - box.height), 0),
    to: at(Math.min(box.bottom - 1, box.bottom + box.height), size),
  };
}

/** Re-highlights the matches around the viewport (after scrolling). Cheap; no re-search. */
export function setVisibleRange(view: EditorView, range: Range): void {
  const s = searchState(view.state);
  if (!s.query || (range.from === s.range.from && range.to === s.range.to)) return;
  view.dispatch(view.state.tr.setMeta(searchKey, { range }));
}

/** Sets the query (empty clears highlights) and selects the first match at or after the cursor. */
export function setQuery(view: EditorView, query: string, range: Range): SearchState {
  const matches = findMatches(view.state.doc, query);
  const head = view.state.selection.from;
  let idx = matches.findIndex((m) => m.from >= head);
  if (idx < 0) idx = 0;
  view.dispatch(view.state.tr.setMeta(searchKey, { query, matches, index: idx, range }));
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
