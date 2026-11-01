/**
 * Inline ($…$) and display ($$…$$) LaTeX.
 *
 * Equations are atom nodes rendered with KaTeX. Selecting one (click, or arrowing onto it) "reveals"
 * it: the node is swapped for its raw `$…$` source, tagged with the `mathSource` mark, and edited as
 * ordinary text. When the cursor leaves the source (or on Enter/Escape), it's parsed back into a
 * node; if the delimiters were deleted it simply stays text.
 */
import { InputRule, Mark, Node, mergeAttributes, type Editor } from '@tiptap/core';
import { NodeSelection, Plugin, PluginKey, TextSelection, type EditorState, type Transaction } from '@tiptap/pm/state';
import type { Mark as PMMark, Node as PMNode, ResolvedPos, Schema } from '@tiptap/pm/model';
import type KaTeX from 'katex';

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    math: {
      /** Converts the selected text to an equation, or inserts an empty one ready for typing. */
      insertMath: () => ReturnType;
    };
  }
}

/** Smart typing may have curled quotes or made dashes/ellipses before the source was finished. */
function normalizeLatex(s: string): string {
  return s
    .replace(/[‘’]/g, "'")
    .replace(/[“”]/g, '"')
    .replace(/—/g, '--')
    .replace(/…/g, '...');
}

/** Marks (font, size, bold…) of the surrounding text, so equations and what follows keep them. */
function contextMarks($pos: ResolvedPos): readonly PMMark[] {
  return $pos.marks().filter((m) => m.type.name !== 'mathSource');
}

// KaTeX is ~270KB, so it loads on first use (a document with math, or typing `$x$`). Until then
// equations show their source; each element re-renders once KaTeX arrives.
let katex: typeof KaTeX | null = null;
let katexLoading: Promise<void> | null = null;
const awaitingKatex = new Map<HTMLElement, [string, boolean]>();

function loadKatex(): void {
  katexLoading ??= import('katex').then((mod) => {
    katex = mod.default;
    for (const [el, [latex, display]] of awaitingKatex) if (el.isConnected) renderMath(el, latex, display);
    awaitingKatex.clear();
  });
}

/** Renders `latex` into `el`, showing the source in red if KaTeX can't parse it. */
export function renderMath(el: HTMLElement, latex: string, displayMode: boolean): void {
  if (!latex.trim()) {
    el.innerHTML = '<span class="math-empty">equation</span>';
    return;
  }
  if (!katex) {
    el.textContent = displayMode ? `$$${latex}$$` : `$${latex}$`;
    awaitingKatex.set(el, [latex, displayMode]);
    loadKatex();
    return;
  }
  awaitingKatex.delete(el);
  try {
    katex.render(latex, el, { displayMode, throwOnError: true, strict: false, output: 'htmlAndMathml' });
  } catch {
    el.innerHTML = '';
    const err = document.createElement('span');
    err.className = 'math-error';
    err.textContent = displayMode ? `$$${latex}$$` : `$${latex}$`;
    el.append(err);
  }
}

function mathNodeView(displayMode: boolean) {
  return ({ node }: { node: PMNode }) => {
    const dom = document.createElement(displayMode ? 'div' : 'span');
    dom.className = displayMode ? 'math-block' : 'math-inline';
    dom.contentEditable = 'false';
    let current = node.attrs.latex as string;
    renderMath(dom, current, displayMode);
    return {
      dom,
      update: (next: PMNode) => {
        if (next.type !== node.type) return false;
        if (next.attrs.latex !== current) {
          current = next.attrs.latex as string;
          renderMath(dom, current, displayMode);
        }
        return true;
      },
      ignoreMutation: () => true,
    };
  };
}

// ---------- reveal / collapse ----------

interface RevealState {
  /** Position of the textblock holding revealed source, or null. */
  block: number | null;
}

const revealKey = new PluginKey<RevealState>('mathReveal');

interface SourceRun {
  from: number;
  to: number;
  text: string;
  /** The latex at reveal time; collapsing back to it unchanged stays out of undo history. */
  orig: string;
  display: boolean;
  /** Marks the equation had before it was revealed (the source text can't carry them). */
  marks: readonly PMMark[] | null;
  blockStart: number;
  blockEnd: number;
  coversBlock: boolean;
}

/** Finds runs of `mathSource`-marked text in the textblock at `blockPos`. */
function sourceRuns(doc: PMNode, blockPos: number): SourceRun[] {
  if (blockPos < 0 || blockPos >= doc.content.size) return [];
  const block = doc.nodeAt(blockPos);
  if (!block?.isTextblock) return [];
  const runs: SourceRun[] = [];
  let cur: SourceRun | null = null;
  block.forEach((child, offset) => {
    const mark = child.isText ? child.marks.find((m) => m.type.name === 'mathSource') : undefined;
    const start = blockPos + 1 + offset;
    if (mark && cur && cur.to === start) {
      cur.to += child.nodeSize;
      cur.text += child.text;
    } else if (mark) {
      cur = {
        from: start,
        to: start + child.nodeSize,
        text: child.text ?? '',
        orig: String(mark.attrs.orig ?? ''),
        display: Boolean(mark.attrs.display),
        marks: Array.isArray(mark.attrs.marks) ? (mark.attrs.marks as unknown[]).map((j) => block.type.schema.markFromJSON(j as never)) : null,
        blockStart: blockPos,
        blockEnd: blockPos + block.nodeSize,
        coversBlock: false,
      };
      runs.push(cur);
    } else {
      cur = null;
    }
  });
  for (const r of runs) r.coversBlock = r.from === blockPos + 1 && r.to === blockPos + block.nodeSize - 1;
  return runs;
}

/** Turns a revealed source run back into an equation (or plain text if its delimiters are gone). */
function collapseRun(tr: Transaction, run: SourceRun, schema: Schema): void {
  const from = tr.mapping.map(run.from);
  const to = tr.mapping.map(run.to);
  const text = run.text;
  const display = text.length >= 4 && text.startsWith('$$') && text.endsWith('$$');
  if (display && run.coversBlock) {
    const latex = normalizeLatex(text.slice(2, -2)).trim();
    const start = tr.mapping.map(run.blockStart);
    const end = tr.mapping.map(run.blockEnd);
    const $start = tr.doc.resolve(start);
    const index = $start.index();
    if (!latex) {
      tr.replaceWith(start, end, schema.nodes.paragraph.create());
    } else if ($start.parent.canReplaceWith(index, index + 1, schema.nodes.mathBlock)) {
      tr.replaceWith(start, end, schema.nodes.mathBlock.create({ latex }));
    } else {
      tr.replaceWith(from, to, schema.nodes.mathInline.create({ latex }));
    }
    if (latex === run.orig && run.display) tr.setMeta('addToHistory', false);
    return;
  }
  const m = /^\$([\s\S]*)\$$/.exec(text);
  if (m && text.length >= 2) {
    const inner = display ? m[1].slice(1, -1) : m[1];
    const latex = normalizeLatex(inner).trim().replace(/\n/g, ' ');
    const marks = run.marks ?? contextMarks(tr.doc.resolve(from));
    if (latex) tr.replaceWith(from, to, schema.nodes.mathInline.create({ latex }, null, marks));
    else tr.delete(from, to);
    if (latex === run.orig && !run.display) tr.setMeta('addToHistory', false);
    return;
  }
  tr.removeMark(from, to, schema.marks.mathSource);
  for (const mk of run.marks ?? []) tr.addMark(from, to, mk);
}

/** Replaces the math node at `pos` with its editable source, puts the cursor inside, returns its block. */
function reveal(tr: Transaction, pos: number, node: PMNode, atStart: boolean, schema: Schema): number {
  const latex = node.attrs.latex as string;
  const display = node.type.name === 'mathBlock';
  const delim = display ? '$$' : '$';
  const mark = schema.marks.mathSource.create({ orig: latex, display, marks: node.marks.map((m) => m.toJSON()) });
  const text = schema.text(`${delim}${latex}${delim}`, [mark]);
  let blockPos: number;
  let cursor: number;
  if (display) {
    tr.replaceWith(pos, pos + node.nodeSize, schema.nodes.paragraph.create(null, text));
    blockPos = pos;
    cursor = pos + 1 + delim.length + (atStart ? 0 : latex.length);
  } else {
    tr.replaceWith(pos, pos + node.nodeSize, text);
    blockPos = tr.doc.resolve(pos).before();
    cursor = pos + delim.length + (atStart ? 0 : latex.length);
  }
  tr.setSelection(TextSelection.create(tr.doc, cursor));
  tr.setMeta('addToHistory', false);
  return blockPos;
}

/** The document with any revealed equation source parsed back — what gets saved and compared. */
export function collapsedDoc(state: EditorState): PMNode {
  const block = revealKey.getState(state)?.block;
  if (block == null) return state.doc;
  const runs = sourceRuns(state.doc, block);
  if (!runs.length) return state.doc;
  const tr = state.tr;
  for (const run of runs) collapseRun(tr, run, state.schema);
  return tr.doc;
}

/** Position of the first textblock containing `mathSource` text, if any. */
function findRevealedBlock(doc: PMNode): number | null {
  let found: number | null = null;
  doc.descendants((node, pos) => {
    if (found != null) return false;
    if (!node.isTextblock) return true;
    if (node.content.size && sourceRuns(doc, pos).length) found = pos;
    return false;
  });
  return found;
}

/** Collapses the source under the cursor and moves past it (Enter/Escape). */
function finishEditing(editor: Editor): boolean {
  const { state } = editor;
  const block = revealKey.getState(state)?.block;
  if (block == null) return false;
  const head = state.selection.head;
  const run = sourceRuns(state.doc, block).find((r) => head >= r.from && head <= r.to);
  if (!run) return false;
  const tr = state.tr;
  collapseRun(tr, run, state.schema);
  const after = Math.min(tr.mapping.map(run.to), tr.doc.content.size);
  tr.setSelection(TextSelection.near(tr.doc.resolve(after), 1));
  if (run.marks?.length) tr.setStoredMarks(run.marks);
  tr.setMeta(revealKey, { block: null });
  editor.view.dispatch(tr);
  return true;
}

function revealPlugin(editor: Editor): Plugin<RevealState> {
  return new Plugin<RevealState>({
    key: revealKey,
    state: {
      init: () => ({ block: null }),
      apply(tr, prev) {
        const meta = tr.getMeta(revealKey) as RevealState | undefined;
        if (meta) return meta;
        if (prev.block == null || !tr.docChanged) return prev;
        const mapped = tr.mapping.mapResult(prev.block, 1);
        return { block: mapped.deleted ? null : mapped.pos };
      },
    },
    appendTransaction(trs, oldState, state) {
      if (!trs.some((t) => t.docChanged || t.selectionSet)) return null;
      const schema = state.schema;
      const tr = state.tr;
      const prevActive = revealKey.getState(state)?.block ?? null;
      let active = prevActive;
      // Undo/redo can bring back revealed source anywhere; find it (a full scan, but only on undo).
      if (trs.some((t) => t.getMeta('history$'))) active = findRevealedBlock(state.doc) ?? active;
      // 1) Leaving revealed source parses it back.
      if (active != null) {
        const runs = sourceRuns(state.doc, active);
        const { from, to } = state.selection;
        const inside = runs.find((r) => from >= r.from && to <= r.to);
        for (const run of runs) if (run !== inside) collapseRun(tr, run, schema);
        if (!inside) active = null;
      }
      // 2) Selecting an equation reveals it. Arriving from the left puts the cursor at the start.
      const sel = tr.selection;
      if (sel instanceof NodeSelection && sel.node.type.name.startsWith('math')) {
        const atStart = oldState.selection.head <= sel.from && sel.from - oldState.selection.head <= 1;
        active = reveal(tr, sel.from, sel.node, atStart, schema);
      }
      if (!tr.docChanged && active === prevActive) return null;
      tr.setMeta(revealKey, { block: active });
      return tr;
    },
    props: {
      handleKeyDown(view, event) {
        if (event.key !== 'Enter' && event.key !== 'Escape') return false;
        const block = revealKey.getState(view.state)?.block;
        if (block == null) return false;
        if (event.key === 'Enter' && event.shiftKey) {
          // Multi-line display equations: newline inside the source.
          const head = view.state.selection.head;
          const run = sourceRuns(view.state.doc, block).find((r) => r.coversBlock && head >= r.from && head <= r.to);
          if (!run) return false;
          view.dispatch(view.state.tr.insertText('\n'));
          return true;
        }
        return finishEditing(editor);
      },
    },
    view() {
      const preview = document.createElement('div');
      preview.className = 'math-preview';
      preview.hidden = true;
      document.body.append(preview);
      let last = '';
      return {
        update(view) {
          const block = revealKey.getState(view.state)?.block;
          const run = block == null ? undefined : sourceRuns(view.state.doc, block)[0];
          if (!run) {
            preview.hidden = true;
            last = '';
            return;
          }
          const display = run.text.length >= 4 && run.text.startsWith('$$') && run.text.endsWith('$$');
          const latex = normalizeLatex(display ? run.text.slice(2, -2) : run.text.replace(/^\$|\$$/g, ''));
          const key = `${display}|${latex}`;
          if (key !== last) {
            renderMath(preview, latex, display);
            last = key;
          }
          const start = view.coordsAtPos(run.from);
          const end = view.coordsAtPos(run.to);
          preview.hidden = false;
          const w = preview.offsetWidth;
          const rect = view.dom.getBoundingClientRect();
          const left = display ? (rect.left + rect.right) / 2 - w / 2 : start.left;
          preview.style.left = `${Math.max(8, Math.min(left, window.innerWidth - w - 8))}px`;
          preview.style.top = `${Math.max(start.bottom, end.bottom) + 6}px`;
        },
        destroy() {
          preview.remove();
        },
      };
    },
  });
}

/** Raw LaTeX being edited. `code: true` keeps markdown/smart-typing input rules out of it. */
export const MathSource = Mark.create({
  name: 'mathSource',
  code: true,
  excludes: '_',
  inclusive: false,
  spanning: false,

  addAttributes() {
    return { orig: { default: '' }, display: { default: false }, marks: { default: null } };
  },

  parseHTML() {
    return [];
  },

  renderHTML() {
    return ['span', { class: 'math-source', spellcheck: 'false' }, 0];
  },
});

export const MathInline = Node.create({
  name: 'mathInline',
  group: 'inline',
  inline: true,
  atom: true,
  selectable: true,

  addAttributes() {
    return { latex: { default: '' } };
  },

  parseHTML() {
    return [{ tag: 'span[data-math-inline]', getAttrs: (el) => ({ latex: (el as HTMLElement).getAttribute('data-latex') ?? '' }) }];
  },

  renderHTML({ node, HTMLAttributes }) {
    return ['span', mergeAttributes(HTMLAttributes, { 'data-math-inline': '', 'data-latex': node.attrs.latex }), `$${node.attrs.latex}$`];
  },

  renderText({ node }) {
    return `$${node.attrs.latex}$`;
  },

  addNodeView() {
    return mathNodeView(false);
  },

  addCommands() {
    return {
      insertMath:
        () =>
        ({ state, dispatch }) => {
          if (!dispatch) return true;
          const { from, to, empty } = state.selection;
          const text = empty ? '' : state.doc.textBetween(from, to, ' ');
          const tr = state.tr.replaceSelectionWith(this.type.create({ latex: normalizeLatex(text) }), false);
          // Selecting the new node reveals its source for typing.
          tr.setSelection(NodeSelection.create(tr.doc, from));
          dispatch(tr);
          return true;
        },
    };
  },

  addInputRules() {
    return [
      // `$$x$$` mid-paragraph → inline math.
      new InputRule({
        find: /(?:^|[^\\$])\$\$([^$\n]+)\$\$$/,
        handler: ({ state, range, match }) => {
          const start = range.from + match[0].indexOf('$$');
          const marks = contextMarks(state.doc.resolve(start));
          state.tr.replaceWith(start, range.to, this.type.create({ latex: normalizeLatex(match[1]) }, null, marks));
          state.tr.setStoredMarks(marks);
        },
      }),
      // `$x$` — opening `$` must not follow a word char/digit/`$`/backslash and must hug its content,
      // so currency like "$5 and $10" never triggers.
      new InputRule({
        find: /(?:^|[^\\$\w])\$([^\s$](?:[^$\n]*[^\s$\\])?)\$$/,
        handler: ({ state, range, match }) => {
          // Currency ranges like "$5-$10", "$5/$10", "$1,000–$2,000": the closing `$` is really
          // the start of the next amount (pandoc's "not followed by a digit" can't be seen yet).
          if (/^\d[\d.,]*\s*[-–—/,~]?\s*$/.test(match[1]) || /^\d.*[-–—/,~:;(]\s*$/.test(match[1])) return null;
          const start = range.from + match[0].lastIndexOf('$' + match[1] + '$');
          const marks = contextMarks(state.doc.resolve(start));
          state.tr.replaceWith(start, range.to, this.type.create({ latex: normalizeLatex(match[1]) }, null, marks));
          // Keep typing in the same font after the equation.
          state.tr.setStoredMarks(marks);
        },
      }),
    ];
  },

  addProseMirrorPlugins() {
    return [revealPlugin(this.editor)];
  },
});

export const MathBlock = Node.create({
  name: 'mathBlock',
  group: 'block',
  atom: true,
  selectable: true,
  // Higher than MathInline so a lone `$$x$$` paragraph becomes a display equation.
  priority: 101,

  addAttributes() {
    return { latex: { default: '' } };
  },

  parseHTML() {
    return [{ tag: 'div[data-math-block]', getAttrs: (el) => ({ latex: (el as HTMLElement).getAttribute('data-latex') ?? '' }) }];
  },

  renderHTML({ node, HTMLAttributes }) {
    return ['div', mergeAttributes(HTMLAttributes, { 'data-math-block': '', 'data-latex': node.attrs.latex }), `$$${node.attrs.latex}$$`];
  },

  renderText({ node }) {
    return `$$${node.attrs.latex}$$`;
  },

  addNodeView() {
    return mathNodeView(true);
  },

  addInputRules() {
    return [
      // A paragraph consisting only of `$$x$$` becomes a display equation; typing continues below.
      new InputRule({
        find: /^\$\$([^$]+)\$\$$/,
        handler: ({ state, range, match }) => {
          const $from = state.doc.resolve(range.from);
          if ($from.parent.type.name !== 'paragraph' || $from.depth !== 1) return null;
          const start = $from.before();
          const end = $from.after();
          const { tr } = state;
          tr.replaceWith(start, end, [
            this.type.create({ latex: normalizeLatex(match[1]) }),
            state.schema.nodes.paragraph.create(),
          ]);
          tr.setSelection(TextSelection.create(tr.doc, start + 2));
          tr.setStoredMarks(contextMarks(state.doc.resolve(range.from)));
        },
      }),
      // `$$ ` on an empty line starts a display equation with its source open for typing.
      new InputRule({
        find: /^\$\$\s$/,
        handler: ({ state, range }) => {
          const $from = state.doc.resolve(range.from);
          if ($from.parent.type.name !== 'paragraph' || $from.depth !== 1) return null;
          const start = $from.before();
          const { tr } = state;
          tr.replaceWith(start, $from.after(), [this.type.create({ latex: '' }), state.schema.nodes.paragraph.create()]);
          tr.setSelection(NodeSelection.create(tr.doc, start));
        },
      }),
    ];
  },
});
