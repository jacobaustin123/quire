/**
 * Typewriter mode (keep the caret's line at a fixed height while the page moves) and focus mode
 * (dim every paragraph except the current one).
 */
import { Extension, type Editor } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';

/** Where the caret's line rests, as a fraction of the visible height (a little above centre reads best). */
const ANCHOR = 0.42;
const ANIMATION_MS = 110;

const focusKey = new PluginKey('quireFocus');

/** Marks the textblock holding the cursor so CSS can keep it bright while focus mode dims the rest. */
export const FocusParagraph = Extension.create({
  name: 'focusParagraph',

  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: focusKey,
        props: {
          decorations(state) {
            if (!document.body.classList.contains('focus-mode')) return null;
            const { $head } = state.selection;
            for (let d = $head.depth; d > 0; d--) {
              const node = $head.node(d);
              if (node.isTextblock) {
                const pos = $head.before(d);
                return DecorationSet.create(state.doc, [Decoration.node(pos, pos + node.nodeSize, { class: 'is-current' })]);
              }
            }
            return null;
          },
        },
      }),
    ];
  },
});

/** Keeps the caret line anchored while typing; returns a function that re-centres immediately. */
export function attachTypewriter(editor: Editor, scroller: HTMLElement): () => void {
  let pointerDown = false;
  let frame = 0;

  const animateBy = (delta: number) => {
    cancelAnimationFrame(frame);
    const start = scroller.scrollTop;
    const t0 = performance.now();
    const step = (now: number) => {
      const t = Math.min(1, (now - t0) / ANIMATION_MS);
      const eased = 1 - (1 - t) ** 3;
      scroller.scrollTop = start + delta * eased;
      if (t < 1) frame = requestAnimationFrame(step);
    };
    frame = requestAnimationFrame(step);
  };

  const recenter = (instant = false) => {
    if (!document.body.classList.contains('typewriter') || pointerDown) return;
    const caret = editor.view.coordsAtPos(editor.state.selection.head);
    const box = scroller.getBoundingClientRect();
    const delta = (caret.top + caret.bottom) / 2 - (box.top + box.height * ANCHOR);
    if (Math.abs(delta) < 2) return;
    if (instant) scroller.scrollTop += delta;
    else animateBy(delta);
  };

  editor.on('transaction', ({ transaction }) => {
    if (transaction.docChanged || transaction.selectionSet) requestAnimationFrame(() => recenter());
  });
  scroller.addEventListener('mousedown', () => {
    pointerDown = true;
  });
  window.addEventListener('mouseup', () => {
    if (!pointerDown) return;
    pointerDown = false;
    // A click (not a drag-selection) moves the line it landed on to the anchor.
    if (editor.state.selection.empty) recenter();
  });
  // The user scrolling by hand cancels any in-flight re-centre; the next keystroke settles back.
  scroller.addEventListener('wheel', () => cancelAnimationFrame(frame), { passive: true });
  return () => recenter(true);
}
