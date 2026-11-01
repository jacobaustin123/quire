/** Custom TipTap extensions: indentation, font size, tabs, smart typing, and paste cleanup. */
import { Extension, textInputRule } from '@tiptap/core';
import { Plugin, PluginKey, TextSelection, type EditorState, type Transaction } from '@tiptap/pm/state';
import { Fragment, Slice, type Mark as PMMark, type Node as PMNode } from '@tiptap/pm/model';
import type { EditorView } from '@tiptap/pm/view';
import { closeHistory, isHistoryTransaction } from '@tiptap/pm/history';
import { liftListItem } from '@tiptap/pm/schema-list';

const MAX_INDENT = 8;
const INDENT_TYPES = ['paragraph', 'heading'];

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    quireIndent: {
      indent: () => ReturnType;
      outdent: () => ReturnType;
    };
    fontSize: {
      setFontSize: (size: string | null) => ReturnType;
    };
  }
}

/** True when the selection's start sits inside a list item. */
export function inList(state: EditorState): boolean {
  const { $from } = state.selection;
  for (let d = $from.depth; d > 0; d--) {
    if ($from.node(d).type.name === 'listItem') return true;
  }
  return false;
}

/** Shifts the `indent` attr of every paragraph/heading touched by the selection. */
function shiftIndent(tr: Transaction, delta: number): boolean {
  const { from, to } = tr.selection;
  let changed = false;
  tr.doc.nodesBetween(from, to, (node: PMNode, pos: number) => {
    if (!INDENT_TYPES.includes(node.type.name)) return true;
    const cur = (node.attrs.indent as number) ?? 0;
    const next = Math.max(0, Math.min(MAX_INDENT, cur + delta));
    if (next !== cur) {
      tr.setNodeMarkup(pos, undefined, { ...node.attrs, indent: next });
      changed = true;
    }
    return false;
  });
  return changed;
}

/** Paragraph/heading indentation (0.5in steps) plus the docx outline-level passthrough. */
export const Indent = Extension.create({
  name: 'quireIndent',

  addGlobalAttributes() {
    return [
      {
        types: INDENT_TYPES,
        attributes: {
          indent: {
            default: 0,
            parseHTML: (el) => Number(el.getAttribute('data-indent')) || 0,
            renderHTML: (attrs) =>
              attrs.indent ? { 'data-indent': attrs.indent, style: `margin-left: ${attrs.indent * 0.5}in` } : {},
          },
        },
      },
      {
        types: ['paragraph'],
        attributes: {
          outline: {
            default: null,
            parseHTML: (el) => (el.hasAttribute('data-outline') ? Number(el.getAttribute('data-outline')) : null),
            renderHTML: (attrs) => (attrs.outline == null ? {} : { 'data-outline': attrs.outline }),
          },
        },
      },
    ];
  },

  addCommands() {
    return {
      indent:
        () =>
        ({ tr, state, dispatch, commands }) => {
          if (inList(state)) return commands.sinkListItem('listItem');
          const ok = shiftIndent(tr, 1);
          if (ok && dispatch) dispatch(tr);
          return ok;
        },
      outdent:
        () =>
        ({ tr, state, dispatch, commands }) => {
          if (inList(state)) return commands.liftListItem('listItem');
          const ok = shiftIndent(tr, -1);
          if (ok && dispatch) dispatch(tr);
          return ok;
        },
    };
  },

  addKeyboardShortcuts() {
    return {
      Tab: ({ editor }) => {
        const { state } = editor;
        if (inList(state)) {
          // Word behavior: Tab nests the item; at the first item (can't nest) we swallow the key.
          editor.commands.sinkListItem('listItem');
          return true;
        }
        const { selection } = state;
        if (!selection.empty) return editor.commands.indent() || true;
        return editor.commands.insertContent('\t');
      },
      'Shift-Tab': ({ editor }) => {
        if (inList(editor.state)) {
          editor.commands.liftListItem('listItem');
          return true;
        }
        editor.commands.outdent();
        return true;
      },
      'Mod-]': ({ editor }) => editor.commands.indent(),
      'Mod-[': ({ editor }) => editor.commands.outdent(),
      Backspace: ({ editor }) => backspaceAtStart(editor),
      // At the start of a line these delete "backwards" too, so they must not create a bullet either.
      'Mod-Backspace': ({ editor }) => backspaceAtStart(editor),
      'Alt-Backspace': ({ editor }) => backspaceAtStart(editor),
      'Shift-Backspace': ({ editor }) => backspaceAtStart(editor),
    };
  },
});

/** Backspace-like keys at the start of a paragraph: un-indent, or join onto the list above. */
function backspaceAtStart(editor: import('@tiptap/core').Editor): boolean {
        const { selection } = editor.state;
        if (!selection.empty || !(selection instanceof TextSelection)) return false;
        const { $from } = selection;
        if ($from.parentOffset !== 0) return false;
        if (inList(editor.state)) return removeBullet(editor.state, editor.view.dispatch);
        const node = $from.parent;
        if (!INDENT_TYPES.includes(node.type.name)) return false;
        // Backspace at the very start of an indented paragraph un-indents before merging.
        if (node.attrs.indent) return editor.commands.outdent();
        return joinOntoListAbove(editor.view);
}

/**
 * Backspace at the start of a paragraph right after a list: like Google Docs, an empty line is
 * removed and the cursor lands at the end of the last item; a line with text joins that item's
 * line. (ProseMirror's default would turn the paragraph into a new, empty bullet.)
 */
export function joinOntoListAbove(view: Pick<EditorView, "state" | "dispatch">): boolean {
  const { state } = view;
  const { $from } = state.selection;
  const depth = $from.depth;
  const index = $from.index(depth - 1);
  if (index === 0) return false;
  const before = $from.node(depth - 1).child(index - 1);
  if (before.type.name !== 'bulletList' && before.type.name !== 'orderedList') return false;
  const paraStart = $from.before(depth);
  const paraEnd = $from.after(depth);
  // End of the deepest last textblock inside the list above.
  let target = -1;
  before.descendants((n, pos) => {
    if (n.isTextblock) target = paraStart - before.nodeSize + 1 + pos + n.nodeSize - 1;
  });
  if (target < 0) return false;
  const tr = state.tr.delete(paraStart, paraEnd);
  const para = $from.parent;
  if (para.content.size) tr.insert(target, para.content);
  // If this leaves two lists touching, ListMerge joins them.
  tr.setSelection(TextSelection.create(tr.doc, target));
  view.dispatch(tr.scrollIntoView());
  return true;
}

/**
 * Backspace at the start of a bullet removes the bullet: the line becomes a plain paragraph and
 * the cursor stays at its start (a second Backspace then joins it onto the line above). Nested
 * items come all the way out rather than moving up a level. ProseMirror's default would instead
 * merge the line into the previous item as an unbulleted continuation.
 */
export function removeBullet(state: EditorState, dispatch: (tr: Transaction) => void): boolean {
  const { $from } = state.selection;
  // Only from the item's own first line; later paragraphs in an item join as usual.
  if ($from.node($from.depth - 1).type.name !== 'listItem' || $from.index($from.depth - 1) !== 0) return false;
  const itemType = state.schema.nodes.listItem;
  // Lift one level at a time until out of every list, gathering the steps into one transaction
  // so a single Undo brings the bullet back.
  const tr = state.tr;
  let cur = state;
  while (cur.selection.$from.node(cur.selection.$from.depth - 1).type === itemType) {
    let lifted = false;
    liftListItem(itemType)(cur, (sub) => {
      for (const step of sub.steps) tr.step(step);
      cur = cur.apply(sub);
      lifted = true;
    });
    if (!lifted) break;
  }
  if (!tr.docChanged) return false;
  tr.setSelection(TextSelection.create(tr.doc, cur.selection.from));
  dispatch(tr.scrollIntoView());
  return true;
}

const LIST_TYPES = ['bulletList', 'orderedList'];

/**
 * Two lists of the same kind left touching by an edit become one, so numbering continues
 * (1, 2 rather than 1, 1). Only the edited region is checked, and loading a document (a
 * no-update transaction) or Undo/Redo never merges: lists kept apart in a docx stay apart.
 */
export const ListMerge = Extension.create({
  name: 'listMerge',

  addProseMirrorPlugins() {
    return [listMergePlugin()];
  },
});

/** The ListMerge plugin on its own (for tests on a bare EditorState). */
export function listMergePlugin(): Plugin {
  return new Plugin({
        key: new PluginKey('quireListMerge'),
        appendTransaction(trs, _old, state) {
          if (!trs.some((t) => t.docChanged) || trs.some((t) => t.getMeta('preventUpdate') || isHistoryTransaction(t))) return null;
          // The changed span in the final document.
          let from = Infinity;
          let to = -Infinity;
          trs.forEach((t, i) => {
            const after = trs.slice(i + 1);
            t.mapping.maps.forEach((map, j) => {
              map.forEach((_a, _b, start, end) => {
                let s = start;
                let e = end;
                for (const m of t.mapping.maps.slice(j + 1)) [s, e] = [m.map(s, -1), m.map(e, 1)];
                for (const later of after) [s, e] = [later.mapping.map(s, -1), later.mapping.map(e, 1)];
                from = Math.min(from, s);
                to = Math.max(to, e);
              });
            });
          });
          if (from > to) return null;
          return mergeTouchingLists(state, Math.max(0, from - 1), Math.min(state.doc.content.size, to + 1));
        },
  });
}

/** Joins same-kind, same-level lists that meet within [from, to]; null when there are none. */
export function mergeTouchingLists(state: EditorState, from: number, to: number): Transaction | null {
  const joins: number[] = [];
  const visit = (parent: PMNode, start: number) => {
    let pos = start;
    for (let i = 0; i < parent.childCount; i++) {
      const child = parent.child(i);
      const end = pos + child.nodeSize;
      const next = i + 1 < parent.childCount ? parent.child(i + 1) : null;
      if (next && end >= from && end <= to && LIST_TYPES.includes(child.type.name) && next.type === child.type && next.attrs.level === child.attrs.level) {
        joins.push(end);
      }
      // Descend only into blocks overlapping the edit (nested lists live inside list items).
      if (!child.isTextblock && !child.isAtom && end >= from && pos <= to) visit(child, pos + 1);
      pos = end;
    }
  };
  visit(state.doc, 0);
  if (!joins.length) return null;
  const tr = state.tr;
  // Right to left, so earlier positions stay valid.
  for (const at of joins.reverse()) tr.join(at);
  return tr;
}

/** Adds a pt-valued `fontSize` attribute to the textStyle mark. */
export const FontSize = Extension.create({
  name: 'fontSize',

  addGlobalAttributes() {
    return [
      {
        types: ['textStyle'],
        attributes: {
          fontSize: {
            default: null,
            parseHTML: (el) => {
              const v = el.style.fontSize;
              if (!v) return null;
              const n = parseFloat(v);
              if (v.endsWith('pt')) return String(n);
              if (v.endsWith('px')) return String(Math.round(n * 0.75 * 2) / 2);
              return null;
            },
            renderHTML: (attrs) => (attrs.fontSize ? { style: `font-size: ${attrs.fontSize}pt` } : {}),
          },
        },
      },
    ];
  },

  addCommands() {
    return {
      setFontSize:
        (size) =>
        ({ chain }) =>
          size
            ? chain().setMark('textStyle', { fontSize: size }).run()
            : chain().setMark('textStyle', { fontSize: null }).removeEmptyTextStyle().run(),
    };
  },
});

/** Smart quotes and dashes, which Word users expect. Undo (⌘Z) right after reverts each one. */
export const SmartTyping = Extension.create({
  name: 'smartTyping',

  addInputRules() {
    return [
      textInputRule({ find: /--$/, replace: '—' }),
      textInputRule({ find: /(?:^|[\s{[(<'"‘“—])(")$/, replace: '“' }),
      textInputRule({ find: /"$/, replace: '”' }),
      textInputRule({ find: /(?:^|[\s{[(<'"‘“—])(')$/, replace: '‘' }),
      textInputRule({ find: /'$/, replace: '’' }),
      textInputRule({ find: /\.\.\.$/, replace: '…' }),
    ];
  },
});

const pasteKey = new PluginKey('quirePaste');

/** HTML from Word, Pages/TextEdit or Google Docs, whose fonts are deliberate and worth keeping. */
function fromWordProcessor(html: string): boolean {
  return /urn:schemas-microsoft-com:office|mso-|Cocoa HTML Writer|docs-internal-guid/.test(html);
}

interface PasteOptions {
  baseStyle: () => { fontFamily: string; fontSize: string };
}

/**
 * Paste policy, matching Word's defaults: text from word processors keeps its fonts and sizes
 * (minus ones equal to this document's defaults); text from anywhere else (web pages, plain text)
 * takes on the font at the cursor, so it doesn't drop back to the default font or bring in web fonts.
 * Pastes and drags within Quire are untouched.
 */
export const PasteCleanup = Extension.create<PasteOptions>({
  name: 'pasteCleanup',

  addOptions() {
    return { baseStyle: () => ({ fontFamily: '', fontSize: '' }) };
  },

  addProseMirrorPlugins() {
    let source: 'internal' | 'office' | 'other' = 'other';
    const options = this.options;
    return [
      new Plugin({
        key: pasteKey,
        props: {
          // transformPasted runs before handlePaste, so the source is read from the DOM event.
          handleDOMEvents: {
            paste: (_view, event) => {
              const html = (event as ClipboardEvent).clipboardData?.getData('text/html') ?? '';
              source = html.includes('data-pm-slice') ? 'internal' : fromWordProcessor(html) ? 'office' : 'other';
              return false;
            },
          },
          transformPasted: (slice, view) => {
            const kind = view.dragging ? 'internal' : source;
            source = 'other';
            if (kind === 'internal') return slice;
            const ts = view.state.schema.marks.textStyle;
            const base = options.baseStyle();
            const { $from } = view.state.selection;
            const here = (view.state.storedMarks ?? $from.marks()).find((m) => m.type === ts);
            const restyle = (marks: readonly PMMark[]): readonly PMMark[] => {
              const others = marks.filter((m) => m.type !== ts);
              if (kind === 'other') return here ? here.addToSet(others) : others;
              const own = marks.find((m) => m.type === ts);
              if (!own) return others;
              const fontFamily = own.attrs.fontFamily && own.attrs.fontFamily !== base.fontFamily ? own.attrs.fontFamily : null;
              const fontSize = own.attrs.fontSize && own.attrs.fontSize !== base.fontSize ? own.attrs.fontSize : null;
              return fontFamily || fontSize ? ts.create({ ...own.attrs, fontFamily, fontSize }).addToSet(others) : others;
            };
            const walk = (node: PMNode): PMNode => {
              if (node.isText) return node.mark(restyle(node.marks));
              const children: PMNode[] = [];
              node.content.forEach((c) => children.push(walk(c)));
              return node.copy(Fragment.fromArray(children));
            };
            const content: PMNode[] = [];
            slice.content.forEach((c) => content.push(walk(c)));
            return new Slice(Fragment.fromArray(content), slice.openStart, slice.openEnd);
          },
        },
      }),
    ];
  },
});

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    quireFontFamily: {
      setFontFamily: (family: string) => ReturnType;
      unsetFontFamily: () => ReturnType;
    };
  }
}

/**
 * Font family on the textStyle mark. Unlike TipTap's stock extension this always falls back to
 * the document font, so a missing font (e.g. Calibri on a Mac) degrades gracefully rather than
 * to the browser's default Times.
 */
export const FontFamily = Extension.create({
  name: 'quireFontFamily',

  addGlobalAttributes() {
    return [
      {
        types: ['textStyle'],
        attributes: {
          fontFamily: {
            default: null,
            parseHTML: (el) => {
              const first = el.style.fontFamily.split(',')[0]?.trim().replace(/^["']|["']$/g, '');
              return first && !first.startsWith('var(') ? first : null;
            },
            renderHTML: (attrs) =>
              attrs.fontFamily ? { style: `font-family: "${String(attrs.fontFamily).replace(/"/g, '')}", var(--doc-font)` } : {},
          },
        },
      },
    ];
  },

  addCommands() {
    return {
      setFontFamily:
        (family) =>
        ({ chain }) =>
          chain().setMark('textStyle', { fontFamily: family }).run(),
      unsetFontFamily:
        () =>
        ({ chain }) =>
          chain().setMark('textStyle', { fontFamily: null }).removeEmptyTextStyle().run(),
    };
  },
});

/**
 * Undo granularity like Word/TextEdit: ProseMirror groups everything within its time window into
 * one undo step; here a run of Enters, a paste, and each typing run around them are separate steps.
 * Formatting commands are separated by the app calling `closeUndoGroup` before running them.
 */
export const UndoGrouping = Extension.create({
  name: 'undoGrouping',
  priority: 1000,

  addStorage() {
    return { last: 'none' as 'none' | 'enter' | 'text' | 'paste' };
  },

  addKeyboardShortcuts() {
    const beforeBreak = () => {
      // Consecutive Enters stay one step; the first one closes the preceding typing run.
      if (this.storage.last !== 'enter') closeUndoGroup(this.editor);
      this.storage.last = 'enter';
      return false;
    };
    // Formatting shortcuts handled directly by the editor (not via the menu) also start a new step.
    const beforeCommand = () => {
      closeUndoGroup(this.editor);
      return false;
    };
    const commands = ['Mod-b', 'Mod-i', 'Mod-u', 'Mod-Shift-x', 'Mod-.', 'Mod-,', 'Mod-Shift-7', 'Mod-Shift-8',
      'Mod-Alt-0', 'Mod-Alt-1', 'Mod-Alt-2', 'Mod-Alt-3', 'Mod-Shift-l', 'Mod-Shift-e', 'Mod-Shift-r', 'Mod-Shift-j',
      'Tab', 'Shift-Tab', 'Mod-]', 'Mod-['];
    return {
      Enter: beforeBreak,
      'Shift-Enter': beforeBreak,
      'Mod-Enter': beforeBreak,
      ...Object.fromEntries(commands.map((k) => [k, beforeCommand])),
    };
  },

  addProseMirrorPlugins() {
    const storage = this.storage;
    const editor = this.editor;
    return [
      new Plugin({
        key: new PluginKey('undoGrouping'),
        props: {
          handleTextInput: () => {
            if (storage.last !== 'text') closeUndoGroup(editor);
            storage.last = 'text';
            return false;
          },
          handleDOMEvents: {
            paste: () => {
              closeUndoGroup(editor);
              storage.last = 'paste';
              return false;
            },
          },
        },
      }),
    ];
  },
});

/** Ends the current undo group, so the next change becomes its own undo step. */
export function closeUndoGroup(editor: import('@tiptap/core').Editor): void {
  editor.view.dispatch(closeHistory(editor.state.tr));
  const storage = (editor.storage as unknown as Record<string, { last?: string }>).undoGrouping;
  if (storage) storage.last = 'none';
}

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    paragraphSpacing: {
      /** Line spacing (multiple of Word "single") for the selected paragraphs. */
      setLineSpacing: (multiple: number) => ReturnType;
      /** Space after the selected paragraphs, in pt. */
      setSpaceAfter: (pt: number) => ReturnType;
    };
  }
}

/** Word's "single" spacing as a CSS line-height multiple (see app.ts WORD_SINGLE). */
const CSS_SINGLE = 1.15;

/**
 * Per-paragraph spacing (docx w:spacing): space before/after in pt and line spacing as a multiple
 * of single. Null means "inherit the document default".
 */
export const ParagraphSpacing = Extension.create({
  name: 'paragraphSpacing',

  addGlobalAttributes() {
    const num = (name: string) => (el: HTMLElement) => {
      const v = el.getAttribute(name);
      return v === null || v === '' ? null : Number(v);
    };
    return [
      {
        types: ['paragraph', 'heading'],
        attributes: {
          spaceBefore: {
            default: null,
            parseHTML: num('data-space-before'),
            renderHTML: (a) => (a.spaceBefore == null ? {} : { 'data-space-before': a.spaceBefore, style: `margin-top: ${a.spaceBefore}pt` }),
          },
          spaceAfter: {
            default: null,
            parseHTML: num('data-space-after'),
            renderHTML: (a) => (a.spaceAfter == null ? {} : { 'data-space-after': a.spaceAfter, style: `margin-bottom: ${a.spaceAfter}pt` }),
          },
          lineSpacing: {
            default: null,
            parseHTML: num('data-line-spacing'),
            renderHTML: (a) =>
              a.lineSpacing == null ? {} : { 'data-line-spacing': a.lineSpacing, style: `line-height: ${(CSS_SINGLE * a.lineSpacing).toFixed(3)}` },
          },
        },
      },
    ];
  },

  addCommands() {
    const setAttr = (key: string, value: number) =>
      ({ tr, dispatch }: { tr: Transaction; dispatch?: (tr: Transaction) => void }) => {
        const { from, to } = tr.selection;
        tr.doc.nodesBetween(from, to, (node, pos) => {
          if (node.type.name !== 'paragraph' && node.type.name !== 'heading') return true;
          tr.setNodeMarkup(pos, undefined, { ...node.attrs, [key]: value });
          return false;
        });
        if (dispatch) dispatch(tr);
        return true;
      };
    return {
      setLineSpacing: (multiple) => setAttr('lineSpacing', multiple),
      setSpaceAfter: (pt) => setAttr('spaceAfter', pt),
    };
  },
});
