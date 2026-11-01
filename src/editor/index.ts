/** Builds the TipTap editor with Quire's schema (see docs/SCHEMA.md). */
import { Editor, getSchema, type JSONContent } from '@tiptap/core';
import StarterKit from '@tiptap/starter-kit';
import Strike from '@tiptap/extension-strike';
import Underline from '@tiptap/extension-underline';
import TextStyle from '@tiptap/extension-text-style';
import TextAlign from '@tiptap/extension-text-align';
import Link from '@tiptap/extension-link';
import Superscript from '@tiptap/extension-superscript';
import Subscript from '@tiptap/extension-subscript';
import { FontFamily, FontSize, Indent, PasteCleanup, SmartTyping } from './extensions';
import { MathBlock, MathInline, MathSource } from './math';
import { DocxRaw, ListLevel } from './raw';
import { Search } from './search';
import { FocusParagraph } from './typewriter';

/** Strikethrough on ⇧⌘X, since TipTap's default ⇧⌘S belongs to Save As. */
const QuireStrike = Strike.extend({
  addKeyboardShortcuts() {
    return { 'Mod-Shift-x': () => this.editor.commands.toggleStrike() };
  },
});

/** Quire's TipTap extensions (the document schema plus editing behaviour). */
export function quireExtensions(baseStyle: EditorHooks['baseStyle'] = () => ({ fontFamily: '', fontSize: '' })) {
  return [
      StarterKit.configure({
        heading: { levels: [1, 2, 3] },
        code: false,
        codeBlock: false,
        blockquote: false,
        strike: false,
        history: { depth: 500, newGroupDelay: 600 },
      }),
      QuireStrike,
      Underline,
      TextStyle,
      FontFamily,
      FontSize,
      TextAlign.configure({ types: ['heading', 'paragraph'], alignments: ['left', 'center', 'right', 'justify'] }),
      Link.configure({ openOnClick: false, autolink: true, linkOnPaste: true, HTMLAttributes: { target: null, rel: null } }),
      Superscript,
      Subscript,
      Indent,
      SmartTyping,
      PasteCleanup.configure({ baseStyle }),
      MathSource,
      MathInline,
      MathBlock,
      DocxRaw,
      ListLevel,
      Search,
      FocusParagraph,
  ];
}

export interface EditorHooks {
  onUpdate: () => void;
  onSelection: () => void;
  onOpenLink: (href: string) => void;
  onDropFiles: (files: File[]) => void;
  onContentError: (error: Error) => void;
  /** The document's default font/size, so pasted text equal to it carries no redundant mark. */
  baseStyle: () => { fontFamily: string; fontSize: string };
}

/** Creates the editor mounted in `element` with initial `content`. */
export function createEditor(element: HTMLElement, content: JSONContent, hooks: EditorHooks): Editor {
  return new Editor({
    element,
    content,
    autofocus: false,
    injectCSS: false,
    enableContentCheck: true,
    extensions: quireExtensions(hooks.baseStyle),
    editorProps: {
      attributes: { spellcheck: 'true', 'aria-label': 'Document' },
      handleClick: (view, pos, event) => {
        if (!event.metaKey) return false;
        const $pos = view.state.doc.resolve(pos);
        const link = $pos.marks().find((m) => m.type.name === 'link');
        if (link) {
          hooks.onOpenLink(link.attrs.href as string);
          return true;
        }
        return false;
      },
      handleDrop: (_view, event) => {
        const files = Array.from(event.dataTransfer?.files ?? []);
        if (!files.length) return false;
        event.preventDefault();
        hooks.onDropFiles(files);
        return true;
      },
    },
    // Without this TipTap silently falls back to an empty doc — which a save would then write out.
    onContentError: ({ error }) => hooks.onContentError(error),
    onUpdate: hooks.onUpdate,
    onSelectionUpdate: hooks.onSelection,
    onTransaction: hooks.onSelection,
  });
}

/** The ProseMirror schema, without an editor (used by tests and tools). */
export function quireSchema() {
  return getSchema(quireExtensions());
}
