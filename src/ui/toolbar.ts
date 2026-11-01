/** The collapsible formatting bar. */
import type { Editor } from '@tiptap/core';
import { createDropdown, type DropdownOption } from './dropdown';

export const FONTS = [
  'Times New Roman',
  'Charter',
  'Iowan Old Style',
  'Georgia',
  'Palatino',
  'Baskerville',
  'Hoefler Text',
  'Helvetica Neue',
  'Avenir Next',
  'Arial',
  'Gill Sans',
  'Menlo',
  'Courier New',
];

export const SIZES = [8, 9, 10, 10.5, 11, 12, 13, 14, 16, 18, 20, 24, 28, 32, 36, 48, 64, 72];

const ICONS: Record<string, string> = {
  'align-left': '<path d="M4 5.5h12M4 9h8M4 12.5h12M4 16h8"/>',
  'align-center': '<path d="M4 5.5h12M6 9h8M4 12.5h12M6 16h8"/>',
  'align-right': '<path d="M4 5.5h12M8 9h8M4 12.5h12M8 16h8"/>',
  'align-justify': '<path d="M4 5.5h12M4 9h12M4 12.5h12M4 16h12"/>',
  'bullet-list': '<circle cx="5" cy="6" r=".9" class="fill"/><circle cx="5" cy="10" r=".9" class="fill"/><circle cx="5" cy="14" r=".9" class="fill"/><path d="M8.5 6H16M8.5 10H16M8.5 14H16"/>',
  'ordered-list': '<path d="M4.2 4.8 5.3 4.3v3.4M4.2 7.7h2.2M4.2 11.4c.2-.6.7-.9 1.2-.9.6 0 1 .4 1 .9 0 .9-2.3 1.6-2.3 2.6h2.4" stroke-width="1.1"/><path d="M8.5 6H16M8.5 10H16M8.5 14H16"/>',
  outdent: '<path d="M4 5h12M9 9h7M9 12h7M4 16h12M6.5 8.5 4 10.5l2.5 2"/>',
  indent: '<path d="M4 5h12M9 9h7M9 12h7M4 16h12M4 8.5l2.5 2-2.5 2"/>',
  minus: '<path d="M6 10h8"/>',
  plus: '<path d="M6 10h8M10 6v8"/>',
};

function icon(name: string): string {
  return `<svg viewBox="0 0 20 20">${ICONS[name]}</svg>`;
}

/** Distinct fonts and sizes in the selection (or at the cursor), with defaults filled in. */
function selectionStyles(editor: Editor, baseFont: string, baseSize: number): { fonts: Set<string>; sizes: Set<string> } {
  const fonts = new Set<string>();
  const sizes = new Set<string>();
  const { from, to, empty } = editor.state.selection;
  const add = (attrs: Record<string, unknown>) => {
    fonts.add((attrs.fontFamily as string) || baseFont);
    sizes.add((attrs.fontSize as string) || String(baseSize));
  };
  if (empty) {
    add(editor.getAttributes('textStyle'));
    return { fonts, sizes };
  }
  editor.state.doc.nodesBetween(from, to, (node) => {
    if (!node.isText) return true;
    add(node.marks.find((m) => m.type.name === 'textStyle')?.attrs ?? {});
    // Two of each is already "mixed"; no need to keep walking a huge selection.
    return fonts.size < 2 || sizes.size < 2;
  });
  if (!fonts.size) add(editor.getAttributes('textStyle'));
  return { fonts, sizes };
}

export interface ToolbarHooks {
  run: (command: string) => void;
  setFont: (family: string) => void;
  setSize: (size: number) => void;
  baseFont: () => string;
  baseSize: () => number;
}

/** Renders the bar into `root` and returns a function that syncs it with the editor's selection. */
export function buildToolbar(root: HTMLElement, editor: Editor, hooks: ToolbarHooks): () => void {
  root.innerHTML = '';
  const mk = <K extends keyof HTMLElementTagNameMap>(tag: K, cls: string, html = ''): HTMLElementTagNameMap[K] => {
    const e = document.createElement(tag);
    e.className = cls;
    e.innerHTML = html;
    return e;
  };
  const sep = () => root.append(mk('span', 'tb-sep'));
  const buttons = new Map<string, HTMLButtonElement>();
  const button = (cmd: string, html: string, title: string) => {
    const b = mk('button', 'ghost', html);
    b.title = title;
    b.setAttribute('aria-label', title);
    // mousedown+preventDefault keeps the editor's selection and focus.
    b.addEventListener('mousedown', (e) => {
      e.preventDefault();
      hooks.run(cmd);
    });
    buttons.set(cmd, b);
    root.append(b);
  };

  const STYLES: DropdownOption[] = [
    { value: 'p', label: 'Body' },
    { value: 'h1', label: 'Heading 1', style: 'font-weight: 700; font-size: 15px' },
    { value: 'h2', label: 'Heading 2', style: 'font-weight: 700; font-size: 14px' },
    { value: 'h3', label: 'Heading 3', style: 'font-weight: 700' },
  ];
  const style = createDropdown({
    className: 'style',
    title: 'Paragraph style',
    options: () => STYLES,
    onSelect: (v) => hooks.run(v === 'p' ? 'paragraph' : v),
  });
  root.append(style.el);
  sep();

  // The document's own font is listed even if it isn't one of ours, so it can be picked back.
  let currentFont = '';
  const font = createDropdown({
    className: 'font',
    title: 'Font',
    options: () => {
      const list = !currentFont || FONTS.includes(currentFont) ? FONTS : [currentFont, ...FONTS];
      return list.map((f) => ({ value: f, label: f, style: `font-family: "${f}"; font-size: 14px` }));
    },
    onSelect: (f) => hooks.setFont(f),
  });
  root.append(font.el);

  const sizeWrap = mk('div', 'tb-size');
  const down = mk('button', 'ghost', icon('minus'));
  const up = mk('button', 'ghost', icon('plus'));
  down.title = 'Smaller (⇧⌘,)';
  up.title = 'Bigger (⇧⌘.)';
  const size = mk('input', '');
  size.title = 'Font size';
  size.inputMode = 'decimal';
  down.addEventListener('mousedown', (e) => { e.preventDefault(); hooks.run('size-down'); });
  up.addEventListener('mousedown', (e) => { e.preventDefault(); hooks.run('size-up'); });
  const applySize = () => {
    const n = parseFloat(size.value);
    if (n >= 1 && n <= 400) hooks.setSize(n);
    editor.commands.focus();
  };
  size.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); applySize(); }
    if (e.key === 'Escape') { e.preventDefault(); editor.commands.focus(); }
  });
  size.addEventListener('change', applySize);
  // Clicking into the size box selects the whole number, so typing replaces it. The mouseup that
  // follows the focusing click would otherwise collapse the selection back to a caret.
  let justFocused = false;
  size.addEventListener('focus', () => {
    size.select();
    justFocused = true;
  });
  size.addEventListener('mouseup', (e) => {
    if (justFocused) e.preventDefault();
    justFocused = false;
  });
  sizeWrap.append(down, size, up);
  root.append(sizeWrap);
  sep();

  button('bold', '<span class="glyph">B</span>', 'Bold (⌘B)');
  button('italic', '<span class="glyph i">I</span>', 'Italic (⌘I)');
  button('underline', '<span class="glyph u">U</span>', 'Underline (⌘U)');
  button('strike', '<span class="glyph s">S</span>', 'Strikethrough (⇧⌘X)');
  sep();
  button('align-left', icon('align-left'), 'Align left (⇧⌘L)');
  button('align-center', icon('align-center'), 'Center (⇧⌘E)');
  button('align-right', icon('align-right'), 'Align right (⇧⌘R)');
  button('align-justify', icon('align-justify'), 'Justify (⇧⌘J)');
  sep();
  button('bullet-list', icon('bullet-list'), 'Bulleted list (⇧⌘8)');
  button('ordered-list', icon('ordered-list'), 'Numbered list (⇧⌘7)');
  button('outdent', icon('outdent'), 'Decrease indent (⇧⇥)');
  button('indent', icon('indent'), 'Increase indent (⇥)');

  return () => {
    const { fonts, sizes } = selectionStyles(editor, hooks.baseFont(), hooks.baseSize());
    const fam = fonts.size === 1 ? [...fonts][0] : '';
    currentFont = fam;
    // A mixed selection shows a placeholder — naming one font would be a lie.
    font.setValue(fam, 'Multiple fonts');
    if (document.activeElement !== size) size.value = sizes.size === 1 ? [...sizes][0] : '';
    style.setValue(editor.isActive('heading', { level: 1 }) ? 'h1'
      : editor.isActive('heading', { level: 2 }) ? 'h2'
      : editor.isActive('heading', { level: 3 }) ? 'h3' : 'p');
    const on = (cmd: string, active: boolean) => buttons.get(cmd)?.classList.toggle('on', active);
    on('bold', editor.isActive('bold'));
    on('italic', editor.isActive('italic'));
    on('underline', editor.isActive('underline'));
    on('strike', editor.isActive('strike'));
    const align = ['left', 'center', 'right', 'justify'].find((a) => editor.isActive({ textAlign: a })) ?? 'left';
    for (const a of ['left', 'center', 'right', 'justify']) on(`align-${a}`, a === align);
    on('bullet-list', editor.isActive('bulletList'));
    on('ordered-list', editor.isActive('orderedList'));
  };
}
