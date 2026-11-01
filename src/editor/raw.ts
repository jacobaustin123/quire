/**
 * Content Quire can't edit (tables, images…) carried through from a .docx verbatim, so saving never
 * changes or drops it. Shown as a quiet, selectable block that can be moved or deleted as a whole.
 */
import { Extension, Node, mergeAttributes } from '@tiptap/core';

const LABELS: Record<string, string> = { table: 'Table', image: 'Image', other: 'Embedded content' };

export const DocxRaw = Node.create({
  name: 'docxRaw',
  group: 'block',
  atom: true,
  selectable: true,
  draggable: true,

  addAttributes() {
    return { xml: { default: '' }, kind: { default: 'other' }, text: { default: '' } };
  },

  parseHTML() {
    return [
      {
        tag: 'div[data-docx-raw]',
        getAttrs: (el) => {
          const e = el as HTMLElement;
          return { xml: e.getAttribute('data-xml') ?? '', kind: e.getAttribute('data-kind') ?? 'other', text: e.textContent ?? '' };
        },
      },
    ];
  },

  renderHTML({ node, HTMLAttributes }) {
    return ['div', mergeAttributes(HTMLAttributes, { 'data-docx-raw': '', 'data-xml': node.attrs.xml, 'data-kind': node.attrs.kind }), node.attrs.text || ''];
  },

  renderText({ node }) {
    return node.attrs.text as string;
  },

  addNodeView() {
    return ({ node }) => {
      const dom = document.createElement('div');
      dom.className = 'docx-raw';
      dom.contentEditable = 'false';
      const label = document.createElement('span');
      label.className = 'docx-raw-label';
      label.textContent = `${LABELS[node.attrs.kind as string] ?? LABELS.other} · kept as-is from the original file`;
      dom.append(label);
      const text = String(node.attrs.text ?? '').trim();
      if (text) {
        const body = document.createElement('div');
        body.className = 'docx-raw-text';
        body.textContent = text;
        dom.append(body);
      }
      return { dom, ignoreMutation: () => true };
    };
  },
});

/** Extra starting depth for lists that begin at a sub-level in Word (see docs/SCHEMA.md). */
export const ListLevel = Extension.create({
  name: 'listLevel',
  addGlobalAttributes() {
    return [
      {
        types: ['bulletList', 'orderedList'],
        attributes: {
          level: {
            default: 0,
            parseHTML: (el) => Number(el.getAttribute('data-level')) || 0,
            renderHTML: (attrs) => (attrs.level ? { 'data-level': attrs.level, style: `margin-left: ${attrs.level * 0.375}in` } : {}),
          },
        },
      },
    ];
  },
});
