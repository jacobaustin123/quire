/** Pictures in the text: insert, paste, drop, and resize by dragging a corner. */
import { Node, mergeAttributes } from '@tiptap/core';
import { closeHistory } from '@tiptap/pm/history';
import { Fragment, type Node as PMNode } from '@tiptap/pm/model';
import { NodeSelection, TextSelection } from '@tiptap/pm/state';
import type { EditorView } from '@tiptap/pm/view';
import { bytesToBase64, COLUMN_PX, dataUriFor, defaultSize, fitSize, imageTypeOf, jpegOrientation, normalizeSrc, pixelSize } from '../formats/images';

const MIN_PX = 16;
/** Longer sides are scaled down to this: over 350 dpi across the column, beyond what print shows. */
const MAX_PIXELS = 2400;

export interface ImageAttrs {
  src: string;
  width: number;
  height: number;
  alt: string;
}

export function isImageFile(f: File): boolean {
  return f.type.startsWith('image/') || /\.(png|jpe?g|gif|webp|bmp|avif|heic|heif|tiff?)$/i.test(f.name);
}

/**
 * Image files to insert for a paste. Word and Excel also put a picture *of the text* on the
 * clipboard; when there is real text, that wins. ("Copy Image" in a browser brings HTML holding
 * only an <img>, which counts as no text.)
 */
export function pastedImageFiles(data: DataTransfer | null): File[] {
  const files = Array.from(data?.files ?? []).filter(isImageFile);
  if (!files.length) return [];
  const html = data?.getData('text/html') ?? '';
  if (!html) return files;
  if (html.includes('data-pm-slice')) return [];
  const text = new DOMParser().parseFromString(html, 'text/html').body.textContent ?? '';
  return text.trim() ? [] : files;
}

/**
 * Turns an image file into attributes for a new picture. PNG, JPEG and GIF are kept byte for byte
 * unless very large or (JPEG) rotated by metadata, which Word may ignore; other formats are converted.
 *
 * Args:
 *   file: The image.
 *   convert: Converts formats the browser can't decode (HEIC, TIFF) to PNG or JPEG bytes.
 */
export async function readImage(file: Blob, convert: (bytes: Uint8Array) => Promise<Uint8Array>): Promise<ImageAttrs> {
  let bytes = new Uint8Array(await file.arrayBuffer());
  let bitmap: ImageBitmap;
  try {
    bitmap = await createImageBitmap(new Blob([bytes], { type: file.type }));
  } catch {
    bytes = new Uint8Array(await convert(bytes));
    bitmap = await createImageBitmap(new Blob([bytes]));
  }
  const name = file instanceof File ? file.name : '';
  // Clipboard images are all called image.png, which says nothing.
  const alt = /^image\.\w+$/i.test(name) ? '' : name.replace(/\.[^.]+$/, '');
  const type = imageTypeOf(bytesToBase64(bytes.subarray(0, 24)));
  const longest = Math.max(bitmap.width, bitmap.height);
  // Re-encoding a GIF would lose its animation, so those are never scaled.
  const keep = type && (type.ext === 'gif' || longest <= MAX_PIXELS) && !(type.ext === 'jpeg' && jpegOrientation(bytes) !== 1);
  const scale = type ? (pixelSize(bytes)?.scale ?? 1) : 1;
  const size = fitSize({ width: bitmap.width, height: bitmap.height, scale });
  if (!keep) {
    const k = Math.min(1, MAX_PIXELS / longest);
    const canvas = new OffscreenCanvas(Math.max(1, Math.round(bitmap.width * k)), Math.max(1, Math.round(bitmap.height * k)));
    const g = canvas.getContext('2d')!;
    g.imageSmoothingQuality = 'high';
    g.drawImage(bitmap, 0, 0, canvas.width, canvas.height);
    const out = await canvas.convertToBlob(type?.ext === 'jpeg' ? { type: 'image/jpeg', quality: 0.92 } : { type: 'image/png' });
    bytes = new Uint8Array(await out.arrayBuffer());
  }
  bitmap.close();
  const src = dataUriFor(bytesToBase64(bytes));
  if (!src) throw new Error('unsupported image format');
  return { src, ...size, alt };
}

/** Inserts pictures at `pos`, or at the cursor (replacing selected text), as one undo step. */
export function insertImages(view: EditorView, images: ImageAttrs[], pos?: number): void {
  if (!images.length) return;
  const { state } = view;
  const type = state.schema.nodes.image;
  const nodes = images.map((a) => type.create(a));
  const sel = state.selection;
  // A selected picture, table or equation is not replaced: the new picture goes after it.
  let from = pos ?? (sel instanceof TextSelection ? sel.from : sel.to);
  let to = pos === undefined && sel instanceof TextSelection ? sel.to : from;
  from = Math.min(from, state.doc.content.size);
  to = Math.min(to, state.doc.content.size);
  const inline = state.doc.resolve(from).parent.inlineContent;
  const content = inline ? Fragment.from(nodes) : Fragment.from(state.schema.nodes.paragraph.create(null, nodes));
  const tr = closeHistory(state.tr).replaceWith(from, to, content);
  // Right after the last picture: the insertion's end, minus a paragraph's closing edge.
  tr.setSelection(TextSelection.near(tr.doc.resolve(tr.mapping.map(to, 1) - (inline ? 0 : 1)), -1));
  view.dispatch(tr.scrollIntoView());
  view.focus();
}

export const QuireImage = Node.create({
  name: 'image',
  group: 'inline',
  inline: true,
  atom: true,
  selectable: true,
  draggable: true,
  marks: '',

  addAttributes() {
    return {
      src: { default: '' },
      width: { default: 300 },
      height: { default: 200 },
      alt: { default: '' },
      // Verbatim source from a .docx. Never put on the clipboard: it only fits its own file.
      xml: { default: null, rendered: false },
    };
  },

  parseHTML() {
    return [
      {
        tag: 'img[src^="data:image/"]',
        getAttrs: (el) => {
          const img = el as HTMLElement;
          const src = normalizeSrc(img.getAttribute('src') ?? '');
          if (!src) return false;
          const width = Number(img.getAttribute('width'));
          const height = Number(img.getAttribute('height'));
          return { src, alt: img.getAttribute('alt') ?? '', xml: null, ...(width > 0 && height > 0 ? { width, height } : defaultSize(src)) };
        },
      },
    ];
  },

  renderHTML({ HTMLAttributes }) {
    return ['img', mergeAttributes(HTMLAttributes)];
  },

  renderText({ node }) {
    return node.attrs.alt as string;
  },

  addNodeView() {
    return ({ node, view, getPos }) => {
      let current = node;
      const dom = document.createElement('span');
      dom.className = 'image-node';
      dom.contentEditable = 'false';
      const img = document.createElement('img');
      img.draggable = false;
      dom.append(img);
      const show = (width: number, height: number) => {
        img.style.width = `${width}px`;
        // Height follows from the ratio, so a narrow window (or list) shrinks it undistorted.
        img.style.aspectRatio = `${width} / ${height}`;
      };
      const sync = (n: PMNode) => {
        if (img.getAttribute('src') !== n.attrs.src) img.src = n.attrs.src as string;
        img.alt = n.attrs.alt as string;
        show(n.attrs.width as number, n.attrs.height as number);
      };
      sync(node);

      for (const corner of ['nw', 'ne', 'sw', 'se']) {
        const handle = document.createElement('span');
        handle.className = `image-handle ${corner}`;
        handle.addEventListener('pointerdown', (e) => startResize(e, corner));
        // Chromium's drag and ProseMirror's selection both start from mousedown.
        handle.addEventListener('mousedown', (e) => e.preventDefault());
        dom.append(handle);
      }
      const tip = document.createElement('span');
      tip.className = 'image-size';
      dom.append(tip);

      const startResize = (e: PointerEvent, corner: string) => {
        if (e.button !== 0 || !view.editable) return;
        e.preventDefault();
        e.stopPropagation();
        const handle = e.currentTarget as HTMLElement;
        handle.setPointerCapture(e.pointerId);
        const ratio = (current.attrs.height as number) / (current.attrs.width as number);
        const startWidth = img.offsetWidth;
        // Screen px per document px (the page is zoomed).
        const zoom = img.getBoundingClientRect().width / Math.max(1, startWidth);
        // A centred picture grows both ways, so its edge moves half as far as its width.
        const centred = getComputedStyle(dom.parentElement ?? dom).textAlign === 'center' ? 2 : 1;
        const dir = corner.endsWith('e') ? 1 : -1;
        const max = Math.max(COLUMN_PX, MIN_PX);
        let width = startWidth;
        const onMove = (m: PointerEvent) => {
          const dx = ((m.clientX - e.clientX) * dir * centred) / zoom;
          const dy = ((m.clientY - e.clientY) * (corner.startsWith('s') ? 1 : -1)) / zoom / ratio;
          // Whichever direction the pointer moved further decides, keeping proportions.
          const d = Math.abs(dx) >= Math.abs(dy) ? dx : dy;
          width = Math.round(Math.max(MIN_PX, Math.min(max, startWidth + d)));
          show(width, width * ratio);
          tip.textContent = `${width} × ${Math.max(1, Math.round(width * ratio))}`;
        };
        const finish = (commit: boolean) => {
          handle.removeEventListener('pointermove', onMove);
          handle.removeEventListener('pointerup', onUp);
          handle.removeEventListener('pointercancel', onCancel);
          window.removeEventListener('keydown', onKey, true);
          dom.classList.remove('resizing');
          const pos = typeof getPos === 'function' ? getPos() : undefined;
          if (!commit || pos === undefined || width === startWidth) {
            sync(current);
            return;
          }
          const height = Math.max(1, Math.round(width * ratio));
          const tr = closeHistory(view.state.tr).setNodeMarkup(pos, undefined, { ...current.attrs, width, height });
          view.dispatch(tr.setSelection(NodeSelection.create(tr.doc, pos)));
        };
        const onUp = () => finish(true);
        const onCancel = () => finish(false);
        const onKey = (k: KeyboardEvent) => {
          if (k.key !== 'Escape') return;
          k.preventDefault();
          k.stopPropagation();
          finish(false);
        };
        dom.classList.add('resizing');
        handle.addEventListener('pointermove', onMove);
        handle.addEventListener('pointerup', onUp);
        handle.addEventListener('pointercancel', onCancel);
        window.addEventListener('keydown', onKey, true);
      };

      return {
        dom,
        update(n: PMNode) {
          if (n.type !== current.type) return false;
          current = n;
          sync(n);
          return true;
        },
        stopEvent: (e: Event) => (e.target as HTMLElement).classList?.contains('image-handle'),
        ignoreMutation: () => true,
      };
    };
  },
});
