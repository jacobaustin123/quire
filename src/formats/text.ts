/** Plain text <-> Quire document model. See docs/SCHEMA.md for the JSON contract. */

import { DEFAULT_META, type LoadedDoc, type PMNode } from './model';

const LIST_INDENT = '    ';

/**
 * Reads plain text: one paragraph per line, with no Markdown or math interpretation.
 *
 * Args:
 *   src: The file contents. Any of `\r\n`, `\n`, `\r` end a line; a final newline does not
 *     start an extra empty paragraph.
 *
 * Returns:
 *   The document and default metadata.
 */
export function importText(src: string): LoadedDoc {
  const lines = src.split(/\r\n|\n|\r/);
  if (lines.length > 1 && lines[lines.length - 1] === '') lines.pop();
  const content: PMNode[] = lines.map((line) =>
    line ? { type: 'paragraph', content: [{ type: 'text', text: line }] } : { type: 'paragraph' },
  );
  return { doc: { type: 'doc', content }, meta: { ...DEFAULT_META } };
}

/** Renders inline content; hard breaks continue on a new line prefixed by `indent`. */
function inlineText(nodes: PMNode[] | undefined, indent: string): string {
  let out = '';
  for (const node of nodes ?? []) {
    if (node.type === 'text') out += node.text ?? '';
    else if (node.type === 'hardBreak') out += `\n${indent}`;
    else if (node.type === 'mathInline') out += `$${String(node.attrs?.latex ?? '')}$`;
  }
  return out;
}

function listLines(list: PMNode, depth: number): string[] {
  const ordered = list.type === 'orderedList';
  let n = Number(list.attrs?.start ?? 1);
  const base = LIST_INDENT.repeat(depth);
  const lines: string[] = [];
  for (const item of list.content ?? []) {
    const marker = ordered ? `${n++}. ` : '- ';
    const cont = base + ' '.repeat(marker.length);
    let first = true;
    for (const child of item.content ?? []) {
      if (child.type === 'bulletList' || child.type === 'orderedList') {
        lines.push(...listLines(child, depth + 1));
      } else {
        lines.push((first ? base + marker : cont) + inlineText(child.content, cont));
      }
      first = false;
    }
    if (first) lines.push(base + marker.trimEnd());
  }
  return lines;
}

function blockLines(node: PMNode): string[] {
  switch (node.type) {
    case 'paragraph':
    case 'heading':
      return [inlineText(node.content, '')];
    case 'bulletList':
    case 'orderedList':
      return listLines(node, 0);
    case 'horizontalRule':
      return ['---'];
    case 'mathBlock':
      return [`$$${String(node.attrs?.latex ?? '')}$$`];
    case 'docxRaw': {
      const text = String(node.attrs?.text ?? '');
      return text ? text.split('\n') : [];
    }
    default:
      return [];
  }
}

/**
 * Writes plain text: one line per paragraph/heading, lists as `- ` / `N. ` with 4 spaces of
 * indent per nesting level, math as `$latex$` / `$$latex$$`. Formatting is dropped.
 *
 * Args:
 *   doc: The document root (`{type: "doc"}`).
 *
 * Returns:
 *   The text, ending in a single newline.
 */
export function exportText(doc: PMNode): string {
  const lines = (doc.content ?? []).flatMap(blockLines);
  return lines.join('\n') + '\n';
}
