/** Content fingerprints used to verify a save before it touches the user's file. */
import type { PMNode } from './model';

export interface Fingerprint {
  /** All text, with block and line boundaries, equations and kept-as-is content inlined. */
  text: string;
  words: number;
}

function walk(node: PMNode, out: string[]): void {
  switch (node.type) {
    case 'text':
      out.push(node.text ?? '');
      return;
    case 'hardBreak':
      out.push('\n');
      return;
    case 'mathInline':
    case 'mathBlock':
      out.push(`$${String(node.attrs?.latex ?? '')}$`);
      return;
    case 'docxRaw':
      out.push(`[raw:${String(node.attrs?.xml ?? '')}]`);
      return;
  }
  for (const c of node.content ?? []) walk(c, out);
  if (node.type === 'paragraph' || node.type === 'heading' || node.type === 'horizontalRule') out.push(' ');
}

/** Fingerprint of a document's content (formatting is deliberately ignored). */
export function fingerprint(doc: PMNode): Fingerprint {
  const parts: string[] = [];
  walk(doc, parts);
  // Characters XML can't hold are stripped when a .docx is written; they're not content.
  const text = parts.join('').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, '');
  const words = text.replace(/\[raw:[^\]]*\]/g, ' ').split(/\s+/).filter(Boolean).length;
  return { text, words };
}
