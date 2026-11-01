/** Shared document-model types used by the editor and every format converter. See docs/SCHEMA.md. */

export interface PMMark {
  type: string;
  attrs?: Record<string, unknown>;
}

export interface PMNode {
  type: string;
  attrs?: Record<string, unknown>;
  content?: PMNode[];
  text?: string;
  marks?: PMMark[];
}

export interface DocMeta {
  baseFont: string;
  baseSize: number;
  paragraphSpacing: number;
  /** Line spacing as a multiple of Word's "single" (docx w:spacing line/240 with lineRule auto). */
  lineSpacing: number;
}

export interface LoadedDoc {
  doc: PMNode;
  meta: DocMeta;
}

export const DEFAULT_META: DocMeta = { baseFont: 'Times New Roman', baseSize: 12, paragraphSpacing: 0, lineSpacing: 1 };

export type Format = 'docx' | 'md' | 'txt';

/** Maps a file path to the format used to read and write it. */
export function formatForPath(path: string): Format {
  const ext = path.toLowerCase().split('.').pop() ?? '';
  if (ext === 'docx') return 'docx';
  if (ext === 'md' || ext === 'markdown') return 'md';
  return 'txt';
}
