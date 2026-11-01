/**
 * Finds content in a .docx body that Quire would lose on save (outside kept-as-is raw blocks).
 * Documents with any of these open in protected mode: saving always goes to a new file.
 */
import type { PMNode } from '../model';

const LOSSY: [RegExp, string][] = [
  [/<w:footnoteReference\b/, 'footnotes'],
  [/<w:endnoteReference\b/, 'endnotes'],
  [/<w:commentReference\b|<w:commentRangeStart\b/, 'comments'],
  [/<w:(ins|del|moveFrom|moveTo)\b[^>]*w:author=/, 'tracked changes'],
  [/<m:oMath\b/, 'Word equations'],
  [/<w:pPr>(?:(?!<\/w:pPr>).)*<w:sectPr\b/s, 'section breaks'],
  [/<w:(drawing|pict|object|txbxContent)\b/, 'inline images or text boxes'],
  [/<w:fldSimple\b(?![^>]*w:instr="\s*(PAGE|NUMPAGES|HYPERLINK)\b)/, 'fields'],
  [/<w:instrText[^>]*>\s*(?!PAGE\b|NUMPAGES\b|HYPERLINK\b)[A-Z]/, 'fields'],
  [/<w:sdt\b/, 'content controls'],
];

function rawXmls(doc: PMNode, out: string[] = []): string[] {
  if (doc.type === 'docxRaw') out.push(String(doc.attrs?.xml ?? ''));
  for (const c of doc.content ?? []) rawXmls(c, out);
  return out;
}

/** Human-readable names of features in `docXml` that Quire can't keep (deduplicated). */
export function lossyFeatures(docXml: string, doc: PMNode): string[] {
  let rest = docXml;
  for (const xml of rawXmls(doc)) if (xml) rest = rest.split(xml).join('');
  const found = new Set<string>();
  for (const [re, name] of LOSSY) if (re.test(rest)) found.add(name);
  return [...found];
}
