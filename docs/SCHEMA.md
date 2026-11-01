# Quire document model (contract between editor and format converters)

All converters operate on plain TipTap/ProseMirror JSON — no editor instance needed —
so they can be unit-tested in Node. Converters live in `src/formats/` and are pure
TypeScript (no Electron / DOM globals except what is passed in).

## Top-level

```ts
interface DocMeta {
  baseFont: string;        // e.g. "Times New Roman" — the document's default font
  baseSize: number;        // pt, e.g. 12
  paragraphSpacing: number; // pt of space after each paragraph (docx Normal spacing after), e.g. 0
  lineSpacing: number;      // multiple of Word "single" (w:spacing w:line/240, lineRule auto); 1.0 if absent in docx
}

interface LoadedDoc {
  doc: PMNode;              // { type: "doc", content: Block[] }
  meta: DocMeta;
}
```

Defaults for new documents / md / txt: `{ baseFont: "Times New Roman", baseSize: 12, paragraphSpacing: 0, lineSpacing: 1 }`
(exported as `DEFAULT_META` from `src/formats/model.ts`).

## Block nodes

| type | attrs | content |
|---|---|---|
| `paragraph` | `textAlign: "left"\|"center"\|"right"\|"justify"` (default `"left"`), `indent: number` (0..8, each level = 0.5in = 720 twips), `outline: number\|null` (docx outlineLvl passthrough, default null) | inline* |
| `heading` | `level: 1\|2\|3`, `textAlign`, `indent` | inline* |
| `bulletList` | `level: number` (default 0) | listItem+ |
| `orderedList` | `start: number` (default 1), `level: number` (default 0) | listItem+ |
| `listItem` | — | paragraph, (paragraph \| bulletList \| orderedList)* — first child always a paragraph |
| `horizontalRule` | — | — |
| `mathBlock` | `latex: string` | — (atom) |
| `docxRaw` | `xml: string` (verbatim outer XML of a body-level element Quire can't edit, e.g. `<w:tbl>…</w:tbl>` or a `<w:p>` containing a drawing), `kind: "table"\|"image"\|"other"`, `text: string` (plain text inside it, for md/txt export and search) | — (atom). Round-trips byte-for-byte through docx. |

`level` is an extra nesting offset for a list whose first item sits at docx `ilvl` > 0 with no
parent item (Word shows it indented as a sub-list). A top-level list with `level: 1` exports at
`ilvl` 1 (+ nesting depth). Nested lists inside it keep level 0 (their depth adds on).

Attrs may be omitted when default. Empty paragraph = `{type:"paragraph"}` with no `content`.

## Inline nodes

| type | attrs |
|---|---|
| `text` | `text: string` (may contain `\t` tab characters; never `\n`), `marks?: Mark[]` |
| `hardBreak` | — (soft line break, docx `<w:br/>`, md trailing `\` newline) |
| `mathInline` | `latex: string` (atom; no marks needed) |

## Marks

| type | attrs |
|---|---|
| `bold`, `italic`, `underline`, `strike`, `superscript`, `subscript` | — |
| `link` | `href: string` (other attrs like target may appear; ignore them) |
| `textStyle` | `fontFamily: string\|null`, `fontSize: string\|null` (pt as a string, e.g. `"14"`) |

`textStyle` is only present when the run differs from `meta.baseFont` / `meta.baseSize`.
Never emit a `textStyle` with both attrs null. Colors/highlights are intentionally dropped.

## Math conventions

- In-editor, LaTeX lives in `mathInline` / `mathBlock` nodes.
- DOCX: written as a run with character style `QuireMath` whose text is `$latex$` (inline) or a
  centered paragraph with a single `QuireMath` run `$$latex$$` (block). On import, only runs with
  the `QuireMath` style become math — plain-text `$` in docx is always literal (the user's notebook
  uses `$` for currency a lot).
- Markdown: `$...$` inline with pandoc rules (opening `$` followed by non-space, closing `$`
  preceded by non-space and NOT followed by a digit; `\$` is a literal dollar), `$$...$$` blocks.
- Text: written as `$latex$` / `$$latex$$`; on txt import, never parsed.

## Converter API (src/formats/)

```ts
// docx.ts
importDocx(bytes: Uint8Array, parseXml: (xml: string) => Document): Promise<LoadedDoc & { original: Uint8Array }>
exportDocx(doc: PMNode, meta: DocMeta, original?: Uint8Array): Promise<Uint8Array>
// markdown.ts
importMarkdown(src: string): LoadedDoc
exportMarkdown(doc: PMNode): string
// text.ts
importText(src: string): LoadedDoc
exportText(doc: PMNode): string
```

`parseXml` is injected so the renderer can use the fast native `DOMParser` and tests can use
`@xmldom/xmldom`. Only use portable DOM APIs: `childNodes`, `nodeType`, `nodeName` (qualified,
e.g. `"w:p"`), `getAttribute("w:val")`, `textContent`, `documentElement`.
