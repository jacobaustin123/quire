# Quire

A quiet editor for `.docx`, `.md` and `.txt`. Infinite-scroll text, a formatting bar you can hide,
markdown shortcuts everywhere, LaTeX, and crash recovery. Built with Electron + TipTap/ProseMirror.

> *quire* (n.) — a gathering of folded sheets nested together; the unit a book is sewn from.

## Install

Download `Quire-mac-arm64.zip` from the [latest release](https://github.com/jacobaustin123/quire/releases/latest)
(Apple silicon), unzip, and move `Quire.app` to `/Applications`. The build isn't signed with an
Apple Developer ID, so the first time macOS will refuse to open it; either right-click ▸ Open, or run:

```sh
xattr -dr com.apple.quarantine /Applications/Quire.app
```

Or build it yourself (below) with `npm run install-app`.

## Commands

```sh
npm install           # once
npm start             # build the renderer and launch
npm test              # unit tests; QUIRE_TEST_DOCX=path/to/file.docx npm test also round-trips a real document
npm run typecheck
npm run dist          # build dist-app/mac-arm64/Quire.app
npm run install-app   # build, ad-hoc sign, and copy to /Applications
npm run icon          # re-render build/icon.icns from build/icon.svg
```

Open a file from the command line during development: `npx electron . path/to/file.docx`.

## Using it

| | |
|---|---|
| Formatting bar | ⌘T or the **Aa** button |
| Light / dark | ⇧⌘D or the half-moon button (View ▸ Appearance for System) |
| Bold / italic / underline / strike | ⌘B ⌘I ⌘U ⇧⌘X |
| Super / subscript | ⌘. ⌘, |
| Headings / body | ⌘1–3 / ⌘0 (or ⌥⌘1–3 / ⌥⌘0), or type `# `, `## `, `### ` |
| Zoom | ⌘= / ⌘− / ⇧⌘0 (actual size) |
| Lists | type `* `, `- ` or `1. `; ⇧⌘8 / ⇧⌘7 |
| Indent | Tab / ⇧Tab (nests list items; indents selected paragraphs; inserts a tab otherwise), ⌘] ⌘[ |
| Align | ⇧⌘L ⇧⌘E ⇧⌘R ⇧⌘J |
| Font size | ⇧⌘. / ⇧⌘, |
| Markdown marks | `**bold**`, `*italic*`, `~~strike~~`, `---` rule |
| Math | `$x^2$` inline, `$$…$$` on its own line (or `$$ ` + space) for display; ⇧⌘M; click to edit |
| Link | ⌘K; ⌘-click to open |
| Find | ⌘F, then ⏎ / ⇧⏎ for next / previous |
| Typewriter mode / focus on paragraph | ⇧⌘T / ⇧⌘F |
| Go to page / start / end | ⌘G, ⌘↑ (or ⌘Home), ⌘↓ / ⇧⌘G (or ⌘End) |
| Switch between documents | ⌘` / ⇧⌘` |
| Export PDF / print | ⌥⌘P / ⌘P |

Pages are Letter with 1" margins; the editor column is exactly 6.5" wide and docx line spacing
follows the file, so the page indicator, Go to Page, and PDF output agree (and roughly match Word).

Inline `$…$` follows pandoc's rules (no space inside the delimiters, closing `$` not followed by a
digit), so currency like "$5 and $10" stays text. In `.docx`, equations are stored as `$…$` text in a
`QuireMath` character style, so they survive Word round trips; plain `$` in a docx is never math.

## Safety

- **Crash recovery**: unsaved changes are snapshotted (fsync'd) every ~1.5s to
  `~/Library/Application Support/Quire/Recovery/`. After a crash or power loss, Quire reopens the
  document and asks **Restore** / **Discard Changes**; discarded snapshots move to
  `Recovery/Discarded/` and are kept for 30 days.
- **Backups** (`~/Library/Application Support/Quire/Backups/`, File ▸ Show Backups): one folder per
  document, with dated whole-file copies. Only documents you've saved from Quire are tracked. When a
  tracked document is opened, a background check copies it if there's no copy from today and it differs
  from the newest copy; before Quire overwrites a file (first save of a session, or after another app
  changed it) the on-disk version is copied if it isn't backed up yet. Caps: 100 copies per document,
  1 GB total, oldest first — the newest copy of each document is always kept. See `main/backups.js`.
- **External changes**: the open file is checked every 2s and on focus; it reloads silently if you
  have no unsaved changes, otherwise you choose. Saving over a file another app changed always asks.
- **Save checks**: saving an unchanged document never rewrites it. Every .docx save is read back
  before it touches disk and must contain exactly the text, equations and kept-as-is content being
  saved, or nothing is written. A save that would cut the document to under half its words asks first.
- **Protected mode**: .docx files with things Quire can't keep (footnotes, comments, tracked changes,
  Word equations, section breaks, inline images/text boxes, complex fields, content controls) open
  normally but are never overwritten — saving always asks for a new file. Tables and block images are
  kept byte-for-byte.
- Saves are atomic (write to a temp file, then rename).

## What docx features are kept

Kept: paragraphs, headings (Heading 1–3 styles), bold/italic/underline/strike, super/subscript,
fonts and sizes, links, bulleted/numbered lists (nested, with numbering continuation), alignment,
indentation, tabs, line breaks, horizontal rules, page setup, headers/footers, and styles.
Dropped on save: images, tables (flattened to paragraphs), comments, tracked changes, colors and
highlights, footnotes.

## Layout

```
main/        Electron main process (windows, menus, file IO, recovery, backups) + test harness
src/app.ts   renderer controller (load/save, autosave, toolbar, find, page navigation)
src/editor/  TipTap extensions: indent/tab, font family/size, math, search, smart typing
src/formats/ pure converters (docx, markdown, text) ⇄ ProseMirror JSON — see docs/SCHEMA.md
test/        vitest suites for the converters
```

End-to-end checks: `QUIRE_USER_DATA=/tmp/q QUIRE_HARNESS=script.js npx electron . file.docx`
runs `script.js` (`module.exports = async (ctx) => …`) against the live app; see `main/harness.js`.
