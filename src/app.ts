/** Renderer entry: loads the document, wires the editor to menus, saving, recovery and chrome. */
import 'katex/dist/katex.min.css';
import './styles.css';
import type { Editor, JSONContent } from '@tiptap/core';
import type { Node as PMNode } from '@tiptap/pm/model';
import { TextSelection } from '@tiptap/pm/state';
import { createEditor } from './editor';
import { collapsedDoc } from './editor/math';
import { closeUndoGroup } from './editor/extensions';
import { attachTypewriter } from './editor/typewriter';
import { searchState, selectCurrent, setQuery, setVisibleRange, step, visibleRange } from './editor/search';
import { fingerprint } from './formats/fingerprint';
import { DEFAULT_META, formatForPath, type DocMeta, type Format, type LoadedDoc, type PMNode as JSONNode } from './formats/model';
import { closePopover, popoverOpen, showPopover } from './ui/popover';
import { buildToolbar, SIZES } from './ui/toolbar';

interface InitPayload {
  path: string | null;
  name: string;
  bytes: Uint8Array | null;
  mtime: number;
  recovery: string | null;
  error?: string;
}

interface QuireApi {
  init(): Promise<InitPayload>;
  chooseSavePath(defaultName: string, format: Format): Promise<string | null>;
  write(filePath: string, data: Uint8Array | string): Promise<{ path: string; name: string; canceled?: boolean }>;
  setDirty(dirty: boolean): void;
  markPristine(): void;
  onSaveRequest(fn: (id: number) => void): void;
  reload(): Promise<{ bytes: Uint8Array; mtime: number }>;
  onReplaceWord(fn: (info: { x: number; y: number; word: string; replacement: string }) => void): void;
  confirm(message: string, detail: string, ok?: string): Promise<boolean>;
  askRestore(name: string, savedAt: number, fileChangedSince: boolean): Promise<boolean>;
  discardRecovery(): void;
  onFileChanged(fn: (info: { missing: boolean }) => void): void;
  saveResult(id: number, ok: boolean): void;
  writeRecovery(json: string): void;
  clearRecovery(): void;
  exportPdf(defaultName: string): Promise<string | null>;
  print(): void;
  openExternal(url: string): void;
  openPath(p: string): void;
  pathForFile(f: File): string;
  toggleTheme(): void;
  onCommand(fn: (name: string) => void): void;
  onMenu(fn: (name: string) => void): void;
}

declare global {
  interface Window {
    quire: QuireApi;
  }
}

const api = window.quire;
console.info(`[quire] script evaluated at ${performance.now().toFixed(0)}ms`);
/** A Letter page's printable height (9in at 96 CSS px/in); the column is 6.5in wide to match. */
const PAGE_PX = 9 * 96;
const RECOVERY_DELAY_MS = 1500;
const WORD_SINGLE = 1.15;
const $ = <T extends HTMLElement = HTMLElement>(id: string) => document.getElementById(id) as T;

const doc = {
  path: null as string | null,
  name: 'Untitled',
  format: 'docx' as Format,
  meta: { ...DEFAULT_META } as DocMeta,
  original: undefined as Uint8Array | undefined,
  saved: null as PMNode | null,
  dirty: false,
  saving: false,
  /** Set when the file couldn't be read faithfully; saving over it then requires Save As. */
  loadFailed: false,
  /** Features Quire can't keep (footnotes, comments…): the original is never overwritten. */
  lossy: [] as string[],
  /** Path that must never be overwritten (the original of a protected document). */
  protectedPath: null as string | null,
};
let editor: Editor;
let syncToolbar: () => void = () => {};

// ---------- loading ----------

// Converters are loaded on first use: an empty document never pays for jszip or markdown-it.
async function decode(format: Format, bytes: Uint8Array): Promise<LoadedDoc & { original?: Uint8Array; lossy?: string[] }> {
  if (format === 'docx') {
    const { importDocx } = await import('./formats/docx');
    const parser = new DOMParser();
    return importDocx(bytes, (xml) => parser.parseFromString(xml, 'application/xml'));
  }
  const text = new TextDecoder('utf-8').decode(bytes).replace(/^\uFEFF/, '');
  if (format === 'md') return (await import('./formats/markdown')).importMarkdown(text);
  return (await import('./formats/text')).importText(text);
}

function applyMeta(meta: DocMeta): void {
  const root = document.documentElement.style;
  root.setProperty('--doc-font', `"${meta.baseFont}", "Times New Roman", Times, serif`);
  root.setProperty('--doc-size', `${meta.baseSize}pt`);
  root.setProperty('--para-gap', `${meta.paragraphSpacing}pt`);
  // Word's "single" spacing is ~1.15× the font size for Times New Roman (its hhea ascent + descent +
  // line gap), confirmed by measuring a Word screenshot against Quire's at equal glyph height.
  root.setProperty('--line-height', String(WORD_SINGLE * (meta.lineSpacing || 1)));
}

function setName(name: string): void {
  doc.name = name;
  $('doc-name').textContent = name;
  document.title = name;
}

async function boot(): Promise<void> {
  const init = await api.init();
  let loaded: LoadedDoc = { doc: { type: 'doc', content: [{ type: 'paragraph' }] }, meta: { ...DEFAULT_META } };
  doc.path = init.path;
  setName(init.name);
  if (init.path) doc.format = formatForPath(init.path);
  if (init.error) {
    doc.loadFailed = true;
    banner(`Couldn’t open this file: ${init.error}`);
  } else if (init.bytes) {
    try {
      const t0 = performance.now();
      const res = await decode(doc.format, new Uint8Array(init.bytes));
      loaded = res;
      doc.original = res.original;
      if (res.lossy?.length) {
        doc.lossy = res.lossy;
        doc.protectedPath = init.path;
        banner(`This document has ${res.lossy.join(', ')}, which Quire can’t keep. You can read and edit it, but saving will ask for a new file so the original stays intact.`);
      }
      console.info(`[quire] decoded ${doc.format} in ${(performance.now() - t0).toFixed(0)}ms`);
    } catch (err) {
      console.error(err);
      doc.loadFailed = true;
      banner(`Couldn’t read this ${doc.format.toUpperCase()} file (${(err as Error).message}). Saving will ask for a new location so the original isn’t overwritten.`);
    }
  }

  let content: JSONContent = loaded.doc as JSONContent;
  doc.meta = loaded.meta;
  let recovered: { savedAt: number; doc: JSONContent; meta: DocMeta } | null = null;
  if (init.recovery) {
    try {
      recovered = JSON.parse(init.recovery);
    } catch {}
  }
  if (recovered?.doc) {
    // Quire didn't close normally last time (crash, power loss…): ask, don't guess.
    const restore = await api.askRestore(init.name, recovered.savedAt, init.mtime > recovered.savedAt);
    if (restore) {
      content = recovered.doc;
      doc.meta = recovered.meta ?? doc.meta;
    } else {
      api.discardRecovery();
      recovered = null;
    }
  }
  applyMeta(doc.meta);

  const t1 = performance.now();
  editor = createEditor($('editor'), content, {
    onUpdate: onDocChanged,
    onSelection: () => scheduleToolbarSync(),
    onOpenLink: (href) => api.openExternal(href),
    onDropFiles: (files) => files.forEach((f) => api.openPath(api.pathForFile(f))),
    baseStyle: () => ({ fontFamily: doc.meta.baseFont, fontSize: String(doc.meta.baseSize) }),
    onContentError: (err) => {
      doc.loadFailed = true;
      console.error(err);
      banner('Part of this document couldn’t be displayed. Saving will ask for a new location so the original isn’t overwritten.');
    },
  });
  // Handle for the e2e harness and devtools debugging.
  (window as unknown as { __quireEditor: Editor }).__quireEditor = editor;
  console.info(`[quire] editor ready in ${(performance.now() - t1).toFixed(0)}ms`);

  if (recovered) {
    try {
      doc.saved = init.path ? editor.schema.nodeFromJSON(loaded.doc) : editor.schema.nodeFromJSON({ type: 'doc', content: [{ type: 'paragraph' }] });
    } catch {
      doc.saved = null;
    }
    const when = new Date(recovered.savedAt).toLocaleString([], { dateStyle: 'medium', timeStyle: 'short' });
    const stale = init.mtime > recovered.savedAt ? ' The file has changed on disk since then.' : '';
    banner(`Restored unsaved changes from ${when}.${stale}`, init.path ? [['Revert to Saved', () => revertTo(loaded)]] : []);
  } else {
    doc.saved = editor.state.doc;
    if (!init.path) api.markPristine();
  }
  updateDirty();

  syncToolbar = buildToolbar($('toolbar'), editor, {
    run: runCommand,
    setFont: (f) => {
      const ch = editor.chain().focus();
      (f === doc.meta.baseFont ? ch.unsetFontFamily() : ch.setFontFamily(f)).run();
    },
    setSize: setSize,
    baseFont: () => doc.meta.baseFont,
    baseSize: () => doc.meta.baseSize,
    baseLineSpacing: () => doc.meta.lineSpacing || 1,
    baseParagraphSpacing: () => doc.meta.paragraphSpacing,
  });
  syncToolbar();
  setToolbarOpen(localStorage.getItem('quire.toolbar') === '1');
  // Writing modes always start off: they're something you switch into, not a resting state.
  recenterNow = attachTypewriter(editor, scroller());
  editor.commands.focus('start', { scrollIntoView: false });
  updatePageIndicator(false);
  console.info(`[quire] boot done at ${performance.now().toFixed(0)}ms`);
}

function revertTo(loaded: LoadedDoc): void {
  editor.commands.setContent(loaded.doc as JSONContent, false);
  doc.meta = loaded.meta;
  applyMeta(doc.meta);
  doc.saved = editor.state.doc;
  cancelRecoveryWrite();
  updateDirty();
  api.clearRecovery();
  hideBanner();
}

// ---------- spelling ----------

/** Replaces the misspelled `word` under the right-click point with `replacement`, keeping its marks. */
function replaceWord({ x, y, word, replacement }: { x: number; y: number; word: string; replacement: string }): void {
  const hit = editor.view.posAtCoords({ left: x, top: y });
  if (!hit) return;
  const $pos = editor.state.doc.resolve(hit.pos);
  const block = $pos.parent;
  if (!block.isTextblock) return;
  const start = $pos.start();
  const text = block.textBetween(0, block.content.size, undefined, '\uFFFC');
  const offset = hit.pos - start;
  // The occurrence of `word` that contains (or touches) the click point.
  let from = -1;
  for (let i = text.indexOf(word); i !== -1; i = text.indexOf(word, i + 1)) {
    if (i <= offset && offset <= i + word.length) {
      from = i;
      break;
    }
  }
  if (from < 0) return;
  const tr = editor.state.tr.insertText(replacement, start + from, start + from + word.length);
  editor.view.dispatch(tr);
  editor.view.focus();
}

// ---------- external changes ----------

/** Whether the banner currently shows an external-change conflict (cleared by a successful save). */
let conflictShown = false;

/** Replaces the document with what's on disk now, keeping the reading position. */
async function reloadFromDisk(): Promise<void> {
  const { bytes } = await api.reload();
  const loaded = await decode(doc.format, new Uint8Array(bytes));
  const s = scroller();
  const ratio = s.scrollTop / Math.max(1, s.scrollHeight);
  const head = editor.state.selection.head;
  editor.commands.setContent(loaded.doc as JSONContent, false);
  doc.original = loaded.original ?? doc.original;
  doc.meta = loaded.meta;
  applyMeta(doc.meta);
  doc.saved = editor.state.doc;
  editor.commands.setTextSelection(Math.min(head, editor.state.doc.content.size - 1));
  s.scrollTop = ratio * s.scrollHeight;
  cancelRecoveryWrite();
  updateDirty();
  api.clearRecovery();
}

function onFileChanged({ missing }: { missing: boolean }): void {
  if (missing) {
    banner(`“${doc.name}” was moved or deleted. Saving will create it again.`);
    return;
  }
  if (!doc.dirty) {
    // Nothing to lose: follow the file, like TextEdit and Pages do.
    reloadFromDisk()
      .then(() => {
        const text = `Reloaded — “${doc.name}” was changed in another app.`;
        banner(text);
        setTimeout(() => {
          if ($('banner').textContent?.startsWith(text)) hideBanner();
        }, 5000);
      })
      .catch((err) => banner(`“${doc.name}” changed in another app, but couldn’t be reloaded: ${(err as Error).message}`));
    return;
  }
  conflictShown = true;
  banner(`“${doc.name}” was changed in another app, and you have unsaved changes here.`, [
    ['Reload (Discard Mine)', () => void reloadFromDisk().then(hideBanner)],
    ['Keep Mine', hideBanner],
  ]);
}

// ---------- dirty state & crash recovery ----------

let recoveryTimer: number | undefined;
let recoveryIdle: number | undefined;

function cancelRecoveryWrite(): void {
  window.clearTimeout(recoveryTimer);
  recoveryTimer = undefined;
  if (recoveryIdle !== undefined) window.cancelIdleCallback(recoveryIdle);
  recoveryIdle = undefined;
}

function updateDirty(): void {
  const dirty = !doc.saved || !collapsedDoc(editor.state).eq(doc.saved);
  if (dirty !== doc.dirty) {
    doc.dirty = dirty;
    document.body.classList.toggle('dirty', dirty);
    api.setDirty(dirty);
  }
}

function onDocChanged(): void {
  updateDirty();
  scheduleToolbarSync();
  if (!doc.dirty) {
    cancelRecoveryWrite();
    api.clearRecovery();
    return;
  }
  // Not debounced-to-idle: continuous typing still gets snapshotted every RECOVERY_DELAY_MS.
  if (recoveryTimer === undefined) {
    recoveryTimer = window.setTimeout(writeRecovery, RECOVERY_DELAY_MS);
  }
}

function writeRecovery(): void {
  recoveryTimer = undefined;
  if (!doc.dirty) return;
  recoveryIdle = window.requestIdleCallback(
    () => {
      recoveryIdle = undefined;
      // Re-checked here: a save may have finished while this was waiting for idle time.
      if (!doc.dirty || doc.saving) return;
      api.writeRecovery(JSON.stringify({ version: 1, path: doc.path, format: doc.format, name: doc.name, savedAt: Date.now(), meta: doc.meta, doc: collapsedDoc(editor.state).toJSON() }));
    },
    { timeout: 1000 },
  );
}

// ---------- saving ----------

function stem(): string {
  return doc.name.replace(/\.[^.]+$/, '');
}

async function encode(format: Format, json: JSONNode): Promise<Uint8Array | string> {
  if (format === 'docx') return (await import('./formats/docx')).exportDocx(json, doc.meta, doc.original);
  if (format === 'md') return (await import('./formats/markdown')).exportMarkdown(json);
  return (await import('./formats/text')).exportText(json);
}

async function save(saveAs = false): Promise<boolean> {
  if (doc.saving) return false;
  // Nothing changed: never rewrite the file (so open → save can't alter a document).
  if (!saveAs && doc.path && !doc.dirty && !doc.loadFailed && !doc.protectedPath) return true;
  let path = doc.path;
  let format = doc.format;
  if (!path || saveAs || doc.loadFailed || doc.protectedPath === path) {
    path = await api.chooseSavePath(doc.protectedPath ? `${stem()} (Quire)` : stem(), path ? format : 'docx');
    if (!path) return false;
    if (path === doc.protectedPath) {
      banner(`Quire won’t overwrite the original — it has ${doc.lossy.join(', ')} that would be lost. Choose a different name.`);
      return false;
    }
    format = formatForPath(path);
  }
  doc.saving = true;
  cancelRecoveryWrite();
  try {
    // An equation being edited is saved as an equation, not as its raw `$…$` source.
    const snapshot = collapsedDoc(editor.state);
    const json = snapshot.toJSON() as JSONNode;
    const data = await encode(format, json);
    if (!(await verifySave(format, json, data, path))) return false;
    const res = await api.write(path, data);
    if (res.canceled) return false;
    doc.loadFailed = false;
    doc.path = res.path;
    doc.format = format;
    setName(res.name);
    doc.saved = snapshot;
    if (conflictShown) {
      conflictShown = false;
      hideBanner();
    }
    updateDirty();
    // Main marked the window clean on write; if typing continued during the save, say so again.
    api.setDirty(doc.dirty);
    if (doc.dirty) onDocChanged();
    return true;
  } catch (err) {
    console.error(err);
    banner(`Couldn’t save: ${(err as Error).message}`);
    return false;
  } finally {
    doc.saving = false;
  }
}

/**
 * Last line of defence before a save touches disk: a .docx must read back with exactly the content
 * being saved, and a save that would shrink the document drastically needs confirmation.
 */
async function verifySave(format: Format, json: JSONNode, data: Uint8Array | string, path: string): Promise<boolean> {
  const want = fingerprint(json);
  if (format === 'docx') {
    const { importDocx } = await import('./formats/docx');
    const parser = new DOMParser();
    const back = await importDocx(data as Uint8Array, (xml) => parser.parseFromString(xml, 'application/xml'));
    if (fingerprint(back.doc).text !== want.text) {
      banner('Save stopped: the file Quire produced didn’t read back identically, so nothing was written. Your document is unchanged on disk; please use Save As to a new file and report this.');
      return false;
    }
  }
  const before = doc.saved ? fingerprint(doc.saved.toJSON() as JSONNode).words : 0;
  if (path === doc.path && before >= 100 && want.words < before * 0.5) {
    return api.confirm(
      `Replace “${doc.name}” with a much shorter version?`,
      `The file on disk has about ${before.toLocaleString()} words; this save has ${want.words.toLocaleString()}. A backup of the current file is kept either way.`,
    );
  }
  return true;
}

// ---------- chrome ----------

function banner(text: string, actions: [string, () => void][] = []): void {
  const el = $('banner');
  el.replaceChildren(Object.assign(document.createElement('span'), { textContent: text }));
  for (const [label, fn] of actions) {
    const b = document.createElement('button');
    b.textContent = label;
    b.onclick = fn;
    el.append(b);
  }
  const x = document.createElement('button');
  x.textContent = 'Dismiss';
  x.onclick = hideBanner;
  el.append(x);
  el.hidden = false;
}

function hideBanner(): void {
  $('banner').hidden = true;
}

function setToolbarOpen(open: boolean): void {
  document.body.classList.toggle('toolbar-open', open);
  $('btn-toolbar').classList.toggle('on', open);
  try {
    localStorage.setItem('quire.toolbar', open ? '1' : '0');
  } catch {}
}

let toolbarQueued = false;
function scheduleToolbarSync(): void {
  if (toolbarQueued) return;
  toolbarQueued = true;
  requestAnimationFrame(() => {
    toolbarQueued = false;
    if (document.body.classList.contains('toolbar-open')) syncToolbar();
  });
}

/** Switches light/dark. Instant on purpose: macOS redraws the window controls immediately, so any
 * fade on the page would visibly lag behind them. */
function toggleTheme(): void {
  api.toggleTheme();
}

// ---------- typewriter & focus ----------

let recenterNow: () => void = () => {};
const MODE_LABELS: Record<string, string> = { typewriter: 'Typewriter', 'focus-mode': 'Focus' };

/** Turns a writing mode on/off for this window (a body class) and briefly says so. */
function setMode(mode: 'typewriter' | 'focus-mode', on: boolean, announce = true): void {
  document.body.classList.toggle(mode, on);
  // Re-run decorations (focus) and re-anchor the caret (typewriter) right away.
  editor.view.dispatch(editor.state.tr.setMeta('quireMode', mode));
  if (mode === 'typewriter' && on) requestAnimationFrame(recenterNow);
  if (announce) {
    const el = $('page-indicator');
    el.textContent = `${MODE_LABELS[mode]} ${on ? 'on' : 'off'}`;
    indicatorHeldUntil = performance.now() + 1100;
    el.classList.add('show');
    window.clearTimeout(indicatorTimer);
    indicatorTimer = window.setTimeout(() => el.classList.remove('show'), 1100);
  }
}

// ---------- pages & navigation ----------

const scroller = () => $('scroller');

function editorTop(): number {
  return $('editor').getBoundingClientRect().top - scroller().getBoundingClientRect().top + scroller().scrollTop;
}

/** One printed page's height in on-screen pixels at the current document zoom. */
function pagePx(): number {
  return PAGE_PX * docZoom;
}

function pageCount(): number {
  return Math.max(1, Math.ceil($('editor').getBoundingClientRect().height / pagePx()));
}

function currentPage(): number {
  const y = scroller().scrollTop - editorTop() + scroller().clientHeight * 0.25;
  return Math.max(1, Math.min(pageCount(), Math.floor(y / pagePx()) + 1));
}

let indicatorTimer: number | undefined;
/** Page numbers while scrolling are opt-in (View ▸ Show Page Numbers); explicit jumps always show. */
let showPageNumbers = (() => {
  try {
    return localStorage.getItem('quire.pageNumbers') === '1';
  } catch {
    return false;
  }
})();
/** While a short message (e.g. "Typewriter on") is showing, page numbers don't overwrite it. */
let indicatorHeldUntil = 0;
function updatePageIndicator(show = true): void {
  if (performance.now() < indicatorHeldUntil) return;
  const el = $('page-indicator');
  el.textContent = `${currentPage()} / ${pageCount()}`;
  if (!show) return;
  el.classList.add('show');
  window.clearTimeout(indicatorTimer);
  indicatorTimer = window.setTimeout(() => el.classList.remove('show'), 1200);
}

function gotoPage(n: number): void {
  const total = pageCount();
  const page = Math.max(1, Math.min(total, Math.round(n)));
  const s = scroller();
  // Page 1 is the top of the document itself, not the first line of text below the margin.
  s.scrollTop = page === 1 ? 0 : editorTop() + (page - 1) * pagePx() - 24;
  const rect = $('editor').getBoundingClientRect();
  const found = editor.view.posAtCoords({ left: rect.left + 4, top: Math.max(rect.top, s.getBoundingClientRect().top + 30) });
  if (found) {
    const tr = editor.state.tr.setSelection(TextSelection.near(editor.state.doc.resolve(found.pos)));
    editor.view.dispatch(tr);
  }
  editor.view.focus();
  updatePageIndicator();
}

/** Scrolls so `pos` sits about a third of the way down, unless it's already comfortably visible. */
function revealPos(pos: number): void {
  const s = scroller();
  const coords = editor.view.coordsAtPos(pos);
  const box = s.getBoundingClientRect();
  if (coords.top > box.top + 40 && coords.bottom < box.bottom - box.height * 0.15) return;
  s.scrollTop += coords.top - box.top - box.height * 0.33;
}

function gotoStart(): void {
  editor.commands.focus('start', { scrollIntoView: false });
  scroller().scrollTop = 0;
  updatePageIndicator();
}

function gotoEnd(): void {
  editor.commands.focus('end', { scrollIntoView: false });
  revealPos(editor.state.doc.content.size);
  updatePageIndicator();
}

// ---------- zoom ----------

const ZOOMS = [0.8, 0.9, 1, 1.07, 1.15, 1.25, 1.35, 1.5, 1.75, 2];
const DEFAULT_ZOOM = 1.15;
let docZoom = DEFAULT_ZOOM;

/** Zooms the document only (not the window chrome), keeping the cursor's line in place. */
function setZoom(z: number): void {
  const anchor = editor ? editor.view.coordsAtPos(editor.state.selection.head).top : 0;
  docZoom = z;
  document.documentElement.style.setProperty('--doc-zoom', String(z));
  try {
    localStorage.setItem('quire.zoom', String(z));
  } catch {}
  if (editor) {
    const after = editor.view.coordsAtPos(editor.state.selection.head).top;
    scroller().scrollTop += after - anchor;
  }
  $('page-indicator').textContent = `${Math.round(z * 100)}%`;
  $('page-indicator').classList.add('show');
  window.clearTimeout(indicatorTimer);
  indicatorTimer = window.setTimeout(() => $('page-indicator').classList.remove('show'), 1000);
}

function stepZoom(dir: 1 | -1): void {
  const next = dir > 0 ? ZOOMS.find((z) => z > docZoom + 1e-6) : [...ZOOMS].reverse().find((z) => z < docZoom - 1e-6);
  if (next) setZoom(next);
}

try {
  const saved = parseFloat(localStorage.getItem('quire.zoom') ?? '');
  if (saved >= 0.5 && saved <= 3) docZoom = saved;
} catch {}
document.documentElement.style.setProperty('--doc-zoom', String(docZoom));

// ---------- find ----------

function openFind(): void {
  const bar = $('findbar');
  const input = $<HTMLInputElement>('find-input');
  const { from, to, empty } = editor.state.selection;
  const selected = empty ? '' : editor.state.doc.textBetween(from, to, ' ');
  if (selected && selected.length < 120 && !selected.includes('\n')) input.value = selected;
  bar.hidden = false;
  input.focus();
  input.select();
  runFind();
}

let findTimer: number | undefined;

/** Searches after a short pause in typing, so each keystroke doesn't rescan a 400-page document. */
function scheduleFind(): void {
  window.clearTimeout(findTimer);
  findTimer = window.setTimeout(runFind, 120);
}

function runFind(): void {
  window.clearTimeout(findTimer);
  findTimer = undefined;
  const q = $<HTMLInputElement>('find-input').value;
  const s = setQuery(editor.view, q, visibleRange(editor.view, scroller()));
  renderFindCount();
  if (s.index >= 0) revealPos(s.matches[s.index].from);
}

function renderFindCount(): void {
  const s = searchState(editor.state);
  const q = $<HTMLInputElement>('find-input').value;
  $('find-count').textContent = !q ? '' : s.matches.length ? `${s.index + 1} / ${s.matches.length}` : 'none';
}

function findStep(dir: 1 | -1): void {
  if ($('findbar').hidden) {
    openFind();
    return;
  }
  // Enter right after typing: run the pending search (landing on the first match) instead of skipping it.
  if (findTimer !== undefined) {
    runFind();
    return;
  }
  const s = step(editor.view, dir);
  renderFindCount();
  if (s.index >= 0) revealPos(s.matches[s.index].from);
}

function closeFind(): void {
  $('findbar').hidden = true;
  selectCurrent(editor.view);
  setQuery(editor.view, '', { from: 0, to: 0 });
  editor.view.focus();
}

// ---------- commands ----------

function currentSize(): number {
  const v = editor.getAttributes('textStyle').fontSize as string | undefined;
  return v ? parseFloat(v) : doc.meta.baseSize;
}

function setSize(n: number): void {
  const ch = editor.chain().focus();
  (n === doc.meta.baseSize ? ch.setFontSize(null) : ch.setFontSize(String(n))).run();
  scheduleToolbarSync();
}

function bumpSize(dir: 1 | -1): void {
  const cur = currentSize();
  const next = dir > 0 ? SIZES.find((s) => s > cur) ?? cur : [...SIZES].reverse().find((s) => s < cur) ?? cur;
  setSize(next);
}

function editLink(): void {
  const { from, to, empty } = editor.state.selection;
  const existing = editor.getAttributes('link').href as string | undefined;
  const rect = editor.view.coordsAtPos(from);
  showPopover({
    anchor: new DOMRect(rect.left, rect.top, Math.max(1, editor.view.coordsAtPos(to).left - rect.left), rect.bottom - rect.top),
    label: 'LINK',
    value: existing ?? '',
    placeholder: 'https://',
    hint: existing ? '⏎ save · clear to remove · esc cancel' : '⏎ save · esc cancel',
    onCommit: (raw) => {
      const href = raw.trim();
      const ch = editor.chain().focus();
      if (!href) {
        ch.extendMarkRange('link').unsetLink().run();
        return;
      }
      const url = /^[a-z][\w+.-]*:/i.test(href) ? href : `https://${href}`;
      if (empty && !existing) {
        ch.insertContent({ type: 'text', text: href, marks: [{ type: 'link', attrs: { href: url } }] }).run();
      } else {
        ch.extendMarkRange('link').setLink({ href: url }).run();
      }
    },
    onCancel: () => editor.commands.focus(),
  });
}

function gotoPagePrompt(): void {
  const bar = $('gotobar');
  const input = $<HTMLInputElement>('goto-input');
  $('goto-total').textContent = `of ${pageCount()}`;
  input.value = String(currentPage());
  bar.hidden = false;
  input.focus();
  input.select();
}

function closeGoto(go: boolean): void {
  const bar = $('gotobar');
  if (bar.hidden) return;
  const input = $<HTMLInputElement>('goto-input');
  const n = parseInt(input.value, 10);
  bar.hidden = true;
  // gotoPage clamps, so a number past the end lands at the top of the last page.
  if (go && Number.isFinite(n)) gotoPage(n);
  else editor.view.focus();
}

function toggleList(kind: 'bullet' | 'ordered'): void {
  const ch = editor.chain().focus();
  (kind === 'bullet' ? ch.toggleBulletList() : ch.toggleOrderedList()).run();
}

/** Commands that change the document; each becomes its own undo step. */
const EDIT_COMMANDS = new Set(['bold', 'italic', 'underline', 'strike', 'superscript', 'subscript', 'paragraph', 'h1', 'h2', 'h3',
  'bullet-list', 'ordered-list', 'indent', 'outdent', 'align-left', 'align-center', 'align-right', 'align-justify',
  'size-up', 'size-down', 'math', 'clear-format']);

function runCommand(name: string): void {
  if (!editor) return;
  if (EDIT_COMMANDS.has(name)) closeUndoGroup(editor);
  const c = () => editor.chain().focus();
  switch (name) {
    case 'bold': c().toggleBold().run(); break;
    case 'italic': c().toggleItalic().run(); break;
    case 'underline': c().toggleUnderline().run(); break;
    case 'strike': c().toggleStrike().run(); break;
    case 'superscript': c().toggleSuperscript().run(); break;
    case 'subscript': c().toggleSubscript().run(); break;
    case 'paragraph': c().setParagraph().run(); break;
    case 'h1': c().toggleHeading({ level: 1 }).run(); break;
    case 'h2': c().toggleHeading({ level: 2 }).run(); break;
    case 'h3': c().toggleHeading({ level: 3 }).run(); break;
    case 'bullet-list': toggleList('bullet'); break;
    case 'ordered-list': toggleList('ordered'); break;
    case 'indent': c().indent().run(); break;
    case 'outdent': c().outdent().run(); break;
    case 'align-left': c().setTextAlign('left').run(); break;
    case 'align-center': c().setTextAlign('center').run(); break;
    case 'align-right': c().setTextAlign('right').run(); break;
    case 'align-justify': c().setTextAlign('justify').run(); break;
    case 'size-up': bumpSize(1); break;
    case 'size-down': bumpSize(-1); break;
    case 'link': editLink(); break;
    case 'math': c().insertMath().run(); break;
    case 'clear-format': c().unsetAllMarks().run(); break;
    case 'toggle-toolbar': setToolbarOpen(!document.body.classList.contains('toolbar-open')); syncToolbar(); break;
    case 'toggle-theme': toggleTheme(); break;
    case 'toggle-page-numbers': {
      showPageNumbers = !showPageNumbers;
      try {
        localStorage.setItem('quire.pageNumbers', showPageNumbers ? '1' : '0');
      } catch {}
      const el = $('page-indicator');
      el.textContent = showPageNumbers ? `Page numbers on · ${currentPage()} / ${pageCount()}` : 'Page numbers off';
      indicatorHeldUntil = performance.now() + 1100;
      el.classList.add('show');
      window.clearTimeout(indicatorTimer);
      indicatorTimer = window.setTimeout(() => el.classList.remove('show'), 1100);
      break;
    }
    case 'toggle-typewriter': setMode('typewriter', !document.body.classList.contains('typewriter')); break;
    case 'toggle-focus': setMode('focus-mode', !document.body.classList.contains('focus-mode')); break;
    case 'goto-page': gotoPagePrompt(); break;
    case 'goto-start': gotoStart(); break;
    case 'goto-end': gotoEnd(); break;
    case 'find': openFind(); break;
    case 'zoom-in': stepZoom(1); break;
    case 'zoom-out': stepZoom(-1); break;
    case 'zoom-reset': setZoom(DEFAULT_ZOOM); break;
    case 'find-next': findStep(1); break;
    case 'find-prev': findStep(-1); break;
  }
  scheduleToolbarSync();
}

async function runMenu(name: string): Promise<void> {
  if (!editor) return;
  switch (name) {
    case 'save': await save(false); break;
    case 'save-as': await save(true); break;
    case 'export-pdf': await api.exportPdf(stem()); break;
    case 'print': api.print(); break;
  }
}

// ---------- wiring ----------

// If the page already handled a shortcut (e.g. ProseMirror's ⌘B), the menu accelerator for the
// same key may still arrive over IPC; running it again would undo the toggle.
let lastHandledKey = 0;
window.addEventListener('keydown', (e) => {
  if (e.defaultPrevented && (e.metaKey || e.ctrlKey)) lastHandledKey = performance.now();
});
/** Commands that make sense while typing in the find bar / size field / link box. */
const CHROME_COMMANDS = new Set(['find', 'find-next', 'find-prev', 'toggle-toolbar', 'toggle-theme', 'zoom-in', 'zoom-out', 'zoom-reset']);
api.onCommand((name) => {
  if (performance.now() - lastHandledKey < 60) return;
  const active = document.activeElement;
  const inField = active instanceof HTMLElement && /^(INPUT|TEXTAREA|SELECT)$/.test(active.tagName);
  if (inField && !CHROME_COMMANDS.has(name)) return;
  runCommand(name);
});
api.onFileChanged(onFileChanged);
api.onReplaceWord(replaceWord);
api.onSaveRequest((id) => {
  void save(false).then((ok) => api.saveResult(id, ok), () => api.saveResult(id, false));
});
api.onMenu((n) => void runMenu(n));

$('btn-toolbar').addEventListener('click', () => runCommand('toggle-toolbar'));
$('btn-theme').addEventListener('click', toggleTheme);
let findScrollQueued = false;
scroller().addEventListener('scroll', () => {
  // Highlights follow the viewport while find is open.
  if (!$('findbar').hidden && !findScrollQueued) {
    findScrollQueued = true;
    requestAnimationFrame(() => {
      findScrollQueued = false;
      setVisibleRange(editor.view, visibleRange(editor.view, scroller()));
    });
  }
  if (showPageNumbers) requestAnimationFrame(() => updatePageIndicator());
}, { passive: true });
// Clicking the empty margin below the text puts the caret at the end, like a page would.
scroller().addEventListener('mousedown', (e) => {
  if (e.target === scroller() || e.target === $('page')) {
    e.preventDefault();
    const rect = $('editor').getBoundingClientRect();
    if (e.clientY > rect.bottom) editor.commands.focus('end');
    else editor.view.focus();
  }
});

const gotoInput = $<HTMLInputElement>('goto-input');
gotoInput.addEventListener('input', () => {
  gotoInput.value = gotoInput.value.replace(/\D/g, '').slice(0, 6);
});
gotoInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') { e.preventDefault(); closeGoto(true); }
  else if (e.key === 'Escape') { e.preventDefault(); closeGoto(false); }
  else if (e.key === 'ArrowUp' || e.key === 'ArrowDown') {
    e.preventDefault();
    const n = (parseInt(gotoInput.value, 10) || currentPage()) + (e.key === 'ArrowUp' ? 1 : -1);
    gotoInput.value = String(Math.max(1, Math.min(pageCount(), n)));
    gotoInput.select();
  }
});
gotoInput.addEventListener('blur', () => setTimeout(() => closeGoto(false), 0));

const findInput = $<HTMLInputElement>('find-input');
findInput.addEventListener('input', scheduleFind);
findInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') { e.preventDefault(); findStep(e.shiftKey ? -1 : 1); }
  if (e.key === 'Escape') { e.preventDefault(); closeFind(); }
});
$('find-next').addEventListener('click', () => findStep(1));
$('find-prev').addEventListener('click', () => findStep(-1));
$('find-close').addEventListener('click', closeFind);

window.addEventListener(
  'keydown',
  (e) => {
    // Standard Mac document navigation (⌘↑/⌘↓) plus Word's ⌘Home/⌘End; ⇧ variants keep selecting.
    if (!e.metaKey || e.shiftKey || e.altKey || e.ctrlKey || !editor) return;
    const t = e.target as HTMLElement | null;
    if (t && /^(INPUT|TEXTAREA|SELECT)$/.test(t.tagName)) return;
    if (e.key === 'ArrowUp' || e.key === 'Home') {
      e.preventDefault();
      gotoStart();
    } else if (e.key === 'ArrowDown' || e.key === 'End') {
      e.preventDefault();
      gotoEnd();
    }
  },
  true,
);

window.addEventListener('keydown', (e) => {
  if (e.key === 'Meta') document.body.classList.add('cmd-down');
  if (e.key === 'Escape' && popoverOpen()) closePopover();
  else if (e.key === 'Escape' && !$('findbar').hidden) closeFind();
});
window.addEventListener('keyup', (e) => {
  if (e.key === 'Meta') document.body.classList.remove('cmd-down');
});
window.addEventListener('blur', () => document.body.classList.remove('cmd-down'));
window.addEventListener('dragover', (e) => e.preventDefault());
window.addEventListener('drop', (e) => {
  e.preventDefault();
  for (const f of Array.from(e.dataTransfer?.files ?? [])) api.openPath(api.pathForFile(f));
});

boot().catch((err) => {
  console.error(err);
  banner(`Something went wrong while opening: ${(err as Error).message}`);
});
