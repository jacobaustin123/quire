// Quire main process: windows, menus, file IO, crash recovery and backups.
// Format conversion happens in the renderer; this process only moves bytes around.
'use strict';
const T0 = Date.now();
const trace = (what) => process.env.QUIRE_TRACE && console.log(`[trace] ${what} +${Date.now() - T0}ms`);

const { app, BrowserWindow, Menu, clipboard, dialog, ipcMain, nativeTheme, net, protocol, shell } = require('electron');
const { pathToFileURL } = require('node:url');
const fs = require('node:fs');
const fsp = require('node:fs/promises');
const path = require('node:path');
const crypto = require('node:crypto');

const isDev = process.argv.includes('--dev');
// Test harness: QUIRE_HARNESS=<module> drives the app from a script (see scripts/harness/README.md).
// QUIRE_USER_DATA isolates settings/recovery/backups from the real profile.
if (process.env.QUIRE_USER_DATA) app.setPath('userData', process.env.QUIRE_USER_DATA);
const DIST_DIR = path.join(__dirname, '..', 'dist');
const TRAFFIC_LIGHTS = { x: 18, y: 16 };
// The UI is served from quire://app/ rather than file:// so ES-module code splitting works (file://
// module loads are blocked by CORS) and V8 can cache compiled code between launches.
protocol.registerSchemesAsPrivileged([
  { scheme: 'quire', privileges: { standard: true, secure: true, supportFetchAPI: true, codeCache: true } },
]);
const SUPPORTED = new Set(['.docx', '.md', '.markdown', '.txt']);
const USER_DIR = app.getPath('userData');
const RECOVERY_DIR = path.join(USER_DIR, 'Recovery');
const BACKUP_DIR = path.join(USER_DIR, 'Backups');
const SETTINGS_PATH = path.join(USER_DIR, 'settings.json');
const { Backups } = require('./backups');
const backups = new Backups(BACKUP_DIR);

/** Per-window document state, keyed by webContents id. */
const docs = new Map();
/** Paths that have already been backed up this session (backup once, before the first overwrite). */
const backedUp = new Set();
const pendingOpen = [];
let quitApproved = false;
let ready = false;
/** Renderer save round-trips: id → resolve(ok). */
const pendingSaves = new Map();
let saveSeq = 0;

// ---------- settings ----------

function loadSettings() {
  try {
    return JSON.parse(fs.readFileSync(SETTINGS_PATH, 'utf8'));
  } catch {
    return {};
  }
}
const settings = { theme: 'system', ...loadSettings() };

function saveSettings() {
  fs.mkdirSync(USER_DIR, { recursive: true });
  fs.writeFileSync(SETTINGS_PATH, JSON.stringify(settings, null, 2));
}

/** Whether Quire currently shows its dark theme (an explicit choice, or System while macOS is dark). */
function effectiveDark() {
  return settings.theme === 'dark' || (settings.theme === 'system' && nativeTheme.shouldUseDarkColors);
}

/**
 * Applies Quire's theme by styling only — never via nativeTheme.themeSource, because changing the
 * app's macOS appearance re-lays out the title bar and the window buttons visibly jump.
 */
function applyTheme(theme) {
  settings.theme = theme;
  for (const w of BrowserWindow.getAllWindows()) {
    w.setBackgroundColor(backgroundColor());
    send(w, 'theme', theme);
  }
  saveSettings();
  buildMenu();
}

// Keep each window's native background in step with the page so a theme switch never shows a
// frame of the old colour behind the content.
// System appearance changed (only matters in System mode): keep native backgrounds in step.
nativeTheme.on('updated', () => {
  for (const w of BrowserWindow.getAllWindows()) w.setBackgroundColor(backgroundColor());
});

ipcMain.on('theme:get', (event) => {
  event.returnValue = settings.theme;
});

// ---------- helpers ----------

/**
 * Writes via temp file + fsync + rename so a crash never leaves a half-written file. Symlinks are
 * resolved first (a journal symlinked into iCloud/Dropbox stays a symlink) and the mode is kept.
 * `shouldCommit` can veto the rename at the last moment (used to drop stale recovery snapshots).
 */
async function atomicWrite(file, data, shouldCommit = () => true) {
  let target = file;
  try {
    target = await fsp.realpath(file);
  } catch {}
  let mode = 0o644;
  try {
    mode = (await fsp.stat(target)).mode & 0o777;
  } catch {}
  const tmp = path.join(path.dirname(target), `.${path.basename(target)}.quire-${process.pid}-${crypto.randomUUID().slice(0, 8)}`);
  const fh = await fsp.open(tmp, 'w', mode);
  try {
    await fh.writeFile(data);
    await fh.sync();
  } finally {
    await fh.close();
  }
  if (!shouldCommit()) {
    await fsp.rm(tmp, { force: true });
    return false;
  }
  try {
    await fsp.rename(tmp, target);
  } catch (err) {
    await fsp.rm(tmp, { force: true });
    throw err;
  }
  return true;
}

function recoveryIdForPath(p) {
  return 'file-' + crypto.createHash('sha1').update(p).digest('hex').slice(0, 16);
}

function recoveryFile(id) {
  return path.join(RECOVERY_DIR, `${id}.json`);
}

const DISK_POLL_MS = 2000;

async function diskVersion(file) {
  try {
    const s = await fsp.stat(file);
    return { mtime: s.mtimeMs, size: s.size };
  } catch {
    return null;
  }
}

/** True if `v` differs from the version this window last read or wrote. */
function changedOnDisk(st, v) {
  return Boolean(v) && (Math.abs(v.mtime - st.mtime) > 1 || (st.size >= 0 && v.size !== st.size));
}

/** Tells the renderer when its file was modified or removed by something else. */
async function checkDisk(win, st) {
  if (!st.path || !st.mtime || st.writing || win.isDestroyed()) return;
  const v = await diskVersion(st.path);
  if (st.writing || win.isDestroyed()) return;
  const key = v ? `${v.mtime}:${v.size}` : 'missing';
  if ((v && !changedOnDisk(st, v)) || st.announced === key) return;
  st.announced = key;
  send(win, 'file-changed', { missing: !v });
}

/** Keeps the on-disk version before the first overwrite this session (or after an external change). */
async function backupOnce(file) {
  if (backedUp.has(file) || !fs.existsSync(file)) return;
  backedUp.add(file);
  await backups.beforeOverwrite(file);
}

function backgroundColor() {
  return effectiveDark() ? '#1b1a19' : '#fbfaf7';
}

function stateFor(event) {
  return docs.get(event.sender.id);
}

function winFor(event) {
  return BrowserWindow.fromWebContents(event.sender);
}

function send(win, channel, ...args) {
  if (win && !win.isDestroyed()) win.webContents.send(channel, ...args);
}

function focused() {
  return BrowserWindow.getFocusedWindow() ?? BrowserWindow.getAllWindows()[0];
}

// ---------- windows ----------

/**
 * Opens a window for `filePath` (or an untitled doc). Reuses an existing window for the same file.
 * `recoveryId` restores an orphaned untitled recovery file.
 */
function createWindow(filePath = null, recoveryId = null) {
  if (filePath) {
    for (const [id, st] of docs) {
      if (st.path === filePath) {
        const w = BrowserWindow.getAllWindows().find((x) => x.webContents.id === id);
        if (w) {
          w.show();
          w.focus();
          return w;
        }
      }
    }
  }
  const parent = focused();
  const win = new BrowserWindow({
    width: 980,
    height: 1100,
    minWidth: 520,
    minHeight: 400,
    // Shown immediately in the page colour (not on ready-to-show), so launch feels instant even
    // while the editor is still starting.
    show: false,
    titleBarStyle: 'hiddenInset',
    trafficLightPosition: TRAFFIC_LIGHTS,
    backgroundColor: backgroundColor(),
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      v8CacheOptions: 'bypassHeatCheck',
      spellcheck: true,
    },
  });
  if (parent && !parent.isFullScreen()) {
    const [x, y] = parent.getPosition();
    win.setPosition(x + 28, y + 28);
  }
  const st = {
    path: filePath,
    dirty: false,
    forceClose: false,
    recoveryId: filePath ? recoveryIdForPath(filePath) : recoveryId ?? `untitled-${crypto.randomUUID()}`,
    recoveryGen: 0,
    pristine: !filePath && !recoveryId,
    prompting: false,
    mtime: 0,
    allowedSavePath: null,
    size: -1,
    writing: false,
    /** Disk version (mtime:size) already reported to the renderer, so it's announced once. */
    announced: null,
  };
  docs.set(win.webContents.id, st);
  const wcId = win.webContents.id;

  // The e2e harness must never steal keyboard focus from whatever the user is doing.
  win.once('ready-to-show', () => trace('ready-to-show'));
  win.webContents.on('console-message', (e) => {
    const msg = e.message ?? '';
    if (process.env.QUIRE_TRACE && msg.startsWith('[quire]')) trace(`renderer: ${msg}`);
  });
  win.webContents.on('render-process-gone', (_e, details) => {
    // Reloading re-runs doc:init, which restores the unsaved work from the recovery file.
    if (details.reason !== 'clean-exit' && !win.isDestroyed()) win.reload();
  });
  win.on('close', (e) => onWindowClose(win, st, e));
  win.on('closed', () => {
    clearInterval(poll);
    docs.delete(wcId);
  });
  // Notice edits made by other apps (Word, Dropbox…): poll cheaply, and check at once on focus.
  const poll = setInterval(() => checkDisk(win, st), DISK_POLL_MS);
  win.on('focus', () => checkDisk(win, st));
  win.webContents.setWindowOpenHandler(({ url }) => {
    openExternal(url);
    return { action: 'deny' };
  });
  win.webContents.on('will-navigate', (e) => e.preventDefault());
  win.webContents.on('context-menu', (_e, params) => {
    Menu.buildFromTemplate(contextMenuTemplate(win, params)).popup({ window: win });
  });
  if (filePath) {
    win.setRepresentedFilename(filePath);
    app.addRecentDocument(filePath);
  }
  win.setTitle(filePath ? path.basename(filePath) : 'Untitled');
  win.loadURL('quire://app/index.html');
  if (process.env.QUIRE_HARNESS) win.showInactive();
  else win.show();
  if (isDev) win.webContents.openDevTools({ mode: 'detach' });
  return win;
}

function onWindowClose(win, st, e) {
  if (st.forceClose || !st.dirty) {
    clearRecovery(st);
    return;
  }
  e.preventDefault();
  confirmClose(win, st).then((ok) => {
    if (!ok || win.isDestroyed()) return;
    st.forceClose = true;
    win.close();
  });
}

/** Asks the renderer to save; resolves with whether it succeeded. */
function requestSave(win) {
  return new Promise((resolve) => {
    const id = ++saveSeq;
    pendingSaves.set(id, resolve);
    send(win, 'save-request', id);
  });
}

/** Prompts about unsaved changes. Resolves true when the window may close without losing work. */
async function confirmClose(win, st) {
  if (st.forceClose || !st.dirty) return true;
  if (st.prompting) return false;
  st.prompting = true;
  try {
    const name = st.path ? path.basename(st.path) : 'Untitled';
    const { response } = await dialog.showMessageBox(win, {
      type: 'warning',
      buttons: ['Save', 'Cancel', "Don't Save"],
      defaultId: 0,
      cancelId: 1,
      message: `Do you want to save the changes you made to “${name}”?`,
      detail: "Your changes will be lost if you don't save them.",
    });
    if (response === 0) return (await requestSave(win)) && !st.dirty;
    return response === 2;
  } finally {
    st.prompting = false;
  }
}

function clearRecovery(st) {
  // Bumping the generation makes any in-flight snapshot discard itself instead of resurrecting.
  st.recoveryGen++;
  try {
    fs.rmSync(recoveryFile(st.recoveryId), { force: true });
  } catch {}
}

function openPath(p) {
  const resolved = path.resolve(p);
  if (!SUPPORTED.has(path.extname(resolved).toLowerCase())) return;
  if (!ready) {
    pendingOpen.push(resolved);
    return;
  }
  // Replace a pristine untitled window instead of stacking a new one on top of it.
  const w = focused();
  const st = w && docs.get(w.webContents.id);
  createWindow(resolved);
  if (st && !st.path && !st.dirty && st.pristine) {
    st.forceClose = true;
    w.close();
  }
}

async function showOpenDialog() {
  const res = await dialog.showOpenDialog({
    properties: ['openFile', 'multiSelections'],
    filters: [
      { name: 'Documents', extensions: ['docx', 'md', 'markdown', 'txt'] },
      { name: 'All Files', extensions: ['*'] },
    ],
  });
  for (const p of res.filePaths) openPath(p);
}

function openExternal(url) {
  if (/^(https?:|mailto:)/i.test(url)) shell.openExternal(url);
}

/** Focuses the next (or previous) Quire window, in the order they were opened. */
function cycleWindows(dir) {
  const wins = BrowserWindow.getAllWindows().filter((w) => w.isVisible() && !w.isMinimized()).sort((a, b) => a.id - b.id);
  if (wins.length < 2) return;
  const i = wins.indexOf(BrowserWindow.getFocusedWindow());
  const next = wins[(i + dir + wins.length) % wins.length];
  next.show();
  next.focus();
}

// ---------- context menu ----------

/**
 * Right-click menu: spelling suggestions (never applied automatically), then the usual editing
 * items. Replacements go through the editor (not webContents.replaceMisspelling) so they keep the
 * word's formatting and are undoable with ⌘Z.
 */
function contextMenuTemplate(win, params) {
  const items = [];
  const word = params.misspelledWord;
  if (word) {
    const suggestions = (params.dictionarySuggestions || []).slice(0, 5);
    for (const s of suggestions) {
      items.push({ label: s, click: () => send(win, 'replace-word', { x: params.x, y: params.y, word, replacement: s }) });
    }
    if (!suggestions.length) items.push({ label: 'No Guesses Found', enabled: false });
    items.push({ type: 'separator' });
    items.push({ label: 'Learn Spelling', click: () => win.webContents.session.addWordToSpellCheckerDictionary(word) });
    items.push({ type: 'separator' });
  }
  if (params.linkURL && /^(https?:|mailto:)/i.test(params.linkURL)) {
    items.push({ label: 'Open Link', click: () => openExternal(params.linkURL) });
    items.push({ label: 'Copy Link', click: () => clipboard.writeText(params.linkURL) });
    items.push({ type: 'separator' });
  }
  const sel = (params.selectionText || '').trim();
  if (sel && sel.length < 60) {
    items.push({ label: `Look Up “${sel.length > 24 ? sel.slice(0, 24) + '…' : sel}”`, click: () => win.webContents.showDefinitionForSelection() });
    items.push({ type: 'separator' });
  }
  items.push(
    { role: 'cut', enabled: params.editFlags.canCut },
    { role: 'copy', enabled: params.editFlags.canCopy },
    { role: 'paste', enabled: params.editFlags.canPaste },
  );
  return items;
}

// ---------- IPC ----------

ipcMain.handle('doc:init', async (event) => {
  const st = stateFor(event);
  const out = { path: st.path, name: st.path ? path.basename(st.path) : 'Untitled', bytes: null, mtime: 0, recovery: null };
  if (st.path) {
    try {
      out.bytes = await fsp.readFile(st.path);
      const stat = await fsp.stat(st.path);
      out.mtime = stat.mtimeMs;
      st.mtime = stat.mtimeMs;
      st.size = stat.size;
      st.announced = null;
      // Daily safety copy of documents edited in Quire; runs in the background.
      backups.onOpen(st.path).catch((err) => console.error('backup failed', err));
    } catch (err) {
      out.error = String(err.message ?? err);
    }
  }
  try {
    out.recovery = await fsp.readFile(recoveryFile(st.recoveryId), 'utf8');
  } catch {}
  return out;
});

ipcMain.handle('doc:choose-save-path', async (event, { defaultName, format }) => {
  const win = winFor(event);
  const st = stateFor(event);
  const ext = format === 'md' ? 'md' : format === 'txt' ? 'txt' : 'docx';
  const res = await dialog.showSaveDialog(win, {
    defaultPath: st.path ? st.path.replace(/\.[^.]+$/, `.${ext}`) : `${defaultName}.${ext}`,
    filters: [
      { name: 'Word Document', extensions: ['docx'] },
      { name: 'Markdown', extensions: ['md'] },
      { name: 'Plain Text', extensions: ['txt'] },
    ],
  });
  if (res.canceled || !res.filePath) return null;
  st.allowedSavePath = res.filePath;
  return res.filePath;
});

ipcMain.handle('doc:write', async (event, { filePath, data }) => {
  const st = stateFor(event);
  const win = winFor(event);
  // Only ever write where this window's document lives or where the user just chose to save it.
  if (filePath !== st.path && filePath !== st.allowedSavePath) throw new Error('Refusing to write to an unexpected path.');
  for (const [id, other] of docs) {
    if (id !== event.sender.id && other.path === filePath) throw new Error('That file is open in another window.');
  }
  if (filePath === st.path && st.mtime) {
    const current = await diskVersion(filePath);
    if (changedOnDisk(st, current)) {
      const { response } = await dialog.showMessageBox(win, {
        type: 'warning',
        buttons: ['Overwrite', 'Cancel'],
        defaultId: 1,
        cancelId: 1,
        message: `“${path.basename(filePath)}” was changed by another app since you opened it.`,
        detail: 'Overwrite it with your version? The version on disk will be kept in Backups.',
      });
      if (response !== 0) return { canceled: true };
      backedUp.delete(filePath);
    }
  }
  st.writing = true;
  try {
    await backupOnce(filePath);
    await atomicWrite(filePath, typeof data === 'string' ? data : Buffer.from(data));
    const v = await diskVersion(filePath);
    if (v) {
      st.mtime = v.mtime;
      st.size = v.size;
    }
    st.announced = null;
  } finally {
    st.writing = false;
  }
  if (st.path !== filePath) {
    clearRecovery(st);
    st.path = filePath;
    st.recoveryId = recoveryIdForPath(filePath);
  }
  st.allowedSavePath = null;
  backups.markEdited(filePath).catch((err) => console.error('backup index failed', err));
  st.dirty = false;
  clearRecovery(st);
  win.setDocumentEdited(false);
  win.setRepresentedFilename(filePath);
  win.setTitle(path.basename(filePath));
  app.addRecentDocument(filePath);
  return { path: filePath, name: path.basename(filePath) };
});

/** Re-reads the file after an external change; it becomes the new baseline for conflict checks. */
ipcMain.handle('doc:reload', async (event) => {
  const st = stateFor(event);
  const bytes = await fsp.readFile(st.path);
  const stat = await fsp.stat(st.path);
  st.mtime = stat.mtimeMs;
  st.size = stat.size;
  st.announced = null;
  return { bytes, mtime: stat.mtimeMs };
});

ipcMain.handle('confirm', async (event, { message, detail, ok }) => {
  const { response } = await dialog.showMessageBox(winFor(event), {
    type: 'warning',
    buttons: [ok || 'Save Anyway', 'Cancel'],
    defaultId: 1,
    cancelId: 1,
    message,
    detail,
  });
  return response === 0;
});

ipcMain.on('doc:dirty', (event, dirty) => {
  const st = stateFor(event);
  if (!st) return;
  st.dirty = dirty;
  st.pristine = false;
  winFor(event)?.setDocumentEdited(dirty);
});

ipcMain.on('doc:pristine', (event) => {
  const st = stateFor(event);
  if (st) st.pristine = true;
});

ipcMain.on('save-result', (_event, { id, ok }) => {
  pendingSaves.get(id)?.(Boolean(ok));
  pendingSaves.delete(id);
});

ipcMain.on('recovery:write', async (event, json) => {
  const st = stateFor(event);
  if (!st || st.forceClose || !st.dirty) return;
  const gen = st.recoveryGen;
  try {
    await fsp.mkdir(RECOVERY_DIR, { recursive: true });
    await atomicWrite(recoveryFile(st.recoveryId), json, () => gen === st.recoveryGen && st.dirty && !st.forceClose);
  } catch (err) {
    console.error('recovery write failed', err);
  }
});

ipcMain.handle('recovery:ask', async (event, { name, savedAt, fileChangedSince }) => {
  const when = new Date(savedAt).toLocaleString([], { dateStyle: 'medium', timeStyle: 'short' });
  const { response } = await dialog.showMessageBox(winFor(event), {
    type: 'warning',
    buttons: ['Restore', 'Discard Changes'],
    defaultId: 0,
    cancelId: 0,
    message: `Restore unsaved changes to “${name}”?`,
    detail:
      `Quire didn’t close normally. Your last unsaved changes are from ${when}.` +
      (fileChangedSince ? '\n\nThe file on disk was changed after that — restoring and saving would replace that newer version (it would be kept in Backups).' : '') +
      '\n\nDiscarded changes are kept for 30 days in Quire’s Recovery/Discarded folder.',
  });
  return response === 0;
});

/** Moves a declined recovery snapshot aside (kept 30 days) rather than deleting it. */
ipcMain.on('recovery:discard', (event) => {
  const st = stateFor(event);
  if (!st) return;
  const src = recoveryFile(st.recoveryId);
  const dir = path.join(RECOVERY_DIR, 'Discarded');
  try {
    fs.mkdirSync(dir, { recursive: true });
    fs.renameSync(src, path.join(dir, `${new Date().toISOString().replace(/[:.]/g, '-')} ${st.recoveryId}.json`));
    const cutoff = Date.now() - 30 * 24 * 3600 * 1000;
    for (const f of fs.readdirSync(dir)) {
      const p = path.join(dir, f);
      if (fs.statSync(p).mtimeMs < cutoff) fs.rmSync(p, { force: true });
    }
  } catch (err) {
    console.error('discard recovery failed', err);
  }
});

ipcMain.on('recovery:clear', (event) => {
  const st = stateFor(event);
  if (st) clearRecovery(st);
});

ipcMain.handle('pdf:export', async (event, { defaultName }) => {
  const win = winFor(event);
  const st = stateFor(event);
  const res = await dialog.showSaveDialog(win, {
    defaultPath: st.path ? st.path.replace(/\.[^.]+$/, '.pdf') : `${defaultName}.pdf`,
    filters: [{ name: 'PDF', extensions: ['pdf'] }],
  });
  if (res.canceled || !res.filePath) return null;
  const pdf = await win.webContents.printToPDF({
    printBackground: false,
    preferCSSPageSize: true,
    generateDocumentOutline: true,
  });
  await atomicWrite(res.filePath, pdf);
  return res.filePath;
});

ipcMain.on('print', (event) => {
  winFor(event)?.webContents.print({ printBackground: false });
});

ipcMain.on('open-external', (_event, url) => openExternal(url));
ipcMain.on('open-path', (_event, p) => openPath(p));
ipcMain.handle('settings:get', () => ({ theme: settings.theme }));

// ---------- menu ----------

/** Menu items that run editor commands in the renderer. */
function cmd(label, command, accelerator) {
  return { label, accelerator, click: () => send(focused(), 'cmd', command) };
}

function buildMenu() {
  const themeItem = (label, value) => ({
    label,
    type: 'radio',
    checked: settings.theme === value,
    click: () => applyTheme(value),
  });
  const template = [
    {
      label: app.name,
      submenu: [
        { role: 'about' },
        { type: 'separator' },
        { role: 'services' },
        { type: 'separator' },
        { role: 'hide' },
        { role: 'hideOthers' },
        { role: 'unhide' },
        { type: 'separator' },
        { role: 'quit' },
      ],
    },
    {
      label: 'File',
      submenu: [
        { label: 'New', accelerator: 'CmdOrCtrl+N', click: () => createWindow() },
        { label: 'Open…', accelerator: 'CmdOrCtrl+O', click: showOpenDialog },
        { role: 'recentDocuments', submenu: [{ role: 'clearRecentDocuments' }] },
        { type: 'separator' },
        { label: 'Close', accelerator: 'CmdOrCtrl+W', click: () => focused()?.close() },
        { label: 'Save', accelerator: 'CmdOrCtrl+S', click: () => send(focused(), 'menu', 'save') },
        { label: 'Save As…', accelerator: 'CmdOrCtrl+Shift+S', click: () => send(focused(), 'menu', 'save-as') },
        { type: 'separator' },
        { label: 'Export as PDF…', accelerator: 'CmdOrCtrl+Alt+P', click: () => send(focused(), 'menu', 'export-pdf') },
        { label: 'Print…', accelerator: 'CmdOrCtrl+P', click: () => send(focused(), 'menu', 'print') },
        { type: 'separator' },
        { label: 'Show Backups in Finder', click: () => { fs.mkdirSync(BACKUP_DIR, { recursive: true }); shell.openPath(BACKUP_DIR); } },
      ],
    },
    {
      label: 'Edit',
      submenu: [
        // Native roles: in text fields they undo the field; in the editor ProseMirror's history
        // picks up the resulting historyUndo/historyRedo input events.
        { role: 'undo' },
        { role: 'redo' },
        { type: 'separator' },
        { role: 'cut' },
        { role: 'copy' },
        { role: 'paste' },
        { role: 'pasteAndMatchStyle' },
        { role: 'selectAll' },
        { type: 'separator' },
        cmd('Find…', 'find', 'CmdOrCtrl+F'),
        cmd('Find Next', 'find-next'),
        cmd('Find Previous', 'find-prev'),
        { type: 'separator' },
        { label: 'Speech', submenu: [{ role: 'startSpeaking' }, { role: 'stopSpeaking' }] },
      ],
    },
    {
      label: 'Format',
      submenu: [
        cmd('Bold', 'bold', 'CmdOrCtrl+B'),
        cmd('Italic', 'italic', 'CmdOrCtrl+I'),
        cmd('Underline', 'underline', 'CmdOrCtrl+U'),
        cmd('Strikethrough', 'strike', 'CmdOrCtrl+Shift+X'),
        cmd('Superscript', 'superscript', 'CmdOrCtrl+.'),
        cmd('Subscript', 'subscript', 'CmdOrCtrl+,'),
        { type: 'separator' },
        cmd('Body Text', 'paragraph', 'CmdOrCtrl+Alt+0'),
        cmd('Heading 1', 'h1', 'CmdOrCtrl+Alt+1'),
        cmd('Heading 2', 'h2', 'CmdOrCtrl+Alt+2'),
        cmd('Heading 3', 'h3', 'CmdOrCtrl+Alt+3'),
        { type: 'separator' },
        cmd('Bulleted List', 'bullet-list', 'CmdOrCtrl+Shift+8'),
        cmd('Numbered List', 'ordered-list', 'CmdOrCtrl+Shift+7'),
        cmd('Increase Indent', 'indent', 'CmdOrCtrl+]'),
        cmd('Decrease Indent', 'outdent', 'CmdOrCtrl+['),
        { type: 'separator' },
        cmd('Align Left', 'align-left', 'CmdOrCtrl+Shift+L'),
        cmd('Center', 'align-center', 'CmdOrCtrl+Shift+E'),
        cmd('Align Right', 'align-right', 'CmdOrCtrl+Shift+R'),
        cmd('Justify', 'align-justify', 'CmdOrCtrl+Shift+J'),
        { type: 'separator' },
        cmd('Bigger', 'size-up', 'CmdOrCtrl+Shift+.'),
        cmd('Smaller', 'size-down', 'CmdOrCtrl+Shift+,'),
        { type: 'separator' },
        cmd('Insert Link…', 'link', 'CmdOrCtrl+K'),
        cmd('Insert Equation', 'math', 'CmdOrCtrl+Shift+M'),
        cmd('Clear Formatting', 'clear-format', 'CmdOrCtrl+\\'),
      ],
    },
    {
      label: 'View',
      submenu: [
        cmd('Show Formatting Bar', 'toggle-toolbar', 'CmdOrCtrl+T'),
        { type: 'separator' },
        { label: 'Appearance', submenu: [themeItem('System', 'system'), themeItem('Light', 'light'), themeItem('Dark', 'dark')] },
        cmd('Toggle Light/Dark', 'toggle-theme', 'CmdOrCtrl+Shift+D'),
        { type: 'separator' },
        cmd('Typewriter Mode', 'toggle-typewriter', 'CmdOrCtrl+Shift+T'),
        cmd('Focus on Paragraph', 'toggle-focus', 'CmdOrCtrl+Alt+F'),
        { type: 'separator' },
        cmd('Go to Page…', 'goto-page', 'CmdOrCtrl+G'),
        cmd('Go to Beginning', 'goto-start', 'CmdOrCtrl+Up'),
        cmd('Go to End', 'goto-end', 'CmdOrCtrl+Shift+G'),
        { type: 'separator' },
        cmd('Actual Size', 'zoom-reset', 'CmdOrCtrl+0'),
        cmd('Zoom In', 'zoom-in', 'CmdOrCtrl+='),
        cmd('Zoom Out', 'zoom-out', 'CmdOrCtrl+-'),
        { type: 'separator' },
        { role: 'togglefullscreen' },
        ...(isDev ? [{ role: 'toggleDevTools' }, { role: 'reload' }] : []),
      ],
    },
    {
      // Built by hand: Electron's stock windowMenu doesn't give ⌘` window cycling for these windows.
      label: 'Window',
      role: 'window',
      submenu: [
        { role: 'minimize' },
        { role: 'zoom' },
        { type: 'separator' },
        { label: 'Cycle Through Windows', accelerator: 'Cmd+`', click: () => cycleWindows(1) },
        { label: 'Cycle Backwards', accelerator: 'Cmd+Shift+`', click: () => cycleWindows(-1) },
        { type: 'separator' },
        { role: 'front' },
      ],
    },
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

ipcMain.on('toggle-theme', () => applyTheme(effectiveDark() ? 'light' : 'dark'));

// ---------- app lifecycle ----------

app.on('open-file', (e, p) => {
  e.preventDefault();
  openPath(p);
});

app.on('before-quit', (e) => {
  if (quitApproved) return;
  const dirty = BrowserWindow.getAllWindows().filter((w) => {
    const st = docs.get(w.webContents.id);
    return st && st.dirty && !st.forceClose;
  });
  if (!dirty.length) return;
  e.preventDefault();
  (async () => {
    for (const w of dirty) {
      const st = docs.get(w.webContents.id);
      if (!st || w.isDestroyed()) continue;
      w.show();
      w.focus();
      if (!(await confirmClose(w, st))) return; // Cancel (or a failed save) aborts the quit.
      st.forceClose = true;
    }
    quitApproved = true;
    app.quit();
  })();
});

app.on('window-all-closed', () => {
  // Stay resident on macOS (standard behavior) — reopening is then instant.
  if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});

app.whenReady().then(() => {
  trace('app ready');
  protocol.handle('quire', (req) => {
    const rel = decodeURIComponent(new URL(req.url).pathname);
    const file = path.normalize(path.join(DIST_DIR, rel));
    if (!file.startsWith(DIST_DIR + path.sep)) return new Response('Not found', { status: 404 });
    return net.fetch(pathToFileURL(file).toString());
  });
  buildMenu();
  ready = true;
  for (const arg of process.argv.slice(1)) {
    if (!arg.startsWith('-') && SUPPORTED.has(path.extname(arg).toLowerCase()) && fs.existsSync(arg)) {
      pendingOpen.push(path.resolve(arg));
    }
  }
  // Documents with unsaved changes during a crash come back automatically.
  let orphans = [];
  try {
    orphans = fs.readdirSync(RECOVERY_DIR).filter((f) => f.endsWith('.json'));
  } catch {}
  for (const f of orphans) {
    if (f.startsWith('untitled-')) {
      createWindow(null, f.replace(/\.json$/, ''));
      continue;
    }
    try {
      const rec = JSON.parse(fs.readFileSync(path.join(RECOVERY_DIR, f), 'utf8'));
      if (rec.path && !pendingOpen.includes(rec.path)) pendingOpen.push(rec.path);
    } catch {}
  }
  for (const p of pendingOpen.splice(0)) createWindow(p);
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
  if (process.env.QUIRE_HARNESS) require('./harness').run(process.env.QUIRE_HARNESS);
});
