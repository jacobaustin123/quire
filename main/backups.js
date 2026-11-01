// Quire backups: dated copies of documents the user has edited in Quire.
//
// Layout: <root>/<basename> (<id>)/<YYYY-MM-DD HH.MM> <basename>, plus <root>/index.json listing
// tracked files. A file becomes tracked once Quire saves it. Copies are taken
//   - when a tracked file is opened, at most once per day, and only if it differs from the newest
//     copy (so unchanged files never accumulate copies), and
//   - before Quire overwrites a file, if the on-disk version differs from the newest copy.
// Retention: at most `perFile` copies per document and `maxBytes` overall, oldest first, but the
// newest copy of every document is always kept.
'use strict';

const fs = require('node:fs');
const fsp = require('node:fs/promises');
const path = require('node:path');
const crypto = require('node:crypto');

const DEFAULTS = { perFile: 100, maxBytes: 1024 ** 3 };

function pad(n) {
  return String(n).padStart(2, '0');
}

/** Local-time stamp used in copy names; sorts chronologically as a string. */
function stamp(d) {
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}.${pad(d.getMinutes())}`;
}

async function sha1(file) {
  return crypto.createHash('sha1').update(await fsp.readFile(file)).digest('hex');
}

class Backups {
  /**
   * @param {string} root Backup directory.
   * @param {{perFile?: number, maxBytes?: number, now?: () => Date}} [opts]
   */
  constructor(root, opts = {}) {
    this.root = root;
    this.perFile = opts.perFile ?? DEFAULTS.perFile;
    this.maxBytes = opts.maxBytes ?? DEFAULTS.maxBytes;
    this.now = opts.now ?? (() => new Date());
    // Serializes all operations so concurrent opens/saves never race on the same folder.
    this.queue = Promise.resolve();
  }

  /** Runs `fn` after every previously scheduled operation. */
  run(fn) {
    const next = this.queue.then(fn, fn);
    this.queue = next.catch(() => {});
    return next;
  }

  indexPath() {
    return path.join(this.root, 'index.json');
  }

  async readIndex() {
    try {
      return JSON.parse(await fsp.readFile(this.indexPath(), 'utf8'));
    } catch {
      return { tracked: {} };
    }
  }

  async writeIndex(index) {
    await fsp.mkdir(this.root, { recursive: true });
    const tmp = `${this.indexPath()}.tmp-${process.pid}`;
    await fsp.writeFile(tmp, JSON.stringify(index, null, 2));
    await fsp.rename(tmp, this.indexPath());
  }

  /** Folder holding copies of `file`. */
  dirFor(file) {
    const id = crypto.createHash('sha1').update(file).digest('hex').slice(0, 8);
    return path.join(this.root, `${path.basename(file).replace(/[/:]/g, '_')} (${id})`);
  }

  /** Copies in `dir`, oldest first. */
  async copies(dir) {
    let names = [];
    try {
      names = (await fsp.readdir(dir)).filter((n) => /^\d{4}-\d\d-\d\d \d\d\.\d\d(-\d+)? /.test(n));
    } catch {}
    return names.sort().map((n) => path.join(dir, n));
  }

  /** Marks `file` as edited in Quire (called after a successful save). */
  markEdited(file) {
    return this.run(async () => {
      const index = await this.readIndex();
      if (index.tracked[file]) return;
      index.tracked[file] = { since: this.now().toISOString() };
      await this.writeIndex(index);
    });
  }

  /** Background check when a document is opened: at most one copy per day, only if changed. */
  onOpen(file) {
    return this.run(async () => {
      const index = await this.readIndex();
      if (!index.tracked[file]) return null;
      return this.snapshot(file, { oncePerDay: true });
    });
  }

  /** Called before Quire overwrites `file`: keeps the on-disk version if it isn't backed up yet. */
  beforeOverwrite(file) {
    return this.run(() => this.snapshot(file, { oncePerDay: false }));
  }

  /** Copies `file` unless the newest copy matches it (or is from today, if `oncePerDay`). */
  async snapshot(file, { oncePerDay }) {
    if (!fs.existsSync(file)) return null;
    const dir = this.dirFor(file);
    const existing = await this.copies(dir);
    const newest = existing[existing.length - 1];
    const now = this.now();
    if (newest) {
      if (oncePerDay && path.basename(newest).slice(0, 10) === stamp(now).slice(0, 10)) return null;
      if ((await sha1(newest)) === (await sha1(file))) return null;
    }
    await fsp.mkdir(dir, { recursive: true });
    let dest = path.join(dir, `${stamp(now)} ${path.basename(file)}`);
    // Same-minute copies (e.g. open, then a save) get "-2", "-3"…: '-' sorts after ' ', so name
    // order stays chronological and the newest copy is always last.
    for (let i = 2; fs.existsSync(dest); i++) dest = path.join(dir, `${stamp(now)}-${pad(i)} ${path.basename(file)}`);
    await fsp.copyFile(file, dest);
    await this.prune();
    return dest;
  }

  /** Enforces the per-document and total-size caps. */
  async prune() {
    let dirs = [];
    try {
      dirs = (await fsp.readdir(this.root, { withFileTypes: true })).filter((d) => d.isDirectory()).map((d) => path.join(this.root, d.name));
    } catch {
      return;
    }
    const all = [];
    for (const dir of dirs) {
      const list = await this.copies(dir);
      for (const old of list.slice(0, Math.max(0, list.length - this.perFile))) await fsp.rm(old, { force: true });
      const kept = list.slice(Math.max(0, list.length - this.perFile));
      for (const [i, file] of kept.entries()) {
        const { size } = await fsp.stat(file);
        all.push({ file, size, key: path.basename(file), newest: i === kept.length - 1 });
      }
    }
    let total = all.reduce((sum, c) => sum + c.size, 0);
    if (total <= this.maxBytes) return;
    // Oldest first across all documents; never the newest copy of any document.
    for (const c of all.filter((c) => !c.newest).sort((a, b) => a.key.localeCompare(b.key))) {
      if (total <= this.maxBytes) break;
      await fsp.rm(c.file, { force: true });
      total -= c.size;
    }
  }
}

module.exports = { Backups, stamp };
