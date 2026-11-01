import { createRequire } from 'node:module';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { beforeEach, describe, expect, it } from 'vitest';

const require = createRequire(import.meta.url);
const { Backups } = require('../main/backups.js') as {
  Backups: new (root: string, opts?: { perFile?: number; maxBytes?: number; now?: () => Date }) => {
    markEdited(f: string): Promise<void>;
    onOpen(f: string): Promise<string | null>;
    beforeOverwrite(f: string): Promise<string | null>;
    dirFor(f: string): string;
    copies(dir: string): Promise<string[]>;
  };
};

let tmp: string;
let clock: Date;
const now = () => clock;
const at = (iso: string) => {
  clock = new Date(iso);
};

beforeEach(() => {
  tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'quire-backups-'));
  at('2026-10-31T10:00:00');
});

function doc(name: string, content: string): string {
  const p = path.join(tmp, 'docs', name);
  fs.mkdirSync(path.dirname(p), { recursive: true });
  fs.writeFileSync(p, content);
  return p;
}

describe('backups', () => {
  it('ignores files never edited in Quire', async () => {
    const b = new Backups(path.join(tmp, 'B'), { now });
    const f = doc('a.docx', 'v1');
    expect(await b.onOpen(f)).toBeNull();
    expect(fs.existsSync(path.join(tmp, 'B', 'index.json'))).toBe(false);
  });

  it('copies at most once per day, and only when the content changed', async () => {
    const b = new Backups(path.join(tmp, 'B'), { now });
    const f = doc('a.docx', 'v1');
    await b.markEdited(f);
    expect(await b.onOpen(f)).toMatch(/2026-10-31 10\.00 a\.docx$/);
    fs.writeFileSync(f, 'v2');
    expect(await b.onOpen(f)).toBeNull(); // same day
    at('2026-11-01T09:00:00');
    expect(await b.onOpen(f)).toMatch(/2026-11-01 09\.00 a\.docx$/);
    at('2026-11-02T09:00:00');
    expect(await b.onOpen(f)).toBeNull(); // unchanged since last copy
    const copies = await b.copies(b.dirFor(f));
    expect(copies.map((c) => fs.readFileSync(c, 'utf8'))).toEqual(['v1', 'v2']);
  });

  it('keeps the on-disk version before an overwrite, even on the same day', async () => {
    const b = new Backups(path.join(tmp, 'B'), { now });
    const f = doc('a.md', 'one');
    await b.markEdited(f);
    await b.onOpen(f);
    fs.writeFileSync(f, 'two (written by another app)');
    expect(await b.beforeOverwrite(f)).toMatch(/2026-10-31 10\.00-02 a\.md$/);
    expect(await b.beforeOverwrite(f)).toBeNull(); // already backed up
    fs.writeFileSync(f, 'three');
    expect(await b.beforeOverwrite(f)).toMatch(/2026-10-31 10\.00-03 a\.md$/);
    const copies = await b.copies(b.dirFor(f));
    expect(copies.map((c) => fs.readFileSync(c, 'utf8'))).toEqual(['one', 'two (written by another app)', 'three']);
  });

  it('caps copies per document', async () => {
    const b = new Backups(path.join(tmp, 'B'), { now, perFile: 3 });
    const f = doc('a.txt', 'x');
    await b.markEdited(f);
    for (let day = 1; day <= 6; day++) {
      at(`2026-11-0${day}T08:00:00`);
      fs.writeFileSync(f, `version ${day}`);
      await b.onOpen(f);
    }
    const copies = await b.copies(b.dirFor(f));
    expect(copies.map((c) => fs.readFileSync(c, 'utf8'))).toEqual(['version 4', 'version 5', 'version 6']);
  });

  it('caps total size oldest-first but always keeps each document’s newest copy', async () => {
    const b = new Backups(path.join(tmp, 'B'), { now, maxBytes: 250 });
    const big = doc('big.docx', 'a'.repeat(100));
    const small = doc('small.docx', 'b'.repeat(100));
    await b.markEdited(big);
    await b.markEdited(small);
    at('2026-11-01T08:00:00');
    await b.onOpen(big);
    at('2026-11-02T08:00:00');
    fs.writeFileSync(big, 'c'.repeat(100));
    await b.onOpen(big);
    at('2026-11-03T08:00:00');
    await b.onOpen(small); // total would be 300 > 250: drops big's oldest copy
    expect((await b.copies(b.dirFor(big))).length).toBe(1);
    expect((await b.copies(b.dirFor(small))).length).toBe(1);
    expect(fs.readFileSync((await b.copies(b.dirFor(big)))[0], 'utf8')).toBe('c'.repeat(100));
  });

  it('serializes concurrent operations', async () => {
    const b = new Backups(path.join(tmp, 'B'), { now });
    const f = doc('a.docx', 'v1');
    await b.markEdited(f);
    const results = await Promise.all([b.onOpen(f), b.onOpen(f), b.beforeOverwrite(f)]);
    expect(results.filter(Boolean).length).toBe(1);
  });
});
