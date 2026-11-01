#!/usr/bin/env node
/**
 * Scans commits about to be pushed for things that must never reach the public repo: credentials,
 * personal paths, private terms, and text copied from private documents (e.g. the journal used
 * for fidelity testing). Run by the pre-push hook (.githooks/pre-push), or by hand:
 *
 *   node scripts/prepush-scan.mjs                 # commits not yet on any remote
 *   node scripts/prepush-scan.mjs origin/main..   # an explicit range
 *
 * Private terms and documents are listed in .git/info/push-denylist, which is never committed
 * (listing them in the repo would itself leak them). One entry per line:
 *
 *   term: Some private phrase      (case-insensitive literal)
 *   doc: /path/to/Private.docx      (any 8-word run of its text is blocked; read-only)
 *
 * Added or changed binary files (images, .docx, ...) are blocked too, because their content
 * can't be checked as text (a screenshot can show a private document). Once you have looked at
 * them, push with QUIRE_ALLOW_BINARIES=1.
 */
import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';

const SHINGLE = 8;

const SECRET_PATTERNS = [
  ['private key', /-----BEGIN [A-Z ]*PRIVATE KEY-----/],
  ['GitHub token', /\b(gh[pousr]_[A-Za-z0-9]{30,}|github_pat_[A-Za-z0-9_]{40,})\b/],
  ['Anthropic/OpenAI key', /\bsk-(ant-)?[A-Za-z0-9_-]{20,}/],
  ['AWS access key', /\b(AKIA|ASIA)[0-9A-Z]{16}\b/],
  ['Slack token', /\bxox[abprs]-[A-Za-z0-9-]{10,}/],
  ['Google API key', /\bAIza[0-9A-Za-z_-]{35}\b/],
  ['assigned secret', /\b(password|passwd|secret|api[_-]?key|token)\s*[:=]\s*['"][^'"\s]{8,}['"]/i],
  // Home directories and the agent scratchpad leak usernames and private file names.
  ['home path', /\/Users\/(?!Shared\/)[^/\s'"`]+\//],
  ['home path', /\/home\/[a-z][^/\s'"`]*\//],
  ['scratch path', /\/private\/tmp\/claude-/],
];

function git(...args) {
  return execFileSync('git', args, { encoding: 'utf8', maxBuffer: 1 << 30 });
}

function words(text) {
  return text.toLowerCase().match(/[\p{L}\p{N}']+/gu) ?? [];
}

/** All runs of SHINGLE consecutive words, joined with spaces. */
function shingles(text) {
  const w = words(text);
  const out = [];
  for (let i = 0; i + SHINGLE <= w.length; i++) out.push(w.slice(i, i + SHINGLE).join(' '));
  return out;
}

function docText(file) {
  // textutil reads the file and writes to stdout only; the document is never modified.
  if (/\.(docx?|rtf|odt|html?)$/i.test(file)) {
    return execFileSync('textutil', ['-convert', 'txt', '-stdout', file], { encoding: 'utf8', maxBuffer: 1 << 30 });
  }
  return fs.readFileSync(file, 'utf8');
}

function loadDenylist() {
  const terms = [];
  const docShingles = new Map(); // shingle -> doc name
  const listFile = path.join(git('rev-parse', '--git-common-dir').trim(), 'info', 'push-denylist');
  if (!fs.existsSync(listFile)) return { terms, docShingles, listFile, found: false };
  for (const raw of fs.readFileSync(listFile, 'utf8').split('\n')) {
    const line = raw.trim();
    if (!line || line.startsWith('#')) continue;
    const m = /^(term|doc):\s*(.+)$/.exec(line);
    if (!m) throw new Error(`push-denylist: can't parse "${line}"`);
    if (m[1] === 'term') terms.push(m[2].toLowerCase());
    else {
      const file = m[2].replace(/^~(?=\/)/, process.env.HOME);
      // A missing private doc must fail loudly: silently skipping it would disable the check.
      if (!fs.existsSync(file)) throw new Error(`push-denylist: document not found: ${file}`);
      for (const s of shingles(docText(file))) docShingles.set(s, path.basename(file));
    }
  }
  return { terms, docShingles, listFile, found: true };
}

/** Commit range(s) to scan: from pre-push stdin, the command line, or everything unpushed. */
function ranges() {
  if (process.argv[2]) return [process.argv[2]];
  if (process.env.QUIRE_PREPUSH_STDIN) {
    const out = [];
    for (const line of process.env.QUIRE_PREPUSH_STDIN.split('\n')) {
      const [, localSha, , remoteSha] = line.trim().split(/\s+/);
      if (!localSha || /^0+$/.test(localSha)) continue; // deleting a remote branch
      out.push(/^0+$/.test(remoteSha) ? `${localSha} --not --remotes` : `${remoteSha}..${localSha}`);
    }
    return out;
  }
  return ['HEAD --not --remotes'];
}

function scan() {
  const { terms, docShingles, listFile, found } = loadDenylist();
  const problems = [];
  const report = (where, what, text) => problems.push(`${where}: ${what}\n    ${text.trim().slice(0, 160)}`);

  for (const range of ranges()) {
    const revs = range.split(' ');
    const commits = git('rev-list', ...revs).split('\n').filter(Boolean);
    for (const sha of commits) {
      const short = sha.slice(0, 7);
      const message = git('log', '-1', '--format=%B', sha);
      const added = [];
      let file = '';
      for (const line of git('show', '--format=', '--no-color', '--unified=0', sha).split('\n')) {
        if (line.startsWith('+++ ')) file = line.replace(/^\+\+\+ (b\/)?/, '');
        else if (line.startsWith('+')) added.push([file, line.slice(1)]);
        else if (line.startsWith('Binary files') && !process.env.QUIRE_ALLOW_BINARIES) {
          report(short, 'binary file added or changed (check it, then push with QUIRE_ALLOW_BINARIES=1)', line);
        }
      }
      const items = [...message.split('\n').map((l) => ['commit message', l]), ...added];
      for (const [where, text] of items) {
        const loc = `${short} ${where}`;
        for (const [name, re] of SECRET_PATTERNS) if (re.test(text)) report(loc, name, text);
        const lower = text.toLowerCase();
        for (const t of terms) if (lower.includes(t)) report(loc, `private term "${t}"`, text);
        for (const s of shingles(text)) {
          const doc = docShingles.get(s);
          if (doc) {
            report(loc, `text from ${doc}`, text);
            break;
          }
        }
      }
    }
  }

  if (!found) console.warn(`prepush-scan: no ${listFile}; only checking secrets and paths.`);
  if (problems.length) {
    console.error(`prepush-scan: ${problems.length} problem(s); push blocked.\n`);
    for (const p of problems) console.error(`  ${p}`);
    console.error('\nFix the commits (e.g. git rebase), or bypass a false alarm with git push --no-verify.');
    process.exit(1);
  }
  console.log(`prepush-scan: clean (${docShingles.size} private-document phrases, ${terms.length} terms checked).`);
}

scan();
