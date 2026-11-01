#!/usr/bin/env node
// Rasterizes build/icon.svg into build/icon.iconset, build/icon.icns,
// build/icon.png (1024) and build/icon-256.png.
//
// ImageMagick's built-in SVG renderer mangles gradients, clip paths and
// filters, so rendering is done by Chromium via headless Electron: the SVG is
// drawn onto a canvas at each target size, so small sizes are rasterized from
// the vector rather than downsampled from 1024.
//
// Usage: node scripts/make-icon.mjs   (or: npm run icon)

import childProcess from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import url from 'node:url';

const ROOT = path.resolve(path.dirname(url.fileURLToPath(import.meta.url)), '..');
const BUILD = path.join(ROOT, 'build');
const SVG = path.join(BUILD, 'icon.svg');
const ICONSET = path.join(BUILD, 'icon.iconset');

/** Iconset entries as [file name, pixel size]. */
const ICONSET_ENTRIES = [16, 32, 128, 256, 512].flatMap((n) => [
  [`icon_${n}x${n}.png`, n],
  [`icon_${n}x${n}@2x.png`, n * 2],
]);

/**
 * Renders an SVG string to PNGs at the given sizes inside an Electron renderer.
 *
 * @param {import('electron').BrowserWindow} win A window with a blank page loaded.
 * @param {string} svg The SVG source.
 * @param {number[]} sizes Square output sizes in pixels.
 * @returns {Promise<Map<number, Buffer>>} PNG bytes keyed by size.
 */
async function renderSizes(win, svg, sizes) {
  const svgB64 = Buffer.from(svg).toString('base64');
  const dataUrls = await win.webContents.executeJavaScript(`(async () => {
    const img = new Image();
    img.src = 'data:image/svg+xml;base64,${svgB64}';
    await img.decode();
    const out = {};
    for (const n of ${JSON.stringify(sizes)}) {
      const canvas = document.createElement('canvas');
      canvas.width = canvas.height = n;
      const ctx = canvas.getContext('2d');
      ctx.imageSmoothingQuality = 'high';
      ctx.drawImage(img, 0, 0, n, n);
      out[n] = canvas.toDataURL('image/png');
    }
    return out;
  })()`);
  return new Map(
    Object.entries(dataUrls).map(([n, u]) => [Number(n), Buffer.from(u.split(',')[1], 'base64')]),
  );
}

/** Electron side: render every size and write the PNG outputs. */
async function runInElectron() {
  const { app, BrowserWindow } = await import('electron');
  // Keep this throwaway run from touching the real app's profile directory.
  app.setPath('userData', fs.mkdtempSync(path.join(os.tmpdir(), 'quire-icon-')));
  app.disableHardwareAcceleration();
  app.dock?.hide();
  await app.whenReady();
  try {
    const win = new BrowserWindow({ show: false, webPreferences: { offscreen: true } });
    await win.loadURL('data:text/html,<!doctype html><title>icon</title>');
    const sizes = [...new Set([...ICONSET_ENTRIES.map(([, n]) => n), 256, 1024])];
    const pngs = await renderSizes(win, fs.readFileSync(SVG, 'utf8'), sizes);

    fs.rmSync(ICONSET, { recursive: true, force: true });
    fs.mkdirSync(ICONSET, { recursive: true });
    for (const [name, n] of ICONSET_ENTRIES) fs.writeFileSync(path.join(ICONSET, name), pngs.get(n));
    fs.writeFileSync(path.join(BUILD, 'icon.png'), pngs.get(1024));
    fs.writeFileSync(path.join(BUILD, 'icon-256.png'), pngs.get(256));
    app.exit(0);
  } catch (err) {
    console.error(err);
    app.exit(1);
  }
}

/** Node side: re-launch this file under Electron, then build the .icns. */
async function runInNode() {
  const { default: electronBin } = await import('electron');
  const result = childProcess.spawnSync(electronBin, [url.fileURLToPath(import.meta.url)], {
    stdio: ['ignore', 'inherit', 'pipe'],
    env: { ...process.env, ELECTRON_ENABLE_LOGGING: '' },
  });
  if (result.status !== 0) {
    process.stderr.write(result.stderr);
    throw new Error(`Electron render failed (exit ${result.status})`);
  }
  childProcess.execFileSync('iconutil', ['-c', 'icns', ICONSET, '-o', path.join(BUILD, 'icon.icns')], {
    stdio: 'inherit',
  });
  console.log(`Wrote ${path.relative(ROOT, ICONSET)}/, build/icon.icns, build/icon.png, build/icon-256.png`);
}

// No top-level await under Electron: an ESM main module that awaits before
// evaluation finishes blocks the 'ready' event and deadlocks.
if (process.versions.electron) runInElectron();
else await runInNode();
