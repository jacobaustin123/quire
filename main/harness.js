// Drives Quire from a script for automated end-to-end checks. Not used in normal runs.
'use strict';

const { app, BrowserWindow } = require('electron');
const fs = require('node:fs');
const path = require('node:path');

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/** Runs the harness module at `scriptPath`, then quits with its exit code. */
exports.run = async function run(scriptPath) {
  const ctx = {
    app,
    sleep,
    windows: () => BrowserWindow.getAllWindows(),
    win: () => BrowserWindow.getAllWindows()[0],
    /** Waits until the first window's editor exists. */
    async ready(win = ctx.win()) {
      for (let i = 0; i < 200; i++) {
        if (win && !win.webContents.isLoading()) {
          const ok = await win.webContents.executeJavaScript('!!document.querySelector(".ProseMirror")').catch(() => false);
          if (ok) return;
        }
        await sleep(50);
        win = win ?? ctx.win();
      }
      throw new Error('editor never became ready');
    },
    eval: (js, win = ctx.win()) => win.webContents.executeJavaScript(js),
    async capture(file, win = ctx.win()) {
      const img = await win.webContents.capturePage();
      fs.mkdirSync(path.dirname(file), { recursive: true });
      fs.writeFileSync(file, img.toPNG());
    },
    /** Sends a real key event (goes through the page's keymaps). */
    async key(keyCode, modifiers = [], win = ctx.win()) {
      win.webContents.sendInputEvent({ type: 'keyDown', keyCode, modifiers });
      if (keyCode.length === 1) win.webContents.sendInputEvent({ type: 'char', keyCode, modifiers });
      win.webContents.sendInputEvent({ type: 'keyUp', keyCode, modifiers });
      await sleep(15);
    },
    async type(text, win = ctx.win()) {
      for (const ch of text) {
        if (ch === '\n') await ctx.key('Enter', [], win);
        else win.webContents.sendInputEvent({ type: 'char', keyCode: ch });
        await sleep(8);
      }
      await sleep(30);
    },
  };
  let code = 0;
  try {
    await require(path.resolve(scriptPath))(ctx);
  } catch (err) {
    console.error('[harness] FAILED:', err && err.stack ? err.stack : err);
    code = 1;
  }
  for (const w of BrowserWindow.getAllWindows()) {
    w.removeAllListeners('close');
    w.destroy();
  }
  app.exit(code);
};
