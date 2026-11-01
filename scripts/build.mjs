// Bundles the renderer into dist/ with esbuild. Kept to a single JS + CSS file so the window paints fast.
import * as esbuild from 'esbuild';
import fs from 'node:fs';

const watch = process.argv.includes('--watch');
fs.rmSync('dist', { recursive: true, force: true });
fs.mkdirSync('dist', { recursive: true });
fs.copyFileSync('src/index.html', 'dist/index.html');

const ctx = await esbuild.context({
  entryPoints: { app: 'src/app.ts' },
  bundle: true,
  outdir: 'dist',
  // ESM + splitting: KaTeX and the file-format converters load only when first needed.
  format: 'esm',
  splitting: true,
  chunkNames: 'chunks/[name]-[hash]',
  target: 'chrome130',
  minify: !watch,
  sourcemap: watch ? 'inline' : false,
  loader: { '.woff2': 'file', '.woff': 'file', '.ttf': 'file' },
  assetNames: 'fonts/[name]',
  logLevel: 'info',
  legalComments: 'none',
});
if (watch) {
  await ctx.watch();
} else {
  await ctx.rebuild();
  await ctx.dispose();
}
