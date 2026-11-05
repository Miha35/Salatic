import { cpSync, createReadStream, existsSync, statSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { defineConfig, type Plugin } from 'vite';

const PDFJS_DIRS = ['cmaps', 'standard_fonts', 'wasm'];
const pdfjsRoot = resolve(import.meta.dirname, 'node_modules/pdfjs-dist');

/** Serve (dev) and copy (build) pdf.js auxiliary data under /pdfjs/. */
function pdfjsAssets(): Plugin {
  let outDir = 'dist';
  return {
    name: 'salatic-pdfjs-assets',
    configResolved(cfg) {
      outDir = cfg.build.outDir;
    },
    configureServer(server) {
      server.middlewares.use('/pdfjs', (req, res, next) => {
        const rel = decodeURIComponent((req.url ?? '').split('?')[0]).replace(/^\/+/, '');
        const [dir] = rel.split('/');
        const file = join(pdfjsRoot, rel);
        if (!PDFJS_DIRS.includes(dir) || !file.startsWith(pdfjsRoot) || !existsSync(file) || !statSync(file).isFile()) return next();
        if (file.endsWith('.wasm')) res.setHeader('Content-Type', 'application/wasm');
        createReadStream(file).pipe(res);
      });
    },
    closeBundle() {
      for (const d of PDFJS_DIRS) cpSync(join(pdfjsRoot, d), join(outDir, 'pdfjs', d), { recursive: true });
    },
  };
}

export default defineConfig({
  plugins: [pdfjsAssets()],
  resolve: {
    // Font packages hide their files behind "exports"; reach them by path
    alias: {
      'office-fonts': resolve(import.meta.dirname, 'node_modules/@docx-editor.dev/fonts/assets'),
      'cjk-fonts': resolve(import.meta.dirname, 'node_modules/@docx-editor.dev/fonts-cjk/assets'),
    },
  },
  build: {
    target: 'es2022',
    outDir: 'dist',
    assetsInlineLimit: 0,
    chunkSizeWarningLimit: 2500,
    sourcemap: false,
  },
  worker: {
    format: 'es',
  },
  server: {
    port: 5173,
  },
});
