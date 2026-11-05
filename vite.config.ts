import { cpSync, createReadStream, existsSync, statSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { defineConfig, type Plugin } from 'vite';

const nm = (p: string): string => resolve(import.meta.dirname, 'node_modules', p);

/** Files served as-is under fixed URLs (pdf.js data, Tesseract worker/core/models). */
const STATIC: { url: string; from: string }[] = [
  ...['cmaps', 'standard_fonts', 'wasm'].map((d) => ({ url: `/pdfjs/${d}`, from: nm(`pdfjs-dist/${d}`) })),
  { url: '/tesseract/worker.min.js', from: nm('tesseract.js/dist/worker.min.js') },
  ...['tesseract-core-lstm.wasm.js', 'tesseract-core-simd-lstm.wasm.js', 'tesseract-core-relaxedsimd-lstm.wasm.js'].map((f) => ({ url: `/tesseract/core/${f}`, from: nm(`tesseract.js-core/${f}`) })),
  ...['rus', 'eng'].map((l) => ({ url: `/tesseract/lang/${l}.traineddata.gz`, from: nm(`@tesseract.js-data/${l}/4.0.0_best_int/${l}.traineddata.gz`) })),
];

/** Serve (dev) and copy (build) the STATIC files. */
function staticFiles(): Plugin {
  let outDir = 'dist';
  return {
    name: 'salatic-static-files',
    configResolved(cfg) {
      outDir = cfg.build.outDir;
    },
    configureServer(server) {
      server.middlewares.use((req, res, next) => {
        const path = decodeURIComponent((req.url ?? '').split('?')[0]);
        for (const s of STATIC) {
          if (path !== s.url && !path.startsWith(s.url + '/')) continue;
          const file = join(s.from, path.slice(s.url.length));
          if (!file.startsWith(s.from) || !existsSync(file) || !statSync(file).isFile()) return next();
          if (file.endsWith('.wasm')) res.setHeader('Content-Type', 'application/wasm');
          else if (file.endsWith('.js')) res.setHeader('Content-Type', 'text/javascript');
          createReadStream(file).pipe(res);
          return;
        }
        next();
      });
    },
    closeBundle() {
      for (const s of STATIC) cpSync(s.from, join(outDir, s.url), { recursive: true });
    },
  };
}

export default defineConfig({
  plugins: [staticFiles()],
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
