import { cpSync, createReadStream, existsSync, mkdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { gzipSync } from 'node:zlib';
import { defaultClientConditions, defineConfig, type Plugin } from 'vite';

const nm = (p: string): string => resolve(import.meta.dirname, 'node_modules', p);

/** ONNX Runtime for the speech worker, served from this site; its WebGPU build is over 25 MB, so gzipped. */
const ORT_VERSION: string = JSON.parse(readFileSync(nm('onnxruntime-web/package.json'), 'utf8')).version;
const ORT_FILES: { url: string; from: string; gzip?: boolean }[] = [
  { url: `/ort/${ORT_VERSION}/ort-wasm-simd-threaded.asyncify.mjs`, from: nm('onnxruntime-web/dist/ort-wasm-simd-threaded.asyncify.mjs') },
  { url: `/ort/${ORT_VERSION}/ort-wasm-simd-threaded.asyncify.wasm.gz`, from: nm('onnxruntime-web/dist/ort-wasm-simd-threaded.asyncify.wasm'), gzip: true },
];
let gzCache: Buffer | null = null;
const gzipped = (file: string): Buffer => (gzCache ??= gzipSync(readFileSync(file), { level: 9 }));

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
        const ort = ORT_FILES.find((f) => f.url === path);
        if (ort) {
          res.setHeader('Content-Type', ort.gzip ? 'application/gzip' : 'text/javascript');
          if (ort.gzip) res.end(gzipped(ort.from));
          else createReadStream(ort.from).pipe(res);
          return;
        }
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
      for (const f of ORT_FILES) {
        const out = join(outDir, f.url);
        mkdirSync(dirname(out), { recursive: true });
        if (f.gzip) writeFileSync(out, gzipped(f.from));
        else cpSync(f.from, out);
      }
    },
  };
}

export default defineConfig({
  plugins: [staticFiles()],
  define: {
    __ORT_VERSION__: JSON.stringify(ORT_VERSION),
  },
  resolve: {
    // onnxruntime-web without its own reference to the 26 MB runtime (served gzipped from /ort/ instead)
    conditions: ['onnxruntime-web-use-extern-wasm', ...defaultClientConditions],
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
