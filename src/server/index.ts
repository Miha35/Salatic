/**
 * Cloudflare Worker.
 *
 * Static assets (the SPA) are served directly by Workers Static Assets with headers
 * from `public/_headers`; this script only runs for `/api/*` (see wrangler.jsonc).
 *
 * The API converts lightweight document/data formats on the server for automation
 * (curl, scripts). PDF and image work stays in the browser: it needs fonts, canvas
 * and more CPU time than an edge request should spend.
 */
import { detectFormat, FORMATS, targetsFor, type FormatId } from '../core/formats';
import { LITE_FORMATS, readLite, writeLite } from '../core/convert-lite';
import { ConvertError, stripExtension } from '../core/util';
import { fetchMedia, MediaLinkError } from '../core/media-links';

export interface Env {
  ASSETS: Fetcher;
}

const MAX_BYTES = 10 * 1024 * 1024;
const VERSION = '1.0.0';

const SECURITY_HEADERS: Record<string, string> = {
  'X-Content-Type-Options': 'nosniff',
  'Referrer-Policy': 'no-referrer',
  'Cache-Control': 'no-store',
};

const CORS_HEADERS: Record<string, string> = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
  'Access-Control-Expose-Headers': 'Content-Disposition, X-Warnings, X-Source-Format, X-Target-Format',
  'Access-Control-Max-Age': '86400',
};

function json(data: unknown, status = 200, extra: Record<string, string> = {}): Response {
  return new Response(JSON.stringify(data, null, 2), {
    status,
    headers: { 'Content-Type': 'application/json; charset=utf-8', ...SECURITY_HEADERS, ...CORS_HEADERS, ...extra },
  });
}

const error = (status: number, code: string, message: string): Response => json({ error: message, code }, status);

/** RFC 5987 filename for Content-Disposition (keeps Cyrillic names). */
function contentDisposition(name: string): string {
  const ascii = name.replace(/[^\x20-\x7e]/g, '_').replace(/["\\]/g, '_');
  return `attachment; filename="${ascii}"; filename*=UTF-8''${encodeURIComponent(name)}`;
}

function serverTargets(from: FormatId): FormatId[] {
  return LITE_FORMATS.includes(from) ? targetsFor(from).filter((t) => LITE_FORMATS.includes(t)) : [];
}

function formatsPayload(): unknown {
  const all = Object.values(FORMATS).map(({ id, label, ext, mime, kind }) => ({ id, label, ext, mime, kind }));
  const browser: Record<string, FormatId[]> = {};
  const server: Record<string, FormatId[]> = {};
  for (const f of Object.keys(FORMATS) as FormatId[]) {
    const t = targetsFor(f, { avif: true });
    if (t.length) browser[f] = t;
    const s = serverTargets(f);
    if (s.length) server[f] = s;
  }
  return {
    version: VERSION,
    formats: all,
    conversions: { browser, server },
    usage: {
      convert: 'POST /api/convert?to=<format>[&from=<format>][&name=<file name>] with the file as the raw request body or as multipart/form-data field "file". Max 10 MB.',
      example: "curl -X POST --data-binary @report.docx 'https://<host>/api/convert?to=md&name=report.docx' -o report.md",
    },
  };
}

async function readUpload(request: Request, url: URL): Promise<{ bytes: Uint8Array; name: string } | Response> {
  const declared = Number(request.headers.get('Content-Length') ?? '0');
  if (declared > MAX_BYTES) return error(413, 'too_large', `Файл больше ${MAX_BYTES / 1024 / 1024} МБ. Используйте конвертер в браузере.`);
  const type = request.headers.get('Content-Type') ?? '';
  let bytes: Uint8Array;
  let name = url.searchParams.get('name') ?? 'file';
  if (type.startsWith('multipart/form-data')) {
    let form: FormData;
    try {
      form = await request.formData();
    } catch {
      return error(400, 'bad_form', 'Не удалось разобрать multipart/form-data.');
    }
    const file = form.get('file');
    if (!file || typeof file === 'string') return error(400, 'no_file', 'Поле "file" с файлом не найдено.');
    bytes = new Uint8Array(await file.arrayBuffer());
    name = url.searchParams.get('name') ?? file.name ?? name;
  } else {
    bytes = new Uint8Array(await request.arrayBuffer());
  }
  if (!bytes.length) return error(400, 'empty', 'Пустой файл.');
  if (bytes.length > MAX_BYTES) return error(413, 'too_large', `Файл больше ${MAX_BYTES / 1024 / 1024} МБ. Используйте конвертер в браузере.`);
  return { bytes, name };
}

async function handleConvert(request: Request, url: URL): Promise<Response> {
  const to = url.searchParams.get('to') as FormatId | null;
  if (!to || !FORMATS[to]) return error(400, 'bad_target', 'Укажите целевой формат: ?to=docx|md|html|txt|csv|tsv|json|xlsx');
  const upload = await readUpload(request, url);
  if (upload instanceof Response) return upload;
  let from = url.searchParams.get('from') as FormatId | null;
  if (from && !FORMATS[from]) return error(400, 'bad_source', `Неизвестный исходный формат: ${from}`);
  if (!from) {
    const det = detectFormat(upload.name, upload.bytes);
    if (!det.format) return error(415, 'unknown_format', det.error ?? 'Не удалось определить формат файла.');
    from = det.format;
  }
  const allowed = serverTargets(from);
  if (!allowed.includes(to)) {
    const inBrowser = targetsFor(from, { avif: true }).includes(to);
    return error(
      422,
      'unsupported',
      inBrowser
        ? `Преобразование ${FORMATS[from].label} → ${FORMATS[to].label} выполняется только в браузере (нужны шрифты/рендеринг). Откройте сайт.`
        : `Преобразование ${FORMATS[from].label} → ${FORMATS[to].label} не поддерживается.`,
    );
  }
  try {
    const title = stripExtension(upload.name);
    const { content, warnings: w1 } = readLite(upload.bytes, from, title);
    const out = writeLite(content, to, {}, title);
    const warnings = [...new Set([...w1, ...out.warnings])];
    const filename = `${title || 'file'}.${out.ext}`;
    return new Response(out.data, {
      status: 200,
      headers: {
        'Content-Type': out.mime,
        'Content-Disposition': contentDisposition(filename),
        'X-Source-Format': from,
        'X-Target-Format': to,
        ...(warnings.length ? { 'X-Warnings': encodeURIComponent(JSON.stringify(warnings)) } : {}),
        ...SECURITY_HEADERS,
        ...CORS_HEADERS,
      },
    });
  } catch (e) {
    if (e instanceof ConvertError) return error(422, e.code, e.message);
    console.error('convert failed', e);
    return error(500, 'internal', 'Внутренняя ошибка при преобразовании.');
  }
}

export async function handleApi(request: Request, url: URL): Promise<Response> {
  const path = url.pathname.replace(/\/+$/, '');
  if (request.method === 'OPTIONS') return new Response(null, { status: 204, headers: CORS_HEADERS });
  if (path === '/api/health') {
    if (request.method !== 'GET' && request.method !== 'HEAD') return error(405, 'method', 'Метод не поддерживается.');
    return json({ status: 'ok', version: VERSION });
  }
  if (path === '/api/formats') {
    if (request.method !== 'GET' && request.method !== 'HEAD') return error(405, 'method', 'Метод не поддерживается.');
    return json(formatsPayload(), 200, { 'Cache-Control': 'public, max-age=3600' });
  }
  if (path === '/api/media') {
    // audio/video by link for the transcription page: only for pages of this site (not an open proxy)
    if (request.method !== 'GET') return error(405, 'method', 'Используйте GET.');
    const site = request.headers.get('Sec-Fetch-Site');
    if (site !== 'same-origin') return error(403, 'forbidden', 'Этот метод доступен только со страниц сайта.');
    try {
      return await fetchMedia(url.searchParams.get('url') ?? '');
    } catch (e) {
      if (e instanceof MediaLinkError) return error(e.status, e.code, e.message);
      return error(502, 'upstream', 'Не удалось скачать файл по ссылке.');
    }
  }
  if (path === '/api/convert') {
    if (request.method !== 'POST') return error(405, 'method', 'Используйте POST.');
    return handleConvert(request, url);
  }
  return error(404, 'not_found', 'Неизвестный метод API.');
}

export default {
  async fetch(request: Request, env: Env): Promise<Response> {
    const url = new URL(request.url);
    if (url.pathname === '/api' || url.pathname.startsWith('/api/')) {
      try {
        return await handleApi(request, url);
      } catch (e) {
        console.error('api failure', e);
        return error(500, 'internal', 'Внутренняя ошибка сервера.');
      }
    }
    return env.ASSETS.fetch(request);
  },
} satisfies ExportedHandler<Env>;
