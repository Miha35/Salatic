/**
 * Audio/video by link for the transcription page: the server fetches the file (browsers cannot read
 * other sites' files) and streams it back. Cloud-drive share links are turned into download links;
 * video-hosting sites are refused (they do not give their files out, and their terms forbid it).
 */

export const MEDIA_MAX_BYTES = 2 * 1024 * 1024 * 1024;

export class MediaLinkError extends Error {
  constructor(
    message: string,
    readonly code: 'bad_url' | 'forbidden_host' | 'video_site' | 'not_media' | 'too_large' | 'upstream',
    readonly status = 400,
  ) {
    super(message);
  }
}

const VIDEO_SITES = [
  'youtube.com',
  'youtu.be',
  'youtube-nocookie.com',
  'vk.com',
  'vk.ru',
  'vkvideo.ru',
  'vk.cc',
  'rutube.ru',
  'tiktok.com',
  'instagram.com',
  'facebook.com',
  'fb.watch',
  'twitter.com',
  'x.com',
  'ok.ru',
  'dzen.ru',
  'twitch.tv',
  'vimeo.com',
  'dailymotion.com',
  'bilibili.com',
  'kinopoisk.ru',
];

const under = (host: string, domain: string): boolean => host === domain || host.endsWith('.' + domain);

/** Hosts a server must not be sent to: loopback, private and link-local addresses, local names. */
function privateHost(host: string): boolean {
  const h = host.replace(/^\[|\]$/g, '').toLowerCase();
  if (h === 'localhost' || h.endsWith('.localhost') || h.endsWith('.local') || h.endsWith('.internal') || h.endsWith('.lan') || !h.includes('.') && !h.includes(':')) return true;
  const v4 = /^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})$/.exec(h);
  if (v4) {
    const [a, b] = [Number(v4[1]), Number(v4[2])];
    return a === 0 || a === 10 || a === 127 || (a === 100 && b >= 64 && b <= 127) || (a === 169 && b === 254) || (a === 172 && b >= 16 && b <= 31) || (a === 192 && b === 168) || a >= 224;
  }
  if (h.includes(':')) return h === '::' || h === '::1' || /^f[cd]/.test(h) || /^fe[89ab]/.test(h) || h.startsWith('::ffff:');
  return false;
}

export interface ResolvedLink {
  /** what to download (for Yandex.Disk: the API call that gives the download link) */
  url: string;
  via: 'direct' | 'yandex' | 'gdrive' | 'dropbox';
}

/** Check a link and turn share links into download links. */
export function resolveMediaLink(input: string): ResolvedLink {
  let u: URL;
  try {
    u = new URL(input.trim());
  } catch {
    throw new MediaLinkError('Это не ссылка. Вставьте адрес, который начинается с https://', 'bad_url');
  }
  if (u.protocol !== 'https:' && u.protocol !== 'http:') throw new MediaLinkError('Поддерживаются только ссылки http(s).', 'bad_url');
  if (u.username || u.password) throw new MediaLinkError('Ссылки с логином и паролем не поддерживаются.', 'bad_url');
  if (u.port && u.port !== '443' && u.port !== '80') throw new MediaLinkError('Ссылки на нестандартные порты не поддерживаются.', 'forbidden_host');
  const host = u.hostname.toLowerCase();
  if (privateHost(host)) throw new MediaLinkError('Ссылки на локальные и внутренние адреса не поддерживаются.', 'forbidden_host');
  if (VIDEO_SITES.some((d) => under(host, d)))
    throw new MediaLinkError('Видеохостинги (YouTube, VK, RuTube и другие) не отдают файлы по ссылке. Скачайте видео на устройство и загрузите файл.', 'video_site');
  if (['disk.yandex.ru', 'disk.yandex.com', 'disk.yandex.kz', 'disk.yandex.by', 'disk.360.yandex.ru', 'yadi.sk'].includes(host))
    return { url: `https://cloud-api.yandex.net/v1/disk/public/resources/download?public_key=${encodeURIComponent(u.href)}`, via: 'yandex' };
  if (host === 'drive.google.com' || host === 'docs.google.com') {
    const id = /\/file\/d\/([\w-]{10,})/.exec(u.pathname)?.[1] ?? u.searchParams.get('id');
    if (!id || !/^[\w-]{10,}$/.test(id)) throw new MediaLinkError('В ссылке Google Диска нет идентификатора файла. Откройте доступ «Все, у кого есть ссылка» и скопируйте ссылку на файл.', 'bad_url');
    return { url: `https://drive.usercontent.google.com/download?id=${id}&export=download&confirm=t`, via: 'gdrive' };
  }
  if (under(host, 'dropbox.com')) {
    u.searchParams.set('dl', '1');
    return { url: u.href, via: 'dropbox' };
  }
  return { url: u.href, via: 'direct' };
}

const MEDIA_TYPE = /^(audio|video)\/|^application\/(ogg|octet-stream|mp4|x-matroska)|^binary\/octet-stream/i;

/** First bytes of common audio/video containers (when the server says nothing useful about the type). */
export function looksLikeMedia(b: Uint8Array): boolean {
  const s = (o: number, t: string): boolean => [...t].every((c, i) => b[o + i] === c.charCodeAt(0));
  return (
    s(4, 'ftyp') || // MP4, MOV, M4A, 3GP
    (b[0] === 0x1a && b[1] === 0x45 && b[2] === 0xdf && b[3] === 0xa3) || // Matroska, WebM
    s(0, 'OggS') ||
    s(0, 'RIFF') || // WAV, AVI
    s(0, 'fLaC') ||
    s(0, 'ID3') ||
    (b[0] === 0xff && (b[1] & 0xe0) === 0xe0) || // MPEG audio, ADTS AAC
    s(0, '#!AMR') ||
    s(0, 'FORM') || // AIFF
    (b[0] === 0x30 && b[1] === 0x26 && b[2] === 0xb2 && b[3] === 0x75) || // ASF (WMA, WMV)
    (b[0] === 0x47 && b[188] === 0x47) // MPEG-TS
  );
}

/** A file name for the download: the server's, else the last part of the link. */
function fileName(res: Response, url: string): string {
  const cd = res.headers.get('Content-Disposition') ?? '';
  const star = /filename\*\s*=\s*UTF-8''([^;]+)/i.exec(cd)?.[1];
  const plain = /filename\s*=\s*"?([^";]+)"?/i.exec(cd)?.[1];
  let name = '';
  try {
    name = star ? decodeURIComponent(star) : (plain ?? decodeURIComponent(new URL(url).pathname.split('/').pop() ?? ''));
  } catch {
    name = plain ?? '';
  }
  return name.replace(/[\\/:*?"<>|\x00-\x1f]/g, '_').slice(0, 200) || 'media';
}

/**
 * Fetch the file of a link and give a response streaming it (the type and size checked, at most
 * MEDIA_MAX_BYTES). `fetcher` is the platform fetch (injected for tests).
 */
export async function fetchMedia(input: string, fetcher: typeof fetch = fetch): Promise<Response> {
  const link = resolveMediaLink(input);
  let target = link.url;
  if (link.via === 'yandex') {
    const api = await fetcher(target, { headers: { Accept: 'application/json' } });
    if (!api.ok) throw new MediaLinkError('Яндекс.Диск не отдал файл: проверьте, что ссылка публичная и файл не удалён.', 'upstream', 502);
    const href = ((await api.json()) as { href?: string }).href;
    if (!href) throw new MediaLinkError('Яндекс.Диск не дал ссылку на скачивание (возможно, это папка, а не файл).', 'upstream', 502);
    target = href;
  }
  const res = await fetcher(target, { redirect: 'follow', headers: { Accept: 'audio/*, video/*, application/octet-stream;q=0.9, */*;q=0.5' } });
  if (!res.ok || !res.body) throw new MediaLinkError(`Сервер по ссылке ответил ошибкой ${res.status}.`, 'upstream', 502);
  const type = (res.headers.get('Content-Type') ?? '').split(';')[0].trim().toLowerCase();
  if (/^text\/html|^application\/(json|xhtml)/.test(type))
    throw new MediaLinkError(
      link.via === 'gdrive' ? 'Google Диск открыл страницу вместо файла: откройте доступ «Все, у кого есть ссылка».' : 'По ссылке открывается страница, а не файл. Нужна прямая ссылка на аудио или видео.',
      'not_media',
      415,
    );
  const length = Number(res.headers.get('Content-Length') ?? NaN);
  if (Number.isFinite(length) && length > MEDIA_MAX_BYTES) throw new MediaLinkError('Файл больше 2 ГБ.', 'too_large', 413);
  // the first bytes decide when the type says nothing; then the rest streams through, counted
  const reader = res.body.getReader();
  const head: Uint8Array[] = [];
  let got = 0;
  while (got < 512) {
    const { value, done } = await reader.read();
    if (done) break;
    head.push(value);
    got += value.length;
  }
  const first = new Uint8Array(got);
  let o = 0;
  for (const c of head) {
    first.set(c, o);
    o += c.length;
  }
  if (!MEDIA_TYPE.test(type) || /octet-stream/.test(type) || !type) {
    if (!looksLikeMedia(first)) {
      await reader.cancel();
      throw new MediaLinkError('По ссылке не аудио и не видео.', 'not_media', 415);
    }
  }
  let total = got;
  const stream = new ReadableStream<Uint8Array>({
    start(ctrl) {
      if (first.length) ctrl.enqueue(first);
    },
    async pull(ctrl) {
      const { value, done } = await reader.read();
      if (done) return ctrl.close();
      total += value.length;
      if (total > MEDIA_MAX_BYTES) {
        await reader.cancel();
        return ctrl.error(new MediaLinkError('Файл больше 2 ГБ.', 'too_large', 413));
      }
      ctrl.enqueue(value);
    },
    cancel() {
      return reader.cancel();
    },
  });
  const name = fileName(res, target);
  const headers: Record<string, string> = {
    'Content-Type': MEDIA_TYPE.test(type) ? type : 'application/octet-stream',
    'Content-Disposition': `attachment; filename="${name.replace(/[^\x20-\x7e]/g, '_')}"; filename*=UTF-8''${encodeURIComponent(name)}`,
    'Cache-Control': 'no-store',
    'X-Content-Type-Options': 'nosniff',
  };
  if (Number.isFinite(length)) headers['Content-Length'] = String(length);
  return new Response(stream, { status: 200, headers });
}
