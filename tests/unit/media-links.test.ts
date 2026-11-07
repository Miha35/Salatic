/** Media by link (transcription page): link checks, cloud-drive links, the streamed response. */
import { describe, expect, it } from 'vitest';
import { fetchMedia, looksLikeMedia, MediaLinkError, resolveMediaLink } from '../../src/core/media-links';

const code = (fn: () => unknown): string => {
  try {
    fn();
  } catch (e) {
    return e instanceof MediaLinkError ? e.code : 'other';
  }
  return 'ok';
};

describe('links', () => {
  it('refuses what a server must not fetch', () => {
    for (const bad of ['ftp://example.com/a.mp3', 'javascript:alert(1)', 'not a link', 'https://user:pw@example.com/a.mp3']) expect(code(() => resolveMediaLink(bad)), bad).toBe('bad_url');
    for (const bad of ['http://localhost/a.mp3', 'http://127.0.0.1/a.mp3', 'http://10.1.2.3/a', 'http://192.168.0.5/a', 'http://172.20.0.1/a', 'http://169.254.169.254/latest/meta-data', 'http://[::1]/a', 'http://[fd00::1]/a', 'http://intranet/a.mp3', 'http://nas.local/a.mp3', 'https://example.com:8443/a.mp3'])
      expect(code(() => resolveMediaLink(bad)), bad).toBe('forbidden_host');
  });

  it('explains that video sites do not give their files out', () => {
    for (const v of ['https://www.youtube.com/watch?v=abc', 'https://youtu.be/abc', 'https://vk.com/video-1_2', 'https://vkvideo.ru/video1', 'https://rutube.ru/video/x/', 'https://m.youtube.com/watch?v=1']) expect(code(() => resolveMediaLink(v)), v).toBe('video_site');
  });

  it('turns cloud-drive share links into download links', () => {
    expect(resolveMediaLink('https://disk.yandex.ru/d/AbCdEf123')).toEqual({ via: 'yandex', url: 'https://cloud-api.yandex.net/v1/disk/public/resources/download?public_key=https%3A%2F%2Fdisk.yandex.ru%2Fd%2FAbCdEf123' });
    expect(resolveMediaLink('https://drive.google.com/file/d/1AbCdEfGhIjKlMn/view?usp=sharing')).toEqual({ via: 'gdrive', url: 'https://drive.usercontent.google.com/download?id=1AbCdEfGhIjKlMn&export=download&confirm=t' });
    expect(resolveMediaLink('https://drive.google.com/open?id=1AbCdEfGhIjKlMn').url).toContain('id=1AbCdEfGhIjKlMn');
    expect(code(() => resolveMediaLink('https://drive.google.com/drive/folders'))).toBe('bad_url');
    expect(resolveMediaLink('https://www.dropbox.com/scl/fi/xyz/talk.mp3?rlkey=k&dl=0').url).toBe('https://www.dropbox.com/scl/fi/xyz/talk.mp3?rlkey=k&dl=1');
    expect(resolveMediaLink('https://cdn.example.org/files/talk.mp3')).toEqual({ via: 'direct', url: 'https://cdn.example.org/files/talk.mp3' });
  });

  it('recognises media by its first bytes', () => {
    const b = (s: string, extra = 16): Uint8Array => Uint8Array.from([...s].map((c) => c.charCodeAt(0)).concat(new Array(extra).fill(0)));
    expect(looksLikeMedia(b('\0\0\0\x20ftypisom'))).toBe(true);
    expect(looksLikeMedia(b('OggS'))).toBe(true);
    expect(looksLikeMedia(b('RIFF\0\0\0\0WAVE'))).toBe(true);
    expect(looksLikeMedia(b('ID3\x04'))).toBe(true);
    expect(looksLikeMedia(Uint8Array.of(0x1a, 0x45, 0xdf, 0xa3, 0, 0))).toBe(true);
    expect(looksLikeMedia(b('<!doctype html><html>'))).toBe(false);
    expect(looksLikeMedia(b('%PDF-1.7'))).toBe(false);
  });
});

/** A fake upstream: url → response. */
function fake(routes: Record<string, () => Response>): typeof fetch & { calls: string[] } {
  const calls: string[] = [];
  const f = (async (input: RequestInfo | URL) => {
    const u = String(input);
    calls.push(u);
    const r = routes[u];
    return r ? r() : new Response('not found', { status: 404 });
  }) as typeof fetch & { calls: string[] };
  f.calls = calls;
  return f;
}
const MP3 = (): Uint8Array => Uint8Array.from([0x49, 0x44, 0x33, 4, 0, 0, ...new Array(2000).fill(7)]);

describe('fetching', () => {
  it('streams a media file through with its name and type', async () => {
    const f = fake({ 'https://cdn.example.org/talk.mp3': () => new Response(MP3() as BodyInit, { headers: { 'Content-Type': 'audio/mpeg', 'Content-Length': '2006' } }) });
    const res = await fetchMedia('https://cdn.example.org/talk.mp3', f);
    expect(res.headers.get('Content-Type')).toBe('audio/mpeg');
    expect(res.headers.get('Content-Disposition')).toContain("filename*=UTF-8''talk.mp3");
    expect(new Uint8Array(await res.arrayBuffer())).toEqual(MP3());
  });

  it('asks Yandex.Disk for the download link first', async () => {
    const f = fake({
      'https://cloud-api.yandex.net/v1/disk/public/resources/download?public_key=https%3A%2F%2Fdisk.yandex.ru%2Fd%2Fabc': () => Response.json({ href: 'https://downloader.disk.yandex.ru/x/file' }),
      'https://downloader.disk.yandex.ru/x/file': () => new Response(MP3() as BodyInit, { headers: { 'Content-Type': 'application/octet-stream', 'Content-Disposition': "attachment; filename*=UTF-8''%D0%B8%D0%BD%D1%82%D0%B5%D1%80%D0%B2%D1%8C%D1%8E.mp3" } }),
    });
    const res = await fetchMedia('https://disk.yandex.ru/d/abc', f);
    expect(res.headers.get('Content-Disposition')).toContain(encodeURIComponent('интервью.mp3'));
    expect(f.calls.length).toBe(2);
  });

  it('refuses pages, other files and files over 2 GB', async () => {
    const page = fake({ 'https://example.org/v': () => new Response('<html>', { headers: { 'Content-Type': 'text/html' } }) });
    await expect(fetchMedia('https://example.org/v', page)).rejects.toMatchObject({ code: 'not_media' });
    const pdf = fake({ 'https://example.org/f': () => new Response(new TextEncoder().encode('%PDF-1.7 ' + 'x'.repeat(600)) as BodyInit, { headers: { 'Content-Type': 'application/octet-stream' } }) });
    await expect(fetchMedia('https://example.org/f', pdf)).rejects.toMatchObject({ code: 'not_media' });
    const huge = fake({ 'https://example.org/big.mp4': () => new Response(MP3() as BodyInit, { headers: { 'Content-Type': 'video/mp4', 'Content-Length': String(3 * 1024 ** 3) } }) });
    await expect(fetchMedia('https://example.org/big.mp4', huge)).rejects.toMatchObject({ code: 'too_large', status: 413 });
    const gone = fake({});
    await expect(fetchMedia('https://example.org/none.mp3', gone)).rejects.toMatchObject({ code: 'upstream' });
  });
});
