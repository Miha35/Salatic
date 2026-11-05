/**
 * Builds a real HEIC still for tests: ffmpeg/libx265 encodes one intra frame,
 * this module wraps the HEVC NAL units into a minimal HEIF container
 * (ftyp, meta with hdlr/pitm/iloc/iinf/iprp, mdat).
 */
import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { withTemp } from './helpers';

const u8 = (...v: number[]): number[] => v;
const u16 = (v: number): number[] => [(v >> 8) & 255, v & 255];
const u32 = (v: number): number[] => [(v >>> 24) & 255, (v >>> 16) & 255, (v >>> 8) & 255, v & 255];
const str = (s: string): number[] => [...s].map((c) => c.charCodeAt(0));
const box = (type: string, ...payload: number[][]): number[] => {
  const body = payload.flat();
  return [...u32(8 + body.length), ...str(type), ...body];
};
const full = (type: string, version: number, flags: number, ...payload: number[][]): number[] => box(type, [version, ...u32(flags).slice(1)], ...payload);

/** Split an Annex-B byte stream into NAL units. */
function nals(stream: Uint8Array): Uint8Array[] {
  const out: Uint8Array[] = [];
  let start = -1;
  for (let i = 0; i + 3 <= stream.length; i++) {
    const sc3 = stream[i] === 0 && stream[i + 1] === 0 && stream[i + 2] === 1;
    if (!sc3) continue;
    if (start >= 0) {
      let end = i;
      while (end > start && stream[end - 1] === 0) end--;
      out.push(stream.subarray(start, end));
    }
    start = i + 3;
    i += 2;
  }
  if (start >= 0) out.push(stream.subarray(start));
  return out;
}

/** HEVC-coded HEIF image of ffmpeg's `testsrc2` pattern (also returns the source PNG for comparison). */
export function makeHeic(width = 64, height = 48): { heic: Uint8Array; png: Uint8Array } {
  return withTemp((dir) => {
    const src = join(dir, 'src.png');
    const hevc = join(dir, 'out.hevc');
    execFileSync('ffmpeg', ['-v', 'error', '-f', 'lavfi', '-i', `testsrc2=size=${width}x${height}:rate=1`, '-frames:v', '1', src]);
    execFileSync('ffmpeg', ['-v', 'error', '-i', src, '-frames:v', '1', '-c:v', 'libx265', '-pix_fmt', 'yuv420p', '-x265-params', 'log-level=error:keyint=1:lossless=0:qp=4', '-f', 'hevc', hevc]);
    const units = nals(new Uint8Array(readFileSync(hevc)));
    const type = (n: Uint8Array): number => (n[0] >> 1) & 0x3f;
    const params = [32, 33, 34].map((t) => units.filter((n) => type(n) === t));
    const slices = units.filter((n) => type(n) < 32);
    const arrays = params.map((list, i) => [0x80 | [32, 33, 34][i], ...u16(list.length), ...list.flatMap((n) => [...u16(n.length), ...n])]);
    // hvcC: version 1, Main profile, level 3.0, 4:2:0 8-bit, 4-byte NAL lengths
    const hvcC = box('hvcC', [1, 0x01, ...u32(0x60000000), ...[0x90, 0, 0, 0, 0, 0], 90, ...u16(0xf000), 0xfc, 0xfd, 0xf8, 0xf8, ...u16(0), 0x0f, params.length], ...arrays);
    const ispe = full('ispe', 0, 0, u32(width), u32(height));
    const data = slices.flatMap((n) => [...u32(n.length), ...n]);
    const ftyp = box('ftyp', str('heic'), u32(0), str('mif1'), str('heic'));
    const metaFor = (offset: number): number[] =>
      full(
        'meta',
        0,
        0,
        full('hdlr', 0, 0, u32(0), str('pict'), u32(0), u32(0), u32(0), [0]),
        full('pitm', 0, 0, u16(1)),
        full('iloc', 0, 0, u8(0x44, 0x00), u16(1), u16(1), u16(0), u16(1), u32(offset), u32(data.length)),
        full('iinf', 0, 0, u16(1), full('infe', 2, 0, u16(1), u16(0), str('hvc1'), [0])),
        box('iprp', box('ipco', hvcC, ispe), full('ipma', 0, 0, u32(1), u16(1), [2, 0x81, 0x02])),
      );
    // the mdat payload follows ftyp + meta + the mdat header
    const offset = ftyp.length + metaFor(0).length + 8;
    const heic = Uint8Array.from([...ftyp, ...metaFor(offset), ...box('mdat', data)]);
    return { heic, png: new Uint8Array(readFileSync(src)) };
  });
}
