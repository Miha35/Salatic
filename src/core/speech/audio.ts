/**
 * Audio for speech recognition: WAV decoding (any PCM/float layout), downmixing and resampling to the
 * 16 kHz mono the models take. Browsers decode compressed audio and video themselves (decodeAudioData);
 * this covers WAV everywhere and everything in tests.
 */

export const SAMPLE_RATE = 16000;

export interface PcmAudio {
  rate: number;
  channels: Float32Array[];
}

/** Decode a RIFF/RIFX WAVE file: integer PCM 8–32 bit, IEEE float 32/64, WAVE_FORMAT_EXTENSIBLE. */
export function decodeWav(bytes: Uint8Array): PcmAudio {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const tag = (o: number): string => String.fromCharCode(bytes[o], bytes[o + 1], bytes[o + 2], bytes[o + 3]);
  if (bytes.length < 12 || (tag(0) !== 'RIFF' && tag(0) !== 'RIFX') || tag(8) !== 'WAVE') throw new Error('not a WAV file');
  const le = tag(0) === 'RIFF';
  let pos = 12;
  let format = 0;
  let channels = 0;
  let rate = 0;
  let bits = 0;
  let data: { off: number; len: number } | null = null;
  while (pos + 8 <= bytes.length) {
    const id = tag(pos);
    const size = dv.getUint32(pos + 4, le);
    const body = pos + 8;
    if (id === 'fmt ' && body + 16 <= bytes.length) {
      format = dv.getUint16(body, le);
      channels = dv.getUint16(body + 2, le);
      rate = dv.getUint32(body + 4, le);
      bits = dv.getUint16(body + 14, le);
      // WAVE_FORMAT_EXTENSIBLE: the real format is the first two bytes of the sub-format GUID
      if (format === 0xfffe && size >= 26 && body + 26 <= bytes.length) format = dv.getUint16(body + 24, le);
    } else if (id === 'data') {
      // a streamed WAV may give 0 or 0xFFFFFFFF as the size: the data runs to the end
      const len = size === 0 || size === 0xffffffff || body + size > bytes.length ? bytes.length - body : size;
      data = { off: body, len };
      break;
    }
    pos = body + size + (size & 1);
  }
  if (!data || !channels || !rate) throw new Error('WAV without format or data');
  if (format !== 1 && format !== 3) throw new Error(`unsupported WAV encoding ${format}`);
  const bps = bits / 8;
  if (!Number.isInteger(bps) || bps < 1 || bps > 8) throw new Error(`unsupported WAV sample size ${bits}`);
  const frames = Math.floor(data.len / (bps * channels));
  const out = Array.from({ length: channels }, () => new Float32Array(frames));
  for (let i = 0; i < frames; i++)
    for (let c = 0; c < channels; c++) {
      const o = data.off + (i * channels + c) * bps;
      let v: number;
      if (format === 3) v = bps === 8 ? dv.getFloat64(o, le) : dv.getFloat32(o, le);
      else if (bps === 1) v = (bytes[o] - 128) / 128;
      else if (bps === 2) v = dv.getInt16(o, le) / 32768;
      else if (bps === 3) {
        const b0 = bytes[o + (le ? 0 : 2)];
        const b1 = bytes[o + 1];
        const b2 = bytes[o + (le ? 2 : 0)];
        v = (((b2 << 24) | (b1 << 16) | (b0 << 8)) >> 8) / 8388608;
      } else v = dv.getInt32(o, le) / 2147483648;
      out[c][i] = v;
    }
  return { rate, channels: out };
}

/** Average of the channels. */
export function downmix(channels: Float32Array[]): Float32Array {
  if (channels.length === 1) return channels[0];
  const n = Math.min(...channels.map((c) => c.length));
  const out = new Float32Array(n);
  for (const c of channels) for (let i = 0; i < n; i++) out[i] += c[i];
  const k = 1 / channels.length;
  for (let i = 0; i < n; i++) out[i] *= k;
  return out;
}

/**
 * Band-limited resampling (windowed sinc, Kaiser window): the cut-off sits below the lower Nyquist
 * frequency, so downsampling from 44.1/48 kHz does not fold high frequencies into the speech band.
 */
export function resample(x: Float32Array, from: number, to: number): Float32Array {
  if (from === to) return x;
  const ratio = to / from;
  const n = Math.floor(x.length * ratio);
  const out = new Float32Array(n);
  const cutoff = 0.94 * Math.min(1, ratio); // of the input Nyquist
  const half = 12; // zero crossings each side of the cut-off sinc
  const width = half / cutoff; // in input samples
  // the kernel, tabulated finely once (linear interpolation between entries)
  const RES = 512;
  const size = Math.ceil(width * RES) + 2;
  const table = new Float32Array(size);
  const beta = 8.6;
  const i0 = (v: number): number => {
    let sum = 1;
    let term = 1;
    for (let k = 1; k < 40; k++) {
      term *= (v / (2 * k)) ** 2;
      sum += term;
      if (term < 1e-12 * sum) break;
    }
    return sum;
  };
  const norm = i0(beta);
  for (let k = 0; k < size; k++) {
    const d = k / RES;
    const r = d / width;
    if (r >= 1) break;
    const arg = Math.PI * d * cutoff;
    table[k] = (i0(beta * Math.sqrt(1 - r * r)) / norm) * (arg === 0 ? 1 : Math.sin(arg) / arg);
  }
  const kernel = (d: number): number => {
    const p = Math.abs(d) * RES;
    const k = Math.floor(p);
    if (k + 1 >= size) return 0;
    return table[k] + (table[k + 1] - table[k]) * (p - k);
  };
  for (let j = 0; j < n; j++) {
    const t = j / ratio; // position in the input
    const lo = Math.max(0, Math.ceil(t - width));
    const hi = Math.min(x.length - 1, Math.floor(t + width));
    let acc = 0;
    let wsum = 0;
    for (let i = lo; i <= hi; i++) {
      const w = kernel(i - t);
      acc += x[i] * w;
      wsum += w;
    }
    out[j] = wsum ? acc / wsum : 0;
  }
  return out;
}

/** 16 kHz mono from any decoded audio. */
export function toSpeechRate(a: PcmAudio): Float32Array {
  return resample(downmix(a.channels), a.rate, SAMPLE_RATE);
}

/** A 16-bit mono WAV of 16 kHz samples (exports, tests). */
export function encodeWav(x: Float32Array, rate = SAMPLE_RATE): Uint8Array {
  const out = new Uint8Array(44 + x.length * 2);
  const dv = new DataView(out.buffer);
  const str = (o: number, s: string): void => {
    for (let i = 0; i < 4; i++) out[o + i] = s.charCodeAt(i);
  };
  str(0, 'RIFF');
  dv.setUint32(4, 36 + x.length * 2, true);
  str(8, 'WAVE');
  str(12, 'fmt ');
  dv.setUint32(16, 16, true);
  dv.setUint16(20, 1, true);
  dv.setUint16(22, 1, true);
  dv.setUint32(24, rate, true);
  dv.setUint32(28, rate * 2, true);
  dv.setUint16(32, 2, true);
  dv.setUint16(34, 16, true);
  str(36, 'data');
  dv.setUint32(40, x.length * 2, true);
  for (let i = 0; i < x.length; i++) dv.setInt16(44 + i * 2, Math.max(-32768, Math.min(32767, Math.round(x[i] * 32767))), true);
  return out;
}

/** Peak-normalise quiet recordings (the models expect a usable level); loud ones are left alone. */
export function normalizeLevel(x: Float32Array, target = 0.5): Float32Array {
  let peak = 0;
  for (let i = 0; i < x.length; i++) peak = Math.max(peak, Math.abs(x[i]));
  if (peak === 0 || peak >= target) return x;
  const k = Math.min(target / peak, 30);
  const out = new Float32Array(x.length);
  for (let i = 0; i < x.length; i++) out[i] = x[i] * k;
  return out;
}
