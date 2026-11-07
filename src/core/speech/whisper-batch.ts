/**
 * Whisper on several pieces at once (a GPU decodes a batch in about the time of one piece), with word
 * timestamps. transformers.js decodes word timestamps for one piece at a time only, so the batch is
 * generated with the base generation loop and the timestamps are worked out here the same way: the
 * alignment heads' cross-attention, normalised and median-filtered, aligned to time by DTW.
 */
import type { Word } from './transcript';

type TF = typeof import('@huggingface/transformers');

/** A tensor as ONNX Runtime returns it: data and dims. */
export interface Tensorish {
  data: Float32Array | ArrayLike<number>;
  dims: number[];
}

function medianFilter(data: Float32Array, width: number): Float32Array {
  const out = new Float32Array(data.length);
  const buf = new Float32Array(width);
  const half = Math.floor(width / 2);
  for (let i = 0; i < data.length; i++) {
    for (let j = -half; j <= half; j++) {
      let k = i + j;
      if (k < 0) k = -k;
      else if (k >= data.length) k = 2 * (data.length - 1) - k;
      buf[j + half] = data[Math.max(0, Math.min(data.length - 1, k))];
    }
    buf.sort();
    out[i] = buf[half];
  }
  return out;
}

/** DTW over a cost matrix (rows: tokens, columns: frames): the aligned (token, frame) path. */
export function dtw(cost: Float32Array[], cols: number): [number[], number[]] {
  const rows = cost.length;
  const W = cols + 1;
  const acc = new Float64Array((rows + 1) * W).fill(Infinity);
  const trace = new Int8Array((rows + 1) * W).fill(-1);
  acc[0] = 0;
  for (let j = 1; j <= cols; j++)
    for (let i = 1; i <= rows; i++) {
      const c0 = acc[(i - 1) * W + j - 1];
      const c1 = acc[(i - 1) * W + j];
      const c2 = acc[i * W + j - 1];
      let c: number;
      let t: number;
      if (c0 < c1 && c0 < c2) {
        c = c0;
        t = 0;
      } else if (c1 < c0 && c1 < c2) {
        c = c1;
        t = 1;
      } else {
        c = c2;
        t = 2;
      }
      acc[i * W + j] = cost[i - 1][j - 1] + c;
      trace[i * W + j] = t;
    }
  for (let j = 0; j <= cols; j++) trace[j] = 2;
  for (let i = 0; i <= rows; i++) trace[i * W] = 1;
  let i = rows;
  let j = cols;
  const ti: number[] = [];
  const fi: number[] = [];
  while (i > 0 || j > 0) {
    ti.push(i - 1);
    fi.push(j - 1);
    const t = trace[i * W + j];
    if (t === 0) {
      i--;
      j--;
    } else if (t === 1) i--;
    else if (t === 2) j--;
    else break;
  }
  return [ti.reverse(), fi.reverse()];
}

/**
 * Timestamps (seconds) of the tokens of batch item `b`: `cross[step][layer]` are the cross-attentions of
 * each generation step ([batch, heads, tokens of the step, encoder frames]); `rows` how many attention
 * rows belong to the item (its tokens before its end); `frames` the encoder frames that hold audio.
 * Mirrors transformers.js' `_extract_token_timestamps` for one item.
 */
export function tokenTimestamps(cross: Tensorish[][], heads: [number, number][], b: number, rows: number, frames: number, numInputIds: number, medianWidth = 7, precision = 0.02): number[] {
  // per alignment head: the attention matrix rows × frames of this item
  const mats: Float32Array[][] = heads.map(([l, h]) => {
    const out: Float32Array[] = [];
    for (const step of cross) {
      const t = step[l];
      const [, H, T, F] = t.dims;
      for (let r = 0; r < T && out.length < rows; r++) {
        const off = ((b * H + h) * T + r) * F;
        out.push(Float32Array.from({ length: frames }, (_, f) => Number(t.data[off + f])));
      }
    }
    return out;
  });
  // normalise each head per frame over the tokens (population std), median-filter along time
  for (const m of mats) {
    const n = m.length;
    for (let f = 0; f < frames; f++) {
      let mean = 0;
      for (let r = 0; r < n; r++) mean += m[r][f];
      mean /= n;
      let v = 0;
      for (let r = 0; r < n; r++) v += (m[r][f] - mean) ** 2;
      const std = Math.sqrt(v / n);
      for (let r = 0; r < n; r++) m[r][f] = (m[r][f] - mean) / std;
    }
    for (let r = 0; r < n; r++) m[r] = medianFilter(m[r], medianWidth);
  }
  // mean over heads, the prompt rows dropped; DTW on the negated weights
  const cost: Float32Array[] = [];
  for (let r = numInputIds; r < rows; r++) {
    const row = new Float32Array(frames);
    for (const m of mats) for (let f = 0; f < frames; f++) row[f] -= m[r][f] / mats.length;
    cost.push(row);
  }
  const out = new Array<number>(numInputIds).fill(0);
  if (!cost.length) return out;
  const [ti, fi] = dtw(cost, frames);
  for (let k = 0; k < ti.length; k++) if (k === 0 || ti[k] !== ti[k - 1]) out.push(fi[k] * precision);
  out.push(out[out.length - 1]);
  return out;
}

interface WhisperParts {
  model: {
    generation_config: Record<string, unknown> & { alignment_heads?: [number, number][]; eos_token_id: number | number[]; no_timestamps_token_id: number; begin_suppress_tokens?: number[] };
    config: { max_target_positions?: number };
  };
  processor: (audio: Float32Array) => Promise<{ input_features: { data: Float32Array; dims: number[] } }>;
  tokenizer: { _decode_asr(chunks: unknown[], o: Record<string, unknown>): [string, { chunks?: { text: string; timestamp: [number, number | null] }[] }] };
}

/**
 * Words of each piece (≤ 30 s each), decoded together. `init`: the decoder prompt, shared by all or one
 * per piece (equally long, ending in the same <|startoftranscript|> <|lang|> <|transcribe|>).
 */
export async function transcribeBatchWith(tf: TF, asr: unknown, pieces: Float32Array[], init: number[] | number[][], alignOver?: (samples: number) => number): Promise<Word[][]> {
  const { model, processor, tokenizer } = asr as WhisperParts;
  const gc = model.generation_config;
  const B = pieces.length;
  const prompts = Array.isArray(init[0]) ? (init as number[][]) : pieces.map(() => init as number[]);
  const P = prompts[0].length;
  if (prompts.length !== B || prompts.some((p) => p.length !== P || p[P - 1] !== prompts[0][P - 1])) throw new Error('batched prompts differ in length');
  // mel features, all padded to 30 s by the processor; stacked into one batch
  const feats = await Promise.all(pieces.map((p) => processor(p)));
  const [, mels, len] = feats[0].input_features.dims;
  const data = new Float32Array(B * mels * len);
  feats.forEach((f, i) => data.set(f.input_features.data, i * mels * len));
  const inputs = new tf.Tensor('float32', data, [B, mels, len]);
  const procs = new tf.LogitsProcessorList();
  // both processors only use the prompt's length and its last token
  procs.push(new tf.WhisperTimeStampLogitsProcessor(gc as never, prompts[0]));
  if (gc.begin_suppress_tokens) procs.push(new tf.SuppressTokensAtBeginLogitsProcessor(gc.begin_suppress_tokens, P));
  const maxTokens = (model.config.max_target_positions ?? 448) - P - 1;
  const generate = (tf.PreTrainedModel.prototype as unknown as { generate(o: Record<string, unknown>): Promise<unknown> }).generate;
  const out = (await generate.call(model, {
    inputs,
    logits_processor: procs,
    decoder_input_ids: prompts,
    max_new_tokens: maxTokens,
    return_timestamps: true,
    output_attentions: true,
    return_dict_in_generate: true,
  })) as { sequences: { tolist(): bigint[][] }; cross_attentions: Tensorish[][] };
  const seqs = out.sequences.tolist();
  const eos = Array.isArray(gc.eos_token_id) ? gc.eos_token_id[0] : gc.eos_token_id;
  const timestampBegin = gc.no_timestamps_token_id + 1;
  const heads = gc.alignment_heads ?? [];
  return seqs.map((seq, b) => {
    // the item ends at its end-of-text token (later positions are padding of the batch)
    let end = seq.length;
    for (let k = P; k < seq.length; k++)
      if (Number(seq[k]) === eos) {
        end = k + 1;
        break;
      }
    // the encoder frames that hold audio (50 per second); `alignOver` lets tests reproduce other choices
    const frames = Math.max(1, Math.min(1500, alignOver ? alignOver(pieces[b].length) : Math.floor(pieces[b].length / 320)));
    const times = tokenTimestamps(out.cross_attentions, heads, b, end - 1, frames, P);
    const prefix = Math.max(0, seq.findIndex((t, k) => k >= P && Number(t) >= timestampBegin));
    const [, opt] = tokenizer._decode_asr([{ tokens: seq.slice(prefix, end), token_timestamps: times.slice(prefix, end).map((x) => Math.round(x * 100) / 100), stride: [pieces[b].length / 16000, 0, 0] }], {
      time_precision: 0.02,
      return_timestamps: 'word',
      force_full_sequences: false,
    });
    const words: Word[] = [];
    for (const c of opt.chunks ?? []) {
      const [s, e] = c.timestamp;
      const start = Number.isFinite(s) ? s : (words[words.length - 1]?.end ?? 0);
      words.push({ text: c.text, start, end: e !== null && Number.isFinite(e) ? e : start + 0.3 });
    }
    return words;
  });
}
