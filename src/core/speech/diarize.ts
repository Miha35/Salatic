/**
 * Speaker diarization after pyannote 3.1, with the same models (segmentation-3.0, WeSpeaker ResNet34):
 *
 * 1. the segmentation model sees 10 s windows sliding over the audio and says, per 16.9 ms frame, which
 *    of up to three local speakers talk (powerset classes, decoded to one-hot activities);
 * 2. each local speaker of each window gets a voice embedding (from the frames where it talks alone);
 * 3. the embeddings are clustered (agglomerative, centroid linkage, distance threshold or a given count);
 * 4. per frame, the number of simultaneous speakers is estimated from all windows covering it, and that
 *    many global speakers with the highest aggregated activity are kept.
 *
 * Pure computation: the models are run by the caller (browser worker, Node tests).
 */

export const FRAME_STEP = 270; // samples per segmentation frame (16 kHz)
export const FRAME_SPAN = 990; // samples one frame covers
export const WINDOW = 160000; // 10 s
export const WINDOW_FRAMES = Math.floor((WINDOW - FRAME_SPAN) / FRAME_STEP) + 1; // 589

/** pyannote segmentation-3.0 powerset classes → local speakers (none, 1, 2, 3, 1+2, 1+3, 2+3). */
const POWERSET: number[][] = [[], [0], [1], [2], [0, 1], [0, 2], [1, 2]];

/** One window: its first sample and, per local speaker, a 0/1 activity per frame. */
export interface SegWindow {
  start: number;
  active: Uint8Array[]; // 3 × frames
}

/** Window starts covering `n` samples with the given step (a whole number of frames), the last one flush with the end. */
export function windowStarts(n: number, stepFrames: number): number[] {
  const step = stepFrames * FRAME_STEP;
  if (n <= WINDOW) return [0];
  const starts: number[] = [];
  for (let s = 0; s + WINDOW <= n; s += step) starts.push(s);
  const last = n - WINDOW;
  // the tail: one more window reaching the end, on the frame grid (it may run a few samples past it: zeros)
  if (starts[starts.length - 1] < last) starts.push(Math.ceil(last / FRAME_STEP) * FRAME_STEP);
  return starts;
}

/** Decode powerset logits (frames × 7) into one-hot activities of the three local speakers. */
export function decodePowerset(logits: Float32Array, frames: number): Uint8Array[] {
  const act = [new Uint8Array(frames), new Uint8Array(frames), new Uint8Array(frames)];
  for (let f = 0; f < frames; f++) {
    let best = 0;
    for (let c = 1; c < 7; c++) if (logits[f * 7 + c] > logits[f * 7 + best]) best = c;
    for (const s of POWERSET[best]) act[s][f] = 1;
  }
  return act;
}

/**
 * The samples a local speaker talks alone in (or, when that is too little, all its frames): the input
 * of its embedding. Null when it says too little to be recognised.
 */
export function speakerAudio(audio: Float32Array, w: SegWindow, s: number, minSeconds = 0.6): Float32Array | null {
  const frames = w.active[s].length;
  const pick = (alone: boolean): Float32Array | null => {
    const keep = new Uint8Array(WINDOW);
    let n = 0;
    for (let f = 0; f < frames; f++) {
      if (!w.active[s][f]) continue;
      if (alone && w.active.some((a, o) => o !== s && a[f])) continue;
      // the frame's own stretch of samples (step), centred in its span
      const a = f * FRAME_STEP + (FRAME_SPAN - FRAME_STEP) / 2;
      for (let i = Math.max(0, Math.floor(a)); i < Math.min(WINDOW, Math.floor(a) + FRAME_STEP); i++) {
        if (!keep[i]) n++;
        keep[i] = 1;
      }
    }
    if (n < minSeconds * 16000) return null;
    const out = new Float32Array(n);
    let k = 0;
    for (let i = 0; i < WINDOW; i++) if (keep[i]) out[k++] = audio[w.start + i] ?? 0;
    return out;
  };
  return pick(true) ?? pick(false);
}

/**
 * Agglomerative clustering, centroid linkage on L2-normalised embeddings (Euclidean distance), as pyannote:
 * merging stops at `threshold`, or at `count` clusters when the number of speakers is known.
 * Small clusters (fewer members than `minSize`) are then attached to the nearest large one.
 */
export function cluster(embeddings: Float32Array[], opts: { threshold?: number; count?: number; minCount?: number; maxCount?: number; minSize?: number } = {}): number[] {
  const n = embeddings.length;
  if (n === 0) return [];
  if (n === 1) return [0];
  const dim = embeddings[0].length;
  const unit = embeddings.map((e) => {
    let s = 0;
    for (let i = 0; i < dim; i++) s += e[i] * e[i];
    const k = 1 / (Math.sqrt(s) || 1);
    return Float64Array.from(e, (v) => v * k);
  });
  // cluster centroids and sizes; squared distances between live clusters
  const cent = unit.map((u) => Float64Array.from(u));
  const size = new Array<number>(n).fill(1);
  const alive = new Array<boolean>(n).fill(true);
  const members: number[][] = Array.from({ length: n }, (_, i) => [i]);
  const d2 = (a: Float64Array, b: Float64Array): number => {
    let s = 0;
    for (let i = 0; i < dim; i++) {
      const t = a[i] - b[i];
      s += t * t;
    }
    return s;
  };
  // nearest neighbour of each cluster (recomputed for rows that change)
  const nn = new Int32Array(n).fill(-1);
  const nnd = new Float64Array(n).fill(Infinity);
  const recompute = (i: number): void => {
    nn[i] = -1;
    nnd[i] = Infinity;
    for (let j = 0; j < n; j++) {
      if (j === i || !alive[j]) continue;
      const v = d2(cent[i], cent[j]);
      if (v < nnd[i]) {
        nnd[i] = v;
        nn[i] = j;
      }
    }
  };
  for (let i = 0; i < n; i++) recompute(i);
  let clusters = n;
  const thr2 = opts.threshold !== undefined ? opts.threshold * opts.threshold : Infinity;
  const target = opts.count ?? opts.minCount ?? 1;
  const ceiling = opts.count ?? opts.maxCount ?? Infinity;
  while (clusters > 1) {
    let a = -1;
    for (let i = 0; i < n; i++) if (alive[i] && nn[i] >= 0 && (a < 0 || nnd[i] < nnd[a])) a = i;
    if (a < 0) break;
    const best = nnd[a];
    // stop at the threshold, unless more clusters remain than allowed; never below the wanted count
    if (clusters <= target) break;
    if (best > thr2 && clusters <= ceiling && opts.count === undefined) break;
    const b = nn[a];
    // merge b into a: size-weighted centroid
    const na = size[a];
    const nb = size[b];
    for (let k = 0; k < dim; k++) cent[a][k] = (cent[a][k] * na + cent[b][k] * nb) / (na + nb);
    size[a] = na + nb;
    members[a].push(...members[b]);
    members[b] = [];
    alive[b] = false;
    clusters--;
    recompute(a);
    for (let i = 0; i < n; i++) {
      if (!alive[i] || i === a) continue;
      if (nn[i] === a || nn[i] === b) recompute(i);
      else {
        const v = d2(cent[i], cent[a]);
        if (v < nnd[i]) {
          nnd[i] = v;
          nn[i] = a;
        }
      }
    }
  }
  let groups = members.filter((m, i) => alive[i] && m.length);
  // small clusters join the nearest large one (unless the count is given and they are needed)
  const minSize = Math.min(opts.minSize ?? 12, Math.max(1, Math.round(0.1 * n)));
  const large = groups.filter((g) => g.length >= minSize);
  if (large.length && large.length < groups.length && (opts.count === undefined || large.length >= opts.count)) {
    const centroid = (g: number[]): Float64Array => {
      const c = new Float64Array(dim);
      for (const i of g) for (let k = 0; k < dim; k++) c[k] += unit[i][k] / g.length;
      return c;
    };
    const lc = large.map(centroid);
    for (const g of groups.filter((x) => x.length < minSize)) {
      const c = centroid(g);
      let best = 0;
      for (let k = 1; k < lc.length; k++) if (d2(c, lc[k]) < d2(c, lc[best])) best = k;
      large[best].push(...g);
    }
    groups = large;
  }
  // labels by first appearance
  groups.sort((x, y) => Math.min(...x) - Math.min(...y));
  const labels = new Array<number>(n).fill(0);
  groups.forEach((g, k) => {
    for (const i of g) labels[i] = k;
  });
  return labels;
}

export interface SpeakerTurn {
  speaker: number;
  start: number; // seconds
  end: number;
}

/**
 * Global diarization: local speakers mapped to clusters (`labels[w][s]`, -1 when it had no embedding),
 * activities aggregated over windows, and per frame the estimated number of speakers kept.
 */
export function reconstruct(windows: SegWindow[], labels: number[][], totalSamples: number, minDuration = 0.25): SpeakerTurn[] {
  const frames = Math.max(1, Math.ceil(totalSamples / FRAME_STEP));
  const speakers = Math.max(0, ...labels.flat()) + 1;
  const act = Array.from({ length: speakers }, () => new Float32Array(frames));
  const cover = new Float32Array(frames);
  const count = new Float32Array(frames);
  windows.forEach((w, wi) => {
    const off = Math.round(w.start / FRAME_STEP);
    const nf = w.active[0].length;
    for (let f = 0; f < nf && off + f < frames; f++) {
      const g = off + f;
      cover[g] += 1;
      let c = 0;
      for (let s = 0; s < w.active.length; s++) {
        if (!w.active[s][f]) continue;
        c++;
        const k = labels[wi]?.[s] ?? -1;
        // a local speaker too brief for an embedding still counts as a speaker: it is given the
        // window's speaker it overlaps least with below (via the count), not an activity of its own
        if (k >= 0) act[k][g] = Math.max(act[k][g], 0) + 1;
      }
      count[g] += c;
    }
  });
  // per frame: the k most active speakers, k = rounded average count
  const on = Array.from({ length: speakers }, () => new Uint8Array(frames));
  for (let g = 0; g < frames; g++) {
    if (!cover[g]) continue;
    const k = Math.round(count[g] / cover[g]);
    if (k <= 0) continue;
    const order = Array.from({ length: speakers }, (_, s) => s)
      .filter((s) => act[s][g] > 0)
      .sort((a, b) => act[b][g] - act[a][g]);
    for (const s of order.slice(0, k)) on[s][g] = 1;
  }
  // frames → turns (frame g covers its step, centred in its span)
  const t = (g: number): number => (g * FRAME_STEP + (FRAME_SPAN - FRAME_STEP) / 2) / 16000;
  const turns: SpeakerTurn[] = [];
  for (let s = 0; s < speakers; s++) {
    let g = 0;
    while (g < frames) {
      if (!on[s][g]) {
        g++;
        continue;
      }
      const a = g;
      while (g < frames && on[s][g]) g++;
      turns.push({ speaker: s, start: t(a), end: Math.min(totalSamples / 16000, t(g)) });
    }
  }
  // fill tiny gaps of one speaker, drop specks
  turns.sort((x, y) => x.start - y.start || x.speaker - y.speaker);
  const merged: SpeakerTurn[] = [];
  for (const tr of turns) {
    const prev = [...merged].reverse().find((m) => m.speaker === tr.speaker);
    if (prev && tr.start - prev.end < 0.1) prev.end = Math.max(prev.end, tr.end);
    else merged.push({ ...tr });
  }
  return merged.filter((tr) => tr.end - tr.start >= minDuration).sort((x, y) => x.start - y.start);
}

/** Speech regions (any speaker) from the windows: where Whisper should listen. */
export function speechRegions(windows: SegWindow[], totalSamples: number, padding = 0.2, joinGap = 0.6): { start: number; end: number }[] {
  const frames = Math.max(1, Math.ceil(totalSamples / FRAME_STEP));
  const votes = new Float32Array(frames);
  const cover = new Float32Array(frames);
  for (const w of windows) {
    const off = Math.round(w.start / FRAME_STEP);
    for (let f = 0; f < w.active[0].length && off + f < frames; f++) {
      cover[off + f]++;
      if (w.active.some((a) => a[f])) votes[off + f]++;
    }
  }
  const t = (g: number): number => (g * FRAME_STEP + (FRAME_SPAN - FRAME_STEP) / 2) / 16000;
  const out: { start: number; end: number }[] = [];
  let g = 0;
  while (g < frames) {
    if (!(cover[g] && votes[g] / cover[g] >= 0.5)) {
      g++;
      continue;
    }
    const a = g;
    while (g < frames && cover[g] && votes[g] / cover[g] >= 0.5) g++;
    const start = Math.max(0, t(a) - padding);
    const end = Math.min(totalSamples / 16000, t(g) + padding);
    const last = out[out.length - 1];
    if (last && start - last.end < joinGap) last.end = end;
    else out.push({ start, end });
  }
  return out;
}

/**
 * Pieces of at most `max` seconds for Whisper, cut in the pauses between speech regions; a region longer
 * than that is cut at its quietest point between `min` and `max` seconds.
 */
export function asrChunks(audio: Float32Array, regions: { start: number; end: number }[], max = 28, min = 12): { start: number; end: number }[] {
  const chunks: { start: number; end: number }[] = [];
  let cur: { start: number; end: number } | null = null;
  const energyMin = (from: number, to: number): number => {
    // the 100 ms window with the least energy in [from, to] (seconds) → its centre
    const win = 1600;
    let best = from;
    let bestE = Infinity;
    for (let s = Math.floor(from * 16000); s + win <= Math.min(audio.length, to * 16000); s += 400) {
      let e = 0;
      for (let i = s; i < s + win; i++) e += audio[i] * audio[i];
      if (e < bestE) {
        bestE = e;
        best = (s + win / 2) / 16000;
      }
    }
    return best;
  };
  const push = (c: { start: number; end: number }): void => {
    // split what is still too long at its quietest points
    let s = c.start;
    while (c.end - s > max) {
      const cut = energyMin(s + min, s + max);
      chunks.push({ start: s, end: cut });
      s = cut;
    }
    chunks.push({ start: s, end: c.end });
  };
  for (const r of regions) {
    if (cur && r.end - cur.start <= max) cur.end = r.end;
    else {
      if (cur) push(cur);
      cur = { ...r };
    }
  }
  if (cur) push(cur);
  return chunks;
}
