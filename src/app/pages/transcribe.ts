/**
 * Speech → text with speakers: an audio or video file (or a link to one) is decoded in the page, the
 * speech worker recognises it on the device, and the transcript can be checked against the recording,
 * corrected, its speakers named, and saved as DOCX, TXT, subtitles, Markdown or JSON.
 */
import { decodeWav, downmix, toSpeechRate } from '../../core/speech/audio';
import { clock, toDoc, toJson, toMarkdown, toSrt, toText, toVtt, type Transcript } from '../../core/speech/transcript';
import type { Preset, SpeechRequest, SpeechResponse } from '../engine/speech-worker';
import { combine } from '../../core/speech/pipeline';
import type { SpeakerTurn } from '../../core/speech/diarize';
import type { Word } from '../../core/speech/transcript';
import { baseName, button, download, dropzone, formatBytes, h, handOff, icon, note, prefs, toast } from '../ui';

const ACCEPT = 'audio/*,video/*,.mp3,.wav,.m4a,.aac,.ogg,.oga,.opus,.flac,.webm,.mp4,.m4v,.mov,.mkv,.3gp,.wma,.amr';
const MAX_BYTES = 2 * 1024 * 1024 * 1024;

const LANGUAGES: [string, string][] = [
  ['russian', 'Русский'],
  ['english', 'Английский'],
  ['ukrainian', 'Украинский'],
  ['belarusian', 'Белорусский'],
  ['kazakh', 'Казахский'],
  ['uzbek', 'Узбекский'],
  ['tatar', 'Татарский'],
  ['german', 'Немецкий'],
  ['french', 'Французский'],
  ['spanish', 'Испанский'],
  ['italian', 'Итальянский'],
  ['turkish', 'Турецкий'],
  ['chinese', 'Китайский'],
];

const PRESETS: { id: Preset; label: string; hint: string }[] = [
  { id: 'accurate', label: 'Максимальная точность', hint: 'Whisper large-v3-turbo, ≈1,6 ГБ; нужна видеокарта с WebGPU' },
  { id: 'balanced', label: 'Высокая точность', hint: 'та же модель в 4-битной версии, ≈0,8 ГБ' },
  { id: 'fast', label: 'Быстрая', hint: 'Whisper small, ≈0,3 ГБ; заметно больше ошибок' },
];

/** Speaker colours (enough contrast on both themes). */
const COLORS = ['#2563eb', '#db2777', '#059669', '#d97706', '#7c3aed', '#0891b2', '#dc2626', '#4d7c0f', '#9333ea', '#0f766e'];

interface Settings {
  language: string;
  speakers: number; // 0: automatic
  preset: Preset | 'auto';
  /** names and terms, spelled as given */
  terms: string;
}

interface Draft {
  name: string;
  transcript: Transcript;
}

const DRAFT_KEY = 'salatic:transcript';

/** Decode any audio/video the browser can play into 16 kHz mono (resampled by the browser). */
async function decodeMedia(file: Blob): Promise<Float32Array> {
  const buf = await file.arrayBuffer();
  const head = new Uint8Array(buf, 0, Math.min(12, buf.byteLength));
  const isWav = String.fromCharCode(...head.subarray(0, 4)) === 'RIFF' && String.fromCharCode(...head.subarray(8, 12)) === 'WAVE';
  try {
    const ctx = new OfflineAudioContext(1, 1, 16000);
    const audio = await ctx.decodeAudioData(buf.slice(0));
    return downmix(Array.from({ length: audio.numberOfChannels }, (_, c) => audio.getChannelData(c)));
  } catch {
    // WAV layouts some browsers do not read (24-bit, float, extensible): our own decoder
    if (isWav) return toSpeechRate(decodeWav(new Uint8Array(buf)));
    throw new Error('Не удалось извлечь звук: браузер не умеет читать этот формат. Сохраните запись в MP3, M4A, WAV или MP4.');
  }
}

export function mount(root: HTMLElement): () => void {
  const saved = prefs.get<Partial<Settings>>('transcribe-settings', {});
  const settings: Settings = { language: 'russian', speakers: 0, preset: 'auto', terms: '', ...saved };
  let worker: Worker | null = null;
  let runId = 0;
  let gpu: { webgpu: boolean; fp16: boolean; adapter?: string } | null = null;
  let media: { file: File; url: string } | null = null;
  let transcript: Transcript | null = null;
  let docName = 'Расшифровка';
  let busy = false;

  const getWorker = (): Worker => {
    if (!worker) {
      worker = new Worker(new URL('../engine/speech-worker.ts', import.meta.url), { type: 'module', name: 'salatic-speech' });
      worker.onmessage = (e: MessageEvent<SpeechResponse>) => onMessage(e.data);
    }
    return worker;
  };
  const send = (m: SpeechRequest, transfer: Transferable[] = []): void => getWorker().postMessage(m, transfer);
  /** With a GPU: a second worker finds the speakers (CPU + GPU) while the first recognises (GPU). */
  let helper: Worker | null = null;
  const getHelper = (): Worker => {
    if (!helper) {
      helper = new Worker(new URL('../engine/speech-worker.ts', import.meta.url), { type: 'module', name: 'salatic-speakers' });
      helper.onmessage = (e: MessageEvent<SpeechResponse>) => onMessage(e.data);
    }
    return helper;
  };
  /** Parts of a split run, put together when both have arrived. */
  let split: { words?: Word[]; turns?: SpeakerTurn[]; regions?: { start: number; end: number }[]; device?: string; asr: number; speakers: number; seg: number; t0: number } | null = null;

  // ------------------------------------------------------------------ layout
  const zone = dropzone({
    accept: ACCEPT,
    multiple: false,
    title: 'Перетащите аудио или видео сюда или нажмите, чтобы выбрать',
    hint: 'MP3, M4A, WAV, OGG/Opus, FLAC, MP4, MOV, WebM и другие форматы, которые воспроизводит браузер · до 2 ГБ',
    onFiles: (files) => void start(files[0]),
  });
  const urlInput = h('input', { class: 'input', type: 'url', placeholder: 'https://… — прямая ссылка, Яндекс.Диск, Google Диск или Dropbox', 'aria-label': 'Ссылка на аудио или видео' });
  const urlBtn = button('Загрузить', () => void fromLink(), { icon: 'link' });
  const urlForm = h(
    'form',
    {
      class: 'stt-link',
      onsubmit: (e: Event) => {
        e.preventDefault();
        void fromLink();
      },
    },
    urlInput,
    urlBtn,
  );

  const select = (label: string, options: [string, string][], value: string, onChange: (v: string) => void, hint?: string): HTMLElement => {
    const s = h('select', { class: 'select input', onchange: (e: Event) => onChange((e.target as HTMLSelectElement).value) }, options.map(([v, t]) => h('option', { value: v, selected: v === value }, t)));
    return h('label', { class: 'field' }, h('span', null, label), s, hint ? h('small', { class: 'muted' }, hint) : null);
  };
  const store = (): void => prefs.set('transcribe-settings', settings);
  const presetHint = h('small', { class: 'muted' });
  const presetSelect = h(
    'select',
    {
      class: 'select input',
      onchange: (e: Event) => {
        settings.preset = (e.target as HTMLSelectElement).value as Settings['preset'];
        store();
        paintDevice();
      },
    },
    h('option', { value: 'auto', selected: settings.preset === 'auto' }, 'Автоматически (лучшее для устройства)'),
    PRESETS.map((p) => h('option', { value: p.id, selected: p.id === settings.preset }, p.label)),
  );
  const deviceLine = h('div', { class: 'stt-device muted' }, 'Проверяю видеокарту…');
  const options = h(
    'div',
    { class: 'stt-options' },
    select('Язык речи', LANGUAGES, settings.language, (v) => {
      settings.language = v;
      store();
    }),
    select(
      'Число спикеров',
      [['0', 'Определить автоматически'], ...Array.from({ length: 10 }, (_, i) => [String(i + 1), String(i + 1)] as [string, string])],
      String(settings.speakers),
      (v) => {
        settings.speakers = Number(v);
        store();
      },
      'Если знаете точно — укажите: так спикеры разделятся надёжнее',
    ),
    h('label', { class: 'field' }, h('span', null, 'Точность'), presetSelect, presetHint),
    h(
      'label',
      { class: 'field stt-terms' },
      h('span', null, 'Имена и термины (необязательно)'),
      h('input', {
        class: 'input',
        value: settings.terms,
        maxlength: 300,
        placeholder: 'Например: Анна Сергеевна Ковальчук, ООО «Ромашка», Kubernetes',
        oninput: (e: Event) => {
          settings.terms = (e.target as HTMLInputElement).value;
          store();
        },
      }),
      h('small', { class: 'muted' }, 'Помогает правильно написать фамилии, названия и специальные слова'),
    ),
  );

  const status = h('div', { class: 'stt-status', 'aria-live': 'polite', hidden: true });
  const result = h('div', { class: 'stt-result' });

  root.append(
    h(
      'div',
      { class: 'page-head' },
      h('h1', null, 'Расшифровка аудио и видео'),
      h('p', null, 'Речь в текст с разделением на спикеров и метками времени. Распознавание работает прямо в браузере: запись не покидает устройство. Модели (Whisper large-v3-turbo, pyannote) скачиваются один раз и сохраняются в браузере.'),
    ),
    h('section', { class: 'card stt-input' }, zone, h('div', { class: 'stt-or muted' }, 'или по ссылке'), urlForm, options, deviceLine),
    status,
    result,
  );

  // ------------------------------------------------------------------ device
  const effectivePreset = (): Preset => {
    if (settings.preset !== 'auto') return settings.preset;
    return gpu?.webgpu && gpu.fp16 ? 'accurate' : 'balanced';
  };
  /** Speed measured on this device (seconds of work per second of recording), by device kind and preset. */
  const speedKey = (): string => `${gpu?.webgpu ? 'gpu' : 'cpu'}|${effectivePreset()}`;
  const measured = (): number | undefined => prefs.get<Record<string, number>>('transcribe-speed', {})[speedKey()];
  const minutes = (sec: number): string => (sec < 90 ? `${Math.max(1, Math.round(sec))} с` : `${Math.round(sec / 60)} мин`);
  const paintDevice = (): void => {
    const p = PRESETS.find((x) => x.id === effectivePreset())!;
    presetHint.textContent = p.hint;
    if (!gpu) return;
    const k = measured();
    const hour = k ? ` По прошлой расшифровке на этом устройстве час записи займёт ≈ ${minutes(k * 3600)}.` : '';
    if (gpu.webgpu)
      deviceLine.replaceChildren(icon('check'), ` Видеокарта доступна (WebGPU${gpu.adapter ? `: ${gpu.adapter}` : ''}${gpu.fp16 ? '' : ', без fp16'}): распознавание идёт на ней пакетами, спикеры ищутся параллельно.${hour}`);
    else
      deviceLine.replaceChildren(
        icon('alert'),
        ` WebGPU недоступен: распознавание пойдёт на процессоре, это в разы медленнее, чем с видеокартой. Для длинных записей откройте сайт в Chrome или Edge на компьютере с видеокартой.${hour}`,
      );
    if (settings.preset === 'accurate' && !(gpu.webgpu && gpu.fp16)) presetHint.textContent += ' — на этом устройстве будет использована 4-битная версия';
  };
  paintDevice();
  send({ type: 'probe' });

  // ------------------------------------------------------------------ progress
  const bar = h('div');
  const label = h('div', { class: 'stt-status__label' });
  const detail = h('div', { class: 'stt-status__detail muted' });
  const cancelBtn = button('Отменить', () => cancel(), { variant: 'ghost', small: true, icon: 'x' });
  status.append(h('div', { class: 'card stt-status__card' }, h('div', { class: 'stt-status__row' }, h('div', { class: 'spinner' }), label, cancelBtn), h('div', { class: 'progress' }, bar), detail));
  const downloads = new Map<string, { loaded: number; total: number }>();
  let stageStart = 0;
  let audioSeconds = 0;
  const show = (text: string, frac: number | null, extra = ''): void => {
    status.hidden = false;
    label.textContent = text;
    bar.style.width = frac === null ? '100%' : `${Math.round(Math.min(1, Math.max(0, frac)) * 100)}%`;
    bar.classList.toggle('progress__indeterminate', frac === null);
    detail.textContent = extra;
  };
  const eta = (frac: number): string => {
    if (frac <= 0.02) return '';
    const spent = (performance.now() - stageStart) / 1000;
    const left = (spent / frac) * (1 - frac);
    return left > 5 ? ` · осталось ≈ ${left > 90 ? `${Math.round(left / 60)} мин` : `${Math.round(left)} с`}` : '';
  };
  let lastStage = '';
  /** When the models were ready and the work began (for the measured speed) */
  let workStart = 0;
  /** Two-worker runs: when the recognition model was ready, and when the speech regions were found */
  let readyAt = 0;
  let regionsAt = 0;

  function onMessage(m: SpeechResponse): void {
    if (m.type === 'probe') {
      gpu = m;
      paintDevice();
      return;
    }
    if ('id' in m && m.id !== runId) return;
    if (m.type === 'note') {
      console.info(`[speech] ${m.message}`);
      return;
    }
    if (m.type === 'progress' && m.stage === 'load' && m.total > 0 && m.done >= m.total) {
      readyAt = performance.now();
      return;
    }
    if (m.type === 'progress' && m.stage === 'segment' && !workStart) workStart = performance.now();
    if (split && onSplit(m)) return;
    if (m.type === 'download') {
      downloads.set(m.file, { loaded: m.loaded, total: m.total });
      const all = [...downloads.values()].filter((d) => d.total > 1_000_000);
      const loaded = all.reduce((s, d) => s + d.loaded, 0);
      const total = all.reduce((s, d) => s + d.total, 0);
      if (total) show('Загрузка моделей распознавания (один раз)', loaded / total, `${formatBytes(loaded)} из ${formatBytes(total)} · следующие расшифровки начнутся сразу`);
      return;
    }
    if (m.type === 'progress') {
      if (m.stage !== lastStage) {
        lastStage = m.stage;
        stageStart = performance.now();
      }
      const f = m.total ? m.done / m.total : 0;
      if (m.stage === 'load') show('Подготовка моделей…', null);
      else if (m.stage === 'segment') show('Поиск речи и голосов', f, eta(f));
      else if (m.stage === 'asr') show('Распознавание речи', f, `${clock(m.done)} из ${clock(m.total)} речи${eta(f)}`);
      else if (m.stage === 'speakers') show('Разделение по спикерам', f, eta(f));
      return;
    }
    if (m.type === 'error') {
      failed(m.message, m.cancelled);
      return;
    }
    if (m.type === 'result') done(m.transcript, m.seconds);
  }

  function failed(message: string, cancelled?: boolean): void {
    if (busy && !cancelled) {
      send({ type: 'cancel', id: runId });
      helper?.postMessage({ type: 'cancel', id: runId } satisfies SpeechRequest);
    }
    split = null;
    finish();
    status.hidden = true;
    if (cancelled) return;
    result.replaceChildren(note('error', `Не удалось распознать: ${message}`));
    // a failed engine (out of memory, lost GPU) is not reused
    worker?.terminate();
    worker = null;
    helper?.terminate();
    helper = null;
    send({ type: 'probe' });
  }

  /** `idle`: seconds spent waiting for a model download after the speech was found (not counted in the speed) */
  function done(tr: Transcript, seconds: number, idle = 0): void {
    split = null;
    finish();
    status.hidden = true;
    transcript = tr;
    saveDraft();
    paintResult();
    const speed = audioSeconds && seconds ? audioSeconds / seconds : 0;
    toast(`Готово за ${clock(seconds)}${speed ? ` (${speed >= 1 ? `в ${speed.toFixed(1)} раза быстрее записи` : `в ${(1 / speed).toFixed(1)} раза медленнее записи`})` : ''}.`);
    // remember the speed of the work itself (model downloads and loading excluded), for recordings of a minute or more
    if (audioSeconds >= 60 && workStart) {
      const all = prefs.get<Record<string, number>>('transcribe-speed', {});
      prefs.set('transcribe-speed', { ...all, [speedKey()]: Math.max(0.01, (performance.now() - workStart) / 1000 - idle) / audioSeconds });
      paintDevice();
    }
  }

  /** Messages of a split run (two workers): progress of both, regions passed on, parts put together. */
  function onSplit(m: SpeechResponse): boolean {
    const sp = split!;
    if (m.type === 'regions') {
      sp.regions = m.regions;
      regionsAt = performance.now();
      send({ type: 'regions', id: runId, regions: m.regions });
      return true;
    }
    if (m.type === 'turns') sp.turns = m.turns;
    else if (m.type === 'words') {
      sp.words = m.words;
      sp.device = m.device;
    } else if (m.type === 'progress' && m.stage !== 'load') {
      const f = m.total ? m.done / m.total : 0;
      if (m.stage === 'segment') sp.seg = f;
      else if (m.stage === 'asr') sp.asr = f;
      else if (m.stage === 'speakers') sp.speakers = f;
      if (m.stage === 'asr' && lastStage !== 'asr') {
        lastStage = 'asr';
        stageStart = performance.now();
      }
      const spk = sp.seg < 1 ? `поиск речи ${Math.round(sp.seg * 100)} %` : `спикеры ${Math.round(sp.speakers * 100)} %`;
      if (!sp.regions) show('Поиск речи и голосов', sp.seg, '');
      else if (sp.words) show('Разделение по спикерам', sp.speakers, 'распознавание речи завершено');
      else show('Распознавание речи', sp.asr, `видеокарта · ${spk}${eta(sp.asr)}`);
      return true;
    } else return false;
    if (sp.words && sp.turns && sp.regions) {
      const idle = readyAt > regionsAt ? (readyAt - regionsAt) / 1000 : 0;
      done(combine(sp.words, sp.turns, sp.regions, audioSeconds, settings.language), (performance.now() - sp.t0) / 1000, idle);
    } else if (m.type === 'words') show('Разделение по спикерам', sp.speakers, 'распознавание речи завершено');
    return true;
  }

  const finish = (): void => {
    busy = false;
    zone.classList.remove('is-disabled');
    urlBtn.disabled = false;
  };

  function cancel(): void {
    if (!busy) return;
    send({ type: 'cancel', id: runId });
    helper?.postMessage({ type: 'cancel', id: runId } satisfies SpeechRequest);
    split = null;
    finish();
    status.hidden = true;
    toast('Распознавание отменено.');
  }

  // ------------------------------------------------------------------ input
  async function fromLink(): Promise<void> {
    const link = urlInput.value.trim();
    if (!link || busy) return;
    busy = true;
    urlBtn.disabled = true;
    try {
      show('Скачивание по ссылке…', null);
      const res = await fetch(`/api/media?url=${encodeURIComponent(link)}`);
      if (!res.ok) {
        const body = (await res.json().catch(() => null)) as { error?: string } | null;
        throw new Error(body?.error ?? `ошибка ${res.status}`);
      }
      const total = Number(res.headers.get('Content-Length') ?? 0);
      const reader = res.body!.getReader();
      const parts: Uint8Array[] = [];
      let got = 0;
      for (;;) {
        const { value, done } = await reader.read();
        if (done) break;
        parts.push(value);
        got += value.length;
        show('Скачивание по ссылке…', total ? got / total : null, total ? `${formatBytes(got)} из ${formatBytes(total)}` : formatBytes(got));
      }
      const name = /filename\*=UTF-8''([^;]+)/.exec(res.headers.get('Content-Disposition') ?? '')?.[1];
      const file = new File(parts as BlobPart[], name ? decodeURIComponent(name) : 'запись', { type: res.headers.get('Content-Type') ?? '' });
      busy = false;
      await start(file);
    } catch (e) {
      status.hidden = true;
      result.replaceChildren(note('error', `Не удалось загрузить по ссылке: ${e instanceof Error ? e.message : e}`));
      finish();
    }
  }

  async function start(file: File | undefined): Promise<void> {
    if (!file || busy) return;
    if (file.size > MAX_BYTES) {
      toast('Файл больше 2 ГБ: сохраните только звуковую дорожку (например, в M4A или MP3).', 'error');
      return;
    }
    busy = true;
    zone.classList.add('is-disabled');
    urlBtn.disabled = true;
    result.replaceChildren();
    if (media) URL.revokeObjectURL(media.url);
    media = { file, url: URL.createObjectURL(file) };
    docName = baseName(file.name) || 'Расшифровка';
    let audio: Float32Array;
    try {
      show('Извлечение звука…', null, `${file.name} · ${formatBytes(file.size)}`);
      audio = await decodeMedia(file);
    } catch (e) {
      finish();
      status.hidden = true;
      result.replaceChildren(note('error', e instanceof Error ? e.message : String(e)));
      return;
    }
    audioSeconds = audio.length / 16000;
    if (audioSeconds < 0.5) {
      finish();
      status.hidden = true;
      result.replaceChildren(note('warn', 'В файле нет звука.'));
      return;
    }
    runId++;
    workStart = 0;
    readyAt = 0;
    regionsAt = 0;
    downloads.clear();
    lastStage = '';
    const k = measured();
    show('Подготовка моделей…', null, `${clock(audioSeconds)} записи${k ? ` · ожидаемое время ≈ ${minutes(k * audioSeconds)}` : ''}`);
    const terms = settings.terms.trim() || undefined;
    const speakers = settings.speakers || undefined;
    if (!gpu?.webgpu) {
      // CPU only: one stage after the other in one worker
      split = null;
      send({ type: 'run', id: runId, audio, preset: effectivePreset(), language: settings.language, speakers, terms }, [audio.buffer]);
      return;
    }
    // GPU: recognition and speaker search side by side, sharing the samples when the page is isolated
    split = { asr: 0, speakers: 0, seg: 0, t0: performance.now() };
    let shared = audio;
    if (self.crossOriginIsolated && typeof SharedArrayBuffer !== 'undefined') {
      shared = new Float32Array(new SharedArrayBuffer(audio.byteLength));
      shared.set(audio);
      getHelper().postMessage({ type: 'diarize', id: runId, audio: shared, speakers } satisfies SpeechRequest);
      send({ type: 'recognize', id: runId, audio: shared, preset: effectivePreset(), language: settings.language, terms });
    } else {
      getHelper().postMessage({ type: 'diarize', id: runId, audio: audio.slice(), speakers } satisfies SpeechRequest);
      send({ type: 'recognize', id: runId, audio, preset: effectivePreset(), language: settings.language, terms }, [audio.buffer]);
    }
  }

  // ------------------------------------------------------------------ result
  function saveDraft(): void {
    if (!transcript) return;
    try {
      localStorage.setItem(DRAFT_KEY, JSON.stringify({ name: docName, transcript } satisfies Draft));
    } catch {
      /* storage full or blocked: the result stays on the page */
    }
  }

  function paintResult(): void {
    if (!transcript) return;
    const tr = transcript;
    const player = media ? (/^video\//.test(media.file.type) ? h('video', { src: media.url, controls: true, class: 'stt-player', preload: 'metadata', playsinline: true }) : h('audio', { src: media.url, controls: true, class: 'stt-player', preload: 'metadata' })) : null;
    const seek = (t: number): void => {
      if (!player) return;
      player.currentTime = Math.max(0, t - 0.2);
      void player.play().catch(() => undefined);
    };
    const names = h(
      'div',
      { class: 'stt-speakers' },
      tr.speakers.map((name, i) =>
        h(
          'label',
          { class: 'stt-speaker' },
          h('span', { class: 'stt-dot', style: `background:${COLORS[i % COLORS.length]}` }),
          h('input', {
            class: 'input',
            value: name,
            'aria-label': `Имя спикера ${i + 1}`,
            oninput: (e: Event) => {
              tr.speakers[i] = (e.target as HTMLInputElement).value || `Спикер ${i + 1}`;
              list.querySelectorAll<HTMLSelectElement>('select.stt-who').forEach((s) => (s.options[i].textContent = tr.speakers[i]));
              saveDraft();
            },
          }),
        ),
      ),
    );
    const list = h(
      'div',
      { class: 'stt-list' },
      tr.utterances.map((u) => {
        const text = h('div', {
          class: 'stt-text',
          contenteditable: 'true',
          spellcheck: 'true',
          lang: 'ru',
          oninput: (e: Event) => {
            u.text = ((e.target as HTMLElement).innerText ?? '').replace(/\s+/g, ' ').trim();
            saveDraft();
          },
        });
        text.textContent = u.text;
        const who = h(
          'select',
          {
            class: 'select stt-who',
            'aria-label': 'Спикер реплики',
            style: `border-color:${COLORS[u.speaker % COLORS.length]}`,
            onchange: (e: Event) => {
              u.speaker = Number((e.target as HTMLSelectElement).value);
              (e.target as HTMLElement).style.borderColor = COLORS[u.speaker % COLORS.length];
              saveDraft();
            },
          },
          tr.speakers.map((n, i) => h('option', { value: String(i), selected: i === u.speaker }, n)),
        );
        return h('div', { class: 'stt-utt' }, h('div', { class: 'stt-meta' }, who, h('button', { type: 'button', class: 'stt-time', title: 'Слушать с этого места', onclick: () => seek(u.start), disabled: !player }, clock(u.start))), text);
      }),
    );
    const save = (data: string | Uint8Array, ext: string, mime: string): void => download(typeof data === 'string' ? new Blob([data], { type: mime }) : data, `${docName}.${ext}`, mime);
    const docx = async (): Promise<Uint8Array> => (await import('../../core/writers/docx')).docToDocx(toDoc(tr, docName)).data;
    const tools = h(
      'div',
      { class: 'toolbar stt-tools' },
      button('Word (DOCX)', async () => save(await docx(), 'docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'), { icon: 'download', variant: 'primary' }),
      button('Текст', () => save(toText(tr), 'txt', 'text/plain;charset=utf-8'), { icon: 'download' }),
      button('Субтитры SRT', () => save(toSrt(tr), 'srt', 'application/x-subrip;charset=utf-8')),
      button('WebVTT', () => save(toVtt(tr), 'vtt', 'text/vtt;charset=utf-8')),
      button('Markdown', () => save(toMarkdown(tr), 'md', 'text/markdown;charset=utf-8')),
      button('JSON', () => save(toJson(tr), 'json', 'application/json')),
      h('span', { class: 'toolbar__spacer' }),
      button('Копировать', async () => {
        await navigator.clipboard.writeText(toText(tr, false));
        toast('Текст скопирован.');
      }, { icon: 'copy', variant: 'ghost' }),
      button('В редактор', async () => handOff(new File([await docx() as BlobPart], `${docName}.docx`), '/doc-editor'), { icon: 'edit-text', variant: 'ghost' }),
    );
    const words = tr.utterances.reduce((s, u) => s + u.text.split(/\s+/).length, 0);
    result.replaceChildren(
      h(
        'section',
        { class: 'card stt-card' },
        h('div', { class: 'stt-head' }, h('h2', null, docName), h('span', { class: 'muted' }, `${clock(tr.duration)} · ${tr.speakers.length} ${plural(tr.speakers.length, 'спикер', 'спикера', 'спикеров')} · ${words} ${plural(words, 'слово', 'слова', 'слов')}`)),
        tr.utterances.length ? null : note('warn', 'Речь не найдена: в записи тишина, музыка или шум.'),
        player,
        h('div', { class: 'stt-hint muted' }, 'Имена спикеров можно задать ниже, текст реплик — исправить прямо в списке; изменения сохраняются в браузере.'),
        names,
        tools,
        list,
      ),
    );
  }

  // the last transcript, if any
  try {
    const raw = localStorage.getItem(DRAFT_KEY);
    if (raw) {
      const d = JSON.parse(raw) as Draft;
      if (d?.transcript?.utterances) {
        transcript = d.transcript;
        docName = d.name || 'Расшифровка';
        paintResult();
        result.prepend(note('ok', 'Последняя расшифровка восстановлена. Загрузите файл, чтобы распознать новый.'));
      }
    }
  } catch {
    /* ignore a broken draft */
  }

  return () => {
    if (busy) send({ type: 'cancel', id: runId });
    worker?.terminate();
    helper?.terminate();
    if (media) URL.revokeObjectURL(media.url);
  };
}

function plural(n: number, one: string, few: string, many: string): string {
  const m10 = n % 10;
  const m100 = n % 100;
  if (m10 === 1 && m100 !== 11) return one;
  if (m10 >= 2 && m10 <= 4 && (m100 < 12 || m100 > 14)) return few;
  return many;
}
