/**
 * Pool of engine workers. Jobs run in parallel on multi-core machines; a job can be
 * cancelled, which terminates its worker (the pool replaces it lazily).
 */
import type { EngineRequest, WorkerMessage } from './protocol';

interface Job {
  id: number;
  req: EngineRequest;
  transfer: Transferable[];
  resolve: (v: unknown) => void;
  reject: (e: Error) => void;
  onProgress?: (v: number) => void;
  slot?: Slot;
  cancelled?: boolean;
}

interface Slot {
  worker: Worker;
  job: Job | null;
}

export class EngineError extends Error {
  constructor(
    message: string,
    readonly code?: string,
  ) {
    super(message);
    this.name = 'EngineError';
  }
}

const MAX_WORKERS = Math.max(1, Math.min(4, (navigator.hardwareConcurrency || 2) - 1));

class EnginePool {
  private slots: Slot[] = [];
  private queue: Job[] = [];
  private nextId = 1;

  private spawn(): Slot {
    const worker = new Worker(new URL('./worker.ts', import.meta.url), { type: 'module', name: 'salatic-engine' });
    const slot: Slot = { worker, job: null };
    worker.addEventListener('message', (e: MessageEvent<WorkerMessage>) => {
      const msg = e.data;
      const job = slot.job;
      if (!job || job.id !== msg.id) return;
      if (msg.type === 'progress') {
        job.onProgress?.(msg.value);
        return;
      }
      slot.job = null;
      if (msg.type === 'done') job.resolve(msg.result);
      else job.reject(new EngineError(msg.message, msg.code));
      this.pump();
    });
    worker.addEventListener('error', (e) => {
      const job = slot.job;
      slot.job = null;
      this.remove(slot);
      job?.reject(new EngineError(`Сбой обработчика: ${e.message || 'неизвестная ошибка'}`));
      this.pump();
    });
    this.slots.push(slot);
    return slot;
  }

  private remove(slot: Slot): void {
    slot.worker.terminate();
    this.slots = this.slots.filter((s) => s !== slot);
  }

  private pump(): void {
    while (this.queue.length) {
      let slot = this.slots.find((s) => !s.job);
      if (!slot && this.slots.length < MAX_WORKERS) slot = this.spawn();
      if (!slot) return;
      const job = this.queue.shift()!;
      if (job.cancelled) continue;
      slot.job = job;
      job.slot = slot;
      slot.worker.postMessage({ id: job.id, req: job.req }, job.transfer);
    }
  }

  run<T>(req: EngineRequest, opts: { transfer?: Transferable[]; onProgress?: (v: number) => void; signal?: AbortSignal } = {}): Promise<T> {
    return new Promise<T>((resolve, reject) => {
      const job: Job = { id: this.nextId++, req, transfer: opts.transfer ?? [], resolve: resolve as (v: unknown) => void, reject, onProgress: opts.onProgress };
      opts.signal?.addEventListener('abort', () => {
        job.cancelled = true;
        if (job.slot?.job === job) {
          this.remove(job.slot);
          this.pump();
        }
        reject(new EngineError('Отменено', 'cancelled'));
      });
      this.queue.push(job);
      this.pump();
    });
  }
}

export const engine = new EnginePool();
