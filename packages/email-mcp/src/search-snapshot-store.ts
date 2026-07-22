export interface SearchSnapshotStoreOptions {
  ttlMs: number;
  maxEntries: number;
  maxBytes: number;
  cleanupIntervalMs?: number;
  now?: () => number;
}

interface StoredSnapshot<T> {
  signature: string;
  value: T;
  bytes: number;
  expiresAt: number;
}

export class SearchSnapshotStore<T> {
  private readonly entries = new Map<string, StoredSnapshot<T>>();
  private readonly now: () => number;
  private readonly timer: NodeJS.Timeout;
  private totalBytes = 0;

  constructor(private readonly options: SearchSnapshotStoreOptions) {
    this.now = options.now ?? Date.now;
    this.timer = setInterval(
      () => this.pruneExpired(),
      options.cleanupIntervalMs ?? Math.min(options.ttlMs, 60_000),
    );
    this.timer.unref();
  }

  get(id: string, signature: string): T | undefined {
    this.pruneExpired();
    const entry = this.entries.get(id);
    if (!entry || entry.signature !== signature) {
      if (entry) this.delete(id);
      return undefined;
    }
    entry.expiresAt = this.now() + this.options.ttlMs;
    return entry.value;
  }

  set(id: string, signature: string, value: T, bytes: number): void {
    if (!Number.isSafeInteger(bytes) || bytes < 0 || bytes > this.options.maxBytes) {
      throw new Error('SEARCH_SNAPSHOT_TOO_LARGE: projected result exceeds the cache byte budget');
    }
    this.pruneExpired();
    this.delete(id);
    while (
      this.entries.size >= this.options.maxEntries ||
      this.totalBytes + bytes > this.options.maxBytes
    ) {
      const oldest = this.entries.keys().next().value as string | undefined;
      if (!oldest) break;
      this.delete(oldest);
    }
    this.entries.set(id, {
      signature,
      value,
      bytes,
      expiresAt: this.now() + this.options.ttlMs,
    });
    this.totalBytes += bytes;
  }

  delete(id: string): void {
    const entry = this.entries.get(id);
    if (!entry) return;
    this.totalBytes -= entry.bytes;
    this.entries.delete(id);
  }

  pruneExpired(): void {
    const now = this.now();
    for (const [id, entry] of this.entries) {
      if (entry.expiresAt <= now) this.delete(id);
    }
  }

  diagnostics(): { entries: number; bytes: number } {
    return { entries: this.entries.size, bytes: this.totalBytes };
  }

  dispose(): void {
    clearInterval(this.timer);
    this.entries.clear();
    this.totalBytes = 0;
  }
}
