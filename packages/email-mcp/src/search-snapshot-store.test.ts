import { afterEach, describe, expect, it, vi } from 'vitest';
import { SearchSnapshotStore } from './search-snapshot-store.js';

const stores: Array<SearchSnapshotStore<unknown>> = [];

function store<T>(options: ConstructorParameters<typeof SearchSnapshotStore<T>>[0]): SearchSnapshotStore<T> {
  const result = new SearchSnapshotStore<T>(options);
  stores.push(result as SearchSnapshotStore<unknown>);
  return result;
}

afterEach(() => {
  for (const item of stores.splice(0)) item.dispose();
  vi.useRealTimers();
});

describe('bounded search snapshot store', () => {
  it('evicts by aggregate bytes and entry count', () => {
    const snapshots = store<string>({ ttlMs: 1000, maxEntries: 2, maxBytes: 10 });
    snapshots.set('a', 'sig', 'a', 6);
    snapshots.set('b', 'sig', 'b', 4);
    snapshots.set('c', 'sig', 'c', 4);
    expect(snapshots.get('a', 'sig')).toBeUndefined();
    expect(snapshots.diagnostics()).toEqual({ entries: 2, bytes: 8 });
    snapshots.set('d', 'sig', 'd', 4);
    expect(snapshots.get('b', 'sig')).toBeUndefined();
    expect(snapshots.diagnostics()).toEqual({ entries: 2, bytes: 8 });
  });

  it('rejects one projection larger than the global byte budget', () => {
    const snapshots = store<string>({ ttlMs: 1000, maxEntries: 2, maxBytes: 10 });
    expect(() => snapshots.set('a', 'sig', 'a', 11)).toThrow(/SEARCH_SNAPSHOT_TOO_LARGE/);
  });

  it('actively expires entries and renews valid access', () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date('2026-07-22T12:00:00Z'));
    const snapshots = store<string>({ ttlMs: 1000, cleanupIntervalMs: 100, maxEntries: 2, maxBytes: 10 });
    snapshots.set('a', 'sig', 'value', 5);
    vi.advanceTimersByTime(900);
    expect(snapshots.get('a', 'sig')).toBe('value');
    vi.advanceTimersByTime(900);
    expect(snapshots.diagnostics()).toEqual({ entries: 1, bytes: 5 });
    vi.advanceTimersByTime(200);
    expect(snapshots.diagnostics()).toEqual({ entries: 0, bytes: 0 });
  });

  it('invalidates a token when its signature changes', () => {
    const snapshots = store<string>({ ttlMs: 1000, maxEntries: 2, maxBytes: 10 });
    snapshots.set('a', 'sig-a', 'value', 5);
    expect(snapshots.get('a', 'sig-b')).toBeUndefined();
    expect(snapshots.diagnostics()).toEqual({ entries: 0, bytes: 0 });
  });
});
