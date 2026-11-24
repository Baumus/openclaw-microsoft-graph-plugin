// Inactive prototype. All I/O is injected by the adapter; this module never authenticates.
import { modelTextBytes } from './collection-output.ts';

const open = new Set(['notStarted', 'inProgress', 'waitingOnOthers', 'deferred']);
const compareText = (a, b) => a < b ? -1 : a > b ? 1 : 0;
export const compare = (a, b) => a.due - b.due ||
  compareText(a.listId, b.listId) || compareText(a.taskId, b.taskId);

// The adapter admits only authorized lists and validates paths/cursors.
// Each list gets its own row/page quota; no timing-dependent shared scan counter.
export async function aggregate(adapter, {
  limit = 5, maxLists = 20, rowsPerList = 100, pagesPerList = 4,
  concurrency = 3, deadlineMs = 5000, outputBytes = 4096, signal,
} = {}) {
  const bounds = [[limit, 20], [maxLists, 20], [rowsPerList, 100],
    [pagesPerList, 4], [concurrency, 3], [deadlineMs, 5000], [outputBytes, 8192]];
  if (bounds.some(([value, cap]) => !Number.isSafeInteger(value) || value < 1 || value > cap)
      || outputBytes < 1024) throw new Error('invalid_budget');
  const deadline = AbortSignal.timeout(deadlineMs);
  const bounded = signal ? AbortSignal.any([signal, deadline]) : deadline;
  bounded.throwIfAborted();
  const inventory = await adapter.listPage(maxLists, bounded);
  bounded.throwIfAborted();
  const lists = [...new Map(inventory.lists.map(list => [list.id, list])).values()].slice(0, maxLists);
  const inventoryComplete = !inventory.more && inventory.lists.length <= maxLists;
  const coverage = lists.map(({id}) => ({listId: id, state: 'not_started', scanned: 0}));
  const candidates = [];
  let position = 0;
  async function worker() {
    while (!bounded.aborted && position < lists.length) {
      const index = position++;
      const list = lists[index];
      const entry = coverage[index];
      let cursor;
      let unknownStatus = false;
      const seen = new Set();
      try {
        for (let page = 0; page < pagesPerList; page++) {
          bounded.throwIfAborted();
          const remaining = rowsPerList - entry.scanned;
          const result = await adapter.taskPage(list.id, cursor, remaining, bounded);
          bounded.throwIfAborted();
          for (const task of result.items.slice(0, remaining)) {
            entry.scanned++;
            if (seen.has(task.id)) continue;
            seen.add(task.id);
            if (!open.has(task.status)) {
              unknownStatus ||= task.status !== 'completed';
              continue;
            }
            // Adapter normalizes valid due instants; null means no due date.
            candidates.push({listId: list.id, taskId: task.id, title: task.title,
              status: task.status, due: task.due ?? Infinity});
          }
          const over = result.items.length > remaining;
          cursor = result.next;
          entry.state = cursor || over ? 'scan_limit' : unknownStatus ? 'unknown_status' : 'complete';
          if (!cursor || entry.scanned >= rowsPerList) break;
        }
      } catch {
        entry.state = bounded.aborted ? 'deadline' : 'read_error';
      }
    }
  }
  await Promise.all(Array.from({length: Math.min(concurrency, lists.length)}, worker));
  signal?.throwIfAborted(); // caller cancellation is not a successfully published result
  const complete = inventoryComplete && coverage.every(row => row.state === 'complete');
  candidates.sort(compare);
  const top = candidates.slice(0, limit).map(item => ({
    ...item, due: Number.isFinite(item.due) ? new Date(item.due).toISOString() : null,
  }));
  const result = { completeness: complete ? 'complete' : 'partial', inventoryComplete,
    coverage, matchedInScan: candidates.length, outputLimited: false, items: top,
    detailAction: 'get_task' };
  while (modelTextBytes(result) > outputBytes && result.items.length) {
    result.items.pop();
    result.outputLimited = true;
  }
  if (modelTextBytes(result) > outputBytes) throw new Error('coverage_exceeds_output_budget');
  return result;
}
