import { safeId } from "./graph.js";

const LISTS = "/me/todo/lists";
const MAX_LISTS = 20;
const MAX_PAGES = 4;
const MAX_ROWS = 100;
const OPEN = new Set(["notStarted", "inProgress", "waitingOnOthers", "deferred"]);

type CoverageState = "complete" | "scan_limit" | "read_error" | "deadline" | "not_started" | "invalid_row";
type Coverage = { listId: string; listName: string; state: CoverageState; scanned: number };
type Task = { listId: string; listName: string; taskId: string; title: string; status: string; dueDateTime: string | null };
type NextPath = (link: unknown, prefix: string) => string | undefined;
type ReadPage = (path: string) => Promise<unknown>;

function record(value: unknown): Record<string, unknown> {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new Error("invalid_provider_response");
  return value as Record<string, unknown>;
}

function page(value: unknown, path: string, nextPath: NextPath) {
  const payload = record(value);
  if (!Array.isArray(payload.value)) throw new Error("invalid_provider_response");
  const link = payload["@odata.nextLink"];
  return {
    rows: payload.value as unknown[],
    next: nextPath(link, path),
  };
}

function dueDate(value: unknown): string | null {
  if (value == null) return null;
  const { dateTime, timeZone } = record(value);
  if (typeof dateTime !== "string" || !/^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d(?:\.\d{1,7})?(?:Z|[+-]\d\d:\d\d)?$/i.test(dateTime)) throw new Error("invalid_provider_response");
  const hasOffset = /(?:Z|[+-]\d\d:\d\d)$/i.test(dateTime);
  if (!hasOffset && timeZone !== "UTC" && timeZone !== "Etc/UTC") throw new Error("invalid_provider_response");
  const calendarDay = dateTime.slice(0, 10);
  const day = new Date(calendarDay + "T00:00:00Z");
  if (!Number.isFinite(day.getTime()) || day.toISOString().slice(0, 10) !== calendarDay) throw new Error("invalid_provider_response");
  const instant = new Date(hasOffset ? dateTime : dateTime + "Z");
  if (!Number.isFinite(instant.getTime())) throw new Error("invalid_provider_response");
  return instant.toISOString();
}

const compareText = (a: string, b: string) => a < b ? -1 : a > b ? 1 : 0;
function dueTime(task: Task) {
  return task.dueDateTime === null ? Infinity : Date.parse(task.dueDateTime);
}

export async function readTodoAggregate({
  read, nextPath, bounded, callerSignal, runWorkers, limit,
}: {
  read: ReadPage;
  nextPath: NextPath;
  bounded: AbortSignal;
  callerSignal?: AbortSignal;
  runWorkers: (worker: () => Promise<void>, count: number) => Promise<void>;
  limit: number;
}) {
  const inventory = page(await read(`${LISTS}?$top=${MAX_LISTS}`), LISTS, nextPath);
  bounded.throwIfAborted();
  const lists = new Map<string, Coverage>();
  let inventoryComplete = !inventory.next && inventory.rows.length <= MAX_LISTS;
  for (const value of inventory.rows.slice(0, MAX_LISTS)) {
    try {
      const { id, displayName } = record(value);
      if (typeof id !== "string") throw new Error("invalid_provider_response");
      safeId(id);
      lists.set(id, { listId: id, listName: typeof displayName === "string" ? displayName : id, state: "not_started", scanned: 0 });
    } catch {
      inventoryComplete = false;
    }
  }
  const coverage = [...lists.values()];
  const candidates: Task[] = [];
  let position = 0;
  await runWorkers(async () => {
    while (!bounded.aborted && position < coverage.length) {
      const entry = coverage[position++];
      const prefix = `${LISTS}/${safeId(entry.listId)}/tasks`;
      const seen = new Set<string>();
      let invalidRow = false;
      let path = `${prefix}?$top=25`;
      try {
        for (let i = 0; i < MAX_PAGES; i++) {
          bounded.throwIfAborted();
          const result = page(await read(path), prefix, nextPath);
          bounded.throwIfAborted();
          const remaining = MAX_ROWS - entry.scanned;
          for (const value of result.rows.slice(0, remaining)) {
            entry.scanned++;
            try {
              const { id, title, status, dueDateTime } = record(value);
              if (status === "completed") continue;
              if (typeof id !== "string" || typeof title !== "string" || typeof status !== "string" || !OPEN.has(status)) throw new Error("invalid_provider_response");
              safeId(id);
              if (seen.has(id)) continue;
              const due = dueDate(dueDateTime);
              seen.add(id);
              candidates.push({ listId: entry.listId, listName: entry.listName,
                taskId: id, title, status, dueDateTime: due });
            } catch {
              invalidRow = true;
            }
          }
          const truncated = Boolean(result.next) || result.rows.length > remaining;
          entry.state = truncated ? "scan_limit" : invalidRow ? "invalid_row" : "complete";
          if (!result.next || entry.scanned >= MAX_ROWS) break;
          path = result.next;
        }
      } catch {
        entry.state = bounded.aborted ? "deadline" : "read_error";
      }
    }
  }, Math.min(coverage.length, 3));
  callerSignal?.throwIfAborted();

  const completeness = inventoryComplete && coverage.every(({ state }) => state === "complete") ? "complete" : "partial";
  candidates.sort((a, b) => dueTime(a) - dueTime(b) || compareText(a.listId, b.listId) || compareText(a.taskId, b.taskId));
  return {
    ok: true, action: "aggregate_open_tasks", completeness, inventoryComplete,
    listsDiscovered: coverage.length, listsComplete: coverage.filter(({ state }) => state === "complete").length,
    coverage: coverage.map(({ listId, state, scanned }) => ({ listId, state, scanned })), matchedInScan: candidates.length, items: candidates.slice(0, limit),
    outputLimited: false, coverageOutputLimited: false, detailAction: "get_task",
  };
}
