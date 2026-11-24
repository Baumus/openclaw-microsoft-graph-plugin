import assert from 'node:assert/strict';
import {setTimeout as delay} from 'node:timers/promises';
import {aggregate} from './aggregate.mjs';
import {boundedCollectionOutput, modelTextBytes} from './collection-output.ts';

const epoch = Date.UTC(2030, 0, 1);
const lists = Array.from({length: 10}, (_, i) => ({id: 'L' + (i + 1)}));
const fixture = new Map(lists.map(({id}, i) => [id,
  Array.from({length: 33}, (_, j) => ({
    id: id + '-' + j,
    title: j < 30 ? 'Erledigt "' + 'synthetischer Platzhalter '.repeat(12) + j : 'Synthetisch ' + id + '-' + j,
    status: j < 30 ? 'completed' : ['notStarted', 'waitingOnOthers', 'deferred'][j - 30],
    due: epoch + (j < 30 ? -100 : (9 - i) + 10 * (j - 30)) * 86400000,
    body: 'nur Mock',
  }))]));
const expected = ['L10-30', 'L9-30', 'L8-30', 'L7-30', 'L6-30'];
const omittedFields = ['body', 'bodyLastModifiedDateTime', 'startDateTime', 'reminderDateTime',
  'isReminderOn', 'completedDateTime', 'recurrence', 'categories', 'hasAttachments',
  'createdDateTime', 'lastModifiedDateTime'];

function lifecycle(result) {
  const phase = result.completeness === 'partial' ? 'partial' : 'complete';
  const nextAction = phase === 'complete' ? 'No further action required.' :
    result.continuation ? 'Repeat the same criteria with continuation for the next page.' :
      'Narrow the query or time range; completeness is not proven.';
  return {...result, phase, code: phase === 'partial' ? 'partial_results' : 'ok',
    nextAction, retrySafety: 'safe', mutationApplied: false};
}

function baselinePage(rows) {
  const summaries = rows.map(({id, title, status, due}) => ({id, title, status,
    importance: 'normal', dueDateTime: {dateTime: new Date(due).toISOString(), timeZone: 'UTC'}}));
  const {value, remaining} = boundedCollectionOutput(summaries, 8192, (items, withheld) =>
    lifecycle({ok: true, action: 'list_tasks', items, truncated: withheld,
      completeness: withheld ? 'partial' : 'complete', view: 'summary', detailAction: 'get_task',
      omittedFields, outputLimited: withheld,
      ...(withheld ? {continuation: 'mgc1_' + 'x'.repeat(43)} : {})}), modelTextBytes);
  if (remaining.length) assert.equal(value.items.length + remaining.length, rows.length);
  return value;
}

async function replay(variant, initialTokens, fault) {
  const start = performance.now();
  const calls = [];
  let activeReads = 0;
  let peakReads = 0;
  const adapter = {
    async listPage(_limit, signal) {
      calls.push('inventory');
      await delay(12, undefined, {signal});
      return {lists, more: false};
    },
    async taskPage(id, cursor, remaining, signal) {
      calls.push(id);
      peakReads = Math.max(peakReads, ++activeReads);
      try {
        await delay(fault === 'deadline' && id === 'L1' ? 120 : 12, undefined, {signal});
        if (fault === 'error' && id === 'L10') throw new Error('mock_503');
        const rows = fixture.get(id);
        const offset = cursor ?? 0;
        const end = Math.min(offset + remaining, rows.length);
        return {items: rows.slice(offset, end), next: end < rows.length ? end : undefined};
      } finally {
        activeReads--;
      }
    },
  };
  const attempts = [];
  const transcript = [];
  const promptBudget = 32768;
  function emit(userAnswer, dataComplete, ids = [], coverage) {
    return {variant, initialTokens, fault, attempts, modelCalls: attempts.filter(a => a.admitted).length,
      actualProviderModelCalls: 0, graphMockCalls: calls.length, peakReads,
      runtimeMs: +(performance.now() - start).toFixed(1), userAnswer, dataComplete, ids, coverage};
  }
  function modelStep(phase) {
    const toolTokens = transcript.reduce((n, event) => n + Math.ceil(modelTextBytes(event) / 4), 0);
    const estimatedTokens = initialTokens + toolTokens;
    const admitted = estimatedTokens <= promptBudget;
    attempts.push({phase, estimatedTokens, toolTokens, admitted});
    return admitted;
  }
  if (!modelStep('initial')) return emit(null, false);
  if (variant === 'aggregate') {
    const value = await aggregate(adapter, fault === 'scan' ? {rowsPerList: 31} :
      fault === 'deadline' ? {deadlineMs: 38} : {});
    transcript.push({action: 'aggregate', result: lifecycle(value)});
    const ids = value.items.map(item => item.taskId);
    if (!modelStep('answer')) return emit(null, false);
    const complete = value.completeness === 'complete';
    if (fault === 'deadline') {
      const observedIds = new Set(value.coverage.filter(row => row.scanned === 33).map(row => row.listId));
      const deadlineOracle = [...fixture.entries()].filter(([id]) => observedIds.has(id))
        .flatMap(([, tasks]) => tasks.slice(30)).sort((a, b) => a.due - b.due)
        .slice(0, 5).map(task => task.id);
      assert.equal(complete, false);
      assert.deepEqual(ids, deadlineOracle);
    } else {
      assert.deepEqual(ids, fault === 'error' ? ['L9-30','L8-30','L7-30','L6-30','L5-30'] : expected);
    }
    if (fault === 'scan') assert.ok(value.coverage.every(x => x.state === 'scan_limit' && x.scanned === 31));
    if (fault === 'error') assert.equal(value.coverage.at(-1).state, 'read_error');
    if (fault === 'deadline') assert.ok(value.coverage.some(x => x.state === 'deadline'));
    assert.ok(peakReads <= 3);
    const full = value.coverage.filter(x => x.state === 'complete').length;
    const scope = complete ? 'Alle 10 Listen vollständig geprüft.' :
      `Teilabdeckung: ${full}/10 Listen vollständig. Die Rangfolge gilt nur für die gelesenen Aufgaben.`;
    const hits = value.items.map(item => `${item.title} (${item.listId}, ${item.due.slice(0, 10)})`);
    const answer = [scope, ...hits, ...(!complete ? value.coverage.filter(x => x.state !== 'complete')
      .map(x => `${x.listId}: ${({scan_limit:'Scanlimit', read_error:'Lesefehler',
        deadline:'Zeitlimit', not_started:'nicht begonnen'})[x.state]}`) : [])].join('\n');
    return emit(answer, complete, ids, value.coverage);
  }
  const inventory = await adapter.listPage(10);
  transcript.push({action: 'list_lists', result: inventory});
  if (!modelStep('fanout')) return emit(null, false);
  const results = await Promise.all(inventory.lists.map(async ({id}, index) => {
    if (index >= 4) return {listId: id, error: 'request_concurrency_exceeded'};
    const page = await adapter.taskPage(id, null, 50);
    return {listId: id, result: baselinePage(page.items)};
  }));
  transcript.push(...results);
  assert.equal(results.filter(x => x.error).length, 6);
  assert.ok(results.slice(0, 4).every(x => x.result.outputLimited));
  assert.ok(results.slice(0, 4).every(x => x.result.items.every(row => row.status === 'completed')));
  if (!modelStep('answer')) return emit(null, false);
  return emit('Teilabdeckung: 0/10 Listen vollständig. L1–L4: nur erste, bytebegrenzte Seite; darin keine offenen Aufgaben sichtbar. L5–L10: nicht gelesen (Parallelitätsgrenze). Keine belastbare Gesamtrangfolge.', false,
    [], results.map(({listId, error}) => ({listId, state: error ?? 'output_limited'})));
}

const runs = [];
for (const initial of [2500, 30000, 32760]) {
  for (const variant of ['#113', 'aggregate']) runs.push(await replay(variant, initial));
}
for (const fault of ['scan', 'error', 'deadline']) runs.push(await replay('aggregate', 2500, fault));
assert.equal(runs[2].userAnswer, null);
assert.deepEqual(runs[3].ids, expected);
assert.equal(runs[5].userAnswer, null);
console.log(JSON.stringify(runs, null, 2));
