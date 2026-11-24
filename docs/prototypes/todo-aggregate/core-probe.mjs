// Offline probe of the installed Core helper; no gateway, history store or provider.
import assert from 'node:assert/strict';
import {pathToFileURL} from 'node:url';

if (!process.argv[2]) throw new Error('Pass the inspected preemptive-compaction bundle path');
const {d: preflight} = await import(pathToFileURL(process.argv[2]).href);
for (const count of [5, 240, 245, 275]) {
  const messages = Array.from({length: count}, (_, i) => ({
    role: 'user', content: 'synthetic '.repeat(36) + ' turn ' + i, timestamp: i,
  }));
  const result = preflight({systemPrompt: 'Synthetic system prompt', prompt: 'List open tasks.',
    messages, toolSchemaTokens: 300, contextTokenBudget: 32768, reserveTokens: 2048});
  assert.equal(result.toolResultReducibleChars, 0);
  assert.equal(result.route, result.estimatedPromptTokens > 30720 ? 'compact_only' : 'fits');
  console.log(JSON.stringify({count, ...result}));
}
