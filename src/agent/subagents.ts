import type { AppSettings } from '../utils/storage';
import { createChatCompletion, type ChatMessage } from './llmClient';
import { TOOL_DEFINITIONS, isMutatingTool, isToolName, type ToolName } from './tools';
import { type ExcelToolName, ToolError, executeTool } from './excel';

/** A read-only agent can inspect and analyze but never change the workbook. */
// use_skill and spawn_agents are handled by the agent loop, not by sub-agents.
const READ_ONLY_TOOLS = TOOL_DEFINITIONS.filter(t => !isMutatingTool(t.function.name as ToolName) && !['use_skill', 'spawn_agents'].includes(t.function.name));
const MAX_SUBAGENT_STEPS = 15;

export interface SubagentSpec {
  name: string;
  task: string;
  /** What to extract from the result, so the orchestrator doesn't get raw noise back. */
  extract?: string;
}

export interface SubagentResult {
  name: string;
  ok: boolean;
  summary: string;
  tokens?: { prompt: number; completion: number };
}

const subagentSystemPrompt = (agentName: string, task: string) => `You are the sub-agent "${agentName}" of an Excel agent team. The orchestrator agent gave you this task:

${task}

You can only inspect and analyze the workbook (read-only tools); you cannot change anything. Do exactly the task and nothing else: no fixes, no improvements, no writing to the workbook.
End with a final message starting with "RESULT:" followed by what the orchestrator needs (or "ERROR: <reason>" if you couldn't complete it).`;

/** Runs one read-only sub-agent conversation to completion. */
const runSubagent = async (settings: AppSettings, spec: SubagentSpec, signal?: AbortSignal) => {
  const messages: ChatMessage[] = [
    { role: 'system', content: subagentSystemPrompt(spec.name, spec.task) },
    { role: 'user', content: 'Do your task.' },
  ];
  let prompt = 0;
  let completion = 0;
  let lastError: Error | null = null;
  let failedCalls = 0;

  for (let step = 0; step < MAX_SUBAGENT_STEPS; step++) {
    signal?.throwIfAborted();
    const { content, toolCalls, finishReason, usage } = await createChatCompletion(
      settings,
      messages,
      READ_ONLY_TOOLS,
      { signal, stream: false }
    );
    if (usage?.promptTokens) prompt += usage.promptTokens;
    if (usage?.completionTokens) completion += usage.completionTokens;
    messages.push({ role: 'assistant', content, ...(toolCalls.length ? { tool_calls: toolCalls } : {}) });

    if (toolCalls.length === 0) {
      const text = content?.trim() ?? '';
      if (finishReason === 'length') return { ok: false, summary: `Ran out of output tokens mid-answer. Last output: ${text.slice(-500)}`, tokens: { prompt, completion } };
      const clean = text.replace(/^RESULT:\s*/i, '').slice(0, 4000);
      // Tool failures (including blocked writes) are visible to the orchestrator.
      const note = failedCalls ? ` (${failedCalls} tool call${failedCalls > 1 ? 's' : ''} failed${lastError instanceof Error ? `; last: ${lastError.message}` : ''})` : '';
      return { ok: !/^ERROR:/i.test(text), summary: clean + note, tokens: { prompt, completion } };
    }

    for (const call of toolCalls) {
      if (!isToolName(call.function.name)) {
        messages.push({ role: 'tool', tool_call_id: call.id, content: JSON.stringify({ error: `Unknown tool "${call.function.name}".` }) });
        continue;
      }
      if (isMutatingTool(call.function.name)) {
        lastError = new ToolError('Sub-agents are read-only; the orchestrator will apply changes.');
        messages.push({ role: 'tool', tool_call_id: call.id, content: JSON.stringify({ error: lastError.message }) });
        failedCalls++;
        continue;
      }
      try {
        let args: Record<string, unknown> = {};
        const parsed: unknown = JSON.parse(call.function.arguments || '{}');
        if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) args = parsed as Record<string, unknown>;
        const result = await executeTool(call.function.name as ExcelToolName, args, { maxReadCells: settings.maxReadCells });
        messages.push({ role: 'tool', tool_call_id: call.id, content: JSON.stringify(result) });
      } catch (e) {
        const message = e instanceof ToolError ? e.message : e instanceof Error ? `${e.message} (${(e as { code?: string }).code ?? 'Excel error'})` : String(e);
        messages.push({ role: 'tool', tool_call_id: call.id, content: JSON.stringify({ error: message }) });
        lastError = e instanceof Error ? e : new Error(message);
        failedCalls++;
      }
    }
  }
  return { ok: false, summary: `Reached ${MAX_SUBAGENT_STEPS} steps without finishing. ${lastError instanceof Error ? `Last error: ${lastError.message}` : ''}`, tokens: { prompt, completion } };
};

/** Applies the orchestrator's own guard: sub-agent tasks must be phrased as inspection. */
const describeSpec = (spec: SubagentSpec) => `${spec.name}: ${spec.extract ? `${spec.task} (report: ${spec.extract})` : spec.task}`;

/**
 * Runs sub-agents with limited concurrency, honoring the abort signal.
 */
const runAll = async (settings: AppSettings, specs: SubagentSpec[], signal?: AbortSignal, onProgress?: (done: number, total: number) => void) => {
  const MAX_CONCURRENT = 3;
  const results: (SubagentResult | null)[] = new Array(specs.length).fill(null);
  let next = 0;
  let done = 0;

  const worker = async () => {
    for (;;) {
      const index = next++;
      if (index >= specs.length) return;
      try {
        const outcome = await runSubagent(settings, specs[index], signal);
        results[index] = { name: specs[index].name, ...outcome };
      } catch (e) {
        if (signal?.aborted) throw e;
        results[index] = { name: specs[index].name, ok: false, summary: e instanceof Error ? e.message : String(e) };
      }
      done++;
      onProgress?.(done, specs.length);
    }
  };

  await Promise.all(Array.from({ length: Math.min(MAX_CONCURRENT, specs.length) }, worker));
  return results.filter(Boolean) as SubagentResult[];
};

export interface SpawnAgentsArgs {
  specs: SubagentSpec[];
  strategy?: string;
}

/** Called by the agent loop when the model decides to split a task. */
export const spawnSubagents = async (
  settings: AppSettings,
  rawArgs: { specs?: unknown; strategy?: unknown },
  signal: AbortSignal | undefined,
  onProgress?: (text: string) => void
) => {
  if (!Array.isArray(rawArgs.specs) || rawArgs.specs.length === 0) throw new ToolError('"specs" must be an array of {name, task} with at least one sub-agent.');
  if (rawArgs.specs.length > 6) throw new ToolError(`At most 6 sub-agents per split (got ${rawArgs.specs.length}); run another split afterwards.`);
  const specs: SubagentSpec[] = rawArgs.specs.map((item, i) => {
    const spec = (item ?? {}) as { name?: unknown; task?: unknown; extract?: unknown };
    if (typeof spec.name !== 'string' || !spec.name.trim() || typeof spec.task !== 'string' || !spec.task.trim()) {
      throw new ToolError(`"specs[${i}]" must have a non-empty "name" and "task".`);
    }
    return { name: spec.name.trim().slice(0, 60), task: spec.task.trim(), ...(typeof spec.extract === 'string' ? { extract: spec.extract.trim() } : {}) };
  });

  const total = specs.length;
  onProgress?.(`Running ${total} sub-agent${total > 1 ? 's' : ''}…`);
  const results = await runAll(settings, specs, signal, (done, count) => onProgress?.(`Sub-agent ${done} of ${count} finished…`));

  const failed = results.filter(r => !r.ok);
  return {
    strategy: typeof rawArgs.strategy === 'string' ? rawArgs.strategy : 'parallel inspection',
    agents: results,
    summary: results.map(r => `${r.ok ? '✓' : '✗'} ${r.name}: ${r.summary}`).join('\n'),
    ...(failed.length ? { note: `${failed.length} of ${total} sub-agents failed. Review their summaries before continuing.` } : {}),
  };
};

export const describeSpecs = (specs: SubagentSpec[]) => specs.map(describeSpec).join(' | ');
