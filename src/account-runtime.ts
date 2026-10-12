import { randomUUID } from "node:crypto";
import { createAccountDirectory } from "./account-directory.js";
import type { OpenClawPluginApi, OpenClawPluginToolContext } from "openclaw/plugin-sdk/core";
import { AccountStore, blankAccountPolicy, type AccountHostContext, type WorkerBindings } from "./accounts.js";
import { accountScope } from "./account-scope.js";
import { normalizedCriteria } from "./continuation.js";
import { validatePolicy, type GraphPolicy } from "./policy.js";

type Config = { enabled?: boolean; policy?: GraphPolicy; accountWorkers?: WorkerBindings; accountPolicyCeilings?: Record<string, GraphPolicy> };
type HookEvent = { toolName: string; toolCallId?: string };
type HookContext = { agentId?: string; sessionKey?: string; sessionId?: string };
type Approval = { onResolution?: (decision: "allow-once" | "allow-always" | "deny" | "timeout" | "cancelled") => void | Promise<void> };
type HookResult = { block?: boolean; blockReason?: string; params?: Record<string, unknown>; requireApproval?: Approval } | undefined | void;
function code(error: unknown): string { const raw = error instanceof Error ? error.message : "internal_error"; return /^(graph_|access_denied|trusted_|approval_context_)/.test(raw) ? raw : "internal_error"; }
function failureResult(error: unknown, toolName?: string) { const errorCode = code(error); const hints: Record<string, string> = { graph_single_user_topology_changed: "Automatic single-user access was permanently disabled after a directory/account change. Use explicit personal consent and session/fixed-worker setup; returning to one user cannot revive the automatic grant.", graph_single_user_directory_unavailable: "Automatic single-user access needs a fresh complete directory from this exact supported local host. No stale or global fallback is allowed. Retry the supported host path or use explicit personal configuration.", graph_execution_account_required: "Use the authenticated Mein Microsoft-Konto page to bind an explicitly shared account to this session, or select an approved fixed worker. No implicit account fallback exists.", graph_execution_account_conflict: "A fixed worker cannot change accounts. Use a different approved worker or session.", graph_session_authority_required: "This host path cannot prove exact current session authority. Use a supported authenticated session; do not invent a profile or binding.", access_denied: "The account owner must explicitly permit shared execution and the executing agent operation within the administrative policy ceiling." }; const changed = ["graph_account_context_changed", "graph_single_user_topology_changed", "graph_single_user_directory_unavailable"].includes(errorCode); const mutation = !!toolName && /(write|upload|update|delete|create_folder)/.test(toolName); const details = { ok: false, error: errorCode, ...(hints[errorCode] ? { nextAction: hints[errorCode] } : {}), ...(changed && mutation ? { outcome: "unknown", remoteEffectsMayHaveOccurred: true, nextAction: "Read back the exact target in a newly authorized account context before retrying. Earlier remote effects cannot be recalled by revocation." } : {}) }; return { details, content: [{ type: "text" as const, text: JSON.stringify(details) }] }; }
export function createAccountRuntime(api: OpenClawPluginApi, runtimeConfig: Config, stateDir?: string) {
  // The host can import the plugin into separate module instances for hooks and tools.
  // Authority and busy state must be shared by the same protected root/config generation.
  const stateKey = Symbol.for("@baumus/openclaw-microsoft-graph/account-runtime-state/v4/single-user-snapshot-v1");
  const slots = globalThis as Record<symbol, unknown>;
  const states = (slots[stateKey] ??= new Map<string, { invocationContexts: Map<string, OpenClawPluginToolContext<2>>; store?: AccountStore; hostInstanceId: string }>()) as Map<string, { invocationContexts: Map<string, OpenClawPluginToolContext<2>>; store?: AccountStore; hostInstanceId: string }>;
  const generationKey = normalizedCriteria([stateDir, runtimeConfig.enabled, runtimeConfig.policy, runtimeConfig.accountWorkers, runtimeConfig.accountPolicyCeilings]);
  let state = states.get(generationKey);
  if (!state) { state = { invocationContexts: new Map(), hostInstanceId: randomUUID() }; states.set(generationKey, state); }
  const invocationContexts = state.invocationContexts;
  const contextKey = (ctx: { agentId?: string; sessionId?: string; sessionKey?: string }) => JSON.stringify([ctx.agentId, ctx.sessionKey, ctx.sessionId]);
  const hostBusy = (target: { agentId: string; sessionKey: string; sessionId: string }): boolean => {
    const invocation = invocationContexts.get(contextKey(target));
    if (!invocation) return false;
    try { invocation.assertInvocationCurrent(); return true; } catch { return false; }
  };
  const accountStore = stateDir ? state.store ??= new AccountStore(stateDir, runtimeConfig.accountWorkers, runtimeConfig.policy ? validatePolicy(runtimeConfig.policy) : blankAccountPolicy(), runtimeConfig.accountPolicyCeilings, hostBusy, createAccountDirectory(api, state.hostInstanceId)) : undefined;
  const hostTarget = (ctx: OpenClawPluginToolContext<2>): AccountHostContext => {
    if (!ctx.agentId || !ctx.sessionKey || !ctx.sessionId || typeof ctx.assertInvocationCurrent !== "function") throw new Error("graph_session_authority_required");
    const lookup = api.runtime?.agent?.session?.getSessionEntry;
    const entry = lookup?.({ agentId: ctx.agentId, sessionKey: ctx.sessionKey, readConsistency: "latest" });
    if (entry?.sessionId !== ctx.sessionId) throw new Error("graph_account_context_changed");
    if (!entry.lifecycleRevision) throw new Error("graph_session_authority_required");
    const initialLifecycle = entry.lifecycleRevision;
    return { agentId: ctx.agentId, sessionKey: ctx.sessionKey, sessionId: ctx.sessionId, lifecycleRevision: initialLifecycle, assertCurrent: () => {
      try { ctx.assertInvocationCurrent(); } catch { throw new Error("graph_account_context_changed"); }
      const latest = lookup?.({ agentId: ctx.agentId, sessionKey: ctx.sessionKey!, readConsistency: "latest" });
      if (!latest || latest.sessionId !== ctx.sessionId || latest.lifecycleRevision !== initialLifecycle) throw new Error("graph_account_context_changed");
      const currentConfig = ctx.getRuntimeConfig?.()?.plugins?.entries?.["microsoft-graph"]?.config;
      if (currentConfig && normalizedCriteria([currentConfig.enabled, currentConfig.policy, currentConfig.accountWorkers, currentConfig.accountPolicyCeilings]) !== normalizedCriteria([runtimeConfig.enabled, runtimeConfig.policy, runtimeConfig.accountWorkers, runtimeConfig.accountPolicyCeilings])) throw new Error("graph_account_context_changed");
    } };
  };
  const scopedConfig = new Proxy(runtimeConfig, { get(target, key, receiver) { return key === "policy" && accountScope.getStore() ? accountScope.getStore()!.policy : Reflect.get(target, key, receiver); } });
  const registrationApi = new Proxy(api, { get(target, key, receiver) {
    if (key === "pluginConfig") return scopedConfig;
    if (key !== "registerTool") return Reflect.get(target, key, receiver);
    return (factory: Parameters<OpenClawPluginApi["registerTool"]>[0], options: Parameters<OpenClawPluginApi["registerTool"]>[1]) => {
      if (typeof factory !== "function") { api.registerTool(factory, options); return; }
      api.registerTool({ contextVersion: 2, create(ctx) {
        if (!invocationContexts.has(contextKey(ctx)) && invocationContexts.size >= 1024) {
          for (const [key, invocation] of invocationContexts) { try { invocation.assertInvocationCurrent(); } catch { invocationContexts.delete(key); } }
          if (invocationContexts.size >= 1024) throw new Error("graph_account_context_capacity");
        }
        invocationContexts.set(contextKey(ctx), ctx);
        const tools = factory(ctx);
        const list = tools == null ? [] : Array.isArray(tools) ? tools : [tools];
        for (const tool of list) {
          const execute = tool.execute;
          tool.execute = async (...args) => {
            try {
              if (!accountStore) throw new Error("graph_account_store_unavailable");
              accountStore.finishApproval(args[0]);
              return await accountStore.run(hostTarget(ctx), () => execute(...args));
            } catch (error) { return failureResult(error, tool.name); }
          };
        }
        return Array.isArray(tools) ? list : list[0] ?? null;
      } }, options);
    };
  } });

  const hook = async <T extends HookResult>(event: HookEvent, ctx: HookContext, action: () => Promise<T>): Promise<T | { block: true; blockReason: string } | undefined> => {
    if (!event.toolName.startsWith("onedrive_") && !event.toolName.startsWith("outlook_") && !event.toolName.startsWith("microsoft_todo_") && event.toolName !== "microsoft_graph_capabilities") return;
    try {
      const invocation = invocationContexts.get(contextKey(ctx));
      if (!accountStore || !invocation) throw new Error("graph_session_authority_required");
      const target = hostTarget(invocation);
      const { scope, holdKey } = await accountStore.resolve(target);
      const release = accountStore.hold(holdKey);
      try {
        return await accountScope.run(scope, async () => {
          const response = await action();
          await scope.assertFresh(); scope.assertCurrent();
          if (response?.requireApproval && event.toolCallId) {
            accountStore.pendingApproval(event.toolCallId, holdKey);
            const resolution = response.requireApproval.onResolution;
            response.requireApproval.onResolution = async decision => {
              try { scope.assertCurrent(); await scope.assertFresh(); await accountScope.run(scope, async () => { await resolution?.(decision); }); }
              finally { if (decision !== "allow-once" && decision !== "allow-always") accountStore.finishApproval(event.toolCallId!); }
            };
          }
          return response;
        });
      } finally { release(); }
    } catch (error) { return { block: true, blockReason: code(error) }; }
  };
  const clearSession = (id: string) => { for (const [key, context] of invocationContexts) if (context.sessionId === id) invocationContexts.delete(key); };
  return { hostInstanceId: state.hostInstanceId, store: accountStore, registrationApi, config: scopedConfig, hook, clearSession };
}
