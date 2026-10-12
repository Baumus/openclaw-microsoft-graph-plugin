import type { OpenClawPluginApi, PluginServiceSchedulerV1 } from "openclaw/plugin-sdk/core";
import { resolve } from "node:path";
import type { AccountStore } from "./accounts.js";
import { migrateLegacyAccount } from "./account-rpc.js";
import { inspectVaultCredential } from "./credential-vault.js";
import { normalizedCriteria } from "./continuation.js";
import type { RuntimeConfig } from "./index.js";

/** Trusted host startup, never a fabricated personal RPC identity. */
export function registerAccountBootstrap(api: OpenClawPluginApi, store: AccountStore, config: RuntimeConfig): void {
  if (api.registrationMode !== "full" || config.enabled !== true || typeof api.registerService !== "function") return;
  const root = resolve(store.root);
  const key = Symbol.for("@baumus/openclaw-microsoft-graph/bootstrap-leases/v1");
  const slots = globalThis as Record<symbol, unknown>;
  const leases = (slots[key] ??= new Map<string, { close: () => void }>()) as Map<string, { close: () => void }>;
  // Replacement with byte-identical configuration still retires old admission.
  leases.get(root)?.close();
  let scheduler: PluginServiceSchedulerV1 | undefined;
  let controller = new AbortController();
  let started = false;
  let retired = false;
  let readySignal: AbortSignal | undefined;
  let epoch = 0;
  let removeAbort: (() => void) | undefined;
  const close = () => { controller.abort(); scheduler?.beginClose(); removeAbort?.(); removeAbort = undefined; };
  const lease = { close: () => { retired = true; readySignal = undefined; close(); } }; leases.set(root, lease);
  const gatewayIdentity = (gateway: ReturnType<typeof api.runtime.config.current>["gateway"]) => normalizedCriteria([gateway?.mode, gateway?.port, gateway?.bind, gateway?.tls, gateway?.auth]);
  const relevant = () => {
    const current = api.runtime.config.current();
    return normalizedCriteria([config, gatewayIdentity(current.gateway), current.plugins?.enabled, current.plugins?.entries?.["microsoft-graph"], process.env.OPENCLAW_GATEWAY_PORT, process.env.OPENCLAW_GATEWAY_TOKEN]);
  };
  const baseline = relevant();
  const current = () => {
    const host = api.runtime.config.current();
    if (retired || leases.get(root) !== lease || controller.signal.aborted || !scheduler || scheduler.signal.aborted || relevant() !== baseline || config.enabled !== true || host.plugins?.enabled === false || host.plugins?.entries?.["microsoft-graph"]?.enabled === false || resolve(api.runtime.state.resolveStateDir(process.env)) !== root) throw new Error("graph_account_context_changed");
  };
  api.registerService({ apiVersion: 2, id: "microsoft-graph-single-user-bootstrap", reload: { configPrefixes: ["plugins.entries.microsoft-graph", "gateway.mode", "gateway.port", "gateway.bind", "gateway.tls", "gateway.auth"] },
    start(ctx) {
      if (retired || leases.get(root) !== lease) return;
      close(); epoch++; controller = new AbortController(); scheduler = ctx.scheduler.scope(); started = false;
      // Core can restart this retained service without replaying gateway_start.
      // Only a previously published, still-live host lifetime permits recovery.
      scheduleReady();
    },
    async stop() { close(); await scheduler?.stop(); },
  });
  api.on("gateway_stop", () => { readySignal = undefined; close(); });
  api.on("gateway_start", (_event, ctx) => {
    if (!ctx.abortSignal || ctx.abortSignal.aborted) return;
    try { current(); if (gatewayIdentity(ctx.config?.gateway) !== gatewayIdentity(api.runtime.config.current().gateway)) return; } catch { return; }
    readySignal = ctx.abortSignal;
    scheduleReady();
  });
  function scheduleReady() {
    if (started || !scheduler || !readySignal || readySignal.aborted) return;
    try { current(); } catch { return; }
    started = true;
    const admittedEpoch = epoch, admittedController = controller, admittedScheduler = scheduler;
    const admittedCurrent = () => { if (epoch !== admittedEpoch) throw new Error("graph_account_context_changed"); current(); };
    const hostSignal = readySignal;
    const abort = () => { if (epoch === admittedEpoch) { readySignal = undefined; close(); } }; hostSignal.addEventListener("abort", abort, { once: true });
    removeAbort = () => hostSignal.removeEventListener("abort", abort);
    let attempts = 0;
    const run = async () => {
      try {
        admittedCurrent();
        const existing = await store.automaticStatus(); admittedCurrent();
        // Never reset a sticky revocation, even when the user returns to singleton.
        if (existing?.invalidated) return;
        const legacy = await inspectVaultCredential(root, config.credentialVaultKey as string); admittedCurrent();
        if (legacy.result === "missing" && !existing) return;
        const snapshot = await store.directory?.read(admittedController.signal);
        // Preserve observed safety facts independently of positive startup authority.
        if (existing) { await store.checkAutomatic(snapshot); admittedCurrent(); return; }
        admittedCurrent();
        if (!snapshot) throw new Error("graph_single_user_directory_unavailable");
        if (!snapshot.soleProfileId || snapshot.soleProfileId === "gateway-owner") return;
        if (legacy.result !== "valid") throw new Error("graph_account_conflict");
        await migrateLegacyAccount(store, config, snapshot.soleProfileId, admittedCurrent, snapshot, { signal: admittedController.signal, requireEmptyRegistry: true });
      } catch (error) {
        // No raw directory, stderr, tokens, IDs or paths reach logs.
        if (error instanceof Error && error.message === "graph_single_user_directory_unavailable" && ++attempts < 3 && !admittedController.signal.aborted && !admittedScheduler.signal.aborted) {
          try { admittedCurrent(); admittedScheduler.schedule({ id: "singleton-bootstrap", delayMs: 2000, run }); } catch { /* retired generation */ }
        } else if (!admittedController.signal.aborted && !admittedScheduler.signal.aborted) api.logger.warn("Microsoft Graph automatic startup setup unavailable; explicit personal setup remains available.");
      }
    };
    scheduler.schedule({ id: "singleton-bootstrap", delayMs: 0, run });
  }
}
