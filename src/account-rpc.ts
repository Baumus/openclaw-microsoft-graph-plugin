import type { DirectorySnapshot } from "./account-directory.js";
import type { OpenClawPluginApi } from "openclaw/plugin-sdk/core";
import { AccountStore, assertPolicyCeiling, SHARING_WARNING, type AccountHostContext, type SessionTarget } from "./accounts.js";
import { acquireVaultLock, assertMigrationRefreshClear, createVaultCredential, inspectVaultCredential, readMigrationVault } from "./credential-vault.js";
import { DeviceCodeSignIn } from "./device-code.js";
import { accountScope } from "./account-scope.js";
import { tokenForAuthorizedOperation } from "./credential.js";
import { graphRequest } from "./graph.js";
import { validatePolicy, type GraphPolicy } from "./policy.js";

/** Migration keeps only legacy grants that the destination profile ceiling still permits. */
function intersectMigrationPolicy(legacy: GraphPolicy, ceiling: GraphPolicy): GraphPolicy {
  const policy = structuredClone(legacy);
  for (const service of ["calendar", "mail", "todo"] as const) {
    const grants = policy.services[service].agents;
    for (const [agent, grant] of Object.entries(grants)) {
      const cap = ceiling.services[service].agents[agent];
      if (!cap) { delete grants[agent]; continue; }
      const operations = grant.operations.filter(operation => cap.operations.includes(operation));
      const resources = (grant.resources ?? ["me"]).filter(resource => (cap.resources ?? ["me"]).includes(resource));
      if (!operations.length || !resources.length) { delete grants[agent]; continue; }
      grant.operations = operations;
      if (grant.resources !== undefined) grant.resources = resources;
    }
  }
  policy.services.onedrive.allowed_roots = policy.services.onedrive.allowed_roots.flatMap(root => {
    const cap = ceiling.services.onedrive.allowed_roots.find(candidate => candidate.label === root.label);
    if (!cap || root.path !== cap.path || root.drive_id !== cap.drive_id || root.item_id !== cap.item_id || root.include_descendants !== cap.include_descendants || root.agents_instructions !== cap.agents_instructions) return [];
    for (const operation of ["read", "write", "delete"] as const) root.permissions[operation] = root.permissions[operation] && cap.permissions[operation];
    for (const [agent, grant] of Object.entries(root.agents)) {
      const capGrant = cap.agents[agent];
      for (const operation of ["read", "write", "delete"] as const) if (grant.permissions[operation]) grant.permissions[operation] = !!(root.permissions[operation] && capGrant?.permissions[operation]);
      if (!(["read", "write", "delete"] as const).some(operation => grant.permissions[operation] === true)) delete root.agents[agent];
    }
    return Object.keys(root.agents).length ? [root] : [];
  });
  for (const service of ["onedrive", "calendar", "mail", "todo"] as const) {
    const former = legacy.rules.warningApprovalsByService?.[service];
    const cap = ceiling.rules.warningApprovalsByService?.[service];
    if (former === false && cap !== false) {
      if (cap === true) (policy.rules.warningApprovalsByService ??= {})[service] = true;
      else delete policy.rules.warningApprovalsByService?.[service];
    } else if (former === undefined && cap === true) (policy.rules.warningApprovalsByService ??= {})[service] = true;
  }
  validatePolicy(policy);
  assertPolicyCeiling(policy, legacy);
  assertPolicyCeiling(policy, ceiling);
  return policy;
}

export function registerAccountMethods(api: OpenClawPluginApi, store: AccountStore, config: { credentialVaultKey?: unknown }): void {
  type Context = Parameters<Parameters<OpenClawPluginApi["registerGatewayMethod"]>[1]>[0];
  const methods = new Map<string, DeviceCodeSignIn>();
  const profile = (ctx: Context): string => {
    const id = ctx.client?.authenticatedUserProfile?.profileId;
    if (typeof id !== "string" || !id || !ctx.hasCurrentClientAuthority?.()) throw new Error("graph_authenticated_profile_required");
    return id;
  };
  const guard = (ctx: Context, id: string): (() => void) => () => {
    ctx.signal?.throwIfAborted();
    if (!ctx.hasCurrentClientAuthority?.() || ctx.client?.authenticatedUserProfile?.profileId !== id) throw new Error("graph_authenticated_profile_required");
  };
  const exact = (params: Record<string, unknown>, keys: string[]) => { if (Object.keys(params).length !== keys.length || keys.some(k => !(k in params))) throw new Error("invalid_rpc_parameters"); };
  const register = (name: string, keys: string[], action: (ctx: Context, id: string, current: () => void) => Promise<unknown>, session = false, admin = false) => {
    api.registerGatewayMethod(`microsoft-graph.accounts.${name}`, async ctx => {
      try {
        exact(ctx.params, keys); const id = profile(ctx); const current = guard(ctx, id); current();
        const value = await action(ctx, id, current); current(); ctx.respond(true, value);
      } catch (error) {
        const raw = error instanceof Error ? error.message : "internal_error";
        const code = /^(graph_|access_denied|invalid_|credential_|device_)/.test(raw) ? raw : "internal_error";
        ctx.respond(false, undefined, { code: "INVALID_REQUEST", message: code });
      }
    }, { scope: admin ? "operator.admin" : session ? "operator.write" : "operator.read", profileAccess: "required", ...(session ? { sessionAccess: { mode: "write", requiredTool: "microsoft_graph_capabilities" } as const } : {}) });
  };
  register("status", [], async (_ctx, id) => {
    const a = await store.readAccount(id);
    const automatic = await store.automaticStatus();
    return { profileId: id, account: a ?? null, ceiling: store.ceilingFor(id), workers: Object.entries(store.workers).filter(([, owner]) => owner === id).map(([agent]) => agent), warning: SHARING_WARNING, automatic: automatic?.profileId === id ? automatic : undefined, credential: a ? await inspectVaultCredential(store.accountPath(id, a.generation), config.credentialVaultKey as string) : { result: "missing" } };
  });
  register("provision", [], async (_ctx, id, current) => store.provision(id, current));
  register("resolve-folder", ["path"], async (ctx, id, current) => {
    const a = await store.readAccount(id);
    const raw = ctx.params.path;
    if (!a?.connected || typeof raw !== "string" || raw.length > 1024 || !raw.startsWith("/") || raw.includes("\\") || /[\u0000-\u001f\u007f]/.test(raw)) throw new Error("invalid_path");
    const segments = raw.split("/").filter(Boolean);
    if (segments.some(s => s === "." || s === ".." || s.length > 255)) throw new Error("invalid_path");
    const signature = JSON.stringify(a);
    const fresh = async () => { current(); if (JSON.stringify(await store.readAccount(id)) !== signature) throw new Error("graph_account_context_changed"); current(); };
    return accountScope.run({ identity: `management:${id}:${a.generation}:${a.revision}`, stateDir: store.accountPath(id, a.generation), policy: a.policy, assertCurrent: current, assertFresh: fresh }, async () => {
      await fresh();
      const token = await tokenForAuthorizedOperation({ config, policy: a.policy, allowedScopes: ["Files.Read"], stateDir: store.accountPath(id, a.generation), signal: AbortSignal.timeout(15000) });
      const selector = segments.length ? `/me/drive/root:/${segments.map(encodeURIComponent).join("/")}:` : "/me/drive/root";
      const item = await graphRequest(token, `${selector}?$select=id,name,folder,parentReference`, { maxBytes: 16384, signal: AbortSignal.timeout(15000) });
      if (typeof item?.id !== "string" || !item.id || !item.folder || typeof item?.parentReference?.driveId !== "string") throw new Error("invalid_folder");
      await fresh(); return { path: "/" + segments.join("/"), drive_id: item.parentReference.driveId, item_id: item.id };
    });
  }, false, true);

  register("configure", ["expectedRevision", "policy", "shared", "unattended"], async (ctx, id, current) => {
    const p = ctx.params;
    if (!Number.isSafeInteger(p.expectedRevision) || typeof p.shared !== "boolean" || typeof p.unattended !== "boolean") throw new Error("invalid_rpc_parameters");
    return store.update(id, p.expectedRevision as number, validatePolicy(p.policy), p.shared, p.unattended, current);
  });
  register("disconnect", ["expectedRevision"], async (ctx, id, current) => {
    if (!Number.isSafeInteger(ctx.params.expectedRevision)) throw new Error("invalid_rpc_parameters");
    const a = await store.readAccount(id); if (a) { const d = methods.get(`${id}:${a.generation}`); if (d) d.stop(); }
    await store.disconnect(id, ctx.params.expectedRevision as number, current); return { disconnected: true };
  });
  register("device-start", ["clientId", "tenant"], async (ctx, id, current) => {
    const a = await store.readAccount(id); if (!a || a.connected) throw new Error("graph_account_conflict");
    if (typeof ctx.params.clientId !== "string" || typeof ctx.params.tenant !== "string") throw new Error("invalid_rpc_parameters");
    const key = `${id}:${a.generation}`;
    let device = methods.get(key);
    if (!device) {
      const fresh = async () => { current(); const latest = await store.readAccount(id); current(); if (!latest || latest.generation !== a.generation || latest.revision !== a.revision) throw new Error("graph_account_context_changed"); };
      device = new DeviceCodeSignIn({ credentialVaultKey: config.credentialVaultKey, policy: a.policy }, () => store.accountPath(id, a.generation), fetch, undefined, fresh, current);
      methods.set(key, device);
    }
    await device.checkAuthority(); return device.start(ctx.params.clientId, ctx.params.tenant);
  });
  register("device-status", ["sessionId"], async (ctx, id, current) => {
    const a = await store.readAccount(id); if (!a || typeof ctx.params.sessionId !== "string") throw new Error("device_session_unavailable");
    const device = methods.get(`${id}:${a.generation}`); if (!device) throw new Error("device_session_unavailable");
    const status = device.status(ctx.params.sessionId);
    if (status.state === "created" && !a.connected) await store.markConnected(id, a.generation, current);
    return status;
  });
  register("device-cancel", ["sessionId"], async (ctx, id) => {
    const a = await store.readAccount(id); const device = a && methods.get(`${id}:${a.generation}`);
    if (!device || typeof ctx.params.sessionId !== "string") throw new Error("device_session_unavailable"); return device.cancel(ctx.params.sessionId);
  });
  const target = (ctx: Context, current: () => void): AccountHostContext => {
    const p = ctx.params; const authority = ctx.sessionAccessAuthority;
    if (!authority?.target.lifecycleRevision || !ctx.withSessionTurnAuthority || typeof p.agentId !== "string" || typeof p.sessionKey !== "string" || typeof p.sessionId !== "string") throw new Error("graph_session_authority_required");
    const exact: SessionTarget = { agentId: p.agentId, sessionKey: p.sessionKey, sessionId: p.sessionId, lifecycleRevision: authority.target.lifecycleRevision };
    if (authority.target.agentId !== exact.agentId || authority.target.sessionKey !== exact.sessionKey || authority.target.sessionId !== exact.sessionId) throw new Error("graph_account_conflict");
    return { ...exact, assertCurrent: () => { current(); authority.assertCurrent(); } };
  };
  register("session-status", ["agentId", "sessionKey", "sessionId"], async (ctx, id, current) => store.status(target(ctx, current), id), true);
  register("bind", ["agentId", "sessionKey", "sessionId", "profileId", "expectedRevision"], async (ctx, id, current) => {
    if (typeof ctx.params.profileId !== "string" || !(ctx.params.expectedRevision === null || Number.isSafeInteger(ctx.params.expectedRevision))) throw new Error("invalid_rpc_parameters");
    const t = target(ctx, current);
    // Host admission gives a synchronous exact-turn guard; store publication rechecks its captured authority.
    await ctx.withSessionTurnAuthority!(t, () => {
      t.assertCurrent();
    });
    return store.bind(id, ctx.params.profileId, t, ctx.params.expectedRevision as number | null);
  }, true);
  register("automatic-migration", [], async (_ctx, id, current) => {
    // Host's operator.admin admission plus the fresh authenticated exact profile;
    // directory read authority is never mistaken for the person's authority.
    const existing = await store.automaticStatus();
    if (existing) { const automatic = await store.checkAutomatic(); if (automatic?.profileId !== id) throw new Error("graph_account_conflict"); return { automatic: true, migrated: true }; }
    const snapshot = await store.directory?.read(); current();
    if (!snapshot) throw new Error("graph_single_user_directory_unavailable");
    if (snapshot.soleProfileId !== id || id === "gateway-owner") return { automatic: false, reason: "not_authenticated_sole_user" };
    const legacy = await inspectVaultCredential(store.root, config.credentialVaultKey as string);
    current(); if (legacy.result === "missing") return { automatic: false, reason: "no_legacy_connection" };
    await migrateLegacyAccount(store, config, id, current, snapshot);
    return { automatic: true, migrated: true, retainedLegacyBackup: true };
  }, false, true);
  // Explicit self migration remains available on genuine multi-user hosts.
  register("migrate-legacy", [], async (_ctx, id, current) => migrateLegacyAccount(store, config, id, current), false, true);

}

/** The only migration writer; both paths retain the unchanged encrypted legacy vault. */
export async function migrateLegacyAccount(store: AccountStore, config: { credentialVaultKey?: unknown }, id: string, current: () => void, snapshot?: DirectorySnapshot) {
  const fresh = async () => {
    current();
    if (snapshot) {
      const latest = await store.directory?.read();
      // Record an actually observed mismatch even when the actor is cancelled
      // while the read finishes. Cancellation must never erase a safety fact.
      if (!latest || latest.soleProfileId !== id || latest.digest !== snapshot.digest) throw new Error("graph_single_user_topology_changed");
      current();
    }
  };
  const a = await store.provision(id, current);
  const path = store.accountPath(id, a.generation);
  const sourceLock = await acquireVaultLock(store.root, "migrate-source");
  try {
    current();
    const destinationLock = await acquireVaultLock(path, "migrate-destination");
    try {
      current();
      let sourceDigest: string | undefined; let destinationDigest: string | undefined;
      const verifyVaults = async () => {
        current();
        const source = await readMigrationVault(store.root, config.credentialVaultKey, sourceLock);
        const destination = await readMigrationVault(path, config.credentialVaultKey, destinationLock);
        if (source.digest !== sourceDigest || destination.digest !== destinationDigest) throw new Error("credential_vault_conflict");
        current();
      };
      return await store.mutate(async r => {
        const account = Object.values(r.accounts).find(x => x.profileId === id)!;
        if (account.generation !== a.generation) throw new Error("graph_account_conflict");
        if (snapshot && r.automatic?.invalidated) throw new Error("graph_single_user_topology_changed");
        if (snapshot && r.legacyMigration && !r.automatic) throw new Error("graph_account_conflict");
        await fresh();
        const legacy = await readMigrationVault(store.root, config.credentialVaultKey, sourceLock);
        sourceDigest = legacy.digest;
        if (r.legacyMigration) {
          if (r.legacyMigration.profileId !== id || r.legacyMigration.generation !== account.generation || r.legacyMigration.sourceDigest !== legacy.digest || !account.connected) throw new Error("graph_account_conflict");
          const destination = await readMigrationVault(path, config.credentialVaultKey, destinationLock);
          destinationDigest = destination.digest;
          await fresh();
          return { migrated: true, profileId: id, generation: account.generation, retainedLegacyBackup: true };
        }
        if (account.connected) throw new Error("graph_account_conflict");
        await fresh();
        await assertMigrationRefreshClear(path, config.credentialVaultKey, destinationLock);
        const status = await inspectVaultCredential(path, config.credentialVaultKey as string);
        current();
        if (status.result === "missing") {
          if (!await sourceLock.verifyStillHeld() || !await destinationLock.verifyStillHeld()) throw new Error("credential_vault_locked");
          current(); await createVaultCredential(path, legacy.credential, config.credentialVaultKey, undefined, current);
        } else if (status.result !== "valid") throw new Error("credential_reauthorization_required");
        const destination = await readMigrationVault(path, config.credentialVaultKey, destinationLock);
        if (JSON.stringify(destination.credential) !== JSON.stringify(legacy.credential)) throw new Error("graph_account_conflict");
        destinationDigest = destination.digest;
        await fresh();
        if (snapshot) r.automatic = { profileId: id, generation: account.generation, topologyDigest: snapshot.digest, invalidated: false };
        current(); r.legacyMigration = { profileId: id, generation: account.generation, sourceDigest: legacy.digest }; account.policy = intersectMigrationPolicy(store.ceiling, store.ceilingFor(id)); account.connected = true; account.shared = false; account.unattended = false; account.revision++;
        return { migrated: true, profileId: id, generation: a.generation, retainedLegacyBackup: true };
      }, current, { beforePublication: verifyVaults, ...(snapshot ? { automaticAttempt: { profileId: id, generation: a.generation, topologyDigest: snapshot.digest } } : {}) });
    } finally { await destinationLock.release().catch(() => undefined); }
  } finally { await sourceLock.release().catch(() => undefined); }
}
