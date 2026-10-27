import { mkdtemp, readFile, rm, stat, symlink, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { tmpdir } from "node:os";
import { createHash } from "node:crypto";
import { afterEach, describe, expect, it } from "vitest";
import { ActivationJournal } from "./config-activation.js";
import { canonicalPolicy } from "./config-policy-identity.js";
const initial = () => ({ version: 2, rules: { default: "deny" }, services: { onedrive: { allowed_roots: [] }, calendar: { agents: { selected: { operations: ["read", "create"], resources: ["old-resource-private"] } } }, mail: { agents: {} }, todo: { agents: {} } } });
const hash = (p: unknown) => createHash("sha256").update(canonicalPolicy(p)).digest("hex");
const dirs: string[] = [];
afterEach(async () => { for (const dir of dirs.splice(0)) await rm(dir, { recursive: true, force: true }); });
async function setup() {
  const dir = await mkdtemp(join(tmpdir(), "graph-activation-test-")); dirs.push(dir);
  const before = initial(); const after = structuredClone(before); after.services.calendar.agents.selected = { operations: ["read"], resources: ["new-resource-private"] };
  const entries: Record<string, { tools: { alsoAllow: string[]; deny: string[] } }> = { selected: { tools: { alsoAllow: ["outlook_mail_read"], deny: ["outlook_mail_write"] } }, unrelated: { tools: { alsoAllow: [], deny: [] } } };
  const snapshot = { hash: "before", parsed: { plugins: { $include: "private-plugins-path" } }, config: { tools: { deny: ["onedrive_delete"] }, agents: { entries }, plugins: { entries: { "microsoft-graph": { config: { policy: structuredClone(before) } } } } } };
  const journal = () => new ActivationJournal(dir, async () => structuredClone(snapshot), "synthetic-operator");
  const persist = () => { snapshot.config.plugins.entries["microsoft-graph"].config.policy = structuredClone(after); snapshot.hash="after"; };
  const file = join(dir, "plugin-data", "microsoft-graph", "activation", "pending.json");
  return { dir, before, after, snapshot, entries, journal, persist, file };
}
describe("durable exact-target activation journal", () => {
  it("survives fresh instance, gates on exact active policy and claims only narrowed selected tools once", async () => {
    const h=await setup(); const intent=await h.journal().prepare("before", h.before, h.after);
    expect(await h.journal().status()).toBeNull(); h.persist();
    expect((await h.journal().status())?.id).toBe(intent.id);
    await expect(h.journal().claim(intent.id, hash(h.before))).rejects.toThrow("activation_conflict");
    const claimed=await h.journal().claim(intent.id, hash(h.after));
    expect(claimed.baseHash).toBe("after"); expect(claimed.patch).toEqual({ selected: { tools: { alsoAllow: ["outlook_mail_read", "outlook_calendar_read"] } } });
    await expect(h.journal().claim(intent.id, hash(h.after))).rejects.toThrow("activation_not_claimable");
    expect(h.entries.unrelated!.tools.alsoAllow).toEqual([]); expect(h.entries.selected!.tools.deny).toEqual(["outlook_mail_write"]);
    const raw=await readFile(h.file,"utf8"); expect(raw).not.toMatch(/resource-private|private-plugins-path|credential|baseline|policy/);
    expect((await stat(h.file)).mode & 0o777).toBe(0o600); expect((await stat(join(h.file,".."))).mode & 0o777).toBe(0o700);
  });
  it("does not resurrect a completed grant after a later revocation", async () => {
    const h=await setup(); const intent=await h.journal().prepare("before",h.before,h.after); h.persist();
    const c=await h.journal().claim(intent.id,hash(h.after)); h.entries.selected!.tools.alsoAllow=c.expected.selected!.value;
    await h.journal().complete(intent.id,hash(h.after)); h.entries.selected!.tools.alsoAllow=["outlook_mail_read"];
    expect(await h.journal().status()).toBeNull(); await expect(h.journal().claim(intent.id,hash(h.after))).rejects.toThrow();
  });
  it("rejects stale tool policy and denies without changing configuration", async () => {
    const h=await setup(); const i=await h.journal().prepare("before",h.before,h.after); h.persist(); h.entries.selected!.tools.deny.push("outlook_calendar_*");
    await expect(h.journal().claim(i.id,hash(h.after))).rejects.toThrow("activation_conflict"); expect(h.entries.selected!.tools.alsoAllow).toEqual(["outlook_mail_read"]);
  });
  it("permits unrelated agent changes but refuses owner/source changes", async () => {
    const h=await setup(); const i=await h.journal().prepare("before",h.before,h.after); h.persist(); h.entries.unrelated!.tools.alsoAllow.push("unrelated");
    h.snapshot.parsed.plugins.$include="other"; expect(await h.journal().status()).toBeNull(); await expect(h.journal().claim(i.id,hash(h.after))).rejects.toThrow();
    h.snapshot.parsed.plugins.$include="private-plugins-path"; expect((await h.journal().claim(i.id,hash(h.after))).patch).toHaveProperty("selected");
  });
  it("rejects wrong policy, stale prepare, expiry and cancelled intents", async () => {
    const h=await setup(); await expect(h.journal().prepare("stale",h.before,h.after)).rejects.toThrow();
    const i=await h.journal().prepare("before",h.before,h.after); h.persist();
    h.snapshot.config.plugins.entries["microsoft-graph"].config.policy.services.calendar.agents.selected.resources=["external"];
    expect(await h.journal().status()).toBeNull(); await expect(h.journal().claim(i.id,hash(h.after))).rejects.toThrow(); h.persist();
    await writeFile(h.file,JSON.stringify({...i,createdAt:Date.now()-25*3600000}),{mode:0o600}); expect(await h.journal().status()).toBeNull(); await expect(h.journal().claim(i.id,hash(h.after))).rejects.toThrow();
    await h.journal().cancel(i.id); expect(await h.journal().status()).toBeNull();
  });
  it("serializes concurrent claims and leaves a lost dispatch ambiguous, never replayable", async () => {
    const h=await setup(); const i=await h.journal().prepare("before",h.before,h.after); h.persist();
    const claims=await Promise.allSettled([h.journal().claim(i.id,hash(h.after)),h.journal().claim(i.id,hash(h.after))]);
    expect(claims.filter(r=>r.status==="fulfilled")).toHaveLength(1); expect((await h.journal().status())?.state).toBe("issued");
    await expect(h.journal().complete(i.id,hash(h.after))).rejects.toThrow(); await expect(h.journal().claim(i.id,hash(h.after))).rejects.toThrow();
  });
  it("binds intent to the authenticated operator, not browser-supplied identity", async () => {
    const h=await setup(); const i=await h.journal().prepare("before",h.before,h.after); h.persist();
    const other=new ActivationJournal(h.dir,async()=>structuredClone(h.snapshot),"other-operator");
    expect(await other.status()).toBeNull(); await expect(other.claim(i.id,hash(h.after))).rejects.toThrow(); await expect(other.cancel(i.id)).rejects.toThrow();
    await expect(other.prepare("after",h.after,h.before)).rejects.toThrow("activation_owner_conflict");
  });
  it("rejects a symlink journal and never follows it", async () => {
    const h=await setup(); await h.journal().prepare("before",h.before,h.after); await rm(h.file);
    const other=join(h.dir,"other"); await writeFile(other,"null",{mode:0o600}); await symlink(other,h.file);
    await expect(h.journal().prepare("before",h.before,h.after)).rejects.toThrow(); expect(await readFile(other,"utf8")).toBe("null");
  });
});
