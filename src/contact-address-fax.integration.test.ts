import { afterEach, describe, expect, it, vi } from "vitest";

const credentials = vi.hoisted(() => ({ tokenForAuthorizedOperation: vi.fn(async () => "synthetic-token") }));
vi.mock("./credential.js", async () => ({ ...await vi.importActual<typeof import("./credential.js")>("./credential.js"), ...credentials }));
import entry from "./index.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";

function harness() {
  const factories: Array<(context: any) => any> = [];
  const hooks: Record<string, (...args: any[]) => any> = {};
  const policy = graphPolicyFixture();
  policy.services.contacts = { agents: { main: { operations: ["read", "create", "update", "delete"], resources: ["me"] } } };
  entry.register({ pluginConfig: { enabled: true, policy }, registerTool: (factory: any) => factories.push(factory), on: (name: string, handler: any) => { hooks[name] = handler; }, logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() } } as any);
  const context = { agentId: "main", sessionId: "contact-address-fax" };
  const tools = Object.fromEntries(factories.map(factory => { const tool = factory(context); return [tool.name, tool]; }));
  let counter = 0;
  return { tools, hooks, context, async execute(name: string, params: Record<string, unknown>) {
    const id = `contact-${++counter}`;
    const gate = await hooks.before_tool_call({ toolName: name, toolCallId: id, params }, context);
    if (gate?.block) return { ok: false, error: gate.blockReason };
    await gate?.requireApproval?.onResolution("allow-once");
    return (await tools[name].execute(id, gate?.params ?? params)).details;
  } };
}
const address = { street: "1 Example Street", city: "Berlin", state: "BE", postalCode: "10115", countryOrRegion: "Germany" };
const businessFax = { id: "String 0x3A24", value: "+49 30 12345" };
const homeFax = { id: "String 0x3A25", value: "+49 30 98765" };
const primaryFax = { id: "String 0x3A23", value: "+49 30 11111" };
const unrelated = { id: "String 0x3A00", value: "do-not-expose" };
afterEach(() => { vi.restoreAllMocks(); credentials.tokenForAuthorizedOperation.mockClear(); });

describe("Outlook contact postal addresses and fixed native fax properties", () => {
  it("returns addresses and normalized fax in list/get/search without raw or unrelated extended properties", async () => {
    const contact = { id: "A=B", displayName: "Ada", businessAddress: { ...address, providerOnly: "hidden" }, homeAddress: address, otherAddress: address, singleValueExtendedProperties: [unrelated, primaryFax, homeFax, { ...businessFax, id: "String 0x3a24" }], providerOnly: "hidden" };
    const fetch = vi.spyOn(globalThis, "fetch").mockImplementation(async input => new URL(String(input)).pathname.endsWith("/contacts/A%3DB") ? Response.json(contact) : Response.json({ value: [contact] }));
    const h = harness();
    for (const params of [{ action: "list_contacts" }, { action: "get_contact", contactId: "A=B" }, { action: "search_contacts", search: "10115" }, { action: "search_contacts", search: "98765" }]) {
      const result = await h.execute("outlook_contacts_read", params);
      expect(result.ok).toBe(true);
      const item = result.item ?? result.items[0];
      expect(item).toMatchObject({ businessAddress: address, homeAddress: address, otherAddress: address, businessFaxNumber: businessFax.value, homeFaxNumber: homeFax.value, primaryFaxNumber: primaryFax.value });
      expect(JSON.stringify(item)).not.toContain("providerOnly");
      expect(JSON.stringify(item)).not.toContain("do-not-expose");
      expect(item).not.toHaveProperty("singleValueExtendedProperties");
    }
    for (const [input] of fetch.mock.calls) {
      const url = new URL(String(input));
      expect(url.pathname).toMatch(/^\/v1\.0\/me\/contacts(?:\/A%3DB)?$/);
      expect(url.searchParams.get("$select")).toContain("businessAddress,homeAddress,otherAddress");
      expect(url.searchParams.get("$select")).not.toContain("Fax");
      expect(url.searchParams.get("$expand")).toBe("singleValueExtendedProperties($filter=id eq 'String 0x3A24' or id eq 'String 0x3A25' or id eq 'String 0x3A23')");
    }
  });

  it.each(["create", "update"])("%s maps typed fields to native provider properties and reads back normalized fax", async action => {
    let saved: any = { id: "A=B", displayName: "Ada", businessPhones: ["unchanged-voice"], singleValueExtendedProperties: [homeFax, unrelated] };
    const writes: any[] = [];
    const fetch = vi.spyOn(globalThis, "fetch").mockImplementation(async (_input, init) => {
      if (init?.method === "POST" || init?.method === "PATCH") {
        const body = JSON.parse(String(init.body)); writes.push(body);
        saved = { ...saved, ...body, singleValueExtendedProperties: [homeFax, unrelated, { id: "String 0x3a24", value: "+49 (30) 12345" }] };
        return Response.json({ id: "A=B", displayName: "Ada", businessAddress: address }); // Graph omits extended properties here.
      }
      return Response.json(saved);
    });
    const result = await harness().execute("outlook_contacts_write", { action, ...(action === "update" ? { contactId: "A=B" } : { displayName: "Ada" }), businessAddress: address, businessFaxNumber: "+493012345" });
    expect(result).toMatchObject({ ok: true, mutationApplied: true, verification: "read_back", item: { businessAddress: address, businessFaxNumber: "+49 (30) 12345", homeFaxNumber: homeFax.value, businessPhones: ["unchanged-voice"] } });
    expect(writes).toEqual([{ ...(action === "create" ? { displayName: "Ada" } : {}), businessAddress: address, singleValueExtendedProperties: [{ id: "String 0x3A24", value: "+493012345" }] }]);
    expect(fetch).toHaveBeenCalledTimes(2);
    expect(new URL(String(fetch.mock.calls[1][0])).pathname).toBe("/v1.0/me/contacts/A%3DB");
    expect(saved.singleValueExtendedProperties).toContainEqual(unrelated);
    expect(result.item).not.toHaveProperty("singleValueExtendedProperties");
  });

  it("omits unchanged address/fax fields and clears only explicitly selected properties", async () => {
    const writes: any[] = [];
    vi.spyOn(globalThis, "fetch").mockImplementation(async (_input, init) => {
      if (init?.method === "PATCH") writes.push(JSON.parse(String(init.body)));
      return Response.json({ id: "contact", businessAddress: { ...address, city: "" }, singleValueExtendedProperties: [] });
    });
    const h = harness();
    expect(await h.execute("outlook_contacts_write", { action: "update", contactId: "contact", companyName: "New company" })).toMatchObject({ ok: true });
    expect(writes[0]).toEqual({ companyName: "New company" });
    expect(await h.execute("outlook_contacts_write", { action: "update", contactId: "contact", homeAddress: null, businessFaxNumber: "", primaryFaxNumber: "" })).toMatchObject({ ok: true, item: { businessFaxNumber: "", primaryFaxNumber: "" } });
    expect(writes[1]).toEqual({ homeAddress: { street: "", city: "", state: "", postalCode: "", countryOrRegion: "" }, singleValueExtendedProperties: [{ id: businessFax.id, value: "" }, { id: primaryFax.id, value: "" }] });
    expect(writes[1]).not.toHaveProperty("businessAddress");
    expect(writes[1]).not.toHaveProperty("homeFaxNumber");
  });

  it("supports full replacements of all address types and strips nested provider-only data", async () => {
    const fetch = vi.spyOn(globalThis, "fetch").mockImplementation(async (_input, init) => Response.json({ id: "contact", ...JSON.parse(String(init?.body)), otherAddress: { ...address, privateProviderField: "hidden" } }));
    const result = await harness().execute("outlook_contacts_write", { action: "update", contactId: "contact", businessAddress: address, homeAddress: address, otherAddress: address });
    expect(JSON.parse(String(fetch.mock.calls[0][1]?.body))).toEqual({ businessAddress: address, homeAddress: address, otherAddress: address });
    expect(result).toMatchObject({ ok: true, item: { businessAddress: address, homeAddress: address, otherAddress: address } });
    expect(result.item.otherAddress).not.toHaveProperty("privateProviderField");
    expect(fetch).toHaveBeenCalledTimes(1);
  });

  it.each([
    { id: "contact" },
    { id: "different-contact", singleValueExtendedProperties: [homeFax] },
  ])("does not manufacture fax-clear success from incomplete or wrong-target readback: %j", async saved => {
    vi.spyOn(globalThis, "fetch").mockImplementation(async (_input, init) => Response.json(init?.method === "PATCH" ? { id: "contact" } : saved));
    const result = await harness().execute("outlook_contacts_write", { action: "update", contactId: "contact", homeFaxNumber: homeFax.value });
    expect(result).toMatchObject({ ok: true, mutationApplied: true, appliedButUnverified: true, verification: "failed" });
    expect(result.item).not.toHaveProperty("homeFaxNumber");
  });

  it.each([
    [], [businessFax], [unrelated],
    [{ ...homeFax, value: "" }], [{ ...homeFax, value: "   " }],
    [{ ...homeFax, value: null }], [{ ...homeFax, value: 42 }],
    [{ ...homeFax, value: "a".repeat(101) }], [{ ...homeFax, value: "+49\n30" }],
    [homeFax, { ...homeFax, id: "String 0x3a25" }],
  ].map(properties => ({ properties })))("requires nonempty requested fax evidence, rejecting missing, invalid or duplicated properties: %j", async ({ properties }) => {
    const fetch = vi.spyOn(globalThis, "fetch").mockImplementation(async (_input, init) => Response.json(init?.method === "PATCH" ? { id: "contact" } : { id: "contact", singleValueExtendedProperties: properties }));
    const result = await harness().execute("outlook_contacts_write", { action: "update", contactId: "contact", homeFaxNumber: homeFax.value });
    expect(result).toMatchObject({ ok: true, mutationApplied: true, phase: "partial", appliedButUnverified: true, verification: "failed", outcome: "applied_with_warning", retrySafety: "readback_before_retry" });
    expect(result.item).not.toHaveProperty("homeFaxNumber");
    expect(fetch.mock.calls.filter(([, init]) => init?.method === "PATCH")).toHaveLength(1);
  });

  it.each([[], [{ ...homeFax, value: "" }], [unrelated]].map(properties => ({ properties })))("verifies explicit clear when the target property is absent or empty: %j", async ({ properties }) => {
    const fetch = vi.spyOn(globalThis, "fetch").mockImplementation(async (_input, init) => Response.json(init?.method === "PATCH" ? { id: "contact" } : { id: "contact", singleValueExtendedProperties: properties }));
    const result = await harness().execute("outlook_contacts_write", { action: "update", contactId: "contact", homeFaxNumber: "" });
    expect(result).toMatchObject({ ok: true, mutationApplied: true, verification: "read_back", item: { homeFaxNumber: "" } });
    expect(result).not.toHaveProperty("appliedButUnverified");
    expect(fetch.mock.calls.filter(([, init]) => init?.method === "PATCH")).toHaveLength(1);
  });

  it("does not verify a clear when the saved fax is still nonempty", async () => {
    vi.spyOn(globalThis, "fetch").mockImplementation(async (_input, init) => Response.json(init?.method === "PATCH" ? { id: "contact" } : { id: "contact", singleValueExtendedProperties: [homeFax] }));
    const result = await harness().execute("outlook_contacts_write", { action: "update", contactId: "contact", homeFaxNumber: "" });
    expect(result).toMatchObject({ ok: true, mutationApplied: true, appliedButUnverified: true, verification: "failed" });
  });

  it("retains a successful mutation receipt when provider readback fails and never repeats the write", async () => {
    const fetch = vi.spyOn(globalThis, "fetch").mockImplementation(async (_input, init) => init?.method === "PATCH" ? Response.json({ id: "contact", displayName: "Ada" }) : Response.json({ error: { code: "AccessDenied" } }, { status: 403 }));
    const result = await harness().execute("outlook_contacts_write", { action: "update", contactId: "contact", homeFaxNumber: homeFax.value });
    expect(result).toMatchObject({ ok: true, mutationApplied: true, phase: "partial", outcome: "applied_with_warning", warning: "contact_readback_failed", appliedButUnverified: true, retrySafety: "readback_before_retry", verification: "failed" });
    expect(result.nextAction).toContain("do not repeat");
    expect(result.item).not.toHaveProperty("homeFaxNumber");
    expect(fetch.mock.calls.filter(([, init]) => init?.method === "PATCH")).toHaveLength(1);
  });

  it.each([
    { businessAddress: { city: "partial-not-safe" } },
    { businessAddress: { ...address, unknown: "not-allowed" } },
    { homeAddress: [] }, { otherAddress: "not-an-object" }, { businessAddress: { ...address, city: 3 } },
    { businessAddress: { ...address, street: "a".repeat(1025) } }, { otherAddress: { ...address, postalCode: "a".repeat(65) } },
    { businessFaxNumber: null }, { homeFaxNumber: [] }, { primaryFaxNumber: 5 }, { homeFaxNumber: "a".repeat(101) }, { businessFaxNumber: "+49\n30" }, { homeFaxNumber: "   " },
    { singleValueExtendedProperties: [unrelated] }, { faxNumber: "ambiguous" },
  ])("rejects malformed/unsupported fields before approval, credentials, or provider access: %j", async fields => {
    const fetch = vi.spyOn(globalThis, "fetch");
    const h = harness();
    const params = { action: "update", contactId: "contact", ...fields };
    const gate = await h.hooks.before_tool_call({ toolName: "outlook_contacts_write", toolCallId: "invalid", params }, h.context);
    expect(gate.block).toBe(true);
    expect(gate.requireApproval).toBeUndefined();
    expect(credentials.tokenForAuthorizedOperation).not.toHaveBeenCalled();
    expect(fetch).not.toHaveBeenCalled();
  });

  it("keeps delete action-coupled and binds approval to exact address and fax content", async () => {
    const fetch = vi.spyOn(globalThis, "fetch");
    const h = harness();
    expect(await h.execute("outlook_contacts_write", { action: "delete", contactId: "contact", homeAddress: null })).toMatchObject({ ok: false, error: "invalid_write_parameter" });
    const params = { action: "update", contactId: "contact", businessAddress: address, businessFaxNumber: businessFax.value };
    const gate = await h.hooks.before_tool_call({ toolName: "outlook_contacts_write", toolCallId: "bound", params }, h.context);
    expect(gate.requireApproval).toBeDefined();
    await gate.requireApproval.onResolution("allow-once");
    const result = (await h.tools.outlook_contacts_write.execute("bound", { ...params, businessAddress: { ...address, street: "different target content" }, businessFaxNumber: "changed" })).details;
    expect(result).toMatchObject({ ok: false, error: "approval_context_invalid_or_changed", mutationApplied: false });
    expect(credentials.tokenForAuthorizedOperation).not.toHaveBeenCalled();
    expect(fetch).not.toHaveBeenCalled();
  });
});
