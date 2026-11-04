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

const fields = {
  assistantName: "Assistant", birthday: "1974-07-22T00:00:00Z", businessHomePage: "https://example.com", categories: ["Friends"], children: ["Child"], department: "Research", fileAs: "Example, Ada", generation: "III", homePhones: ["+49 30 123", "+49 30 456"], imAddresses: ["sip:ada@example.com"], initials: "AE", manager: "Manager", middleName: "Middle", nickName: "Nickname", officeLocation: "Room 1", parentFolderId: "folder", personalNotes: "Synthetic notes", primaryEmailAddress: { address: "ada@example.com", name: "Ada" }, secondaryEmailAddress: { address: "ada2@example.com" }, tertiaryEmailAddress: { address: "ada3@example.com" }, profession: "Engineer", spouseName: "Spouse", title: "Dr", yomiCompanyName: "Company", yomiGivenName: "Ada", yomiSurname: "Example",
};
afterEach(() => { vi.restoreAllMocks(); credentials.tokenForAuthorizedOperation.mockClear(); });
describe("complete Graph contact fields", () => {
  it.each(["create", "update"])("%s sends every added writable property and preserves provider output", async action => {
    const fetch = vi.spyOn(globalThis, "fetch").mockImplementation(async (_url, init) => Response.json({ id: "contact", ...JSON.parse(String(init?.body)), changeKey: "version", createdDateTime: "2026-01-01T00:00:00Z", lastModifiedDateTime: "2026-01-02T00:00:00Z" }));
    const result = await harness().execute("outlook_contacts_write", { action, ...(action === "update" ? { contactId: "contact" } : {}), ...fields });
    expect(result).toMatchObject({ ok: true, item: { ...fields, changeKey: "version" } });
    expect(JSON.parse(String(fetch.mock.calls[0][1]?.body))).toEqual(fields);
  });
  it("returns and searches added properties across list/get/search with explicit selects", async () => {
    const item = { id: "contact", ...fields, changeKey: "version", createdDateTime: "2026-01-01T00:00:00Z", lastModifiedDateTime: "2026-01-02T00:00:00Z", primaryEmailAddress: { ...fields.primaryEmailAddress, providerOnly: "hidden" }, singleValueExtendedProperties: [], photo: { hidden: true } };
    const fetch = vi.spyOn(globalThis, "fetch").mockImplementation(async input => new URL(String(input)).pathname.endsWith("/contact") ? Response.json(item) : Response.json({ value: [item] }));
    const h = harness();
    for (const params of [{ action: "list_contacts" }, { action: "get_contact", contactId: "contact" }, { action: "search_contacts", search: "123" }, { action: "search_contacts", search: "Synthetic notes" }]) {
      const result = await h.execute("outlook_contacts_read", params);
      expect(result.ok).toBe(true);
      const actual = result.item ?? result.items[0];
      expect(actual).toMatchObject(fields);
      expect(actual).not.toHaveProperty("photo");
      expect(actual.primaryEmailAddress).not.toHaveProperty("providerOnly");
    }
    for (const [url] of fetch.mock.calls) for (const field of Object.keys(fields)) expect(new URL(String(url)).searchParams.get("$select")?.split(",")).toContain(field);
  });
  it("normalizes date-only birthdays, preserves offsets, clears explicitly and omits unchanged fields", async () => {
    const fetch = vi.spyOn(globalThis, "fetch").mockImplementation(async (_url, init) => Response.json({ id: "contact", ...JSON.parse(String(init?.body)) }));
    const h = harness();
    for (const birthday of ["2000-02-29", "2000-02-29T12:30:00+02:00", null]) {
      expect(await h.execute("outlook_contacts_write", { action: "update", contactId: "contact", birthday, homePhones: [], children: [], personalNotes: null, primaryEmailAddress: null })).toMatchObject({ ok: true });
      const body = JSON.parse(String(fetch.mock.calls.at(-1)?.[1]?.body));
      expect(body).toEqual({ birthday: birthday === "2000-02-29" ? "2000-02-29T00:00:00Z" : birthday, homePhones: [], children: [], personalNotes: null, primaryEmailAddress: null });
      expect(body).not.toHaveProperty("businessPhones");
    }
  });
  it.each(["2025-02-29", "2000-13-01", "2000-02-30", "yesterday", "2000-01-01T12:00:00", 123, "2000-01-01T24:99:00Z", "2000-01-01T24:00:00Z"])("rejects invalid birthday %s before credentials/network/approval", async birthday => {
    const fetch = vi.spyOn(globalThis, "fetch");
    const h = harness();
    const gate = await h.hooks.before_tool_call({ toolName: "outlook_contacts_write", toolCallId: "invalid", params: { action: "update", contactId: "contact", birthday } }, h.context);
    expect(gate?.block).toBe(true); expect(gate?.requireApproval).toBeUndefined();
    expect(fetch).not.toHaveBeenCalled(); expect(credentials.tokenForAuthorizedOperation).not.toHaveBeenCalled();
  });
  it.each([{ homePhones: ["a", "b", "c"] }, { children: "child" }, { primaryEmailAddress: { address: "bad" } }, { changeKey: "readonly" }, { createdDateTime: "2026-01-01T00:00:00Z" }, { photo: {} }, { extensions: [] }])("rejects malformed, read-only and relationship writes %j", async fields => {
    const fetch = vi.spyOn(globalThis, "fetch");
    const result = await harness().execute("outlook_contacts_write", { action: "update", contactId: "contact", ...fields });
    expect(result.ok).toBe(false); expect(fetch).not.toHaveBeenCalled(); expect(credentials.tokenForAuthorizedOperation).not.toHaveBeenCalled();
  });
});
