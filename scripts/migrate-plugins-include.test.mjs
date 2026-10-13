import { mkdtemp, readFile, readdir, symlink, writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import JSON5 from "json5";
import { migrate } from "./migrate-plugins-include.mjs";

async function fixture(config = `{
  // Keep unrelated host configuration and comments intact.
  agents: { defaults: { model: "example/model" } },
  plugins: {
    allow: ["microsoft-graph", "other"],
    entries: {
      "other": { enabled: true },
      "microsoft-graph": { enabled: true, config: {
        credentialVaultKey: { source: "store", provider: "default", id: "SYNTHETIC" },
        policy: { $include: "./microsoft-graph-policy.json5" },
      } },
    },
  },
}`) {
  const dir = await mkdtemp(join(tmpdir(), "msgraph-include-test-"));
  const configPath = join(dir, "openclaw.json");
  await writeFile(configPath, config);
  await writeFile(join(dir, "microsoft-graph-policy.json5"), `{ version: 2, rules: { default: "deny" }, services: { mail: { agents: {} } } }`);
  return { dir, configPath, config };
}

describe("single plugins include migration", () => {
  it("previews without changing files and preserves all plugins, policy, comments, and rollback bytes", async () => {
    const f = await fixture();
    expect((await migrate(f.configPath)).status).toBe("ready");
    expect(await readdir(f.dir)).toHaveLength(2);
    const result = await migrate(f.configPath, { apply: true });
    expect(result.status).toBe("migrated");
    const root = await readFile(f.configPath, "utf8");
    const plugins = JSON5.parse(await readFile(join(f.dir, "plugins.json"), "utf8"));
    expect(root).toContain("// Keep unrelated host configuration and comments intact.");
    expect(JSON5.parse(root).plugins).toEqual({ $include: "plugins.json" });
    expect(plugins.allow).toEqual(["microsoft-graph", "other"]);
    expect(plugins.entries.other.enabled).toBe(true);
    expect(plugins.entries["microsoft-graph"].config.credentialVaultKey).toEqual({ source: "store", provider: "default", id: "SYNTHETIC" });
    expect(plugins.entries["microsoft-graph"].config.policy).toEqual({ version: 2, rules: { default: "deny" }, services: { mail: { agents: {} } } });
    expect(await readFile(result.backupPath, "utf8")).toBe(f.config);
    expect((await migrate(f.configPath, { apply: true })).status).toBe("already-migrated");
  });

  it("refuses other includes and existing plugins.json without changing root", async () => {
    const f = await fixture();
    await writeFile(join(f.dir, "plugins.json"), "{}");
    await expect(migrate(f.configPath, { apply: true })).rejects.toThrow("already exists");
    expect(await readFile(f.configPath, "utf8")).toBe(f.config);
    const g = await fixture(f.config.replace("enabled: true },\n      \"microsoft-graph\"", "enabled: true, config: { other: { $include: './other.json5' } } },\n      \"microsoft-graph\""));
    await expect(migrate(g.configPath)).rejects.toThrow("Unexpected nested include");
    expect(await readFile(g.configPath, "utf8")).toBe(g.config);
  });

  it("does not call an empty or non-object plugins include migrated", async () => {
    for (const contents of ["{}", "[]"]) {
      const f = await fixture(`{ plugins: { $include: "plugins.json" } }`);
      await writeFile(join(f.dir, "plugins.json"), contents);
      await expect(migrate(f.configPath, { apply: true })).rejects.toThrow("no Microsoft Graph policy");
    }
  });

  it("refuses to trust a pre-existing world-readable backup", async () => {
    const f = await fixture();
    const hash = createHash("sha256").update(f.config).digest("hex").slice(0, 12);
    await writeFile(`${f.configPath}.before-plugins-include.${hash}.bak`, f.config, { mode: 0o644 });
    await expect(migrate(f.configPath, { apply: true })).rejects.toThrow("not private");
    expect(await readFile(f.configPath, "utf8")).toBe(f.config);
    expect(await readdir(f.dir)).not.toContain("plugins.json");
  });

  it("does not mistake JSON5 string contents for object delimiters or plugin keys", async () => {
    const f = await fixture();
    const source = f.config.replace('agents: { defaults: { model: "example/model" } }', 'agents: { defaults: { model: "}" }, note: "plugins", marker: "{" }');
    await writeFile(f.configPath, source);
    await migrate(f.configPath, { apply: true });
    const root = await readFile(f.configPath, "utf8");
    expect(root).toContain('marker: "{"');
    expect(JSON5.parse(root).plugins).toEqual({ $include: "plugins.json" });
  });

  it("refuses a policy include outside the config directory, including symlink traversal", async () => {
    const f = await fixture();
    await writeFile(f.configPath, f.config.replace("./microsoft-graph-policy.json5", "../outside.json5"));
    await expect(migrate(f.configPath)).rejects.toThrow("inside the config directory");
    const g = await fixture();
    const external = await mkdtemp(join(tmpdir(), "msgraph-external-test-"));
    await writeFile(join(external, "policy.json5"), "{}");
    await symlink(external, join(g.dir, "link"));
    await writeFile(g.configPath, g.config.replace("./microsoft-graph-policy.json5", "./link/policy.json5"));
    await expect(migrate(g.configPath)).rejects.toThrow("symlink");
  });
});
