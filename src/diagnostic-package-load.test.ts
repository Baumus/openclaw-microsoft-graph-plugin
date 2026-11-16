import { readFile } from "node:fs/promises";
import { pathToFileURL } from "node:url";
import { createHash } from "node:crypto";
import { expect, it } from "vitest";
const artifact = process.env.MS_GRAPH_DIAGNOSTIC_PACKAGE_ROOT;
it.skipIf(!artifact)("loads the actual unpacked temporary artifact without host activation", async () => {
 const root = pathToFileURL(artifact! + "/");
 const manifest = JSON.parse(await readFile(new URL("package.json", root), "utf8"));
 expect(manifest.version).toBe("3.25.3");
 const hooks = new Map(); const factories: unknown[] = []; const logs: any[] = [];
 const url = new URL("dist/index.js", root).href;
 const { default: plugin } = await import(/* @vite-ignore */ url);
 plugin.register({
  pluginConfig: { enabled: false }, runtime: {},
  on: (name: string, handler: any) => hooks.set(name, handler),
  registerTool: (factory: unknown) => factories.push(factory),
  logger: { info: (line: string) => logs.push(JSON.parse(line)), warn() {}, error() {} },
 });
 expect(await hooks.get("before_tool_call")({
  toolName: "onedrive_upload", toolCallId: "synthetic-private-id",
  params: { rootLabel: "synthetic", relativePath: "private.md", sourceWorkspacePath: "private.md" },
 }, {})).toEqual({ block: true, blockReason: "connector_disabled" });
 const diagnostic = logs.filter(x => x.event === "microsoft_graph_approval_binding");
 expect(diagnostic.some(x => x.stage === "hook_result" && x.category === "disabled")).toBe(true);
 expect(JSON.stringify(diagnostic)).not.toContain("synthetic-private-id");
 expect(JSON.stringify(diagnostic)).not.toContain("private.md");
 expect(factories.length).toBeGreaterThan(0);
 expect(createHash("sha256").update(await readFile(new URL("dist/index.js", root))).digest("hex"))
   .toBe(process.env.MS_GRAPH_DIAGNOSTIC_BUNDLE_SHA256);
});
