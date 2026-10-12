import { execFile } from "node:child_process";
import { readFileSync } from "node:fs";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);
const expected = [
  "CHANGELOG.md",
  "CONTRIBUTING.md",
  "LICENSE",
  "README.md",
  "SECURITY.md",
  "assets/ICON_PROVENANCE.md",
  "assets/icon.png",
  "dist/index.js",
  "docs/ARCHITECTURE.md",
  "docs/CREDENTIAL_VAULT_DESIGN.md",
  "docs/OAUTH_ACCESS_MATRIX.md",
  "docs/SUPPORT.md",
  "docs/MULTI_USER.md",
  "src/account-control-ui.ts",
  "src/device-verification.ts",
  "src/config-policy-identity.ts",
  "src/control-ui-status.ts",
  "examples/microsoft-graph-policy.json5",
  "examples/microsoft-graph-policy-v2.json5",
  "openclaw.plugin.json",
  "package.json",
  "scripts/bounded-json.mjs",
  "src/control-ui.ts",
  "src/control-ui-i18n.ts",
  "src/control-ui.css",
].sort();
const manifest = JSON.parse(readFileSync(new URL("../openclaw.plugin.json", import.meta.url), "utf8"));
if (manifest.controlUi?.entry) expected.push(manifest.controlUi.entry);
for (const style of manifest.controlUi?.styles ?? []) expected.push(style);
expected.sort();

const { stdout } = await execFileAsync("npm", ["pack", "--dry-run", "--json", "--ignore-scripts"], {
  encoding: "utf8",
  maxBuffer: 4 * 1024 * 1024,
});
const payload = JSON.parse(stdout);
const pkg = JSON.parse(readFileSync(new URL("../package.json", import.meta.url), "utf8"));
// npm <=11 returns an array; npm 12 keys the same records by package name.
const records = Array.isArray(payload) ? payload
  : payload && typeof payload === "object" && Object.keys(payload).length === 1 && Object.hasOwn(payload, pkg.name)
    ? [payload[pkg.name]] : [];
const record = records[0];
if (records.length !== 1 || record?.name !== pkg.name || record?.version !== pkg.version
  || !Array.isArray(record?.files) || record.files.some(entry => typeof entry?.path !== "string" || !entry.path)) {
  throw new Error("invalid_npm_pack_output");
}
const actual = record.files.map((entry) => entry.path).sort();
if (JSON.stringify(actual) !== JSON.stringify(expected)) {
  const missing = expected.filter((entry) => !actual.includes(entry));
  const unexpected = actual.filter((entry) => !expected.includes(entry));
  throw new Error(`package_contents_mismatch missing=${JSON.stringify(missing)} unexpected=${JSON.stringify(unexpected)}`);
}
process.stdout.write(`${JSON.stringify({ ok: true, files: actual }, null, 2)}\n`);
