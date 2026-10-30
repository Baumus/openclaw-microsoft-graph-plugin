import { execFile } from "node:child_process";
import { readFileSync } from "node:fs";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
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
  "docs/CONTACT_ADDRESS_FAX.md",
  "docs/OAUTH_ACCESS_MATRIX.md",
  "docs/SUPPORT.md",
  "examples/microsoft-graph-policy.json5",
  "examples/microsoft-graph-policy-v2.json5",
  "openclaw.plugin.json",
  "package.json",
  "scripts/bounded-json.mjs",
  "scripts/migrate-plugins-include.mjs",
  "src/control-ui.ts",
  "src/control-ui-i18n.ts",
  "src/control-ui.css",
].sort();
const manifest = JSON.parse(readFileSync(new URL("../openclaw.plugin.json", import.meta.url), "utf8"));
if (manifest.controlUi?.entry) expected.push(manifest.controlUi.entry);
for (const style of manifest.controlUi?.styles ?? []) expected.push(style);
expected.sort();

// Pack into a disposable directory so this check inspects the actual release
// files without leaving an artifact in the worktree.
const packDir = await mkdtemp(join(tmpdir(), "microsoft-graph-package-check-"));
let payload;
try {
  const { stdout } = await execFileAsync("npm", ["pack", "--json", "--ignore-scripts", "--pack-destination", packDir], {
    encoding: "utf8",
    maxBuffer: 4 * 1024 * 1024,
  });
  payload = JSON.parse(stdout);
} finally {
  await rm(packDir, { recursive: true, force: true });
}
// npm emits an array when called directly, but some versions key the JSON by
// package name when npm pack runs inside npm run.
const packageName = JSON.parse(readFileSync(new URL("../package.json", import.meta.url), "utf8")).name;
const packEntry = Array.isArray(payload)
  ? (payload.length === 1 ? payload[0] : undefined)
  : (payload && typeof payload === "object" && Object.keys(payload).length === 1 ? payload[packageName] : undefined);
if (!Array.isArray(packEntry?.files)) throw new Error("invalid_npm_pack_output");
const actual = packEntry.files.map((entry) => entry.path).sort();
if (JSON.stringify(actual) !== JSON.stringify(expected)) {
  const missing = expected.filter((entry) => !actual.includes(entry));
  const unexpected = actual.filter((entry) => !expected.includes(entry));
  throw new Error(`package_contents_mismatch missing=${JSON.stringify(missing)} unexpected=${JSON.stringify(unexpected)}`);
}
process.stdout.write(`${JSON.stringify({ ok: true, files: actual }, null, 2)}\n`);
