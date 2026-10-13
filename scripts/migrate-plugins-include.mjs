#!/usr/bin/env node
import { createHash } from "node:crypto";
import { readFile, writeFile, rename, lstat, realpath, rm } from "node:fs/promises";
import { dirname, isAbsolute, join, relative, resolve } from "node:path";
import JSON5 from "json5";

const INCLUDE = "$include";
const PLUGIN = "microsoft-graph";
const digest = (bytes) => createHash("sha256").update(bytes).digest("hex");
const object = (value) => value !== null && typeof value === "object" && !Array.isArray(value);
const own = (value, key) => Object.prototype.hasOwnProperty.call(value, key);

function fail(message) { throw new Error(message); }

function tokens(source) {
  const out = [];
  for (let i = 0; i < source.length;) {
    const start = i;
    const ch = source[i];
    if (/\s/.test(ch)) { i++; continue; }
    if (ch === "/" && source[i + 1] === "/") { i = source.indexOf("\n", i + 2); if (i < 0) break; continue; }
    if (ch === "/" && source[i + 1] === "*") { i = source.indexOf("*/", i + 2); if (i < 0) fail("Unterminated JSON5 comment"); i += 2; continue; }
    if (ch === '"' || ch === "'") {
      const quote = ch; i++;
      while (i < source.length) {
        if (source[i] === "\\") { i += 2; continue; }
        if (source[i++] === quote) break;
      }
      if (source[i - 1] !== quote) fail("Unterminated JSON5 string");
      out.push({ kind: "string", value: JSON5.parse(source.slice(start, i)), start, end: i });
      continue;
    }
    if ("{}[]:,".includes(ch)) { out.push({ kind: "punct", value: ch, start, end: ++i }); continue; }
    while (i < source.length && !/[\s{}\[\]:,]/.test(source[i]) && source.slice(i, i + 2) !== "//" && source.slice(i, i + 2) !== "/*") i++;
    if (i === start) fail("Unsupported JSON5 token");
    out.push({ kind: "word", value: source.slice(start, i), start, end: i });
  }
  return out;
}

function topLevelPluginsValueSpan(source) {
  const items = tokens(source);
  const punct = (item, value) => item?.kind === "punct" && item.value === value;
  if (!punct(items[0], "{")) fail("Root config must be a JSON5 object");
  let depth = 0; let span;
  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    if (depth === 1 && item.kind !== "punct" && item.value === "plugins" && punct(items[i + 1], ":")) {
      if (span) fail("Duplicate top-level plugins keys are unsupported");
      const first = items[i + 2];
      if (!punct(first, "{")) fail("Top-level plugins must be an object");
      let nested = 0; let last;
      for (let j = i + 2; j < items.length; j++) {
        if (punct(items[j], "{") || punct(items[j], "[")) nested++;
        if (punct(items[j], "}") || punct(items[j], "]")) nested--;
        if (nested === 0) { last = items[j]; break; }
      }
      if (!last) fail("Unterminated plugins object");
      span = { start: first.start, end: last.end };
    }
    if (punct(item, "{") || punct(item, "[")) depth++;
    if (punct(item, "}") || punct(item, "]")) depth--;
  }
  if (!span) fail("Top-level plugins object is missing");
  return span;
}

function rejectIncludes(value, path = "plugins") {
  if (Array.isArray(value)) return value.forEach((entry, i) => rejectIncludes(entry, `${path}[${i}]`));
  if (!object(value)) return;
  if (own(value, INCLUDE)) fail(`Unexpected nested include at ${path}`);
  for (const [key, entry] of Object.entries(value)) rejectIncludes(entry, `${path}.${key}`);
}

async function regularFile(path) {
  const stat = await lstat(path);
  if (!stat.isFile()) fail("Config inputs must be regular files, not symlinks or directories");
  if (stat.size > 2 * 1024 * 1024) fail("Config input exceeds OpenClaw's 2 MB include limit");
  return stat;
}

function inside(root, target) {
  const rel = relative(root, target);
  return rel && !rel.startsWith("..") && !isAbsolute(rel);
}

export async function migrate(configPath, { apply = false, outputName = "plugins.json" } = {}) {
  const absoluteConfig = resolve(configPath);
  const root = dirname(absoluteConfig);
  if (outputName !== "plugins.json") fail("Only plugins.json is supported");
  const outputPath = join(root, outputName);
  const rootStat = await regularFile(absoluteConfig);
  const rootBytes = await readFile(absoluteConfig);
  const source = rootBytes.toString("utf8");
  const parsed = JSON5.parse(source);
  if (!object(parsed) || !object(parsed.plugins)) fail("Top-level plugins object is missing");
  const span = topLevelPluginsValueSpan(source);
  const current = parsed.plugins;
  if (Object.keys(current).length === 1 && [outputName, `./${outputName}`].includes(current[INCLUDE])) {
    await regularFile(outputPath);
    const output = JSON5.parse(await readFile(outputPath, "utf8"));
    if (!object(output) || !object(output.entries?.[PLUGIN]?.config?.policy)) fail("Existing plugins.json has no Microsoft Graph policy; migration is not complete");
    rejectIncludes(output);
    return { status: "already-migrated", changed: false };
  }
  if (own(current, INCLUDE)) fail("Existing plugins include has an unsupported shape");
  const policy = current.entries?.[PLUGIN]?.config?.policy;
  if (!object(policy) || Object.keys(policy).length !== 1 || typeof policy[INCLUDE] !== "string") fail("Expected one Microsoft Graph policy file include");
  const policyPath = resolve(root, policy[INCLUDE]);
  if (!inside(root, policyPath)) fail("Policy include must be inside the config directory");
  await regularFile(policyPath);
  if (!inside(await realpath(root), await realpath(policyPath))) fail("Policy include escapes the config directory through a symlink");
  const policyBytes = await readFile(policyPath);
  const resolvedPolicy = JSON5.parse(policyBytes.toString("utf8"));
  if (!object(resolvedPolicy)) fail("Included policy must be an object");
  rejectIncludes(resolvedPolicy, "policy");
  const plugins = structuredClone(current);
  plugins.entries[PLUGIN].config.policy = resolvedPolicy;
  rejectIncludes(plugins);
  const outputBytes = Buffer.from(`${JSON.stringify(plugins, null, 2)}\n`);
  if (outputBytes.length > 2 * 1024 * 1024) fail("Resulting plugins.json exceeds OpenClaw's 2 MB include limit");
  const newRoot = Buffer.from(source.slice(0, span.start) + `{ $include: "${outputName}" }` + source.slice(span.end));
  const parsedNewRoot = JSON5.parse(newRoot.toString("utf8"));
  if (parsedNewRoot.plugins?.[INCLUDE] !== outputName) fail("Root replacement failed validation");
  try { await lstat(outputPath); fail("plugins.json already exists; refusing to overwrite it"); }
  catch (error) { if (error.code !== "ENOENT") throw error; }
  if (!apply) return { status: "ready", changed: false, configPath: absoluteConfig, outputPath };

  // The include target is fully written before the root pointer changes. A byte
  // backup is retained for rollback; the old policy file is never removed.
  const backupPath = `${absoluteConfig}.before-plugins-include.${digest(rootBytes).slice(0, 12)}.bak`;
  const tempRoot = `${absoluteConfig}.plugins-include-${process.pid}.tmp`;
  let outputCreated = false;
  try {
    await writeFile(outputPath, outputBytes, { flag: "wx", mode: 0o600 });
    outputCreated = true;
    try { await writeFile(backupPath, rootBytes, { flag: "wx", mode: 0o600 }); }
    catch (error) {
      if (error.code !== "EEXIST") throw error;
      const backupStat = await lstat(backupPath);
      if (!backupStat.isFile() || (backupStat.mode & 0o077) !== 0 || digest(await readFile(backupPath)) !== digest(rootBytes)) fail("An existing backup is not private or does not match the current config");
    }
    if (digest(await readFile(absoluteConfig)) !== digest(rootBytes) || digest(await readFile(policyPath)) !== digest(policyBytes) || digest(await readFile(outputPath)) !== digest(outputBytes)) fail("Config changed during migration; no root replacement made");
    await writeFile(tempRoot, newRoot, { flag: "wx", mode: rootStat.mode & 0o777 });
    await rename(tempRoot, absoluteConfig);
    return { status: "migrated", changed: true, backupPath, outputPath };
  } catch (error) {
    await rm(tempRoot, { force: true });
    if (outputCreated) await rm(outputPath, { force: true });
    throw error;
  }
}

if (process.argv[1] && resolve(process.argv[1]) === new URL(import.meta.url).pathname) {
  const args = process.argv.slice(2);
  const apply = args.includes("--apply");
  const configIndex = args.indexOf("--config");
  if (args.some((arg, i) => arg !== "--apply" && arg !== "--config" && i !== configIndex + 1) || configIndex < 0 || !args[configIndex + 1]) {
    process.stderr.write("Usage: node scripts/migrate-plugins-include.mjs --config /path/to/openclaw.json [--apply]\n");
    process.exitCode = 2;
  } else {
    migrate(args[configIndex + 1], { apply }).then((result) => {
      process.stdout.write(`${JSON.stringify(result)}\n`);
    }, (error) => {
      process.stderr.write(`Migration refused: ${error.message}\n`);
      process.exitCode = 1;
    });
  }
}
