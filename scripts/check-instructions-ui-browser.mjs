// Synthetic native-browser regression and screenshots. No Gateway/provider access.
// PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH=/path/to/chrome node scripts/check-instructions-ui-browser.mjs
import assert from "node:assert/strict";
import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { build } from "esbuild";
import { chromium } from "playwright-core";

const output = path.resolve(process.env.UI_EVIDENCE_DIR ?? "/tmp/msgraph-instructions-ui-evidence");
await mkdir(output, { recursive: true });
const bundle = await build({ entryPoints: ["src/control-ui.ts"], bundle: true, write: false, format: "iife", globalName: "GraphFixture", platform: "browser", loader: { ".css": "text" }, plugins: [{ name: "synthetic-host-contract", setup(builder) {
  builder.onResolve({ filter: /^openclaw\/plugin-sdk\/control-ui$/ }, () => ({ path: "control-ui", namespace: "synthetic-host" }));
  builder.onLoad({ filter: /.*/, namespace: "synthetic-host" }, () => ({ contents: "export const defineControlUiPlugin = plugin => plugin;", loader: "js" }));
} }] });
const browser = await chromium.launch({ headless: true, ...(process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH } : {}) });
const report = [];
try {
  for (const [name, viewport, mobile] of [["desktop", { width: 1280, height: 1000 }, false], ["mobile", { width: 390, height: 844 }, true]]) {
    const context = await browser.newContext({ viewport, isMobile: mobile, hasTouch: mobile });
    // Fixtures are entirely local. Reject every accidental network request.
    await context.route("**/*", route => route.abort());
    const page = await context.newPage(); page.setDefaultTimeout(5000); const errors = [];
    console.log(`Checking ${name} native popover interactions`);
    page.on("pageerror", error => errors.push(error.message));
    await page.setContent('<!doctype html><html lang="de"><meta name="viewport" content="width=device-width,initial-scale=1"><style>body{margin:0;background:#f7f6f3;font:16px/1.4 system-ui,sans-serif}#fixture{height:100vh;overflow:auto}</style><div id="fixture"></div></html>');
    await page.addScriptTag({ content: bundle.outputFiles[0].text });
    await page.evaluate(() => {
      const root = (label, trusted = false) => ({ label, path: `/Synthetic/${label}`, drive_id: "synthetic-drive", item_id: label, include_descendants: true, ...(trusted ? { agents_instructions: "trusted" } : {}), permissions: { read: true, write: true, delete: false }, agents: { reader: { permissions: { read: true } }, editor: { permissions: { read: true, write: true } } } });
      const policy = { version: 2, rules: { default: "deny" }, services: { onedrive: { allowed_roots: [root("Documents"), root("Projects", true)] }, calendar: { agents: {} }, mail: { agents: {} }, todo: { agents: {} } } };
      let pageDefinition; let listener; let delayedStatus;
      const host = { locale: "de", connection: { connected: true, canAdmin: true }, agents: { rows: [{ id: "reader", name: "Reader" }, { id: "editor", name: "Editor" }] }, subscribe: fn => { listener = fn; return () => {}; },
        request: async method => {
          if (method === "config.get") return { hash: "fixture", configRevisionHash: "fixture", appliedConfigHash: "fixture", config: { plugins: { entries: { "microsoft-graph": { enabled: true, config: { credentialVaultKey: "synthetic-ref", policy } } } } } };
          if (method === "microsoft-graph.credentials.status") return { ok: true, value: { policyVersion: 2, credential: { result: "valid" } } };
          if (method === "microsoft-graph.updateStatus") return { updateAvailable: false };
          if (method === "microsoft-graph.configuration.applicationStatus") { if (delayedStatus) await delayedStatus; return { activePolicyHash: null }; }
          throw new Error(`Disallowed fixture RPC: ${method}`);
        }, ui: { registerPage: definition => { pageDefinition = definition; return () => {}; }, registerNavigation: () => () => {} } };
      const stop = GraphFixture.default.activate(host); const controller = new AbortController();
      const view = pageDefinition.mount(document.querySelector("#fixture"), { host, signal: controller.signal });
      window.fixture = { notify: () => listener(), delay: () => { delayedStatus = new Promise(resolve => { window.fixture.finish = resolve; }); }, dispose: () => { view.dispose(); controller.abort(); stop(); } };
    });
    const info = page.locator(".mg-instructions-info"); const inputs = page.locator(".mg-instructions-control input");
    await inputs.first().waitFor();
    assert.equal(await inputs.count(), 2); assert.equal(await inputs.nth(0).isChecked(), false); assert.equal(await inputs.nth(1).isChecked(), true);
    await info.first().scrollIntoViewIfNeeded();
    await page.screenshot({ path: path.join(output, `${name}-closed.png`), fullPage: true });
    const open = page.locator(".mg-instructions-help:popover-open");
    const checkOpen = async index => { await open.waitFor(); await page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)))); assert.equal(await info.nth(index).getAttribute("aria-expanded"), "true"); };
    const checkClosed = async () => { await open.waitFor({ state: "hidden" }); assert.equal(await info.first().getAttribute("aria-expanded"), "false"); };
    await info.first().focus(); await page.keyboard.press("Enter"); await checkOpen(0);
    await page.mouse.move(1, 1); await page.waitForTimeout(50); assert.equal(await open.count(), 1);
    assert.equal(await open.getAttribute("role"), "note"); assert.ok(await info.first().getAttribute("aria-label"));
    const bounds = await open.boundingBox();
    assert.ok(bounds.x >= 0 && bounds.y >= 0 && bounds.x + bounds.width <= viewport.width && bounds.y + bounds.height <= viewport.height, JSON.stringify(bounds));
    assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth), true);
    assert.equal(await open.evaluate(node => node.scrollWidth <= node.clientWidth), true);
    await page.screenshot({ path: path.join(output, `${name}-open.png`), fullPage: true });
    await page.keyboard.press("Escape"); await checkClosed(); assert.equal(await info.first().evaluate(node => node === document.activeElement), true);
    await page.keyboard.press("Space"); await checkOpen(0);
    await page.keyboard.press("Space"); await checkClosed();
    await info.first().click(); await checkOpen(0); await page.mouse.click(2, 2); await checkClosed();
    if (mobile) { await info.first().tap(); await checkOpen(0); await info.first().tap(); await checkClosed(); }
    // Opening the other root dismisses the first and background renders retain
    // the correct root's popup/focus, not the first matching glyph/checkbox.
    await info.first().click(); await checkOpen(0);
    await info.nth(1).scrollIntoViewIfNeeded(); await info.nth(1).click(); await checkOpen(1);
    assert.equal(await open.count(), 1); assert.equal(await info.first().getAttribute("aria-expanded"), "false");
    await page.evaluate(() => window.fixture.notify()); await checkOpen(1);
    assert.equal(await info.nth(1).evaluate(node => node === document.activeElement), true);
    await page.keyboard.press("Escape");
    // A delayed status completion must not steal newer semantic focus.
    await page.evaluate(() => { window.fixture.delay(); window.fixture.notify(); });
    await inputs.nth(1).focus(); await page.evaluate(() => window.fixture.finish());
    await page.waitForTimeout(50); assert.equal(await inputs.nth(1).evaluate(node => node === document.activeElement), true);
    // Exercise native keyboard toggling and root-specific focus restoration.
    await page.keyboard.press("Space"); assert.equal(await inputs.nth(1).isChecked(), false);
    assert.equal(await inputs.nth(1).evaluate(node => node === document.activeElement), true);
    await page.keyboard.press("Space"); assert.equal(await inputs.nth(1).isChecked(), true);
    // Short viewport still permits all help text through its own scroller.
    await page.setViewportSize({ width: viewport.width, height: 420 });
    await info.first().scrollIntoViewIfNeeded(); await info.first().click(); await checkOpen(0);
    const smallBounds = await open.boundingBox(); assert.ok(smallBounds.y >= 0 && smallBounds.y + smallBounds.height <= 420, JSON.stringify(smallBounds));
    await open.evaluate(node => { node.scrollTop = node.scrollHeight; });
    assert.equal(await open.evaluate(node => node.scrollTop + node.clientHeight >= node.scrollHeight - 1), true);
    assert.deepEqual(errors, []);
    report.push({ viewport: name, checks: "default/preexisting; native Enter/Space/Escape; outside/same-button/touch; persistence; correct-root/delayed focus; bounded top layer; short viewport scrolling", screenshots: [`${name}-closed.png`, `${name}-open.png`], consoleErrors: errors });
    await page.evaluate(() => window.fixture.dispose()); await context.close();
  }
  await writeFile(path.join(output, "report.json"), JSON.stringify(report, null, 2) + "\n");
  console.log(JSON.stringify({ result: "PASS", output, report }, null, 2));
} finally { await browser.close(); }
