import type { OpenClawPluginApi } from "openclaw/plugin-sdk/core";
/** No userless credential administration: use the authenticated personal Control UI. */
export function registerAccountCli(api: OpenClawPluginApi): void {
  api.registerCli(({ program }) => {
    program.command("microsoft-graph").description("Personal Microsoft Graph accounts").command("accounts").description("Show personal account setup guidance").action(() => {
      process.stdout.write("Open the authenticated Control UI → Mein Microsoft-Konto to sign in, migrate your legacy connection, choose session accounts, or disconnect. No global credential or token-export command is available in version 4.\n");
    });
  }, { commands: ["microsoft-graph"], descriptors: [{ name: "microsoft-graph", description: "Microsoft Graph operator commands", hasSubcommands: true }] });
}
