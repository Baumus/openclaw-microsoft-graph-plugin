import { join } from "node:path";
import { Type } from "typebox";
import { GraphPolicySchema, validatePolicy, type GraphPolicy } from "./policy.js";

const ACCOUNT_ID_PATTERN = "^[a-z][a-z0-9-]{0,62}$";
const accountIdRegex = new RegExp(ACCOUNT_ID_PATTERN);
export const AccountSelector = Type.Optional(Type.String({
  pattern: ACCOUNT_ID_PATTERN,
  description: "Connection ID from microsoft_graph_capabilities. Omit for the legacy connection; always repeat the same ID for follow-up calls.",
}));
export const AccountsSchema = Type.Record(Type.String({ pattern: ACCOUNT_ID_PATTERN }), GraphPolicySchema, { additionalProperties: false });
export type AccountConfig = { policy?: GraphPolicy; accounts?: Record<string, GraphPolicy> };

export function selectAccount<T extends AccountConfig>(config: T, accountId: unknown): T {
  if (accountId === undefined) return config;
  if (typeof accountId !== "string" || !accountIdRegex.test(accountId) || !config.accounts || !Object.hasOwn(config.accounts, accountId)) {
    throw new Error("access_denied");
  }
  return { ...config, policy: validatePolicy(config.accounts[accountId]) };
}

export function accountStateDir(stateDir: string, accountId: string | undefined): string {
  return accountId === undefined ? stateDir : join(stateDir, "plugin-data", "microsoft-graph", "accounts", accountId);
}
