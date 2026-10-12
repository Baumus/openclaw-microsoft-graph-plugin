import { AsyncLocalStorage } from "node:async_hooks";

export type AccountScope = {
  identity: string;
  stateDir: string;
  policy: unknown;
  /** Bounded local directory admission allowance; caller/native cancellation wins. */
  directoryAdmissionAllowanceMs?: number;
  assertCurrent: () => void;
  assertFresh: () => Promise<void>;
};
const key = Symbol.for("@baumus/openclaw-microsoft-graph/account-scope/v4");
const slots = globalThis as Record<symbol, unknown>;
export const accountScope = (slots[key] ??= new AsyncLocalStorage<AccountScope>()) as AsyncLocalStorage<AccountScope>;
export function currentAccountIdentity(): string | undefined { return accountScope.getStore()?.identity; }
export async function assertAccountEffect(): Promise<void> {
  const scope = accountScope.getStore();
  if (!scope) return; // Pure exported helpers remain independently testable; registered tools always install a scope.
  scope.assertCurrent();
  await scope.assertFresh();
  scope.assertCurrent();
}
/** Capture at callback creation; a host approval callback need not retain the original ALS context. */
export function captureAccountGuard(): () => void {
  const scope = accountScope.getStore();
  return () => scope?.assertCurrent();
}
