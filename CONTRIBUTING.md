# Contributing

Contributions are welcome through normal repository review. Keep changes bounded to the plugin and preserve its fail-closed behavior.

## Development setup

Requirements:

- Node.js `>=24.16.0 <25` or `>=26.1.0`
- npm
- OpenClaw-compatible Linux or macOS development environment

Install from the committed lockfile:

```bash
npm ci
```

Do not use real Microsoft credentials, accounts, mailbox data, drive data, or private policy files in tests. Use synthetic fixtures and mocked `fetch`, OAuth, credential, filesystem, and media-store boundaries.

## Required checks

```bash
npm test
npm run typecheck
npm run build
npm run plugin:build:check
npm run plugin:validate
npm run package:check
npx --no-install clawhub package validate .
git diff --check
```

If runtime metadata changes, run `npm run plugin:build` once to regenerate `openclaw.plugin.json`, inspect the diff, then rerun `npm run plugin:build:check`.

## Pull requests

- Explain the behavior and security boundary being changed.
- Add focused regression coverage for changed behavior and failure ordering.
- Keep package/manifest versions aligned.
- Do not add publishing, deployment, activation, or secret-bearing workflows.
- Do not weaken policy, confirmation, approval, path-containment, size, timeout, or credential-scope checks to make a test pass.
- Document new delegated permissions before using them.

By contributing, you agree that your contribution is licensed under Apache-2.0.
