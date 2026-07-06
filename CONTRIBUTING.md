# Contributing

## Secret & PII protection

This repo has layered guards to keep credentials and personal information out of
commits and published artifacts. Please keep them working.

### Enable the pre-commit hook (one-time, per clone)

```bash
git config core.hooksPath .githooks
```

This runs `scripts/scan-secrets.mjs` on your staged changes before each commit
and blocks anything that looks like a credential or a personal email address.

### What the scanner checks

- **Credentials**: Fastmail API tokens (`fmu…`), `Bearer`/`Basic` auth values,
  and hardcoded `token`/`secret`/`password`/`api_key` assignments.
- **Personal information**: email addresses on any domain outside a small
  allowlist of placeholder/service domains (`example.com`, `fastmail.com`, …).

Run it manually anytime:

```bash
npm run scan:secrets
```

The same scan runs in CI (`.github/workflows/secret-scan.yml`) on every push and
pull request, so it catches anything the local hook missed.

### Test fixtures must be synthetic

Never paste a real token, password, or personal email into a test — not even a
revoked one. Use obviously-fake values (`example.com` addresses, zero/`a`-filled
token shapes). If a synthetic value is unavoidably credential-shaped and the
scanner flags it, add the marker `allowlist-secret` to that line.

### Local denylist for your own identifiers (optional but recommended)

To make the scanner also flag *your* real domains/addresses without publishing
them, copy the template and fill it in — the target file is gitignored:

```bash
cp .secret-scan-local.txt.example .secret-scan-local.txt
# then add your personal domains/addresses, one per line
```

### Packaging

The published `.dxt` is built from `dist/` plus runtime dependencies only.
`.dxtignore` excludes `src/`, all `*.test.*` files, scripts, and CI config, so
source and test files never ship inside a release binary.
