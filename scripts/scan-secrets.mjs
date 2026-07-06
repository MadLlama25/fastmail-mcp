#!/usr/bin/env node
// Secret & PII scanner — blocks credentials and personal information from being
// committed or published. Runs in CI (all tracked files) and as a pre-commit
// hook (staged files). Self-contained: no third-party dependencies, and it
// embeds NO personal data — personal domains live only in a gitignored local
// denylist (see .secret-scan-local.txt.example).
//
// Usage:
//   node scripts/scan-secrets.mjs --all         scan all git-tracked files
//   node scripts/scan-secrets.mjs --staged      scan staged (pre-commit) content
//   node scripts/scan-secrets.mjs file1 file2   scan specific files
//
// Suppress a known-safe match by putting the marker  allowlist-secret  on the
// line (e.g. a synthetic test fixture). Exit code 1 on any finding.

import { execSync } from 'node:child_process';
import { readFileSync, existsSync } from 'node:fs';

const PRAGMA = 'allowlist-secret';

// Paths never scanned (vendored code, build output, binaries, the lockfile, and
// this scanner's own pattern definitions).
const IGNORE = [
  /^node_modules\//,
  /^dist\//,
  /(^|\/)package-lock\.json$/,
  /\.dxt$/,
  /(^|\/)scan-secrets\.mjs$/,
  /\.(png|jpg|jpeg|gif|ico|pdf|zip|gz|lock)$/i,
];

// Domains considered non-personal placeholders/services. An email on any other
// domain is flagged as possible real PII. This list is intentionally generic —
// it names no personal domains.
const SAFE_EMAIL_DOMAINS = new Set([
  'example.com', 'example.org', 'example.net', 'test.com', 'localhost',
  'fastmail.com', 'api.fastmail.com', 'caldav.fastmail.com',
  'www.fastmailusercontent.com', 'fastmailusercontent.com',
  'anthropic.com',
  // short placeholder domains used in coerce comma-split tests
  'b.com', 'd.com', 'f.com', 'h.com', 'j.com', 'l.com',
  // adversary/other placeholders that are semantically meaningful in tests
  'evil.com', 'other.com',
]);

const RULES = [
  { name: 'Fastmail API token', re: /fmu\d+-[A-Za-z0-9_-]{20,}/g },
  { name: 'Bearer credential', re: /Bearer\s+[A-Za-z0-9._~+/=-]{16,}/g },
  { name: 'Basic auth credential', re: /Basic\s+[A-Za-z0-9+/=]{16,}/g },
  {
    name: 'hardcoded secret assignment',
    re: /\b(?:api[_-]?key|secret|password|passwd|token)\b\s*[:=]\s*["'`][A-Za-z0-9_\-]{16,}["'`]/gi,
    // skip env refs / obvious placeholders
    skip: (m) => /\$\{|process\.env|your-|REDACTED|example|placeholder|x{6,}|0{6,}|changeme|<[^>]+>/i.test(m),
  },
];

const EMAIL_RE = /[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g;

// Optional local denylist (gitignored): one literal string or domain per line.
// Lets a developer catch their own personal domains without publishing them.
function loadLocalDenylist() {
  const path = '.secret-scan-local.txt';
  if (!existsSync(path)) return [];
  return readFileSync(path, 'utf8')
    .split('\n')
    .map((l) => l.trim())
    .filter((l) => l && !l.startsWith('#'));
}

function targetFiles() {
  const mode = process.argv[2];
  if (mode === '--all') {
    return execSync('git ls-files', { encoding: 'utf8' }).split('\n').filter(Boolean);
  }
  if (mode === '--staged') {
    return execSync('git diff --cached --name-only --diff-filter=ACM', { encoding: 'utf8' })
      .split('\n').filter(Boolean);
  }
  return process.argv.slice(2);
}

const denylist = loadLocalDenylist();
const findings = [];

for (const file of targetFiles()) {
  if (IGNORE.some((re) => re.test(file))) continue;
  let text;
  try { text = readFileSync(file, 'utf8'); } catch { continue; }
  if (/\x00/.test(text)) continue; // skip binary files

  text.split('\n').forEach((line, i) => {
    if (line.includes(PRAGMA)) return;
    const lineNo = i + 1;

    for (const rule of RULES) {
      const matches = line.match(rule.re);
      if (!matches) continue;
      for (const m of matches) {
        if (rule.skip && rule.skip(m)) continue;
        findings.push({ file, lineNo, rule: rule.name, snippet: m.slice(0, 40) });
      }
    }

    const emails = line.match(EMAIL_RE) || [];
    for (const e of emails) {
      const domain = e.split('@')[1].toLowerCase();
      if (!SAFE_EMAIL_DOMAINS.has(domain)) {
        findings.push({ file, lineNo, rule: 'possible personal email', snippet: e });
      }
    }

    for (const bad of denylist) {
      if (line.toLowerCase().includes(bad.toLowerCase())) {
        findings.push({ file, lineNo, rule: 'local denylist match', snippet: bad });
      }
    }
  });
}

if (findings.length === 0) {
  console.log('✓ secret/PII scan clean');
  process.exit(0);
}

console.error(`\n✗ secret/PII scan found ${findings.length} issue(s):\n`);
for (const f of findings) {
  console.error(`  ${f.file}:${f.lineNo}  [${f.rule}]  ${f.snippet}`);
}
console.error(`
If a match is a deliberate synthetic fixture (not a real secret), append the
marker "${PRAGMA}" to that line. If it is a real credential or personal
datum, remove it — and if it was ever committed, rotate/revoke it.
`);
process.exit(1);
