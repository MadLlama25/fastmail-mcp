// Some MCP clients (e.g. Claude Cowork as of 2026-04-08, issue #54) stringify
// structured params before dispatch. These helpers coerce such values back to
// their expected shapes so the handlers work against both strict and lenient clients.

// Defense-in-depth: scrub credential-shaped substrings from any string that
// might be reflected back to the MCP caller (e.g. a JMAP error message). This
// is intentionally narrow — provider error messages are useful for the LLM to
// recover from, so we don't want to over-sanitize.
const BEARER_PATTERN = /Bearer\s+\S+/gi;
// Basic auth (CalDAV path uses HTTP Basic) — redact the base64 credential blob.
const BASIC_PATTERN = /Basic\s+[A-Za-z0-9+/=]+/gi;
// Fastmail token shape. `[\w-]` (not `[A-Za-z0-9-]`) so an underscore mid-token
// can't end the match early and leak the remaining entropy.
const FASTMAIL_TOKEN_PATTERN = /fmu\d+-[\w-]{20,}/g;

// Exact known secret values registered at startup (API token, CalDAV password,
// self-hosted tokens without a `Bearer`/`fmu` shape). Value-based redaction
// catches credentials the pattern-based rules would miss. Populated by
// registerSecret(); never logged.
const KNOWN_SECRETS = new Set<string>();

// Register a literal secret value so redactBearerTokens will scrub any exact
// occurrence of it. No-ops on empty/short values to avoid over-broad matches.
export function registerSecret(value: string | undefined): void {
  if (typeof value === 'string' && value.length >= 8) {
    KNOWN_SECRETS.add(value);
  }
}

function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

export function redactBearerTokens(input: string): string {
  let out = input
    .replace(BEARER_PATTERN, 'Bearer [REDACTED]')
    .replace(BASIC_PATTERN, 'Basic [REDACTED]')
    .replace(FASTMAIL_TOKEN_PATTERN, 'fmu[REDACTED]');
  for (const secret of KNOWN_SECRETS) {
    out = out.replace(new RegExp(escapeRegExp(secret), 'g'), '[REDACTED]');
  }
  return out;
}

export function coerceStringArray(value: unknown): string[] | undefined {
  if (value === undefined || value === null) return undefined;
  if (Array.isArray(value)) return value.map(String);
  if (typeof value !== 'string') return undefined;
  const trimmed = value.trim();
  if (!trimmed) return [];
  if (trimmed.startsWith('[') && trimmed.endsWith(']')) {
    try {
      const parsed = JSON.parse(trimmed);
      if (Array.isArray(parsed)) return parsed.map(String);
    } catch { /* fall through to comma-split */ }
  }
  return trimmed.split(',').map(s => s.trim()).filter(Boolean);
}

// Coerce the four recipient list fields from whatever shape a (possibly lenient)
// client sent into string[] | undefined, so the JMAP client's .map(parseAddress)
// calls never receive a bare string (issue #54). Pass the raw tool args; reads
// only to/cc/bcc/replyTo and returns the coerced quartet.
export function coerceRecipients(args: { to?: unknown; cc?: unknown; bcc?: unknown; replyTo?: unknown }): {
  to?: string[]; cc?: string[]; bcc?: string[]; replyTo?: string[];
} {
  return {
    to: coerceStringArray(args.to),
    cc: coerceStringArray(args.cc),
    bcc: coerceStringArray(args.bcc),
    replyTo: coerceStringArray(args.replyTo),
  };
}

export function coerceBool(value: unknown): boolean | undefined {
  if (typeof value === 'boolean') return value;
  if (value === 'true') return true;
  if (value === 'false') return false;
  return undefined;
}
