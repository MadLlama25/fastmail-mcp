// Some MCP clients (e.g. Claude Cowork as of 2026-04-08, issue #54) stringify
// structured params before dispatch. These helpers coerce such values back to
// their expected shapes so the handlers work against both strict and lenient clients.

// Defense-in-depth: scrub bearer-token-shaped substrings from any string that
// might be reflected back to the MCP caller (e.g. a JMAP error message). This
// is intentionally narrow — provider error messages are useful for the LLM to
// recover from, so we don't want to over-sanitize.
const BEARER_PATTERN = /Bearer\s+\S+/gi;
const FASTMAIL_TOKEN_PATTERN = /fmu\d+-[A-Za-z0-9-]{20,}/g;

export function redactBearerTokens(input: string): string {
  return input
    .replace(BEARER_PATTERN, 'Bearer [REDACTED]')
    .replace(FASTMAIL_TOKEN_PATTERN, 'fmu[REDACTED]');
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
