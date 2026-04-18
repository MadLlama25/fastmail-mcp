// Validates URLs that will receive the bearer token. Restricts to approved
// Fastmail origins by default, with an explicit opt-out for self-hosted JMAP.

const FASTMAIL_ALLOWED_HOSTS: ReadonlySet<string> = new Set([
  'api.fastmail.com',
  'www.fastmailusercontent.com',
]);

/**
 * Validate that a URL is acceptable for sending the bearer token to.
 *
 * Default policy:
 *   - Must be HTTPS.
 *   - Hostname must be in FASTMAIL_ALLOWED_HOSTS.
 *
 * When `allowUnsafe=true` (e.g. user opted in via FASTMAIL_ALLOW_UNSAFE_BASE_URL
 * for a self-hosted JMAP server):
 *   - Must still be HTTPS (plain HTTP is never allowed; the token would be sent
 *     in cleartext).
 *   - Any hostname is accepted.
 *
 * Throws on rejection; returns the parsed URL on success.
 */
export function validateFastmailUrl(input: string, fieldName: string, allowUnsafe = false): URL {
  let parsed: URL;
  try {
    parsed = new URL(input);
  } catch {
    throw new Error(`${fieldName} is not a valid URL`);
  }
  if (parsed.protocol !== 'https:') {
    throw new Error(
      `${fieldName} must use HTTPS (got: ${parsed.protocol}). ` +
      `Plain HTTP is rejected because the bearer token would be sent in cleartext.`,
    );
  }
  if (!allowUnsafe && !FASTMAIL_ALLOWED_HOSTS.has(parsed.hostname)) {
    throw new Error(
      `${fieldName} host '${parsed.hostname}' is not in the Fastmail allowlist. ` +
      `Set FASTMAIL_ALLOW_UNSAFE_BASE_URL=true to opt in for self-hosted JMAP servers.`,
    );
  }
  return parsed;
}

export const FASTMAIL_ALLOWED_HOSTS_FOR_TEST = FASTMAIL_ALLOWED_HOSTS;
