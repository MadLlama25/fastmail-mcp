/**
 * Identity signature handling.
 *
 * JMAP identities carry `textSignature` and `htmlSignature` (RFC 8621 §6.1). The
 * Fastmail web client appends them to every message it composes, so a signature
 * configured in Settings is part of what "sending as this identity" means.
 *
 * The send paths already resolve an identity in order to pick `identityId`; these
 * helpers let them apply that identity's signature to the body as well, instead of
 * discarding the fields and emitting unsigned mail.
 *
 * Appending is idempotent — a body that already ends with the signature is returned
 * unchanged (trailing whitespace ignored on both sides, since appending trims it).
 * That keeps re-signing safe when a draft is edited or a server-returned body is
 * replayed, rather than stacking a second copy.
 */

/**
 * Tags that carry visible content even with no accompanying text — a logo-only
 * signature is a real signature, so a naive "strip tags, is it blank?" test would
 * wrongly discard it.
 */
const CONTENTFUL_HTML_TAG = /<(?:img|a|table|hr|video|audio|object|iframe|svg|embed|picture|source)\b/i;

/** Entities that render as whitespace and so must not count as signature content. */
const HTML_SPACE_ENTITY = /&(?:nbsp|#160|#xa0|ensp|emsp|thinsp|#8194|#8195|#8201);/gi;

/** A text signature of only whitespace (a stray newline, say) is nothing to append. */
export function hasTextSignature(signature?: string | null): boolean {
  return typeof signature === 'string' && signature.trim() !== '';
}

/**
 * An HTML signature is empty when it renders as nothing. Fastmail stores
 * `<div><br></div><div><br></div>` for identities whose signature was never set,
 * and appending that to every message would be pure noise.
 */
export function hasHtmlSignature(signature?: string | null): boolean {
  if (typeof signature !== 'string' || signature.trim() === '') return false;
  if (CONTENTFUL_HTML_TAG.test(signature)) return true;
  return signature.replace(/<[^>]*>/g, '').replace(HTML_SPACE_ENTITY, ' ').trim() !== '';
}

/**
 * Append a text signature, separated from the body by one blank line.
 *
 * An empty body is returned untouched: the callers treat "" as "no such part", and
 * signing it would synthesise a body part the message never had.
 */
export function appendTextSignature(body: string, signature?: string | null): string {
  if (!body || !hasTextSignature(signature)) return body;
  const sig = signature as string;
  if (body.trimEnd().endsWith(sig.trimEnd())) return body;
  return `${body.trimEnd()}\n\n${sig}`;
}

/**
 * Append an HTML signature. No separator is inserted — Fastmail stores signatures
 * as leading block elements, which already break the line.
 */
export function appendHtmlSignature(body: string, signature?: string | null): string {
  if (!body || !hasHtmlSignature(signature)) return body;
  const sig = signature as string;
  if (body.trimEnd().endsWith(sig.trimEnd())) return body;
  return `${body.trimEnd()}${sig}`;
}

/** The subset of a JMAP Identity this module reads. */
export interface IdentitySignature {
  textSignature?: string | null;
  htmlSignature?: string | null;
}

export interface SignedBodies {
  textBody?: string;
  htmlBody?: string;
}

/**
 * Apply an identity's signature to whichever bodies are present.
 *
 * A part that is absent stays absent: a plain-text message is never promoted to
 * multipart just because the identity happens to define an HTML signature.
 */
export function applyIdentitySignature(
  bodies: SignedBodies,
  identity?: IdentitySignature | null,
): SignedBodies {
  if (!identity) return bodies;
  return {
    textBody: bodies.textBody === undefined
      ? undefined
      : appendTextSignature(bodies.textBody, identity.textSignature),
    htmlBody: bodies.htmlBody === undefined
      ? undefined
      : appendHtmlSignature(bodies.htmlBody, identity.htmlSignature),
  };
}
