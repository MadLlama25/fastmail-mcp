import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { coerceStringArray, coerceRecipients, coerceBool, redactBearerTokens, registerSecret } from './coerce.js';

describe('coerceStringArray', () => {
  it('returns undefined for undefined input', () => {
    assert.equal(coerceStringArray(undefined), undefined);
  });

  it('returns undefined for null input', () => {
    assert.equal(coerceStringArray(null), undefined);
  });

  it('returns undefined for non-array, non-string input', () => {
    assert.equal(coerceStringArray(123), undefined);
    assert.equal(coerceStringArray({}), undefined);
    assert.equal(coerceStringArray(true), undefined);
  });

  it('returns array as-is', () => {
    assert.deepEqual(coerceStringArray(['a@b.com', 'c@d.com']), ['a@b.com', 'c@d.com']);
  });

  it('stringifies array elements', () => {
    assert.deepEqual(coerceStringArray([1, 2, 3] as any), ['1', '2', '3']);
  });

  it('parses JSON-stringified array', () => {
    assert.deepEqual(coerceStringArray('["a@b.com", "c@d.com"]'), ['a@b.com', 'c@d.com']);
  });

  it('parses JSON-stringified array with whitespace', () => {
    assert.deepEqual(coerceStringArray('  ["a@b.com"]  '), ['a@b.com']);
  });

  it('splits comma-separated string', () => {
    assert.deepEqual(coerceStringArray('a@b.com, c@d.com'), ['a@b.com', 'c@d.com']);
  });

  it('wraps single address as one-item array', () => {
    assert.deepEqual(coerceStringArray('single@example.com'), ['single@example.com']);
  });

  it('returns empty array for empty string', () => {
    assert.deepEqual(coerceStringArray(''), []);
  });

  it('trims whitespace and filters empty segments in comma-split', () => {
    assert.deepEqual(coerceStringArray('a@b.com, ,c@d.com,'), ['a@b.com', 'c@d.com']);
  });

  it('falls back to comma-split when JSON parsing fails', () => {
    assert.deepEqual(coerceStringArray('[not valid json]'), ['[not valid json]']);
  });
});

describe('coerceRecipients', () => {
  it('coerces all four fields from arrays, JSON-strings, comma-strings, and bare strings', () => {
    const result = coerceRecipients({
      to: ['a@b.com'],
      cc: '["c@d.com", "e@f.com"]',
      bcc: 'g@h.com, i@j.com',
      replyTo: 'k@l.com',
    });
    assert.deepEqual(result, {
      to: ['a@b.com'],
      cc: ['c@d.com', 'e@f.com'],
      bcc: ['g@h.com', 'i@j.com'],
      replyTo: ['k@l.com'],
    });
  });

  it('coerces empty string to empty array for each field', () => {
    assert.deepEqual(coerceRecipients({ to: '', cc: '', bcc: '', replyTo: '' }), {
      to: [],
      cc: [],
      bcc: [],
      replyTo: [],
    });
  });

  it('returns undefined for omitted fields', () => {
    assert.deepEqual(coerceRecipients({}), {
      to: undefined,
      cc: undefined,
      bcc: undefined,
      replyTo: undefined,
    });
  });

  it('returns undefined for non-string, non-array values', () => {
    assert.deepEqual(coerceRecipients({ to: 123, cc: {}, bcc: true, replyTo: null } as any), {
      to: undefined,
      cc: undefined,
      bcc: undefined,
      replyTo: undefined,
    });
  });
});

describe('coerceBool', () => {
  it('returns boolean as-is', () => {
    assert.equal(coerceBool(true), true);
    assert.equal(coerceBool(false), false);
  });

  it('coerces "true" string to true', () => {
    assert.equal(coerceBool('true'), true);
  });

  it('coerces "false" string to false', () => {
    assert.equal(coerceBool('false'), false);
  });

  it('returns undefined for unrecognized strings', () => {
    assert.equal(coerceBool('yes'), undefined);
    assert.equal(coerceBool('1'), undefined);
    assert.equal(coerceBool(''), undefined);
  });

  it('returns undefined for null/undefined', () => {
    assert.equal(coerceBool(undefined), undefined);
    assert.equal(coerceBool(null), undefined);
  });

  it('returns undefined for numbers', () => {
    assert.equal(coerceBool(1), undefined);
    assert.equal(coerceBool(0), undefined);
  });
});

describe('redactBearerTokens', () => {
  it('redacts Bearer header pattern', () => {
    const out = redactBearerTokens('Authorization: Bearer abc.def.ghi failed');
    assert.equal(out, 'Authorization: Bearer [REDACTED] failed');
  });

  it('redacts case-insensitive Bearer', () => {
    assert.equal(redactBearerTokens('bearer secret'), 'Bearer [REDACTED]');
    assert.equal(redactBearerTokens('BEARER xyz'), 'Bearer [REDACTED]');
  });

  it('redacts Fastmail token shape (fmu...)', () => {
    // Synthetic value — matches the fmuN-<hex>-<hex>-N-<hex> shape only, never a real token.
    const out = redactBearerTokens(
      'Failed: token fmu0-00000000-1111111111111111111111111111111a-0-2222222222222222222222222222222b invalid' // allowlist-secret (synthetic)
    );
    assert.match(out, /fmu\[REDACTED\]/);
    assert.ok(!out.includes('fmu0-0000'));
  });

  it('does not redact unrelated text', () => {
    const original = 'JMAP error: invalidArguments — mailbox not found';
    assert.equal(redactBearerTokens(original), original);
  });

  it('redacts multiple tokens in one string', () => {
    const out = redactBearerTokens('Bearer one and Bearer two');
    assert.equal(out, 'Bearer [REDACTED] and Bearer [REDACTED]');
  });

  it('handles empty string', () => {
    assert.equal(redactBearerTokens(''), '');
  });

  it('redacts Basic auth credentials (CalDAV path)', () => {
    const out = redactBearerTokens('401 on Authorization: Basic dXNlcjpwYXNzd29yZA== failed'); // allowlist-secret (synthetic base64 of "user:password")
    assert.match(out, /Basic \[REDACTED\]/);
    assert.ok(!out.includes('dXNlcjpwYXNz'));
  });

  it('redacts a Fastmail token containing an underscore fully (no tail leak)', () => {
    const out = redactBearerTokens('token fmu9-abcd1234-aaaaaaaaaaaaaaaaaaaa_bbbbbbbbbb invalid'); // allowlist-secret (synthetic)
    assert.match(out, /fmu\[REDACTED\]/);
    assert.ok(!out.includes('bbbbbbbbbb'), 'the post-underscore tail must not survive');
  });

  it('redacts an exact registered secret value even without a recognizable shape', () => {
    registerSecret('sk-pla1n-cr3dential-with-no-prefix');
    const out = redactBearerTokens('self-hosted auth failed for sk-pla1n-cr3dential-with-no-prefix here');
    assert.ok(!out.includes('sk-pla1n-cr3dential'));
    assert.match(out, /\[REDACTED\]/);
  });

  it('registerSecret ignores short/empty values (avoids over-broad matches)', () => {
    registerSecret('abc');
    registerSecret('');
    assert.equal(redactBearerTokens('the word abc should survive'), 'the word abc should survive');
  });
});
