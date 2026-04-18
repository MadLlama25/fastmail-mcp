import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { coerceStringArray, coerceBool, redactBearerTokens } from './coerce.js';

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
    const out = redactBearerTokens(
      'Failed: token fmu1-3b1e4048-036f4f86690cd04d8d05105a369ee30b-0-dbfc727af72d5e3e27dd324675869337 invalid'
    );
    assert.match(out, /fmu\[REDACTED\]/);
    assert.ok(!out.includes('fmu1-3b1e'));
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
});
