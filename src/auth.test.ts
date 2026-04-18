import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { FastmailAuth } from './auth.js';

describe('FastmailAuth (URL validation)', () => {
  it('uses the default base URL when none provided', () => {
    const auth = new FastmailAuth({ apiToken: 'fake-token' });
    assert.equal(auth.getSessionUrl(), 'https://api.fastmail.com/jmap/session');
    assert.equal(auth.getApiUrl(), 'https://api.fastmail.com/jmap/api/');
    assert.equal(auth.getAllowUnsafe(), false);
  });

  it('accepts an allowlisted explicit base URL', () => {
    const auth = new FastmailAuth({
      apiToken: 'fake-token',
      baseUrl: 'https://api.fastmail.com',
    });
    assert.equal(auth.getApiUrl(), 'https://api.fastmail.com/jmap/api/');
  });

  it('rejects HTTP base URL', () => {
    assert.throws(
      () => new FastmailAuth({ apiToken: 'fake-token', baseUrl: 'http://api.fastmail.com' }),
      /must use HTTPS/,
    );
  });

  it('rejects non-allowlisted base URL by default', () => {
    assert.throws(
      () => new FastmailAuth({ apiToken: 'fake-token', baseUrl: 'https://attacker.example.com' }),
      /not in the Fastmail allowlist/,
    );
  });

  it('rejects bare-domain attacker hostnames', () => {
    assert.throws(
      () => new FastmailAuth({ apiToken: 'fake-token', baseUrl: 'attacker.example.com' }),
      /not in the Fastmail allowlist/,
    );
  });

  it('accepts arbitrary HTTPS base URL when allowUnsafeBaseUrl is true', () => {
    const auth = new FastmailAuth({
      apiToken: 'fake-token',
      baseUrl: 'https://jmap.self-hosted.example',
      allowUnsafeBaseUrl: true,
    });
    assert.equal(auth.getApiUrl(), 'https://jmap.self-hosted.example/jmap/api/');
    assert.equal(auth.getAllowUnsafe(), true);
  });

  it('still rejects HTTP even when allowUnsafeBaseUrl is true', () => {
    assert.throws(
      () => new FastmailAuth({
        apiToken: 'fake-token',
        baseUrl: 'http://jmap.self-hosted.example',
        allowUnsafeBaseUrl: true,
      }),
      /must use HTTPS/,
    );
  });

  it('strips trailing slashes from base URL', () => {
    const auth = new FastmailAuth({ apiToken: 'fake-token', baseUrl: 'https://api.fastmail.com///' });
    assert.equal(auth.getSessionUrl(), 'https://api.fastmail.com/jmap/session');
  });
});
