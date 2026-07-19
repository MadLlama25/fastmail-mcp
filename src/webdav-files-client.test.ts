import { describe, it, mock } from 'node:test';
import assert from 'node:assert/strict';
import { WebDAVFilesClient, sanitizeRemotePath, filesAvailabilitySection } from './webdav-files-client.js';
import { validateHttpsUrl } from './url-validation.js';

// ---------- sanitizeRemotePath ----------

describe('sanitizeRemotePath', () => {
  it('accepts clean nested paths and returns raw segments', () => {
    assert.deepEqual(sanitizeRemotePath('invoices/2026/receipt.pdf'), ['invoices', '2026', 'receipt.pdf']);
    assert.deepEqual(sanitizeRemotePath('file.txt'), ['file.txt']);
    assert.deepEqual(sanitizeRemotePath('name with spaces & symbols!.txt'), ['name with spaces & symbols!.txt']);
  });

  const rejects: Array<[string, any, RegExp]> = [
    ['empty', '', /required/],
    ['absolute path', '/etc/passwd', /relative/],
    ['parent traversal', '../secrets.txt', /\.\. segments|must not contain/],
    ['embedded traversal', 'a/../../b', /must not contain/],
    ['dot segment', 'a/./b', /must not contain/],
    ['double slash', 'a//b', /empty segment/],
    ['trailing slash', 'a/b/', /empty segment/],
    ['backslash', 'a\\b.txt', /forward slashes/],
    ['URL smuggling', 'https://evil.example/x', /not a URL/],
    ['NUL byte', 'a\u0000b', /control characters/],
    ['newline', 'a\nb', /control characters/],
    ['DEL char', 'a\u007fb', /control characters/],
    ['overlong path', 'x'.repeat(1025), /too long/],
    ['too many segments', Array(33).fill('d').join('/'), /too many segments/],
    ['overlong segment', 'y'.repeat(256), /255 bytes/],
  ];
  for (const [label, input, pattern] of rejects) {
    it(`rejects ${label}`, () => {
      assert.throws(() => sanitizeRemotePath(input), pattern);
    });
  }

  it('percent-encoded traversal survives as literal text, not traversal', () => {
    // %2e%2e is validated as literal characters, then encoded again on URL
    // build — it can never decode into ".." server-side.
    assert.deepEqual(sanitizeRemotePath('%2e%2e/file.txt'), ['%2e%2e', 'file.txt']);
  });
});

// ---------- validateHttpsUrl ----------

describe('validateHttpsUrl', () => {
  it('accepts a plain HTTPS collection URL', () => {
    assert.equal(validateHttpsUrl('https://myfiles.fastmail.com/', 'x').hostname, 'myfiles.fastmail.com');
    assert.equal(validateHttpsUrl('https://cloud.example.com/remote.php/dav/files/user/', 'x').pathname, '/remote.php/dav/files/user/');
  });
  it('rejects HTTP', () => {
    assert.throws(() => validateHttpsUrl('http://myfiles.fastmail.com/', 'x'), /HTTPS/);
  });
  it('rejects embedded userinfo', () => {
    assert.throws(() => validateHttpsUrl('https://user:pass@host.example/', 'x'), /embed credentials/); // allowlist-secret (synthetic URL fixture)
  });
  it('rejects query and fragment', () => {
    assert.throws(() => validateHttpsUrl('https://host.example/dav/?x=1', 'x'), /query or fragment/);
    assert.throws(() => validateHttpsUrl('https://host.example/dav/#frag', 'x'), /query or fragment/);
  });
});

// ---------- WebDAVFilesClient.uploadBuffer ----------

describe('WebDAVFilesClient.uploadBuffer', () => {
  function makeClient() {
    const client = new WebDAVFilesClient({
      baseUrl: 'https://files.example.com/dav', // no trailing slash on purpose
      username: 'user@example.com',
      password: 'app-password',
    });
    const dav = {
      createObject: mock.fn(async () => new Response(null, { status: 201, statusText: 'Created' })),
      // PROPFIND existence pre-check → absent; MKCOL → created.
      davRequest: mock.fn(async ({ init }: any) => [{ status: init.method === 'PROPFIND' ? 404 : 201 }]),
    };
    (client as any).dav = dav;
    const callsBy = (method: string) => dav.davRequest.mock.calls.filter((c: any) => c.arguments[0].init.method === method);
    return { client, dav, callsBy };
  }

  it('PUTs to base + encoded segments with If-None-Match by default', async () => {
    const { client, dav } = makeClient();
    const result = await client.uploadBuffer(Buffer.from('data'), 'sub dir/räport.pdf', { contentType: 'application/pdf' });

    const call = dav.createObject.mock.calls[0].arguments[0];
    assert.equal(call.url, 'https://files.example.com/dav/sub%20dir/r%C3%A4port.pdf');
    assert.equal(call.headers['If-None-Match'], '*');
    assert.equal(call.headers['Content-Type'], 'application/pdf');
    assert.ok(call.headers['authorization'] || call.headers['Authorization'], 'must send Basic auth header');
    assert.deepEqual(result, { savedPath: 'sub dir/räport.pdf', host: 'files.example.com', bytes: 4, contentType: 'application/pdf' });
  });

  it('overwrite: true drops the If-None-Match precondition', async () => {
    const { client, dav } = makeClient();
    await client.uploadBuffer(Buffer.from('x'), 'a.txt', { overwrite: true });
    assert.equal('If-None-Match' in dav.createObject.mock.calls[0].arguments[0].headers, false);
  });

  it('maps 412 to the exists/overwrite guidance', async () => {
    const { client } = makeClient();
    (client as any).dav.createObject = mock.fn(async () => new Response(null, { status: 412, statusText: 'Precondition Failed' }));
    await assert.rejects(() => client.uploadBuffer(Buffer.from('x'), 'a.txt'), /already exists.*overwrite/s);
  });

  it('creates parent collections via MKCOL, tolerating 405 (exists)', async () => {
    const { client, dav, callsBy } = makeClient();
    dav.davRequest.mock.mockImplementation(async ({ init }: any) =>
      [{ status: init.method === 'PROPFIND' ? 404 : 405 }]);
    await client.uploadBuffer(Buffer.from('x'), 'a/b/c.txt');
    const mkcols = callsBy('MKCOL');
    assert.equal(mkcols.length, 2); // a/, a/b/
    assert.equal(mkcols[0].arguments[0].url, 'https://files.example.com/dav/a/');
  });

  it('pre-checks existence with PROPFIND and rejects when present (servers may ignore If-None-Match)', async () => {
    const { client, dav } = makeClient();
    dav.davRequest.mock.mockImplementation(async ({ init }: any) =>
      [{ status: init.method === 'PROPFIND' ? 207 : 201 }]);
    await assert.rejects(() => client.uploadBuffer(Buffer.from('x'), 'a.txt'), /already exists.*overwrite/s);
    assert.equal(dav.createObject.mock.calls.length, 0);
  });

  it('overwrite: true skips the existence pre-check entirely', async () => {
    const { client, callsBy } = makeClient();
    await client.uploadBuffer(Buffer.from('x'), 'a.txt', { overwrite: true });
    assert.equal(callsBy('PROPFIND').length, 0);
  });

  it('fails hard when MKCOL returns an unexpected status', async () => {
    const { client } = makeClient();
    (client as any).dav.davRequest = mock.fn(async ({ init }: any) =>
      [{ status: init.method === 'PROPFIND' ? 404 : 403 }]);
    await assert.rejects(() => client.uploadBuffer(Buffer.from('x'), 'a/b.txt'), /Could not create remote directory/);
  });

  it('skips MKCOL when createParents is false', async () => {
    const { client, callsBy } = makeClient();
    await client.uploadBuffer(Buffer.from('x'), 'a/b/c.txt', { createParents: false });
    assert.equal(callsBy('MKCOL').length, 0);
  });

  it('rejects unsanitary paths before any network call', async () => {
    const { client, dav } = makeClient();
    await assert.rejects(() => client.uploadBuffer(Buffer.from('x'), '../escape.txt'), /must not contain|relative/);
    assert.equal(dav.createObject.mock.calls.length, 0);
    assert.equal(dav.davRequest.mock.calls.length, 0);
  });
});

// ---------- filesAvailabilitySection ----------

describe('filesAvailabilitySection', () => {
  it('reports available with no guide when configured and healthy', () => {
    const s = filesAvailabilitySection(true);
    assert.equal(s.available, true);
    assert.deepEqual(s.functions, ['save_attachment_to_webdav']);
    assert.equal(s.enablementGuide, null);
    assert.match(s.note, /configured/);
  });

  it('reports unavailable with setup guide when unconfigured', () => {
    const s = filesAvailabilitySection(false);
    assert.equal(s.available, false);
    assert.match(s.note, /not configured/);
    assert.ok(s.enablementGuide && s.enablementGuide.steps.some((x) => x.includes('myfiles.fastmail.com')));
  });

  it('configured-but-broken reports unavailable with the error, not a throw', () => {
    const s = filesAvailabilitySection(true, 'FASTMAIL_WEBDAV_URL must use HTTPS (got: http:).');
    assert.equal(s.available, false);
    assert.match(s.note, /invalid.*HTTPS/s);
    assert.ok(s.enablementGuide, 'broken config should still include the guide');
  });
});
