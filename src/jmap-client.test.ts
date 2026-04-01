import { describe, it, beforeEach, mock } from 'node:test';
import assert from 'node:assert/strict';
import { homedir } from 'os';
import { resolve, join } from 'path';
import { JmapClient } from './jmap-client.js';
import { FastmailAuth } from './auth.js';

// ---------- helpers ----------

const ACCOUNT_ID = 'acct-123';
const IDENTITY = { id: 'id-1', email: 'me@example.com', mayDelete: false };
const DRAFTS_MAILBOX = { id: 'mb-drafts', name: 'Drafts', role: 'drafts' };

function makeClient(): JmapClient {
  const auth = new FastmailAuth({ apiToken: 'fake-token' });
  const client = new JmapClient(auth);

  // Stub getSession so no network call is made
  mock.method(client, 'getSession', async () => ({
    apiUrl: 'https://api.example.com/jmap/api/',
    accountId: ACCOUNT_ID,
    capabilities: {},
  }));

  // Default stubs — tests override as needed
  mock.method(client, 'getIdentities', async () => [IDENTITY]);
  mock.method(client, 'getMailboxes', async () => [DRAFTS_MAILBOX]);

  return client;
}

function stubMakeRequest(client: JmapClient, response: any) {
  mock.method(client, 'makeRequest', async () => response);
}

// ---------- tests ----------

describe('createDraft', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  // 1. Happy path
  it('returns email ID on success', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['Email/set', { created: { draft: { id: 'email-42' } } }, 'createDraft'],
      ],
    });

    const id = await client.createDraft({ subject: 'Hello' });
    assert.equal(id, 'email-42');
  });

  // 2. Correct JMAP request structure
  it('sends correct JMAP request structure', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/set', { created: { draft: { id: 'email-1' } } }, 'createDraft'],
      ],
    }));

    await client.createDraft({ subject: 'Test', textBody: 'body' });

    assert.equal(makeReq.mock.calls.length, 1);
    const request = makeReq.mock.calls[0].arguments[0];

    // capabilities
    assert.deepEqual(request.using, [
      'urn:ietf:params:jmap:core',
      'urn:ietf:params:jmap:mail',
    ]);

    // method
    assert.equal(request.methodCalls[0][0], 'Email/set');

    // accountId
    assert.equal(request.methodCalls[0][1].accountId, ACCOUNT_ID);

    // email object shape
    const emailObj = request.methodCalls[0][1].create.draft;
    assert.equal(emailObj.subject, 'Test');
    assert.deepEqual(emailObj.from, [{ email: 'me@example.com' }]);
    assert.deepEqual(emailObj.keywords, { $draft: true });
    assert.equal(emailObj.mailboxIds[DRAFTS_MAILBOX.id], true);
  });

  // 3. Bug 1 regression — JMAP method-level error throws
  it('throws on JMAP method-level error', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['error', { type: 'unknownMethod', description: 'bad call' }, 'createDraft'],
      ],
    });

    await assert.rejects(
      () => client.createDraft({ subject: 'X' }),
      (err: Error) => {
        assert.match(err.message, /unknownMethod/);
        assert.match(err.message, /bad call/);
        return true;
      },
    );
  });

  // 4. Bug 2 regression — notCreated includes server type + description
  it('throws with server-provided error details from notCreated', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        [
          'Email/set',
          {
            notCreated: {
              draft: { type: 'invalidProperties', description: 'subject too long' },
            },
          },
          'createDraft',
        ],
      ],
    });

    await assert.rejects(
      () => client.createDraft({ subject: 'X' }),
      (err: Error) => {
        assert.match(err.message, /invalidProperties/);
        assert.match(err.message, /subject too long/);
        return true;
      },
    );
  });

  // 5. Bug 3 regression — missing created.draft.id throws
  it('throws when created.draft.id is missing', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['Email/set', { created: { draft: {} } }, 'createDraft'],
      ],
    });

    await assert.rejects(
      () => client.createDraft({ subject: 'X' }),
      (err: Error) => {
        assert.match(err.message, /no email ID/);
        return true;
      },
    );
  });

  // 6. Validation — empty input throws
  it('throws when no meaningful fields are provided', async () => {
    await assert.rejects(
      () => client.createDraft({}),
      (err: Error) => {
        assert.match(err.message, /at least one/i);
        return true;
      },
    );
  });

  // 7. Custom from address used correctly
  it('uses custom from address when provided', async () => {
    const altIdentity = { id: 'id-2', email: 'alias@example.com', mayDelete: true };
    mock.method(client, 'getIdentities', async () => [IDENTITY, altIdentity]);

    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/set', { created: { draft: { id: 'email-7' } } }, 'createDraft'],
      ],
    }));

    await client.createDraft({ subject: 'Hi', from: 'alias@example.com' });

    const emailObj = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].create.draft;
    assert.deepEqual(emailObj.from, [{ email: 'alias@example.com' }]);
  });

  // 8. Invalid from address throws
  it('throws when from address is not a verified identity', async () => {
    await assert.rejects(
      () => client.createDraft({ subject: 'Hi', from: 'nobody@example.com' }),
      (err: Error) => {
        assert.match(err.message, /not verified/i);
        return true;
      },
    );
  });

  // 9. Custom mailboxId used instead of auto-lookup
  it('uses provided mailboxId without looking up mailboxes', async () => {
    const getMailboxes = client.getMailboxes as ReturnType<typeof mock.method>;

    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/set', { created: { draft: { id: 'email-9' } } }, 'createDraft'],
      ],
    }));

    await client.createDraft({ subject: 'Custom', mailboxId: 'mb-custom' });

    // getMailboxes should not have been called
    assert.equal(getMailboxes.mock.calls.length, 0);

    const emailObj = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].create.draft;
    assert.equal(emailObj.mailboxIds['mb-custom'], true);
  });

  // 10. HTML body constructed correctly
  it('constructs HTML body parts correctly', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/set', { created: { draft: { id: 'email-10' } } }, 'createDraft'],
      ],
    }));

    await client.createDraft({ subject: 'Rich', htmlBody: '<p>Hello</p>' });

    const emailObj = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].create.draft;
    assert.deepEqual(emailObj.htmlBody, [{ partId: 'html', type: 'text/html' }]);
    assert.equal(emailObj.textBody, undefined);
    assert.deepEqual(emailObj.bodyValues, { html: { value: '<p>Hello</p>' } });
  });
});

// ---------- updateDraft ----------

const EXISTING_DRAFT = {
  id: 'draft-1',
  subject: 'Old Subject',
  from: [{ email: 'me@example.com' }],
  to: [{ email: 'bob@example.com' }],
  cc: [],
  bcc: [],
  textBody: [{ partId: 'text', type: 'text/plain' }],
  htmlBody: null,
  bodyValues: { text: { value: 'Old body' } },
  mailboxIds: { 'mb-drafts': true },
  keywords: { $draft: true },
};

describe('updateDraft', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('returns new email ID on success', async () => {
    const makeReq = mock.method(client, 'makeRequest', async (req: any) => {
      // First call: Email/get to fetch existing draft
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [EXISTING_DRAFT] }, 'getEmail']] };
      }
      // Second call: Email/set with create+destroy
      return { methodResponses: [['Email/set', { created: { draft: { id: 'draft-2' } }, destroyed: ['draft-1'] }, 'updateDraft']] };
    });

    const newId = await client.updateDraft('draft-1', { subject: 'New Subject' });
    assert.equal(newId, 'draft-2');

    // Verify the second call has both create and destroy
    const setCall = makeReq.mock.calls[1].arguments[0];
    assert.equal(setCall.methodCalls[0][0], 'Email/set');
    assert.deepEqual(setCall.methodCalls[0][1].destroy, ['draft-1']);
    assert.equal(setCall.methodCalls[0][1].create.draft.subject, 'New Subject');
  });

  it('merges fields — preserves existing values for unspecified fields', async () => {
    mock.method(client, 'makeRequest', async (req: any) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [EXISTING_DRAFT] }, 'getEmail']] };
      }
      return { methodResponses: [['Email/set', { created: { draft: { id: 'draft-2' } }, destroyed: ['draft-1'] }, 'updateDraft']] };
    });

    await client.updateDraft('draft-1', { subject: 'Updated' });

    // The create call should keep existing to address
    const makeReq = client.makeRequest as ReturnType<typeof mock.method>;
    const emailObj = makeReq.mock.calls[1].arguments[0].methodCalls[0][1].create.draft;
    assert.deepEqual(emailObj.to, [{ email: 'bob@example.com' }]);
    assert.equal(emailObj.subject, 'Updated');
  });

  it('rejects non-draft email', async () => {
    const nonDraft = { ...EXISTING_DRAFT, keywords: { $seen: true } };
    mock.method(client, 'makeRequest', async () => ({
      methodResponses: [['Email/get', { list: [nonDraft] }, 'getEmail']],
    }));

    await assert.rejects(
      () => client.updateDraft('email-1', { subject: 'X' }),
      (err: Error) => {
        assert.match(err.message, /non-draft/i);
        return true;
      },
    );
  });

  it('throws when email not found', async () => {
    mock.method(client, 'makeRequest', async () => ({
      methodResponses: [['Email/get', { list: [] }, 'getEmail']],
    }));

    await assert.rejects(
      () => client.updateDraft('missing-id', { subject: 'X' }),
      (err: Error) => {
        assert.match(err.message, /not found/i);
        return true;
      },
    );
  });

  it('throws on JMAP error during create+destroy', async () => {
    mock.method(client, 'makeRequest', async (req: any) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [EXISTING_DRAFT] }, 'getEmail']] };
      }
      return { methodResponses: [['error', { type: 'serverFail', description: 'oops' }, 'updateDraft']] };
    });

    await assert.rejects(
      () => client.updateDraft('draft-1', { subject: 'X' }),
      (err: Error) => {
        assert.match(err.message, /serverFail/);
        return true;
      },
    );
  });
});

// ---------- sendDraft ----------

const SENDABLE_DRAFT = {
  id: 'draft-1',
  from: [{ email: 'me@example.com' }],
  to: [{ email: 'bob@example.com' }],
  cc: [{ email: 'cc@example.com' }],
  bcc: [],
  keywords: { $draft: true },
};

const SENT_MAILBOX = { id: 'mb-sent', name: 'Sent', role: 'sent' };

describe('sendDraft', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
    mock.method(client, 'getMailboxes', async () => [DRAFTS_MAILBOX, SENT_MAILBOX]);
  });

  it('returns submission ID on success', async () => {
    const makeReq = mock.method(client, 'makeRequest', async (req: any) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [SENDABLE_DRAFT] }, 'getEmail']] };
      }
      return { methodResponses: [['EmailSubmission/set', { created: { submission: { id: 'sub-1' } } }, 'submitDraft']] };
    });

    const subId = await client.sendDraft('draft-1');
    assert.equal(subId, 'sub-1');

    // Verify submission call structure
    const submitCall = makeReq.mock.calls[1].arguments[0];
    assert.equal(submitCall.methodCalls[0][0], 'EmailSubmission/set');
    assert.equal(submitCall.methodCalls[0][1].create.submission.emailId, 'draft-1');
    assert.equal(submitCall.methodCalls[0][1].create.submission.identityId, IDENTITY.id);

    // Verify envelope has all recipients (to + cc)
    const rcptTo = submitCall.methodCalls[0][1].create.submission.envelope.rcptTo;
    assert.equal(rcptTo.length, 2);
    assert.deepEqual(rcptTo[0], { email: 'bob@example.com' });
    assert.deepEqual(rcptTo[1], { email: 'cc@example.com' });
  });

  it('rejects non-draft email', async () => {
    const nonDraft = { ...SENDABLE_DRAFT, keywords: { $seen: true } };
    mock.method(client, 'makeRequest', async () => ({
      methodResponses: [['Email/get', { list: [nonDraft] }, 'getEmail']],
    }));

    await assert.rejects(
      () => client.sendDraft('email-1'),
      (err: Error) => {
        assert.match(err.message, /non-draft/i);
        return true;
      },
    );
  });

  it('rejects draft with no recipients', async () => {
    const noRecipients = { ...SENDABLE_DRAFT, to: [], cc: [], bcc: [] };
    mock.method(client, 'makeRequest', async () => ({
      methodResponses: [['Email/get', { list: [noRecipients] }, 'getEmail']],
    }));

    await assert.rejects(
      () => client.sendDraft('draft-1'),
      (err: Error) => {
        assert.match(err.message, /no recipients/i);
        return true;
      },
    );
  });

  it('throws when email not found', async () => {
    mock.method(client, 'makeRequest', async () => ({
      methodResponses: [['Email/get', { list: [] }, 'getEmail']],
    }));

    await assert.rejects(
      () => client.sendDraft('missing-id'),
      (err: Error) => {
        assert.match(err.message, /not found/i);
        return true;
      },
    );
  });

  it('throws on JMAP submission error', async () => {
    mock.method(client, 'makeRequest', async (req: any) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [SENDABLE_DRAFT] }, 'getEmail']] };
      }
      return { methodResponses: [['error', { type: 'forbidden', description: 'not allowed' }, 'submitDraft']] };
    });

    await assert.rejects(
      () => client.sendDraft('draft-1'),
      (err: Error) => {
        assert.match(err.message, /forbidden/);
        return true;
      },
    );
  });
});

// ---------- JMAP response validation ----------

describe('JMAP response validation', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('throws when methodResponses is missing', async () => {
    stubMakeRequest(client, { sessionState: 's1' });
    await assert.rejects(
      () => client.getEmailById('email-1'),
      (err: Error) => {
        assert.match(err.message, /missing expected method/i);
        return true;
      },
    );
  });

  it('throws when index exceeds methodResponses length', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['error', { type: 'serverFail', description: 'oops' }, 'query'],
      ],
    });
    // getEmails uses getListResult(response, 1) but only 1 response exists
    await assert.rejects(
      () => client.getEmails(undefined, 10),
      (err: Error) => {
        assert.ok(err.message.length > 0);
        return true;
      },
    );
  });

  it('throws on malformed methodResponses entry', async () => {
    stubMakeRequest(client, {
      methodResponses: ['not-a-tuple' as any],
    });
    await assert.rejects(
      () => client.getEmailById('email-1'),
      (err: Error) => {
        assert.match(err.message, /malformed/i);
        return true;
      },
    );
  });
});

// ---------- searchEmails ----------

describe('searchEmails', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('returns email list on success', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['Email/query', { ids: ['e1'] }, 'query'],
        ['Email/get', { list: [{ id: 'e1', subject: 'Test' }] }, 'emails'],
      ],
    });
    const results = await client.searchEmails('test', 10);
    assert.equal(results.length, 1);
    assert.equal(results[0].subject, 'Test');
  });

  it('returns empty array when no results', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['Email/query', { ids: [] }, 'query'],
        ['Email/get', { list: [] }, 'emails'],
      ],
    });
    const results = await client.searchEmails('nonexistent');
    assert.deepEqual(results, []);
  });

  it('throws on JMAP error in query', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['error', { type: 'invalidArguments', description: 'bad filter' }, 'query'],
        ['error', { type: 'invalidArguments', description: 'bad filter' }, 'emails'],
      ],
    });
    await assert.rejects(
      () => client.searchEmails('test'),
      (err: Error) => {
        assert.match(err.message, /invalidArguments/);
        return true;
      },
    );
  });
});

// ---------- validateSavePath tests ----------

describe('validateSavePath', () => {
  const allowedDir = resolve(homedir(), 'Downloads', 'fastmail-mcp');

  it('accepts paths within the allowed directory', () => {
    const result = JmapClient.validateSavePath(`${allowedDir}/photo.jpg`);
    assert.equal(result, resolve(allowedDir, 'photo.jpg'));
  });

  it('accepts paths in subdirectories', () => {
    const result = JmapClient.validateSavePath(`${allowedDir}/andrew/assets/logo.png`);
    assert.equal(result, resolve(allowedDir, 'andrew', 'assets', 'logo.png'));
  });

  it('rejects paths outside the allowed directory', () => {
    assert.throws(
      () => JmapClient.validateSavePath('/tmp/evil.sh'),
      (err: Error) => {
        assert.match(err.message, /must be within/);
        return true;
      },
    );
  });

  it('rejects path traversal attempts', () => {
    assert.throws(
      () => JmapClient.validateSavePath(`${allowedDir}/../../../.bashrc`),
      (err: Error) => {
        assert.match(err.message, /must be within/);
        return true;
      },
    );
  });

  it('rejects home directory writes', () => {
    assert.throws(
      () => JmapClient.validateSavePath(`${homedir()}/.ssh/authorized_keys`),
      (err: Error) => {
        assert.match(err.message, /must be within/);
        return true;
      },
    );
  });

  it('rejects null bytes', () => {
    assert.throws(
      () => JmapClient.validateSavePath(`${allowedDir}/file\0.txt`),
      (err: Error) => {
        assert.match(err.message, /null bytes/);
        return true;
      },
    );
  });

  it('allows any path when allowAnyPath is true', () => {
    const result = JmapClient.validateSavePath('/tmp/any/path/file.txt', true);
    assert.equal(result, resolve('/tmp/any/path/file.txt'));
  });

  it('resolves bare filenames to downloads dir in restricted mode', () => {
    const result = JmapClient.validateSavePath('report.pdf');
    assert.equal(result, join(JmapClient.DEFAULT_DOWNLOADS_DIR, 'report.pdf'));
  });

  it('resolves relative paths to downloads dir in restricted mode', () => {
    const result = JmapClient.validateSavePath('2026/march/report.pdf');
    assert.equal(result, join(JmapClient.DEFAULT_DOWNLOADS_DIR, '2026', 'march', 'report.pdf'));
  });

  it('resolves relative paths against CWD when allowAnyPath is true', () => {
    const result = JmapClient.validateSavePath('output/file.txt', true);
    assert.equal(result, resolve('output/file.txt'));
  });

  it('rejects traversal escaping downloads dir even with allowAnyPath false', () => {
    assert.throws(
      () => JmapClient.validateSavePath(`${allowedDir}/../../etc/passwd`, false),
      (err: Error) => {
        assert.match(err.message, /must be within/);
        return true;
      },
    );
  });

  it('error message mentions FASTMAIL_ALLOW_ANY_PATH', () => {
    assert.throws(
      () => JmapClient.validateSavePath('/tmp/evil.sh'),
      (err: Error) => {
        assert.match(err.message, /FASTMAIL_ALLOW_ANY_PATH/);
        return true;
      },
    );
  });
});
