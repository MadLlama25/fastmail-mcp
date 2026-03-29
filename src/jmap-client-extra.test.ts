import assert from 'node:assert/strict';
import { beforeEach, describe, it, mock } from 'node:test';
import { FastmailAuth } from './auth.js';
import { JmapClient, type JmapRequest } from './jmap-client.js';

// ---------- helpers ----------

const ACCOUNT_ID = 'acct-123';
const INBOX_MAILBOX = {
  id: 'mb-inbox',
  name: 'Inbox',
  role: 'inbox',
  totalEmails: 42,
  unreadEmails: 5,
};
const DRAFTS_MAILBOX = { id: 'mb-drafts', name: 'Drafts', role: 'drafts' };
const TRASH_MAILBOX = { id: 'mb-trash', name: 'Trash', role: 'trash' };
const SENT_MAILBOX = { id: 'mb-sent', name: 'Sent', role: 'sent' };

function makeClient(): JmapClient {
  const auth = new FastmailAuth({ apiToken: 'fake-token' });
  const client = new JmapClient(auth);

  mock.method(client, 'getSession', async () => ({
    apiUrl: 'https://api.example.com/jmap/api/',
    accountId: ACCOUNT_ID,
    capabilities: {},
  }));

  return client;
}

function stubMakeRequest(client: JmapClient, response: Record<string, unknown>) {
  mock.method(client, 'makeRequest', async () => response);
}

function stubMailboxes(
  client: JmapClient,
  mailboxes: Record<string, unknown>[] = [
    INBOX_MAILBOX,
    DRAFTS_MAILBOX,
    TRASH_MAILBOX,
    SENT_MAILBOX,
  ],
) {
  mock.method(client, 'getMailboxes', async () => mailboxes);
}

// ---------- getMailboxes ----------

describe('getMailboxes', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('returns list of mailboxes on valid response', async () => {
    stubMakeRequest(client, {
      methodResponses: [['Mailbox/get', { list: [INBOX_MAILBOX, DRAFTS_MAILBOX] }, 'mailboxes']],
    });

    const mailboxes = await client.getMailboxes();
    assert.equal(mailboxes.length, 2);
    assert.equal(mailboxes[0].role, 'inbox');
    assert.equal(mailboxes[1].id, 'mb-drafts');
  });

  it('returns empty array when response list is missing', async () => {
    stubMakeRequest(client, {
      methodResponses: [['Mailbox/get', {}, 'mailboxes']],
    });

    const mailboxes = await client.getMailboxes();
    assert.deepEqual(mailboxes, []);
  });
});

// ---------- getRecentEmails ----------

describe('getRecentEmails', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('returns recent emails on valid response', async () => {
    stubMailboxes(client);
    stubMakeRequest(client, {
      methodResponses: [
        ['Email/query', { ids: ['e1', 'e2'] }, 'query'],
        [
          'Email/get',
          {
            list: [
              { id: 'e1', subject: 'First' },
              { id: 'e2', subject: 'Second' },
            ],
          },
          'emails',
        ],
      ],
    });

    const emails = await client.getRecentEmails(10, 'inbox');
    assert.equal(emails.length, 2);
    assert.equal(emails[0].subject, 'First');
  });

  it('throws when mailbox is not found', async () => {
    stubMailboxes(client, [INBOX_MAILBOX]);

    await assert.rejects(
      () => client.getRecentEmails(10, 'nonexistent'),
      (err: Error) => {
        assert.match(err.message, /Could not find mailbox/);
        return true;
      },
    );
  });

  it('matches mailbox by role', async () => {
    stubMailboxes(client, [{ id: 'mb-custom', name: 'My Inbox', role: 'inbox' }]);
    stubMakeRequest(client, {
      methodResponses: [
        ['Email/query', { ids: [] }, 'query'],
        ['Email/get', { list: [] }, 'emails'],
      ],
    });

    const emails = await client.getRecentEmails(5, 'inbox');
    assert.deepEqual(emails, []);
  });
});

// ---------- getEmails ----------

describe('getEmails', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('returns emails with mailboxId filter', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/query', { ids: ['e1'] }, 'query'],
        ['Email/get', { list: [{ id: 'e1', subject: 'Filtered' }] }, 'emails'],
      ],
    }));

    const emails = await client.getEmails('mb-inbox', 5);
    assert.equal(emails.length, 1);

    const filter = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].filter;
    assert.equal(filter.inMailbox, 'mb-inbox');
  });

  it('returns emails without mailboxId filter', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/query', { ids: ['e1'] }, 'query'],
        ['Email/get', { list: [{ id: 'e1', subject: 'All' }] }, 'emails'],
      ],
    }));

    const emails = await client.getEmails(undefined, 10);
    assert.equal(emails.length, 1);

    const filter = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].filter;
    assert.deepEqual(filter, {});
  });
});

// ---------- getEmailById ----------

describe('getEmailById', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('returns email on valid response', async () => {
    stubMakeRequest(client, {
      methodResponses: [['Email/get', { list: [{ id: 'e1', subject: 'Found' }] }, 'email']],
    });

    const email = await client.getEmailById('e1');
    assert.equal(email.id, 'e1');
    assert.equal(email.subject, 'Found');
  });

  it('throws when email is not found (empty list)', async () => {
    stubMakeRequest(client, {
      methodResponses: [['Email/get', { list: [] }, 'email']],
    });

    await assert.rejects(
      () => client.getEmailById('missing'),
      (err: Error) => {
        assert.match(err.message, /not found/);
        return true;
      },
    );
  });

  it('throws when email is in notFound list', async () => {
    stubMakeRequest(client, {
      methodResponses: [['Email/get', { list: [], notFound: ['gone'] }, 'email']],
    });

    await assert.rejects(
      () => client.getEmailById('gone'),
      (err: Error) => {
        assert.match(err.message, /not found/);
        return true;
      },
    );
  });
});

// ---------- moveEmail ----------

describe('moveEmail', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('moves email successfully', async () => {
    // First call: getEmail to read current mailboxIds
    // Second call: Email/set to move
    let callCount = 0;
    mock.method(client, 'makeRequest', async () => {
      callCount++;
      if (callCount === 1) {
        return {
          methodResponses: [
            ['Email/get', { list: [{ id: 'e1', mailboxIds: { 'mb-inbox': true } }] }, 'getEmail'],
          ],
        };
      }
      return {
        methodResponses: [['Email/set', { updated: { e1: null } }, 'moveEmail']],
      };
    });

    await client.moveEmail('e1', 'mb-archive');
    assert.equal(callCount, 2);
  });

  it('throws when update fails', async () => {
    let callCount = 0;
    mock.method(client, 'makeRequest', async () => {
      callCount++;
      if (callCount === 1) {
        return {
          methodResponses: [
            ['Email/get', { list: [{ id: 'e1', mailboxIds: { 'mb-inbox': true } }] }, 'getEmail'],
          ],
        };
      }
      return {
        methodResponses: [['Email/set', { notUpdated: { e1: { type: 'notFound' } } }, 'moveEmail']],
      };
    });

    await assert.rejects(
      () => client.moveEmail('e1', 'mb-archive'),
      (err: Error) => {
        assert.match(err.message, /Failed to move/);
        return true;
      },
    );
  });
});

// ---------- deleteEmail ----------

describe('deleteEmail', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('deletes email by moving to trash', async () => {
    stubMailboxes(client);
    stubMakeRequest(client, {
      methodResponses: [['Email/set', { updated: { e1: null } }, 'moveToTrash']],
    });

    await client.deleteEmail('e1');
    // No error means success
  });

  it('throws when trash mailbox is not found', async () => {
    stubMailboxes(client, [INBOX_MAILBOX, DRAFTS_MAILBOX]);

    await assert.rejects(
      () => client.deleteEmail('e1'),
      (err: Error) => {
        assert.match(err.message, /Trash/);
        return true;
      },
    );
  });
});

// ---------- markEmailRead ----------

describe('markEmailRead', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('marks email as read', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [['Email/set', { updated: { e1: null } }, 'updateEmail']],
    }));

    await client.markEmailRead('e1', true);

    const update = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].update;
    assert.equal(update.e1['keywords/$seen'], true);
  });

  it('marks email as unread', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [['Email/set', { updated: { e1: null } }, 'updateEmail']],
    }));

    await client.markEmailRead('e1', false);

    const update = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].update;
    assert.equal(update.e1['keywords/$seen'], null);
  });

  it('throws when update fails', async () => {
    stubMakeRequest(client, {
      methodResponses: [['Email/set', { notUpdated: { e1: { type: 'notFound' } } }, 'updateEmail']],
    });

    await assert.rejects(
      () => client.markEmailRead('e1'),
      (err: Error) => {
        assert.match(err.message, /Failed to mark/);
        return true;
      },
    );
  });
});

// ---------- bulkMarkRead ----------

describe('bulkMarkRead', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('marks multiple emails as read in one request', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [['Email/set', { updated: { e1: null, e2: null, e3: null } }, 'bulkUpdate']],
    }));

    await client.bulkMarkRead(['e1', 'e2', 'e3'], true);

    const update = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].update;
    assert.equal(update.e1['keywords/$seen'], true);
    assert.equal(update.e2['keywords/$seen'], true);
    assert.equal(update.e3['keywords/$seen'], true);
  });

  it('marks multiple emails as unread', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [['Email/set', { updated: { e1: null, e2: null } }, 'bulkUpdate']],
    }));

    await client.bulkMarkRead(['e1', 'e2'], false);

    const update = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].update;
    assert.equal(update.e1['keywords/$seen'], null);
    assert.equal(update.e2['keywords/$seen'], null);
  });

  it('throws when some emails fail to update', async () => {
    stubMakeRequest(client, {
      methodResponses: [['Email/set', { notUpdated: { e2: { type: 'notFound' } } }, 'bulkUpdate']],
    });

    await assert.rejects(
      () => client.bulkMarkRead(['e1', 'e2']),
      (err: Error) => {
        assert.match(err.message, /Failed to update/);
        return true;
      },
    );
  });
});

// ---------- getMethodResult ----------

describe('getMethodResult', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('throws on JMAP error response', async () => {
    stubMakeRequest(client, {
      methodResponses: [['error', { type: 'serverFail', description: 'internal error' }, 'op']],
    });

    await assert.rejects(
      () => client.getEmailById('e1'),
      (err: Error) => {
        assert.match(err.message, /serverFail/);
        assert.match(err.message, /internal error/);
        return true;
      },
    );
  });

  it('throws when index exceeds response length', async () => {
    stubMakeRequest(client, {
      methodResponses: [],
    });

    await assert.rejects(
      () => client.getEmailById('e1'),
      (err: Error) => {
        assert.match(err.message, /missing expected method/i);
        return true;
      },
    );
  });

  it('throws on malformed entry (not an array)', async () => {
    stubMakeRequest(client, {
      methodResponses: ['not-a-tuple' as unknown],
    });

    await assert.rejects(
      () => client.getEmailById('e1'),
      (err: Error) => {
        assert.match(err.message, /malformed/i);
        return true;
      },
    );
  });

  it('throws on error without description', async () => {
    stubMakeRequest(client, {
      methodResponses: [['error', { type: 'unknownMethod' }, 'op']],
    });

    await assert.rejects(
      () => client.getEmailById('e1'),
      (err: Error) => {
        assert.match(err.message, /unknownMethod/);
        return true;
      },
    );
  });
});

// ---------- getListResult ----------

describe('getListResult', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('extracts list from valid response', async () => {
    stubMakeRequest(client, {
      methodResponses: [['Mailbox/get', { list: [INBOX_MAILBOX, DRAFTS_MAILBOX] }, 'mailboxes']],
    });

    const mailboxes = await client.getMailboxes();
    assert.equal(mailboxes.length, 2);
  });

  it('returns empty array when list property is missing', async () => {
    stubMakeRequest(client, {
      methodResponses: [['Mailbox/get', { notList: 'something' }, 'mailboxes']],
    });

    const mailboxes = await client.getMailboxes();
    assert.deepEqual(mailboxes, []);
  });

  it('returns empty array when result is null-ish', async () => {
    stubMakeRequest(client, {
      methodResponses: [['Mailbox/get', null, 'mailboxes']],
    });

    const mailboxes = await client.getMailboxes();
    assert.deepEqual(mailboxes, []);
  });
});

// ---------- sendEmail ----------

describe('sendEmail', () => {
  let client: JmapClient;
  const IDENTITY = { id: 'id-1', email: 'me@example.com', mayDelete: false, name: 'Me' };

  beforeEach(() => {
    client = makeClient();
    mock.method(client, 'getIdentities', async () => [IDENTITY]);
    mock.method(client, 'getMailboxes', async () => [DRAFTS_MAILBOX, SENT_MAILBOX]);
  });

  it('returns submission ID on success', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['Email/set', { created: { draft: { id: 'email-1' } } }, 'createEmail'],
        ['EmailSubmission/set', { created: { submission: { id: 'sub-1' } } }, 'submitEmail'],
      ],
    });

    const subId = await client.sendEmail({
      to: ['bob@example.com'],
      subject: 'Test',
      textBody: 'Hello',
    });
    assert.equal(subId, 'sub-1');
  });

  it('throws when Drafts mailbox not found', async () => {
    mock.method(client, 'getMailboxes', async () => [SENT_MAILBOX]);
    await assert.rejects(
      () => client.sendEmail({ to: ['bob@example.com'], subject: 'X', textBody: 'Y' }),
      (err: Error) => {
        assert.match(err.message, /Drafts/);
        return true;
      },
    );
  });

  it('throws when Sent mailbox not found', async () => {
    mock.method(client, 'getMailboxes', async () => [DRAFTS_MAILBOX]);
    await assert.rejects(
      () => client.sendEmail({ to: ['bob@example.com'], subject: 'X', textBody: 'Y' }),
      (err: Error) => {
        assert.match(err.message, /Sent/);
        return true;
      },
    );
  });

  it('throws when no body provided', async () => {
    await assert.rejects(
      () => client.sendEmail({ to: ['bob@example.com'], subject: 'X' }),
      (err: Error) => {
        assert.match(err.message, /textBody or htmlBody/);
        return true;
      },
    );
  });

  it('throws on email creation failure (notCreated)', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        [
          'Email/set',
          { notCreated: { draft: { type: 'invalidProperties', description: 'bad' } } },
          'createEmail',
        ],
        ['EmailSubmission/set', { created: { submission: { id: 'sub-1' } } }, 'submitEmail'],
      ],
    });

    await assert.rejects(
      () => client.sendEmail({ to: ['bob@example.com'], subject: 'X', textBody: 'Y' }),
      (err: Error) => {
        assert.match(err.message, /invalidProperties/);
        return true;
      },
    );
  });

  it('throws on submission failure (notCreated)', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['Email/set', { created: { draft: { id: 'email-1' } } }, 'createEmail'],
        [
          'EmailSubmission/set',
          { notCreated: { submission: { type: 'forbidden', description: 'blocked' } } },
          'submitEmail',
        ],
      ],
    });

    await assert.rejects(
      () => client.sendEmail({ to: ['bob@example.com'], subject: 'X', textBody: 'Y' }),
      (err: Error) => {
        assert.match(err.message, /forbidden/);
        return true;
      },
    );
  });

  it('throws when no submission ID returned', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['Email/set', { created: { draft: { id: 'email-1' } } }, 'createEmail'],
        ['EmailSubmission/set', { created: { submission: {} } }, 'submitEmail'],
      ],
    });

    await assert.rejects(
      () => client.sendEmail({ to: ['bob@example.com'], subject: 'X', textBody: 'Y' }),
      (err: Error) => {
        assert.match(err.message, /no submission ID/);
        return true;
      },
    );
  });

  it('includes cc and bcc in envelope rcptTo', async () => {
    let capturedRequest: JmapRequest | undefined;
    mock.method(client, 'makeRequest', async (req: JmapRequest) => {
      capturedRequest = req;
      return {
        methodResponses: [
          ['Email/set', { created: { draft: { id: 'email-1' } } }, 'createEmail'],
          ['EmailSubmission/set', { created: { submission: { id: 'sub-1' } } }, 'submitEmail'],
        ],
      };
    });

    await client.sendEmail({
      to: ['to@example.com'],
      cc: ['cc@example.com'],
      bcc: ['bcc@example.com'],
      subject: 'Test',
      textBody: 'Body',
    });

    const submission = capturedRequest?.methodCalls[1]?.[1] as Record<string, unknown>;
    const create = submission?.create as Record<string, Record<string, unknown>>;
    const envelope = create?.submission?.envelope as Record<string, unknown>;
    const rcptTo = envelope?.rcptTo as Array<{ email: string }>;
    const recipients = rcptTo.map((r) => r.email);

    assert.ok(recipients.includes('to@example.com'), 'to should be in rcptTo');
    assert.ok(recipients.includes('cc@example.com'), 'cc should be in rcptTo');
    assert.ok(recipients.includes('bcc@example.com'), 'bcc should be in rcptTo');
  });
});

// ---------- pinEmail ----------

describe('pinEmail', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('pins an email', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [['Email/set', { updated: { e1: null } }, 'pinEmail']],
    }));

    await client.pinEmail('e1', true);

    const update = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].update;
    assert.deepEqual(update.e1, { 'keywords/$flagged': true });
  });

  it('unpins an email', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [['Email/set', { updated: { e1: null } }, 'pinEmail']],
    }));

    await client.pinEmail('e1', false);

    const update = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].update;
    assert.deepEqual(update.e1, { 'keywords/$flagged': null });
  });

  it('throws when update fails', async () => {
    stubMakeRequest(client, {
      methodResponses: [['Email/set', { notUpdated: { e1: { type: 'notFound' } } }, 'pinEmail']],
    });

    await assert.rejects(
      () => client.pinEmail('e1'),
      (err: Error) => {
        assert.match(err.message, /pin/i);
        return true;
      },
    );
  });
});

// ---------- addLabels / removeLabels ----------

describe('addLabels', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('adds labels to an email', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [['Email/set', { updated: { e1: null } }, 'addLabels']],
    }));

    await client.addLabels('e1', ['mb-1', 'mb-2']);

    const update = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].update;
    assert.deepEqual(update.e1, { 'mailboxIds/mb-1': true, 'mailboxIds/mb-2': true });
  });

  it('throws when update fails', async () => {
    stubMakeRequest(client, {
      methodResponses: [['Email/set', { notUpdated: { e1: { type: 'notFound' } } }, 'addLabels']],
    });

    await assert.rejects(
      () => client.addLabels('e1', ['mb-1']),
      (err: Error) => {
        assert.match(err.message, /Failed to add labels/);
        return true;
      },
    );
  });
});

describe('removeLabels', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('removes labels from an email', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [['Email/set', { updated: { e1: null } }, 'removeLabels']],
    }));

    await client.removeLabels('e1', ['mb-1', 'mb-2']);

    const update = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].update;
    assert.deepEqual(update.e1, { 'mailboxIds/mb-1': null, 'mailboxIds/mb-2': null });
  });

  it('throws when update fails', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['Email/set', { notUpdated: { e1: { type: 'notFound' } } }, 'removeLabels'],
      ],
    });

    await assert.rejects(
      () => client.removeLabels('e1', ['mb-1']),
      (err: Error) => {
        assert.match(err.message, /Failed to remove labels/);
        return true;
      },
    );
  });
});

// ---------- bulkAddLabels / bulkRemoveLabels ----------

describe('bulkAddLabels', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('adds labels to multiple emails', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [['Email/set', { updated: { e1: null, e2: null } }, 'bulkAddLabels']],
    }));

    await client.bulkAddLabels(['e1', 'e2'], ['mb-1']);

    const updates = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].update;
    assert.deepEqual(updates.e1, { 'mailboxIds/mb-1': true });
    assert.deepEqual(updates.e2, { 'mailboxIds/mb-1': true });
  });

  it('throws when some updates fail', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['Email/set', { notUpdated: { e2: { type: 'notFound' } } }, 'bulkAddLabels'],
      ],
    });

    await assert.rejects(
      () => client.bulkAddLabels(['e1', 'e2'], ['mb-1']),
      (err: Error) => {
        assert.match(err.message, /Failed to add labels/);
        return true;
      },
    );
  });
});

describe('bulkRemoveLabels', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('removes labels from multiple emails', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [['Email/set', { updated: { e1: null, e2: null } }, 'bulkRemoveLabels']],
    }));

    await client.bulkRemoveLabels(['e1', 'e2'], ['mb-1']);

    const updates = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].update;
    assert.deepEqual(updates.e1, { 'mailboxIds/mb-1': null });
    assert.deepEqual(updates.e2, { 'mailboxIds/mb-1': null });
  });

  it('throws when some updates fail', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['Email/set', { notUpdated: { e2: { type: 'notFound' } } }, 'bulkRemoveLabels'],
      ],
    });

    await assert.rejects(
      () => client.bulkRemoveLabels(['e1', 'e2'], ['mb-1']),
      (err: Error) => {
        assert.match(err.message, /Failed to remove labels/);
        return true;
      },
    );
  });
});

// ---------- getEmailAttachments ----------

describe('getEmailAttachments', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('returns attachments for an email', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        [
          'Email/get',
          {
            list: [
              {
                id: 'e1',
                attachments: [
                  { blobId: 'b1', name: 'file.pdf', type: 'application/pdf', size: 1024 },
                ],
              },
            ],
          },
          'getAttachments',
        ],
      ],
    });

    const attachments = await client.getEmailAttachments('e1');
    assert.equal(attachments.length, 1);
    assert.equal(attachments[0].name, 'file.pdf');
  });

  it('returns empty array when no attachments', async () => {
    stubMakeRequest(client, {
      methodResponses: [['Email/get', { list: [{ id: 'e1' }] }, 'getAttachments']],
    });

    const attachments = await client.getEmailAttachments('e1');
    assert.deepEqual(attachments, []);
  });
});

// ---------- downloadAttachment ----------

describe('downloadAttachment', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
    // Override getSession to include downloadUrl
    mock.method(client, 'getSession', async () => ({
      apiUrl: 'https://api.example.com/jmap/api/',
      accountId: ACCOUNT_ID,
      capabilities: {},
      downloadUrl: 'https://api.example.com/download/{accountId}/{blobId}/{type}/{name}',
    }));
  });

  it('returns download URL for attachment by partId', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        [
          'Email/get',
          {
            list: [
              {
                id: 'e1',
                attachments: [
                  { partId: 'p1', blobId: 'b1', name: 'photo.jpg', type: 'image/jpeg' },
                ],
              },
            ],
          },
          'getEmail',
        ],
      ],
    });

    const url = await client.downloadAttachment('e1', 'p1');
    assert.ok(url.includes('b1'));
    assert.ok(url.includes(ACCOUNT_ID));
  });

  it('returns download URL for attachment by numeric index', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        [
          'Email/get',
          {
            list: [
              {
                id: 'e1',
                attachments: [{ partId: 'p1', blobId: 'b1', name: 'file.txt', type: 'text/plain' }],
              },
            ],
          },
          'getEmail',
        ],
      ],
    });

    const url = await client.downloadAttachment('e1', '0');
    assert.ok(url.includes('b1'));
  });

  it('throws when email not found', async () => {
    stubMakeRequest(client, {
      methodResponses: [['Email/get', { list: [] }, 'getEmail']],
    });

    await assert.rejects(
      () => client.downloadAttachment('missing', 'p1'),
      (err: Error) => {
        assert.match(err.message, /not found/i);
        return true;
      },
    );
  });

  it('throws when attachment not found', async () => {
    stubMakeRequest(client, {
      methodResponses: [['Email/get', { list: [{ id: 'e1', attachments: [] }] }, 'getEmail']],
    });

    await assert.rejects(
      () => client.downloadAttachment('e1', 'nonexistent'),
      (err: Error) => {
        assert.match(err.message, /Attachment not found/);
        return true;
      },
    );
  });

  it('throws when downloadUrl not available', async () => {
    mock.method(client, 'getSession', async () => ({
      apiUrl: 'https://api.example.com/jmap/api/',
      accountId: ACCOUNT_ID,
      capabilities: {},
    }));

    stubMakeRequest(client, {
      methodResponses: [
        [
          'Email/get',
          {
            list: [
              {
                id: 'e1',
                attachments: [{ partId: 'p1', blobId: 'b1', name: 'file.txt', type: 'text/plain' }],
              },
            ],
          },
          'getEmail',
        ],
      ],
    });

    await assert.rejects(
      () => client.downloadAttachment('e1', 'p1'),
      (err: Error) => {
        assert.match(err.message, /Download capability/);
        return true;
      },
    );
  });
});

// ---------- advancedSearch ----------

describe('advancedSearch', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('returns results with various filters', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/query', { ids: ['e1'] }, 'query'],
        ['Email/get', { list: [{ id: 'e1', subject: 'Found' }] }, 'emails'],
      ],
    }));

    const results = await client.advancedSearch({
      query: 'test',
      from: 'alice@example.com',
      to: 'bob@example.com',
      subject: 'meeting',
      hasAttachment: true,
      mailboxId: 'mb-inbox',
      after: '2026-01-01',
      before: '2026-12-31',
    });
    assert.equal(results.length, 1);

    const filter = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].filter;
    assert.equal(filter.text, 'test');
    assert.equal(filter.from, 'alice@example.com');
    assert.equal(filter.to, 'bob@example.com');
    assert.equal(filter.subject, 'meeting');
    assert.equal(filter.hasAttachment, true);
    assert.equal(filter.inMailbox, 'mb-inbox');
    assert.equal(filter.after, '2026-01-01T00:00:00Z');
    assert.equal(filter.before, '2026-12-31T00:00:00Z');
  });

  it('applies isUnread filter', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/query', { ids: [] }, 'query'],
        ['Email/get', { list: [] }, 'emails'],
      ],
    }));

    await client.advancedSearch({ isUnread: true });

    const filter = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].filter;
    assert.equal(filter.notKeyword, '$seen');
  });

  it('applies isPinned filter', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/query', { ids: [] }, 'query'],
        ['Email/get', { list: [] }, 'emails'],
      ],
    }));

    await client.advancedSearch({ isPinned: true });

    const filter = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].filter;
    assert.equal(filter.hasKeyword, '$flagged');
  });

  it('combines isUnread + isPinned with AND operator', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/query', { ids: [] }, 'query'],
        ['Email/get', { list: [] }, 'emails'],
      ],
    }));

    await client.advancedSearch({ isUnread: true, isPinned: true });

    const filter = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].filter;
    assert.equal(filter.operator, 'AND');
    assert.ok(Array.isArray(filter.conditions));
    assert.equal(filter.conditions.length, 3);
  });

  it('applies isUnread=false and isPinned=false', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/query', { ids: [] }, 'query'],
        ['Email/get', { list: [] }, 'emails'],
      ],
    }));

    await client.advancedSearch({ isUnread: false, isPinned: false });

    const filter = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].filter;
    assert.equal(filter.operator, 'AND');
  });
});

// ---------- getThread ----------

describe('getThread', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('returns thread emails', async () => {
    let callCount = 0;
    mock.method(client, 'makeRequest', async () => {
      callCount++;
      if (callCount === 1) {
        // First call: resolve email to thread
        return {
          methodResponses: [['Email/get', { list: [{ threadId: 'thread-1' }] }, 'checkEmail']],
        };
      }
      // Second call: get thread
      return {
        methodResponses: [
          ['Thread/get', { list: [{ id: 'thread-1', emailIds: ['e1', 'e2'] }] }, 'getThread'],
          [
            'Email/get',
            {
              list: [
                { id: 'e1', subject: 'First', threadId: 'thread-1' },
                { id: 'e2', subject: 'Re: First', threadId: 'thread-1' },
              ],
            },
            'emails',
          ],
        ],
      };
    });

    const emails = await client.getThread('e1');
    assert.equal(emails.length, 2);
    assert.equal(emails[0].subject, 'First');
  });

  it('throws when thread not found', async () => {
    let callCount = 0;
    mock.method(client, 'makeRequest', async () => {
      callCount++;
      if (callCount === 1) {
        return { methodResponses: [['Email/get', { list: [] }, 'checkEmail']] };
      }
      return {
        methodResponses: [
          ['Thread/get', { list: [], notFound: ['missing-thread'] }, 'getThread'],
          ['Email/get', { list: [] }, 'emails'],
        ],
      };
    });

    await assert.rejects(
      () => client.getThread('missing-thread'),
      (err: Error) => {
        assert.match(err.message, /not found/i);
        return true;
      },
    );
  });

  it('handles email lookup failure gracefully', async () => {
    let callCount = 0;
    mock.method(client, 'makeRequest', async () => {
      callCount++;
      if (callCount === 1) {
        // Simulate email lookup error
        throw new Error('lookup failed');
      }
      return {
        methodResponses: [
          ['Thread/get', { list: [{ id: 'thread-1', emailIds: ['e1'] }] }, 'getThread'],
          ['Email/get', { list: [{ id: 'e1', subject: 'Solo' }] }, 'emails'],
        ],
      };
    });

    // Should fall through to using the original threadId
    const emails = await client.getThread('thread-1');
    assert.equal(emails.length, 1);
  });
});

// ---------- getMailboxStats ----------

describe('getMailboxStats', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('returns stats for specific mailbox', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        [
          'Mailbox/get',
          {
            list: [
              {
                id: 'mb-inbox',
                name: 'Inbox',
                role: 'inbox',
                totalEmails: 42,
                unreadEmails: 5,
                totalThreads: 30,
                unreadThreads: 3,
              },
            ],
          },
          'mailbox',
        ],
      ],
    });

    const stats = await client.getMailboxStats('mb-inbox');
    assert.equal((stats as Record<string, unknown>).totalEmails, 42);
    assert.equal((stats as Record<string, unknown>).unreadEmails, 5);
  });

  it('returns stats for all mailboxes when no id specified', async () => {
    stubMailboxes(client);

    const stats = await client.getMailboxStats();
    assert.ok(Array.isArray(stats));
    assert.equal(stats.length, 4);
    assert.equal((stats[0] as Record<string, unknown>).name, 'Inbox');
  });
});

// ---------- getAccountSummary ----------

describe('getAccountSummary', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
    mock.method(client, 'getIdentities', async () => [
      { id: 'id-1', email: 'me@example.com', mayDelete: false },
    ]);
    stubMailboxes(client);
  });

  it('returns account summary with totals', async () => {
    const summary = await client.getAccountSummary();
    assert.equal(summary.accountId, ACCOUNT_ID);
    assert.equal(summary.mailboxCount, 4);
    assert.equal(summary.identityCount, 1);
    assert.ok(Array.isArray(summary.mailboxes));
  });
});

// ---------- bulkPinEmails ----------

describe('bulkPinEmails', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('pins multiple emails', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [['Email/set', { updated: { e1: null, e2: null } }, 'bulkFlag']],
    }));

    await client.bulkPinEmails(['e1', 'e2'], true);

    const updates = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].update;
    assert.deepEqual(updates.e1, { 'keywords/$flagged': true });
    assert.deepEqual(updates.e2, { 'keywords/$flagged': true });
  });

  it('unpins multiple emails', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [['Email/set', { updated: { e1: null, e2: null } }, 'bulkFlag']],
    }));

    await client.bulkPinEmails(['e1', 'e2'], false);

    const updates = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].update;
    assert.deepEqual(updates.e1, { 'keywords/$flagged': null });
    assert.deepEqual(updates.e2, { 'keywords/$flagged': null });
  });

  it('throws when some updates fail', async () => {
    stubMakeRequest(client, {
      methodResponses: [['Email/set', { notUpdated: { e2: { type: 'notFound' } } }, 'bulkFlag']],
    });

    await assert.rejects(
      () => client.bulkPinEmails(['e1', 'e2']),
      (err: Error) => {
        assert.match(err.message, /pin\/unpin/);
        return true;
      },
    );
  });
});

// ---------- bulkMove ----------

describe('bulkMove', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('moves multiple emails to target mailbox', async () => {
    let callCount = 0;
    mock.method(client, 'makeRequest', async () => {
      callCount++;
      if (callCount === 1) {
        return {
          methodResponses: [
            [
              'Email/get',
              {
                list: [
                  { id: 'e1', mailboxIds: { 'mb-inbox': true } },
                  { id: 'e2', mailboxIds: { 'mb-inbox': true } },
                ],
              },
              'getEmails',
            ],
          ],
        };
      }
      return {
        methodResponses: [['Email/set', { updated: { e1: null, e2: null } }, 'bulkMove']],
      };
    });

    await client.bulkMove(['e1', 'e2'], 'mb-archive');
    assert.equal(callCount, 2);
  });

  it('throws when some moves fail', async () => {
    let callCount = 0;
    mock.method(client, 'makeRequest', async () => {
      callCount++;
      if (callCount === 1) {
        return {
          methodResponses: [['Email/get', { list: [{ id: 'e1', mailboxIds: {} }] }, 'getEmails']],
        };
      }
      return {
        methodResponses: [['Email/set', { notUpdated: { e1: { type: 'notFound' } } }, 'bulkMove']],
      };
    });

    await assert.rejects(
      () => client.bulkMove(['e1'], 'mb-archive'),
      (err: Error) => {
        assert.match(err.message, /Failed to move/);
        return true;
      },
    );
  });
});

// ---------- bulkDelete ----------

describe('bulkDelete', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
    stubMailboxes(client);
  });

  it('deletes multiple emails by moving to trash', async () => {
    stubMakeRequest(client, {
      methodResponses: [['Email/set', { updated: { e1: null, e2: null } }, 'bulkDelete']],
    });

    await client.bulkDelete(['e1', 'e2']);
  });

  it('throws when Trash not found', async () => {
    stubMailboxes(client, [INBOX_MAILBOX]);

    await assert.rejects(
      () => client.bulkDelete(['e1']),
      (err: Error) => {
        assert.match(err.message, /Trash/);
        return true;
      },
    );
  });

  it('throws when some deletes fail', async () => {
    stubMakeRequest(client, {
      methodResponses: [['Email/set', { notUpdated: { e1: { type: 'notFound' } } }, 'bulkDelete']],
    });

    await assert.rejects(
      () => client.bulkDelete(['e1']),
      (err: Error) => {
        assert.match(err.message, /Failed to delete/);
        return true;
      },
    );
  });
});

// ---------- getDefaultIdentity ----------

describe('getDefaultIdentity', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('returns non-deletable identity as default', async () => {
    mock.method(client, 'getIdentities', async () => [
      { id: 'id-1', email: 'alias@example.com', mayDelete: true },
      { id: 'id-2', email: 'me@example.com', mayDelete: false },
    ]);

    const identity = await client.getDefaultIdentity();
    assert.equal(identity.email, 'me@example.com');
  });

  it('returns first identity when none has mayDelete=false', async () => {
    mock.method(client, 'getIdentities', async () => [
      { id: 'id-1', email: 'first@example.com', mayDelete: true },
      { id: 'id-2', email: 'second@example.com', mayDelete: true },
    ]);

    const identity = await client.getDefaultIdentity();
    assert.equal(identity.email, 'first@example.com');
  });
});

// ---------- sendDraft: additional edge cases ----------

describe('sendDraft edge cases', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
    mock.method(client, 'getIdentities', async () => [
      { id: 'id-1', email: 'me@example.com', mayDelete: false },
    ]);
    mock.method(client, 'getMailboxes', async () => [
      { id: 'mb-sent', name: 'Sent', role: 'sent' },
    ]);
  });

  it('throws when draft has no from address', async () => {
    mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        [
          'Email/get',
          { list: [{ id: 'd1', keywords: { $draft: true }, to: [{ email: 'bob@example.com' }] }] },
          'getEmail',
        ],
      ],
    }));

    await assert.rejects(
      () => client.sendDraft('d1'),
      (err: Error) => {
        assert.match(err.message, /no from address/i);
        return true;
      },
    );
  });

  it('throws when from address does not match any identity', async () => {
    mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        [
          'Email/get',
          {
            list: [
              {
                id: 'd1',
                keywords: { $draft: true },
                from: [{ email: 'unknown@example.com' }],
                to: [{ email: 'bob@example.com' }],
              },
            ],
          },
          'getEmail',
        ],
      ],
    }));

    await assert.rejects(
      () => client.sendDraft('d1'),
      (err: Error) => {
        assert.match(err.message, /identity/i);
        return true;
      },
    );
  });

  it('throws when Sent mailbox not found', async () => {
    mock.method(client, 'getMailboxes', async () => [DRAFTS_MAILBOX]);

    mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        [
          'Email/get',
          {
            list: [
              {
                id: 'd1',
                keywords: { $draft: true },
                from: [{ email: 'me@example.com' }],
                to: [{ email: 'bob@example.com' }],
              },
            ],
          },
          'getEmail',
        ],
      ],
    }));

    await assert.rejects(
      () => client.sendDraft('d1'),
      (err: Error) => {
        assert.match(err.message, /Sent/);
        return true;
      },
    );
  });

  it('throws when submission returns no ID', async () => {
    let callCount = 0;
    mock.method(client, 'makeRequest', async () => {
      callCount++;
      if (callCount === 1) {
        return {
          methodResponses: [
            [
              'Email/get',
              {
                list: [
                  {
                    id: 'd1',
                    keywords: { $draft: true },
                    from: [{ email: 'me@example.com' }],
                    to: [{ email: 'bob@example.com' }],
                  },
                ],
              },
              'getEmail',
            ],
          ],
        };
      }
      return {
        methodResponses: [['EmailSubmission/set', { created: { submission: {} } }, 'submitDraft']],
      };
    });

    await assert.rejects(
      () => client.sendDraft('d1'),
      (err: Error) => {
        assert.match(err.message, /no submission ID/i);
        return true;
      },
    );
  });

  it('includes bcc recipients in envelope', async () => {
    let callCount = 0;
    const makeReq = mock.method(client, 'makeRequest', async () => {
      callCount++;
      if (callCount === 1) {
        return {
          methodResponses: [
            [
              'Email/get',
              {
                list: [
                  {
                    id: 'd1',
                    keywords: { $draft: true },
                    from: [{ email: 'me@example.com' }],
                    to: [{ email: 'bob@example.com' }],
                    cc: [],
                    bcc: [{ email: 'secret@example.com' }],
                  },
                ],
              },
              'getEmail',
            ],
          ],
        };
      }
      return {
        methodResponses: [
          ['EmailSubmission/set', { created: { submission: { id: 'sub-1' } } }, 'submitDraft'],
        ],
      };
    });

    const subId = await client.sendDraft('d1');
    assert.equal(subId, 'sub-1');

    // Check that bcc was included in rcptTo
    const submitCall = makeReq.mock.calls[1].arguments[0];
    const rcptTo = submitCall.methodCalls[0][1].create.submission.envelope.rcptTo;
    assert.equal(rcptTo.length, 2);
  });
});

// ---------- validateEmailAddress ----------

describe('validateEmailAddress', () => {
  it('accepts valid email', () => {
    assert.doesNotThrow(() => JmapClient.validateEmailAddress('user@example.com'));
  });

  it('rejects missing @', () => {
    assert.throws(() => JmapClient.validateEmailAddress('noatsign'), /Invalid email/);
  });

  it('rejects control characters', () => {
    assert.throws(() => JmapClient.validateEmailAddress('user\x00@evil.com'), /Invalid email/);
  });

  it('rejects newlines', () => {
    assert.throws(() => JmapClient.validateEmailAddress('user\n@evil.com'), /Invalid email/);
  });

  it('rejects empty string', () => {
    assert.throws(() => JmapClient.validateEmailAddress(''), /Invalid email/);
  });
});
