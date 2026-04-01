import { describe, it, beforeEach, mock } from 'node:test';
import assert from 'node:assert/strict';
import { JmapClient } from './jmap-client.js';
import { FastmailAuth } from './auth.js';

// ---------- helpers ----------

const ACCOUNT_ID = 'acct-123';
const INBOX_MAILBOX = { id: 'mb-inbox', name: 'Inbox', role: 'inbox', totalEmails: 42, unreadEmails: 5 };
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

function stubMakeRequest(client: JmapClient, response: any) {
  mock.method(client, 'makeRequest', async () => response);
}

function stubMailboxes(client: JmapClient, mailboxes: any[] = [INBOX_MAILBOX, DRAFTS_MAILBOX, TRASH_MAILBOX, SENT_MAILBOX]) {
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
      methodResponses: [
        ['Mailbox/get', { list: [INBOX_MAILBOX, DRAFTS_MAILBOX] }, 'mailboxes'],
      ],
    });

    const mailboxes = await client.getMailboxes();
    assert.equal(mailboxes.length, 2);
    assert.equal(mailboxes[0].role, 'inbox');
    assert.equal(mailboxes[1].id, 'mb-drafts');
  });

  it('returns empty array when response list is missing', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['Mailbox/get', {}, 'mailboxes'],
      ],
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
        ['Email/get', { list: [
          { id: 'e1', subject: 'First' },
          { id: 'e2', subject: 'Second' },
        ] }, 'emails'],
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
      methodResponses: [
        ['Email/get', { list: [{ id: 'e1', subject: 'Found' }] }, 'email'],
      ],
    });

    const email = await client.getEmailById('e1');
    assert.equal(email.id, 'e1');
    assert.equal(email.subject, 'Found');
  });

  it('throws when email is not found (empty list)', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['Email/get', { list: [] }, 'email'],
      ],
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
      methodResponses: [
        ['Email/get', { list: [], notFound: ['gone'] }, 'email'],
      ],
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
        methodResponses: [
          ['Email/set', { updated: { 'e1': null } }, 'moveEmail'],
        ],
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
        methodResponses: [
          ['Email/set', { notUpdated: { 'e1': { type: 'notFound' } } }, 'moveEmail'],
        ],
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
      methodResponses: [
        ['Email/set', { updated: { 'e1': null } }, 'moveToTrash'],
      ],
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
      methodResponses: [
        ['Email/set', { updated: { 'e1': null } }, 'updateEmail'],
      ],
    }));

    await client.markEmailRead('e1', true);

    const update = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].update;
    assert.deepEqual(update['e1'], { 'keywords/$seen': true });
  });

  it('marks email as unread', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/set', { updated: { 'e1': null } }, 'updateEmail'],
      ],
    }));

    await client.markEmailRead('e1', false);

    const update = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].update;
    assert.deepEqual(update['e1'], { 'keywords/$seen': null });
  });

  it('throws when update fails', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['Email/set', { notUpdated: { 'e1': { type: 'notFound' } } }, 'updateEmail'],
      ],
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
      methodResponses: [
        ['Email/set', { updated: { 'e1': null, 'e2': null, 'e3': null } }, 'bulkUpdate'],
      ],
    }));

    await client.bulkMarkRead(['e1', 'e2', 'e3'], true);

    const update = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].update;
    assert.deepEqual(update['e1'], { 'keywords/$seen': true });
    assert.deepEqual(update['e2'], { 'keywords/$seen': true });
    assert.deepEqual(update['e3'], { 'keywords/$seen': true });
  });

  it('marks multiple emails as unread', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/set', { updated: { 'e1': null, 'e2': null } }, 'bulkUpdate'],
      ],
    }));

    await client.bulkMarkRead(['e1', 'e2'], false);

    const update = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].update;
    assert.deepEqual(update['e1'], { 'keywords/$seen': null });
    assert.deepEqual(update['e2'], { 'keywords/$seen': null });
  });

  it('throws when some emails fail to update', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['Email/set', { notUpdated: { 'e2': { type: 'notFound' } } }, 'bulkUpdate'],
      ],
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
      methodResponses: [
        ['error', { type: 'serverFail', description: 'internal error' }, 'op'],
      ],
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
      methodResponses: ['not-a-tuple' as any],
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
      methodResponses: [
        ['error', { type: 'unknownMethod' }, 'op'],
      ],
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
      methodResponses: [
        ['Mailbox/get', { list: [INBOX_MAILBOX, DRAFTS_MAILBOX] }, 'mailboxes'],
      ],
    });

    const mailboxes = await client.getMailboxes();
    assert.equal(mailboxes.length, 2);
  });

  it('returns empty array when list property is missing', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['Mailbox/get', { notList: 'something' }, 'mailboxes'],
      ],
    });

    const mailboxes = await client.getMailboxes();
    assert.deepEqual(mailboxes, []);
  });

  it('returns empty array when result is null-ish', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['Mailbox/get', null, 'mailboxes'],
      ],
    });

    const mailboxes = await client.getMailboxes();
    assert.deepEqual(mailboxes, []);
  });
});
