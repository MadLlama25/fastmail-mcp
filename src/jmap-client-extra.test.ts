import { describe, it, beforeEach, mock } from 'node:test';
import assert from 'node:assert/strict';
import { JmapClient, buildEmailQueryFilter } from './jmap-client.js';
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

// ---------- ascending sort parameter ----------

describe('ascending sort parameter', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  const QUERY_GET_RESPONSE = {
    methodResponses: [
      ['Email/query', { ids: ['e1'] }, 'query'],
      ['Email/get', { list: [{ id: 'e1', subject: 'Test' }] }, 'emails'],
    ],
  };

  describe('getEmails', () => {
    it('defaults to isAscending: false', async () => {
      const makeReq = mock.method(client, 'makeRequest', async () => QUERY_GET_RESPONSE);

      await client.getEmails('mb-inbox', 5);

      const sort = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].sort;
      assert.deepEqual(sort, [{ property: 'receivedAt', isAscending: false }]);
    });

    it('passes ascending=true as isAscending: true', async () => {
      const makeReq = mock.method(client, 'makeRequest', async () => QUERY_GET_RESPONSE);

      await client.getEmails('mb-inbox', 5, true);

      const sort = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].sort;
      assert.deepEqual(sort, [{ property: 'receivedAt', isAscending: true }]);
    });
  });

  describe('searchEmails', () => {
    it('defaults to isAscending: false', async () => {
      const makeReq = mock.method(client, 'makeRequest', async () => QUERY_GET_RESPONSE);

      await client.searchEmails('test', 10);

      const sort = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].sort;
      assert.deepEqual(sort, [{ property: 'receivedAt', isAscending: false }]);
    });

    it('passes ascending=true as isAscending: true', async () => {
      const makeReq = mock.method(client, 'makeRequest', async () => QUERY_GET_RESPONSE);

      await client.searchEmails('test', 10, true);

      const sort = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].sort;
      assert.deepEqual(sort, [{ property: 'receivedAt', isAscending: true }]);
    });
  });

  describe('getRecentEmails', () => {
    it('defaults to isAscending: false', async () => {
      stubMailboxes(client);
      const makeReq = mock.method(client, 'makeRequest', async () => QUERY_GET_RESPONSE);

      await client.getRecentEmails(10, 'inbox');

      const sort = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].sort;
      assert.deepEqual(sort, [{ property: 'receivedAt', isAscending: false }]);
    });

    it('passes ascending=true as isAscending: true', async () => {
      stubMailboxes(client);
      const makeReq = mock.method(client, 'makeRequest', async () => QUERY_GET_RESPONSE);

      await client.getRecentEmails(10, 'inbox', true);

      const sort = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].sort;
      assert.deepEqual(sort, [{ property: 'receivedAt', isAscending: true }]);
    });
  });

  describe('advancedSearch', () => {
    it('defaults to isAscending: false when ascending not specified', async () => {
      const makeReq = mock.method(client, 'makeRequest', async () => QUERY_GET_RESPONSE);

      await client.advancedSearch({ query: 'test' });

      const sort = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].sort;
      assert.deepEqual(sort, [{ property: 'receivedAt', isAscending: false }]);
    });

    it('passes ascending: true as isAscending: true', async () => {
      const makeReq = mock.method(client, 'makeRequest', async () => QUERY_GET_RESPONSE);

      await client.advancedSearch({ query: 'test', ascending: true });

      const sort = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].sort;
      assert.deepEqual(sort, [{ property: 'receivedAt', isAscending: true }]);
    });
  });
});

// ---------- *_metadata variants: privacy invariant ----------
//
// These tests pin the load-bearing privacy invariant of the four metadata
// variants: their JMAP `Email/get` properties allowlist must never contain
// `preview` (or any body-derived field). A future refactor that accidentally
// re-introduces preview will fail here loudly, rather than silently leaking
// body excerpts to callers operating under least-privilege constraints.

describe('metadata variants — JMAP properties allowlist', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  const QUERY_GET_RESPONSE = {
    methodResponses: [
      ['Email/query', { ids: ['e1'] }, 'query'],
      ['Email/get', { list: [{ id: 'e1', subject: 'Test' }] }, 'emails'],
    ],
  };

  const FORBIDDEN_PROPERTIES = ['preview', 'textBody', 'htmlBody', 'bodyValues', 'body', 'bodyStructure'];

  function assertNoBodyProperties(props: any) {
    assert.ok(Array.isArray(props), 'properties must be an array');
    for (const forbidden of FORBIDDEN_PROPERTIES) {
      assert.ok(
        !props.includes(forbidden),
        `properties allowlist must not contain '${forbidden}' (got: ${JSON.stringify(props)})`,
      );
    }
  }

  it('getEmailsMetadata excludes preview and body fields from Email/get properties', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => QUERY_GET_RESPONSE);

    await client.getEmailsMetadata('mb-inbox', 5);

    const props = makeReq.mock.calls[0].arguments[0].methodCalls[1][1].properties;
    assertNoBodyProperties(props);
  });

  it('searchEmailsMetadata excludes preview and body fields from Email/get properties', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => QUERY_GET_RESPONSE);

    await client.searchEmailsMetadata('test', 10);

    const props = makeReq.mock.calls[0].arguments[0].methodCalls[1][1].properties;
    assertNoBodyProperties(props);
  });

  it('advancedSearchMetadata excludes preview and body fields from Email/get properties', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => QUERY_GET_RESPONSE);

    await client.advancedSearchMetadata({ query: 'test', mailboxId: 'mb-inbox' });

    const props = makeReq.mock.calls[0].arguments[0].methodCalls[1][1].properties;
    assertNoBodyProperties(props);
  });

  it('advancedSearchMetadata preserves the full filter logic of advancedSearch', async () => {
    // The metadata variant is structurally identical to advancedSearch except
    // for the property list — confirm filter handling is intact (mailbox,
    // recipient, attachment, isUnread/isPinned conjunction).
    const makeReq = mock.method(client, 'makeRequest', async () => QUERY_GET_RESPONSE);

    await client.advancedSearchMetadata({
      mailboxId: 'mb-inbox',
      to: 'someone@example.com',
      hasAttachment: true,
      isUnread: true,
      isPinned: true,
    });

    const filter = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].filter;
    // When both isUnread and isPinned are set, advancedSearch wraps the
    // conditions in an AND operator — the metadata variant must do the same.
    assert.equal(filter.operator, 'AND');
    assert.ok(Array.isArray(filter.conditions));
  });

  it('getThreadMetadata excludes preview and body fields when fetching thread emails', async () => {
    // Two-step: the threadId-resolution probe also calls Email/get, but with
    // a `properties: ['threadId']` minimum payload — that's fine. The
    // privacy-critical call is the second `makeRequest` (the Thread/get +
    // Email/get composite). We only inspect that one's properties list.
    let callIndex = 0;
    const makeReq = mock.method(client, 'makeRequest', async () => {
      callIndex += 1;
      if (callIndex === 1) {
        // First call: threadId resolution probe.
        return { methodResponses: [['Email/get', { list: [{ id: 'e1', threadId: 't1' }] }, 'checkEmail']] };
      }
      // Second call: Thread/get + Email/get composite — this is the one that
      // would leak preview if the allowlist were wrong.
      return {
        methodResponses: [
          ['Thread/get', { list: [{ id: 't1', emailIds: ['e1'] }] }, 'getThread'],
          ['Email/get', { list: [{ id: 'e1', subject: 'Test' }] }, 'emails'],
        ],
      };
    });

    await client.getThreadMetadata('e1');

    // Inspect the second call's Email/get properties (the composite).
    const compositeProps = makeReq.mock.calls[1].arguments[0].methodCalls[1][1].properties;
    assertNoBodyProperties(compositeProps);
  });
});

// ---------- buildEmailQueryFilter ----------
//
// Direct unit tests for the JMAP filter assembler. These pin the wire-level
// shapes that advancedSearch / advancedSearchMetadata depend on. The helper is
// pure, so testing it directly is much cheaper than round-tripping through a
// mocked makeRequest for every shape.

describe('buildEmailQueryFilter', () => {
  it('returns an empty filter when no fields are set', () => {
    assert.deepEqual(buildEmailQueryFilter({}), {});
  });

  it('passes through standard FilterCondition fields on a flat condition', () => {
    const filter = buildEmailQueryFilter({
      query: 'invoice',
      from: 'a@example.com',
      to: 'b@example.com',
      subject: 'Q1 2025',
      hasAttachment: true,
      after: '2025-01-01T00:00:00Z',
      before: '2025-04-01T00:00:00Z',
    });
    assert.deepEqual(filter, {
      text: 'invoice',
      from: 'a@example.com',
      to: 'b@example.com',
      subject: 'Q1 2025',
      hasAttachment: true,
      after: '2025-01-01T00:00:00Z',
      before: '2025-04-01T00:00:00Z',
    });
  });

  it('flattens a single mailboxId into a flat FilterCondition', () => {
    assert.deepEqual(
      buildEmailQueryFilter({ mailboxId: 'mb-1' }),
      { inMailbox: 'mb-1' },
    );
  });

  it('flattens a single requiredMailboxIds entry to the same shape as mailboxId', () => {
    const viaArray = buildEmailQueryFilter({ requiredMailboxIds: ['mb-1'] });
    const viaScalar = buildEmailQueryFilter({ mailboxId: 'mb-1' });
    assert.deepEqual(viaArray, viaScalar);
  });

  it('builds FilterOperator AND when requiredMailboxIds has multiple entries', () => {
    const filter = buildEmailQueryFilter({ requiredMailboxIds: ['mb-1', 'mb-2'] });
    assert.equal(filter.operator, 'AND');
    assert.equal(filter.conditions.length, 2);
    assert.deepEqual(filter.conditions[0], { inMailbox: 'mb-1' });
    assert.deepEqual(filter.conditions[1], { inMailbox: 'mb-2' });
  });

  it('combines mailboxId with requiredMailboxIds and de-duplicates overlap', () => {
    const filter = buildEmailQueryFilter({
      mailboxId: 'mb-1',
      requiredMailboxIds: ['mb-1', 'mb-2'], // 'mb-1' duplicates the scalar; should drop
    });
    assert.equal(filter.operator, 'AND');
    assert.equal(filter.conditions.length, 2);
    assert.deepEqual(filter.conditions[0], { inMailbox: 'mb-1' });
    assert.deepEqual(filter.conditions[1], { inMailbox: 'mb-2' });
  });

  it('emits inMailboxOtherThan on a single-condition filter', () => {
    const filter = buildEmailQueryFilter({
      mailboxId: 'mb-parent',
      excludeMailboxIds: ['mb-archive'],
    });
    assert.deepEqual(filter, {
      inMailbox: 'mb-parent',
      inMailboxOtherThan: ['mb-archive'],
    });
  });

  it('emits inMailboxOtherThan inside the AND-base when multi-required is needed', () => {
    const filter = buildEmailQueryFilter({
      requiredMailboxIds: ['mb-inbox', 'mb-label'],
      excludeMailboxIds: ['mb-archive'],
      isPinned: true,
    });
    assert.equal(filter.operator, 'AND');
    assert.deepEqual(filter.conditions[0], {
      inMailboxOtherThan: ['mb-archive'],
      hasKeyword: '$flagged',
    });
    assert.deepEqual(filter.conditions[1], { inMailbox: 'mb-inbox' });
    assert.deepEqual(filter.conditions[2], { inMailbox: 'mb-label' });
  });

  it('keeps a single keyword condition flat when only one of isUnread/isPinned is set', () => {
    assert.deepEqual(
      buildEmailQueryFilter({ mailboxId: 'mb-1', isPinned: true }),
      { inMailbox: 'mb-1', hasKeyword: '$flagged' },
    );
  });

  it('preserves the legacy isUnread+isPinned AND-split when both are set', () => {
    // Both keyword conditions must be split out of base to avoid
    // hasKeyword/notKeyword collisions on a single FilterCondition.
    const filter = buildEmailQueryFilter({
      mailboxId: 'mb-inbox',
      isUnread: true,
      isPinned: true,
    });
    assert.equal(filter.operator, 'AND');
    assert.equal(filter.conditions.length, 3);
    assert.deepEqual(filter.conditions[0], { inMailbox: 'mb-inbox' });
    assert.ok(filter.conditions.some((c: any) => c.notKeyword === '$seen'));
    assert.ok(filter.conditions.some((c: any) => c.hasKeyword === '$flagged'));
  });

  it('combines every primitive in one shot', () => {
    const filter = buildEmailQueryFilter({
      query: 'invoice',
      hasAttachment: true,
      requiredMailboxIds: ['mb-inbox', 'mb-receipts'],
      excludeMailboxIds: ['mb-archive'],
      isPinned: true,
      after: '2025-01-01T00:00:00Z',
    });
    assert.equal(filter.operator, 'AND');
    assert.deepEqual(filter.conditions[0], {
      text: 'invoice',
      hasAttachment: true,
      after: '2025-01-01T00:00:00Z',
      inMailboxOtherThan: ['mb-archive'],
      hasKeyword: '$flagged',
    });
    assert.deepEqual(filter.conditions[1], { inMailbox: 'mb-inbox' });
    assert.deepEqual(filter.conditions[2], { inMailbox: 'mb-receipts' });
  });

  it('treats empty requiredMailboxIds and empty excludeMailboxIds as absent', () => {
    assert.deepEqual(
      buildEmailQueryFilter({
        mailboxId: 'mb-1',
        requiredMailboxIds: [],
        excludeMailboxIds: [],
      }),
      { inMailbox: 'mb-1' },
    );
  });
});

// ---------- advancedSearch / advancedSearchMetadata wiring ----------

describe('advancedSearch and advancedSearchMetadata wire requiredMailboxIds correctly', () => {
  let client: JmapClient;
  const QUERY_GET_RESPONSE = {
    methodResponses: [
      ['Email/query', { ids: [] }, 'query'],
      ['Email/get', { list: [] }, 'emails'],
    ],
  };

  beforeEach(() => {
    client = makeClient();
  });

  it('advancedSearch sends a FilterOperator AND when requiredMailboxIds has 2 entries', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => QUERY_GET_RESPONSE);

    await client.advancedSearch({
      requiredMailboxIds: ['mb-inbox', 'mb-label'],
      isPinned: true,
    });

    const filter = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].filter;
    assert.equal(filter.operator, 'AND');
    assert.equal(filter.conditions.length, 3);
    assert.ok(filter.conditions.some((c: any) => c.inMailbox === 'mb-inbox'));
    assert.ok(filter.conditions.some((c: any) => c.inMailbox === 'mb-label'));
    assert.ok(filter.conditions.some((c: any) => c.hasKeyword === '$flagged'));
  });

  it('advancedSearchMetadata sends inMailboxOtherThan when excludeMailboxIds is set', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => QUERY_GET_RESPONSE);

    await client.advancedSearchMetadata({
      mailboxId: 'mb-parent',
      excludeMailboxIds: ['mb-child-archive'],
    });

    const filter = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].filter;
    assert.deepEqual(filter, {
      inMailbox: 'mb-parent',
      inMailboxOtherThan: ['mb-child-archive'],
    });
  });

  it('advancedSearch backwards-compat: legacy single-mailbox calls still produce flat filter', async () => {
    // Pre-1.11 callers passing only mailboxId must continue to produce a flat
    // FilterCondition shape.
    const makeReq = mock.method(client, 'makeRequest', async () => QUERY_GET_RESPONSE);

    await client.advancedSearch({ mailboxId: 'mb-only' });

    const filter = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].filter;
    assert.deepEqual(filter, { inMailbox: 'mb-only' });
  });
});
