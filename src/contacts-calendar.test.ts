import { describe, it, beforeEach, mock } from 'node:test';
import assert from 'node:assert/strict';
import { ContactsCalendarClient } from './contacts-calendar.js';
import { FastmailAuth } from './auth.js';

// ---------- helpers ----------

const MAIL_ACCOUNT = 'acct-mail';
const CONTACTS_ACCOUNT = 'acct-contacts';

function makeClient(opts: { contactsPrimary?: boolean } = { contactsPrimary: true }): ContactsCalendarClient {
  const auth = new FastmailAuth({ apiToken: 'fake-token' });
  const client = new ContactsCalendarClient(auth);

  mock.method(client, 'getSession', async () => ({
    apiUrl: 'https://api.example.com/jmap/api/',
    accountId: MAIL_ACCOUNT,
    capabilities: { 'urn:ietf:params:jmap:contacts': {} },
    primaryAccounts: opts.contactsPrimary
      ? { 'urn:ietf:params:jmap:contacts': CONTACTS_ACCOUNT, 'urn:ietf:params:jmap:mail': MAIL_ACCOUNT }
      : { 'urn:ietf:params:jmap:mail': MAIL_ACCOUNT },
  }));

  return client;
}

function stubMakeRequest(client: ContactsCalendarClient, response: any) {
  return mock.method(client, 'makeRequest', async () => response);
}

// ---------- createContact ----------

describe('createContact', () => {
  let client: ContactsCalendarClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('builds an RFC 9610 Card and returns the created id', async () => {
    const makeReq = stubMakeRequest(client, {
      methodResponses: [
        ['ContactCard/set', { created: { newContact: { id: 'C1', uid: 'u-1' } } }, 'createContact'],
      ],
    });

    const id = await client.createContact({
      name: { given: 'Ada', surname: 'Lovelace', full: 'Ada Lovelace' },
      emails: [{ address: 'ada@example.com', label: 'work' }],
      phones: [{ number: '+1 555 0100' }],
      notes: 'test note',
    });

    assert.equal(id, 'C1');
    const [method, params] = makeReq.mock.calls[0].arguments[0].methodCalls[0];
    assert.equal(method, 'ContactCard/set');
    assert.equal(params.accountId, CONTACTS_ACCOUNT);
    const card = params.create.newContact;
    assert.equal(card['@type'], 'Card');
    assert.equal(card.name.full, 'Ada Lovelace');
    assert.deepEqual(card.name.components, [
      { kind: 'given', value: 'Ada' },
      { kind: 'surname', value: 'Lovelace' },
    ]);
    assert.deepEqual(card.emails, { e0: { address: 'ada@example.com', label: 'work' } });
    assert.deepEqual(card.phones, { p0: { number: '+1 555 0100' } });
    assert.deepEqual(card.notes, { n0: { note: 'test note' } });
  });

  it('falls back to the mail account when no contacts primary exists', async () => {
    client = makeClient({ contactsPrimary: false });
    const makeReq = stubMakeRequest(client, {
      methodResponses: [
        ['ContactCard/set', { created: { newContact: { id: 'C2' } } }, 'createContact'],
      ],
    });
    await client.createContact({ name: { full: 'Solo Mail' } });
    assert.equal(makeReq.mock.calls[0].arguments[0].methodCalls[0][1].accountId, MAIL_ACCOUNT);
  });

  it('passes addressBookIds when addressBookId is supplied', async () => {
    const makeReq = stubMakeRequest(client, {
      methodResponses: [
        ['ContactCard/set', { created: { newContact: { id: 'C3' } } }, 'createContact'],
      ],
    });
    await client.createContact({ name: { full: 'Booked' }, addressBookId: 'ab-1' });
    const card = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].create.newContact;
    assert.deepEqual(card.addressBookIds, { 'ab-1': true });
  });

  it('rejects empty input client-side', async () => {
    await assert.rejects(
      () => client.createContact({}),
      /name or at least one email/,
    );
  });

  it('surfaces notCreated errors with type and description', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['ContactCard/set', { notCreated: { newContact: { type: 'invalidProperties', description: 'bad email' } } }, 'createContact'],
      ],
    });
    await assert.rejects(
      () => client.createContact({ name: { full: 'X' } }),
      /invalidProperties.*bad email/s,
    );
  });
});

// ---------- updateContact ----------

describe('updateContact', () => {
  let client: ContactsCalendarClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('checks existence then sends a top-level patch', async () => {
    const makeReq = mock.method(client, 'makeRequest', async (req: any) => {
      if (req.methodCalls[0][0] === 'ContactCard/get') {
        return { methodResponses: [['ContactCard/get', { list: [{ id: 'C1' }] }, 'g']] };
      }
      return { methodResponses: [['ContactCard/set', { updated: { C1: null } }, 'u']] };
    });

    await client.updateContact('C1', { emails: [{ address: 'new@example.com' }] });

    const setCall = makeReq.mock.calls.find((c: any) => c.arguments[0].methodCalls[0][0] === 'ContactCard/set');
    const update = setCall.arguments[0].methodCalls[0][1].update.C1;
    assert.deepEqual(update, { emails: { e0: { address: 'new@example.com' } } });
  });

  it('passes expectState through as ifInState', async () => {
    const makeReq = mock.method(client, 'makeRequest', async (req: any) => {
      if (req.methodCalls[0][0] === 'ContactCard/get') {
        return { methodResponses: [['ContactCard/get', { list: [{ id: 'C1' }] }, 'g']] };
      }
      return { methodResponses: [['ContactCard/set', { updated: { C1: null } }, 'u']] };
    });
    await client.updateContact('C1', { notes: 'x', expectState: 'state-42' });
    const setCall = makeReq.mock.calls.find((c: any) => c.arguments[0].methodCalls[0][0] === 'ContactCard/set');
    assert.equal(setCall.arguments[0].methodCalls[0][1].ifInState, 'state-42');
  });

  it('throws not-found before attempting the update', async () => {
    const makeReq = stubMakeRequest(client, {
      methodResponses: [
        ['ContactCard/get', { list: [], notFound: ['ghost'] }, 'g'],
      ],
    });
    await assert.rejects(() => client.updateContact('ghost', { notes: 'x' }), /Contact not found: ghost/);
    assert.equal(makeReq.mock.calls.length, 1);
  });

  it('surfaces notUpdated errors', async () => {
    mock.method(client, 'makeRequest', async (req: any) => {
      if (req.methodCalls[0][0] === 'ContactCard/get') {
        return { methodResponses: [['ContactCard/get', { list: [{ id: 'C1' }] }, 'g']] };
      }
      return { methodResponses: [['ContactCard/set', { notUpdated: { C1: { type: 'stateMismatch' } } }, 'u']] };
    });
    await assert.rejects(() => client.updateContact('C1', { notes: 'x' }), /stateMismatch/);
  });

  it('rejects an empty patch client-side', async () => {
    await assert.rejects(() => client.updateContact('C1', {}), /at least one field/i);
  });
});

// ---------- deleteContact ----------

describe('deleteContact', () => {
  let client: ContactsCalendarClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('destroys by id', async () => {
    const makeReq = stubMakeRequest(client, {
      methodResponses: [
        ['ContactCard/set', { destroyed: ['C1'] }, 'd'],
      ],
    });
    await client.deleteContact('C1');
    assert.deepEqual(makeReq.mock.calls[0].arguments[0].methodCalls[0][1].destroy, ['C1']);
  });

  it('maps notFound destroy errors to the not-found convention', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['ContactCard/set', { notDestroyed: { ghost: { type: 'notFound' } } }, 'd'],
      ],
    });
    await assert.rejects(() => client.deleteContact('ghost'), /Contact not found: ghost/);
  });

  it('surfaces other notDestroyed errors with their type', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['ContactCard/set', { notDestroyed: { C1: { type: 'forbidden', description: 'read-only scope' } } }, 'd'],
      ],
    });
    await assert.rejects(() => client.deleteContact('C1'), /forbidden.*read-only scope/s);
  });

  it('passes expectState through as ifInState', async () => {
    const makeReq = stubMakeRequest(client, {
      methodResponses: [
        ['ContactCard/set', { destroyed: ['C1'] }, 'd'],
      ],
    });
    await client.deleteContact('C1', 'state-7');
    assert.equal(makeReq.mock.calls[0].arguments[0].methodCalls[0][1].ifInState, 'state-7');
  });
});
