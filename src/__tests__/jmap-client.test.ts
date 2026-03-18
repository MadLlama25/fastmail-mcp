import { describe, it, expect, vi, beforeEach } from 'vitest';
import { JmapClient } from '../jmap-client.js';
import { FastmailAuth } from '../auth.js';

function createMockClient() {
  const auth = new FastmailAuth({ apiToken: 'test-token' });
  const client = new JmapClient(auth);

  vi.spyOn(client, 'getSession' as any).mockResolvedValue({
    apiUrl: 'https://api.fastmail.com/jmap/api/',
    accountId: 'test-account-id',
    capabilities: {},
    downloadUrl: 'https://api.fastmail.com/jmap/download/{accountId}/{blobId}/{name}',
    uploadUrl: 'https://api.fastmail.com/jmap/upload/{accountId}/'
  });

  return client;
}

describe('getMailboxes', () => {
  it('returns mailbox list from valid JMAP response', async () => {
    const client = createMockClient();
    const mailboxes = [
      { id: 'mb-1', name: 'Inbox', role: 'inbox', totalEmails: 42 },
      { id: 'mb-2', name: 'Sent', role: 'sent', totalEmails: 10 }
    ];

    vi.spyOn(client, 'makeRequest').mockResolvedValue({
      methodResponses: [['Mailbox/get', { list: mailboxes, accountId: 'test-account-id', state: '1' }, 'mailboxes']],
      sessionState: 'state-1'
    });

    const result = await client.getMailboxes();
    expect(result).toEqual(mailboxes);
    expect(result).toHaveLength(2);
    expect(result[0].role).toBe('inbox');
  });

  it('throws descriptive error when response has no .list property', async () => {
    const client = createMockClient();

    vi.spyOn(client, 'makeRequest').mockResolvedValue({
      methodResponses: [['Mailbox/get', { accountId: 'test-account-id', state: '1' }, 'mailboxes']],
      sessionState: 'state-1'
    });

    await expect(client.getMailboxes()).rejects.toThrow('getMailboxes: unexpected JMAP response');
  });

  it('throws descriptive error when methodResponses is empty', async () => {
    const client = createMockClient();

    vi.spyOn(client, 'makeRequest').mockResolvedValue({
      methodResponses: [],
      sessionState: 'state-1'
    });

    await expect(client.getMailboxes()).rejects.toThrow('getMailboxes: unexpected JMAP response');
  });

  it('throws descriptive error when methodResponses is undefined', async () => {
    const client = createMockClient();

    vi.spyOn(client, 'makeRequest').mockResolvedValue({
      methodResponses: undefined as any,
      sessionState: 'state-1'
    });

    await expect(client.getMailboxes()).rejects.toThrow('getMailboxes: unexpected JMAP response');
  });
});

describe('getRecentEmails', () => {
  it('returns emails from valid JMAP response', async () => {
    const client = createMockClient();
    const mailboxes = [
      { id: 'mb-1', name: 'Inbox', role: 'inbox' }
    ];
    const emails = [
      { id: 'email-1', subject: 'Hello', from: [{ email: 'test@example.com' }], receivedAt: '2026-03-18T10:00:00Z' },
      { id: 'email-2', subject: 'World', from: [{ email: 'foo@example.com' }], receivedAt: '2026-03-18T09:00:00Z' }
    ];

    // getRecentEmails calls getMailboxes internally, so we mock makeRequest to handle both calls
    const makeRequestSpy = vi.spyOn(client, 'makeRequest');

    // First call: getMailboxes (from inside getRecentEmails)
    makeRequestSpy.mockResolvedValueOnce({
      methodResponses: [['Mailbox/get', { list: mailboxes, accountId: 'test-account-id', state: '1' }, 'mailboxes']],
      sessionState: 'state-1'
    });

    // Second call: Email/query + Email/get
    makeRequestSpy.mockResolvedValueOnce({
      methodResponses: [
        ['Email/query', { ids: ['email-1', 'email-2'] }, 'query'],
        ['Email/get', { list: emails }, 'emails']
      ],
      sessionState: 'state-1'
    });

    const result = await client.getRecentEmails(10, 'inbox');
    expect(result).toEqual(emails);
    expect(result).toHaveLength(2);
  });

  it('throws error when target mailbox is not found', async () => {
    const client = createMockClient();
    const mailboxes = [
      { id: 'mb-1', name: 'Inbox', role: 'inbox' }
    ];

    vi.spyOn(client, 'makeRequest').mockResolvedValueOnce({
      methodResponses: [['Mailbox/get', { list: mailboxes, accountId: 'test-account-id', state: '1' }, 'mailboxes']],
      sessionState: 'state-1'
    });

    await expect(client.getRecentEmails(10, 'nonexistent')).rejects.toThrow('Could not find mailbox: nonexistent');
  });

  it('correctly maps mailbox by role name', async () => {
    const client = createMockClient();
    const mailboxes = [
      { id: 'mb-1', name: 'INBOX', role: 'inbox' },
      { id: 'mb-2', name: 'Sent Mail', role: 'sent' },
      { id: 'mb-3', name: 'Drafts', role: 'drafts' }
    ];
    const emails = [
      { id: 'email-1', subject: 'Sent message' }
    ];

    const makeRequestSpy = vi.spyOn(client, 'makeRequest');

    // getMailboxes call
    makeRequestSpy.mockResolvedValueOnce({
      methodResponses: [['Mailbox/get', { list: mailboxes, accountId: 'test-account-id', state: '1' }, 'mailboxes']],
      sessionState: 'state-1'
    });

    // Email/query + Email/get call
    makeRequestSpy.mockResolvedValueOnce({
      methodResponses: [
        ['Email/query', { ids: ['email-1'] }, 'query'],
        ['Email/get', { list: emails }, 'emails']
      ],
      sessionState: 'state-1'
    });

    const result = await client.getRecentEmails(10, 'sent');
    expect(result).toEqual(emails);

    // Verify the second makeRequest was called with the correct mailbox ID (mb-2 for sent)
    const secondCall = makeRequestSpy.mock.calls[1][0];
    expect(secondCall.methodCalls[0][1].filter.inMailbox).toBe('mb-2');
  });
});
