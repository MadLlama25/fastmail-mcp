import { after, before, beforeEach, describe, it, mock } from 'node:test';
import assert from 'node:assert/strict';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { server } from './index.js';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

const MASKED_CAPABILITY = 'https://www.fastmail.com/dev/maskedemail';
const MASK = 'masked.sender@example.com';
const IDENTITY = { id: 'identity', email: 'owner@example.com', name: 'Account Owner', mayDelete: false };
const original = {
  id: 'original', subject: 'Request', from: [{ email: 'recipient@example.com' }],
  messageId: ['original@example.com'], references: ['ancestor@example.com'],
};

describe('masked authoring through MCP tools', () => {
  const client = new Client({ name: 'authoring-test', version: '1.0' });
  let emails: Map<string, any>;
  let calls: any[];
  let masks: any;
  let maskError: boolean;
  let identities: any[];
  let identityError: 'denied' | 'missingId' | undefined;
  let nextId: number;
  let hasSentMailbox: boolean;
  let attachmentDir: string;
  let fetchMock: ReturnType<typeof mock.method>;
  const envKeys = ['FASTMAIL_API_TOKEN', 'FASTMAIL_BASE_URL', 'FASTMAIL_ALLOW_UNSAFE_BASE_URL', 'FASTMAIL_DOWNLOAD_DIR'];
  const savedEnv = envKeys.map(key => process.env[key]);

  before(async () => {
    process.env.FASTMAIL_API_TOKEN = 'synthetic-test-value';
    process.env.FASTMAIL_BASE_URL = 'https://jmap.example.com';
    process.env.FASTMAIL_ALLOW_UNSAFE_BASE_URL = 'true';
    attachmentDir = await mkdtemp(join(tmpdir(), 'fastmail-authoring-test-'));
    process.env.FASTMAIL_DOWNLOAD_DIR = attachmentDir;
    // No test request can leave this synthetic transport, even if configuration changes.
    fetchMock = mock.method(globalThis, 'fetch', async (input: any, init?: RequestInit) => {
      const url = String(input);
      assert.ok(url === 'https://jmap.example.com/jmap/session' || url === 'https://jmap.example.com/jmap/api/');
      if (url.endsWith('/session')) {
        return Response.json({
          apiUrl: 'https://jmap.example.com/jmap/api/',
          accounts: { account: { accountCapabilities: { [MASKED_CAPABILITY]: {} } } },
          primaryAccounts: { 'urn:ietf:params:jmap:mail': 'account' },
          capabilities: { [MASKED_CAPABILITY]: {} },
        });
      }
      const request = JSON.parse(init!.body as string);
      const methodResponses = request.methodCalls.map(([method, args, id]: any) => {
        calls.push({ method, args, using: request.using });
        switch (method) {
          case 'Identity/get': return [method, { list: identities }, id];
          case 'Identity/set': {
            if (identityError === 'denied') return [method, { notCreated: { maskedSender: { type: 'forbiddenFrom', description: 'Mask cannot send' } } }, id];
            if (identityError === 'missingId') return [method, { created: { maskedSender: {} } }, id];
            const identity = { ...args.create.maskedSender, id: 'masked-identity' };
            identities.push(identity);
            return [method, { created: { maskedSender: identity } }, id];
          }
          case 'Mailbox/get': return [method, { list: [
            { id: 'drafts', role: 'drafts', name: 'Drafts' },
            ...(hasSentMailbox ? [{ id: 'sent', role: 'sent', name: 'Sent' }] : []),
          ] }, id];
          case 'MaskedEmail/get': return maskError
            ? ['error', { type: 'forbidden' }, id]
            : [method, { accountId: 'account', list: masks }, id];
          case 'Email/get': return [method, {
            list: args.ids.filter((emailId: string) => emails.has(emailId)).map((emailId: string) => {
              const stored = emails.get(emailId);
              return Object.fromEntries(args.properties.filter((key: string) => key in stored).map((key: string) => [key, stored[key]]));
            }),
          }, id];
          case 'Email/set': {
            const emailId = `draft-${++nextId}`;
            emails.set(emailId, { ...args.create.draft, id: emailId });
            for (const oldId of args.destroy ?? []) emails.delete(oldId);
            return [method, { created: { draft: { id: emailId } } }, id];
          }
          case 'EmailSubmission/set': return [method, { created: { submission: { id: 'submission' } } }, id];
          default: throw new Error(`Unexpected method: ${method}`);
        }
      });
      return Response.json({ methodResponses, sessionState: 'fixture' });
    });
    const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
    await server.connect(serverTransport);
    await client.connect(clientTransport);
  });

  beforeEach(() => {
    emails = new Map([['original', original]]);
    calls = [];
    masks = [{ id: 'mask', email: MASK, state: 'enabled' }];
    maskError = false;
    identities = [IDENTITY];
    identityError = undefined;
    nextId = 0;
    hasSentMailbox = true;
  });

  after(async () => {
    await client.close();
    await server.close();
    fetchMock.mock.restore();
    await rm(attachmentDir, { recursive: true });
    envKeys.forEach((key, i) => {
      if (savedEnv[i] === undefined) delete process.env[key];
      else process.env[key] = savedEnv[i];
    });
  });

  it('creates and edits a draft from an owned mask without using the primary identity name', async () => {
    await client.callTool({ name: 'create_draft', arguments: {
      from: MASK, to: ['recipient@example.com'], subject: 'Request', textBody: 'Draft body',
    } });
    await client.callTool({ name: 'edit_draft', arguments: { emailId: 'draft-1', subject: 'Updated request' } });
    assert.equal(emails.has('draft-1'), false);
    const draft = emails.get('draft-2');
    assert.deepEqual(draft.from, [{ email: MASK }]);
    assert.equal(draft.subject, 'Updated request');
    assert.equal(draft.bodyValues.text.value, 'Draft body');
    assert.deepEqual(draft.to, [{ email: 'recipient@example.com' }]);
    const ownershipRead = calls.find(call => call.method === 'MaskedEmail/get');
    assert.equal(ownershipRead.args.accountId, 'account');
    assert.ok(ownershipRead.using.includes(MASKED_CAPABILITY));
    assert.equal(calls.some(call => call.method === 'Identity/set'), false);
  });

  it('keeps masked reply headers and both bodies when removing Re: from its subject', async () => {
    await client.callTool({ name: 'reply_email', arguments: {
      originalEmailId: 'original', from: MASK, send: false, textBody: 'Reply', htmlBody: '<p>Reply</p>',
    } });
    await client.callTool({ name: 'edit_draft', arguments: { emailId: 'draft-1', subject: 'Request' } });
    const draft = emails.get('draft-2');
    assert.deepEqual(draft.from, [{ email: MASK }]);
    assert.equal(draft.subject, 'Request');
    assert.deepEqual(draft.inReplyTo, ['original@example.com']);
    assert.deepEqual(draft.references, ['ancestor@example.com', 'original@example.com']);
    assert.equal(draft.bodyValues.text.value, 'Reply');
    assert.equal(draft.bodyValues.html.value, '<p>Reply</p>');
  });

  it('preserves ordinary reply threading during a subject-only edit', async () => {
    await client.callTool({ name: 'reply_email', arguments: {
      originalEmailId: 'original', from: IDENTITY.email, send: false, textBody: 'Reply',
    } });
    await client.callTool({ name: 'edit_draft', arguments: { emailId: 'draft-1', subject: 'Request' } });
    assert.deepEqual(emails.get('draft-2').inReplyTo, ['original@example.com']);
    assert.deepEqual(emails.get('draft-2').references, ['ancestor@example.com', 'original@example.com']);
    assert.equal(calls.some(call => call.method === 'MaskedEmail/get'), false);
  });

  it('allows an owned disabled mask and an explicit switch to it when editing', async () => {
    masks[0].state = 'disabled';
    await client.callTool({ name: 'create_draft', arguments: { subject: 'Draft', textBody: 'Body' } });
    await client.callTool({ name: 'edit_draft', arguments: { emailId: 'draft-1', from: MASK } });
    assert.deepEqual(emails.get('draft-2').from, [{ email: MASK }]);
  });

  for (const failure of ['unowned', 'deleted', 'pending', 'unavailable', 'malformed']) {
    it(`rejects ${failure} mask ownership before creating or replacing mail`, async () => {
      if (failure === 'unowned') masks = [];
      if (failure === 'deleted' || failure === 'pending') masks[0].state = failure;
      if (failure === 'unavailable') maskError = true;
      if (failure === 'malformed') masks = { email: MASK, state: 'enabled' };
      await assert.rejects(client.callTool({ name: 'create_draft', arguments: { from: MASK, subject: 'Draft' } }));
      emails.set('existing', {
        id: 'existing', keywords: { $draft: true }, from: [{ email: MASK }],
        to: [{ email: 'recipient@example.com' }], subject: 'Keep',
      });
      await assert.rejects(client.callTool({ name: 'edit_draft', arguments: { emailId: 'existing', subject: 'Changed' } }));
      await assert.rejects(client.callTool({ name: 'send_draft', arguments: { emailId: 'existing' } }));
      assert.equal(calls.some(call => call.method === 'Email/set'), false);
      assert.equal(calls.some(call => call.method === 'Identity/set' || call.method === 'EmailSubmission/set'), false);
      assert.equal(emails.get('existing').subject, 'Keep');
    });
  }

  for (const tool of ['send_email', 'reply_email', 'send_draft']) {
    it(`uses a real mask sending identity for ${tool} and reuses it on the next send`, async () => {
      const args = tool === 'send_email'
        ? { from: MASK, to: ['recipient@example.com'], subject: 'Request', textBody: 'Message' }
        : tool === 'reply_email'
          ? { from: MASK, originalEmailId: 'original', textBody: 'Message', send: true }
          : { emailId: 'existing' };
      emails.set('existing', { id: 'existing', from: [{ email: MASK }], to: [{ email: 'recipient@example.com' }], keywords: { $draft: true } });
      await client.callTool({ name: tool, arguments: args });
      await client.callTool({ name: tool, arguments: args });
      const creation = calls.filter(call => call.method === 'Identity/set');
      assert.equal(creation.length, 1);
      assert.equal(creation[0].args.accountId, 'account');
      assert.deepEqual(creation[0].args.create.maskedSender, { email: MASK, name: '' });
      for (const call of calls.filter(call => call.method === 'EmailSubmission/set')) {
        assert.equal(call.args.create.submission.identityId, 'masked-identity');
        assert.equal(call.args.create.submission.envelope.mailFrom.email, MASK);
        assert.deepEqual(call.args.create.submission.envelope.rcptTo, [{ email: 'recipient@example.com' }]);
      }
      if (tool !== 'send_draft') assert.equal(emails.get('draft-1').from[0].email, MASK);
    });

    it(`does not mutate mail if sending identity creation is denied for ${tool}`, async () => {
      identityError = 'denied';
      const args = tool === 'send_email'
        ? { from: MASK, to: ['recipient@example.com'], subject: 'Request', textBody: 'Message' }
        : tool === 'reply_email'
          ? { from: MASK, originalEmailId: 'original', textBody: 'Message', send: true }
          : { emailId: 'existing' };
      emails.set('existing', { id: 'existing', from: [{ email: MASK }], to: [{ email: 'recipient@example.com' }], keywords: { $draft: true } });
      await assert.rejects(client.callTool({ name: tool, arguments: args }), /forbiddenFrom/);
      assert.equal(calls.some(call => call.method === 'Email/set' || call.method === 'EmailSubmission/set'), false);
      assert.ok(emails.has('existing'));
    });
  }

  it('does not write mail when sending identity creation returns no real ID', async () => {
    identityError = 'missingId';
    await assert.rejects(client.callTool({ name: 'send_email', arguments: {
      from: MASK, to: ['recipient@example.com'], subject: 'Request', textBody: 'Body',
    } }), /no identity ID/);
    assert.equal(calls.some(call => call.method === 'Email/set' || call.method === 'EmailSubmission/set'), false);
  });

  for (const failure of ['missingBody', 'missingAttachment', 'missingSentMailbox']) {
    it(`does not create a sending identity for an invalid send: ${failure}`, async () => {
      hasSentMailbox = failure !== 'missingSentMailbox';
      await assert.rejects(client.callTool({ name: 'send_email', arguments: {
        from: MASK, to: ['recipient@example.com'], subject: 'Request',
        ...(failure !== 'missingBody' && { textBody: 'Body' }),
        ...(failure === 'missingAttachment' && { attachments: [{ localPath: 'missing-synthetic-attachment.txt' }] }),
      } }), failure === 'missingAttachment' ? /Attachment file not found/ : undefined);
      assert.equal(calls.some(call => ['Identity/set', 'Email/set', 'EmailSubmission/set'].includes(call.method)), false);
    });
  }

  it('does not create a sending identity for a draft without a Sent mailbox', async () => {
    hasSentMailbox = false;
    emails.set('existing', { id: 'existing', from: [{ email: MASK }], to: [{ email: 'recipient@example.com' }], keywords: { $draft: true } });
    await assert.rejects(client.callTool({ name: 'send_draft', arguments: { emailId: 'existing' } }), /Sent mailbox/);
    assert.equal(calls.some(call => ['Identity/set', 'Email/set', 'EmailSubmission/set'].includes(call.method)), false);
  });
});
