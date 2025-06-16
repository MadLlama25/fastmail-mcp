import fetch from 'node-fetch';
import { FastmailAuth } from './auth.js';

export interface JmapSession {
  apiUrl: string;
  accountId: string;
  capabilities: Record<string, any>;
}

export interface JmapRequest {
  using: string[];
  methodCalls: [string, any, string][];
}

export interface JmapResponse {
  methodResponses: Array<[string, any, string]>;
  sessionState: string;
}

export class JmapClient {
  private auth: FastmailAuth;
  private session: JmapSession | null = null;

  constructor(auth: FastmailAuth) {
    this.auth = auth;
  }

  async getSession(): Promise<JmapSession> {
    if (this.session) {
      return this.session;
    }

    const response = await fetch(this.auth.getSessionUrl(), {
      method: 'GET',
      headers: this.auth.getAuthHeaders()
    });

    if (!response.ok) {
      throw new Error(`Failed to get session: ${response.statusText}`);
    }

    const sessionData = await response.json() as any;
    
    this.session = {
      apiUrl: sessionData.apiUrl,
      accountId: Object.keys(sessionData.accounts)[0],
      capabilities: sessionData.capabilities
    };

    return this.session;
  }

  async getUserEmail(): Promise<string> {
    try {
      const identity = await this.getDefaultIdentity();
      return identity?.email || 'user@example.com';
    } catch (error) {
      // Fallback if Identity/get is not available
      return 'user@example.com';
    }
  }

  async makeRequest(request: JmapRequest): Promise<JmapResponse> {
    const session = await this.getSession();
    
    const response = await fetch(session.apiUrl, {
      method: 'POST',
      headers: this.auth.getAuthHeaders(),
      body: JSON.stringify(request)
    });

    if (!response.ok) {
      throw new Error(`JMAP request failed: ${response.statusText}`);
    }

    return await response.json() as JmapResponse;
  }

  async getMailboxes(): Promise<any[]> {
    const session = await this.getSession();
    
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Mailbox/get', { accountId: session.accountId }, 'mailboxes']
      ]
    };

    const response = await this.makeRequest(request);
    return response.methodResponses[0][1].list;
  }

  async getEmails(mailboxId?: string, limit: number = 20): Promise<any[]> {
    const session = await this.getSession();
    
    const filter = mailboxId ? { inMailbox: mailboxId } : {};
    
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/query', {
          accountId: session.accountId,
          filter,
          sort: [{ property: 'receivedAt', isAscending: false }],
          limit
        }, 'query'],
        ['Email/get', {
          accountId: session.accountId,
          '#ids': { resultOf: 'query', name: 'Email/query', path: '/ids' },
          properties: ['id', 'subject', 'from', 'to', 'receivedAt', 'preview', 'hasAttachment']
        }, 'emails']
      ]
    };

    const response = await this.makeRequest(request);
    return response.methodResponses[1][1].list;
  }

  async getEmailById(id: string): Promise<any> {
    const session = await this.getSession();
    
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/get', {
          accountId: session.accountId,
          ids: [id],
          properties: ['id', 'subject', 'from', 'to', 'cc', 'bcc', 'receivedAt', 'textBody', 'htmlBody', 'attachments']
        }, 'email']
      ]
    };

    const response = await this.makeRequest(request);
    return response.methodResponses[0][1].list[0];
  }

  async getIdentities(): Promise<any[]> {
    const session = await this.getSession();
    
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:submission'],
      methodCalls: [
        ['Identity/get', {
          accountId: session.accountId
        }, 'identities']
      ]
    };

    const response = await this.makeRequest(request);
    return response.methodResponses[0][1].list;
  }

  async getDefaultIdentity(): Promise<any> {
    const identities = await this.getIdentities();
    
    // Find the default identity (usually the one that can't be deleted)
    return identities.find((id: any) => id.mayDelete === false) || identities[0];
  }

  async sendEmail(email: {
    to: string[];
    cc?: string[];
    bcc?: string[];
    subject: string;
    textBody?: string;
    htmlBody?: string;
    from?: string;
    mailboxId?: string;
  }): Promise<string> {
    const session = await this.getSession();

    // Get the default identity for sending
    const identity = await this.getDefaultIdentity();
    if (!identity) {
      throw new Error('No sending identity found');
    }

    // Use the identity email or the provided from email
    const fromEmail = email.from || identity.email;

    // Get the mailbox IDs we need
    const mailboxes = await this.getMailboxes();
    const draftsMailbox = mailboxes.find(mb => mb.role === 'drafts') || mailboxes.find(mb => mb.name.toLowerCase().includes('draft'));
    const sentMailbox = mailboxes.find(mb => mb.role === 'sent') || mailboxes.find(mb => mb.name.toLowerCase().includes('sent'));
    
    if (!draftsMailbox) {
      throw new Error('Could not find Drafts mailbox to save email');
    }
    if (!sentMailbox) {
      throw new Error('Could not find Sent mailbox to move email after sending');
    }

    // Use provided mailboxId or default to drafts for initial creation
    const initialMailboxId = email.mailboxId || draftsMailbox.id;

    // Ensure we have at least one body type
    if (!email.textBody && !email.htmlBody) {
      throw new Error('Either textBody or htmlBody must be provided');
    }

    const initialMailboxIds: Record<string, boolean> = {};
    initialMailboxIds[initialMailboxId] = true;

    const sentMailboxIds: Record<string, boolean> = {};
    sentMailboxIds[sentMailbox.id] = true;

    const emailObject = {
      mailboxIds: initialMailboxIds,
      keywords: { $draft: true },
      from: [{ email: fromEmail }],
      to: email.to.map(addr => ({ email: addr })),
      cc: email.cc?.map(addr => ({ email: addr })) || [],
      bcc: email.bcc?.map(addr => ({ email: addr })) || [],
      subject: email.subject,
      textBody: email.textBody ? [{ partId: 'text', type: 'text/plain' }] : undefined,
      htmlBody: email.htmlBody ? [{ partId: 'html', type: 'text/html' }] : undefined,
      bodyValues: {
        ...(email.textBody && { text: { value: email.textBody } }),
        ...(email.htmlBody && { html: { value: email.htmlBody } })
      }
    };

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail', 'urn:ietf:params:jmap:submission'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          create: { draft: emailObject }
        }, 'createEmail'],
        ['EmailSubmission/set', {
          accountId: session.accountId,
          create: {
            submission: {
              emailId: '#draft',
              identityId: identity.id,
              envelope: {
                mailFrom: { email: fromEmail },
                rcptTo: email.to.map(addr => ({ email: addr }))
              }
            }
          },
          onSuccessUpdateEmail: {
            '#submission': {
              mailboxIds: sentMailboxIds,
              keywords: {}
            }
          }
        }, 'submitEmail']
      ]
    };

    const response = await this.makeRequest(request);
    
    // Check if email creation was successful
    const emailResult = response.methodResponses[0][1];
    if (emailResult.notCreated && emailResult.notCreated.draft) {
      throw new Error(`Failed to create email: ${JSON.stringify(emailResult.notCreated.draft)}`);
    }
    
    // Check if email submission was successful
    const submissionResult = response.methodResponses[1][1];
    if (submissionResult.notCreated && submissionResult.notCreated.submission) {
      throw new Error(`Failed to submit email: ${JSON.stringify(submissionResult.notCreated.submission)}`);
    }
    
    return submissionResult.created?.submission?.id || 'unknown';
  }

  async getRecentEmails(limit: number = 10, mailboxName: string = 'inbox'): Promise<any[]> {
    const session = await this.getSession();
    
    // Find the specified mailbox (default to inbox)
    const mailboxes = await this.getMailboxes();
    const targetMailbox = mailboxes.find(mb => 
      mb.role === mailboxName.toLowerCase() || 
      mb.name.toLowerCase().includes(mailboxName.toLowerCase())
    );
    
    if (!targetMailbox) {
      throw new Error(`Could not find mailbox: ${mailboxName}`);
    }

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/query', {
          accountId: session.accountId,
          filter: { inMailbox: targetMailbox.id },
          sort: [{ property: 'receivedAt', isAscending: false }],
          limit: Math.min(limit, 50)
        }, 'query'],
        ['Email/get', {
          accountId: session.accountId,
          '#ids': { resultOf: 'query', name: 'Email/query', path: '/ids' },
          properties: ['id', 'subject', 'from', 'to', 'receivedAt', 'preview', 'hasAttachment', 'keywords']
        }, 'emails']
      ]
    };

    const response = await this.makeRequest(request);
    return response.methodResponses[1][1].list;
  }

  async markEmailRead(emailId: string, read: boolean = true): Promise<void> {
    const session = await this.getSession();
    
    const keywords = read ? { $seen: true } : {};
    
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update: {
            [emailId]: {
              keywords
            }
          }
        }, 'updateEmail']
      ]
    };

    const response = await this.makeRequest(request);
    const result = response.methodResponses[0][1];
    
    if (result.notUpdated && result.notUpdated[emailId]) {
      throw new Error(`Failed to mark email as ${read ? 'read' : 'unread'}: ${JSON.stringify(result.notUpdated[emailId])}`);
    }
  }

  async deleteEmail(emailId: string): Promise<void> {
    const session = await this.getSession();
    
    // Find the trash mailbox
    const mailboxes = await this.getMailboxes();
    const trashMailbox = mailboxes.find(mb => mb.role === 'trash') || mailboxes.find(mb => mb.name.toLowerCase().includes('trash'));
    
    if (!trashMailbox) {
      throw new Error('Could not find Trash mailbox');
    }

    const trashMailboxIds: Record<string, boolean> = {};
    trashMailboxIds[trashMailbox.id] = true;

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update: {
            [emailId]: {
              mailboxIds: trashMailboxIds
            }
          }
        }, 'moveToTrash']
      ]
    };

    const response = await this.makeRequest(request);
    const result = response.methodResponses[0][1];
    
    if (result.notUpdated && result.notUpdated[emailId]) {
      throw new Error(`Failed to delete email: ${JSON.stringify(result.notUpdated[emailId])}`);
    }
  }

  async moveEmail(emailId: string, targetMailboxId: string): Promise<void> {
    const session = await this.getSession();

    const targetMailboxIds: Record<string, boolean> = {};
    targetMailboxIds[targetMailboxId] = true;

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update: {
            [emailId]: {
              mailboxIds: targetMailboxIds
            }
          }
        }, 'moveEmail']
      ]
    };

    const response = await this.makeRequest(request);
    const result = response.methodResponses[0][1];
    
    if (result.notUpdated && result.notUpdated[emailId]) {
      throw new Error(`Failed to move email: ${JSON.stringify(result.notUpdated[emailId])}`);
    }
  }
}