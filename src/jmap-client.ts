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
    const session = await this.getSession();
    
    // Try to get the user's primary email from the Identity object
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:submission'],
      methodCalls: [
        ['Identity/get', {
          accountId: session.accountId
        }, 'identities']
      ]
    };

    try {
      const response = await this.makeRequest(request);
      const identities = response.methodResponses[0][1].list;
      
      // Find the default identity or use the first one
      const defaultIdentity = identities.find((id: any) => id.mayDelete === false) || identities[0];
      return defaultIdentity?.email || 'user@example.com';
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

  async sendEmail(email: {
    to: string[];
    cc?: string[];
    bcc?: string[];
    subject: string;
    textBody?: string;
    htmlBody?: string;
    from?: string;
  }): Promise<string> {
    const session = await this.getSession();

    // Get the primary email address from session if not provided
    const fromEmail = email.from || await this.getUserEmail();

    const emailObject = {
      from: [{ email: fromEmail }],
      to: email.to.map(addr => ({ email: addr })),
      cc: email.cc?.map(addr => ({ email: addr })) || [],
      bcc: email.bcc?.map(addr => ({ email: addr })) || [],
      subject: email.subject,
      textBody: email.textBody ? [{ partId: 'text', type: 'text/plain' }] : [],
      htmlBody: email.htmlBody ? [{ partId: 'html', type: 'text/html' }] : [],
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
              envelope: {
                mailFrom: { email: fromEmail },
                rcptTo: email.to.map(addr => ({ email: addr }))
              }
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
}