import { FastmailAuth } from './auth.js';
import { writeFile, mkdir } from 'fs/promises';
import { dirname, resolve, normalize, sep } from 'path';
import { homedir } from 'os';

export interface JmapSession {
  apiUrl: string;
  accountId: string;
  capabilities: Record<string, any>;
  downloadUrl?: string;
  uploadUrl?: string;
}

export interface JmapRequest {
  using: string[];
  methodCalls: [string, any, string][];
}

export interface JmapResponse {
  methodResponses: Array<[string, any, string]>;
  sessionState: string;
}

/** Match an email address against an identity, supporting wildcard identities (e.g. *@example.com). */
function matchesIdentity(identityEmail: string, address: string): boolean {
  const identity = identityEmail.toLowerCase();
  const addr = address.toLowerCase();
  if (identity === addr) return true;
  if (identity.startsWith('*@')) {
    const domain = identity.slice(1); // "@example.com"
    return addr.endsWith(domain) && addr.indexOf('@') > 0;
  }
  return false;
}
export class JmapClient {
  private auth: FastmailAuth;
  private session: JmapSession | null = null;

  constructor(auth: FastmailAuth) {
    this.auth = auth;
  }

  /**
   * Extract the result from a JMAP method response, throwing on method-level errors.
   */
  protected getMethodResult(response: JmapResponse, index: number): any {
    if (!response.methodResponses || index >= response.methodResponses.length) {
      throw new Error(
        `JMAP response missing expected method at index ${index} (got ${response.methodResponses?.length ?? 0} responses)`
      );
    }
    const entry = response.methodResponses[index];
    if (!Array.isArray(entry) || entry.length < 2) {
      throw new Error(`JMAP response entry at index ${index} is malformed`);
    }
    const [tag, result] = entry;
    if (tag === 'error') {
      throw new Error(`JMAP error: ${result.type}${result.description ? ' - ' + result.description : ''}`);
    }
    return result;
  }

  /**
   * Extract the .list array from a JMAP method response, with null safety.
   */
  protected getListResult(response: JmapResponse, index: number): any[] {
    const result = this.getMethodResult(response, index);
    return result?.list || [];
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
      accountId: sessionData.primaryAccounts?.['urn:ietf:params:jmap:mail']
        || sessionData.primaryAccounts?.['urn:ietf:params:jmap:core']
        || Object.keys(sessionData.accounts)[0],
      capabilities: sessionData.capabilities,
      downloadUrl: sessionData.downloadUrl,
      uploadUrl: sessionData.uploadUrl
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

    const data = await response.json();
    if (!data || !Array.isArray(data.methodResponses)) {
      throw new Error('Invalid JMAP response: missing or malformed methodResponses');
    }
    return data as JmapResponse;
  }

  protected findMailboxByRoleOrName(mailboxes: any[], role: string, nameFallback?: string): any | undefined {
    return mailboxes.find(mb => mb.role === role) ||
           (nameFallback ? mailboxes.find(mb => mb.name.toLowerCase().includes(nameFallback)) : undefined);
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
    return this.getListResult(response, 0);
  }

  async getEmails(mailboxId?: string, limit: number = 20, ascending: boolean = false): Promise<any[]> {
    const session = await this.getSession();

    const filter = mailboxId ? { inMailbox: mailboxId } : {};

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/query', {
          accountId: session.accountId,
          filter,
          sort: [{ property: 'receivedAt', isAscending: ascending }],
          limit
        }, 'query'],
        ['Email/get', {
          accountId: session.accountId,
          '#ids': { resultOf: 'query', name: 'Email/query', path: '/ids' },
          properties: ['id', 'subject', 'from', 'to', 'replyTo', 'receivedAt', 'preview', 'hasAttachment']
        }, 'emails']
      ]
    };

    const response = await this.makeRequest(request);
    return this.getListResult(response, 1);
  }

  async getEmailById(id: string): Promise<any> {
    const session = await this.getSession();
    
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/get', {
          accountId: session.accountId,
          ids: [id],
          properties: ['id', 'subject', 'from', 'to', 'cc', 'bcc', 'replyTo', 'receivedAt', 'textBody', 'htmlBody', 'attachments', 'bodyValues', 'messageId', 'threadId', 'inReplyTo', 'references', 'keywords', 'header:List-Unsubscribe:asURLs'],
          bodyProperties: ['partId', 'blobId', 'type', 'size'],
          fetchTextBodyValues: true,
          fetchHTMLBodyValues: true,
        }, 'email']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notFound && result.notFound.includes(id)) {
      throw new Error(`Email with ID '${id}' not found`);
    }

    const email = result.list?.[0];
    if (!email) {
      throw new Error(`Email with ID '${id}' not found or not accessible`);
    }
    
    return email;
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
    return this.getListResult(response, 0);
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
    inReplyTo?: string[];
    references?: string[];
    replyTo?: string[];
  }): Promise<string> {
    const session = await this.getSession();

    // Get all identities to validate from address
    const identities = await this.getIdentities();
    if (!identities || identities.length === 0) {
      throw new Error('No sending identities found');
    }

    // Determine which identity to use
    let selectedIdentity;
    if (email.from) {
      // Validate that the from address matches an available identity
      selectedIdentity = identities.find(id => matchesIdentity(id.email, email.from!));
      if (!selectedIdentity) {
        throw new Error('From address is not verified for sending. Choose one of your verified identities.');
      }
    } else {
      // Use default identity
      selectedIdentity = identities.find(id => id.mayDelete === false) || identities[0];
    }

    // Use the requested from address (not the identity email, which may be a wildcard like *@domain)
    const fromEmail = email.from || selectedIdentity.email;

    // Get the mailbox IDs we need
    const mailboxes = await this.getMailboxes();
    const draftsMailbox = this.findMailboxByRoleOrName(mailboxes, 'drafts', 'draft');
    const sentMailbox = this.findMailboxByRoleOrName(mailboxes, 'sent', 'sent');

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
      from: [{ name: selectedIdentity.name, email: fromEmail }],
      to: email.to.map(addr => ({ email: addr })),
      cc: email.cc?.map(addr => ({ email: addr })) || [],
      bcc: email.bcc?.map(addr => ({ email: addr })) || [],
      subject: email.subject,
      ...(email.inReplyTo && { inReplyTo: email.inReplyTo }),
      ...(email.references && { references: email.references }),
      ...(email.replyTo?.length && { replyTo: email.replyTo.map(addr => ({ email: addr })) }),
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
              identityId: selectedIdentity.id,
              envelope: {
                mailFrom: { email: fromEmail },
                rcptTo: [
                  ...email.to.map(addr => ({ email: addr })),
                  ...(email.cc || []).map(addr => ({ email: addr })),
                  ...(email.bcc || []).map(addr => ({ email: addr })),
                ]
              }
            }
          },
          onSuccessUpdateEmail: {
            '#submission': {
              mailboxIds: sentMailboxIds,
              keywords: { $seen: true }
            }
          }
        }, 'submitEmail']
      ]
    };

    const response = await this.makeRequest(request);

    const emailResult = this.getMethodResult(response, 0);
    if (emailResult.notCreated?.draft) {
      const err = emailResult.notCreated.draft;
      throw new Error(`Failed to create email: ${err.type}${err.description ? ' - ' + err.description : ''}`);
    }

    const emailId = emailResult.created?.draft?.id;
    if (!emailId) {
      throw new Error('Email creation returned no email ID');
    }

    const submissionResult = this.getMethodResult(response, 1);
    if (submissionResult.notCreated?.submission) {
      const err = submissionResult.notCreated.submission;
      throw new Error(`Failed to submit email: ${err.type}${err.description ? ' - ' + err.description : ''}`);
    }

    const submissionId = submissionResult.created?.submission?.id;
    if (!submissionId) {
      throw new Error('Email submission returned no submission ID');
    }

    return submissionId;
  }

  async createDraft(email: {
    to?: string[];
    cc?: string[];
    bcc?: string[];
    subject?: string;
    textBody?: string;
    htmlBody?: string;
    from?: string;
    mailboxId?: string;
    inReplyTo?: string[];
    references?: string[];
    replyTo?: string[];
  }): Promise<string> {
    const session = await this.getSession();

    // Validate at least one meaningful field is present
    if (!email.to?.length && !email.subject && !email.textBody && !email.htmlBody) {
      throw new Error('At least one of to, subject, textBody, or htmlBody must be provided');
    }

    // Get all identities to resolve from address
    const identities = await this.getIdentities();
    if (!identities || identities.length === 0) {
      throw new Error('No sending identities found');
    }

    let selectedIdentity;
    if (email.from) {
      selectedIdentity = identities.find(id => matchesIdentity(id.email, email.from!));
      if (!selectedIdentity) {
        throw new Error('From address is not verified for sending. Choose one of your verified identities.');
      }
    } else {
      selectedIdentity = identities.find(id => id.mayDelete === false) || identities[0];
    }

    const fromEmail = email.from || selectedIdentity.email;

    // Resolve drafts mailbox
    let draftMailboxId: string;
    if (email.mailboxId) {
      draftMailboxId = email.mailboxId;
    } else {
      const mailboxes = await this.getMailboxes();
      const draftsMailbox = this.findMailboxByRoleOrName(mailboxes, 'drafts', 'draft');
      if (!draftsMailbox) {
        throw new Error('Could not find Drafts mailbox');
      }
      draftMailboxId = draftsMailbox.id;
    }

    const mailboxIds: Record<string, boolean> = {};
    mailboxIds[draftMailboxId] = true;

    const emailObject: any = {
      mailboxIds,
      keywords: { $draft: true },
      from: [{ email: fromEmail }],
    };

    if (email.to?.length) emailObject.to = email.to.map(addr => ({ email: addr }));
    if (email.cc?.length) emailObject.cc = email.cc.map(addr => ({ email: addr }));
    if (email.bcc?.length) emailObject.bcc = email.bcc.map(addr => ({ email: addr }));
    if (email.subject) emailObject.subject = email.subject;
    if (email.inReplyTo?.length) emailObject.inReplyTo = email.inReplyTo;
    if (email.references?.length) emailObject.references = email.references;
    if (email.replyTo?.length) emailObject.replyTo = email.replyTo.map(addr => ({ email: addr }));
    if (email.textBody) emailObject.textBody = [{ partId: 'text', type: 'text/plain' }];
    if (email.htmlBody) emailObject.htmlBody = [{ partId: 'html', type: 'text/html' }];
    if (email.textBody || email.htmlBody) {
      emailObject.bodyValues = {
        ...(email.textBody && { text: { value: email.textBody } }),
        ...(email.htmlBody && { html: { value: email.htmlBody } })
      };
    }

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          create: { draft: emailObject }
        }, 'createDraft']
      ]
    };

    const response = await this.makeRequest(request);

    const result = this.getMethodResult(response, 0);

    // Bug 2: Propagate server-provided error details from notCreated
    if (result.notCreated?.draft) {
      const err = result.notCreated.draft;
      throw new Error(`Failed to create draft: ${err.type}${err.description ? ' - ' + err.description : ''}`);
    }

    // Bug 3: Throw if created ID is missing instead of returning 'unknown'
    const emailId = result.created?.draft?.id;
    if (!emailId) {
      throw new Error('Draft creation returned no email ID');
    }

    return emailId;
  }

  async updateDraft(emailId: string, updates: {
    to?: string[];
    cc?: string[];
    bcc?: string[];
    subject?: string;
    textBody?: string;
    htmlBody?: string;
    from?: string;
    replyTo?: string[];
  }): Promise<string> {
    const session = await this.getSession();

    // Fetch the existing email
    const getRequest: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/get', {
          accountId: session.accountId,
          ids: [emailId],
          properties: ['id', 'subject', 'from', 'to', 'cc', 'bcc', 'replyTo', 'textBody', 'htmlBody', 'bodyValues', 'mailboxIds', 'keywords'],
          bodyProperties: ['partId', 'blobId', 'type', 'size'],
          fetchTextBodyValues: true,
          fetchHTMLBodyValues: true,
        }, 'getEmail']
      ]
    };

    const getResponse = await this.makeRequest(getRequest);
    const existingEmail = this.getListResult(getResponse, 0)[0];
    if (!existingEmail) {
      throw new Error(`Email with ID '${emailId}' not found`);
    }

    // Verify it's a draft
    if (!existingEmail.keywords?.$draft) {
      throw new Error('Cannot edit a non-draft email');
    }

    // Resolve identity
    const identities = await this.getIdentities();
    if (!identities || identities.length === 0) {
      throw new Error('No sending identities found');
    }

    let selectedIdentity;
    if (updates.from) {
      selectedIdentity = identities.find(id => matchesIdentity(id.email, updates.from!));
      if (!selectedIdentity) {
        throw new Error('From address is not verified for sending. Choose one of your verified identities.');
      }
    } else {
      // Use existing from, or fall back to default identity
      const existingFrom = existingEmail.from?.[0]?.email;
      if (existingFrom) {
        selectedIdentity = identities.find(id => matchesIdentity(id.email, existingFrom))
          || identities.find(id => id.mayDelete === false) || identities[0];
      } else {
        selectedIdentity = identities.find(id => id.mayDelete === false) || identities[0];
      }
    }

    // Extract existing body values
    const existingTextBody = existingEmail.bodyValues
      ? Object.values(existingEmail.bodyValues).find((bv: any) =>
          existingEmail.textBody?.some((tb: any) => tb.partId === (bv as any).partId || true)
        )
      : null;
    const existingHtmlBody = existingEmail.bodyValues
      ? Object.values(existingEmail.bodyValues).find((bv: any) =>
          existingEmail.htmlBody?.some((hb: any) => hb.partId === (bv as any).partId || true)
        )
      : null;

    // Merge: updates override existing values
    const mergedSubject = updates.subject !== undefined ? updates.subject : (existingEmail.subject || '');
    const mergedTo = updates.to !== undefined ? updates.to.map(addr => ({ email: addr })) : (existingEmail.to || []);
    const mergedCc = updates.cc !== undefined ? updates.cc.map(addr => ({ email: addr })) : (existingEmail.cc || []);
    const mergedBcc = updates.bcc !== undefined ? updates.bcc.map(addr => ({ email: addr })) : (existingEmail.bcc || []);
    const mergedReplyTo = updates.replyTo !== undefined ? updates.replyTo.map(addr => ({ email: addr })) : (existingEmail.replyTo || null);

    const textBodyValue = updates.textBody !== undefined ? updates.textBody : (existingTextBody as any)?.value;
    const htmlBodyValue = updates.htmlBody !== undefined ? updates.htmlBody : (existingHtmlBody as any)?.value;

    const emailObject: any = {
      mailboxIds: existingEmail.mailboxIds,
      keywords: { $draft: true },
      from: [{ email: updates.from || existingEmail.from?.[0]?.email || selectedIdentity.email }],
      to: mergedTo,
      cc: mergedCc,
      bcc: mergedBcc,
      subject: mergedSubject,
      ...(mergedReplyTo?.length && { replyTo: mergedReplyTo }),
    };

    if (textBodyValue) emailObject.textBody = [{ partId: 'text', type: 'text/plain' }];
    if (htmlBodyValue) emailObject.htmlBody = [{ partId: 'html', type: 'text/html' }];
    if (textBodyValue || htmlBodyValue) {
      emailObject.bodyValues = {
        ...(textBodyValue && { text: { value: textBodyValue } }),
        ...(htmlBodyValue && { html: { value: htmlBodyValue } }),
      };
    }

    // Atomic create + destroy in a single Email/set call
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          create: { draft: emailObject },
          destroy: [emailId],
        }, 'updateDraft']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notCreated?.draft) {
      const err = result.notCreated.draft;
      throw new Error(`Failed to create updated draft: ${err.type}${err.description ? ' - ' + err.description : ''}`);
    }

    const newEmailId = result.created?.draft?.id;
    if (!newEmailId) {
      throw new Error('Draft update returned no email ID');
    }

    return newEmailId;
  }

  async sendDraft(emailId: string): Promise<string> {
    const session = await this.getSession();

    // Fetch the existing email to verify it's a draft
    const getRequest: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/get', {
          accountId: session.accountId,
          ids: [emailId],
          properties: ['id', 'from', 'to', 'cc', 'bcc', 'replyTo', 'keywords'],
        }, 'getEmail']
      ]
    };

    const getResponse = await this.makeRequest(getRequest);
    const email = this.getListResult(getResponse, 0)[0];
    if (!email) {
      throw new Error(`Email with ID '${emailId}' not found`);
    }

    if (!email.keywords?.$draft) {
      throw new Error('Cannot send a non-draft email');
    }

    // Collect all recipients for the envelope
    const allRecipients: { email: string }[] = [
      ...(email.to || []),
      ...(email.cc || []),
      ...(email.bcc || []),
    ];

    if (allRecipients.length === 0) {
      throw new Error('Draft has no recipients');
    }

    // Determine identity from the email's from field
    const fromEmail = email.from?.[0]?.email;
    if (!fromEmail) {
      throw new Error('Draft has no from address');
    }

    const identities = await this.getIdentities();
    const selectedIdentity = identities.find(id => matchesIdentity(id.email, fromEmail));
    if (!selectedIdentity) {
      throw new Error('From address on draft does not match any sending identity');
    }

    // Find the Sent mailbox
    const mailboxes = await this.getMailboxes();
    const sentMailbox = this.findMailboxByRoleOrName(mailboxes, 'sent', 'sent');
    if (!sentMailbox) {
      throw new Error('Could not find Sent mailbox');
    }

    const sentMailboxIds: Record<string, boolean> = {};
    sentMailboxIds[sentMailbox.id] = true;

    // Submit the draft
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail', 'urn:ietf:params:jmap:submission'],
      methodCalls: [
        ['EmailSubmission/set', {
          accountId: session.accountId,
          create: {
            submission: {
              emailId,
              identityId: selectedIdentity.id,
              envelope: {
                mailFrom: { email: fromEmail },
                rcptTo: allRecipients.map(addr => ({ email: addr.email })),
              }
            }
          },
          onSuccessUpdateEmail: {
            '#submission': {
              mailboxIds: sentMailboxIds,
              'keywords/$draft': null,
              'keywords/$seen': true,
            }
          }
        }, 'submitDraft']
      ]
    };

    const response = await this.makeRequest(request);
    const submissionResult = this.getMethodResult(response, 0);
    if (submissionResult.notCreated?.submission) {
      const err = submissionResult.notCreated.submission;
      throw new Error(`Failed to submit draft: ${err.type}${err.description ? ' - ' + err.description : ''}`);
    }

    const submissionId = submissionResult.created?.submission?.id;
    if (!submissionId) {
      throw new Error('Draft submission returned no submission ID');
    }

    return submissionId;
  }

  async getRecentEmails(limit: number = 10, mailboxName: string = 'inbox', ascending: boolean = false): Promise<any[]> {
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
          sort: [{ property: 'receivedAt', isAscending: ascending }],
          limit: Math.min(limit, 50)
        }, 'query'],
        ['Email/get', {
          accountId: session.accountId,
          '#ids': { resultOf: 'query', name: 'Email/query', path: '/ids' },
          properties: ['id', 'subject', 'from', 'to', 'replyTo', 'receivedAt', 'preview', 'hasAttachment', 'keywords', 'header:List-Unsubscribe:asURLs']
        }, 'emails']
      ]
    };

    const response = await this.makeRequest(request);
    return this.getListResult(response, 1);
  }

  async markEmailRead(emailId: string, read: boolean = true): Promise<void> {
    const session = await this.getSession();

    const update: Record<string, any> = {};
    update[emailId] = read
      ? { 'keywords/$seen': true }
      : { 'keywords/$seen': null };

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update
        }, 'updateEmail']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);
    
    if (result.notUpdated && result.notUpdated[emailId]) {
      throw new Error(`Failed to mark email as ${read ? 'read' : 'unread'}.`);
    }
  }

  async pinEmail(emailId: string, pinned: boolean = true): Promise<void> {
    const session = await this.getSession();

    const update: Record<string, any> = {};
    update[emailId] = pinned
      ? { 'keywords/$flagged': true }
      : { 'keywords/$flagged': null };

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update
        }, 'pinEmail']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated && result.notUpdated[emailId]) {
      throw new Error(`Failed to ${pinned ? 'pin' : 'unpin'} email.`);
    }
  }

  async deleteEmail(emailId: string): Promise<void> {
    const session = await this.getSession();
    
    // Find the trash mailbox
    const mailboxes = await this.getMailboxes();
    const trashMailbox = this.findMailboxByRoleOrName(mailboxes, 'trash', 'trash');

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
    const result = this.getMethodResult(response, 0);
    
    if (result.notUpdated && result.notUpdated[emailId]) {
      throw new Error('Failed to delete email.');
    }
  }

  async moveEmail(emailId: string, targetMailboxId: string): Promise<void> {
    const session = await this.getSession();

    // Fetch current mailboxIds to build a proper JMAP patch
    const getRequest: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/get', {
          accountId: session.accountId,
          ids: [emailId],
          properties: ['mailboxIds']
        }, 'getEmail']
      ]
    };
    const getResponse = await this.makeRequest(getRequest);
    const email = this.getListResult(getResponse, 0)[0];

    // Build patch: remove from all current mailboxes, add to target
    const patch: Record<string, boolean | null> = {};
    if (email?.mailboxIds) {
      for (const mbId of Object.keys(email.mailboxIds)) {
        patch[`mailboxIds/${mbId}`] = null;
      }
    }
    patch[`mailboxIds/${targetMailboxId}`] = true;

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update: {
            [emailId]: patch
          }
        }, 'moveEmail']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated && result.notUpdated[emailId]) {
      throw new Error('Failed to move email.');
    }
  }

  async addLabels(emailId: string, mailboxIds: string[]): Promise<void> {
    const session = await this.getSession();

    // Build patch object to add specific mailboxIds
    const patch: Record<string, any> = {};
    mailboxIds.forEach(mailboxId => {
      patch[`mailboxIds/${mailboxId}`] = true;
    });

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update: {
            [emailId]: patch
          }
        }, 'addLabels']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated && result.notUpdated[emailId]) {
      throw new Error('Failed to add labels to email.');
    }
  }

  async removeLabels(emailId: string, mailboxIds: string[]): Promise<void> {
    const session = await this.getSession();

    // Build patch object to remove specific mailboxIds
    const patch: Record<string, any> = {};
    mailboxIds.forEach(mailboxId => {
      patch[`mailboxIds/${mailboxId}`] = null;
    });

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update: {
            [emailId]: patch
          }
        }, 'removeLabels']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated && result.notUpdated[emailId]) {
      throw new Error('Failed to remove labels from email.');
    }
  }

  async bulkAddLabels(emailIds: string[], mailboxIds: string[]): Promise<void> {
    const session = await this.getSession();

    // Build patch object to add specific mailboxIds
    const patch: Record<string, any> = {};
    mailboxIds.forEach(mailboxId => {
      patch[`mailboxIds/${mailboxId}`] = true;
    });

    const updates: Record<string, any> = {};
    emailIds.forEach(id => {
      updates[id] = patch;
    });

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update: updates
        }, 'bulkAddLabels']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated && Object.keys(result.notUpdated).length > 0) {
      throw new Error('Failed to add labels to some emails.');
    }
  }

  async bulkRemoveLabels(emailIds: string[], mailboxIds: string[]): Promise<void> {
    const session = await this.getSession();

    // Build patch object to remove specific mailboxIds
    const patch: Record<string, any> = {};
    mailboxIds.forEach(mailboxId => {
      patch[`mailboxIds/${mailboxId}`] = null;
    });

    const updates: Record<string, any> = {};
    emailIds.forEach(id => {
      updates[id] = patch;
    });

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update: updates
        }, 'bulkRemoveLabels']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated && Object.keys(result.notUpdated).length > 0) {
      throw new Error('Failed to remove labels from some emails.');
    }
  }

  async getEmailAttachments(emailId: string): Promise<any[]> {
    const session = await this.getSession();

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/get', {
          accountId: session.accountId,
          ids: [emailId],
          properties: ['attachments']
        }, 'getAttachments']
      ]
    };

    const response = await this.makeRequest(request);
    const email = this.getListResult(response, 0)[0];
    return email?.attachments || [];
  }

  async downloadAttachment(emailId: string, attachmentId: string): Promise<string> {
    const session = await this.getSession();

    // Get the email with full attachment details
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/get', {
          accountId: session.accountId,
          ids: [emailId],
          properties: ['attachments', 'bodyValues'],
          bodyProperties: ['partId', 'blobId', 'size', 'name', 'type']
        }, 'getEmail']
      ]
    };

    const response = await this.makeRequest(request);
    const email = this.getListResult(response, 0)[0];

    if (!email) {
      throw new Error('Email not found');
    }

    // Find attachment by partId or by index
    let attachment = email.attachments?.find((att: any) => 
      att.partId === attachmentId || att.blobId === attachmentId
    );

    // If not found, try by array index
    if (!attachment) {
      const index = parseInt(attachmentId, 10);
      if (!isNaN(index)) {
        attachment = email.attachments?.[index];
      }
    }
    
    if (!attachment) {
      throw new Error('Attachment not found.');
    }

    // Get the download URL from session
    const downloadUrl = session.downloadUrl;
    if (!downloadUrl) {
      throw new Error('Download capability not available in session');
    }

    // Build download URL
    const url = downloadUrl
      .replace('{accountId}', session.accountId)
      .replace('{blobId}', attachment.blobId)
      .replace('{type}', encodeURIComponent(attachment.type || 'application/octet-stream'))
      .replace('{name}', encodeURIComponent(attachment.name || 'attachment'));

    return url;
  }

  static readonly DEFAULT_DOWNLOADS_DIR = resolve(homedir(), 'Downloads', 'fastmail-mcp');

  static validateSavePath(savePath: string, downloadDir?: string): string {
    const allowedDir = downloadDir ? resolve(normalize(downloadDir)) : JmapClient.DEFAULT_DOWNLOADS_DIR;
    const resolved = resolve(normalize(savePath));

    if (resolved.includes('\0')) {
      throw new Error('Save path contains null bytes');
    }

    if (!resolved.startsWith(allowedDir + sep) && resolved !== allowedDir) {
      throw new Error(
        `Save path must be within ${allowedDir}. ` +
        `Received: ${savePath}`
      );
    }

    return resolved;
  }

  async downloadAttachmentToFile(emailId: string, attachmentId: string, savePath: string, downloadDir?: string): Promise<{ url: string; bytesWritten: number }> {
    const validatedPath = JmapClient.validateSavePath(savePath, downloadDir);
    const url = await this.downloadAttachment(emailId, attachmentId);

    const response = await fetch(url, {
      headers: { 'Authorization': this.auth.getAuthHeaders()['Authorization'] }
    });

    if (!response.ok) {
      throw new Error(`Download failed: ${response.status} ${response.statusText}`);
    }

    const buffer = Buffer.from(await response.arrayBuffer());

    await mkdir(dirname(validatedPath), { recursive: true });
    await writeFile(validatedPath, buffer);

    return { url, bytesWritten: buffer.length };
  }

  async advancedSearch(filters: {
    query?: string;
    from?: string;
    to?: string;
    subject?: string;
    hasAttachment?: boolean;
    isUnread?: boolean;
    isPinned?: boolean;
    mailboxId?: string;
    after?: string;
    before?: string;
    limit?: number;
    ascending?: boolean;
  }): Promise<any[]> {
    const session = await this.getSession();
    
    // Build JMAP filter object
    const filter: any = {};
    
    if (filters.query) filter.text = filters.query;
    if (filters.from) filter.from = filters.from;
    if (filters.to) filter.to = filters.to;
    if (filters.subject) filter.subject = filters.subject;
    if (filters.hasAttachment !== undefined) filter.hasAttachment = filters.hasAttachment;
    if (filters.isUnread === true) filter.notKeyword = '$seen';
    else if (filters.isUnread === false) filter.hasKeyword = '$seen';
    if (filters.isPinned === true) filter.hasKeyword = '$flagged';
    if (filters.isPinned === false) filter.notKeyword = '$flagged';
    if (filters.mailboxId) filter.inMailbox = filters.mailboxId;
    if (filters.after) filter.after = filters.after;
    if (filters.before) filter.before = filters.before;

    // When both isUnread and isPinned are set, hasKeyword/notKeyword may conflict.
    // JMAP FilterCondition only supports one hasKeyword, so wrap in an AND operator.
    let finalFilter: any = filter;
    if (filters.isUnread !== undefined && filters.isPinned !== undefined) {
      delete filter.hasKeyword;
      delete filter.notKeyword;
      const conditions: any[] = [filter];
      conditions.push(filters.isUnread ? { notKeyword: '$seen' } : { hasKeyword: '$seen' });
      conditions.push(filters.isPinned ? { hasKeyword: '$flagged' } : { notKeyword: '$flagged' });
      finalFilter = { operator: 'AND', conditions };
    }

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/query', {
          accountId: session.accountId,
          filter: finalFilter,
          sort: [{ property: 'receivedAt', isAscending: filters.ascending ?? false }],
          limit: Math.min(filters.limit || 50, 100)
        }, 'query'],
        ['Email/get', {
          accountId: session.accountId,
          '#ids': { resultOf: 'query', name: 'Email/query', path: '/ids' },
          properties: ['id', 'subject', 'from', 'to', 'cc', 'replyTo', 'receivedAt', 'preview', 'hasAttachment', 'keywords', 'threadId']
        }, 'emails']
      ]
    };

    const response = await this.makeRequest(request);
    return this.getListResult(response, 1);
  }

  async searchEmails(query: string, limit: number = 20, ascending: boolean = false): Promise<any[]> {
    const session = await this.getSession();

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/query', {
          accountId: session.accountId,
          filter: { text: query },
          sort: [{ property: 'receivedAt', isAscending: ascending }],
          limit
        }, 'query'],
        ['Email/get', {
          accountId: session.accountId,
          '#ids': { resultOf: 'query', name: 'Email/query', path: '/ids' },
          properties: ['id', 'subject', 'from', 'to', 'replyTo', 'receivedAt', 'preview', 'hasAttachment']
        }, 'emails']
      ]
    };

    const response = await this.makeRequest(request);
    return this.getListResult(response, 1);
  }

  async getThread(threadId: string): Promise<any[]> {
    const session = await this.getSession();

    // First, check if threadId is actually an email ID and resolve the thread
    let actualThreadId = threadId;
    
    // Try to get the email first to see if we need to resolve thread ID
    try {
      const emailRequest: JmapRequest = {
        using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
        methodCalls: [
          ['Email/get', {
            accountId: session.accountId,
            ids: [threadId],
            properties: ['threadId']
          }, 'checkEmail']
        ]
      };
      
      const emailResponse = await this.makeRequest(emailRequest);
      const email = this.getListResult(emailResponse, 0)[0];

      if (email && email.threadId) {
        actualThreadId = email.threadId;
      }
    } catch (error) {
      // If email lookup fails, assume threadId is correct
    }

    // Use Thread/get with the resolved thread ID
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Thread/get', {
          accountId: session.accountId,
          ids: [actualThreadId]
        }, 'getThread'],
        ['Email/get', {
          accountId: session.accountId,
          '#ids': { resultOf: 'getThread', name: 'Thread/get', path: '/list/*/emailIds' },
          properties: ['id', 'subject', 'from', 'to', 'cc', 'replyTo', 'receivedAt', 'preview', 'hasAttachment', 'keywords', 'threadId']
        }, 'emails']
      ]
    };

    const response = await this.makeRequest(request);
    const threadResult = this.getMethodResult(response, 0);

    // Check if thread was found
    if (threadResult.notFound && threadResult.notFound.includes(actualThreadId)) {
      throw new Error(`Thread with ID '${actualThreadId}' not found`);
    }

    return this.getListResult(response, 1);
  }

  async getMailboxStats(mailboxId?: string): Promise<any> {
    const session = await this.getSession();
    
    if (mailboxId) {
      // Get stats for specific mailbox
      const request: JmapRequest = {
        using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
        methodCalls: [
          ['Mailbox/get', {
            accountId: session.accountId,
            ids: [mailboxId],
            properties: ['id', 'name', 'role', 'totalEmails', 'unreadEmails', 'totalThreads', 'unreadThreads']
          }, 'mailbox']
        ]
      };

      const response = await this.makeRequest(request);
      return this.getListResult(response, 0)[0];
    } else {
      // Get stats for all mailboxes
      const mailboxes = await this.getMailboxes();
      return mailboxes.map(mb => ({
        id: mb.id,
        name: mb.name,
        role: mb.role,
        totalEmails: mb.totalEmails || 0,
        unreadEmails: mb.unreadEmails || 0,
        totalThreads: mb.totalThreads || 0,
        unreadThreads: mb.unreadThreads || 0
      }));
    }
  }

  async getAccountSummary(): Promise<any> {
    const session = await this.getSession();
    const mailboxes = await this.getMailboxes();
    const identities = await this.getIdentities();

    // Calculate totals
    const totals = mailboxes.reduce((acc, mb) => ({
      totalEmails: acc.totalEmails + (mb.totalEmails || 0),
      unreadEmails: acc.unreadEmails + (mb.unreadEmails || 0),
      totalThreads: acc.totalThreads + (mb.totalThreads || 0),
      unreadThreads: acc.unreadThreads + (mb.unreadThreads || 0)
    }), { totalEmails: 0, unreadEmails: 0, totalThreads: 0, unreadThreads: 0 });

    return {
      accountId: session.accountId,
      mailboxCount: mailboxes.length,
      identityCount: identities.length,
      ...totals,
      mailboxes: mailboxes.map(mb => ({
        id: mb.id,
        name: mb.name,
        role: mb.role,
        totalEmails: mb.totalEmails || 0,
        unreadEmails: mb.unreadEmails || 0
      }))
    };
  }

  async bulkMarkRead(emailIds: string[], read: boolean = true): Promise<void> {
    const session = await this.getSession();

    const updates: Record<string, any> = {};
    emailIds.forEach(id => {
      updates[id] = read
        ? { 'keywords/$seen': true }
        : { 'keywords/$seen': null };
    });

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update: updates
        }, 'bulkUpdate']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);
    
    if (result.notUpdated && Object.keys(result.notUpdated).length > 0) {
      throw new Error('Failed to update some emails.');
    }
  }

  async bulkPinEmails(emailIds: string[], pinned: boolean = true): Promise<void> {
    const session = await this.getSession();

    const updates: Record<string, any> = {};
    emailIds.forEach(id => {
      updates[id] = pinned
        ? { 'keywords/$flagged': true }
        : { 'keywords/$flagged': null };
    });

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update: updates
        }, 'bulkFlag']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated && Object.keys(result.notUpdated).length > 0) {
      throw new Error('Failed to pin/unpin some emails.');
    }
  }

  async bulkMove(emailIds: string[], targetMailboxId: string): Promise<void> {
    const session = await this.getSession();

    // Fetch current mailboxIds for all emails to build proper JMAP patches
    const getRequest: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/get', {
          accountId: session.accountId,
          ids: emailIds,
          properties: ['id', 'mailboxIds']
        }, 'getEmails']
      ]
    };
    const getResponse = await this.makeRequest(getRequest);
    const emails: any[] = this.getListResult(getResponse, 0);
    const mailboxMap: Record<string, Record<string, boolean>> = {};
    emails.forEach((e: any) => { mailboxMap[e.id] = e.mailboxIds || {}; });

    // Build patch per email: remove all current mailboxes, add target
    const updates: Record<string, any> = {};
    emailIds.forEach(id => {
      const patch: Record<string, boolean | null> = {};
      for (const mbId of Object.keys(mailboxMap[id] || {})) {
        patch[`mailboxIds/${mbId}`] = null;
      }
      patch[`mailboxIds/${targetMailboxId}`] = true;
      updates[id] = patch;
    });

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update: updates
        }, 'bulkMove']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated && Object.keys(result.notUpdated).length > 0) {
      throw new Error('Failed to move some emails.');
    }
  }

  async bulkDelete(emailIds: string[]): Promise<void> {
    const session = await this.getSession();

    // Find the trash mailbox
    const mailboxes = await this.getMailboxes();
    const trashMailbox = this.findMailboxByRoleOrName(mailboxes, 'trash', 'trash');

    if (!trashMailbox) {
      throw new Error('Could not find Trash mailbox');
    }

    const trashMailboxIds: Record<string, boolean> = {};
    trashMailboxIds[trashMailbox.id] = true;

    const updates: Record<string, any> = {};
    emailIds.forEach(id => {
      updates[id] = { mailboxIds: trashMailboxIds };
    });

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update: updates
        }, 'bulkDelete']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated && Object.keys(result.notUpdated).length > 0) {
      throw new Error('Failed to delete some emails.');
    }
  }
}