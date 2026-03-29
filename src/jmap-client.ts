import { mkdir, writeFile } from 'node:fs/promises';
import { homedir } from 'node:os';
import { dirname, normalize, resolve } from 'node:path';
import type { FastmailAuth } from './auth.js';

export interface JmapSession {
  apiUrl: string;
  accountId: string;
  capabilities: Record<string, unknown>;
  downloadUrl?: string;
  uploadUrl?: string;
}

export interface JmapRequest {
  using: string[];
  methodCalls: [string, Record<string, unknown>, string][];
}

export interface JmapResponse {
  methodResponses: Array<[string, Record<string, unknown>, string]>;
  sessionState: string;
}

/** Represents a JMAP mailbox (folder/label). */
interface JmapMailbox {
  id: string;
  name: string;
  role: string | null;
  totalEmails?: number;
  unreadEmails?: number;
  totalThreads?: number;
  unreadThreads?: number;
  [key: string]: unknown;
}

/** Represents a JMAP identity used for sending email. */
interface JmapIdentity {
  id: string;
  email: string;
  name?: string;
  mayDelete?: boolean;
  [key: string]: unknown;
}

/** Represents a JMAP email attachment. */
interface JmapAttachment {
  partId?: string;
  blobId: string;
  type?: string;
  size?: number;
  name?: string;
  [key: string]: unknown;
}

/** Represents a JMAP email object. */
interface JmapEmail {
  id: string;
  subject: string;
  from?: { name?: string; email: string }[];
  to?: { email: string }[];
  cc?: { email: string }[];
  bcc?: { email: string }[];
  receivedAt: string;
  preview?: string;
  hasAttachment?: boolean;
  keywords?: Record<string, boolean>;
  threadId?: string;
  messageId?: string[];
  inReplyTo?: string[];
  references?: string[];
  textBody?: { partId: string; blobId?: string; type?: string; size?: number }[];
  htmlBody?: { partId: string; blobId?: string; type?: string; size?: number }[];
  bodyValues?: Record<string, { value: string; [key: string]: unknown }>;
  attachments?: JmapAttachment[];
  mailboxIds?: Record<string, boolean>;
  [key: string]: unknown;
}

/** Session data returned by the JMAP server. */
interface JmapSessionData {
  apiUrl: string;
  primaryAccounts?: Record<string, string>;
  accounts: Record<string, unknown>;
  capabilities: Record<string, unknown>;
  downloadUrl?: string;
  uploadUrl?: string;
}

/** Result of a JMAP method call. */
interface JmapMethodResult {
  list?: Record<string, unknown>[];
  created?: Record<string, { id: string; [key: string]: unknown }>;
  notCreated?: Record<string, { type: string; description?: string }>;
  updated?: Record<string, unknown>;
  notUpdated?: Record<string, { type: string; description?: string }>;
  notFound?: string[];
  [key: string]: unknown;
}

/** JMAP email filter conditions for Email/query. */
interface JmapEmailFilter {
  text?: string;
  from?: string;
  to?: string;
  subject?: string;
  hasAttachment?: boolean;
  hasKeyword?: string;
  notKeyword?: string;
  inMailbox?: string;
  after?: string;
  before?: string;
  operator?: string;
  conditions?: JmapEmailFilter[];
  [key: string]: unknown;
}

/** Email object shape used in Email/set create calls. */
interface JmapEmailCreate {
  mailboxIds: Record<string, boolean>;
  keywords: Record<string, boolean>;
  from: { name?: string; email: string }[];
  to?: { email: string }[];
  cc?: { email: string }[];
  bcc?: { email: string }[];
  subject?: string;
  inReplyTo?: string[];
  references?: string[];
  textBody?: { partId: string; type: string }[];
  htmlBody?: { partId: string; type: string }[];
  bodyValues?: Record<string, { value: string }>;
}

export class JmapClient {
  private auth: FastmailAuth;
  private session: JmapSession | null = null;

  constructor(auth: FastmailAuth) {
    this.auth = auth;
  }

  /** Validate an email address: must contain @ and have no control characters. */
  static validateEmailAddress(email: string): void {
    const hasControlChar = email.split('').some((c) => {
      const code = c.charCodeAt(0);
      return code <= 31 || code === 127;
    });
    if (!email || !email.includes('@') || hasControlChar) {
      throw new Error(`Invalid email address: ${email}`);
    }
  }

  /** Validate all addresses in an array (no-op on undefined), throwing on the first invalid one. */
  private static validateEmailAddresses(addrs: string[] | undefined): void {
    if (!addrs) return;
    for (const addr of addrs) {
      JmapClient.validateEmailAddress(addr);
    }
  }

  /**
   * Extract the result from a JMAP method response, throwing on method-level errors.
   */
  protected getMethodResult(response: JmapResponse, index: number): JmapMethodResult {
    if (!response.methodResponses || index >= response.methodResponses.length) {
      throw new Error(
        `JMAP response missing expected method at index ${index} (got ${response.methodResponses?.length ?? 0} responses)`,
      );
    }
    const entry = response.methodResponses[index];
    if (!Array.isArray(entry) || entry.length < 2) {
      throw new Error(`JMAP response entry at index ${index} is malformed`);
    }
    const [tag, result] = entry;
    if (tag === 'error') {
      const errResult = result as Record<string, unknown>;
      throw new Error(
        `JMAP error: ${errResult.type}${errResult.description ? ` - ${errResult.description}` : ''}`,
      );
    }
    return result as JmapMethodResult;
  }

  /**
   * Extract the .list array from a JMAP method response, with null safety.
   */
  protected getListResult(response: JmapResponse, index: number): Record<string, unknown>[] {
    const result = this.getMethodResult(response, index);
    return result?.list || [];
  }

  async getSession(): Promise<JmapSession> {
    if (this.session) {
      return this.session;
    }

    const response = await fetch(this.auth.getSessionUrl(), {
      method: 'GET',
      headers: this.auth.getAuthHeaders(),
    });

    if (response.status === 401 || response.status === 403) {
      throw new Error(`Authentication failed (${response.status}): check your API credentials`);
    }

    if (!response.ok) {
      throw new Error(`Failed to get session: ${response.statusText}`);
    }

    const sessionData = (await response.json()) as JmapSessionData;

    this.session = {
      apiUrl: sessionData.apiUrl,
      accountId:
        sessionData.primaryAccounts?.['urn:ietf:params:jmap:mail'] ||
        sessionData.primaryAccounts?.['urn:ietf:params:jmap:core'] ||
        Object.keys(sessionData.accounts)[0],
      capabilities: sessionData.capabilities,
      downloadUrl: sessionData.downloadUrl,
      uploadUrl: sessionData.uploadUrl,
    };

    return this.session;
  }

  async makeRequest(request: JmapRequest): Promise<JmapResponse> {
    let session = await this.getSession();

    let response = await fetch(session.apiUrl, {
      method: 'POST',
      headers: this.auth.getAuthHeaders(),
      body: JSON.stringify(request),
    });

    // Retry once on auth failure — session may have expired
    if (response.status === 401 || response.status === 403) {
      this.session = null;
      session = await this.getSession();
      response = await fetch(session.apiUrl, {
        method: 'POST',
        headers: this.auth.getAuthHeaders(),
        body: JSON.stringify(request),
      });
    }

    if (!response.ok) {
      throw new Error(`JMAP request failed: ${response.statusText}`);
    }

    const data = await response.json();
    if (!data || !Array.isArray(data.methodResponses)) {
      throw new Error('Invalid JMAP response: missing or malformed methodResponses');
    }
    return data as JmapResponse;
  }

  protected findMailboxByRoleOrName(
    mailboxes: JmapMailbox[],
    role: string,
    nameFallback?: string,
  ): JmapMailbox | undefined {
    return (
      mailboxes.find((mb) => mb.role === role) ||
      (nameFallback
        ? mailboxes.find((mb) => mb.name.toLowerCase().includes(nameFallback))
        : undefined)
    );
  }

  async getMailboxes(): Promise<JmapMailbox[]> {
    const session = await this.getSession();

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [['Mailbox/get', { accountId: session.accountId }, 'mailboxes']],
    };

    const response = await this.makeRequest(request);
    return this.getListResult(response, 0) as JmapMailbox[];
  }

  async getEmails(
    mailboxId?: string,
    limit = 20,
    sortAscending = false,
    position = 0,
  ): Promise<JmapEmail[]> {
    const session = await this.getSession();

    const filter = mailboxId ? { inMailbox: mailboxId } : {};

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        [
          'Email/query',
          {
            accountId: session.accountId,
            filter,
            sort: [{ property: 'receivedAt', isAscending: sortAscending }],
            limit,
            ...(position > 0 && { position }),
          },
          'query',
        ],
        [
          'Email/get',
          {
            accountId: session.accountId,
            '#ids': { resultOf: 'query', name: 'Email/query', path: '/ids' },
            properties: ['id', 'subject', 'from', 'to', 'receivedAt', 'preview', 'hasAttachment'],
          },
          'emails',
        ],
      ],
    };

    const response = await this.makeRequest(request);
    return this.getListResult(response, 1) as JmapEmail[];
  }

  async getEmailById(id: string): Promise<JmapEmail> {
    const session = await this.getSession();

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        [
          'Email/get',
          {
            accountId: session.accountId,
            ids: [id],
            properties: [
              'id',
              'subject',
              'from',
              'to',
              'cc',
              'bcc',
              'receivedAt',
              'textBody',
              'htmlBody',
              'attachments',
              'bodyValues',
              'messageId',
              'threadId',
              'inReplyTo',
              'references',
              'keywords',
              'header:List-Unsubscribe:asURLs',
            ],
            bodyProperties: ['partId', 'blobId', 'type', 'size'],
            fetchTextBodyValues: true,
            fetchHTMLBodyValues: true,
          },
          'email',
        ],
      ],
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notFound?.includes(id)) {
      throw new Error(`Email with ID '${id}' not found`);
    }

    const email = result.list?.[0];
    if (!email) {
      throw new Error(`Email with ID '${id}' not found or not accessible`);
    }

    return email as JmapEmail;
  }

  async getIdentities(): Promise<JmapIdentity[]> {
    const session = await this.getSession();

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:submission'],
      methodCalls: [
        [
          'Identity/get',
          {
            accountId: session.accountId,
          },
          'identities',
        ],
      ],
    };

    const response = await this.makeRequest(request);
    return this.getListResult(response, 0) as JmapIdentity[];
  }

  async getDefaultIdentity(): Promise<JmapIdentity> {
    const identities = await this.getIdentities();

    // Find the default identity (usually the one that can't be deleted)
    return identities.find((id) => id.mayDelete === false) || identities[0];
  }

  /**
   * Resolve a sending identity from a list. If `fromAddress` is provided, find the matching
   * identity or throw. Otherwise return the default (non-deletable) identity.
   */
  protected resolveIdentity(identities: JmapIdentity[], fromAddress?: string): JmapIdentity {
    if (!identities || identities.length === 0) {
      throw new Error('No sending identities found');
    }

    if (fromAddress) {
      const match = identities.find((id) => id.email.toLowerCase() === fromAddress.toLowerCase());
      if (!match) {
        throw new Error(
          'From address is not verified for sending. Choose one of your verified identities.',
        );
      }
      return match;
    }

    return identities.find((id) => id.mayDelete === false) || identities[0];
  }

  /**
   * Build body parts (textBody, htmlBody, bodyValues) for an email create object.
   */
  protected buildBodyParts(
    textBody?: string,
    htmlBody?: string,
  ): {
    textBody?: { partId: string; type: string }[];
    htmlBody?: { partId: string; type: string }[];
    bodyValues?: Record<string, { value: string }>;
  } {
    const result: {
      textBody?: { partId: string; type: string }[];
      htmlBody?: { partId: string; type: string }[];
      bodyValues?: Record<string, { value: string }>;
    } = {};

    if (textBody) result.textBody = [{ partId: 'text', type: 'text/plain' }];
    if (htmlBody) result.htmlBody = [{ partId: 'html', type: 'text/html' }];
    if (textBody || htmlBody) {
      result.bodyValues = {
        ...(textBody && { text: { value: textBody } }),
        ...(htmlBody && { html: { value: htmlBody } }),
      };
    }

    return result;
  }

  /**
   * Check a JMAP Email/set result for creation errors and return the created email ID.
   */
  protected extractCreatedEmailId(result: JmapMethodResult, label: string): string {
    if (result.notCreated?.draft) {
      const err = result.notCreated.draft;
      throw new Error(
        `Failed to create ${label}: ${err.type}${err.description ? ` - ${err.description}` : ''}`,
      );
    }

    const emailId = result.created?.draft?.id;
    if (!emailId) {
      throw new Error(`${label} creation returned no email ID`);
    }

    return emailId;
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
  }): Promise<string> {
    JmapClient.validateEmailAddresses(email.to);
    JmapClient.validateEmailAddresses(email.cc);
    JmapClient.validateEmailAddresses(email.bcc);

    const session = await this.getSession();

    const identities = await this.getIdentities();
    const selectedIdentity = this.resolveIdentity(identities, email.from);
    const fromEmail = selectedIdentity.email;

    const mailboxes = await this.getMailboxes();
    const draftsMailbox = this.findMailboxByRoleOrName(mailboxes, 'drafts', 'draft');
    const sentMailbox = this.findMailboxByRoleOrName(mailboxes, 'sent', 'sent');

    if (!draftsMailbox) {
      throw new Error('Could not find Drafts mailbox to save email');
    }
    if (!sentMailbox) {
      throw new Error('Could not find Sent mailbox to move email after sending');
    }
    if (!email.textBody && !email.htmlBody) {
      throw new Error('Either textBody or htmlBody must be provided');
    }

    const initialMailboxId = email.mailboxId || draftsMailbox.id;
    const initialMailboxIds: Record<string, boolean> = { [initialMailboxId]: true };
    const sentMailboxIds: Record<string, boolean> = { [sentMailbox.id]: true };

    const emailObject: JmapEmailCreate = {
      mailboxIds: initialMailboxIds,
      keywords: { $draft: true },
      from: [{ name: selectedIdentity.name, email: fromEmail }],
      to: email.to.map((addr) => ({ email: addr })),
      cc: email.cc?.map((addr) => ({ email: addr })) || [],
      bcc: email.bcc?.map((addr) => ({ email: addr })) || [],
      subject: email.subject,
      ...(email.inReplyTo && { inReplyTo: email.inReplyTo }),
      ...(email.references && { references: email.references }),
      ...this.buildBodyParts(email.textBody, email.htmlBody),
    };

    const request: JmapRequest = {
      using: [
        'urn:ietf:params:jmap:core',
        'urn:ietf:params:jmap:mail',
        'urn:ietf:params:jmap:submission',
      ],
      methodCalls: [
        [
          'Email/set',
          {
            accountId: session.accountId,
            create: { draft: emailObject },
          },
          'createEmail',
        ],
        [
          'EmailSubmission/set',
          {
            accountId: session.accountId,
            create: {
              submission: {
                emailId: '#draft',
                identityId: selectedIdentity.id,
                envelope: {
                  mailFrom: { email: fromEmail },
                  rcptTo: [
                    ...email.to.map((addr) => ({ email: addr })),
                    ...(email.cc || []).map((addr) => ({ email: addr })),
                    ...(email.bcc || []).map((addr) => ({ email: addr })),
                  ],
                },
              },
            },
            onSuccessUpdateEmail: {
              '#submission': {
                mailboxIds: sentMailboxIds,
                keywords: { $seen: true },
              },
            },
          },
          'submitEmail',
        ],
      ],
    };

    const response = await this.makeRequest(request);

    this.extractCreatedEmailId(this.getMethodResult(response, 0), 'email');

    const submissionResult = this.getMethodResult(response, 1);
    if (submissionResult.notCreated?.submission) {
      const err = submissionResult.notCreated.submission;
      throw new Error(
        `Failed to submit email: ${err.type}${err.description ? ` - ${err.description}` : ''}`,
      );
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
  }): Promise<string> {
    JmapClient.validateEmailAddresses(email.to);
    JmapClient.validateEmailAddresses(email.cc);
    JmapClient.validateEmailAddresses(email.bcc);

    const session = await this.getSession();

    if (!email.to?.length && !email.subject && !email.textBody && !email.htmlBody) {
      throw new Error('At least one of to, subject, textBody, or htmlBody must be provided');
    }

    const identities = await this.getIdentities();
    const selectedIdentity = this.resolveIdentity(identities, email.from);
    const fromEmail = selectedIdentity.email;

    const draftMailboxId = await this.resolveDraftMailboxId(email.mailboxId);
    const mailboxIds: Record<string, boolean> = { [draftMailboxId]: true };

    const emailObject: Record<string, unknown> = {
      mailboxIds,
      keywords: { $draft: true },
      from: [{ email: fromEmail }],
      ...(email.to?.length && { to: email.to.map((addr) => ({ email: addr })) }),
      ...(email.cc?.length && { cc: email.cc.map((addr) => ({ email: addr })) }),
      ...(email.bcc?.length && { bcc: email.bcc.map((addr) => ({ email: addr })) }),
      ...(email.subject && { subject: email.subject }),
      ...(email.inReplyTo?.length && { inReplyTo: email.inReplyTo }),
      ...(email.references?.length && { references: email.references }),
      ...this.buildBodyParts(email.textBody, email.htmlBody),
    };

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        [
          'Email/set',
          {
            accountId: session.accountId,
            create: { draft: emailObject },
          },
          'createDraft',
        ],
      ],
    };

    const response = await this.makeRequest(request);
    return this.extractCreatedEmailId(this.getMethodResult(response, 0), 'draft');
  }

  /**
   * Resolve the draft mailbox ID: use the provided one or look up the Drafts mailbox.
   */
  private async resolveDraftMailboxId(mailboxId?: string): Promise<string> {
    if (mailboxId) return mailboxId;

    const mailboxes = await this.getMailboxes();
    const draftsMailbox = this.findMailboxByRoleOrName(mailboxes, 'drafts', 'draft');
    if (!draftsMailbox) {
      throw new Error('Could not find Drafts mailbox');
    }
    return draftsMailbox.id;
  }

  /**
   * Fetch an existing draft email by ID, verifying it has the $draft keyword.
   */
  private async fetchExistingDraft(session: JmapSession, emailId: string): Promise<JmapEmail> {
    const getRequest: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        [
          'Email/get',
          {
            accountId: session.accountId,
            ids: [emailId],
            properties: [
              'id',
              'subject',
              'from',
              'to',
              'cc',
              'bcc',
              'textBody',
              'htmlBody',
              'bodyValues',
              'mailboxIds',
              'keywords',
            ],
            bodyProperties: ['partId', 'blobId', 'type', 'size'],
            fetchTextBodyValues: true,
            fetchHTMLBodyValues: true,
          },
          'getEmail',
        ],
      ],
    };

    const getResponse = await this.makeRequest(getRequest);
    const existing = this.getListResult(getResponse, 0)[0] as JmapEmail | undefined;
    if (!existing) {
      throw new Error(`Email with ID '${emailId}' not found`);
    }
    if (!existing.keywords?.$draft) {
      throw new Error('Cannot edit a non-draft email');
    }
    return existing;
  }

  /**
   * Resolve the identity for an update, considering the existing email's from address.
   */
  private resolveUpdateIdentity(
    identities: JmapIdentity[],
    fromOverride?: string,
    existingFrom?: string,
  ): JmapIdentity {
    if (fromOverride) {
      return this.resolveIdentity(identities, fromOverride);
    }
    if (existingFrom) {
      return (
        identities.find((id) => id.email.toLowerCase() === existingFrom.toLowerCase()) ||
        this.resolveIdentity(identities)
      );
    }
    return this.resolveIdentity(identities);
  }

  /**
   * Extract the first body value text from bodyValues using the corresponding body part list.
   */
  private extractBodyValue(
    bodyValues?: Record<string, { value: string; [key: string]: unknown }>,
    bodyParts?: { partId: string }[],
  ): string | undefined {
    if (!bodyValues || !bodyParts) return undefined;
    const firstPart = bodyParts[0];
    if (!firstPart) return undefined;
    // Try direct partId lookup first, then fall back to first value
    const entry = bodyValues[firstPart.partId] ?? Object.values(bodyValues)[0];
    return entry?.value;
  }

  async updateDraft(
    emailId: string,
    updates: {
      to?: string[];
      cc?: string[];
      bcc?: string[];
      subject?: string;
      textBody?: string;
      htmlBody?: string;
      from?: string;
    },
  ): Promise<string> {
    JmapClient.validateEmailAddresses(updates.to);
    JmapClient.validateEmailAddresses(updates.cc);
    JmapClient.validateEmailAddresses(updates.bcc);

    const session = await this.getSession();
    const existingEmail = await this.fetchExistingDraft(session, emailId);

    const identities = await this.getIdentities();
    const selectedIdentity = this.resolveUpdateIdentity(
      identities,
      updates.from,
      existingEmail.from?.[0]?.email,
    );

    const existingTextValue = this.extractBodyValue(
      existingEmail.bodyValues,
      existingEmail.textBody,
    );
    const existingHtmlValue = this.extractBodyValue(
      existingEmail.bodyValues,
      existingEmail.htmlBody,
    );

    const mergedSubject =
      updates.subject !== undefined ? updates.subject : existingEmail.subject || '';
    const mergedTo =
      updates.to !== undefined
        ? updates.to.map((addr) => ({ email: addr }))
        : existingEmail.to || [];
    const mergedCc =
      updates.cc !== undefined
        ? updates.cc.map((addr) => ({ email: addr }))
        : existingEmail.cc || [];
    const mergedBcc =
      updates.bcc !== undefined
        ? updates.bcc.map((addr) => ({ email: addr }))
        : existingEmail.bcc || [];

    const textBodyValue = updates.textBody !== undefined ? updates.textBody : existingTextValue;
    const htmlBodyValue = updates.htmlBody !== undefined ? updates.htmlBody : existingHtmlValue;

    const emailObject: Record<string, unknown> = {
      mailboxIds: existingEmail.mailboxIds,
      keywords: { $draft: true },
      from: [{ email: selectedIdentity.email }],
      to: mergedTo,
      cc: mergedCc,
      bcc: mergedBcc,
      subject: mergedSubject,
      ...this.buildBodyParts(textBodyValue, htmlBodyValue),
    };

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        [
          'Email/set',
          {
            accountId: session.accountId,
            create: { draft: emailObject },
            destroy: [emailId],
          },
          'updateDraft',
        ],
      ],
    };

    const response = await this.makeRequest(request);
    return this.extractCreatedEmailId(this.getMethodResult(response, 0), 'updated draft');
  }

  async sendDraft(emailId: string): Promise<string> {
    const session = await this.getSession();

    // Fetch the existing email to verify it's a draft
    const getRequest: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        [
          'Email/get',
          {
            accountId: session.accountId,
            ids: [emailId],
            properties: ['id', 'from', 'to', 'cc', 'bcc', 'keywords'],
          },
          'getEmail',
        ],
      ],
    };

    const getResponse = await this.makeRequest(getRequest);
    const email = this.getListResult(getResponse, 0)[0] as JmapEmail | undefined;
    if (!email) {
      throw new Error(`Email with ID '${emailId}' not found`);
    }

    if (!email.keywords?.$draft) {
      throw new Error('Cannot send a non-draft email');
    }

    // Collect all recipients for the envelope
    const toAddrs = (email.to || []) as { email: string }[];
    const ccAddrs = (email.cc || []) as { email: string }[];
    const bccAddrs = (email.bcc || []) as { email: string }[];
    const allRecipients: { email: string }[] = [...toAddrs, ...ccAddrs, ...bccAddrs];

    if (allRecipients.length === 0) {
      throw new Error('Draft has no recipients');
    }

    // Determine identity from the email's from field
    const fromEmail = email.from?.[0]?.email;
    if (!fromEmail) {
      throw new Error('Draft has no from address');
    }

    const identities = await this.getIdentities();
    const selectedIdentity = identities.find(
      (id) => id.email.toLowerCase() === fromEmail.toLowerCase(),
    );
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
      using: [
        'urn:ietf:params:jmap:core',
        'urn:ietf:params:jmap:mail',
        'urn:ietf:params:jmap:submission',
      ],
      methodCalls: [
        [
          'EmailSubmission/set',
          {
            accountId: session.accountId,
            create: {
              submission: {
                emailId,
                identityId: selectedIdentity.id,
                envelope: {
                  mailFrom: { email: fromEmail },
                  rcptTo: allRecipients.map((addr) => ({ email: addr.email })),
                },
              },
            },
            onSuccessUpdateEmail: {
              '#submission': {
                mailboxIds: sentMailboxIds,
                'keywords/$draft': null,
                'keywords/$seen': true,
              },
            },
          },
          'submitDraft',
        ],
      ],
    };

    const response = await this.makeRequest(request);
    const submissionResult = this.getMethodResult(response, 0);
    if (submissionResult.notCreated?.submission) {
      const err = submissionResult.notCreated.submission;
      throw new Error(
        `Failed to submit draft: ${err.type}${err.description ? ` - ${err.description}` : ''}`,
      );
    }

    const submissionId = submissionResult.created?.submission?.id;
    if (!submissionId) {
      throw new Error('Draft submission returned no submission ID');
    }

    return submissionId;
  }

  async getRecentEmails(limit = 10, mailboxName = 'inbox'): Promise<JmapEmail[]> {
    const session = await this.getSession();

    // Find the specified mailbox (default to inbox)
    const mailboxes = await this.getMailboxes();
    const targetMailbox = mailboxes.find(
      (mb) =>
        mb.role === mailboxName.toLowerCase() ||
        mb.name.toLowerCase().includes(mailboxName.toLowerCase()),
    );

    if (!targetMailbox) {
      throw new Error(`Could not find mailbox: ${mailboxName}`);
    }

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        [
          'Email/query',
          {
            accountId: session.accountId,
            filter: { inMailbox: targetMailbox.id },
            sort: [{ property: 'receivedAt', isAscending: false }],
            limit: Math.min(limit, 50),
          },
          'query',
        ],
        [
          'Email/get',
          {
            accountId: session.accountId,
            '#ids': { resultOf: 'query', name: 'Email/query', path: '/ids' },
            properties: [
              'id',
              'subject',
              'from',
              'to',
              'receivedAt',
              'preview',
              'hasAttachment',
              'keywords',
              'header:List-Unsubscribe:asURLs',
            ],
          },
          'emails',
        ],
      ],
    };

    const response = await this.makeRequest(request);
    return this.getListResult(response, 1) as JmapEmail[];
  }

  async markEmailRead(emailId: string, read = true): Promise<void> {
    const session = await this.getSession();

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        [
          'Email/set',
          {
            accountId: session.accountId,
            update: {
              [emailId]: {
                'keywords/$seen': read ? true : null,
              },
            },
          },
          'updateEmail',
        ],
      ],
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated?.[emailId]) {
      throw new Error(`Failed to mark email as ${read ? 'read' : 'unread'}.`);
    }
  }

  async pinEmail(emailId: string, pinned = true): Promise<void> {
    const session = await this.getSession();

    const update: Record<string, Record<string, boolean | null>> = {};
    update[emailId] = pinned ? { 'keywords/$flagged': true } : { 'keywords/$flagged': null };

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        [
          'Email/set',
          {
            accountId: session.accountId,
            update,
          },
          'pinEmail',
        ],
      ],
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated?.[emailId]) {
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
        [
          'Email/set',
          {
            accountId: session.accountId,
            update: {
              [emailId]: {
                mailboxIds: trashMailboxIds,
              },
            },
          },
          'moveToTrash',
        ],
      ],
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated?.[emailId]) {
      throw new Error('Failed to delete email.');
    }
  }

  async moveEmail(emailId: string, targetMailboxId: string): Promise<void> {
    const session = await this.getSession();

    // Fetch current mailboxIds to build a proper JMAP patch
    const getRequest: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        [
          'Email/get',
          {
            accountId: session.accountId,
            ids: [emailId],
            properties: ['mailboxIds'],
          },
          'getEmail',
        ],
      ],
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
        [
          'Email/set',
          {
            accountId: session.accountId,
            update: {
              [emailId]: patch,
            },
          },
          'moveEmail',
        ],
      ],
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated?.[emailId]) {
      throw new Error('Failed to move email.');
    }
  }

  async addLabels(emailId: string, mailboxIds: string[]): Promise<void> {
    const session = await this.getSession();

    // Build patch object to add specific mailboxIds
    const patch: Record<string, boolean | null> = {};
    for (const mbId of mailboxIds) {
      patch[`mailboxIds/${mbId}`] = true;
    }

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        [
          'Email/set',
          {
            accountId: session.accountId,
            update: {
              [emailId]: patch,
            },
          },
          'addLabels',
        ],
      ],
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated?.[emailId]) {
      throw new Error('Failed to add labels to email.');
    }
  }

  async removeLabels(emailId: string, mailboxIds: string[]): Promise<void> {
    const session = await this.getSession();

    // Build patch object to remove specific mailboxIds
    const patch: Record<string, boolean | null> = {};
    for (const mbId of mailboxIds) {
      patch[`mailboxIds/${mbId}`] = null;
    }

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        [
          'Email/set',
          {
            accountId: session.accountId,
            update: {
              [emailId]: patch,
            },
          },
          'removeLabels',
        ],
      ],
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated?.[emailId]) {
      throw new Error('Failed to remove labels from email.');
    }
  }

  async bulkAddLabels(emailIds: string[], mailboxIds: string[]): Promise<void> {
    const session = await this.getSession();

    // Build patch object to add specific mailboxIds
    const patch: Record<string, boolean | null> = {};
    for (const mbId of mailboxIds) {
      patch[`mailboxIds/${mbId}`] = true;
    }

    const updates: Record<string, Record<string, boolean | null>> = {};
    for (const id of emailIds) {
      updates[id] = patch;
    }

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        [
          'Email/set',
          {
            accountId: session.accountId,
            update: updates,
          },
          'bulkAddLabels',
        ],
      ],
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
    const patch: Record<string, boolean | null> = {};
    for (const mbId of mailboxIds) {
      patch[`mailboxIds/${mbId}`] = null;
    }

    const updates: Record<string, Record<string, boolean | null>> = {};
    for (const id of emailIds) {
      updates[id] = patch;
    }

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        [
          'Email/set',
          {
            accountId: session.accountId,
            update: updates,
          },
          'bulkRemoveLabels',
        ],
      ],
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated && Object.keys(result.notUpdated).length > 0) {
      throw new Error('Failed to remove labels from some emails.');
    }
  }

  async getEmailAttachments(emailId: string): Promise<JmapAttachment[]> {
    const session = await this.getSession();

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        [
          'Email/get',
          {
            accountId: session.accountId,
            ids: [emailId],
            properties: ['attachments'],
          },
          'getAttachments',
        ],
      ],
    };

    const response = await this.makeRequest(request);
    const email = this.getListResult(response, 0)[0] as JmapEmail | undefined;
    return email?.attachments || [];
  }

  async downloadAttachment(emailId: string, attachmentId: string): Promise<string> {
    const session = await this.getSession();

    // Get the email with full attachment details
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        [
          'Email/get',
          {
            accountId: session.accountId,
            ids: [emailId],
            properties: ['attachments', 'bodyValues'],
            bodyProperties: ['partId', 'blobId', 'size', 'name', 'type'],
          },
          'getEmail',
        ],
      ],
    };

    const response = await this.makeRequest(request);
    const email = this.getListResult(response, 0)[0] as JmapEmail | undefined;

    if (!email) {
      throw new Error('Email not found');
    }

    const attachments = email.attachments || [];

    // Find attachment by partId or by index
    let attachment: JmapAttachment | undefined = attachments.find(
      (att) => att.partId === attachmentId || att.blobId === attachmentId,
    );

    // If not found, try by array index
    if (!attachment) {
      const index = Number.parseInt(attachmentId, 10);
      if (!Number.isNaN(index)) {
        attachment = attachments[index];
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

  static validateSavePath(savePath: string): string {
    const allowedDir = JmapClient.DEFAULT_DOWNLOADS_DIR;
    const resolved = resolve(normalize(savePath));

    if (!resolved.startsWith(`${allowedDir}/`) && resolved !== allowedDir) {
      throw new Error(`Save path must be within ${allowedDir}. ` + `Received: ${savePath}`);
    }

    if (resolved.includes('\0')) {
      throw new Error('Save path contains null bytes');
    }

    return resolved;
  }

  async downloadAttachmentToFile(
    emailId: string,
    attachmentId: string,
    savePath: string,
  ): Promise<{ url: string; bytesWritten: number }> {
    const validatedPath = JmapClient.validateSavePath(savePath);
    const url = await this.downloadAttachment(emailId, attachmentId);

    const response = await fetch(url, {
      headers: { Authorization: this.auth.getAuthHeaders().Authorization },
    });

    if (!response.ok) {
      throw new Error(`Download failed: ${response.status} ${response.statusText}`);
    }

    const buffer = Buffer.from(await response.arrayBuffer());

    await mkdir(dirname(validatedPath), { recursive: true });
    await writeFile(validatedPath, buffer);

    return { url, bytesWritten: buffer.length };
  }

  /**
   * Normalize a date string to JMAP UTCDate format (RFC 3339).
   * Accepts "2015-01-01" or "2015-01-01T00:00:00Z" and ensures the result
   * always has a time component and Z suffix.
   */
  private static normalizeUtcDate(date: string): string {
    if (/^\d{4}-\d{2}-\d{2}$/.test(date)) {
      return `${date}T00:00:00Z`;
    }
    if (/^\d{4}-\d{2}-\d{2}T/.test(date) && !date.endsWith('Z') && !/[+-]\d{2}:\d{2}$/.test(date)) {
      return `${date}Z`;
    }
    return date;
  }

  /**
   * Apply keyword-based filters (isUnread, isPinned) to a JMAP email filter.
   * When both are set, wraps in an AND operator to avoid hasKeyword/notKeyword conflicts.
   */
  private applyKeywordFilters(
    filter: JmapEmailFilter,
    isUnread?: boolean,
    isPinned?: boolean,
  ): JmapEmailFilter {
    if (isUnread !== undefined && isPinned !== undefined) {
      const conditions: JmapEmailFilter[] = [filter];
      conditions.push(isUnread ? { notKeyword: '$seen' } : { hasKeyword: '$seen' });
      conditions.push(isPinned ? { hasKeyword: '$flagged' } : { notKeyword: '$flagged' });
      return { operator: 'AND', conditions };
    }

    if (isUnread === true) filter.notKeyword = '$seen';
    else if (isUnread === false) filter.hasKeyword = '$seen';
    if (isPinned === true) filter.hasKeyword = '$flagged';
    else if (isPinned === false) filter.notKeyword = '$flagged';

    return filter;
  }

  /**
   * Build a JMAP email filter from the advanced search parameters.
   */
  private buildSearchFilter(filters: {
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
  }): JmapEmailFilter {
    const filter: JmapEmailFilter = {};

    if (filters.query) filter.text = filters.query;
    if (filters.from) filter.from = filters.from;
    if (filters.to) filter.to = filters.to;
    if (filters.subject) filter.subject = filters.subject;
    if (filters.hasAttachment !== undefined) filter.hasAttachment = filters.hasAttachment;
    if (filters.mailboxId) filter.inMailbox = filters.mailboxId;
    if (filters.after) filter.after = JmapClient.normalizeUtcDate(filters.after);
    if (filters.before) filter.before = JmapClient.normalizeUtcDate(filters.before);

    return this.applyKeywordFilters(filter, filters.isUnread, filters.isPinned);
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
    sortAscending?: boolean;
    position?: number;
  }): Promise<JmapEmail[]> {
    const session = await this.getSession();
    const finalFilter = this.buildSearchFilter(filters);
    const position = filters.position || 0;

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        [
          'Email/query',
          {
            accountId: session.accountId,
            filter: finalFilter,
            sort: [{ property: 'receivedAt', isAscending: filters.sortAscending ?? false }],
            limit: Math.min(filters.limit || 50, 100),
            ...(position > 0 && { position }),
          },
          'query',
        ],
        [
          'Email/get',
          {
            accountId: session.accountId,
            '#ids': { resultOf: 'query', name: 'Email/query', path: '/ids' },
            properties: [
              'id',
              'subject',
              'from',
              'to',
              'cc',
              'receivedAt',
              'preview',
              'hasAttachment',
              'keywords',
              'threadId',
            ],
          },
          'emails',
        ],
      ],
    };

    const response = await this.makeRequest(request);
    // Check Email/query result first to surface real errors instead of cascading invalidResultReference
    this.getMethodResult(response, 0);
    return this.getListResult(response, 1) as JmapEmail[];
  }

  async searchEmails(
    query: string,
    limit = 20,
    sortAscending = false,
    position = 0,
  ): Promise<JmapEmail[]> {
    const session = await this.getSession();

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        [
          'Email/query',
          {
            accountId: session.accountId,
            filter: { text: query },
            sort: [{ property: 'receivedAt', isAscending: sortAscending }],
            limit,
            ...(position > 0 && { position }),
          },
          'query',
        ],
        [
          'Email/get',
          {
            accountId: session.accountId,
            '#ids': { resultOf: 'query', name: 'Email/query', path: '/ids' },
            properties: ['id', 'subject', 'from', 'to', 'receivedAt', 'preview', 'hasAttachment'],
          },
          'emails',
        ],
      ],
    };

    const response = await this.makeRequest(request);
    return this.getListResult(response, 1) as JmapEmail[];
  }

  async getThread(threadId: string): Promise<JmapEmail[]> {
    const session = await this.getSession();

    // First, check if threadId is actually an email ID and resolve the thread
    let actualThreadId = threadId;

    // Try to get the email first to see if we need to resolve thread ID
    try {
      const emailRequest: JmapRequest = {
        using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
        methodCalls: [
          [
            'Email/get',
            {
              accountId: session.accountId,
              ids: [threadId],
              properties: ['threadId'],
            },
            'checkEmail',
          ],
        ],
      };

      const emailResponse = await this.makeRequest(emailRequest);
      const email = this.getListResult(emailResponse, 0)[0];

      if (email?.threadId && typeof email.threadId === 'string') {
        actualThreadId = email.threadId;
      }
    } catch (error) {
      // If email lookup fails, assume threadId is correct
    }

    // Use Thread/get with the resolved thread ID
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        [
          'Thread/get',
          {
            accountId: session.accountId,
            ids: [actualThreadId],
          },
          'getThread',
        ],
        [
          'Email/get',
          {
            accountId: session.accountId,
            '#ids': { resultOf: 'getThread', name: 'Thread/get', path: '/list/*/emailIds' },
            properties: [
              'id',
              'subject',
              'from',
              'to',
              'cc',
              'receivedAt',
              'preview',
              'hasAttachment',
              'keywords',
              'threadId',
            ],
          },
          'emails',
        ],
      ],
    };

    const response = await this.makeRequest(request);
    const threadResult = this.getMethodResult(response, 0);

    // Check if thread was found
    if (threadResult.notFound?.includes(actualThreadId)) {
      throw new Error(`Thread with ID '${actualThreadId}' not found`);
    }

    return this.getListResult(response, 1) as JmapEmail[];
  }

  async getMailboxStats(
    mailboxId?: string,
  ): Promise<Record<string, unknown> | Record<string, unknown>[]> {
    const session = await this.getSession();

    if (mailboxId) {
      // Get stats for specific mailbox
      const request: JmapRequest = {
        using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
        methodCalls: [
          [
            'Mailbox/get',
            {
              accountId: session.accountId,
              ids: [mailboxId],
              properties: [
                'id',
                'name',
                'role',
                'totalEmails',
                'unreadEmails',
                'totalThreads',
                'unreadThreads',
              ],
            },
            'mailbox',
          ],
        ],
      };

      const response = await this.makeRequest(request);
      return this.getListResult(response, 0)[0];
    }
    // Get stats for all mailboxes
    const mailboxes = await this.getMailboxes();
    return mailboxes.map((mb) => ({
      id: mb.id,
      name: mb.name,
      role: mb.role,
      totalEmails: mb.totalEmails || 0,
      unreadEmails: mb.unreadEmails || 0,
      totalThreads: mb.totalThreads || 0,
      unreadThreads: mb.unreadThreads || 0,
    }));
  }

  async getAccountSummary(): Promise<Record<string, unknown>> {
    const session = await this.getSession();
    const mailboxes = await this.getMailboxes();

    let identityCount: number | string;
    try {
      const identities = await this.getIdentities();
      identityCount = identities.length;
    } catch {
      identityCount = 'unavailable (submission scope may not be enabled)';
    }

    // Calculate totals
    const totals = mailboxes.reduce(
      (acc, mb) => ({
        totalEmails: acc.totalEmails + (mb.totalEmails || 0),
        unreadEmails: acc.unreadEmails + (mb.unreadEmails || 0),
        totalThreads: acc.totalThreads + (mb.totalThreads || 0),
        unreadThreads: acc.unreadThreads + (mb.unreadThreads || 0),
      }),
      { totalEmails: 0, unreadEmails: 0, totalThreads: 0, unreadThreads: 0 },
    );

    return {
      accountId: session.accountId,
      mailboxCount: mailboxes.length,
      identityCount,
      ...totals,
      mailboxes: mailboxes.map((mb) => ({
        id: mb.id,
        name: mb.name,
        role: mb.role,
        totalEmails: mb.totalEmails || 0,
        unreadEmails: mb.unreadEmails || 0,
      })),
    };
  }

  async bulkMarkRead(emailIds: string[], read = true): Promise<void> {
    const session = await this.getSession();

    const updates: Record<string, Record<string, unknown>> = {};

    for (const id of emailIds) {
      updates[id] = { 'keywords/$seen': read ? true : null };
    }

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        [
          'Email/set',
          {
            accountId: session.accountId,
            update: updates,
          },
          'bulkUpdate',
        ],
      ],
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated && Object.keys(result.notUpdated).length > 0) {
      throw new Error('Failed to update some emails.');
    }
  }

  async bulkPinEmails(emailIds: string[], pinned = true): Promise<void> {
    const session = await this.getSession();

    const updates: Record<string, Record<string, boolean | null>> = {};
    for (const id of emailIds) {
      updates[id] = pinned ? { 'keywords/$flagged': true } : { 'keywords/$flagged': null };
    }

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        [
          'Email/set',
          {
            accountId: session.accountId,
            update: updates,
          },
          'bulkFlag',
        ],
      ],
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
        [
          'Email/get',
          {
            accountId: session.accountId,
            ids: emailIds,
            properties: ['id', 'mailboxIds'],
          },
          'getEmails',
        ],
      ],
    };
    const getResponse = await this.makeRequest(getRequest);
    const emails = this.getListResult(getResponse, 0);
    const mailboxMap: Record<string, Record<string, boolean>> = {};
    for (const e of emails) {
      const id = e.id as string;
      mailboxMap[id] = (e.mailboxIds as Record<string, boolean>) || {};
    }

    // Build patch per email: remove all current mailboxes, add target
    const updates: Record<string, Record<string, boolean | null>> = {};
    for (const id of emailIds) {
      const patch: Record<string, boolean | null> = {};
      for (const mbId of Object.keys(mailboxMap[id] || {})) {
        patch[`mailboxIds/${mbId}`] = null;
      }
      patch[`mailboxIds/${targetMailboxId}`] = true;
      updates[id] = patch;
    }

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        [
          'Email/set',
          {
            accountId: session.accountId,
            update: updates,
          },
          'bulkMove',
        ],
      ],
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

    const updates: Record<string, Record<string, Record<string, boolean>>> = {};
    for (const id of emailIds) {
      updates[id] = { mailboxIds: trashMailboxIds };
    }

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        [
          'Email/set',
          {
            accountId: session.accountId,
            update: updates,
          },
          'bulkDelete',
        ],
      ],
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated && Object.keys(result.notUpdated).length > 0) {
      throw new Error('Failed to delete some emails.');
    }
  }
}
