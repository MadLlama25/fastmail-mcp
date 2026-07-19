import { FastmailAuth } from './auth.js';
import { validateFastmailUrl } from './url-validation.js';
import { writeFile, mkdir, realpath, stat, lstat, unlink, readFile } from 'fs/promises';
import { dirname, resolve, normalize, sep, basename, join } from 'path';
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
    // Require a single well-formed addr-spec before accepting the wildcard —
    // otherwise a composite/spoofed value like "a@evil.com,b@example.com" or one
    // carrying CR/LF would pass the client-side "verified identity" guard.
    if (!/^[^\s@,;"]+@[^\s@,;"]+$/.test(addr)) return false;
    return addr.endsWith(domain);
  }
  return false;
}

export interface EmailQueryFilters {
  query?: string;
  from?: string;
  to?: string;
  subject?: string;
  hasAttachment?: boolean;
  isUnread?: boolean;
  isPinned?: boolean;
  mailboxId?: string;
  requiredMailboxIds?: string[];
  excludeMailboxIds?: string[];
  after?: string;
  before?: string;
  limit?: number;
  ascending?: boolean;
}

/**
 * Build a JMAP Email/query filter from the high-level filter parameters used by
 * advancedSearch / advancedSearchMetadata.
 *
 * Output is one of three shapes:
 *  - {} when no fields are set (matches all messages)
 *  - a single FilterCondition object when every field can coexist on one condition
 *  - { operator: 'AND', conditions: [...] } when JMAP's per-FilterCondition limits
 *    require splitting across multiple conditions (multi-mailbox AND, or
 *    hasKeyword/notKeyword conflicts between isUnread and isPinned)
 *
 * Exported (rather than file-private) so the test suite can verify the filter
 * shape directly without round-tripping through a mocked makeRequest.
 */
export function buildEmailQueryFilter(filters: EmailQueryFilters): any {
  // Base FilterCondition: fields that always coexist cleanly on a single condition.
  const base: any = {};
  if (filters.query) base.text = filters.query;
  if (filters.from) base.from = filters.from;
  if (filters.to) base.to = filters.to;
  if (filters.subject) base.subject = filters.subject;
  if (filters.hasAttachment !== undefined) base.hasAttachment = filters.hasAttachment;
  if (filters.after) base.after = filters.after;
  if (filters.before) base.before = filters.before;
  if (filters.excludeMailboxIds && filters.excludeMailboxIds.length > 0) {
    base.inMailboxOtherThan = filters.excludeMailboxIds;
  }

  // Combine mailboxId + requiredMailboxIds, de-dup, preserve insertion order.
  const seenIds = new Set<string>();
  const requiredMailboxes: string[] = [];
  const candidateIds = [filters.mailboxId, ...(filters.requiredMailboxIds ?? [])];
  for (const id of candidateIds) {
    if (id && !seenIds.has(id)) {
      seenIds.add(id);
      requiredMailboxes.push(id);
    }
  }

  // A single required mailbox folds into the base condition. Multiple
  // memberships have to live in separate FilterConditions because JMAP's
  // inMailbox is singular per condition.
  if (requiredMailboxes.length === 1) {
    base.inMailbox = requiredMailboxes[0];
  }

  // When both isUnread and isPinned are defined, their keyword conditions
  // can collide on hasKeyword or notKeyword (each is singular per
  // FilterCondition), so split them across separate conditions.
  const splitKeywords = filters.isUnread !== undefined && filters.isPinned !== undefined;
  if (!splitKeywords) {
    if (filters.isUnread === true) base.notKeyword = '$seen';
    else if (filters.isUnread === false) base.hasKeyword = '$seen';
    if (filters.isPinned === true) base.hasKeyword = '$flagged';
    else if (filters.isPinned === false) base.notKeyword = '$flagged';
  }

  // Assemble the final shape.
  const conditions: any[] = [];
  if (Object.keys(base).length > 0) conditions.push(base);

  if (requiredMailboxes.length > 1) {
    for (const id of requiredMailboxes) conditions.push({ inMailbox: id });
  }

  if (splitKeywords) {
    conditions.push(filters.isUnread ? { notKeyword: '$seen' } : { hasKeyword: '$seen' });
    conditions.push(filters.isPinned ? { hasKeyword: '$flagged' } : { notKeyword: '$flagged' });
  }

  if (conditions.length === 0) return {};
  if (conditions.length === 1) return conditions[0];
  return { operator: 'AND', conditions };
}

export interface QueryResult<T = any> {
  items: T[];
  total?: number;
}

/**
 * Tool-facing attachment source. Exactly one of localPath, emailId(+attachmentId),
 * or blobId must be set; name/type optionally override the inferred values.
 */
export interface AttachmentInput {
  localPath?: string;
  emailId?: string;
  attachmentId?: string;
  blobId?: string;
  name?: string;
  type?: string;
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

  /**
   * Build a QueryResult from a query + get pair.
   * queryIndex is the /query response; listIndex is the /get response.
   */
  protected getQueryResult(response: JmapResponse, queryIndex: number, listIndex: number): QueryResult {
    const queryResult = this.getMethodResult(response, queryIndex);
    const items = this.getListResult(response, listIndex);
    const total = queryResult?.total;
    return total != null ? { items, total } : { items };
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

    // Validate every URL the server hands us before we send the bearer token to it.
    // The downloadUrl/uploadUrl are URL templates with {accountId}/{blobId}/etc.
    // placeholders, so we strip those for parsing and validate origin only.
    const allowUnsafe = this.auth.getAllowUnsafe();
    const stripTemplate = (url: string) => url.replace(/\{[^}]+\}/g, 'x');
    if (typeof sessionData.apiUrl !== 'string') {
      throw new Error('Invalid session response: apiUrl missing');
    }
    validateFastmailUrl(sessionData.apiUrl, 'session.apiUrl', allowUnsafe);
    // Reject non-string download/upload URLs rather than storing them unvalidated
    // (validate/store must not diverge — a later coercion would otherwise inherit
    // an unchecked value).
    if (sessionData.downloadUrl !== undefined) {
      if (typeof sessionData.downloadUrl !== 'string') {
        throw new Error('Invalid session response: downloadUrl is not a string');
      }
      validateFastmailUrl(stripTemplate(sessionData.downloadUrl), 'session.downloadUrl', allowUnsafe);
    }
    if (sessionData.uploadUrl !== undefined) {
      if (typeof sessionData.uploadUrl !== 'string') {
        throw new Error('Invalid session response: uploadUrl is not a string');
      }
      validateFastmailUrl(stripTemplate(sessionData.uploadUrl), 'session.uploadUrl', allowUnsafe);
    }

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

  async getMailboxes(options?: { properties?: string[]; parentId?: string | null }): Promise<any[]> {
    const session = await this.getSession();

    const args: Record<string, any> = { accountId: session.accountId };
    if (options?.properties && options.properties.length > 0) {
      args.properties = options.properties;
    }

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Mailbox/get', args, 'mailboxes']
      ]
    };

    const response = await this.makeRequest(request);
    let list = this.getListResult(response, 0);
    if (options && Object.prototype.hasOwnProperty.call(options, 'parentId')) {
      const filterParent = options.parentId ?? null;
      list = list.filter((mb: any) => (mb.parentId ?? null) === filterParent);
    }
    return list;
  }

  async getMailboxByName(path: string): Promise<{ id: string; name: string; parentId: string | null; path: string }> {
    if (!path || typeof path !== 'string') {
      throw new Error('path is required and must be a non-empty string');
    }
    const mailboxes = await this.getMailboxes({ properties: ['id', 'name', 'parentId'] });
    const byId = new Map<string, any>();
    for (const mb of mailboxes) byId.set(mb.id, mb);

    const buildPath = (mb: any): string => {
      const segments: string[] = [];
      let cursor: any = mb;
      let depth = 0;
      while (cursor && depth < 100) {
        segments.unshift(cursor.name);
        cursor = cursor.parentId ? byId.get(cursor.parentId) : null;
        depth++;
      }
      return segments.join('/');
    };

    for (const mb of mailboxes) {
      if (buildPath(mb) === path) {
        return { id: mb.id, name: mb.name, parentId: mb.parentId ?? null, path };
      }
    }
    throw new Error(`Mailbox not found: ${path}`);
  }

  async createMailbox(name: string, parentId?: string | null): Promise<string> {
    if (!name || typeof name !== 'string') {
      throw new Error('name is required and must be a non-empty string');
    }
    const session = await this.getSession();

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Mailbox/set', {
          accountId: session.accountId,
          create: {
            new1: {
              name,
              parentId: parentId ?? null
            }
          }
        }, 'createMailbox']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notCreated && result.notCreated.new1) {
      const err = result.notCreated.new1;
      const detail = err.description ? ` - ${err.description}` : '';
      const props = err.properties ? ` (properties: ${err.properties.join(', ')})` : '';
      throw new Error(`Failed to create mailbox: ${err.type}${detail}${props}`);
    }

    const created = result.created?.new1;
    if (!created?.id) {
      throw new Error('Mailbox creation reported success but server did not return an ID');
    }
    return created.id;
  }

  async getEmails(mailboxId?: string, limit: number = 20, ascending: boolean = false): Promise<QueryResult> {
    const session = await this.getSession();

    const filter = mailboxId ? { inMailbox: mailboxId } : {};

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/query', {
          accountId: session.accountId,
          filter,
          sort: [{ property: 'receivedAt', isAscending: ascending }],
          limit,
          calculateTotal: true
        }, 'query'],
        ['Email/get', {
          accountId: session.accountId,
          '#ids': { resultOf: 'query', name: 'Email/query', path: '/ids' },
          properties: ['id', 'subject', 'from', 'to', 'replyTo', 'receivedAt', 'preview', 'hasAttachment']
        }, 'emails']
      ]
    };

    const response = await this.makeRequest(request);
    return this.getQueryResult(response, 0, 1);
  }

  async getEmailsMetadata(mailboxId?: string, limit: number = 20, ascending: boolean = false): Promise<QueryResult> {
    const session = await this.getSession();

    const filter = mailboxId ? { inMailbox: mailboxId } : {};

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/query', {
          accountId: session.accountId,
          filter,
          sort: [{ property: 'receivedAt', isAscending: ascending }],
          limit,
          calculateTotal: true
        }, 'query'],
        ['Email/get', {
          accountId: session.accountId,
          '#ids': { resultOf: 'query', name: 'Email/query', path: '/ids' },
          properties: ['id', 'threadId', 'subject', 'from', 'to', 'replyTo', 'receivedAt', 'hasAttachment', 'keywords']
        }, 'emails']
      ]
    };

    const response = await this.makeRequest(request);
    return this.getQueryResult(response, 0, 1);
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

  async getEmailMetadata(id: string): Promise<any> {
    const session = await this.getSession();

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/get', {
          accountId: session.accountId,
          ids: [id],
          properties: [
            'id',
            'threadId',
            'mailboxIds',
            'keywords',
            'receivedAt',
            'sentAt',
            'subject',
            'from',
            'to',
            'cc',
            'bcc',
            'replyTo',
            'messageId',
            'inReplyTo',
            'references',
            'size',
            'hasAttachment',
          ],
        }, 'emailMetadata']
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
    attachments?: AttachmentInput[];
    downloadDir?: string;
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

    // Resolve attachment sources (upload local files, reuse existing blobs)
    // before creating the email so a failed upload never leaves a draft behind.
    const resolvedAttachments = email.attachments?.length
      ? await this.resolveAttachments(email.attachments, email.downloadDir)
      : undefined;

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
      },
      // RFC 8621 §4.6 convenience property: the server assembles the final
      // multipart bodyStructure from textBody/htmlBody/attachments itself.
      // Never hand-build bodyStructure alongside these — mixing is forbidden.
      ...(resolvedAttachments?.length && { attachments: resolvedAttachments }),
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
    attachments?: AttachmentInput[];
    downloadDir?: string;
  }): Promise<string> {
    const session = await this.getSession();

    // Validate at least one meaningful field is present
    if (!email.to?.length && !email.subject && !email.textBody && !email.htmlBody && !email.attachments?.length) {
      throw new Error('At least one of to, subject, textBody, htmlBody, or attachments must be provided');
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
      from: [{ name: selectedIdentity.name, email: fromEmail }],
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
    if (email.attachments?.length) {
      // RFC 8621 convenience property — server assembles bodyStructure.
      emailObject.attachments = await this.resolveAttachments(email.attachments, email.downloadDir);
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

    // Propagate server-provided error details from notCreated
    if (result.notCreated?.draft) {
      const err = result.notCreated.draft;
      throw new Error(`Failed to create draft: ${err.type}${err.description ? ' - ' + err.description : ''}`);
    }

    // Throw if created ID is missing instead of returning silently
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
    attachments?: AttachmentInput[];
    downloadDir?: string;
  }): Promise<string> {
    const session = await this.getSession();

    // Fetch the existing email
    const getRequest: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/get', {
          accountId: session.accountId,
          ids: [emailId],
          properties: ['id', 'subject', 'from', 'to', 'cc', 'bcc', 'replyTo', 'textBody', 'htmlBody', 'bodyValues', 'mailboxIds', 'keywords', 'attachments'],
          bodyProperties: ['partId', 'blobId', 'type', 'size', 'name', 'disposition', 'cid'],
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

    // Extract existing body values by MIME type, keyed into bodyValues by partId.
    //
    // The previous lookup used `tb.partId === bv.partId || true` — always true, since the
    // bodyValue objects from Object.values() have no partId field (partId is the map key).
    // So both bodies collapsed to Object.values(bodyValues)[0], and a subject-only or
    // single-body edit would overwrite or lose the other format (and clients render the
    // HTML alternative, RFC 2046 5.1.4, so recipients saw the wrong content).
    //
    // Note also: when a draft has only one body, the server aliases that one part into
    // BOTH the textBody and htmlBody lists (e.g. a text-only draft lists the text/plain
    // part under htmlBody too, with type "text/plain"). So we must select by the part's
    // actual MIME type — not mere presence in a list — or we'd read the text value into
    // the html slot and synthesise a phantom text/html part on recreate.
    const bodyValues = existingEmail.bodyValues || {};
    const bodyValueForType = (parts: any[] | undefined, mimeType: string): string | undefined => {
      const part = parts?.find((p: any) => p.type === mimeType && p.partId != null && bodyValues[p.partId]);
      return part ? bodyValues[part.partId].value : undefined;
    };
    const existingTextValue = bodyValueForType(existingEmail.textBody, 'text/plain');
    const existingHtmlValue = bodyValueForType(existingEmail.htmlBody, 'text/html');

    // Merge: updates override existing values
    const mergedSubject = updates.subject !== undefined ? updates.subject : (existingEmail.subject || '');
    const mergedTo = updates.to !== undefined ? updates.to.map(addr => ({ email: addr })) : (existingEmail.to || []);
    const mergedCc = updates.cc !== undefined ? updates.cc.map(addr => ({ email: addr })) : (existingEmail.cc || []);
    const mergedBcc = updates.bcc !== undefined ? updates.bcc.map(addr => ({ email: addr })) : (existingEmail.bcc || []);
    const mergedReplyTo = updates.replyTo !== undefined ? updates.replyTo.map(addr => ({ email: addr })) : (existingEmail.replyTo || null);

    const textBodyValue = updates.textBody !== undefined ? updates.textBody : existingTextValue;
    const htmlBodyValue = updates.htmlBody !== undefined ? updates.htmlBody : existingHtmlValue;

    const emailObject: any = {
      mailboxIds: existingEmail.mailboxIds,
      keywords: { $draft: true },
      from: [{ name: selectedIdentity.name, email: updates.from || existingEmail.from?.[0]?.email || selectedIdentity.email }],
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

    // Carry existing attachments into the recreated draft — the recreate used
    // to silently strip them. Same-account blobIds are directly reusable, so
    // this is metadata-only. Newly supplied attachments are appended.
    const carriedAttachments = (existingEmail.attachments ?? []).map((att: any) => ({
      blobId: att.blobId,
      type: att.type || 'application/octet-stream',
      name: att.name || 'attachment',
      disposition: att.disposition || 'attachment',
      ...(att.cid && { cid: att.cid }),
    }));
    const newAttachments = updates.attachments?.length
      ? await this.resolveAttachments(updates.attachments, updates.downloadDir)
      : [];
    if (carriedAttachments.length || newAttachments.length) {
      emailObject.attachments = [...carriedAttachments, ...newAttachments];
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

  async getRecentEmails(limit: number = 10, mailboxName: string | null = null, ascending: boolean = false): Promise<QueryResult> {
    const session = await this.getSession();

    const mailboxes = await this.getMailboxes();

    let filter: any;
    if (mailboxName) {
      const targetMailbox = mailboxes.find(mb =>
        mb.role === mailboxName.toLowerCase() ||
        mb.name.toLowerCase().includes(mailboxName.toLowerCase())
      );

      if (!targetMailbox) {
        throw new Error(`Could not find mailbox: ${mailboxName}`);
      }
      filter = { inMailbox: targetMailbox.id };
    } else {
      // No mailbox given: span all folders (Sent, custom folders, ...) but
      // keep Trash and Spam out of "recent emails".
      const excludedIds = mailboxes
        .filter(mb => mb.role === 'trash' || mb.role === 'junk' || mb.role === 'spam')
        .map(mb => mb.id);
      filter = excludedIds.length > 0 ? { inMailboxOtherThan: excludedIds } : {};
    }

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/query', {
          accountId: session.accountId,
          filter,
          sort: [{ property: 'receivedAt', isAscending: ascending }],
          limit: Math.min(limit, 50),
          calculateTotal: true
        }, 'query'],
        ['Email/get', {
          accountId: session.accountId,
          '#ids': { resultOf: 'query', name: 'Email/query', path: '/ids' },
          properties: ['id', 'subject', 'from', 'to', 'replyTo', 'receivedAt', 'preview', 'hasAttachment', 'keywords', 'header:List-Unsubscribe:asURLs']
        }, 'emails']
      ]
    };

    const response = await this.makeRequest(request);
    return this.getQueryResult(response, 0, 1);
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
      const err = result.notUpdated[emailId];
      const detail = err.description ? ` - ${err.description}` : '';
      throw new Error(`Failed to ${pinned ? 'pin' : 'unpin'} email: ${err.type}${detail}`);
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

  async archiveEmail(emailId: string, targetMailboxId: string): Promise<void> {
    const session = await this.getSession();

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

    if (!email) {
      throw new Error(`Email not found: ${emailId}`);
    }

    const patch: Record<string, boolean | null> = {};
    if (email.mailboxIds) {
      for (const mbId of Object.keys(email.mailboxIds)) {
        patch[`mailboxIds/${mbId}`] = null;
      }
    }
    patch[`mailboxIds/${targetMailboxId}`] = true;
    patch['keywords/$seen'] = true;

    const setRequest: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update: {
            [emailId]: patch
          }
        }, 'archiveEmail']
      ]
    };

    const response = await this.makeRequest(setRequest);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated && result.notUpdated[emailId]) {
      const err = result.notUpdated[emailId];
      const detail = err.description ? ` - ${err.description}` : '';
      throw new Error(`Failed to archive email: ${err.type}${detail}`);
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
      const err = result.notUpdated[emailId];
      const detail = err.description ? ` - ${err.description}` : '';
      throw new Error(`Failed to add labels to email: ${err.type}${detail}`);
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
      const err = result.notUpdated[emailId];
      const detail = err.description ? ` - ${err.description}` : '';
      throw new Error(`Failed to remove labels from email: ${err.type}${detail}`);
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

  /**
   * Resolve an attachment reference (partId, blobId, or array index) on an
   * email to its metadata. Shared by the download path and by attach-on-send's
   * zero-copy "reuse an existing attachment" source — same-account blobIds are
   * directly valid in Email/set, so no bytes ever need to move for a forward.
   */
  async getAttachmentInfo(emailId: string, attachmentId: string): Promise<{ blobId: string; type: string; name: string; size?: number }> {
    const session = await this.getSession();

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

    return {
      blobId: attachment.blobId,
      type: attachment.type || 'application/octet-stream',
      name: attachment.name || 'attachment',
      size: attachment.size,
    };
  }

  async downloadAttachment(emailId: string, attachmentId: string): Promise<string> {
    const session = await this.getSession();
    const attachment = await this.getAttachmentInfo(emailId, attachmentId);

    // Get the download URL from session
    const downloadUrl = session.downloadUrl;
    if (!downloadUrl) {
      throw new Error('Download capability not available in session');
    }

    // Build download URL
    const url = downloadUrl
      .replace('{accountId}', session.accountId)
      .replace('{blobId}', attachment.blobId)
      .replace('{type}', encodeURIComponent(attachment.type))
      .replace('{name}', encodeURIComponent(attachment.name));

    // Re-validate the substituted URL before it's used to send the bearer token —
    // defends against a template whose placeholders were filled with values that
    // rewrote the origin (belt-and-suspenders over the session-time origin check).
    validateFastmailUrl(url, 'downloadUrl', this.auth.getAllowUnsafe());

    return url;
  }

  static readonly DEFAULT_DOWNLOADS_DIR = resolve(homedir(), 'Downloads', 'fastmail-mcp');

  static validateSavePath(savePath: string, downloadDir?: string): string {
    const allowedDir = downloadDir ? resolve(normalize(downloadDir)) : JmapClient.DEFAULT_DOWNLOADS_DIR;
    // Resolve relative paths against the allowed download directory rather than
    // the process cwd (which is unpredictable for an MCP server launched by a
    // client). Absolute paths are taken as-is; either way the containment check
    // below is the security boundary. So a bare filename lands safely in the
    // configured dir in one step, and an absolute path inside that dir writes
    // exactly there.
    const resolved = resolve(allowedDir, normalize(savePath));

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

  /**
   * Symlink-safe canonicalization of a save path. Walks up to the longest
   * existing ancestor, realpaths it, and verifies it lives under the canonical
   * allowed directory. Refuses to overwrite an existing symlink at the target.
   *
   * Returns the canonical path that is safe to write to. Throws on escape.
   */
  static async safeWritePath(savePath: string, downloadDir?: string): Promise<string> {
    // Lexical pre-check first (cheap and gives nice errors)
    const lexical = JmapClient.validateSavePath(savePath, downloadDir);
    const allowedDir = downloadDir ? resolve(normalize(downloadDir)) : JmapClient.DEFAULT_DOWNLOADS_DIR;

    // Ensure allowed dir exists so realpath can resolve it.
    await mkdir(allowedDir, { recursive: true });
    const canonicalAllowed = await realpath(allowedDir);

    // Walk up from the target until we find an existing ancestor.
    let ancestor = dirname(lexical);
    const missingSegments: string[] = [];
    while (true) {
      try {
        await stat(ancestor);
        break;
      } catch (e: any) {
        if (e.code !== 'ENOENT') throw e;
        missingSegments.unshift(basename(ancestor));
        const parent = dirname(ancestor);
        if (parent === ancestor) {
          throw new Error(`Could not find existing ancestor for save path: ${lexical}`);
        }
        ancestor = parent;
      }
    }

    // Canonicalize the existing ancestor — this is what catches symlink escapes.
    const canonicalAncestor = await realpath(ancestor);
    if (canonicalAncestor !== canonicalAllowed && !canonicalAncestor.startsWith(canonicalAllowed + sep)) {
      throw new Error(
        `Save path resolves to '${canonicalAncestor}' which is outside the allowed directory '${canonicalAllowed}'. ` +
        `Refusing to follow symlink escape.`,
      );
    }

    // Reconstruct the safe canonical path under the canonical ancestor.
    const safePath = join(canonicalAncestor, ...missingSegments, basename(lexical));

    // If a symlink already exists at the target, refuse — writing through it
    // would still escape the allowed directory.
    try {
      const lst = await lstat(safePath);
      if (lst.isSymbolicLink()) {
        throw new Error(`Refusing to overwrite an existing symlink at the target: ${safePath}`);
      }
    } catch (e: any) {
      if (e.code !== 'ENOENT') throw e;
    }

    return safePath;
  }

  /**
   * Fetch an attachment's bytes into memory. Shared by the local-file download
   * path and any consumer that needs the raw bytes (e.g. re-upload targets).
   */
  async fetchAttachmentBuffer(emailId: string, attachmentId: string): Promise<{ buffer: Buffer; url: string; blobId: string; type: string; name: string }> {
    const info = await this.getAttachmentInfo(emailId, attachmentId);
    const url = await this.downloadAttachment(emailId, attachmentId);

    const response = await fetch(url, {
      headers: { 'Authorization': this.auth.getAuthHeaders()['Authorization'] },
      // Never follow a redirect on a token-bearing request — a validated host
      // that 3xx-redirects cross-origin would otherwise have the attachment body
      // silently sourced from a non-allowlisted host.
      redirect: 'error',
    });

    if (!response.ok) {
      throw new Error(`Download failed: ${response.status} ${response.statusText}`);
    }

    return { buffer: Buffer.from(await response.arrayBuffer()), url, ...info };
  }

  async downloadAttachmentToFile(emailId: string, attachmentId: string, savePath: string, downloadDir?: string): Promise<{ url: string; bytesWritten: number; savedPath: string }> {
    // First check + create the parent dir before the (slow) network fetch, to
    // shrink the window in which a co-resident process could swap a checked
    // directory for a symlink.
    await JmapClient.safeWritePath(savePath, downloadDir);
    const { buffer, url } = await this.fetchAttachmentBuffer(emailId, attachmentId);

    // Re-validate immediately before writing (TOCTOU: the target may have been
    // swapped for a symlink during the fetch), then write with O_EXCL so we
    // never follow a symlink planted at the path. Overwriting a pre-existing
    // regular file stays supported: on EEXIST we re-run the symlink-safe check
    // (which refuses a symlink) and replace the plain file.
    const safePath = await JmapClient.safeWritePath(savePath, downloadDir);
    await mkdir(dirname(safePath), { recursive: true });
    try {
      await writeFile(safePath, buffer, { flag: 'wx' });
    } catch (e: any) {
      if (e?.code !== 'EEXIST') throw e;
      await JmapClient.safeWritePath(savePath, downloadDir); // refuses a symlink at target
      await unlink(safePath);
      await writeFile(safePath, buffer, { flag: 'wx' });
    }

    return { url, bytesWritten: buffer.length, savedPath: safePath };
  }

  /**
   * Read-side counterpart of validateSavePath/safeWritePath: confine local
   * attachment sources to the download directory. Lexical containment first,
   * then realpath the EXISTING file and require the canonical location to stay
   * under the canonical allowed dir — a symlink inside the dir pointing out
   * (e.g. ~/Downloads/fastmail-mcp/link -> /etc/passwd) must not be readable.
   */
  static async validateReadPath(readPath: string, downloadDir?: string): Promise<string> {
    const lexical = JmapClient.validateSavePath(readPath, downloadDir);
    const allowedDir = downloadDir ? resolve(normalize(downloadDir)) : JmapClient.DEFAULT_DOWNLOADS_DIR;
    let canonicalAllowed: string;
    let canonicalFile: string;
    try {
      canonicalAllowed = await realpath(allowedDir);
      canonicalFile = await realpath(lexical);
    } catch {
      throw new Error(`Attachment file not found: ${readPath}`);
    }
    if (!canonicalFile.startsWith(canonicalAllowed + sep) && canonicalFile !== canonicalAllowed) {
      throw new Error(`Attachment path must resolve within ${allowedDir}. Received: ${readPath}`);
    }
    return canonicalFile;
  }

  // Minimal extension → MIME map for local attachment sources. Anything
  // unknown ships as application/octet-stream; callers can override per-file.
  static readonly MIME_BY_EXT: Record<string, string> = {
    pdf: 'application/pdf', png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg',
    gif: 'image/gif', webp: 'image/webp', svg: 'image/svg+xml', txt: 'text/plain',
    md: 'text/markdown', csv: 'text/csv', json: 'application/json', html: 'text/html',
    ics: 'text/calendar', eml: 'message/rfc822', zip: 'application/zip',
    doc: 'application/msword', docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    xls: 'application/vnd.ms-excel', xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    ppt: 'application/vnd.ms-powerpoint', pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  };

  /**
   * Upload raw bytes to the account's JMAP blob endpoint. The uploadUrl
   * template comes from the session (validated at session time) and is
   * re-validated after substitution, mirroring the download path.
   */
  async uploadBlob(buffer: Buffer, type: string): Promise<{ blobId: string; type: string; size: number }> {
    const session = await this.getSession();
    if (!session.uploadUrl) {
      throw new Error('Upload capability not available in session');
    }
    const maxSize = session.capabilities?.['urn:ietf:params:jmap:core']?.maxSizeUpload;
    if (typeof maxSize === 'number' && buffer.length > maxSize) {
      throw new Error(`Attachment is ${buffer.length} bytes; the server's upload limit is ${maxSize} bytes`);
    }

    const url = session.uploadUrl.replace('{accountId}', session.accountId);
    validateFastmailUrl(url, 'uploadUrl', this.auth.getAllowUnsafe());

    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Authorization': this.auth.getAuthHeaders()['Authorization'],
        'Content-Type': type || 'application/octet-stream',
      },
      body: new Uint8Array(buffer),
      // Same rationale as the download path: never follow a redirect with a
      // token-bearing request.
      redirect: 'error',
    });

    if (!response.ok) {
      throw new Error(`Blob upload failed: ${response.status} ${response.statusText}`);
    }

    const result: any = await response.json();
    if (!result?.blobId) {
      throw new Error('Blob upload returned no blobId');
    }
    return { blobId: result.blobId, type: result.type || type, size: result.size ?? buffer.length };
  }

  /**
   * Resolve tool-facing attachment inputs to RFC 8621 attachment parts.
   * Exactly one source per entry:
   *  - { localPath }             file confined to the download directory
   *  - { emailId, attachmentId } zero-copy reuse of an existing attachment's blob
   *  - { blobId, name, type }    a blob already uploaded to this account
   */
  async resolveAttachments(inputs: AttachmentInput[], downloadDir?: string): Promise<any[]> {
    const parts: any[] = [];
    for (const input of inputs) {
      const sources = [input.localPath, input.emailId, input.blobId].filter((v) => v != null).length;
      if (sources !== 1) {
        throw new Error('Each attachment must specify exactly one source: localPath, emailId+attachmentId, or blobId');
      }
      if (input.blobId) {
        parts.push({
          blobId: input.blobId,
          type: input.type || 'application/octet-stream',
          name: input.name || 'attachment',
          disposition: 'attachment',
        });
      } else if (input.emailId) {
        if (!input.attachmentId) {
          throw new Error('attachmentId is required when attaching from an existing email');
        }
        const info = await this.getAttachmentInfo(input.emailId, input.attachmentId);
        parts.push({
          blobId: info.blobId,
          type: input.type || info.type,
          name: input.name || info.name,
          disposition: 'attachment',
        });
      } else {
        const canonical = await JmapClient.validateReadPath(input.localPath!, downloadDir);
        const buffer = await readFile(canonical);
        const ext = canonical.split('.').pop()?.toLowerCase() ?? '';
        const type = input.type || JmapClient.MIME_BY_EXT[ext] || 'application/octet-stream';
        const uploaded = await this.uploadBlob(buffer, type);
        parts.push({
          blobId: uploaded.blobId,
          type,
          name: input.name || basename(canonical),
          disposition: 'attachment',
        });
      }
    }
    return parts;
  }

  async advancedSearch(filters: EmailQueryFilters): Promise<QueryResult> {
    const session = await this.getSession();
    const finalFilter = buildEmailQueryFilter(filters);

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/query', {
          accountId: session.accountId,
          filter: finalFilter,
          sort: [{ property: 'receivedAt', isAscending: filters.ascending ?? false }],
          limit: Math.min(filters.limit || 50, 100),
          calculateTotal: true
        }, 'query'],
        ['Email/get', {
          accountId: session.accountId,
          '#ids': { resultOf: 'query', name: 'Email/query', path: '/ids' },
          properties: ['id', 'subject', 'from', 'to', 'cc', 'replyTo', 'receivedAt', 'preview', 'hasAttachment', 'keywords', 'threadId']
        }, 'emails']
      ]
    };

    const response = await this.makeRequest(request);
    return this.getQueryResult(response, 0, 1);
  }

  async advancedSearchMetadata(filters: EmailQueryFilters): Promise<QueryResult> {
    const session = await this.getSession();
    const finalFilter = buildEmailQueryFilter(filters);

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/query', {
          accountId: session.accountId,
          filter: finalFilter,
          sort: [{ property: 'receivedAt', isAscending: filters.ascending ?? false }],
          limit: Math.min(filters.limit || 50, 100),
          calculateTotal: true
        }, 'query'],
        ['Email/get', {
          accountId: session.accountId,
          '#ids': { resultOf: 'query', name: 'Email/query', path: '/ids' },
          properties: ['id', 'threadId', 'subject', 'from', 'to', 'cc', 'replyTo', 'receivedAt', 'hasAttachment', 'keywords']
        }, 'emails']
      ]
    };

    const response = await this.makeRequest(request);
    return this.getQueryResult(response, 0, 1);
  }

  async searchEmails(query: string, limit: number = 20, ascending: boolean = false, excludeDrafts: boolean = false): Promise<QueryResult> {
    const session = await this.getSession();

    // A JMAP FilterCondition ANDs its properties, so text + notKeyword means
    // "matches the query AND is not a draft". Applied server-side in Email/query.
    const filter: any = { text: query };
    if (excludeDrafts) filter.notKeyword = '$draft';

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/query', {
          accountId: session.accountId,
          filter,
          sort: [{ property: 'receivedAt', isAscending: ascending }],
          limit,
          calculateTotal: true
        }, 'query'],
        ['Email/get', {
          accountId: session.accountId,
          '#ids': { resultOf: 'query', name: 'Email/query', path: '/ids' },
          properties: ['id', 'subject', 'from', 'to', 'replyTo', 'receivedAt', 'preview', 'hasAttachment']
        }, 'emails']
      ]
    };

    const response = await this.makeRequest(request);
    return this.getQueryResult(response, 0, 1);
  }

  async searchEmailsMetadata(query: string, limit: number = 20, ascending: boolean = false): Promise<QueryResult> {
    const session = await this.getSession();

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/query', {
          accountId: session.accountId,
          filter: { text: query },
          sort: [{ property: 'receivedAt', isAscending: ascending }],
          limit,
          calculateTotal: true
        }, 'query'],
        ['Email/get', {
          accountId: session.accountId,
          '#ids': { resultOf: 'query', name: 'Email/query', path: '/ids' },
          properties: ['id', 'threadId', 'subject', 'from', 'to', 'replyTo', 'receivedAt', 'hasAttachment', 'keywords']
        }, 'emails']
      ]
    };

    const response = await this.makeRequest(request);
    return this.getQueryResult(response, 0, 1);
  }

  async getThread(threadId: string, includeDrafts: boolean = false): Promise<any[]> {
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

    // Drafts (e.g. an in-progress reply) are noise when reading a conversation,
    // so exclude them by default. Identify by the $draft keyword (survives a
    // draft moved out of the Drafts mailbox); opt back in via includeDrafts.
    const emails = this.getListResult(response, 1);
    return includeDrafts ? emails : emails.filter((e: any) => !e.keywords?.$draft);
  }

  async getThreadMetadata(threadId: string): Promise<any[]> {
    const session = await this.getSession();

    // Resolve threadId — accept either an email ID or a thread ID, mirroring getThread.
    let actualThreadId = threadId;

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
          properties: ['id', 'threadId', 'subject', 'from', 'to', 'cc', 'replyTo', 'receivedAt', 'hasAttachment', 'keywords']
        }, 'emails']
      ]
    };

    const response = await this.makeRequest(request);
    const threadResult = this.getMethodResult(response, 0);

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