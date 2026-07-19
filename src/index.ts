#!/usr/bin/env node
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ErrorCode,
  ListToolsRequestSchema,
  McpError,
} from '@modelcontextprotocol/sdk/types.js';
import { FastmailAuth, FastmailConfig } from './auth.js';
import { JmapClient, QueryResult } from './jmap-client.js';
import { ContactsCalendarClient } from './contacts-calendar.js';
import { CalDAVCalendarClient } from './caldav-client.js';
import { WebDAVFilesClient } from './webdav-files-client.js';
import { validateHttpsUrl } from './url-validation.js';
import { coerceRecipients, coerceStringArray, coerceBool, redactBearerTokens, registerSecret } from './coerce.js';

const server = new Server(
  {
    name: 'fastmail-mcp',
    version: '1.12.0',
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

let jmapClient: JmapClient | null = null;
let contactsCalendarClient: ContactsCalendarClient | null = null;
let caldavClient: CalDAVCalendarClient | null = null;

function findEnvValue(keys: string[]): { value?: string; key?: string; wasPlaceholder: boolean } {
  const isPlaceholder = (val: string) => /\$\{[^}]+\}/.test(val.trim());
  for (const key of keys) {
    const raw = process.env[key];
    if (typeof raw === 'string' && raw.trim().length > 0) {
      if (isPlaceholder(raw)) {
        return { value: undefined, key, wasPlaceholder: true };
      }
      return { value: raw.trim(), key, wasPlaceholder: false };
    }
  }
  return { value: undefined, key: undefined, wasPlaceholder: false };
}


function getAuthConfig(): FastmailConfig {
  const tokenInfo = findEnvValue([
    'FASTMAIL_API_TOKEN',
    'USER_CONFIG_FASTMAIL_API_TOKEN',
    'USER_CONFIG_fastmail_api_token',
    'fastmail_api_token',
  ]);
  const apiToken = tokenInfo.value;
  if (!apiToken) {
    throw new McpError(
      ErrorCode.InvalidRequest,
      'FASTMAIL_API_TOKEN environment variable is required'
    );
  }
  // Register for value-based redaction so an exact token occurrence in any
  // error string is scrubbed even if it doesn't match the token-shape pattern.
  registerSecret(apiToken);

  const baseInfo = findEnvValue([
    'FASTMAIL_BASE_URL',
    'USER_CONFIG_FASTMAIL_BASE_URL',
    'USER_CONFIG_fastmail_base_url',
    'fastmail_base_url',
  ]);

  // Opt-in for self-hosted JMAP servers. Required to use any base URL outside
  // the api.fastmail.com / www.fastmailusercontent.com allowlist.
  const unsafeInfo = findEnvValue([
    'FASTMAIL_ALLOW_UNSAFE_BASE_URL',
    'USER_CONFIG_FASTMAIL_ALLOW_UNSAFE_BASE_URL',
  ]);
  const allowUnsafeBaseUrl = unsafeInfo.value === 'true' || unsafeInfo.value === '1';

  return { apiToken, baseUrl: baseInfo.value, allowUnsafeBaseUrl };
}

function initializeClient(): JmapClient {
  if (jmapClient) {
    return jmapClient;
  }

  const auth = new FastmailAuth(getAuthConfig());
  jmapClient = new JmapClient(auth);
  return jmapClient;
}

function initializeContactsCalendarClient(): ContactsCalendarClient {
  if (contactsCalendarClient) {
    return contactsCalendarClient;
  }

  const auth = new FastmailAuth(getAuthConfig());
  contactsCalendarClient = new ContactsCalendarClient(auth);
  return contactsCalendarClient;
}

function initializeCalDAVClient(): CalDAVCalendarClient | null {
  if (caldavClient) return caldavClient;

  const username = findEnvValue([
    'FASTMAIL_CALDAV_USERNAME',
    'USER_CONFIG_FASTMAIL_CALDAV_USERNAME',
  ]).value;
  const password = findEnvValue([
    'FASTMAIL_CALDAV_PASSWORD',
    'USER_CONFIG_FASTMAIL_CALDAV_PASSWORD',
  ]).value;

  if (!username || !password) return null;

  // Register the CalDAV password (and username) for value-based redaction —
  // Basic-auth credentials aren't covered by the token-shape patterns.
  registerSecret(password);
  registerSecret(username);

  caldavClient = new CalDAVCalendarClient({ username, password });
  return caldavClient;
}

let webdavClient: WebDAVFilesClient | null = null;

function initializeWebDAVClient(): WebDAVFilesClient | null {
  if (webdavClient) return webdavClient;

  const baseUrl = findEnvValue([
    'FASTMAIL_WEBDAV_URL',
    'USER_CONFIG_FASTMAIL_WEBDAV_URL',
  ]).value;
  const username = findEnvValue([
    'FASTMAIL_WEBDAV_USERNAME',
    'USER_CONFIG_FASTMAIL_WEBDAV_USERNAME',
  ]).value;
  const password = findEnvValue([
    'FASTMAIL_WEBDAV_PASSWORD',
    'USER_CONFIG_FASTMAIL_WEBDAV_PASSWORD',
  ]).value;

  if (!baseUrl || !username || !password) return null;

  // The base URL is server config, never a tool argument — tools may only
  // supply relative paths beneath it. HTTPS-only, no embedded credentials.
  validateHttpsUrl(baseUrl, 'FASTMAIL_WEBDAV_URL');
  registerSecret(password);
  registerSecret(username);

  webdavClient = new WebDAVFilesClient({ baseUrl, username, password });
  return webdavClient;
}

function getDownloadDir(): string | undefined {
  return findEnvValue([
    'FASTMAIL_DOWNLOAD_DIR',
    'USER_CONFIG_FASTMAIL_DOWNLOAD_DIR',
    'USER_CONFIG_fastmail_download_dir',
    'fastmail_download_dir',
  ]).value;
}

// Clamp a caller-supplied limit into [1, max], tolerating string and NaN input
// (a lenient client may send "20", and a bare Number("abc") would yield NaN →
// an unbounded/negative JMAP query).
function clampLimit(value: unknown, fallback: number, max: number): number {
  return Math.min(Math.max(Number(value) || fallback, 1), max);
}

// When the JMAP server reports a total match count (calculateTotal: true), wrap
// the items in a { total, items } envelope so callers can tell a truncated page
// from a complete result. Without a total the bare array is returned unchanged,
// keeping output byte-identical for paths that never had one (CalDAV, contacts
// fallback).
function formatQueryResult(result: QueryResult): string {
  const { items, total } = result;
  if (total != null) {
    return JSON.stringify({ total, items }, null, 2);
  }
  return JSON.stringify(items, null, 2);
}

server.setRequestHandler(ListToolsRequestSchema, async () => {
  return {
    tools: [
      {
        name: 'list_mailboxes',
        description: 'List mailboxes in the Fastmail account. By default returns all mailboxes with full metadata; on accounts with hundreds of mailboxes the full result can exceed the MCP tool result window. Use `properties: ["id","name","parentId"]` for a slim view, and/or `parentId` to filter to one level of children.',
        inputSchema: {
          type: 'object',
          properties: {
            properties: {
              type: 'array',
              items: { type: 'string' },
              description: 'JMAP Mailbox properties to return (e.g. ["id","name","parentId"]). Default: all properties. The slim form roughly halves payload size on large accounts.',
            },
            parentId: {
              type: ['string', 'null'],
              description: 'Filter to direct children of this mailbox ID. Pass null for top-level mailboxes. Filter is applied client-side after Mailbox/get.',
            },
          },
        },
      },
      {
        name: 'get_mailbox_by_name',
        description: 'Look up a single mailbox by its full path from root (e.g. "Folder/Subfolder/Leaf"). Returns the mailbox ID and minimal metadata, or throws "Mailbox not found" if no exact match. The path separator is "/"; folder names containing a literal "/" are not supported.',
        inputSchema: {
          type: 'object',
          properties: {
            path: {
              type: 'string',
              description: 'Full path from root, separated by "/" (e.g. "Inbox" or "Archive/2026/Suppliers/ExampleCo").',
            },
          },
          required: ['path'],
        },
      },
      {
        name: 'create_mailbox',
        description: 'Create a new mailbox (folder). Returns the new mailbox ID. The caller is responsible for validating the name is appropriate (length, character set, parent-folder allow-list) before calling — JMAP itself only enforces uniqueness within a parent.',
        inputSchema: {
          type: 'object',
          properties: {
            name: {
              type: 'string',
              description: 'Leaf name of the new mailbox (not a full path). Must not contain "/".',
            },
            parentId: {
              type: ['string', 'null'],
              description: 'Parent mailbox ID. Pass null (or omit) to create at top level.',
            },
          },
          required: ['name'],
        },
      },
      {
        name: 'list_emails',
        description: 'List emails from a mailbox. When the server reports a total match count, results are wrapped in a {"total", "items"} JSON envelope; otherwise a bare JSON array is returned.',
        inputSchema: {
          type: 'object',
          properties: {
            mailboxId: {
              type: 'string',
              description: 'ID of the mailbox to list emails from (optional, defaults to all)',
            },
            limit: {
              type: ['number', 'string'],
              description: 'Maximum number of emails to return (default: 20)',
              default: 20,
            },
            ascending: {
              type: 'boolean',
              description: 'Sort oldest first instead of newest first (default: false)',
            },
          },
        },
      },
      {
        name: 'list_emails_metadata',
        description: 'Same as list_emails (lists emails from a mailbox, optionally filtered by mailboxId, with paging and sort) but returns ONLY metadata fields on each result — id, threadId, subject, from, to, replyTo, receivedAt, hasAttachment, keywords. Does NOT return preview or any body-derived content. Use in privacy-sensitive flows where the workflow needs only the envelope (e.g. customer-mail least-privilege scans, or any caller forbidden from ingesting message bodies). Pair with get_email_metadata for follow-up lookups that should also stay header-only. When the server reports a total match count, results are wrapped in a {"total", "items"} JSON envelope; otherwise a bare JSON array is returned.',
        inputSchema: {
          type: 'object',
          properties: {
            mailboxId: {
              type: 'string',
              description: 'ID of the mailbox to list emails from (optional, defaults to all)',
            },
            limit: {
              type: ['number', 'string'],
              description: 'Maximum number of emails to return (default: 20)',
              default: 20,
            },
            ascending: {
              type: 'boolean',
              description: 'Sort oldest first instead of newest first (default: false)',
            },
          },
        },
      },
      {
        name: 'get_email',
        description: 'Get a specific email by ID',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: 'ID of the email to retrieve',
            },
          },
          required: ['emailId'],
        },
        _meta: {
          // Raise the per-tool result size limit honoured by Claude Code
          // (v2.1.91+) and other MCP clients that respect this annotation.
          // get_email returns the full Email object including textBody/
          // htmlBody/bodyValues/attachments — promotional newsletters and
          // policy-update emails routinely exceed the default ~25KB inline
          // budget and get spilled to a temp file by the harness, which
          // then forces the caller to do its own file-read recovery.
          // 500000 chars (~500KB) covers virtually all real-world email
          // payloads while remaining well under the MCP hard ceiling.
          'anthropic/maxResultSizeChars': 500000,
        },
      },
      {
        name: 'get_email_metadata',
        description: 'Get headers/metadata for an email — sender, recipients, subject, date, threading, mailbox membership, keywords (read/flagged/etc.), size, and whether an attachment is present — but NOT the body, preview, or any rendered text. Useful when a workflow needs to classify or route an email without ingesting its content (e.g. customer-mail least-privilege flows where reading bodies is forbidden, or skills that only need to verify post-archive folder placement). The return shape is the standard JMAP Email object restricted to a strict header-only allowlist.',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: 'ID of the email to retrieve metadata for',
            },
          },
          required: ['emailId'],
        },
      },
      {
        name: 'send_email',
        description: 'Send an email',
        inputSchema: {
          type: 'object',
          properties: {
            to: {
              oneOf: [
                { type: 'array', items: { type: 'string' } },
                { type: 'string' },
              ],
              description: 'Recipient email addresses (array of strings, or a comma-separated string)',
            },
            cc: {
              type: 'array',
              items: { type: 'string' },
              description: 'CC email addresses (optional)',
            },
            bcc: {
              type: 'array',
              items: { type: 'string' },
              description: 'BCC email addresses (optional)',
            },
            from: {
              type: 'string',
              description: 'Sender email address (optional, defaults to account primary email)',
            },
            mailboxId: {
              type: 'string',
              description: 'Mailbox ID to save the email to (optional, defaults to Drafts folder)',
            },
            subject: {
              type: 'string',
              description: 'Email subject',
            },
            textBody: {
              type: 'string',
              description: 'Plain text body (optional)',
            },
            htmlBody: {
              type: 'string',
              description: 'HTML body (optional)',
            },
            inReplyTo: {
              type: 'array',
              items: { type: 'string' },
              description: 'Message-ID(s) of the email being replied to (optional, for threading)',
            },
            references: {
              type: 'array',
              items: { type: 'string' },
              description: 'Full reference chain of Message-IDs (optional, for threading)',
            },
            replyTo: {
              type: 'array',
              items: { type: 'string' },
              description: 'Reply-To email addresses (replies go here instead of to the sender)',
            },
            attachments: {
              type: 'array',
              description: 'Files to attach. Each entry must use EXACTLY ONE source: localPath (a file inside the configured download directory), emailId + attachmentId (re-attach an attachment from an existing email — no bytes are copied), or blobId (an already-uploaded JMAP blob). Optional name/type override the inferred filename and MIME type.',
              items: {
                type: 'object',
                properties: {
                  localPath: { type: 'string', description: 'Path to a file within FASTMAIL_DOWNLOAD_DIR (relative paths resolve against it)' },
                  emailId: { type: 'string', description: 'Source email ID (use with attachmentId)' },
                  attachmentId: { type: 'string', description: 'Attachment partId, blobId, or zero-based index within the source email' },
                  blobId: { type: 'string', description: 'An existing blob ID in this account' },
                  name: { type: 'string', description: 'Override the attachment filename' },
                  type: { type: 'string', description: 'Override the MIME type' },
                },
              },
            },
          },
          required: ['to', 'subject'],
        },
      },
      {
        name: 'reply_email',
        description: 'Reply to an existing email with proper threading headers (In-Reply-To, References). Automatically fetches the original email to build the reply chain. By default sends immediately; set send=false to save as a draft instead.',
        inputSchema: {
          type: 'object',
          properties: {
            originalEmailId: {
              type: 'string',
              description: 'ID of the email to reply to',
            },
            to: {
              type: 'array',
              items: { type: 'string' },
              description: 'Recipient email addresses (optional, defaults to the original sender)',
            },
            cc: {
              type: 'array',
              items: { type: 'string' },
              description: 'CC email addresses (optional)',
            },
            bcc: {
              type: 'array',
              items: { type: 'string' },
              description: 'BCC email addresses (optional)',
            },
            from: {
              type: 'string',
              description: 'Sender email address (optional, defaults to account primary email)',
            },
            textBody: {
              type: 'string',
              description: 'Plain text body (optional)',
            },
            htmlBody: {
              type: 'string',
              description: 'HTML body (optional)',
            },
            send: {
              type: ['boolean', 'string'],
              description: 'Whether to send the reply immediately (default: true). Set to false to save as draft instead.',
            },
            replyTo: {
              type: 'array',
              items: { type: 'string' },
              description: 'Reply-To email addresses (replies go here instead of to the sender)',
            },
            attachments: {
              type: 'array',
              description: 'Files to attach. Each entry must use EXACTLY ONE source: localPath (a file inside the configured download directory), emailId + attachmentId (re-attach an attachment from an existing email — no bytes are copied), or blobId (an already-uploaded JMAP blob). Optional name/type override the inferred filename and MIME type.',
              items: {
                type: 'object',
                properties: {
                  localPath: { type: 'string', description: 'Path to a file within FASTMAIL_DOWNLOAD_DIR (relative paths resolve against it)' },
                  emailId: { type: 'string', description: 'Source email ID (use with attachmentId)' },
                  attachmentId: { type: 'string', description: 'Attachment partId, blobId, or zero-based index within the source email' },
                  blobId: { type: 'string', description: 'An existing blob ID in this account' },
                  name: { type: 'string', description: 'Override the attachment filename' },
                  type: { type: 'string', description: 'Override the MIME type' },
                },
              },
            },
          },
          required: ['originalEmailId'],
        },
      },
      {
        name: 'create_draft',
        description: 'Create an email draft without sending it. Supports threading headers for replies. IMPORTANT: each call creates a new draft — do not call twice for the same message.',
        inputSchema: {
          type: 'object',
          properties: {
            to: {
              type: 'array',
              items: { type: 'string' },
              description: 'Recipient email addresses (optional)',
            },
            cc: {
              type: 'array',
              items: { type: 'string' },
              description: 'CC email addresses (optional)',
            },
            bcc: {
              type: 'array',
              items: { type: 'string' },
              description: 'BCC email addresses (optional)',
            },
            from: {
              type: 'string',
              description: 'Sender email address (optional, defaults to account primary email)',
            },
            mailboxId: {
              type: 'string',
              description: 'Mailbox ID to save the draft to (optional, defaults to Drafts folder)',
            },
            subject: {
              type: 'string',
              description: 'Email subject (optional)',
            },
            textBody: {
              type: 'string',
              description: 'Plain text body (optional)',
            },
            htmlBody: {
              type: 'string',
              description: 'HTML body (optional)',
            },
            inReplyTo: {
              type: 'array',
              items: { type: 'string' },
              description: 'Message-IDs to reply to (optional, for threading)',
            },
            references: {
              type: 'array',
              items: { type: 'string' },
              description: 'Message-IDs for References header (optional, for threading)',
            },
            replyTo: {
              type: 'array',
              items: { type: 'string' },
              description: 'Reply-To email addresses (replies go here instead of to the sender)',
            },
            attachments: {
              type: 'array',
              description: 'Files to attach. Each entry must use EXACTLY ONE source: localPath (a file inside the configured download directory), emailId + attachmentId (re-attach an attachment from an existing email — no bytes are copied), or blobId (an already-uploaded JMAP blob). Optional name/type override the inferred filename and MIME type.',
              items: {
                type: 'object',
                properties: {
                  localPath: { type: 'string', description: 'Path to a file within FASTMAIL_DOWNLOAD_DIR (relative paths resolve against it)' },
                  emailId: { type: 'string', description: 'Source email ID (use with attachmentId)' },
                  attachmentId: { type: 'string', description: 'Attachment partId, blobId, or zero-based index within the source email' },
                  blobId: { type: 'string', description: 'An existing blob ID in this account' },
                  name: { type: 'string', description: 'Override the attachment filename' },
                  type: { type: 'string', description: 'Override the MIME type' },
                },
              },
            },
          },
        },
      },
      {
        name: 'edit_draft',
        description: 'Edit an existing draft email. Since JMAP emails are immutable, this atomically destroys the old draft and creates a new one with the updated fields. Only fields you provide will be changed; others are preserved from the original draft.',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: 'The ID of the draft email to edit',
            },
            to: {
              type: 'array',
              items: { type: 'string' },
              description: 'Updated recipient email addresses (optional, keeps existing if omitted)',
            },
            cc: {
              type: 'array',
              items: { type: 'string' },
              description: 'Updated CC email addresses (optional)',
            },
            bcc: {
              type: 'array',
              items: { type: 'string' },
              description: 'Updated BCC email addresses (optional)',
            },
            from: {
              type: 'string',
              description: 'Updated sender email address (optional)',
            },
            subject: {
              type: 'string',
              description: 'Updated email subject (optional)',
            },
            textBody: {
              type: 'string',
              description: 'Updated plain text body (optional)',
            },
            htmlBody: {
              type: 'string',
              description: 'Updated HTML body (optional)',
            },
            replyTo: {
              type: 'array',
              items: { type: 'string' },
              description: 'Reply-To email addresses (replies go here instead of to the sender)',
            },
            attachments: {
              type: 'array',
              description: 'Files to attach. Each entry must use EXACTLY ONE source: localPath (a file inside the configured download directory), emailId + attachmentId (re-attach an attachment from an existing email — no bytes are copied), or blobId (an already-uploaded JMAP blob). Optional name/type override the inferred filename and MIME type.',
              items: {
                type: 'object',
                properties: {
                  localPath: { type: 'string', description: 'Path to a file within FASTMAIL_DOWNLOAD_DIR (relative paths resolve against it)' },
                  emailId: { type: 'string', description: 'Source email ID (use with attachmentId)' },
                  attachmentId: { type: 'string', description: 'Attachment partId, blobId, or zero-based index within the source email' },
                  blobId: { type: 'string', description: 'An existing blob ID in this account' },
                  name: { type: 'string', description: 'Override the attachment filename' },
                  type: { type: 'string', description: 'Override the MIME type' },
                },
              },
            },
          },
          required: ['emailId'],
        },
      },
      {
        name: 'send_draft',
        description: 'Send an existing draft email. The draft must have recipients (to/cc/bcc) and a from address. After sending, the email is moved to the Sent folder and the draft keyword is removed.',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: 'The ID of the draft email to send',
            },
          },
          required: ['emailId'],
        },
      },
      {
        name: 'search_emails',
        description:
          'Full-text search of email body and subject. Does not filter by sender, recipient, or date — use advanced_search for field-specific filtering. Drafts are included by default; set excludeDrafts=true to omit draft messages from results. When the server reports a total match count, results are wrapped in a {"total", "items"} JSON envelope; otherwise a bare JSON array is returned.',
        inputSchema: {
          type: 'object',
          properties: {
            query: {
              type: 'string',
              description: 'Text to search for in email body and subject lines',
            },
            limit: {
              type: ['number', 'string'],
              description: 'Maximum number of results (default: 20)',
              default: 20,
            },
            ascending: {
              type: 'boolean',
              description: 'Sort oldest first instead of newest first (default: false)',
            },
            excludeDrafts: {
              type: 'boolean',
              description: 'Omit draft messages from results (default: false, drafts included). Filtered server-side via the $draft keyword.',
            },
          },
          required: ['query'],
        },
      },
      {
        name: 'search_emails_metadata',
        description: 'Same as search_emails (free-text search across subject and body) but returns ONLY metadata on each match — id, threadId, subject, from, to, replyTo, receivedAt, hasAttachment, keywords. The query still searches body text on the server side; only the result envelopes come back, never preview or body excerpts. Use when a content match is required (e.g. "find all messages mentioning X") but the matches must not surface body fragments to the caller. When the server reports a total match count, results are wrapped in a {"total", "items"} JSON envelope; otherwise a bare JSON array is returned.',
        inputSchema: {
          type: 'object',
          properties: {
            query: {
              type: 'string',
              description: 'Search query string',
            },
            limit: {
              type: ['number', 'string'],
              description: 'Maximum number of results (default: 20)',
              default: 20,
            },
            ascending: {
              type: 'boolean',
              description: 'Sort oldest first instead of newest first (default: false)',
            },
          },
          required: ['query'],
        },
      },
      {
        name: 'list_contacts',
        description: 'List contacts from the address book. When the server reports a total match count, results are wrapped in a {"total", "items"} JSON envelope; otherwise a bare JSON array is returned.',
        inputSchema: {
          type: 'object',
          properties: {
            limit: {
              type: 'number',
              description: 'Maximum number of contacts to return (default: 50)',
              default: 50,
            },
          },
        },
      },
      {
        name: 'get_contact',
        description: 'Get a specific contact by ID',
        inputSchema: {
          type: 'object',
          properties: {
            contactId: {
              type: 'string',
              description: 'ID of the contact to retrieve',
            },
          },
          required: ['contactId'],
        },
      },
      {
        name: 'search_contacts',
        description: 'Search contacts by name or email. When the server reports a total match count, results are wrapped in a {"total", "items"} JSON envelope; otherwise a bare JSON array is returned.',
        inputSchema: {
          type: 'object',
          properties: {
            query: {
              type: 'string',
              description: 'Search query string',
            },
            limit: {
              type: 'number',
              description: 'Maximum number of results (default: 20)',
              default: 20,
            },
          },
          required: ['query'],
        },
      },
      {
        name: 'create_contact',
        description: 'Create a new contact in the address book. Requires a name or at least one email address. Requires an API token with read-write contacts scope.',
        inputSchema: {
          type: 'object',
          properties: {
            name: {
              type: 'object',
              description: 'Structured name; provide full and/or given/surname',
              properties: {
                given: { type: 'string' },
                surname: { type: 'string' },
                full: { type: 'string', description: 'Full display name' },
              },
            },
            emails: {
              type: 'array',
              description: 'Email addresses (replaces ALL existing emails on update; [] clears)',
              items: {
                type: 'object',
                properties: {
                  address: { type: 'string' },
                  label: { type: 'string', description: 'Optional label, e.g. work / home' },
                },
                required: ['address'],
              },
            },
            phones: {
              type: 'array',
              description: 'Phone numbers (replaces ALL existing phones on update)',
              items: {
                type: 'object',
                properties: {
                  number: { type: 'string' },
                  label: { type: 'string' },
                },
                required: ['number'],
              },
            },
            addresses: {
              type: 'array',
              description: 'Postal addresses as free-form text (replaces ALL existing on update)',
              items: {
                type: 'object',
                properties: {
                  full: { type: 'string', description: 'Full address as one string' },
                  label: { type: 'string' },
                },
                required: ['full'],
              },
            },
            notes: { type: 'string', description: 'Free-form note (replaces the existing note on update)' },
            addressBookId: { type: 'string', description: 'Target address book id (default book when omitted)' },
          },
        },
      },
      {
        name: 'update_contact',
        description: 'Update an existing contact. Each provided field WHOLLY REPLACES the stored value (e.g. emails: [] removes all emails) — unspecified fields are left untouched. Requires read-write contacts scope.',
        inputSchema: {
          type: 'object',
          properties: {
            contactId: { type: 'string', description: 'ID of the contact to update' },
            name: {
              type: 'object',
              description: 'Structured name; provide full and/or given/surname',
              properties: {
                given: { type: 'string' },
                surname: { type: 'string' },
                full: { type: 'string', description: 'Full display name' },
              },
            },
            emails: {
              type: 'array',
              description: 'Email addresses (replaces ALL existing emails on update; [] clears)',
              items: {
                type: 'object',
                properties: {
                  address: { type: 'string' },
                  label: { type: 'string', description: 'Optional label, e.g. work / home' },
                },
                required: ['address'],
              },
            },
            phones: {
              type: 'array',
              description: 'Phone numbers (replaces ALL existing phones on update)',
              items: {
                type: 'object',
                properties: {
                  number: { type: 'string' },
                  label: { type: 'string' },
                },
                required: ['number'],
              },
            },
            addresses: {
              type: 'array',
              description: 'Postal addresses as free-form text (replaces ALL existing on update)',
              items: {
                type: 'object',
                properties: {
                  full: { type: 'string', description: 'Full address as one string' },
                  label: { type: 'string' },
                },
                required: ['full'],
              },
            },
            notes: { type: 'string', description: 'Free-form note (replaces the existing note on update)' },
            expectState: { type: 'string', description: 'Optional JMAP state precondition (ifInState); update fails with stateMismatch if contacts changed since this state' },
          },
          required: ['contactId'],
        },
      },
      {
        name: 'delete_contact',
        description: 'Permanently delete a contact from the address book. This cannot be undone. Requires read-write contacts scope.',
        inputSchema: {
          type: 'object',
          properties: {
            contactId: { type: 'string', description: 'ID of the contact to delete' },
            expectState: { type: 'string', description: 'Optional JMAP state precondition (ifInState)' },
          },
          required: ['contactId'],
        },
      },
      {
        name: 'list_calendars',
        description: 'List all calendars',
        inputSchema: {
          type: 'object',
          properties: {},
        },
      },
      {
        name: 'list_calendar_events',
        description: 'List events from a calendar',
        inputSchema: {
          type: 'object',
          properties: {
            calendarId: {
              type: 'string',
              description: 'ID of the calendar (optional, defaults to all calendars)',
            },
            startDate: {
              type: 'string',
              description: 'Filter events starting from this date (ISO 8601, e.g. 2026-03-23T00:00:00Z)',
            },
            endDate: {
              type: 'string',
              description: 'Filter events ending before this date (ISO 8601, e.g. 2026-03-30T00:00:00Z)',
            },
            limit: {
              type: 'number',
              description: 'Maximum number of events to return (default: 50)',
              default: 50,
            },
          },
        },
      },
      {
        name: 'get_calendar_event',
        description: 'Get a specific calendar event by ID. Returns organizer and participants when available.',
        inputSchema: {
          type: 'object',
          properties: {
            eventId: {
              type: 'string',
              description: 'ID of the event to retrieve',
            },
          },
          required: ['eventId'],
        },
      },
      {
        name: 'create_calendar_event',
        description: 'Create a new calendar event. Supports date-only (e.g. 2026-04-01) for all-day events. DTEND is exclusive per RFC 5545 — a one-day event on April 1 needs end: 2026-04-02.',
        inputSchema: {
          type: 'object',
          properties: {
            calendarId: {
              type: 'string',
              description: 'ID of the calendar to create the event in',
            },
            title: {
              type: 'string',
              description: 'Event title',
            },
            description: {
              type: 'string',
              description: 'Event description (optional)',
            },
            start: {
              type: 'string',
              description: 'Start time in ISO 8601 format (e.g. 2026-04-07T14:00:00Z) or date-only for all-day events (e.g. 2026-04-07)',
            },
            end: {
              type: 'string',
              description: 'End time in ISO 8601 format. For all-day events, DTEND is exclusive — a one-day event on April 1 requires end: 2026-04-02',
            },
            location: {
              type: 'string',
              description: 'Event location (optional)',
            },
            participants: {
              type: 'array',
              items: {
                type: 'object',
                properties: {
                  email: { type: 'string', description: 'Participant email address' },
                  name: { type: 'string', description: 'Participant display name (optional)' }
                },
                required: ['email'],
              },
              description: 'Event participants (optional). Automatically adds ORGANIZER from CalDAV username.',
            },
          },
          required: ['calendarId', 'title', 'start', 'end'],
        },
      },
      {
        name: 'update_calendar_event',
        description: 'Update an existing calendar event. Preserves all existing data (attendees, reminders, recurrence rules, etc.) not being changed. Omit a field to leave it unchanged; passing an empty/whitespace string for title, description, or location is rejected (use clearFields to delete description/location). Floating times preserve the original timezone; explicit UTC/offset times convert to UTC. WARNING: providing participants replaces ALL existing attendee data (acceptance status, roles, etc.). participants: [] removes all attendees.',
        inputSchema: {
          type: 'object',
          properties: {
            eventId: {
              type: 'string',
              description: 'ID of the event to update',
            },
            title: {
              type: 'string',
              description: 'New event title',
            },
            description: {
              type: 'string',
              description: 'New event description',
            },
            start: {
              type: 'string',
              description: 'New start time in ISO 8601 format. Floating times (no Z/offset) preserve original timezone',
            },
            end: {
              type: 'string',
              description: 'New end time in ISO 8601 format. DTEND is exclusive per RFC 5545',
            },
            location: {
              type: 'string',
              description: 'New event location',
            },
            participants: {
              type: 'array',
              items: {
                type: 'object',
                properties: {
                  email: { type: 'string', description: 'Participant email address' },
                  name: { type: 'string', description: 'Participant display name (optional)' }
                },
                required: ['email'],
              },
              description: 'Replaces ALL existing attendees. Empty array removes all attendees. Omit to preserve existing attendees.',
            },
            clearFields: {
              type: 'array',
              items: { type: 'string', enum: ['description', 'location'] },
              description: 'Property names to delete from the event. Allowed: description, location. Cannot also pass the same field as a value.',
            },
            confirmRecurring: {
              type: 'boolean',
              description: 'Required when changing start/end on a recurring event with exceptions. Acknowledges that orphaned exception overrides will be removed.',
            },
          },
          required: ['eventId'],
        },
      },
      {
        name: 'delete_calendar_event',
        description: 'Delete a calendar event by ID',
        inputSchema: {
          type: 'object',
          properties: {
            eventId: {
              type: 'string',
              description: 'ID of the event to delete',
            },
          },
          required: ['eventId'],
        },
      },
      {
        name: 'list_identities',
        description: 'List sending identities (email addresses that can be used for sending)',
        inputSchema: {
          type: 'object',
          properties: {},
        },
      },
      {
        name: 'get_recent_emails',
        description: 'Get the most recent emails across all mailboxes except Trash and Spam (pass mailboxName to scope to one folder, e.g. "inbox"). When the server reports a total match count, results are wrapped in a {"total", "items"} JSON envelope; otherwise a bare JSON array is returned.',
        inputSchema: {
          type: 'object',
          properties: {
            limit: {
              type: ['number', 'string'],
              description: 'Number of recent emails to retrieve (default: 10, max: 50)',
              default: 10,
            },
            mailboxName: {
              type: 'string',
              description: 'Mailbox to search (optional; when omitted, all mailboxes except Trash and Spam are searched)',
            },
            ascending: {
              type: 'boolean',
              description: 'Sort oldest first instead of newest first (default: false)',
            },
          },
        },
      },
      {
        name: 'mark_email_read',
        description: 'Mark an email as read or unread',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: 'ID of the email to mark',
            },
            read: {
              type: 'boolean',
              description: 'true to mark as read, false to mark as unread',
              default: true,
            },
          },
          required: ['emailId'],
        },
      },
      {
        name: 'pin_email',
        description: 'Pin or unpin an email',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: 'ID of the email to pin/unpin',
            },
            pinned: {
              type: 'boolean',
              description: 'true to pin, false to unpin',
              default: true,
            },
          },
          required: ['emailId'],
        },
      },
      {
        name: 'delete_email',
        description: 'Delete an email (move to trash)',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: 'ID of the email to delete',
            },
          },
          required: ['emailId'],
        },
      },
      {
        name: 'move_email',
        description: 'Move an email to a different mailbox',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: 'ID of the email to move',
            },
            targetMailboxId: {
              type: 'string',
              description: 'ID of the target mailbox',
            },
          },
          required: ['emailId', 'targetMailboxId'],
        },
      },
      {
        name: 'archive_email',
        description: 'Archive an email — move it to the target mailbox AND mark it as read in a single atomic JMAP operation. Equivalent to calling move_email followed by mark_email_read, but in one MCP call and one Email/set patch (the move and the read flag land together or not at all). For trashing an email, use delete_email instead — that follows a different convention and does not auto-mark-read.',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: 'ID of the email to archive',
            },
            targetMailboxId: {
              type: 'string',
              description: 'ID of the destination mailbox',
            },
          },
          required: ['emailId', 'targetMailboxId'],
        },
      },
      {
        name: 'add_labels',
        description: 'Add labels (mailboxes) to an email without removing existing ones',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: 'ID of the email to add labels to',
            },
            mailboxIds: {
              type: 'array',
              items: { type: 'string' },
              description: 'Array of mailbox IDs to add as labels',
            },
          },
          required: ['emailId', 'mailboxIds'],
        },
      },
      {
        name: 'remove_labels',
        description: 'Remove specific labels (mailboxes) from an email',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: 'ID of the email to remove labels from',
            },
            mailboxIds: {
              type: 'array',
              items: { type: 'string' },
              description: 'Array of mailbox IDs to remove as labels',
            },
          },
          required: ['emailId', 'mailboxIds'],
        },
      },
      {
        name: 'get_email_attachments',
        description: 'Get list of attachments for an email',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: 'ID of the email',
            },
          },
          required: ['emailId'],
        },
      },
      {
        name: 'download_attachment',
        description: 'Download an email attachment. If savePath is provided, saves the file to disk and returns the file path and size. Otherwise returns a download URL.',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: 'ID of the email',
            },
            attachmentId: {
              type: 'string',
              description: 'ID of the attachment',
            },
            savePath: {
              type: 'string',
              description: `File path to save the attachment to. May be absolute or relative; relative paths resolve against ${getDownloadDir() || '~/Downloads/fastmail-mcp/'} (configurable via FASTMAIL_DOWNLOAD_DIR), so a bare filename lands there in one step. Absolute paths must fall within that directory; traversal or symlink escape outside it is rejected for security. To save directly into your own location, set FASTMAIL_DOWNLOAD_DIR to that root. Parent directories will be created automatically.`,
            },
          },
          required: ['emailId', 'attachmentId'],
        },
      },
      {
        name: 'save_attachment_to_webdav',
        description: 'Save an email attachment directly to WebDAV cloud storage (e.g. Fastmail Files or Nextcloud) without touching local disk. The storage server and credentials come from server configuration (FASTMAIL_WEBDAV_URL / _USERNAME / _PASSWORD); this tool only chooses the relative path beneath that base. Fails if the remote file exists unless overwrite is set.',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: 'ID of the email',
            },
            attachmentId: {
              type: 'string',
              description: 'Attachment partId, blobId, or zero-based index',
            },
            remotePath: {
              type: 'string',
              description: 'Relative path under the configured WebDAV base (e.g. "invoices/2026/receipt.pdf"). No leading slash, no "..", forward slashes only. Missing parent folders are created unless createParents is false.',
            },
            overwrite: {
              type: 'boolean',
              description: 'Replace an existing remote file (default false: fail if it exists)',
            },
            createParents: {
              type: 'boolean',
              description: 'Create missing parent collections via MKCOL (default true)',
            },
          },
          required: ['emailId', 'attachmentId', 'remotePath'],
        },
      },
      {
        name: 'advanced_search',
        description: 'Advanced email search with multiple criteria. Mailbox scoping supports a single mailbox (mailboxId), an intersection of multiple mailboxes (requiredMailboxIds — must be a member of ALL listed mailboxes), and exclusion (excludeMailboxIds — member of NONE of the listed mailboxes), alongside the standard sender / recipient / subject / free-text / date / attachment / unread / pinned filters. When the server reports a total match count, results are wrapped in a {"total", "items"} JSON envelope; otherwise a bare JSON array is returned.',
        inputSchema: {
          type: 'object',
          properties: {
            query: {
              type: 'string',
              description: 'Text to search for in subject/body',
            },
            from: {
              type: 'string',
              description: 'Filter by sender email',
            },
            to: {
              type: 'string',
              description: 'Filter by recipient email',
            },
            subject: {
              type: 'string',
              description: 'Filter by subject',
            },
            hasAttachment: {
              type: 'boolean',
              description: 'Filter emails with attachments',
            },
            isUnread: {
              type: 'boolean',
              description: 'Filter unread emails',
            },
            isPinned: {
              type: 'boolean',
              description: 'Filter pinned emails',
            },
            mailboxId: {
              type: 'string',
              description: 'Search within a single mailbox. For an intersection across multiple mailboxes (e.g. Inbox AND a label folder), use requiredMailboxIds instead.',
            },
            requiredMailboxIds: {
              type: 'array',
              items: { type: 'string' },
              description: 'Require membership in ALL of these mailbox IDs (intersection / AND semantic). Use this for queries like "in Inbox AND a label folder" — pass both mailbox IDs in the array. If mailboxId is also passed, it is folded into the intersection (de-duplicated). JMAP cannot express multi-mailbox membership in a single FilterCondition, so this builds a FilterOperator AND over multiple inMailbox conditions on the server.',
            },
            excludeMailboxIds: {
              type: 'array',
              items: { type: 'string' },
              description: 'Exclude emails that are members of ANY of these mailbox IDs (maps to JMAP inMailboxOtherThan). Useful for queries like "in a parent label but not its archive sub-folder". Combines cleanly with mailboxId / requiredMailboxIds.',
            },
            after: {
              type: 'string',
              description: 'Emails after this date (ISO 8601)',
            },
            before: {
              type: 'string',
              description: 'Emails before this date (ISO 8601)',
            },
            limit: {
              type: ['number', 'string'],
              description: 'Maximum results (default: 50)',
              default: 50,
            },
            ascending: {
              type: 'boolean',
              description: 'Sort oldest first instead of newest first (default: false)',
            },
          },
        },
      },
      {
        name: 'advanced_search_metadata',
        description: 'Same filter capabilities as advanced_search (single-mailbox scoping via mailboxId, multi-mailbox intersection via requiredMailboxIds, exclusion via excludeMailboxIds, plus sender / recipient / subject / free text / date / attachment / unread / pinned) but returns ONLY metadata on each match — id, threadId, subject, from, to, cc, replyTo, receivedAt, hasAttachment, keywords. Does NOT return preview or any body-derived content. Use in privacy-sensitive flows where the routing decision is made from headers alone — for example, when classifying customer mail by sender / recipient / subject / thread state without ingesting body content. The free-text query still searches body content on the server side; only the result envelope comes back without body excerpts. When the server reports a total match count, results are wrapped in a {"total", "items"} JSON envelope; otherwise a bare JSON array is returned.',
        inputSchema: {
          type: 'object',
          properties: {
            query: {
              type: 'string',
              description: 'Text to search for in subject/body',
            },
            from: {
              type: 'string',
              description: 'Filter by sender email',
            },
            to: {
              type: 'string',
              description: 'Filter by recipient email',
            },
            subject: {
              type: 'string',
              description: 'Filter by subject',
            },
            hasAttachment: {
              type: 'boolean',
              description: 'Filter emails with attachments',
            },
            isUnread: {
              type: 'boolean',
              description: 'Filter unread emails',
            },
            isPinned: {
              type: 'boolean',
              description: 'Filter pinned emails',
            },
            mailboxId: {
              type: 'string',
              description: 'Search within a single mailbox. For an intersection across multiple mailboxes (e.g. Inbox AND a label folder), use requiredMailboxIds instead.',
            },
            requiredMailboxIds: {
              type: 'array',
              items: { type: 'string' },
              description: 'Require membership in ALL of these mailbox IDs (intersection / AND semantic). Use this for queries like "in Inbox AND a label folder" — pass both mailbox IDs in the array. If mailboxId is also passed, it is folded into the intersection (de-duplicated). JMAP cannot express multi-mailbox membership in a single FilterCondition, so this builds a FilterOperator AND over multiple inMailbox conditions on the server.',
            },
            excludeMailboxIds: {
              type: 'array',
              items: { type: 'string' },
              description: 'Exclude emails that are members of ANY of these mailbox IDs (maps to JMAP inMailboxOtherThan). Useful for queries like "in a parent label but not its archive sub-folder". Combines cleanly with mailboxId / requiredMailboxIds.',
            },
            after: {
              type: 'string',
              description: 'Emails after this date (ISO 8601)',
            },
            before: {
              type: 'string',
              description: 'Emails before this date (ISO 8601)',
            },
            limit: {
              type: ['number', 'string'],
              description: 'Maximum results (default: 50)',
              default: 50,
            },
            ascending: {
              type: 'boolean',
              description: 'Sort oldest first instead of newest first (default: false)',
            },
          },
        },
      },
      {
        name: 'get_thread',
        description: 'Get all emails in a conversation thread. Draft messages are excluded by default; set includeDrafts=true to include in-progress drafts in the thread.',
        inputSchema: {
          type: 'object',
          properties: {
            threadId: {
              type: 'string',
              description: 'ID of the thread/conversation',
            },
            includeDrafts: {
              type: 'boolean',
              description: 'Include draft messages in the thread (default: false, drafts excluded).',
            },
          },
          required: ['threadId'],
        },
      },
      {
        name: 'get_thread_metadata',
        description: 'Same as get_thread (enumerate every message in a conversation thread) but returns ONLY metadata on each thread message — id, threadId, subject, from, to, cc, replyTo, receivedAt, hasAttachment, keywords. Does NOT return preview or any body-derived content. Use for thread-state checks (reply-presence detection, sender enumeration, date comparison, read/flagged status) without ingesting message bodies — particularly in customer-mail least-privilege flows where the skill needs to know "did we reply, when, and from which alias" but is forbidden from reading what was said. Accepts either a thread ID or an email ID and resolves to the parent thread, mirroring get_thread.',
        inputSchema: {
          type: 'object',
          properties: {
            threadId: {
              type: 'string',
              description: 'ID of the thread/conversation (an email ID is also accepted and will be resolved to its parent thread)',
            },
          },
          required: ['threadId'],
        },
      },
      {
        name: 'get_mailbox_stats',
        description: 'Get statistics for a mailbox (unread count, total emails, etc.)',
        inputSchema: {
          type: 'object',
          properties: {
            mailboxId: {
              type: 'string',
              description: 'ID of the mailbox (optional, defaults to all mailboxes)',
            },
          },
        },
      },
      {
        name: 'get_account_summary',
        description: 'Get overall account summary with statistics',
        inputSchema: {
          type: 'object',
          properties: {},
        },
      },
      {
        name: 'bulk_mark_read',
        description: 'Mark multiple emails as read/unread',
        inputSchema: {
          type: 'object',
          properties: {
            emailIds: {
              type: 'array',
              items: { type: 'string' },
              description: 'Array of email IDs to mark',
            },
            read: {
              type: 'boolean',
              description: 'true to mark as read, false as unread',
              default: true,
            },
          },
          required: ['emailIds'],
        },
      },
      {
        name: 'bulk_pin',
        description: 'Pin or unpin multiple emails',
        inputSchema: {
          type: 'object',
          properties: {
            emailIds: {
              type: 'array',
              items: { type: 'string' },
              description: 'Array of email IDs to pin/unpin',
            },
            pinned: {
              type: 'boolean',
              description: 'true to pin, false to unpin',
              default: true,
            },
          },
          required: ['emailIds'],
        },
      },
      {
        name: 'bulk_move',
        description: 'Move multiple emails to a mailbox',
        inputSchema: {
          type: 'object',
          properties: {
            emailIds: {
              type: 'array',
              items: { type: 'string' },
              description: 'Array of email IDs to move',
            },
            targetMailboxId: {
              type: 'string',
              description: 'ID of target mailbox',
            },
          },
          required: ['emailIds', 'targetMailboxId'],
        },
      },
      {
        name: 'bulk_delete',
        description: 'Delete multiple emails (move to trash)',
        inputSchema: {
          type: 'object',
          properties: {
            emailIds: {
              type: 'array',
              items: { type: 'string' },
              description: 'Array of email IDs to delete',
            },
          },
          required: ['emailIds'],
        },
      },
      {
        name: 'bulk_add_labels',
        description: 'Add labels to multiple emails simultaneously',
        inputSchema: {
          type: 'object',
          properties: {
            emailIds: {
              type: 'array',
              items: { type: 'string' },
              description: 'Array of email IDs to add labels to',
            },
            mailboxIds: {
              type: 'array',
              items: { type: 'string' },
              description: 'Array of mailbox IDs to add as labels',
            },
          },
          required: ['emailIds', 'mailboxIds'],
        },
      },
      {
        name: 'bulk_remove_labels',
        description: 'Remove labels from multiple emails simultaneously',
        inputSchema: {
          type: 'object',
          properties: {
            emailIds: {
              type: 'array',
              items: { type: 'string' },
              description: 'Array of email IDs to remove labels from',
            },
            mailboxIds: {
              type: 'array',
              items: { type: 'string' },
              description: 'Array of mailbox IDs to remove as labels',
            },
          },
          required: ['emailIds', 'mailboxIds'],
        },
      },
      {
        name: 'check_function_availability',
        description: 'Check which MCP functions are available based on account permissions. Calendar tools run over CalDAV, so calendar is reported available when CalDAV credentials are configured, regardless of the JMAP calendar capability.',
        inputSchema: {
          type: 'object',
          properties: {},
        },
      },
      {
        name: 'test_bulk_operations',
        description: 'Test bulk operations by finding recent emails and performing safe operations (mark read/unread)',
        inputSchema: {
          type: 'object',
          properties: {
            dryRun: {
              type: 'boolean',
              description: 'If true, only shows what would be done without making changes (default: true)',
              default: true,
            },
            limit: {
              type: 'number',
              description: 'Number of emails to test with (default: 3, max: 10)',
              default: 3,
            },
          },
        },
      },
    ],
  };
});

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  try {

    const client = initializeClient();

    switch (name) {
      case 'list_mailboxes': {
        const { properties, parentId } = (args ?? {}) as any;
        const options: { properties?: string[]; parentId?: string | null } = {};
        if (Array.isArray(properties) && properties.length > 0) {
          options.properties = properties;
        }
        if (args && Object.prototype.hasOwnProperty.call(args, 'parentId')) {
          options.parentId = parentId ?? null;
        }
        const mailboxes = await client.getMailboxes(options);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(mailboxes, null, 2),
            },
          ],
        };
      }

      case 'get_mailbox_by_name': {
        const { path } = (args ?? {}) as any;
        if (!path || typeof path !== 'string') {
          throw new McpError(ErrorCode.InvalidParams, 'path is required and must be a non-empty string');
        }
        const mailbox = await client.getMailboxByName(path);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(mailbox, null, 2),
            },
          ],
        };
      }

      case 'create_mailbox': {
        const { name: mailboxName, parentId } = (args ?? {}) as any;
        if (!mailboxName || typeof mailboxName !== 'string') {
          throw new McpError(ErrorCode.InvalidParams, 'name is required and must be a non-empty string');
        }
        if (mailboxName.includes('/')) {
          throw new McpError(ErrorCode.InvalidParams, 'name must not contain "/" — pass a leaf name and use parentId to nest');
        }
        const newId = await client.createMailbox(mailboxName, parentId ?? null);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ id: newId, name: mailboxName, parentId: parentId ?? null }, null, 2),
            },
          ],
        };
      }

      case 'list_emails': {
        const { mailboxId, limit, ascending } = args as any;
        const validLimit = Math.min(Math.max(Number(limit) || 20, 1), 100);
        const result = await client.getEmails(mailboxId, validLimit, !!ascending);
        return {
          content: [
            {
              type: 'text',
              text: formatQueryResult(result),
            },
          ],
        };
      }

      case 'list_emails_metadata': {
        const { mailboxId, limit, ascending } = (args ?? {}) as any;
        const validLimit = Math.min(Math.max(Number(limit) || 20, 1), 100);
        const result = await client.getEmailsMetadata(mailboxId, validLimit, !!ascending);
        return {
          content: [
            {
              type: 'text',
              text: formatQueryResult(result),
            },
          ],
        };
      }

      case 'get_email': {
        const { emailId } = args as any;
        if (!emailId) {
          throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
        }
        const email = await client.getEmailById(emailId);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(email, null, 2),
            },
          ],
        };
      }

      case 'get_email_metadata': {
        const { emailId } = (args ?? {}) as any;
        if (!emailId) {
          throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
        }
        const email = await client.getEmailMetadata(emailId);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(email, null, 2),
            },
          ],
        };
      }

      case 'send_email': {
        const { from, mailboxId, subject, textBody, htmlBody, inReplyTo, references } = args as any;
        const { to: toArray, cc, bcc, replyTo } = coerceRecipients(args as any);
        if (!toArray || toArray.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'to field is required and must be a non-empty array');
        }
        if (!subject) {
          throw new McpError(ErrorCode.InvalidParams, 'subject is required');
        }
        if (!textBody && !htmlBody) {
          throw new McpError(ErrorCode.InvalidParams, 'Either textBody or htmlBody is required');
        }

        const submissionId = await client.sendEmail({
          to: toArray,
          cc,
          bcc,
          from,
          mailboxId,
          subject,
          textBody,
          htmlBody,
          // Coerce so a lenient client's stringified array doesn't reach the JMAP
          // Email object as a bare string (consistent with the recipient fields).
          inReplyTo: coerceStringArray(inReplyTo),
          references: coerceStringArray(references),
          replyTo,
          attachments: (args as any).attachments,
          downloadDir: getDownloadDir(),
        });

        return {
          content: [
            {
              type: 'text',
              text: `Email sent successfully. Submission ID: ${submissionId}`,
            },
          ],
        };
      }

      case 'reply_email': {
        const { originalEmailId, from, textBody, htmlBody, send } = args as any;
        const { to: toArray, cc, bcc, replyTo } = coerceRecipients(args as any);
        const shouldSend = coerceBool(send) ?? true;
        if (!originalEmailId) {
          throw new McpError(ErrorCode.InvalidParams, 'originalEmailId is required');
        }
        if (shouldSend && !textBody && !htmlBody) {
          throw new McpError(ErrorCode.InvalidParams, 'Either textBody or htmlBody is required');
        }

        // Fetch the original email to get threading headers
        const originalEmail = await client.getEmailById(originalEmailId);

        // Build threading headers
        const originalMessageId = originalEmail.messageId?.[0];
        if (!originalMessageId) {
          throw new McpError(ErrorCode.InternalError, 'Original email does not have a Message-ID; cannot thread reply');
        }

        const inReplyToHeader = [originalMessageId];
        const referencesHeader = [
          ...(originalEmail.references || []),
          originalMessageId,
        ];

        // Build subject with Re: prefix
        let replySubject = originalEmail.subject || '';
        if (!/^Re:/i.test(replySubject)) {
          replySubject = `Re: ${replySubject}`;
        }

        // Default recipients to the original sender
        const replyRecipients = (toArray && toArray.length > 0)
          ? toArray
          : (Array.isArray(originalEmail.from) ? originalEmail.from.map((addr: any) => addr.email).filter(Boolean) : []);

        if (replyRecipients.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'Could not determine reply recipient. Please provide "to" explicitly.');
        }

        const replyParams = {
          to: replyRecipients,
          cc,
          bcc,
          from,
          subject: replySubject,
          textBody,
          htmlBody,
          inReplyTo: inReplyToHeader,
          references: referencesHeader,
          replyTo,
          attachments: (args as any).attachments,
          downloadDir: getDownloadDir(),
        };

        if (!shouldSend) {
          const emailId = await client.createDraft(replyParams);
          return {
            content: [
              {
                type: 'text',
                text: `Reply draft saved successfully (Email ID: ${emailId}). Subject: ${replySubject}`,
              },
            ],
          };
        }

        const submissionId = await client.sendEmail(replyParams);

        return {
          content: [
            {
              type: 'text',
              text: `Reply sent successfully. Submission ID: ${submissionId}`,
            },
          ],
        };
      }

      case 'create_draft': {
        const { from, mailboxId, subject, textBody, htmlBody, inReplyTo, references } = args as any;
        const { to, cc, bcc, replyTo } = coerceRecipients(args as any);

        if (!to?.length && !subject && !textBody && !htmlBody) {
          throw new McpError(ErrorCode.InvalidParams, 'At least one of to, subject, textBody, or htmlBody must be provided');
        }

        const emailId = await client.createDraft({
          to,
          cc,
          bcc,
          from,
          mailboxId,
          subject,
          textBody,
          htmlBody,
          inReplyTo: coerceStringArray(inReplyTo),
          references: coerceStringArray(references),
          replyTo,
          attachments: (args as any).attachments,
          downloadDir: getDownloadDir(),
        });

        const summary = [
          `Draft created successfully (Email ID: ${emailId}).`,
          subject ? `Subject: ${subject}` : null,
          to?.length ? `To: ${to.join(', ')}` : null,
          cc?.length ? `CC: ${cc.join(', ')}` : null,
        ].filter(Boolean).join(' ');

        return {
          content: [
            {
              type: 'text',
              text: summary,
            },
          ],
        };
      }

      case 'edit_draft': {
        const { emailId, from, subject, textBody, htmlBody } = args as any;
        const { to, cc, bcc, replyTo } = coerceRecipients(args as any);
        if (!emailId) {
          throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
        }

        const newEmailId = await client.updateDraft(emailId, {
          to,
          cc,
          bcc,
          from,
          subject,
          textBody,
          htmlBody,
          replyTo,
          attachments: (args as any).attachments,
          downloadDir: getDownloadDir(),
        });

        return {
          content: [
            {
              type: 'text',
              text: `Draft updated successfully. New Email ID: ${newEmailId} (old draft ${emailId} was replaced)`,
            },
          ],
        };
      }

      case 'send_draft': {
        const { emailId } = args as any;
        if (!emailId) {
          throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
        }

        const submissionId = await client.sendDraft(emailId);

        return {
          content: [
            {
              type: 'text',
              text: `Draft sent successfully. Submission ID: ${submissionId}`,
            },
          ],
        };
      }

      case 'search_emails': {
        const { query, limit, ascending, excludeDrafts } = args as any;
        if (!query) {
          throw new McpError(ErrorCode.InvalidParams, 'query is required');
        }
        const validLimit = Math.min(Math.max(Number(limit) || 20, 1), 100);
        const result = await client.searchEmails(query, validLimit, !!ascending, coerceBool(excludeDrafts) ?? false);
        return {
          content: [
            {
              type: 'text',
              text: formatQueryResult(result),
            },
          ],
        };
      }

      case 'search_emails_metadata': {
        const { query, limit, ascending } = (args ?? {}) as any;
        if (!query) {
          throw new McpError(ErrorCode.InvalidParams, 'query is required');
        }
        const validLimit = Math.min(Math.max(Number(limit) || 20, 1), 100);
        const result = await client.searchEmailsMetadata(query, validLimit, !!ascending);
        return {
          content: [
            {
              type: 'text',
              text: formatQueryResult(result),
            },
          ],
        };
      }

      case 'list_contacts': {
        const { limit } = args as any;
        const contactsClient = initializeContactsCalendarClient();
        const result = await contactsClient.getContacts(clampLimit(limit, 50, 200));
        return {
          content: [
            {
              type: 'text',
              text: formatQueryResult(result),
            },
          ],
        };
      }

      case 'get_contact': {
        const { contactId } = args as any;
        if (!contactId) {
          throw new McpError(ErrorCode.InvalidParams, 'contactId is required');
        }
        const contactsClient = initializeContactsCalendarClient();
        const contact = await contactsClient.getContactById(contactId);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(contact, null, 2),
            },
          ],
        };
      }

      case 'search_contacts': {
        const { query, limit } = args as any;
        if (!query) {
          throw new McpError(ErrorCode.InvalidParams, 'query is required');
        }
        const contactsClient = initializeContactsCalendarClient();
        const result = await contactsClient.searchContacts(query, clampLimit(limit, 20, 100));
        return {
          content: [
            {
              type: 'text',
              text: formatQueryResult(result),
            },
          ],
        };
      }

      // Calendar operations use CalDAV directly.
      // JMAP Calendars: spec not yet finalized, Fastmail has not enabled JMAP calendar support.
      // Existing JMAP calendar code in contacts-calendar.ts has known bugs and must not be used.
      // When Fastmail enables JMAP calendars: re-enable the path, fix to match finalized spec,
      // do a parity pass with CalDAV implementation, and test against live Fastmail.
      // CalDAV tests should be structured so they can serve as a basis for JMAP tests later.

      case 'create_contact': {
        const { name, emails, phones, addresses, notes, addressBookId } = args as any;
        const contactsClient = initializeContactsCalendarClient();
        const contactId = await contactsClient.createContact({ name, emails, phones, addresses, notes, addressBookId });
        return {
          content: [
            {
              type: 'text',
              text: `Contact created successfully. Contact ID: ${contactId}`,
            },
          ],
        };
      }

      case 'update_contact': {
        const { contactId, name, emails, phones, addresses, notes, expectState } = args as any;
        if (!contactId) {
          throw new McpError(ErrorCode.InvalidParams, 'contactId is required');
        }
        const contactsClient = initializeContactsCalendarClient();
        await contactsClient.updateContact(contactId, { name, emails, phones, addresses, notes, expectState });
        return {
          content: [
            {
              type: 'text',
              text: `Contact updated successfully. Contact ID: ${contactId}`,
            },
          ],
        };
      }

      case 'delete_contact': {
        const { contactId, expectState } = args as any;
        if (!contactId) {
          throw new McpError(ErrorCode.InvalidParams, 'contactId is required');
        }
        const contactsClient = initializeContactsCalendarClient();
        await contactsClient.deleteContact(contactId, expectState);
        return {
          content: [
            {
              type: 'text',
              text: `Contact deleted. Contact ID: ${contactId}`,
            },
          ],
        };
      }

      case 'list_calendars': {
        const davClient = initializeCalDAVClient();
        if (!davClient) {
          throw new McpError(ErrorCode.InvalidRequest, 'CalDAV not configured. Set FASTMAIL_CALDAV_USERNAME and FASTMAIL_CALDAV_PASSWORD.');
        }
        const calendars = await davClient.getCalendars();
        return { content: [{ type: 'text', text: JSON.stringify(calendars, null, 2) }] };
      }

      case 'list_calendar_events': {
        const { calendarId, limit, startDate, endDate } = args as any;
        // JMAP Calendars: disabled — spec not yet finalized, Fastmail has not enabled support.
        // Existing JMAP calendar code in contacts-calendar.ts has known bugs.
        // When Fastmail enables JMAP calendars: re-enable, fix to match finalized spec,
        // do a parity pass with CalDAV implementation, and test against live Fastmail.
        const davClient = initializeCalDAVClient();
        if (!davClient) {
          throw new McpError(ErrorCode.InvalidRequest, 'CalDAV not configured. Set FASTMAIL_CALDAV_USERNAME and FASTMAIL_CALDAV_PASSWORD.');
        }
        const events = await davClient.getCalendarEvents(calendarId, clampLimit(limit, 50, 500), startDate, endDate);
        return { content: [{ type: 'text', text: JSON.stringify(events, null, 2) }] };
      }

      case 'get_calendar_event': {
        const { eventId } = args as any;
        if (!eventId) {
          throw new McpError(ErrorCode.InvalidParams, 'eventId is required');
        }
        const davClient = initializeCalDAVClient();
        if (!davClient) {
          throw new McpError(ErrorCode.InvalidRequest, 'CalDAV not configured. Set FASTMAIL_CALDAV_USERNAME and FASTMAIL_CALDAV_PASSWORD.');
        }
        const event = await davClient.getCalendarEventById(eventId);
        return { content: [{ type: 'text', text: JSON.stringify(event, null, 2) }] };
      }

      case 'create_calendar_event': {
        const { calendarId, title, description, start, end, location, participants } = args as any;
        if (!calendarId || !title || !start || !end) {
          throw new McpError(ErrorCode.InvalidParams, 'calendarId, title, start, and end are required');
        }
        const davClient = initializeCalDAVClient();
        if (!davClient) {
          throw new McpError(ErrorCode.InvalidRequest, 'CalDAV not configured. Set FASTMAIL_CALDAV_USERNAME and FASTMAIL_CALDAV_PASSWORD.');
        }
        const eventId = await davClient.createCalendarEvent({
          calendarId, title, description, start, end, location, participants,
        });
        return { content: [{ type: 'text', text: `Calendar event created. Event ID: ${eventId}` }] };
      }

      case 'update_calendar_event': {
        const { eventId, title, description, start, end, location, participants, clearFields, confirmRecurring } = args as any;
        if (!eventId) {
          throw new McpError(ErrorCode.InvalidParams, 'eventId is required');
        }
        const hasClearFields = Array.isArray(clearFields) && clearFields.length > 0;
        if (title === undefined && description === undefined && start === undefined && end === undefined && location === undefined && participants === undefined && !hasClearFields) {
          throw new McpError(ErrorCode.InvalidParams, 'At least one field to update must be provided (title, description, start, end, location, participants, or clearFields)');
        }
        const davClient = initializeCalDAVClient();
        if (!davClient) {
          throw new McpError(ErrorCode.InvalidRequest, 'CalDAV not configured. Set FASTMAIL_CALDAV_USERNAME and FASTMAIL_CALDAV_PASSWORD.');
        }
        // Coerce confirmRecurring — a lenient client sending the string "false"
        // would otherwise read as truthy and authorize destructive pruning of
        // recurrence exceptions the caller explicitly declined to remove.
        const fields = {
          title, description, start, end, location, participants, clearFields,
          confirmRecurring: coerceBool(confirmRecurring) ?? false,
        };
        await davClient.updateCalendarEvent(eventId, fields);
        return { content: [{ type: 'text', text: `Calendar event updated. Event ID: ${eventId}` }] };
      }

      case 'delete_calendar_event': {
        const { eventId } = args as any;
        if (!eventId) {
          throw new McpError(ErrorCode.InvalidParams, 'eventId is required');
        }
        const davClient = initializeCalDAVClient();
        if (!davClient) {
          throw new McpError(ErrorCode.InvalidRequest, 'CalDAV not configured. Set FASTMAIL_CALDAV_USERNAME and FASTMAIL_CALDAV_PASSWORD.');
        }
        await davClient.deleteCalendarEvent(eventId);
        return { content: [{ type: 'text', text: `Calendar event deleted. Event ID: ${eventId}` }] };
      }

      case 'list_identities': {
        const client = initializeClient();
        const identities = await client.getIdentities();
        
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(identities, null, 2),
            },
          ],
        };
      }

      case 'get_recent_emails': {
        const { limit, mailboxName = null, ascending } = args as any;
        const client = initializeClient();
        const result = await client.getRecentEmails(clampLimit(limit, 10, 50), mailboxName, coerceBool(ascending) ?? false);
        return {
          content: [
            {
              type: 'text',
              text: formatQueryResult(result),
            },
          ],
        };
      }

      case 'mark_email_read': {
        const { emailId } = args as any;
        // Coerce so a stringified "false" from a lenient client doesn't read as truthy.
        const read = coerceBool((args as any).read) ?? true;
        if (!emailId) {
          throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
        }
        const client = initializeClient();
        await client.markEmailRead(emailId, read);
        return {
          content: [
            {
              type: 'text',
              text: `Email ${read ? 'marked as read' : 'marked as unread'} successfully`,
            },
          ],
        };
      }

      case 'pin_email': {
        const { emailId } = args as any;
        const pinned = coerceBool((args as any).pinned) ?? true;
        if (!emailId) {
          throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
        }
        const client = initializeClient();
        await client.pinEmail(emailId, pinned);
        return {
          content: [
            {
              type: 'text',
              text: `Email ${pinned ? 'pinned' : 'unpinned'} successfully`,
            },
          ],
        };
      }

      case 'delete_email': {
        const { emailId } = args as any;
        if (!emailId) {
          throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
        }
        const client = initializeClient();
        await client.deleteEmail(emailId);
        return {
          content: [
            {
              type: 'text',
              text: 'Email deleted successfully (moved to trash)',
            },
          ],
        };
      }

      case 'move_email': {
        const { emailId, targetMailboxId } = args as any;
        if (!emailId || !targetMailboxId) {
          throw new McpError(ErrorCode.InvalidParams, 'emailId and targetMailboxId are required');
        }
        const client = initializeClient();
        await client.moveEmail(emailId, targetMailboxId);
        return {
          content: [
            {
              type: 'text',
              text: 'Email moved successfully',
            },
          ],
        };
      }

      case 'archive_email': {
        const { emailId, targetMailboxId } = (args ?? {}) as any;
        if (!emailId || !targetMailboxId) {
          throw new McpError(ErrorCode.InvalidParams, 'emailId and targetMailboxId are required');
        }
        const client = initializeClient();
        await client.archiveEmail(emailId, targetMailboxId);
        return {
          content: [
            {
              type: 'text',
              text: 'Email archived successfully (moved to target mailbox and marked as read)',
            },
          ],
        };
      }

      case 'add_labels': {
        const { emailId, mailboxIds } = args as any;
        if (!emailId) {
          throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
        }
        if (!mailboxIds || !Array.isArray(mailboxIds) || mailboxIds.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'mailboxIds array is required and must not be empty');
        }
        const client = initializeClient();
        await client.addLabels(emailId, mailboxIds);
        return {
          content: [
            {
              type: 'text',
              text: `Labels added successfully to email`,
            },
          ],
        };
      }

      case 'remove_labels': {
        const { emailId, mailboxIds } = args as any;
        if (!emailId) {
          throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
        }
        if (!mailboxIds || !Array.isArray(mailboxIds) || mailboxIds.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'mailboxIds array is required and must not be empty');
        }
        const client = initializeClient();
        await client.removeLabels(emailId, mailboxIds);
        return {
          content: [
            {
              type: 'text',
              text: `Labels removed successfully from email`,
            },
          ],
        };
      }

      case 'get_email_attachments': {
        const { emailId } = args as any;
        if (!emailId) {
          throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
        }
        const client = initializeClient();
        const attachments = await client.getEmailAttachments(emailId);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(attachments, null, 2),
            },
          ],
        };
      }

      case 'download_attachment': {
        const { emailId, attachmentId, savePath } = args as any;
        if (!emailId || !attachmentId) {
          throw new McpError(ErrorCode.InvalidParams, 'emailId and attachmentId are required');
        }
        const client = initializeClient();
        try {
          if (savePath) {
            const result = await client.downloadAttachmentToFile(emailId, attachmentId, savePath, getDownloadDir());
            return {
              content: [
                {
                  type: 'text',
                  text: `Saved to: ${result.savedPath} (${result.bytesWritten} bytes)`,
                },
              ],
            };
          } else {
            const downloadUrl = await client.downloadAttachment(emailId, attachmentId);
            return {
              content: [
                {
                  type: 'text',
                  text: `Download URL: ${downloadUrl}`,
                },
              ],
            };
          }
        } catch (error) {
          // Let path validation errors through so users see why their savePath was rejected.
          // Redact defensively — the message echoes a caller-influenced path, and a future
          // upstream error merely containing "Save path" would otherwise pass through raw.
          if (error instanceof Error && (error.message.includes('Save path') || error.message.includes('null bytes'))) {
            throw new McpError(ErrorCode.InvalidParams, redactBearerTokens(error.message));
          }
          // Sanitize other errors to avoid leaking attachment metadata
          throw new McpError(
            ErrorCode.InternalError,
            'Attachment download failed. Verify emailId and attachmentId and try again.'
          );
        }
      }

      case 'save_attachment_to_webdav': {
        const { emailId, attachmentId, remotePath, overwrite, createParents } = args as any;
        if (!emailId || !attachmentId) {
          throw new McpError(ErrorCode.InvalidParams, 'emailId and attachmentId are required');
        }
        if (!remotePath) {
          throw new McpError(ErrorCode.InvalidParams, 'remotePath is required');
        }
        const dav = initializeWebDAVClient();
        if (!dav) {
          throw new McpError(ErrorCode.InvalidRequest, 'WebDAV storage not configured. Set FASTMAIL_WEBDAV_URL, FASTMAIL_WEBDAV_USERNAME, and FASTMAIL_WEBDAV_PASSWORD (for Fastmail Files use https://myfiles.fastmail.com/ with a Files-scoped app password).');
        }
        const client = initializeClient();
        const { buffer, type, name } = await client.fetchAttachmentBuffer(emailId, attachmentId);
        const result = await dav.uploadBuffer(buffer, remotePath, {
          contentType: type,
          overwrite: coerceBool(overwrite) ?? false,
          createParents: coerceBool(createParents) ?? true,
        });
        return {
          content: [
            {
              type: 'text',
              text: redactBearerTokens(JSON.stringify({ ...result, attachmentName: name }, null, 2)),
            },
          ],
        };
      }

      case 'advanced_search': {
        const { query, from, to, subject, hasAttachment, isUnread, isPinned, mailboxId, requiredMailboxIds, excludeMailboxIds, after, before, limit, ascending } = args as any;
        const client = initializeClient();
        const validLimit = Math.min(Math.max(Number(limit) || 50, 1), 100);
        const result = await client.advancedSearch({
          query, from, to, subject, hasAttachment, isUnread, isPinned, mailboxId, requiredMailboxIds, excludeMailboxIds, after, before, limit: validLimit, ascending
        });
        return {
          content: [
            {
              type: 'text',
              text: formatQueryResult(result),
            },
          ],
        };
      }

      case 'advanced_search_metadata': {
        const { query, from, to, subject, hasAttachment, isUnread, isPinned, mailboxId, requiredMailboxIds, excludeMailboxIds, after, before, limit, ascending } = (args ?? {}) as any;
        const client = initializeClient();
        const validLimit = Math.min(Math.max(Number(limit) || 50, 1), 100);
        const result = await client.advancedSearchMetadata({
          query, from, to, subject, hasAttachment, isUnread, isPinned, mailboxId, requiredMailboxIds, excludeMailboxIds, after, before, limit: validLimit, ascending
        });
        return {
          content: [
            {
              type: 'text',
              text: formatQueryResult(result),
            },
          ],
        };
      }

      case 'get_thread': {
        const { threadId, includeDrafts } = args as any;
        if (!threadId) {
          throw new McpError(ErrorCode.InvalidParams, 'threadId is required');
        }
        const client = initializeClient();
        try {
          const thread = await client.getThread(threadId, coerceBool(includeDrafts) ?? false);
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify(thread, null, 2),
              },
            ],
          };
        } catch (error) {
          // Provide helpful error information
          throw new McpError(ErrorCode.InternalError, `Thread access failed: ${redactBearerTokens(error instanceof Error ? error.message : String(error))}`);
        }
      }

      case 'get_thread_metadata': {
        const { threadId } = (args ?? {}) as any;
        if (!threadId) {
          throw new McpError(ErrorCode.InvalidParams, 'threadId is required');
        }
        const client = initializeClient();
        try {
          const thread = await client.getThreadMetadata(threadId);
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify(thread, null, 2),
              },
            ],
          };
        } catch (error) {
          throw new McpError(ErrorCode.InternalError, `Thread access failed: ${redactBearerTokens(error instanceof Error ? error.message : String(error))}`);
        }
      }

      case 'get_mailbox_stats': {
        const { mailboxId } = args as any;
        const client = initializeClient();
        const stats = await client.getMailboxStats(mailboxId);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(stats, null, 2),
            },
          ],
        };
      }

      case 'get_account_summary': {
        const client = initializeClient();
        const summary = await client.getAccountSummary();
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(summary, null, 2),
            },
          ],
        };
      }

      case 'bulk_mark_read': {
        const { emailIds, read = true } = args as any;
        if (!emailIds || !Array.isArray(emailIds) || emailIds.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'emailIds array is required and must not be empty');
        }
        const client = initializeClient();
        await client.bulkMarkRead(emailIds, read);
        return {
          content: [
            {
              type: 'text',
              text: `${emailIds.length} emails ${read ? 'marked as read' : 'marked as unread'} successfully`,
            },
          ],
        };
      }

      case 'bulk_pin': {
        const { emailIds, pinned = true } = args as any;
        if (!emailIds || !Array.isArray(emailIds) || emailIds.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'emailIds array is required and must not be empty');
        }
        const client = initializeClient();
        await client.bulkPinEmails(emailIds, pinned);
        return {
          content: [
            {
              type: 'text',
              text: `${emailIds.length} emails ${pinned ? 'pinned' : 'unpinned'} successfully`,
            },
          ],
        };
      }

      case 'bulk_move': {
        const { emailIds, targetMailboxId } = args as any;
        if (!emailIds || !Array.isArray(emailIds) || emailIds.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'emailIds array is required and must not be empty');
        }
        if (!targetMailboxId) {
          throw new McpError(ErrorCode.InvalidParams, 'targetMailboxId is required');
        }
        const client = initializeClient();
        await client.bulkMove(emailIds, targetMailboxId);
        return {
          content: [
            {
              type: 'text',
              text: `${emailIds.length} emails moved successfully`,
            },
          ],
        };
      }

      case 'bulk_delete': {
        const { emailIds } = args as any;
        if (!emailIds || !Array.isArray(emailIds) || emailIds.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'emailIds array is required and must not be empty');
        }
        const client = initializeClient();
        await client.bulkDelete(emailIds);
        return {
          content: [
            {
              type: 'text',
              text: `${emailIds.length} emails deleted successfully (moved to trash)`,
            },
          ],
        };
      }

      case 'bulk_add_labels': {
        const { emailIds, mailboxIds } = args as any;
        if (!emailIds || !Array.isArray(emailIds) || emailIds.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'emailIds array is required and must not be empty');
        }
        if (!mailboxIds || !Array.isArray(mailboxIds) || mailboxIds.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'mailboxIds array is required and must not be empty');
        }
        const client = initializeClient();
        await client.bulkAddLabels(emailIds, mailboxIds);
        return {
          content: [
            {
              type: 'text',
              text: `Labels added successfully to ${emailIds.length} emails`,
            },
          ],
        };
      }

      case 'bulk_remove_labels': {
        const { emailIds, mailboxIds } = args as any;
        if (!emailIds || !Array.isArray(emailIds) || emailIds.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'emailIds array is required and must not be empty');
        }
        if (!mailboxIds || !Array.isArray(mailboxIds) || mailboxIds.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'mailboxIds array is required and must not be empty');
        }
        const client = initializeClient();
        await client.bulkRemoveLabels(emailIds, mailboxIds);
        return {
          content: [
            {
              type: 'text',
              text: `Labels removed successfully from ${emailIds.length} emails`,
            },
          ],
        };
      }

      case 'check_function_availability': {
        const client = initializeClient();
        const session = await client.getSession();

        // Calendar tools run on CalDAV, not JMAP. So calendar is available if
        // EITHER the JMAP calendar capability is present OR CalDAV credentials
        // are configured (FASTMAIL_CALDAV_USERNAME / FASTMAIL_CALDAV_PASSWORD).
        const jmapCalendar = !!session.capabilities['urn:ietf:params:jmap:calendars'];
        const caldavConfigured = initializeCalDAVClient() !== null;
        const calendarAvailable = jmapCalendar || caldavConfigured;
        const calendarNote = jmapCalendar
          ? 'Calendar is available (JMAP)'
          : caldavConfigured
            ? 'Calendar is available via CalDAV'
            : 'Calendar access not available - set FASTMAIL_CALDAV_USERNAME and FASTMAIL_CALDAV_PASSWORD, or enable calendar scope in Fastmail account settings';

        const availability = {
          email: {
            available: true,
            functions: [
              'list_mailboxes', 'get_mailbox_by_name', 'create_mailbox', 'list_emails', 'list_emails_metadata', 'get_email', 'get_email_metadata', 'send_email',
              'create_draft', 'edit_draft', 'send_draft', 'search_emails', 'search_emails_metadata',
              'get_recent_emails', 'mark_email_read', 'pin_email', 'delete_email', 'move_email', 'archive_email',
              'get_email_attachments', 'download_attachment', 'advanced_search', 'advanced_search_metadata', 'get_thread', 'get_thread_metadata',
              'get_mailbox_stats', 'get_account_summary', 'bulk_mark_read', 'bulk_pin', 'bulk_move', 'bulk_delete',
              'add_labels', 'remove_labels', 'bulk_add_labels', 'bulk_remove_labels'
            ]
          },
          identity: {
            available: true,
            functions: ['list_identities']
          },
          contacts: {
            available: !!session.capabilities['urn:ietf:params:jmap:contacts'],
            functions: ['list_contacts', 'get_contact', 'search_contacts', 'create_contact', 'update_contact', 'delete_contact'],
            note: session.capabilities['urn:ietf:params:jmap:contacts'] ?
              'Contacts are available. Write tools (create/update/delete) additionally require the API token to have read-write contacts scope.' :
              'Contacts access not available - may require enabling in Fastmail account settings',
            enablementGuide: session.capabilities['urn:ietf:params:jmap:contacts'] ? null : {
              steps: [
                '1. Log into Fastmail web interface',
                '2. Go to Settings → Privacy & Security → Connected Apps & API tokens',
                '3. Check if contacts scope is enabled for your API token',
                '4. If not available, you may need to upgrade your Fastmail plan or contact support'
              ],
              documentation: 'https://www.fastmail.com/help/technical/jmap-api.html'
            }
          },
          calendar: {
            available: calendarAvailable,
            functions: ['list_calendars', 'list_calendar_events', 'get_calendar_event', 'create_calendar_event'],
            note: calendarNote,
            enablementGuide: calendarAvailable ? null : {
              steps: [
                'Option A (CalDAV): set FASTMAIL_CALDAV_USERNAME and FASTMAIL_CALDAV_PASSWORD (app password) — calendar tools run over CalDAV',
                'Option B (JMAP scope): 1. Log into Fastmail web interface',
                '2. Go to Settings → Privacy & Security → Connected Apps & API tokens',
                '3. Check if calendar scope is enabled for your API token',
                '4. If not available, you may need to upgrade your Fastmail plan or contact support'
              ],
              documentation: 'https://www.fastmail.com/help/technical/jmap-api.html'
            }
          },
          capabilities: Object.keys(session.capabilities)
        };
        
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(availability, null, 2),
            },
          ],
        };
      }

      case 'test_bulk_operations': {
        const { dryRun = true, limit = 3 } = args as any;
        const client = initializeClient();
        
        // Get some recent emails to test with
        const testLimit = Math.min(Math.max(limit, 1), 10);
        const { items: emails } = await client.getRecentEmails(testLimit, 'inbox');

        if (emails.length === 0) {
          return {
            content: [
              {
                type: 'text',
                text: 'No emails found for bulk operation testing. Try sending yourself a test email first.',
              },
            ],
          };
        }
        
        const emailIds = emails.slice(0, testLimit).map(email => email.id);
        const operations = [
          {
            name: 'bulk_mark_read',
            description: `Mark ${emailIds.length} emails as read`,
            parameters: { emailIds, read: true }
          },
          {
            name: 'bulk_mark_read (undo)',
            description: `Mark ${emailIds.length} emails as unread (undo previous)`,
            parameters: { emailIds, read: false }
          }
        ];
        
        const results = {
          testEmails: emails.map(email => ({
            id: email.id,
            subject: email.subject,
            from: email.from?.[0]?.email || 'unknown',
            receivedAt: email.receivedAt
          })),
          operations: [] as any[]
        };
        
        if (dryRun) {
          results.operations = operations.map(op => ({
            ...op,
            status: 'DRY RUN - Would execute but not actually performed',
            executed: false
          }));
          
          return {
            content: [
              {
                type: 'text',
                text: `BULK OPERATIONS TEST (DRY RUN)\n\n${JSON.stringify(results, null, 2)}\n\nTo actually execute the test, set dryRun: false`,
              },
            ],
          };
        } else {
          // Execute the test operations
          for (const operation of operations) {
            try {
              await client.bulkMarkRead(operation.parameters.emailIds, operation.parameters.read);
              results.operations.push({
                ...operation,
                status: 'SUCCESS',
                executed: true,
                timestamp: new Date().toISOString()
              });
              
              // Small delay between operations
              await new Promise(resolve => setTimeout(resolve, 500));
            } catch (error) {
              results.operations.push({
                ...operation,
                status: 'FAILED',
                executed: false,
                error: redactBearerTokens(error instanceof Error ? error.message : String(error)),
                timestamp: new Date().toISOString()
              });
            }
          }
          
          return {
            content: [
              {
                type: 'text',
                text: `BULK OPERATIONS TEST (EXECUTED)\n\n${JSON.stringify(results, null, 2)}`,
              },
            ],
          };
        }
      }

      default:
        throw new McpError(ErrorCode.MethodNotFound, `Unknown tool: ${name}`);
    }
  } catch (error) {
    if (error instanceof McpError) {
      // Redact defensively even on the McpError path — in-file McpErrors are
      // static today, but this makes redaction a single choke point so a future
      // McpError built from dynamic content can't slip a secret through.
      const safe = redactBearerTokens(error.message);
      if (safe === error.message) throw error;
      throw new McpError((error as any).code ?? ErrorCode.InternalError, safe);
    }
    const raw = error instanceof Error ? error.message : String(error);
    throw new McpError(
      ErrorCode.InternalError,
      `Tool execution failed: ${redactBearerTokens(raw)}`
    );
  }
});

async function runServer() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error('Fastmail MCP server running on stdio');
}

runServer().catch(() => {
  // Avoid logging raw error objects to prevent accidental PII leakage
  console.error('Fastmail MCP server failed to start');
  process.exit(1);
});