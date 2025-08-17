import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { CallToolRequestSchema, ErrorCode, ListToolsRequestSchema, McpError } from '@modelcontextprotocol/sdk/types.js';
import { JmapClient, JmapRequest } from './jmap-client.js';
import { ContactsCalendarClient } from './contacts-calendar.js';

export interface HandlerProviders {
  getJmapClient: () => JmapClient;
  getContactsClient: () => ContactsCalendarClient;
}

export function registerHandlers(server: Server, providers: HandlerProviders) {
  server.setRequestHandler(ListToolsRequestSchema, async () => {
    return {
      tools: [
        {
          name: 'list_mailboxes',
          description: 'List all mailboxes in the Fastmail account',
          inputSchema: {
            type: 'object',
            properties: {},
          },
        },
        {
          name: 'list_emails',
          description: 'List emails from a mailbox',
          inputSchema: {
            type: 'object',
            properties: {
              mailboxId: {
                type: 'string',
                description: 'ID of the mailbox to list emails from (optional, defaults to all)',
              },
              limit: {
                type: 'number',
                description: 'Maximum number of emails to return (default: 20)',
                default: 20,
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
        },
        {
          name: 'send_email',
          description: 'Send an email',
          inputSchema: {
            type: 'object',
            properties: {
              to: {
                type: 'array',
                items: { type: 'string' },
                description: 'Recipient email addresses',
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
                description: 'Sender email address (optional, must be a verified identity)',
              },
              mailboxId: {
                type: 'string',
                description: 'Mailbox to draft the email in (optional)',
              },
              subject: { type: 'string' },
              textBody: { type: 'string' },
              htmlBody: { type: 'string' },
            },
            required: ['to', 'subject'],
          },
        },
        { name: 'search_emails', description: 'Search emails by content', inputSchema: { type: 'object', properties: { query: { type: 'string' }, limit: { type: 'number', default: 20 } }, required: ['query'] } },
        { name: 'get_recent_emails', description: 'Get the most recent emails from a mailbox', inputSchema: { type: 'object', properties: { limit: { type: 'number', default: 10 }, mailboxName: { type: 'string', default: 'inbox' } } } },
        { name: 'mark_email_read', description: 'Mark an email as read or unread', inputSchema: { type: 'object', properties: { emailId: { type: 'string' }, read: { type: 'boolean', default: true } }, required: ['emailId'] } },
        { name: 'delete_email', description: 'Delete an email (move to trash)', inputSchema: { type: 'object', properties: { emailId: { type: 'string' } }, required: ['emailId'] } },
        { name: 'move_email', description: 'Move an email to a different mailbox', inputSchema: { type: 'object', properties: { emailId: { type: 'string' }, targetMailboxId: { type: 'string' } }, required: ['emailId', 'targetMailboxId'] } },
        { name: 'get_email_attachments', description: 'Get list of attachments for an email', inputSchema: { type: 'object', properties: { emailId: { type: 'string' } }, required: ['emailId'] } },
        { name: 'download_attachment', description: 'Get download URL for an email attachment', inputSchema: { type: 'object', properties: { emailId: { type: 'string' }, attachmentId: { type: 'string' } }, required: ['emailId', 'attachmentId'] } },
        { name: 'advanced_search', description: 'Advanced email search with multiple criteria', inputSchema: { type: 'object', properties: { query: { type: 'string' }, from: { type: 'string' }, to: { type: 'string' }, subject: { type: 'string' }, hasAttachment: { type: 'boolean' }, isUnread: { type: 'boolean' }, mailboxId: { type: 'string' }, after: { type: 'string' }, before: { type: 'string' }, limit: { type: 'number', default: 50 } } } },
        { name: 'get_thread', description: 'Get all emails in a conversation thread', inputSchema: { type: 'object', properties: { threadId: { type: 'string' } }, required: ['threadId'] } },
        { name: 'get_mailbox_stats', description: 'Get statistics for a mailbox', inputSchema: { type: 'object', properties: { mailboxId: { type: 'string' } } } },
        { name: 'get_account_summary', description: 'Get overall account summary', inputSchema: { type: 'object', properties: {} } },
        { name: 'bulk_mark_read', description: 'Mark multiple emails read/unread', inputSchema: { type: 'object', properties: { emailIds: { type: 'array', items: { type: 'string' } }, read: { type: 'boolean', default: true } }, required: ['emailIds'] } },
        { name: 'bulk_move', description: 'Move multiple emails to a mailbox', inputSchema: { type: 'object', properties: { emailIds: { type: 'array', items: { type: 'string' } }, targetMailboxId: { type: 'string' } }, required: ['emailIds', 'targetMailboxId'] } },
        { name: 'bulk_delete', description: 'Delete multiple emails (move to trash)', inputSchema: { type: 'object', properties: { emailIds: { type: 'array', items: { type: 'string' } } }, required: ['emailIds'] } },
        { name: 'check_function_availability', description: 'Check which MCP functions are available based on account permissions', inputSchema: { type: 'object', properties: {} } },
        { name: 'test_bulk_operations', description: 'Test bulk operations by finding recent emails and performing safe operations (mark read/unread)', inputSchema: { type: 'object', properties: { dryRun: { type: 'boolean', default: true }, limit: { type: 'number', default: 3 } } } },
        { name: 'list_contacts', description: 'List all contacts', inputSchema: { type: 'object', properties: { limit: { type: 'number', default: 50 } } } },
        { name: 'get_contact', description: 'Get a specific contact by ID', inputSchema: { type: 'object', properties: { contactId: { type: 'string' } }, required: ['contactId'] } },
        { name: 'search_contacts', description: 'Search contacts by name or email', inputSchema: { type: 'object', properties: { query: { type: 'string' }, limit: { type: 'number', default: 20 } }, required: ['query'] } },
        { name: 'list_calendars', description: 'List all calendars', inputSchema: { type: 'object', properties: {} } },
        { name: 'list_calendar_events', description: 'List calendar events', inputSchema: { type: 'object', properties: { calendarId: { type: 'string' }, limit: { type: 'number', default: 50 } } } },
        { name: 'get_calendar_event', description: 'Get a specific calendar event by ID', inputSchema: { type: 'object', properties: { eventId: { type: 'string' } }, required: ['eventId'] } },
        { name: 'create_calendar_event', description: 'Create a new calendar event', inputSchema: { type: 'object', properties: { calendarId: { type: 'string' }, title: { type: 'string' }, description: { type: 'string' }, start: { type: 'string' }, end: { type: 'string' }, location: { type: 'string' }, participants: { type: 'array', items: { type: 'string' } } }, required: ['calendarId', 'title', 'start', 'end'] } },
        { name: 'list_identities', description: 'List sending identities', inputSchema: { type: 'object', properties: {} } }
      ],
    };
  });

  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    const { name, arguments: args } = request.params;

    try {
      switch (name) {
        case 'list_mailboxes': {
          const client = providers.getJmapClient();
          const mailboxes = await client.getMailboxes();
          return { content: [{ type: 'text', text: JSON.stringify(mailboxes, null, 2) }] };
        }
        case 'list_emails': {
          const { mailboxId, limit = 20 } = args as any;
          const client = providers.getJmapClient();
          const emails = await client.getEmails(mailboxId, limit);
          return { content: [{ type: 'text', text: JSON.stringify(emails, null, 2) }] };
        }
        case 'get_email': {
          const { emailId } = args as any;
          if (!emailId) throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
          const client = providers.getJmapClient();
          const email = await client.getEmailById(emailId);
          return { content: [{ type: 'text', text: JSON.stringify(email, null, 2) }] };
        }
        case 'send_email': {
          const { to, cc, bcc, from, mailboxId, subject, textBody, htmlBody } = args as any;
          if (!to || !Array.isArray(to) || to.length === 0) throw new McpError(ErrorCode.InvalidParams, 'to field is required and must be a non-empty array');
          if (!subject) throw new McpError(ErrorCode.InvalidParams, 'subject is required');
          if (!textBody && !htmlBody) throw new McpError(ErrorCode.InvalidParams, 'Either textBody or htmlBody is required');
          const client = providers.getJmapClient();
          const submissionId = await client.sendEmail({ to, cc, bcc, subject, textBody, htmlBody, from, mailboxId });
          return { content: [{ type: 'text', text: `Email sent successfully. Submission ID: ${submissionId}` }] };
        }
        case 'search_emails': {
          const { query, limit = 20 } = args as any;
          if (!query) throw new McpError(ErrorCode.InvalidParams, 'query is required');
          const client = providers.getJmapClient();
          const session = await client.getSession();
          const request: JmapRequest = {
            using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
            methodCalls: [
              ['Email/query', { accountId: session.accountId, filter: { text: query }, sort: [{ property: 'receivedAt', isAscending: false }], limit }, 'query'],
              ['Email/get', { accountId: session.accountId, '#ids': { resultOf: 'query', name: 'Email/query', path: '/ids' }, properties: ['id', 'subject', 'from', 'to', 'receivedAt', 'preview', 'hasAttachment'] }, 'emails']
            ]
          };
          const response = await client.makeRequest(request);
          const emails = response.methodResponses[1][1].list;
          return { content: [{ type: 'text', text: JSON.stringify(emails, null, 2) }] };
        }
        case 'list_contacts': {
          const { limit = 50 } = args as any;
          const contactsClient = providers.getContactsClient();
          const contacts = await contactsClient.getContacts(limit);
          return { content: [{ type: 'text', text: JSON.stringify(contacts, null, 2) }] };
        }
        case 'get_contact': {
          const { contactId } = args as any;
          if (!contactId) throw new McpError(ErrorCode.InvalidParams, 'contactId is required');
          const contactsClient = providers.getContactsClient();
          const contact = await contactsClient.getContactById(contactId);
          return { content: [{ type: 'text', text: JSON.stringify(contact, null, 2) }] };
        }
        case 'search_contacts': {
          const { query, limit = 20 } = args as any;
          if (!query) throw new McpError(ErrorCode.InvalidParams, 'query is required');
          const contactsClient = providers.getContactsClient();
          const contacts = await contactsClient.searchContacts(query, limit);
          return { content: [{ type: 'text', text: JSON.stringify(contacts, null, 2) }] };
        }
        case 'list_calendars': {
          const contactsClient = providers.getContactsClient();
          const calendars = await contactsClient.getCalendars();
          return { content: [{ type: 'text', text: JSON.stringify(calendars, null, 2) }] };
        }
        case 'list_calendar_events': {
          const { calendarId, limit = 50 } = args as any;
          const contactsClient = providers.getContactsClient();
          const events = await contactsClient.getCalendarEvents(calendarId, limit);
          return { content: [{ type: 'text', text: JSON.stringify(events, null, 2) }] };
        }
        case 'get_calendar_event': {
          const { eventId } = args as any;
          if (!eventId) throw new McpError(ErrorCode.InvalidParams, 'eventId is required');
          const contactsClient = providers.getContactsClient();
          const event = await contactsClient.getCalendarEventById(eventId);
          return { content: [{ type: 'text', text: JSON.stringify(event, null, 2) }] };
        }
        case 'create_calendar_event': {
          const { calendarId, title, description, start, end, location, participants } = args as any;
          if (!calendarId || !title || !start || !end) throw new McpError(ErrorCode.InvalidParams, 'calendarId, title, start, and end are required');
          const contactsClient = providers.getContactsClient();
          const eventId = await contactsClient.createCalendarEvent({ calendarId, title, description, start, end, location, participants });
          return { content: [{ type: 'text', text: `Calendar event created successfully. Event ID: ${eventId}` }] };
        }
        case 'list_identities': {
          const client = providers.getJmapClient();
          const identities = await client.getIdentities();
          return { content: [{ type: 'text', text: JSON.stringify(identities, null, 2) }] };
        }
        case 'get_recent_emails': {
          const { limit = 10, mailboxName = 'inbox' } = args as any;
          const client = providers.getJmapClient();
          const emails = await client.getRecentEmails(limit, mailboxName);
          return { content: [{ type: 'text', text: JSON.stringify(emails, null, 2) }] };
        }
        case 'mark_email_read': {
          const { emailId, read = true } = args as any;
          if (!emailId) throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
          const client = providers.getJmapClient();
          await client.markEmailRead(emailId, read);
          return { content: [{ type: 'text', text: `Email ${read ? 'marked as read' : 'marked as unread'} successfully` }] };
        }
        case 'delete_email': {
          const { emailId } = args as any;
          if (!emailId) throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
          const client = providers.getJmapClient();
          await client.deleteEmail(emailId);
          return { content: [{ type: 'text', text: 'Email deleted successfully (moved to trash)' }] };
        }
        case 'move_email': {
          const { emailId, targetMailboxId } = args as any;
          if (!emailId || !targetMailboxId) throw new McpError(ErrorCode.InvalidParams, 'emailId and targetMailboxId are required');
          const client = providers.getJmapClient();
          await client.moveEmail(emailId, targetMailboxId);
          return { content: [{ type: 'text', text: 'Email moved successfully' }] };
        }
        case 'get_email_attachments': {
          const { emailId } = args as any;
          if (!emailId) throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
          const client = providers.getJmapClient();
          const attachments = await client.getEmailAttachments(emailId);
          return { content: [{ type: 'text', text: JSON.stringify(attachments, null, 2) }] };
        }
        case 'download_attachment': {
          const { emailId, attachmentId } = args as any;
          if (!emailId || !attachmentId) throw new McpError(ErrorCode.InvalidParams, 'emailId and attachmentId are required');
          const client = providers.getJmapClient();
          const downloadUrl = await client.downloadAttachment(emailId, attachmentId);
          return { content: [{ type: 'text', text: `Download URL: ${downloadUrl}` }] };
        }
        case 'advanced_search': {
          const client = providers.getJmapClient();
          const emails = await client.advancedSearch(args as any);
          return { content: [{ type: 'text', text: JSON.stringify(emails, null, 2) }] };
        }
        case 'get_thread': {
          const { threadId } = args as any;
          if (!threadId) throw new McpError(ErrorCode.InvalidParams, 'threadId is required');
          const client = providers.getJmapClient();
          const thread = await client.getThread(threadId);
          return { content: [{ type: 'text', text: JSON.stringify(thread, null, 2) }] };
        }
        case 'get_mailbox_stats': {
          const { mailboxId } = args as any;
          const client = providers.getJmapClient();
          const stats = await client.getMailboxStats(mailboxId);
          return { content: [{ type: 'text', text: JSON.stringify(stats, null, 2) }] };
        }
        case 'get_account_summary': {
          const client = providers.getJmapClient();
          const summary = await client.getAccountSummary();
          return { content: [{ type: 'text', text: JSON.stringify(summary, null, 2) }] };
        }
        case 'bulk_mark_read': {
          const { emailIds, read = true } = args as any;
          if (!emailIds || !Array.isArray(emailIds) || emailIds.length === 0) throw new McpError(ErrorCode.InvalidParams, 'emailIds must be a non-empty array');
          const client = providers.getJmapClient();
          await client.bulkMarkRead(emailIds, read);
          return { content: [{ type: 'text', text: 'Bulk mark read/unread completed' }] };
        }
        case 'bulk_move': {
          const { emailIds, targetMailboxId } = args as any;
          if (!emailIds || !Array.isArray(emailIds) || emailIds.length === 0) throw new McpError(ErrorCode.InvalidParams, 'emailIds must be a non-empty array');
          if (!targetMailboxId) throw new McpError(ErrorCode.InvalidParams, 'targetMailboxId is required');
          const client = providers.getJmapClient();
          await client.bulkMove(emailIds, targetMailboxId);
          return { content: [{ type: 'text', text: 'Bulk move completed' }] };
        }
        case 'bulk_delete': {
          const { emailIds } = args as any;
          if (!emailIds || !Array.isArray(emailIds) || emailIds.length === 0) throw new McpError(ErrorCode.InvalidParams, 'emailIds must be a non-empty array');
          const client = providers.getJmapClient();
          await client.bulkDelete(emailIds);
          return { content: [{ type: 'text', text: 'Bulk delete completed' }] };
        }
        case 'check_function_availability': {
          // Minimal implementation: if session call works, basic mail features are available
          const client = providers.getJmapClient();
          const session = await client.getSession();
          const hasMail = Boolean(session.capabilities['urn:ietf:params:jmap:mail']);
          const hasContacts = Boolean(session.capabilities['urn:ietf:params:jmap:contacts']);
          const hasCalendar = Boolean(session.capabilities['urn:ietf:params:jmap:calendars']);
          return { content: [{ type: 'text', text: JSON.stringify({ mail: hasMail, contacts: hasContacts, calendar: hasCalendar }) }] };
        }
        case 'test_bulk_operations': {
          const { dryRun = true, limit = 3 } = args as any;
          const client = providers.getJmapClient();
          const emails = await client.getRecentEmails(Math.min(limit, 10));
          if (!emails.length) {
            return { content: [{ type: 'text', text: 'No recent emails to test.' }] };
          }
          if (dryRun) {
            return { content: [{ type: 'text', text: `Would mark ${emails.length} emails as read.` }] };
          }
          await client.bulkMarkRead(emails.map((e: any) => e.id), true);
          return { content: [{ type: 'text', text: `Marked ${emails.length} emails as read.` }] };
        }
        default:
          throw new McpError(ErrorCode.MethodNotFound, `Unknown tool: ${name}`);
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      return { content: [{ type: 'text', text: `Error: ${message}` }] };
    }
  });
}


