import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ErrorCode,
  ListToolsRequestSchema,
  McpError,
} from '@modelcontextprotocol/sdk/types.js';
import { FastmailAuth, FastmailConfig } from './auth.js';
import { JmapClient, JmapRequest } from './jmap-client.js';
import { ContactsCalendarClient } from './contacts-calendar.js';

const server = new Server(
  {
    name: 'fastmail-mcp',
    version: '1.0.0',
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

let jmapClient: JmapClient | null = null;
let contactsCalendarClient: ContactsCalendarClient | null = null;

function initializeClient(): JmapClient {
  if (jmapClient) {
    return jmapClient;
  }

  const apiToken = process.env.FASTMAIL_API_TOKEN;
  if (!apiToken) {
    throw new McpError(
      ErrorCode.InvalidRequest,
      'FASTMAIL_API_TOKEN environment variable is required'
    );
  }

  const config: FastmailConfig = {
    apiToken,
    baseUrl: process.env.FASTMAIL_BASE_URL
  };

  const auth = new FastmailAuth(config);
  jmapClient = new JmapClient(auth);
  return jmapClient;
}

function initializeContactsCalendarClient(): ContactsCalendarClient {
  if (contactsCalendarClient) {
    return contactsCalendarClient;
  }

  const apiToken = process.env.FASTMAIL_API_TOKEN;
  if (!apiToken) {
    throw new McpError(
      ErrorCode.InvalidRequest,
      'FASTMAIL_API_TOKEN environment variable is required'
    );
  }

  const config: FastmailConfig = {
    apiToken,
    baseUrl: process.env.FASTMAIL_BASE_URL
  };

  const auth = new FastmailAuth(config);
  contactsCalendarClient = new ContactsCalendarClient(auth);
  return contactsCalendarClient;
}

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
          },
          required: ['to', 'subject'],
        },
      },
      {
        name: 'search_emails',
        description: 'Search emails by subject or content',
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
        name: 'list_contacts',
        description: 'List contacts from the address book',
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
        description: 'Search contacts by name or email',
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
        description: 'Get a specific calendar event by ID',
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
        description: 'Create a new calendar event',
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
              description: 'Start time in ISO 8601 format',
            },
            end: {
              type: 'string',
              description: 'End time in ISO 8601 format',
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
                  email: { type: 'string' },
                  name: { type: 'string' }
                }
              },
              description: 'Event participants (optional)',
            },
          },
          required: ['calendarId', 'title', 'start', 'end'],
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
        description: 'Get the most recent emails from inbox (like top-ten)',
        inputSchema: {
          type: 'object',
          properties: {
            limit: {
              type: 'number',
              description: 'Number of recent emails to retrieve (default: 10, max: 50)',
              default: 10,
            },
            mailboxName: {
              type: 'string',
              description: 'Mailbox to search (default: inbox)',
              default: 'inbox',
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
        description: 'Download an email attachment',
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
          },
          required: ['emailId', 'attachmentId'],
        },
      },
      {
        name: 'advanced_search',
        description: 'Advanced email search with multiple criteria',
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
            mailboxId: {
              type: 'string',
              description: 'Search within specific mailbox',
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
              type: 'number',
              description: 'Maximum results (default: 50)',
              default: 50,
            },
          },
        },
      },
      {
        name: 'get_thread',
        description: 'Get all emails in a conversation thread',
        inputSchema: {
          type: 'object',
          properties: {
            threadId: {
              type: 'string',
              description: 'ID of the thread/conversation',
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
    ],
  };
});

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  try {
    const client = initializeClient();

    switch (name) {
      case 'list_mailboxes': {
        const mailboxes = await client.getMailboxes();
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(mailboxes, null, 2),
            },
          ],
        };
      }

      case 'list_emails': {
        const { mailboxId, limit = 20 } = args as any;
        const emails = await client.getEmails(mailboxId, limit);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(emails, null, 2),
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

      case 'send_email': {
        const { to, cc, bcc, from, mailboxId, subject, textBody, htmlBody } = args as any;
        if (!to || !Array.isArray(to) || to.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'to field is required and must be a non-empty array');
        }
        if (!subject) {
          throw new McpError(ErrorCode.InvalidParams, 'subject is required');
        }
        if (!textBody && !htmlBody) {
          throw new McpError(ErrorCode.InvalidParams, 'Either textBody or htmlBody is required');
        }

        const submissionId = await client.sendEmail({
          to,
          cc,
          bcc,
          from,
          mailboxId,
          subject,
          textBody,
          htmlBody,
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

      case 'search_emails': {
        const { query, limit = 20 } = args as any;
        if (!query) {
          throw new McpError(ErrorCode.InvalidParams, 'query is required');
        }

        // For search, we'll use the Email/query method with a text filter
        const session = await client.getSession();
        const request: JmapRequest = {
          using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
          methodCalls: [
            ['Email/query', {
              accountId: session.accountId,
              filter: { text: query },
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

        const response = await client.makeRequest(request);
        const emails = response.methodResponses[1][1].list;

        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(emails, null, 2),
            },
          ],
        };
      }

      case 'list_contacts': {
        const { limit = 50 } = args as any;
        const contactsClient = initializeContactsCalendarClient();
        const contacts = await contactsClient.getContacts(limit);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(contacts, null, 2),
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
        const { query, limit = 20 } = args as any;
        if (!query) {
          throw new McpError(ErrorCode.InvalidParams, 'query is required');
        }
        const contactsClient = initializeContactsCalendarClient();
        const contacts = await contactsClient.searchContacts(query, limit);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(contacts, null, 2),
            },
          ],
        };
      }

      case 'list_calendars': {
        const contactsClient = initializeContactsCalendarClient();
        const calendars = await contactsClient.getCalendars();
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(calendars, null, 2),
            },
          ],
        };
      }

      case 'list_calendar_events': {
        const { calendarId, limit = 50 } = args as any;
        const contactsClient = initializeContactsCalendarClient();
        const events = await contactsClient.getCalendarEvents(calendarId, limit);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(events, null, 2),
            },
          ],
        };
      }

      case 'get_calendar_event': {
        const { eventId } = args as any;
        if (!eventId) {
          throw new McpError(ErrorCode.InvalidParams, 'eventId is required');
        }
        const contactsClient = initializeContactsCalendarClient();
        const event = await contactsClient.getCalendarEventById(eventId);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(event, null, 2),
            },
          ],
        };
      }

      case 'create_calendar_event': {
        const { calendarId, title, description, start, end, location, participants } = args as any;
        if (!calendarId || !title || !start || !end) {
          throw new McpError(ErrorCode.InvalidParams, 'calendarId, title, start, and end are required');
        }
        const contactsClient = initializeContactsCalendarClient();
        const eventId = await contactsClient.createCalendarEvent({
          calendarId,
          title,
          description,
          start,
          end,
          location,
          participants,
        });
        return {
          content: [
            {
              type: 'text',
              text: `Calendar event created successfully. Event ID: ${eventId}`,
            },
          ],
        };
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
        const { limit = 10, mailboxName = 'inbox' } = args as any;
        const client = initializeClient();
        const emails = await client.getRecentEmails(limit, mailboxName);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(emails, null, 2),
            },
          ],
        };
      }

      case 'mark_email_read': {
        const { emailId, read = true } = args as any;
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
        const { emailId, attachmentId } = args as any;
        if (!emailId || !attachmentId) {
          throw new McpError(ErrorCode.InvalidParams, 'emailId and attachmentId are required');
        }
        const client = initializeClient();
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

      case 'advanced_search': {
        const { query, from, to, subject, hasAttachment, isUnread, mailboxId, after, before, limit } = args as any;
        const client = initializeClient();
        const emails = await client.advancedSearch({
          query, from, to, subject, hasAttachment, isUnread, mailboxId, after, before, limit
        });
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(emails, null, 2),
            },
          ],
        };
      }

      case 'get_thread': {
        const { threadId } = args as any;
        if (!threadId) {
          throw new McpError(ErrorCode.InvalidParams, 'threadId is required');
        }
        const client = initializeClient();
        const thread = await client.getThread(threadId);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(thread, null, 2),
            },
          ],
        };
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

      default:
        throw new McpError(ErrorCode.MethodNotFound, `Unknown tool: ${name}`);
    }
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      `Tool execution failed: ${error instanceof Error ? error.message : String(error)}`
    );
  }
});

async function runServer() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error('Fastmail MCP server running on stdio');
}

runServer().catch((error) => {
  console.error('Fatal error running server:', error);
  process.exit(1);
});