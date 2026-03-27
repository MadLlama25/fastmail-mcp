# Fastmail MCP Server
[![CI](https://github.com/MadLlama25/fastmail-mcp/actions/workflows/ci.yml/badge.svg)](https://github.com/MadLlama25/fastmail-mcp/actions/workflows/ci.yml)
[![codecov](https://codecov.io/gh/MadLlama25/fastmail-mcp/graph/badge.svg)](https://codecov.io/gh/MadLlama25/fastmail-mcp)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![npm](https://img.shields.io/npm/v/fastmail-mcp)](https://www.npmjs.com/package/fastmail-mcp)

MCP server for the Fastmail API. Gives AI assistants access to email, contacts, and calendars via JMAP (with CalDAV fallback for calendar operations).

## Features

- **Email**: list, search, send, reply, draft, thread, mark read/unread, delete, move, label
- **Attachments**: list and download email attachments
- **Bulk operations**: mark read, move, delete, add/remove labels for multiple emails at once
- **Search**: multi-criteria filtering (sender, date range, attachments, read status)
- **Contacts**: list, get, search by name or email
- **Calendars**: list, get, create events (with CalDAV fallback)
- **Account**: sending identities, mailbox stats, account summary
- **Labels vs moves**: `move_email`/`bulk_move` replaces all mailboxes (folder behavior); `add_labels`/`remove_labels` preserves existing mailboxes (label behavior)

## Installation

### Prerequisites

- Node.js 18+
- Fastmail API token (Settings -> Privacy & Security -> Manage API tokens)

### From npm

```bash
npx fastmail-mcp
```

### From GitHub

```bash
npx --yes github:MadLlama25/fastmail-mcp fastmail-mcp
```

Pin to a tagged release:

```bash
npx --yes github:MadLlama25/fastmail-mcp@v1.8.2 fastmail-mcp
```

### From source

```bash
git clone https://github.com/MadLlama25/fastmail-mcp.git
cd fastmail-mcp
npm install
npm run build
npm start
```

For development with auto-reload: `npm run dev`

## Configuration

Set your API token:

```bash
export FASTMAIL_API_TOKEN="your_api_token_here"
# Optional: override base URL (defaults to https://api.fastmail.com)
export FASTMAIL_BASE_URL="https://api.fastmail.com"
```

Windows PowerShell:

```powershell
$env:FASTMAIL_API_TOKEN="your_token"
$env:FASTMAIL_BASE_URL="https://api.fastmail.com"
```

## Claude Desktop Extension (DXT)

1. Build and pack:
   ```bash
   npm run build
   npx @anthropic-ai/dxt pack
   ```
2. Open `fastmail-mcp.dxt` or drag it into Claude Desktop.
3. When prompted, paste your Fastmail API token (stored encrypted by Claude). Leave the base URL blank for the default.

## Available Tools (38 Total)

### Email Tools

- **list_mailboxes**: Get all mailboxes in your account
- **list_emails**: List emails from a specific mailbox or all mailboxes
  - Parameters: `mailboxId` (optional), `limit` (default: 20)
- **get_email**: Get a specific email by ID
  - Parameters: `emailId` (required)
- **send_email**: Send an email (supports threading via optional `inReplyTo` and `references` headers)
  - Parameters: `to` (required array), `cc` (optional array), `bcc` (optional array), `from` (optional), `mailboxId` (optional), `subject` (required), `textBody` (optional), `htmlBody` (optional), `inReplyTo` (optional array), `references` (optional array)
- **reply_email**: Reply to an existing email with proper threading headers (automatically builds In-Reply-To and References). Set `send=false` to save as draft instead of sending.
  - Parameters: `originalEmailId` (required), `to` (optional array, defaults to original sender), `cc` (optional array), `bcc` (optional array), `from` (optional), `textBody` (optional), `htmlBody` (optional), `send` (optional boolean, default: true)
- **save_draft**: Save an email as a draft without sending (supports threading headers for reply drafts)
  - Parameters: `to` (required array), `cc` (optional array), `bcc` (optional array), `from` (optional), `subject` (required), `textBody` (optional), `htmlBody` (optional), `inReplyTo` (optional array), `references` (optional array)
- **create_draft**: Create a minimal email draft (at least one of to/subject/body required)
  - Parameters: `to` (optional array), `cc` (optional array), `bcc` (optional array), `from` (optional), `mailboxId` (optional), `subject` (optional), `textBody` (optional), `htmlBody` (optional)
- **search_emails**: Search emails by content
  - Parameters: `query` (required), `limit` (default: 20)
- **get_recent_emails**: Get the most recent emails from a mailbox (inspired by JMAP-Samples top-ten)
  - Parameters: `limit` (default: 10, max: 50), `mailboxName` (default: 'inbox')
- **mark_email_read**: Mark an email as read or unread
  - Parameters: `emailId` (required), `read` (default: true)
- **delete_email**: Delete an email (move to trash)
  - Parameters: `emailId` (required)
- **move_email**: Move an email to a different mailbox (replaces all mailboxes)
  - Parameters: `emailId` (required), `targetMailboxId` (required)
- **add_labels**: Add labels (mailboxes) to an email without removing existing ones
  - Parameters: `emailId` (required), `mailboxIds` (required array)
- **remove_labels**: Remove specific labels (mailboxes) from an email
  - Parameters: `emailId` (required), `mailboxIds` (required array)

### Advanced Email Features

- **get_email_attachments**: Get list of attachments for an email
  - Parameters: `emailId` (required)
- **download_attachment**: Download an email attachment. If savePath is provided, saves the file to disk and returns the file path and size. Otherwise returns a download URL.
  - Parameters: `emailId` (required), `attachmentId` (required), `savePath` (optional)
- **advanced_search**: Advanced email search with multiple criteria
  - Parameters: `query` (optional), `from` (optional), `to` (optional), `subject` (optional), `hasAttachment` (optional), `isUnread` (optional), `mailboxId` (optional), `after` (optional), `before` (optional), `limit` (default: 50)
- **get_thread**: Get all emails in a conversation thread
  - Parameters: `threadId` (required)

### Email Statistics & Analytics

- **get_mailbox_stats**: Get statistics for a mailbox (unread count, total emails, etc.)
  - Parameters: `mailboxId` (optional, defaults to all mailboxes)
- **get_account_summary**: Get overall account summary with statistics

### Bulk Operations

- **bulk_mark_read**: Mark multiple emails as read/unread
  - Parameters: `emailIds` (required array), `read` (default: true)
- **bulk_move**: Move multiple emails to a mailbox
  - Parameters: `emailIds` (required array), `targetMailboxId` (required)
- **bulk_delete**: Delete multiple emails (move to trash)
  - Parameters: `emailIds` (required array)
- **bulk_add_labels**: Add labels to multiple emails simultaneously
  - Parameters: `emailIds` (required array), `mailboxIds` (required array)
- **bulk_remove_labels**: Remove labels from multiple emails simultaneously
  - Parameters: `emailIds` (required array), `mailboxIds` (required array)

### Contact Tools

- **list_contacts**: List all contacts
  - Parameters: `limit` (default: 50)
- **get_contact**: Get a specific contact by ID
  - Parameters: `contactId` (required)
- **search_contacts**: Search contacts by name or email
  - Parameters: `query` (required), `limit` (default: 20)

### Calendar Tools

- **list_calendars**: List all calendars
- **list_calendar_events**: List calendar events
  - Parameters: `calendarId` (optional), `limit` (default: 50)
- **get_calendar_event**: Get a specific calendar event by ID
  - Parameters: `eventId` (required)
- **create_calendar_event**: Create a new calendar event
  - Parameters: `calendarId` (required), `title` (required), `description` (optional), `start` (required, ISO 8601), `end` (required, ISO 8601), `location` (optional), `participants` (optional array)

### Identity & Testing Tools

- **list_identities**: List sending identities (email addresses that can be used for sending)
- **check_function_availability**: Check which functions are available based on account permissions (includes setup guidance)
- **test_bulk_operations**: Safely test bulk operations with dry-run mode
  - Parameters: `dryRun` (default: true), `limit` (default: 3)

## API Information

Uses the [JMAP](https://jmap.io/) protocol (JSON Meta Application Protocol) -- a modern alternative to IMAP. Many features are inspired by the official [Fastmail JMAP-Samples](https://github.com/fastmail/JMAP-Samples) repository.

Authentication is via bearer token. Fastmail applies rate limits to API requests; the server handles standard rate limiting, but excessive requests may be throttled.

## CalDAV Calendar Support

Fastmail does not expose JMAP calendar access yet (`urn:ietf:params:jmap:calendars` is still an [IETF Internet-Draft](https://datatracker.ietf.org/doc/draft-ietf-jmap-calendars/)). The server automatically falls back to **CalDAV** via `caldav.fastmail.com` when JMAP calendars are unavailable.

To enable CalDAV, create an app-specific password (Settings -> Privacy & Security -> Manage app passwords) and set:

```bash
export FASTMAIL_CALDAV_USERNAME="your-email@fastmail.com"
export FASTMAIL_CALDAV_PASSWORD="your-app-specific-password"
```

Without these variables, the server uses JMAP only (calendar tools will fail if JMAP calendars are not available on your account).

## Development

```text
src/
├── index.ts              # MCP server entry point
├── auth.ts               # Authentication
├── jmap-client.ts        # JMAP client wrapper
├── contacts-calendar.ts  # Contacts and calendar extensions
└── caldav-client.ts      # CalDAV fallback client
```

```bash
npm run build   # compile TypeScript
npm run dev     # dev mode with auto-reload
npm test        # run tests
```

## Troubleshooting

- **Authentication errors**: verify your API token is valid and has the necessary permissions.
- **Build errors**: run `npm run build` and check for TypeScript compilation errors.
- **Serialization errors** in email tools: upgrade to v1.7.1+ (caused by incomplete JMAP response validation).
- **"Forbidden" errors** on calendar/contacts: may require a business/professional plan, broader API token scope, or CalDAV credentials. Run `check_function_availability` for step-by-step guidance.
- **Testing your setup**: use `check_function_availability` and `test_bulk_operations` (dry-run mode) to verify permissions without side effects.

## Privacy & Security

- API tokens are encrypted at rest when installed via the DXT and never logged by this server.
- Error messages are sanitized -- tokens, email addresses, identities, and attachment blob IDs are not included.
- Tool responses include email metadata/content by design, but credentials and internal identifiers are not disclosed.

## License

MIT
