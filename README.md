# Fastmail MCP Server (Unofficial)

An **unofficial** Model Context Protocol (MCP) server that provides access to the Fastmail API, enabling AI assistants to interact with email, contacts, and calendar data.

> **Disclaimer:** This is a community project. It is **not affiliated with, endorsed by, or supported by Fastmail**. "Fastmail" is a trademark of Fastmail Pty Ltd; it is used here only to describe compatibility with their public JMAP/CalDAV/WebDAV APIs. Use at your own risk under the terms of the project license.

## Features

### Core Email Operations
- List mailboxes and get mailbox statistics
- List, search, and filter emails with advanced criteria
- Get specific emails by ID with full content
- Send emails (text and HTML) with proper draft/sent handling
- Reply to emails with proper threading (In-Reply-To, References headers)
- Create, edit, and send email drafts (with or without threading)
- Email management: mark read/unread, delete, move between folders

### Advanced Email Features
- **Attachment Handling**: List, download, and send attachments; save attachments straight to WebDAV cloud storage
- **Privacy-lean metadata tools**: Metadata-only variants of list/search/thread tools (no body content)
- **Threading Support**: Get complete conversation threads
- **Advanced Search**: Multi-criteria filtering (sender, date range, attachments, read status)
- **Bulk Operations**: Process multiple emails simultaneously
- **Statistics & Analytics**: Account summaries and mailbox statistics

### Contacts Operations
- List all contacts with full contact information
- Get specific contacts by ID
- Search contacts by name or email
- Create, update, and delete contacts (JMAP ContactCard/set; requires an API token with read-write contacts scope)

### Calendar Operations
- List, get, create, update, and delete calendar events (via CalDAV)
- All-day and timed events, participants, recurrence-aware updates

### Label vs Move Operations
- **move_email/bulk_move**: Replaces ALL mailboxes for an email (folder behavior)
- **add_labels/remove_labels**: Adds/removes SPECIFIC mailboxes while preserving others (label behavior)

### Identity & Account Management
- List available sending identities
- Account summary with comprehensive statistics

## Setup

### Prerequisites
- Node.js 20+ 
- A Fastmail account with API access
- Fastmail API token

### Installation

1. Clone or download this repository
2. Install dependencies:
   ```bash
   npm install
   ```

3. Build the project:
   ```bash
   npm run build
   ```

### Configuration

1. Get your Fastmail API token:
   - Log in to Fastmail web interface
   - Go to Settings → Privacy & Security
   - Find "Connected apps & API tokens" section
   - Click "Manage API tokens"
   - Click "New API token"
   - Copy the generated token

2. Set environment variables:
   ```bash
   export FASTMAIL_API_TOKEN="your_api_token_here"
   # Optional: customize base URL (defaults to https://api.fastmail.com)
   # Only api.fastmail.com and www.fastmailusercontent.com are accepted by default.
   # For self-hosted JMAP servers, also set FASTMAIL_ALLOW_UNSAFE_BASE_URL=true.
   export FASTMAIL_BASE_URL="https://api.fastmail.com"
   # Optional: customize attachment download directory (defaults to ~/Downloads/fastmail-mcp/).
   # download_attachment savePaths are confined to this directory; set it to the root
   # you want attachments saved under to write there directly in one step.
   export FASTMAIL_DOWNLOAD_DIR="/path/to/your/downloads"
   ```

### Running the Server

Start the MCP server:
```bash
npm start
```

For development with auto-reload:
```bash
npm run dev
```

### Run from a clone

```bash
git clone https://github.com/MadLlama25/fastmail-mcp && cd fastmail-mcp
npm install && npm run build
FASTMAIL_API_TOKEN="your_token" node dist/index.js
```

> Note: `npx github:MadLlama25/fastmail-mcp` does **not** work on npm 10 (a known npm
> `GitFetcher` bug). Use the clone above, or install the packaged
> [Desktop Extension](#install-as-a-claude-desktop-extension-dxt).

## Install as a Claude Desktop Extension (DXT)

You can install this server as a Desktop Extension for Claude Desktop using the packaged `.dxt` file.

1. Build and pack:
   ```bash
   npm run build
   npx @anthropic-ai/dxt pack
   ```
   This produces `fastmail-mcp.dxt` in the project root.

2. Install into Claude Desktop:
   - Open the `.dxt` file, or drag it into Claude Desktop
   - When prompted:
     - Fastmail API Token: paste your token (stored encrypted by Claude) — required
     - Fastmail Base URL: leave blank to use `https://api.fastmail.com` (default)
     - Download Directory: leave blank for `~/Downloads/fastmail-mcp/`
     - CalDAV Username / Password / Display Name: optional — required for calendar tools (use an app-specific password; see [CalDAV Calendar Support](#caldav-calendar-support))
     - WebDAV URL / Username / Password: optional — required for `save_attachment_to_webdav` (see [WebDAV file storage](#webdav-file-storage-optional))

3. Use any of the tools (e.g. `get_recent_emails`).

## Available Tools (52 Total)

**Response shape for list/search tools:** the query tools (`list_emails`, `list_emails_metadata`, `search_emails`, `search_emails_metadata`, `get_recent_emails`, `advanced_search`, `advanced_search_metadata`, `list_contacts`, `search_contacts`) return a `{"total", "items"}` JSON envelope — `total` is the server-reported match count, `items` the returned page. When the server reports no total, a bare array is returned.

**🎯 Most Popular Tools:**
- **check_function_availability**: Check what's available and get setup guidance  
- **test_bulk_operations**: Safely test bulk operations with dry-run mode
- **send_email**: Full-featured email sending with proper draft/sent handling
- **advanced_search**: Powerful multi-criteria email filtering
- **get_recent_emails**: Quick access to recent emails from any mailbox

### Email Tools

- **list_mailboxes**: Get all mailboxes in your account. On accounts with many mailboxes the full output can be large — pass `properties` for a slim view.
  - Parameters: `properties` (optional array of fields to return), `parentId` (optional; only children of this mailbox, `null` for top level)
- **get_mailbox_by_name**: Look up a mailbox by its full path from the root (e.g. `Inbox/Receipts`)
  - Parameters: `path` (required)
- **create_mailbox**: Create a new mailbox (folder/label)
  - Parameters: `name` (required), `parentId` (optional, omit or null for top level)
- **list_emails**: List emails from a specific mailbox or all mailboxes
  - Parameters: `mailboxId` (optional), `limit` (default: 20, max: 100), `ascending` (optional, oldest first)
- **list_emails_metadata**: List emails from a mailbox, metadata only (headers, no body content)
  - Parameters: `mailboxId` (optional), `limit` (default: 20, max: 100), `ascending` (optional, oldest first)
- **get_email**: Get a specific email by ID
  - Parameters: `emailId` (required)
- **get_email_metadata**: Get a specific email's metadata only (allowlisted headers, no body)
  - Parameters: `emailId` (required)
- **send_email**: Send an email (supports threading via optional `inReplyTo` and `references` headers)
  - Parameters: `to` (required — array or comma-separated string), `cc` (optional array), `bcc` (optional array), `from` (optional), `mailboxId` (optional), `subject` (required), `textBody` (optional), `htmlBody` (optional), `inReplyTo` (optional array), `references` (optional array), `replyTo` (optional array), `attachments` (optional array — see [Email attachments on send](#email-attachments-on-send))
- **reply_email**: Reply to an existing email with proper threading headers (automatically builds In-Reply-To and References). Set `send=false` to save as draft instead of sending.
  - Parameters: `originalEmailId` (required), `to` (optional array, defaults to original sender), `cc` (optional array), `bcc` (optional array), `from` (optional), `textBody` (optional), `htmlBody` (optional), `send` (optional boolean, default: true), `replyTo` (optional array), `attachments` (optional array — see [Email attachments on send](#email-attachments-on-send))
- **create_draft**: Create an email draft (at least one of to/subject/body/attachments required; supports threading headers for reply drafts)
  - Parameters: `to` (optional array), `cc` (optional array), `bcc` (optional array), `from` (optional), `mailboxId` (optional), `subject` (optional), `textBody` (optional), `htmlBody` (optional), `replyTo` (optional array), `inReplyTo` (optional array), `references` (optional array), `attachments` (optional array — see [Email attachments on send](#email-attachments-on-send))
- **edit_draft**: Edit an existing draft in place — only provided fields change; existing attachments are preserved
  - Parameters: `emailId` (required), `to`, `cc`, `bcc`, `from`, `subject`, `textBody`, `htmlBody`, `replyTo`, `attachments` (all optional)
- **send_draft**: Send an existing draft
  - Parameters: `emailId` (required)
- **search_emails**: Search emails by content
  - Parameters: `query` (required), `limit` (default: 20, max: 100), `ascending` (optional, oldest first), `excludeDrafts` (optional, omit draft messages)
  - Drafts are **included by default**. Set `excludeDrafts: true` to filter them out server-side.
  - Searches **all mailboxes including Trash and Spam**. For cleanup/verification flows, exclude the Trash mailbox explicitly (e.g. `advanced_search` with `excludeMailboxIds`) rather than trusting a bare search count.
- **get_recent_emails**: Get the most recent emails (inspired by JMAP-Samples top-ten)
  - Parameters: `limit` (default: 10, max: 50), `mailboxName` (optional), `ascending` (optional, oldest first)
  - When `mailboxName` is omitted, all mailboxes are searched **except Trash and Spam**. Pass a mailbox name (e.g. `'inbox'`, `'sent'`) to scope to one folder.
- **search_emails_metadata**: Search emails by content, returning metadata only
  - Parameters: `query` (required), `limit` (default: 20, max: 100), `ascending` (optional, oldest first)
- **mark_email_read**: Mark an email as read or unread
  - Parameters: `emailId` (required), `read` (default: true)
- **pin_email**: Pin or unpin an email
  - Parameters: `emailId` (required), `pinned` (default: true)
- **archive_email**: Archive an email — moves it and marks it read in one atomic step
  - Parameters: `emailId` (required), `targetMailboxId` (required)
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
  - `savePath` may be absolute or relative. Relative paths (including a bare filename) resolve against the download directory, so an attachment lands there in one step. Absolute paths must fall within that directory; traversal or symlink escape outside it is rejected. To save directly into your own location, set `FASTMAIL_DOWNLOAD_DIR` to that root — confinement stays on, scoped to the directory you choose.
- **save_attachment_to_webdav**: Save an attachment directly to WebDAV cloud storage (Fastmail Files, Nextcloud, ...) without touching local disk
  - Parameters: `emailId` (required), `attachmentId` (required), `remotePath` (required, relative), `overwrite` (default false), `createParents` (default true)
  - The storage server and credentials come from `FASTMAIL_WEBDAV_*` env config; the tool only chooses the relative path beneath that base. Existing files are never replaced unless `overwrite: true`.
- **advanced_search**: Advanced email search with multiple criteria
  - Parameters: `query` (optional), `from` (optional), `to` (optional), `subject` (optional), `hasAttachment` (optional), `isUnread` (optional), `isPinned` (optional), `mailboxId` (optional), `requiredMailboxIds` (optional array — email must be in ALL of these), `excludeMailboxIds` (optional array — exclude emails in any of these), `after` (optional), `before` (optional), `limit` (default: 50, max: 100), `ascending` (optional, oldest first)
  - Like `search_emails`, searches **all mailboxes including Trash and Spam** — scope with `mailboxId`/`excludeMailboxIds` when that matters. (`get_recent_emails` is the one that excludes Trash/Spam by default.)
- **advanced_search_metadata**: Same filters as `advanced_search`, metadata-only results (no body content)
- **get_thread**: Get all emails in a conversation thread
  - Parameters: `threadId` (required), `includeDrafts` (optional, include in-progress drafts)
  - Draft messages are **excluded by default** (an in-progress reply is noise when reading a conversation). Set `includeDrafts: true` to include them. Drafts are identified by the `$draft` keyword, so the asymmetry with `search_emails` (which includes drafts by default) is deliberate: a search should still find everything you've written.
- **get_thread_metadata**: Get all emails in a thread, metadata only. Also accepts an email ID and resolves its parent thread.
  - Parameters: `threadId` (required), `includeDrafts` (optional)

### Email Statistics & Analytics

- **get_mailbox_stats**: Get statistics for a mailbox (unread count, total emails, etc.)
  - Parameters: `mailboxId` (optional, defaults to all mailboxes)
- **get_account_summary**: Get overall account summary with statistics

### Bulk Operations

- **bulk_mark_read**: Mark multiple emails as read/unread
  - Parameters: `emailIds` (required array), `read` (default: true)
- **bulk_pin**: Pin or unpin multiple emails
  - Parameters: `emailIds` (required array), `pinned` (default: true)
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
  - Parameters: `limit` (default: 50, max: 200)
- **get_contact**: Get a specific contact by ID
  - Parameters: `contactId` (required)
- **search_contacts**: Search contacts by name or email
  - Parameters: `query` (required), `limit` (default: 20, max: 100)
- **create_contact**: Create a new contact (requires read-write contacts scope on the API token)
  - Parameters: `name` `{given, surname, full}`, `emails` `[{address, label}]`, `phones` `[{number, label}]`, `addresses` `[{full, label}]`, `notes`, `addressBookId` (all optional, but a name or one email is required)
- **update_contact**: Update an existing contact — each provided field **wholly replaces** the stored value (`emails: []` removes all emails); unspecified fields are untouched
  - Parameters: `contactId` (required), same fields as create, `expectState` (optional JMAP state precondition)
- **delete_contact**: Permanently delete a contact (cannot be undone)
  - Parameters: `contactId` (required), `expectState` (optional)

### Calendar Tools

- **list_calendars**: List all calendars
- **list_calendar_events**: List calendar events (core fields only — no participants for token efficiency)
  - Parameters: `calendarId` (optional), `startDate` (optional, ISO 8601), `endDate` (optional, ISO 8601), `limit` (default: 50, max: 500)
- **get_calendar_event**: Get a specific calendar event by ID. Returns organizer and participants when available.
  - Parameters: `eventId` (required)
- **create_calendar_event**: Create a new calendar event. Supports date-only (e.g. `2026-04-01`) for all-day events. DTEND is exclusive per RFC 5545 — a one-day event on April 1 needs `end: "2026-04-02"`.
  - Parameters: `calendarId` (required), `title` (required), `description` (optional), `start` (required, ISO 8601 or date-only), `end` (required, ISO 8601 or date-only), `location` (optional), `participants` (optional array of `{email, name?}`)
- **update_calendar_event**: Patch an existing calendar event. Preserves all existing data (attendees, reminders, recurrence rules, etc.) not being changed. Omit a field to leave it unchanged; passing an empty or whitespace-only string for `title`, `description`, or `location` is rejected (it won't silently blank the property). To delete `description` or `location`, list them in `clearFields`. Floating times (no Z/offset) preserve the original timezone. WARNING: providing `participants` replaces ALL existing attendee data; `participants: []` removes all attendees (and the now-orphaned ORGANIZER).
  - Parameters: `eventId` (required), `title`, `description`, `start`, `end`, `location`, `participants` (array of `{email, name?}`), `clearFields` (array of `"description"`/`"location"` to delete), `confirmRecurring` (boolean)
- **delete_calendar_event**: Delete a calendar event
  - Parameters: `eventId` (required)

#### Calendar known limitations

- **Recurring events**: Only "all events" modification is supported (master VEVENT). "This event only" or "this and future events" are not supported. Changing start/end on recurring events with exception overrides requires `confirmRecurring: true` — orphaned exceptions are pruned to prevent server errors.
- **Attendee parameters**: RSVP, ROLE, CUTYPE and other attendee parameters are parsed on read but not settable on create/update — only `email` and `name` are accepted.

### Identity & Testing Tools

- **list_identities**: List sending identities (email addresses that can be used for sending)
- **check_function_availability**: Check which functions are available based on account permissions (includes setup guidance). Calendar tools run over CalDAV, so calendar is reported available when CalDAV credentials are configured, regardless of the JMAP calendar capability.
- **test_bulk_operations**: Safely test bulk operations with dry-run mode
  - Parameters: `dryRun` (default: true), `limit` (default: 3)

## API Information

This server uses the JMAP (JSON Meta Application Protocol) API provided by Fastmail. JMAP is a modern, efficient alternative to IMAP for email access.

### Inspired by Fastmail JMAP-Samples

Many features in this MCP server are inspired by the official [Fastmail JMAP-Samples](https://github.com/fastmail/JMAP-Samples) repository, including:
- Recent emails retrieval (based on top-ten example)
- Email management operations
- Efficient chained JMAP method calls

### Authentication
The server uses bearer token authentication with Fastmail's API. API tokens provide secure access without exposing your main account password.

### Rate Limits
Fastmail applies rate limits to API requests. The server handles standard rate limiting, but excessive requests may be throttled.

## CalDAV Calendar Support

Fastmail does not currently expose calendar access via JMAP API tokens — the `urn:ietf:params:jmap:calendars` scope is not available because the JMAP Calendars specification is still an IETF Internet-Draft ([draft-ietf-jmap-calendars](https://datatracker.ietf.org/doc/draft-ietf-jmap-calendars/)). Fastmail has stated they will add JMAP calendar support once the spec becomes an RFC, but there is no public timeline.

However, Fastmail fully supports **CalDAV** for calendar access via `caldav.fastmail.com`. All calendar tools use CalDAV directly.

### Setup

1. Create an app-specific password on Fastmail:
   - Go to **Settings → Privacy & Security → Manage app passwords**
   - Create a new app password (you can name it "CalDAV MCP" or similar)

2. Set the following environment variables:
   ```bash
   export FASTMAIL_CALDAV_USERNAME="your-email@fastmail.com"
   export FASTMAIL_CALDAV_PASSWORD="your-app-specific-password"
   # Optional: display name for ORGANIZER when creating events with participants
   export FASTMAIL_CALDAV_DISPLAY_NAME="Your Name"
   ```

When these variables are set, all calendar tools are available. When they are not set, calendar tools will return an error with setup instructions.

### WebDAV file storage (optional)

`save_attachment_to_webdav` saves attachments straight to cloud storage. Configure the target (never supplied by tools at runtime — this is deliberate, so a misbehaving caller cannot redirect uploads):

```bash
# Fastmail Files:
export FASTMAIL_WEBDAV_URL="https://myfiles.fastmail.com/"
export FASTMAIL_WEBDAV_USERNAME="your-email@fastmail.com"
export FASTMAIL_WEBDAV_PASSWORD="app-password-with-files-scope"

# ...or any WebDAV server, e.g. Nextcloud:
# export FASTMAIL_WEBDAV_URL="https://cloud.example.com/remote.php/dav/files/USERNAME/"
```

The URL must be HTTPS. Note: Fastmail Files ignores the WebDAV `If-None-Match` precondition, so the tool performs an explicit existence check before non-overwrite uploads.

### Email attachments on send

`send_email`, `create_draft`, `edit_draft`, and `reply_email` accept an `attachments` array. Each entry uses exactly one source:

- `{ "localPath": "report.pdf" }` — a file inside `FASTMAIL_DOWNLOAD_DIR` (same confinement as downloads)
- `{ "emailId": "...", "attachmentId": "..." }` — re-attach from an existing email (zero-copy: no bytes are transferred)
- `{ "blobId": "...", "name": "...", "type": "..." }` — an already-uploaded JMAP blob

Uploads respect the server's `maxSizeUpload` (~50 MB on Fastmail). Editing a draft preserves its existing attachments.

### Contacts write scope

`create_contact` / `update_contact` / `delete_contact` need the API token to have **read-write** contacts scope (Settings → Privacy & Security → API tokens). Read-only tokens keep the three read tools working and fail writes with a `forbidden` error.

## Development

### Project Structure
```
src/
├── index.ts               # Main MCP server implementation
├── auth.ts                # Authentication handling
├── jmap-client.ts         # JMAP client wrapper
├── contacts-calendar.ts   # Contacts extensions (JMAP)
├── caldav-client.ts       # CalDAV calendar client (the calendar path — JMAP calendars are not available)
├── webdav-files-client.ts # WebDAV file storage client (save_attachment_to_webdav)
├── url-validation.ts      # Base-URL allowlist / HTTPS validation
├── coerce.ts              # Input coercion helpers
└── *.test.ts              # Unit tests (colocated)
```

### Building
```bash
npm run build
```

### Development Mode
```bash
npm run dev
```

## License

MIT

## Contributing

Contributions are welcome! Please ensure that:
1. Code follows the existing style
2. All functions are properly typed
3. Error handling is implemented
4. Documentation is updated for new features

## Troubleshooting

### Common Issues

1. **Authentication Errors**: Ensure your API token is valid and has the necessary permissions
2. **Missing Dependencies**: Run `npm install` to ensure all dependencies are installed  
3. **Build Errors**: Check that TypeScript compilation completes without errors using `npm run build`
4. **Calendar/Contacts "Forbidden" Errors**: Use `check_function_availability` to see setup guidance

### Email Tools Failing with Serialization Errors?

If `get_email`, `list_emails`, `search_emails`, or `advanced_search` fail with "content serialization" or "Cannot read properties of undefined" errors, upgrade to v1.7.1 or later (any current release includes the fix). This was caused by incomplete JMAP response validation that surfaced after the MCP SDK v1.x upgrade added stricter result checking.

### Calendar Not Working?

Calendar tools run over CalDAV, not JMAP. If they return "CalDAV not configured", set `FASTMAIL_CALDAV_USERNAME` and `FASTMAIL_CALDAV_PASSWORD` (see [CalDAV Calendar Support](#caldav-calendar-support)).

### Contacts Not Working?

If contacts functions return "Forbidden" errors:

1. **API Token Scope**: writes (`create_contact`/`update_contact`/`delete_contact`) need read-write contacts scope (see [Contacts write scope](#contacts-write-scope))
2. **Account Plan**: the contacts API may require certain Fastmail plans

`Contact not found` / `Calendar event not found` errors mean the ID is stale — re-list and retry.

**Solution**: Run `check_function_availability` for step-by-step setup guidance.

### Testing Your Setup

Use the built-in testing tools:
- **check_function_availability**: See what's available and get setup help
- **test_bulk_operations**: Safely test bulk operations without making changes

For more detailed error information, check the console output when running the server.

## Privacy & Security

- API tokens are stored encrypted by Claude Desktop when installed via the DXT and are never logged by this server.
- The server avoids logging raw errors and sensitive data (tokens, email addresses, identities, attachment names/blobIds) in error messages.
- Tool responses may include your email metadata/content by design (e.g., listing emails) but internal identifiers and credentials are not disclosed beyond what Fastmail returns for the requested data.
- If you encounter errors, messages are sanitized and summarized to prevent leaking personal information.
