# Fastmail MCP Server

A Model Context Protocol (MCP) server that provides access to the Fastmail API, enabling AI assistants to interact with email, contacts, and calendar data.

## Features

### Core Email Operations
- List mailboxes and get mailbox statistics
- List, search, and filter emails with advanced criteria
- Get specific emails by ID with full content
- Send emails (text and HTML) with proper draft/sent handling
- Email management: mark read/unread, delete, move between folders

### Advanced Email Features
- **Attachment Handling**: List and download email attachments
- **Threading Support**: Get complete conversation threads
- **Advanced Search**: Multi-criteria filtering (sender, date range, attachments, read status)
- **Bulk Operations**: Process multiple emails simultaneously
- **Statistics & Analytics**: Account summaries and mailbox statistics

### Contacts Operations
- List all contacts with full contact information
- Get specific contacts by ID
- Search contacts by name or email

### Calendar Operations
- List all calendars and calendar events
- Get specific calendar events by ID
- Create new calendar events with participants and details

### Identity & Account Management
- List available sending identities
- Account summary with comprehensive statistics

## Setup

### Prerequisites
- Node.js 18+ 
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
   export FASTMAIL_BASE_URL="https://api.fastmail.com"
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

### Run via npx (GitHub)

Default to `main` branch:

```bash
FASTMAIL_API_TOKEN="your_token" FASTMAIL_BASE_URL="https://api.fastmail.com" \
  npx --yes github:MadLlama25/fastmail-mcp fastmail-mcp
```

Windows PowerShell:

```powershell
$env:FASTMAIL_API_TOKEN="your_token"
$env:FASTMAIL_BASE_URL="https://api.fastmail.com"
npx --yes github:MadLlama25/fastmail-mcp fastmail-mcp
```

Pin to a tagged release:

```bash
FASTMAIL_API_TOKEN="your_token" \
  npx --yes github:MadLlama25/fastmail-mcp@v1.6.1 fastmail-mcp
```

## Install as a Claude Desktop Extension (DXT)

You can install this server as a Desktop Extension for Claude Desktop using the packaged `.dxt` file.

1. Build and pack:
   ```bash
   npm run build
   npx dxt pack
   ```
   This produces `fastmail-mcp.example.com` in the project root.

2. Install into Claude Desktop:
   - Open the `.dxt` file, or drag it into Claude Desktop
   - When prompted:
     - Fastmail API Token: paste your token (stored encrypted by Claude)
     - Fastmail Base URL: leave blank to use `https://api.fastmail.com` (default)

3. Use any of the tools (e.g. `get_recent_emails`).

## Run as a remote HTTPS (SSE) connector (beta)

This mode enables multi-user connections from Claude Web/Desktop custom connectors. Each user provides their own Fastmail API token in the Claude client; the server does not store tokens.

1. Build:
   ```bash
   npm run build
   ```
2. Run (SSE mode):
   ```bash
   MCP_TRANSPORT=sse PORT=3000 HOST=0.0.0.0 \
   AUTH_HEADER=Authorization AUTH_SCHEME=Bearer \
   FASTMAIL_BASE_URL="https://api.fastmail.com" \
   node dist/index.js
   ```
3. Add a custom connector in Claude (HTTPS):
   - URL: `https://<your-domain>/mcp`
   - Secret: paste your Fastmail API token (it will be forwarded as `Authorization: Bearer <token>`)

Notes:
- The server keeps tokens in-memory per connection only. No persistence.
- Keep DXT/npx flows unchanged by omitting `MCP_TRANSPORT` (defaults to stdio).
- You can optionally require `X-Connector-Secret` to match `CONNECTOR_SHARED_SECRET` for extra protection.

### Docker Compose (example)

Create `docker-compose.yml` and `.env` (do not commit secrets):

```yaml
services:
  app:
    build: .
    image: fastmail-mcp:beta
    restart: unless-stopped
    environment:
      MCP_TRANSPORT: sse
      PORT: 3000
      HOST: 0.0.0.0
      AUTH_HEADER: Authorization
      AUTH_SCHEME: Bearer
      CONNECTOR_SHARED_SECRET: ${CONNECTOR_SHARED_SECRET}
      FASTMAIL_BASE_URL: ${FASTMAIL_BASE_URL}
    ports:
      - "3000:3000"
```

Put this behind a TLS reverse proxy (Caddy/Nginx) and expose `https://` on your domain.

#### Caddy example (SSE)
```
fastmail-mcp.example.com {
  encode zstd gzip
  reverse_proxy /mcp app:3000
  reverse_proxy /mcp/messages app:3000
  reverse_proxy /healthz app:3000
}
```

### .env file (example)
Create a `.env` file alongside `docker-compose.yml`:
```env
FASTMAIL_BASE_URL=https://api.fastmail.com
# Optional extra gate for WS connects; comment out to disable
# CONNECTOR_SHARED_SECRET=your_shared_secret
```

### Rate limits and safety
- Tools are unchanged; respect Fastmail API limits. Prefer small `limit` values and staggered bulk ops.
- The server does not cache responses containing personal data.
- Built-in gentle concurrency limits (2 in-flight per connection) and exponential backoff on 429/5xx help avoid rate limits.
 - Logs are JSON lines; keep 7 days by rotating via your process manager or logrotate.

## Available Tools (31 Total)

**🎯 Most Popular Tools:**
- **check_function_availability**: Check what's available and get setup guidance  
- **test_bulk_operations**: Safely test bulk operations with dry-run mode
- **send_email**: Full-featured email sending with proper draft/sent handling
- **advanced_search**: Powerful multi-criteria email filtering
- **get_recent_emails**: Quick access to recent emails from any mailbox

### Email Tools

- **list_mailboxes**: Get all mailboxes in your account
- **list_emails**: List emails from a specific mailbox or all mailboxes
  - Parameters: `mailboxId` (optional), `limit` (default: 20)
- **get_email**: Get a specific email by ID
  - Parameters: `emailId` (required)
- **send_email**: Send an email
  - Parameters: `to` (required array), `cc` (optional array), `bcc` (optional array), `from` (optional), `mailboxId` (optional), `subject` (required), `textBody` (optional), `htmlBody` (optional)
- **search_emails**: Search emails by content
  - Parameters: `query` (required), `limit` (default: 20)
- **get_recent_emails**: Get the most recent emails from a mailbox (inspired by JMAP-Samples top-ten)
  - Parameters: `limit` (default: 10, max: 50), `mailboxName` (default: 'inbox')
- **mark_email_read**: Mark an email as read or unread
  - Parameters: `emailId` (required), `read` (default: true)
- **delete_email**: Delete an email (move to trash)
  - Parameters: `emailId` (required)
- **move_email**: Move an email to a different mailbox
  - Parameters: `emailId` (required), `targetMailboxId` (required)

### Advanced Email Features

- **get_email_attachments**: Get list of attachments for an email
  - Parameters: `emailId` (required)
- **download_attachment**: Get download URL for an email attachment
  - Parameters: `emailId` (required), `attachmentId` (required)
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

## Development

### Project Structure
```
src/
├── index.ts              # Main MCP server implementation
├── auth.ts              # Authentication handling
├── jmap-client.ts       # JMAP client wrapper
└── contacts-calendar.ts # Contacts and calendar extensions
└── handlers.ts          # Tool registry (shared by stdio/ws)
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

### Calendar/Contacts Not Working?

If calendar and contacts functions return "Forbidden" errors, this is likely due to:

1. **Account Plan**: Calendar/contacts API may require business/professional Fastmail plans
2. **API Token Scope**: Your API token may need calendar/contacts permissions enabled
3. **Feature Enablement**: These features may need explicit activation in your account

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