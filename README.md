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

## MCP Client Configuration

Add this server to your MCP client configuration. For example, with Claude Desktop:

```json
{
  "mcpServers": {
    "fastmail": {
      "command": "node",
      "args": ["/path/to/fastmail-mcp/dist/index.js"],
      "env": {
        "FASTMAIL_API_TOKEN": "your_api_token_here"
      }
    }
  }
}
```

## Available Tools

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

### Identity Tools

- **list_identities**: List sending identities (email addresses that can be used for sending)

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

For more detailed error information, check the console output when running the server.