# Fastmail MCP Server

A Model Context Protocol (MCP) server that provides access to the Fastmail API, enabling AI assistants to interact with email, contacts, and calendar data.

## Features

### Email Operations
- List mailboxes
- List emails from specific mailboxes or all mailboxes
- Get specific emails by ID
- Send emails (text and HTML)
- Search emails by content

### Contacts Operations
- List all contacts
- Get specific contacts by ID
- Search contacts by name or email

### Calendar Operations
- List all calendars
- List calendar events
- Get specific calendar events by ID
- Create new calendar events

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