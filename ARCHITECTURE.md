# Agent 365 MCP - Architecture Documentation

This document describes the technical architecture of the Agent 365 MCP proxy server.

## System Overview

```
┌─────────────────┐     ┌──────────────────────┐     ┌─────────────────────┐
│                 │     │                      │     │                     │
│   MCP Client    │◄───►│  Agent 365 MCP Proxy │◄───►│  Microsoft Agent365 │
│  (Any Editor)   │     │     (This Server)    │     │    Backend APIs     │
│                 │     │                      │     │                     │
└─────────────────┘     └──────────────────────┘     └─────────────────────┘
       │                         │                           │
       │ stdio                   │ HTTPS + SSE               │
       │ JSON-RPC                │ JSON-RPC over SSE         │
       │                         │                           │
       └─────────────────────────┴───────────────────────────┘
```

## Component Architecture

### MCP Protocol Layer

The server implements the Model Context Protocol (MCP) specification:

- **Transport**: stdio (stdin/stdout)
- **Protocol**: JSON-RPC 2.0
- **Capabilities**: tools (list and call)

```javascript
// MCP message flow
Client → Server: { jsonrpc: "2.0", method: "tools/list", id: 1 }
Server → Client: { jsonrpc: "2.0", result: { tools: [...] }, id: 1 }
```

### Server Aggregation

The proxy aggregates 8 Microsoft Agent 365 backend servers:

| Server Key | Endpoint | Prefix | Primary Use |
|------------|----------|--------|-------------|
| sharepoint | mcp_ODSPRemoteServer | sharepoint_ | Files, folders |
| word | mcp_WordServer | word_ | Document content |
| teams | mcp_TeamsServer | teams_ | Chats, messages |
| copilot | mcp_M365Copilot | copilot_ | Semantic search |
| mail | mcp_MailTools | mail_ | Email |
| calendar | mcp_CalendarTools | calendar_ | Events |
| me | mcp_MeServer | me_ | User profile |
| excel | mcp_ExcelServer | excel_ | Spreadsheets |

### Tool Naming Convention

Tools are prefixed to avoid collisions:
- Original: `findFileOrFolder`
- Exposed: `sharepoint_findFileOrFolder`

### Authentication Flow

```
┌──────────┐     ┌─────────────┐     ┌──────────────┐     ┌─────────────┐
│  User    │     │  CLI Auth   │     │  Entra ID    │     │  Token File │
└────┬─────┘     └──────┬──────┘     └──────┬───────┘     └──────┬──────┘
     │                  │                   │                    │
     │ npx auth         │                   │                    │
     │─────────────────►│                   │                    │
     │                  │ Device code req   │                    │
     │                  │──────────────────►│                    │
     │                  │ Code + URL        │                    │
     │                  │◄──────────────────│                    │
     │ Visit URL        │                   │                    │
     │◄─────────────────│                   │                    │
     │ Sign in          │                   │                    │
     │─────────────────────────────────────►│                    │
     │                  │ Poll for token    │                    │
     │                  │──────────────────►│                    │
     │                  │ Access token      │                    │
     │                  │◄──────────────────│                    │
     │                  │                   │  Save token        │
     │                  │──────────────────────────────────────►│
     │                  │                   │                    │
```

Token storage:
- Primary: `~/.agent365-mcp/tokens.json`
- Legacy: `~/.claude/agent365-tokens.json`

### Request Flow

```
┌─────────────┐     ┌─────────────┐     ┌──────────────┐     ┌──────────────┐
│ MCP Client  │     │ MCP Proxy   │     │ Agent 365 API│     │ Microsoft 365│
└──────┬──────┘     └──────┬──────┘     └──────┬───────┘     └──────┬───────┘
       │                   │                   │                    │
       │ tools/call        │                   │                    │
       │──────────────────►│                   │                    │
       │                   │ Initialize        │                    │
       │                   │──────────────────►│                    │
       │                   │ OK                │                    │
       │                   │◄──────────────────│                    │
       │                   │ tools/call        │                    │
       │                   │──────────────────►│                    │
       │                   │                   │ Graph API call     │
       │                   │                   │───────────────────►│
       │                   │                   │ Data               │
       │                   │                   │◄───────────────────│
       │                   │ SSE response      │                    │
       │                   │◄──────────────────│                    │
       │ MCP result        │                   │                    │
       │◄──────────────────│                   │                    │
```

## Data Flow

### Large Response Handling

```
┌───────────────────────────────────────────────────────────────────────┐
│                        Response Processing                             │
└───────────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
                    ┌───────────────────────────────┐
                    │ Is response > threshold?      │
                    │ (LARGE_FILE_THRESHOLD=100KB)  │
                    └───────────────┬───────────────┘
                                    │
                    ┌───────────────┴───────────────┐
                    │                               │
                    ▼ Yes                           ▼ No
    ┌───────────────────────────┐       ┌───────────────────────┐
    │ LARGE_FILE_DIR set?       │       │ Return as-is          │
    └───────────────┬───────────┘       └───────────────────────┘
                    │
        ┌───────────┴───────────┐
        │                       │
        ▼ Yes                   ▼ No
┌───────────────────┐   ┌───────────────────────────────┐
│ Save to file      │   │ Is > MAX_RESPONSE_SIZE?       │
│ Return filepath   │   │ (50KB default)                │
└───────────────────┘   └───────────────┬───────────────┘
                                        │
                            ┌───────────┴───────────┐
                            │                       │
                            ▼ Yes                   ▼ No
                    ┌───────────────────┐   ┌───────────────────┐
                    │ Truncate + hint   │   │ Return as-is      │
                    └───────────────────┘   └───────────────────┘
```

### Session Management

```javascript
// Server state structure
serverSessions = {
  sharepoint: {
    initialized: true,
    tools: [/* cached tool definitions */]
  },
  copilot: {
    initialized: false  // Not yet accessed
  },
  // etc.
}

// Disabled servers (license issues)
disabledServers = Set { "copilot" }  // If user lacks Copilot license
```

## Configuration Reference

### Environment Variables

| Variable | Type | Default | Description |
|----------|------|---------|-------------|
| `AGENT365_TENANT_ID` | string | required | Azure tenant ID |
| `AGENT365_CLIENT_ID` | string | required | App registration ID |
| `AGENT365_TOKEN_PATH` | string | ~/.agent365-mcp/tokens.json | Token file location |
| `AGENT365_MAX_RESPONSE_SIZE` | int | 50000 | Truncation threshold (chars) |
| `AGENT365_LARGE_FILE_THRESHOLD` | int | 100000 | File save threshold (chars) |
| `AGENT365_LARGE_FILE_DIR` | string | "" | Directory for large responses |
| `AGENT365_TIMEOUT` | int | 60000 | Request timeout (ms) |
| `AGENT365_ALLOW_DANGEROUS` | bool | false | Enable delete/remove tools |
| `AGENT365_DISABLED_SERVERS` | string | "" | Comma-separated server list |

### Tool Filtering

Dangerous tool patterns (disabled by default):
- `/delete/i`
- `/remove/i`
- `/destroy/i`
- `/purge/i`
- `/erase/i`
- `/drop/i`

## Error Handling

### License Errors

When a server returns a license error, it's automatically disabled:

```javascript
// Error patterns that trigger server disabling
- "license"
- "not entitled"
- "Copilot"
- "subscription"
- HTTP 401/403
```

### Graceful Degradation

```
┌─────────────────────────────────────────────────────────────────────┐
│                     Server Initialization                            │
└─────────────────────────────────────────────────────────────────────┘
                                │
                    ┌───────────┴───────────────┐
                    │ Try initialize server     │
                    └───────────────┬───────────┘
                                    │
                    ┌───────────────┴───────────────┐
                    │                               │
                    ▼ Success                       ▼ License Error
            ┌───────────────────┐       ┌───────────────────────┐
            │ Mark initialized  │       │ Add to disabledServers│
            │ Fetch tools       │       │ Log warning           │
            └───────────────────┘       │ Continue with others  │
                                        └───────────────────────┘
```

## File Structure

```
agent365-mcp-proxy/
├── index.js          # Main MCP server implementation
├── cli.js            # CLI entry point (auth, serve, status)
├── package.json      # Package configuration
├── README.md         # User documentation
├── ADMIN.md          # IT admin setup guide
├── architecture.md   # This file
└── agents.md         # Agent integration guide
```

## Performance Considerations

### Parallel Server Initialization

Tools from all servers are fetched in parallel:

```javascript
const results = await Promise.allSettled(
  serverKeys.map(key => getServerTools(key))
);
```

### Response Buffering

- Max buffer: 10MB (prevents memory issues)
- Responses exceeding buffer are partially processed

### Tool Caching

Tool definitions are cached per session to avoid repeated list requests.

## Security Model

### Token Security

- Tokens stored locally, not transmitted to third parties
- 1-hour expiration with refresh on next use
- No passwords stored

### Dangerous Operation Blocking

1. Tool-level: Tools matching dangerous patterns are filtered
2. Argument-level: Arguments containing "delete"/"remove" are blocked
3. Configurable: Can be enabled with `AGENT365_ALLOW_DANGEROUS=true`

### Data Access

- All access respects user's M365 permissions
- No elevation of privilege
- Audit trail in M365 admin center
