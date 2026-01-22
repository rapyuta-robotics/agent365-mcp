# Agent 365 MCP

MCP (Model Context Protocol) server for Microsoft 365 integration with AI coding assistants.

Access SharePoint, Word, Teams, Outlook, Calendar, Excel, and M365 Copilot directly from any MCP-compatible client.

## Supported Clients

- **Claude Code** (Anthropic CLI)
- **Claude Desktop** (Anthropic app)
- **VS Code** with GitHub Copilot
- **GitHub Copilot** with MCP support
- **Cursor**
- Any MCP-compatible AI assistant

## Features

- **SharePoint & OneDrive**: Read/write files, search, share
- **Word**: Read documents, add comments, create new docs
- **Teams**: Read chats, channels, post messages
- **Outlook Mail**: Read/send emails
- **Calendar**: Manage events
- **Excel**: Read/write spreadsheets
- **M365 Copilot**: Search and chat
- **User Profile**: Org chart, manager info

80+ tools available. Dangerous operations (delete/remove) are filtered for safety.

## Prerequisites

1. **Copilot for Microsoft 365 license** assigned to your account
2. **IT Admin setup** (one-time per organization):
   - Create Entra ID app registration
   - Grant admin consent for Agent 365 permissions
   - Provide tenant ID and client ID to users

## Quick Start

### Option 1: Zero-Config (If IT provided env vars)

Just add to your MCP config - authentication happens automatically on first use:

```json
{
  "mcpServers": {
    "agent365": {
      "command": "npx",
      "args": ["-y", "github:rapyuta-robotics/agent365-mcp", "serve"],
      "env": {
        "AGENT365_TENANT_ID": "your-tenant-id",
        "AGENT365_CLIENT_ID": "your-client-id"
      }
    }
  }
}
```

On first use, your browser opens automatically for Microsoft login. That's it!

### Option 2: Interactive Setup

Run the setup wizard which configures everything for you:

```bash
npx github:rapyuta-robotics/agent365-mcp setup
```

This will:
1. Prompt for your Tenant ID and Client ID (get from IT admin)
2. Open Microsoft login in your browser
3. Auto-configure Claude Code and/or VS Code
4. Store refresh token for ~90 day sessions

**That's it!** Restart your coding assistant and the tools are available.

---

## Manual Installation

If you prefer manual setup:

### 1. Authenticate

```bash
export AGENT365_TENANT_ID="your-tenant-id"
export AGENT365_CLIENT_ID="your-client-id"
npx github:rapyuta-robotics/agent365-mcp auth
```

### 2. Configure your MCP client

#### Claude Code (`~/.claude.json`)

```json
{
  "mcpServers": {
    "agent365": {
      "type": "stdio",
      "command": "npx",
      "args": ["-y", "github:rapyuta-robotics/agent365-mcp", "serve"],
      "env": {
        "AGENT365_TENANT_ID": "your-tenant-id",
        "AGENT365_CLIENT_ID": "your-client-id"
      }
    }
  }
}
```

#### VS Code

> **Quick way:** `Ctrl+Shift+P` (or `Cmd+Shift+P` on Mac) → `MCP: Open User Configuration` to directly edit your MCP config file.

Add this to your config (`.vscode/mcp.json` or user settings):

```json
{
  "mcpServers": {
    "agent365": {
      "command": "npx",
      "args": ["-y", "github:rapyuta-robotics/agent365-mcp", "serve"],
      "env": {
        "AGENT365_TENANT_ID": "your-tenant-id",
        "AGENT365_CLIENT_ID": "your-client-id"
      }
    }
  }
}
```

**Alternative - Add MCP Server UI:**

1. `Ctrl+Shift+P` → `MCP: Add Server`
2. Select: `Command (stdio)`
3. Command: `npx`
4. Args: `-y github:rapyuta-robotics/agent365-mcp serve`
5. Name: `agent365`
6. Scope: `User Settings` or `Workspace`
7. Then edit config to add the `env` section above

#### Claude Desktop

Config file location:
- **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`

Or: Open Claude Desktop → Settings → Developer → Edit Config

```json
{
  "mcpServers": {
    "agent365": {
      "command": "npx",
      "args": ["-y", "github:rapyuta-robotics/agent365-mcp", "serve"],
      "env": {
        "AGENT365_TENANT_ID": "your-tenant-id",
        "AGENT365_CLIENT_ID": "your-client-id"
      }
    }
  }
}
```

After saving, **completely quit and restart Claude Desktop** to load the new configuration.

#### GitHub Copilot / Cursor

Use similar MCP configuration in your editor's settings. Consult your editor's MCP documentation for the exact location.

3. **Restart your editor** - The tools will be available.

### Check Status

```bash
npx github:rapyuta-robotics/agent365-mcp status
```

### Re-authenticate (when token expires)

```bash
npx github:rapyuta-robotics/agent365-mcp auth
```

## IT Admin Setup

See **[ADMIN.md](ADMIN.md)** for detailed setup instructions.

### Quick Start

1. **Create Service Principal** (one-time)

```powershell
curl -O https://raw.githubusercontent.com/microsoft/Agent365-devTools/main/scripts/cli/Auth/New-Agent365ToolsServicePrincipalProdPublic.ps1
pwsh ./New-Agent365ToolsServicePrincipalProdPublic.ps1
```

2. **Create App Registration**

```bash
az ad app create --display-name "Agent 365 MCP" --sign-in-audience "AzureADMyOrg"
az ad app update --id <app-id> --is-fallback-public-client true
az ad app update --id <app-id> --public-client-redirect-uris "http://localhost"
az ad sp create --id <app-id>
```

3. **Add Permissions & Grant Consent**

See [ADMIN.md](ADMIN.md) for complete permission setup and admin consent commands.

4. **Distribute to Users**

Provide users with:
- **Tenant ID**: `az account show --query tenantId -o tsv`
- **Client ID**: The app ID from step 2

## Available Tools

| Prefix | Service | Tools |
|--------|---------|-------|
| `sharepoint_*` | SharePoint/OneDrive | File operations, search, share |
| `word_*` | Word | Read docs, comments, create |
| `teams_*` | Teams | Chats, channels, messages |
| `mail_*` | Outlook | Email read/send/reply |
| `calendar_*` | Calendar | Events, scheduling |
| `me_*` | Profile | User/org info |
| `excel_*` | Excel | Spreadsheet operations |
| `copilot_*` | M365 Copilot | Search, chat |

## Security

- Each user authenticates with their own M365 account
- Access is limited to what the user can access in M365
- Dangerous operations (delete, remove) are blocked by default
- Tokens stored locally in `~/.agent365-mcp/`
- Tokens expire after 1 hour and auto-refresh on next use

## Configuration

Optional environment variables to customize behavior:

| Variable | Default | Description |
|----------|---------|-------------|
| `AGENT365_TENANT_ID` | (required) | Azure tenant ID |
| `AGENT365_CLIENT_ID` | (required) | App registration client ID |
| `AGENT365_TOKEN_PATH` | `~/.agent365-mcp/tokens.json` | Custom token storage path |
| `AGENT365_MAX_RESPONSE_SIZE` | `50000` | Max characters before truncation |
| `AGENT365_TIMEOUT` | `60000` | Request timeout in milliseconds |
| `AGENT365_LARGE_FILE_DIR` | (none) | Directory to save large responses |
| `AGENT365_LARGE_FILE_THRESHOLD` | `100000` | Size threshold for file save |
| `AGENT365_ALLOW_DANGEROUS` | `false` | Enable dangerous tools (delete/remove) |
| `AGENT365_DISABLED_SERVERS` | (none) | Comma-separated servers to disable |

### Large Response Handling

By default, responses over 50KB are truncated. For better handling of large data:

```json
{
  "env": {
    "AGENT365_LARGE_FILE_DIR": "/tmp/agent365-responses"
  }
}
```

When set, responses over 100KB are saved to files instead of truncated. The agent receives the filepath and can use Read tools to access the content.

### Disabling Servers

If you don't have Copilot license or want to disable specific services:

```json
{
  "env": {
    "AGENT365_DISABLED_SERVERS": "copilot,excel"
  }
}
```

The proxy also auto-detects license errors and gracefully disables affected servers.

### Full Configuration Example

```json
{
  "mcpServers": {
    "agent365": {
      "command": "npx",
      "args": ["-y", "github:rapyuta-robotics/agent365-mcp", "serve"],
      "env": {
        "AGENT365_TENANT_ID": "your-tenant-id",
        "AGENT365_CLIENT_ID": "your-client-id",
        "AGENT365_MAX_RESPONSE_SIZE": "100000",
        "AGENT365_TIMEOUT": "120000",
        "AGENT365_LARGE_FILE_DIR": "/tmp/agent365",
        "AGENT365_DISABLED_SERVERS": "copilot"
      }
    }
  }
}
```

## Troubleshooting

### "Scope not present" error
Your token doesn't have the required permissions. Re-authenticate:
```bash
npx github:rapyuta-robotics/agent365-mcp auth
```

### "No valid token" error
Token expired or not authenticated. Run:
```bash
npx github:rapyuta-robotics/agent365-mcp auth
```

### Tools not loading
1. Check status: `npx github:rapyuta-robotics/agent365-mcp status`
2. Verify Copilot license is assigned to your account
3. Ask IT admin to verify admin consent was granted

### Large response errors / Content truncated
Some M365 queries return very large responses. The proxy automatically handles this:
1. Use more specific queries (date ranges, filters, pagination)
2. Set `AGENT365_LARGE_FILE_DIR` to save large responses to files
3. Increase `AGENT365_MAX_RESPONSE_SIZE` for larger truncation limit
4. Use M365 Copilot search instead of listing all items

### Request timeout
Long-running queries may timeout. Increase timeout with `AGENT365_TIMEOUT` (in milliseconds).

## Documentation

- **[ADMIN.md](ADMIN.md)** - IT administrator setup guide with detailed Azure CLI commands
- **[ARCHITECTURE.md](ARCHITECTURE.md)** - Technical architecture, data flows, and system design
- **[AGENTS.md](AGENTS.md)** - AI agent integration guide with tool selection patterns

## License

MIT
