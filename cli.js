#!/usr/bin/env node
const { PublicClientApplication } = require("@azure/msal-node");
const fs = require("fs");
const path = require("path");

// Default configuration - can be overridden via environment variables
const CONFIG = {
  TENANT_ID: process.env.AGENT365_TENANT_ID || process.env.TENANT_ID,
  CLIENT_ID: process.env.AGENT365_CLIENT_ID || process.env.CLIENT_ID,
  AGENT365_API: "ea9ffc3e-8a23-4a7d-836d-234d7c7565c1",
};

// Token storage location
function getTokenPath() {
  const home = process.env.HOME || process.env.USERPROFILE || "";
  return path.join(home, ".agent365-mcp", "tokens.json");
}

function getConfigPath() {
  const home = process.env.HOME || process.env.USERPROFILE || "";
  return path.join(home, ".agent365-mcp", "config.json");
}

function loadConfig() {
  const configPath = getConfigPath();
  if (fs.existsSync(configPath)) {
    try {
      return JSON.parse(fs.readFileSync(configPath, "utf8"));
    } catch (e) {
      // Ignore
    }
  }
  return {};
}

function saveConfig(config) {
  const configPath = getConfigPath();
  const dir = path.dirname(configPath);
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
  fs.writeFileSync(configPath, JSON.stringify(config, null, 2));
}

function getMsalCachePath() {
  const home = process.env.HOME || process.env.USERPROFILE || "";
  return path.join(home, ".agent365-mcp", "msal-cache.json");
}

function createMsalCachePlugin() {
  const cachePath = getMsalCachePath();

  const beforeCacheAccess = async (cacheContext) => {
    try {
      if (fs.existsSync(cachePath)) {
        cacheContext.tokenCache.deserialize(fs.readFileSync(cachePath, "utf8"));
      }
    } catch (e) {
      // Ignore cache read errors
    }
  };

  const afterCacheAccess = async (cacheContext) => {
    if (cacheContext.cacheHasChanged) {
      const dir = path.dirname(cachePath);
      if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
      }
      fs.writeFileSync(cachePath, cacheContext.tokenCache.serialize());
    }
  };

  return { beforeCacheAccess, afterCacheAccess };
}

async function authenticate(tenantId, clientId) {
  const tokenPath = getTokenPath();
  const dir = path.dirname(tokenPath);
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }

  const pca = new PublicClientApplication({
    auth: {
      clientId: clientId,
      authority: `https://login.microsoftonline.com/${tenantId}`,
    },
    cache: {
      cachePlugin: createMsalCachePlugin(),
    },
  });

  console.log("\n🔐 Microsoft 365 Agent Authentication\n");
  console.log("This will authenticate you with your organization's Microsoft 365 account.");
  console.log("Your access will be limited to what your account has permission to view.\n");

  const scopes = [`${CONFIG.AGENT365_API}/.default`];
  let response;

  // Try silent auth first (uses refresh token if available)
  const accounts = await pca.getTokenCache().getAllAccounts();
  if (accounts.length > 0) {
    try {
      console.log(`🔄 Found existing session for ${accounts[0].username}, refreshing...`);
      response = await pca.acquireTokenSilent({
        account: accounts[0],
        scopes: scopes,
      });
      console.log("✅ Token refreshed silently!\n");
    } catch (e) {
      console.log("⚠️  Silent refresh failed, starting interactive login...\n");
      response = null;
    }
  }

  // Fall back to device code if silent auth failed
  if (!response) {
    response = await pca.acquireTokenByDeviceCode({
      scopes: scopes,
      deviceCodeCallback: (resp) => {
        const code = resp.userCode;
        const completeUrl = resp.verificationUriComplete || resp.verificationUri;

        console.log("\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
        console.log("📱 AUTHENTICATION REQUIRED");
        console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
        console.log(`\n  Your code: ${code}\n`);
        console.log(`  Go to: ${completeUrl}\n`);
        console.log("  Enter the code above and sign in with Microsoft.\n");
        console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n");
      },
    });
  }

  // Save token info (access token for quick access, MSAL cache has refresh token)
  fs.writeFileSync(tokenPath, JSON.stringify({
    accessToken: response.accessToken,
    expiresOn: response.expiresOn,
    account: response.account,
    tenantId: tenantId,
    clientId: clientId,
  }));

  // Save config for future use
  saveConfig({ tenantId, clientId });

  console.log("✅ Authentication successful!");
  console.log(`📧 Logged in as: ${response.account?.username || "Unknown"}`);
  console.log(`⏰ Access token expires: ${response.expiresOn}`);
  console.log(`🔄 Refresh token valid for ~90 days`);
  console.log(`📁 Token saved to: ${tokenPath}\n`);
  console.log("You can now use the Agent 365 MCP tools in your MCP client.\n");
}

function showHelp() {
  console.log(`
Agent 365 MCP - Microsoft 365 Integration for AI Coding Assistants

USAGE:
  agent365-mcp setup             Interactive setup (recommended for first use)
  agent365-mcp auth              Authenticate with Microsoft 365
  agent365-mcp serve             Start the MCP server (used by MCP clients)
  agent365-mcp status            Check authentication status
  agent365-mcp logout            Remove saved authentication

QUICK START:
  npx github:rapyuta-robotics/agent365-mcp setup

AUTHENTICATION:
  Before using, authenticate with your Microsoft 365 account:

    npx github:rapyuta-robotics/agent365-mcp auth

MCP CLIENT CONFIGURATION:
  Add to your MCP client config (e.g., ~/.claude.json, .vscode/mcp.json):

  {
    "mcpServers": {
      "agent365": {
        "type": "stdio",
        "command": "npx",
        "args": ["-y", "github:rapyuta-robotics/agent365-mcp", "serve"],
        "env": {
          "AGENT365_TENANT_ID": "<your-tenant-id>",
          "AGENT365_CLIENT_ID": "<your-client-id>"
        }
      }
    }
  }

ENVIRONMENT VARIABLES:
  AGENT365_TENANT_ID    Microsoft Entra tenant ID
  AGENT365_CLIENT_ID    Application (client) ID from Entra app registration

AVAILABLE TOOLS (80+):
  sharepoint_*    SharePoint & OneDrive file operations
  word_*          Word document read/write/comments
  teams_*         Teams chats, channels, messages
  mail_*          Outlook email operations
  calendar_*      Calendar event management
  me_*            User profile and org info
  excel_*         Excel spreadsheet operations
  copilot_*       M365 Copilot search

For more info: https://github.com/rapyuta-robotics/agent365-mcp
`);
}

function showStatus() {
  const tokenPath = getTokenPath();
  const config = loadConfig();

  console.log("\n📊 Agent 365 MCP Status\n");

  if (config.tenantId) {
    console.log(`🏢 Tenant ID: ${config.tenantId}`);
    console.log(`🔑 Client ID: ${config.clientId}`);
  } else {
    console.log("⚠️  No configuration saved. Run 'agent365-mcp auth' first.");
  }

  if (fs.existsSync(tokenPath)) {
    try {
      const token = JSON.parse(fs.readFileSync(tokenPath, "utf8"));
      const expiresOn = new Date(token.expiresOn);
      const now = new Date();

      if (expiresOn > now) {
        console.log(`\n✅ Authenticated`);
        console.log(`📧 Account: ${token.account?.username || "Unknown"}`);
        console.log(`⏰ Expires: ${expiresOn.toLocaleString()}`);
      } else {
        console.log(`\n⚠️  Token expired at ${expiresOn.toLocaleString()}`);
        console.log("   Run 'agent365-mcp auth' to re-authenticate.");
      }
    } catch (e) {
      console.log("\n❌ Token file corrupted. Run 'agent365-mcp auth' to fix.");
    }
  } else {
    console.log("\n❌ Not authenticated. Run 'agent365-mcp auth' first.");
  }
  console.log("");
}

function logout() {
  const tokenPath = getTokenPath();
  const msalCachePath = getMsalCachePath();

  if (fs.existsSync(tokenPath)) {
    fs.unlinkSync(tokenPath);
  }
  if (fs.existsSync(msalCachePath)) {
    fs.unlinkSync(msalCachePath);
  }
  console.log("✅ Logged out successfully. Tokens removed.");
}

// ============================================================================
// INTERACTIVE SETUP
// ============================================================================

const readline = require("readline");

function prompt(question) {
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });
  return new Promise((resolve) => {
    rl.question(question, (answer) => {
      rl.close();
      resolve(answer.trim());
    });
  });
}

function getClaudeConfigPath() {
  const home = process.env.HOME || process.env.USERPROFILE || "";
  return path.join(home, ".claude.json");
}

function getVSCodeConfigPath() {
  const home = process.env.HOME || process.env.USERPROFILE || "";
  // Check common VS Code config locations
  const paths = [
    path.join(home, ".vscode", "mcp.json"),
    path.join(home, "Library", "Application Support", "Code", "User", "settings.json"), // macOS
    path.join(home, ".config", "Code", "User", "settings.json"), // Linux
    path.join(home, "AppData", "Roaming", "Code", "User", "settings.json"), // Windows
  ];
  for (const p of paths) {
    if (fs.existsSync(p)) return p;
  }
  return path.join(home, ".vscode", "mcp.json"); // Default
}

function getMcpServerConfig(tenantId, clientId) {
  return {
    type: "stdio",
    command: "npx",
    args: ["-y", "github:rapyuta-robotics/agent365-mcp", "serve"],
    env: {
      AGENT365_TENANT_ID: tenantId,
      AGENT365_CLIENT_ID: clientId,
    },
  };
}

function configureClaudeCode(tenantId, clientId) {
  const configPath = getClaudeConfigPath();
  let config = {};

  if (fs.existsSync(configPath)) {
    try {
      config = JSON.parse(fs.readFileSync(configPath, "utf8"));
    } catch (e) {
      console.log("⚠️  Could not parse existing Claude config, creating new one.");
    }
  }

  if (!config.mcpServers) {
    config.mcpServers = {};
  }

  config.mcpServers.agent365 = getMcpServerConfig(tenantId, clientId);

  fs.writeFileSync(configPath, JSON.stringify(config, null, 2));
  console.log(`✅ Claude Code configured: ${configPath}`);
  return true;
}

function configureVSCode(tenantId, clientId) {
  const configPath = getVSCodeConfigPath();
  const dir = path.dirname(configPath);

  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }

  let config = {};
  if (fs.existsSync(configPath)) {
    try {
      config = JSON.parse(fs.readFileSync(configPath, "utf8"));
    } catch (e) {
      // Start fresh
    }
  }

  if (!config.mcpServers) {
    config.mcpServers = {};
  }

  config.mcpServers.agent365 = getMcpServerConfig(tenantId, clientId);

  fs.writeFileSync(configPath, JSON.stringify(config, null, 2));
  console.log(`✅ VS Code configured: ${configPath}`);
  return true;
}

async function interactiveSetup() {
  console.log(`
╔══════════════════════════════════════════════════════════════╗
║           Agent 365 MCP - Interactive Setup                   ║
╚══════════════════════════════════════════════════════════════╝
`);

  // Check for existing config
  const existingConfig = loadConfig();
  let tenantId = process.env.AGENT365_TENANT_ID || existingConfig.tenantId;
  let clientId = process.env.AGENT365_CLIENT_ID || existingConfig.clientId;

  // Get credentials
  if (tenantId && clientId) {
    console.log("Found existing configuration:");
    console.log(`  Tenant ID: ${tenantId}`);
    console.log(`  Client ID: ${clientId}\n`);
    const useExisting = await prompt("Use existing credentials? (Y/n): ");
    if (useExisting.toLowerCase() === "n") {
      tenantId = null;
      clientId = null;
    }
  }

  if (!tenantId) {
    console.log("\n📋 Get these values from your IT admin or Azure Portal:\n");
    tenantId = await prompt("Enter Tenant ID: ");
    if (!tenantId) {
      console.error("❌ Tenant ID is required.");
      process.exit(1);
    }
  }

  if (!clientId) {
    clientId = await prompt("Enter Client ID: ");
    if (!clientId) {
      console.error("❌ Client ID is required.");
      process.exit(1);
    }
  }

  // Authenticate
  console.log("\n🔐 Starting authentication...\n");
  await authenticate(tenantId, clientId);

  // Configure MCP clients
  console.log("\n📝 Configuring MCP clients...\n");

  const configureCC = await prompt("Configure Claude Code? (Y/n): ");
  if (configureCC.toLowerCase() !== "n") {
    configureClaudeCode(tenantId, clientId);
  }

  const configureVS = await prompt("Configure VS Code? (y/N): ");
  if (configureVS.toLowerCase() === "y") {
    configureVSCode(tenantId, clientId);
  }

  console.log(`
╔══════════════════════════════════════════════════════════════╗
║                    Setup Complete! 🎉                         ║
╚══════════════════════════════════════════════════════════════╝

Next steps:
  1. Restart your coding assistant (Claude Code, VS Code, etc.)
  2. The agent365 tools will be available automatically

To verify: npx github:rapyuta-robotics/agent365-mcp status

Your session is valid for ~90 days before re-authentication is needed.
`);
}

async function main() {
  const args = process.argv.slice(2);
  const command = args[0];

  switch (command) {
    case "auth":
    case "login":
      const tenantId = process.env.AGENT365_TENANT_ID || process.env.TENANT_ID || args[1];
      const clientId = process.env.AGENT365_CLIENT_ID || process.env.CLIENT_ID || args[2];

      if (!tenantId || !clientId) {
        console.error("❌ Error: Tenant ID and Client ID are required.\n");
        console.error("Set environment variables:");
        console.error("  export AGENT365_TENANT_ID=<your-tenant-id>");
        console.error("  export AGENT365_CLIENT_ID=<your-client-id>\n");
        console.error("Or provide as arguments:");
        console.error("  agent365-mcp auth <tenant-id> <client-id>\n");
        console.error("Get these from your IT admin or Azure portal.");
        process.exit(1);
      }

      await authenticate(tenantId, clientId);
      break;

    case "serve":
    case "start":
    case undefined:
      // When run without args or with 'serve', start the MCP server
      if (command === undefined && process.stdin.isTTY) {
        // Interactive mode - show help
        showHelp();
      } else {
        // Non-interactive (piped) or explicit serve - run MCP server
        require("./index.js");
      }
      break;

    case "status":
      showStatus();
      break;

    case "logout":
      logout();
      break;

    case "setup":
    case "install":
    case "configure":
      await interactiveSetup();
      break;

    case "help":
    case "--help":
    case "-h":
      showHelp();
      break;

    default:
      console.error(`Unknown command: ${command}`);
      showHelp();
      process.exit(1);
  }
}

main().catch((error) => {
  console.error("❌ Error:", error.message);
  process.exit(1);
});
