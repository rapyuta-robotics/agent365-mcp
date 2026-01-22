#!/usr/bin/env node
/**
 * Agent 365 MCP Proxy Server
 *
 * Aggregates Microsoft 365 MCP servers into a single endpoint for MCP clients.
 * Handles authentication, response size management, and graceful degradation.
 *
 * @see https://github.com/anthropics/agent365-mcp for documentation
 * @see ADMIN.md for IT administrator setup guide
 * @see ARCHITECTURE.md for technical details
 */
const { Server } = require("@modelcontextprotocol/sdk/server/index.js");
const { StdioServerTransport } = require("@modelcontextprotocol/sdk/server/stdio.js");
const {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} = require("@modelcontextprotocol/sdk/types.js");
const fs = require("fs");
const path = require("path");
const https = require("https");

// ============================================================================
// CONFIGURATION
// ============================================================================

const HOME = process.env.HOME || process.env.USERPROFILE || "";
const TOKEN_CACHE_PATH = process.env.AGENT365_TOKEN_PATH ||
  path.join(HOME, ".agent365-mcp", "tokens.json");
const LEGACY_TOKEN_PATH = path.join(HOME, ".claude", "agent365-tokens.json");

// Response handling configuration
const MAX_RESPONSE_SIZE = parseInt(process.env.AGENT365_MAX_RESPONSE_SIZE) || 50000;
const MAX_BUFFER_SIZE = 10 * 1024 * 1024; // 10MB max buffer
const REQUEST_TIMEOUT = parseInt(process.env.AGENT365_TIMEOUT) || 60000;

// Large file handling - if set, saves large responses to this directory
const LARGE_FILE_DIR = process.env.AGENT365_LARGE_FILE_DIR || "";
const LARGE_FILE_THRESHOLD = parseInt(process.env.AGENT365_LARGE_FILE_THRESHOLD) || 100000;

// Safety configuration
const ALLOW_DANGEROUS_TOOLS = process.env.AGENT365_ALLOW_DANGEROUS === "true";

// Servers to disable (comma-separated list, e.g., "copilot,excel")
const DISABLED_SERVERS = (process.env.AGENT365_DISABLED_SERVERS || "").split(",").filter(Boolean);

// ============================================================================
// MCP SERVER DEFINITIONS
// Each server requires Copilot for M365 license unless noted otherwise.
// Tool descriptions include hints for better client tool selection.
// ============================================================================

const MCP_SERVERS = {
  sharepoint: {
    url: "https://agent365.svc.cloud.microsoft/agents/servers/mcp_ODSPRemoteServer",
    prefix: "sharepoint",
    description: "SharePoint & OneDrive files",
    requiresCopilot: true,
    toolHints: {
      // Hints to help clients choose the right tool
      findFileOrFolder: "Use this to search for files by name. For Excel files (.xlsx), consider using excel_* tools after finding. For Word files (.docx), consider using word_* tools.",
      readSmallTextFile: "Reads raw file content. For structured documents, prefer word_WordGetDocumentContent or excel_ExcelGetDocumentContent instead.",
      getFileOrFolderMetadataByUrl: "Use this when you have a SharePoint URL and need file details before deciding which tool to use.",
    },
  },
  word: {
    url: "https://agent365.svc.cloud.microsoft/agents/servers/mcp_WordServer",
    prefix: "word",
    description: "Word documents - read content, comments, create docs",
    requiresCopilot: true,
    toolHints: {
      WordGetDocumentContent: "Best for reading Word documents (.docx). Extracts text, comments, and structure. Use this instead of sharepoint_readSmallTextFile for Word files.",
      WordCreateNewDocument: "Creates new Word documents. Provide content in markdown or plain text.",
      WordCreateNewComment: "Add comments to specific parts of a Word document.",
    },
  },
  teams: {
    url: "https://agent365.svc.cloud.microsoft/agents/servers/mcp_TeamsServer",
    prefix: "teams",
    description: "Teams chats, channels, messages, meetings",
    requiresCopilot: true,
    toolHints: {
      listChats: "Lists recent chats including meeting chats. Use expand=lastMessagePreview for recent activity.",
      listChatMessages: "Gets messages from a specific chat. For meeting discussions, find the meeting chat first.",
      postMessage: "Sends a message to a chat. Requires explicit user confirmation.",
    },
  },
  copilot: {
    url: "https://agent365.svc.cloud.microsoft/agents/servers/mcp_M365Copilot",
    prefix: "copilot",
    description: "M365 Copilot - semantic search across all M365 content, meeting transcripts, file summaries",
    requiresCopilot: true, // Strictly requires Copilot license
    toolHints: {
      CopilotChat: "Powerful semantic search across M365. Use for finding information when you don't know exact file names, summarizing meeting transcripts, or searching across emails/chats/files simultaneously.",
    },
  },
  mail: {
    url: "https://agent365.svc.cloud.microsoft/agents/servers/mcp_MailTools",
    prefix: "mail",
    description: "Outlook email - read, search, send, reply",
    requiresCopilot: true,
    toolHints: {
      SearchMessagesAsync: "Search emails by keyword, sender, date range. More efficient than listing all messages.",
      GetMessageAsync: "Get full email content including attachments list.",
      SendEmailWithAttachmentsAsync: "Send email. Requires explicit user confirmation before sending.",
    },
  },
  calendar: {
    url: "https://agent365.svc.cloud.microsoft/agents/servers/mcp_CalendarTools",
    prefix: "calendar",
    description: "Outlook calendar - events, meetings, scheduling",
    requiresCopilot: true,
    toolHints: {
      listCalendarView: "List events in a date range. Use specific date ranges to avoid large responses.",
      getEvent: "Get details of a specific event including attendees.",
      createEvent: "Create calendar event. Requires explicit user confirmation.",
    },
  },
  me: {
    url: "https://agent365.svc.cloud.microsoft/agents/servers/mcp_MeServer",
    prefix: "me",
    description: "User profile, org chart, direct reports, manager info",
    requiresCopilot: true,
    toolHints: {
      getMyProfile: "Get current user's profile information.",
      getDirectReports: "List people who report to the current user.",
      getMyManager: "Get the current user's manager information.",
    },
  },
  excel: {
    url: "https://agent365.svc.cloud.microsoft/agents/servers/mcp_ExcelServer",
    prefix: "excel",
    description: "Excel spreadsheets - read content, comments, formulas",
    requiresCopilot: true,
    toolHints: {
      ExcelGetDocumentContent: "Best for reading Excel files (.xlsx). Use this instead of sharepoint_readSmallTextFile for spreadsheets. Returns cell data as structured text.",
      ExcelCreateNewWorkbook: "Create new Excel workbook.",
    },
  },
};

// ============================================================================
// DANGEROUS TOOL FILTERING
// These patterns identify tools that can permanently modify or delete data.
// Set AGENT365_ALLOW_DANGEROUS=true to enable these tools.
// ============================================================================

const DANGEROUS_TOOL_PATTERNS = [
  /delete/i,
  /remove/i,
  /destroy/i,
  /purge/i,
  /erase/i,
  /drop/i,
];

function isDangerousTool(toolName) {
  if (ALLOW_DANGEROUS_TOOLS) return false;
  return DANGEROUS_TOOL_PATTERNS.some(pattern => pattern.test(toolName));
}

// ============================================================================
// LARGE CONTENT HANDLING
// Responses exceeding thresholds can be saved to filesystem for access.
// ============================================================================

/**
 * Save large content to filesystem and return a reference.
 * Falls back to truncation if filesystem saving is not configured.
 */
function handleLargeContent(text, toolName) {
  if (!text || typeof text !== "string") return text;

  const size = text.length;
  if (size <= LARGE_FILE_THRESHOLD) return text;

  // If large file directory is configured, save to file
  if (LARGE_FILE_DIR) {
    try {
      const dir = path.resolve(LARGE_FILE_DIR);
      if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
      }

      const timestamp = Date.now();
      const safeName = toolName.replace(/[^a-zA-Z0-9_-]/g, "_");
      const filename = `agent365-${safeName}-${timestamp}.txt`;
      const filepath = path.join(dir, filename);

      fs.writeFileSync(filepath, text, "utf8");

      return JSON.stringify({
        message: `Response saved to file (${(size / 1024).toFixed(1)}KB)`,
        filepath: filepath,
        size: size,
        hint: "Use the Read tool to access this file, or use more specific queries to reduce response size.",
      });
    } catch (err) {
      console.error(`Failed to save large response to file: ${err.message}`);
      // Fall through to truncation
    }
  }

  // Truncate with helpful message
  if (size > MAX_RESPONSE_SIZE) {
    const truncated = text.slice(0, MAX_RESPONSE_SIZE);
    const remaining = size - MAX_RESPONSE_SIZE;
    return `${truncated}\n\n... [Content truncated: ${remaining.toLocaleString()} more characters. Set AGENT365_LARGE_FILE_DIR to save large responses to disk, or use more specific queries.]`;
  }

  return text;
}

/**
 * Process MCP result content, handling large responses appropriately.
 */
function processResult(result, toolName = "unknown") {
  if (!result) return result;

  // Handle content array (standard MCP format)
  if (result.content && Array.isArray(result.content)) {
    result.content = result.content.map(item => {
      if (item.type === "text" && item.text) {
        item.text = handleLargeContent(item.text, toolName);
      }
      return item;
    });
  }

  // Handle raw text content
  if (typeof result === "string") {
    return handleLargeContent(result, toolName);
  }

  // Handle response field (Agent 365 format)
  if (result.response && typeof result.response === "string") {
    result.response = handleLargeContent(result.response, toolName);
  }

  // Handle message field
  if (result.message && typeof result.message === "string" && result.message.length > LARGE_FILE_THRESHOLD) {
    result.message = handleLargeContent(result.message, toolName);
  }

  return result;
}

// ============================================================================
// TOKEN MANAGEMENT (with auto-refresh support)
// ============================================================================

const { PublicClientApplication } = require("@azure/msal-node");

const MSAL_CACHE_PATH = path.join(HOME, ".agent365-mcp", "msal-cache.json");
const AGENT365_API = "ea9ffc3e-8a23-4a7d-836d-234d7c7565c1";

// In-memory token cache for the session
let cachedToken = null;
let cachedTokenExpiry = null;

function createMsalCachePlugin() {
  const beforeCacheAccess = async (cacheContext) => {
    try {
      if (fs.existsSync(MSAL_CACHE_PATH)) {
        cacheContext.tokenCache.deserialize(fs.readFileSync(MSAL_CACHE_PATH, "utf8"));
      }
    } catch (e) {
      // Ignore cache read errors
    }
  };

  const afterCacheAccess = async (cacheContext) => {
    if (cacheContext.cacheHasChanged) {
      const dir = path.dirname(MSAL_CACHE_PATH);
      if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
      }
      fs.writeFileSync(MSAL_CACHE_PATH, cacheContext.tokenCache.serialize());
    }
  };

  return { beforeCacheAccess, afterCacheAccess };
}

async function refreshTokenSilently() {
  // Load config to get tenant/client IDs
  const tokenData = loadTokenData();
  if (!tokenData || !tokenData.tenantId || !tokenData.clientId) {
    return null;
  }

  const pca = new PublicClientApplication({
    auth: {
      clientId: tokenData.clientId,
      authority: `https://login.microsoftonline.com/${tokenData.tenantId}`,
    },
    cache: {
      cachePlugin: createMsalCachePlugin(),
    },
  });

  const accounts = await pca.getTokenCache().getAllAccounts();
  if (accounts.length === 0) {
    return null;
  }

  try {
    const response = await pca.acquireTokenSilent({
      account: accounts[0],
      scopes: [`${AGENT365_API}/.default`],
    });

    // Update the stored token
    const tokenPath = TOKEN_CACHE_PATH;
    fs.writeFileSync(tokenPath, JSON.stringify({
      accessToken: response.accessToken,
      expiresOn: response.expiresOn,
      account: response.account,
      tenantId: tokenData.tenantId,
      clientId: tokenData.clientId,
    }));

    console.error(`Token refreshed silently, expires: ${response.expiresOn}`);
    return response.accessToken;
  } catch (e) {
    console.error(`Silent token refresh failed: ${e.message}`);
    return null;
  }
}

function loadTokenData() {
  const paths = [TOKEN_CACHE_PATH, LEGACY_TOKEN_PATH];
  for (const tokenPath of paths) {
    try {
      if (fs.existsSync(tokenPath)) {
        return JSON.parse(fs.readFileSync(tokenPath, "utf8"));
      }
    } catch (e) {
      // Try next path
    }
  }
  return null;
}

async function loadToken() {
  // Check in-memory cache first
  if (cachedToken && cachedTokenExpiry && new Date(cachedTokenExpiry) > new Date()) {
    return cachedToken;
  }

  const paths = [TOKEN_CACHE_PATH, LEGACY_TOKEN_PATH];

  for (const tokenPath of paths) {
    try {
      if (fs.existsSync(tokenPath)) {
        const data = JSON.parse(fs.readFileSync(tokenPath, "utf8"));
        if (data.expiresOn && new Date(data.expiresOn) > new Date()) {
          cachedToken = data.accessToken;
          cachedTokenExpiry = data.expiresOn;
          return data.accessToken;
        }
      }
    } catch (e) {
      // Try next path
    }
  }

  // Token expired - try to refresh silently
  console.error("Access token expired, attempting silent refresh...");
  const refreshedToken = await refreshTokenSilently();
  if (refreshedToken) {
    cachedToken = refreshedToken;
    cachedTokenExpiry = new Date(Date.now() + 3600 * 1000).toISOString(); // ~1 hour
    return refreshedToken;
  }

  console.error("No valid token found. Run: npx github:rapyuta-robotics/agent365-mcp auth");
  return null;
}

// ============================================================================
// AGENT 365 API COMMUNICATION
// ============================================================================

async function makeAgent365Request(serverUrl, method, params, id) {
  const token = await loadToken();
  if (!token) {
    throw new Error("No valid token. Run: npx github:rapyuta-robotics/agent365-mcp auth");
  }

  return new Promise((resolve, reject) => {

    const url = new URL(serverUrl);
    const body = JSON.stringify({ jsonrpc: "2.0", method, params, id });

    const options = {
      hostname: url.hostname,
      path: url.pathname,
      method: "POST",
      headers: {
        "Authorization": `Bearer ${token}`,
        "Content-Type": "application/json",
        "Accept": "application/json, text/event-stream",
        "Content-Length": Buffer.byteLength(body),
      },
    };

    const req = https.request(options, (res) => {
      let data = "";
      let bufferExceeded = false;

      res.on("data", (chunk) => {
        if (data.length + chunk.length > MAX_BUFFER_SIZE) {
          if (!bufferExceeded) {
            bufferExceeded = true;
            console.error(`Response buffer exceeded ${MAX_BUFFER_SIZE} bytes for ${serverUrl}`);
          }
          return;
        }
        data += chunk;
      });

      res.on("end", () => {
        // Parse SSE format
        const lines = data.split("\n");
        for (const line of lines) {
          if (line.startsWith("data: ")) {
            try {
              const jsonData = JSON.parse(line.slice(6));
              resolve(jsonData);
              return;
            } catch (e) {
              // Continue
            }
          }
        }
        // Try plain JSON
        try {
          resolve(JSON.parse(data));
        } catch (e) {
          if (bufferExceeded) {
            reject(new Error(`Response too large (>${MAX_BUFFER_SIZE / 1024 / 1024}MB). Try a more specific query.`));
          } else {
            reject(new Error(`Invalid response from ${serverUrl}: ${data.slice(0, 200)}`));
          }
        }
      });

      res.on("error", (err) => {
        reject(new Error(`Response error from ${serverUrl}: ${err.message}`));
      });
    });

    req.on("error", reject);
    req.setTimeout(REQUEST_TIMEOUT, () => {
      req.destroy();
      reject(new Error(`Request timeout (${REQUEST_TIMEOUT / 1000}s) for ${serverUrl}. Try a more specific query.`));
    });
    req.write(body);
    req.end();
  });
}

// ============================================================================
// SERVER SESSION MANAGEMENT
// Tracks initialization state and handles license/permission errors gracefully.
// ============================================================================

const serverSessions = {};

// Track servers that failed due to license issues
const disabledServers = new Set(DISABLED_SERVERS);

async function initializeServer(serverKey) {
  if (serverSessions[serverKey]?.initialized) return true;
  if (disabledServers.has(serverKey)) return false;

  const server = MCP_SERVERS[serverKey];
  if (!server) throw new Error(`Unknown server: ${serverKey}`);

  try {
    const result = await makeAgent365Request(server.url, "initialize", {
      protocolVersion: "2024-11-05",
      capabilities: {},
      clientInfo: { name: "agent365-proxy", version: "1.1.0" },
    }, 1);

    if (result.error) {
      // Check for license-related errors
      const errorMsg = result.error.message || "";
      if (errorMsg.includes("license") || errorMsg.includes("not entitled") ||
          errorMsg.includes("Copilot") || errorMsg.includes("subscription")) {
        console.error(`${serverKey}: License not available - disabling server. Error: ${errorMsg}`);
        disabledServers.add(serverKey);
        return false;
      }
      throw new Error(result.error.message);
    }

    serverSessions[serverKey] = { initialized: true, tools: null };
    console.error(`Initialized ${serverKey} server`);
    return true;
  } catch (error) {
    // Handle network/permission errors gracefully
    const errorMsg = error.message || "";
    if (errorMsg.includes("403") || errorMsg.includes("401") ||
        errorMsg.includes("license") || errorMsg.includes("permission")) {
      console.error(`${serverKey}: Access denied - disabling server. Error: ${errorMsg}`);
      disabledServers.add(serverKey);
      return false;
    }
    console.error(`Failed to initialize ${serverKey}: ${error.message}`);
    throw error;
  }
}

async function getServerTools(serverKey) {
  const server = MCP_SERVERS[serverKey];
  if (!server) return [];
  if (disabledServers.has(serverKey)) return [];

  if (serverSessions[serverKey]?.tools) {
    return serverSessions[serverKey].tools;
  }

  try {
    const initialized = await initializeServer(serverKey);
    if (!initialized) return [];

    const result = await makeAgent365Request(server.url, "tools/list", {}, 2);

    if (result.error) {
      console.error(`Error getting tools for ${serverKey}: ${result.error.message}`);
      return [];
    }

    // Process tools: prefix names, filter dangerous, enhance descriptions
    const tools = (result.result?.tools || [])
      .filter(tool => !isDangerousTool(tool.name))
      .map(tool => {
        const hint = server.toolHints?.[tool.name] || "";
        const enhancedDescription = hint
          ? `[${server.description}] ${tool.description || ""}\n\nHint: ${hint}`
          : `[${server.description}] ${tool.description || ""}`;

        return {
          ...tool,
          name: `${server.prefix}_${tool.name}`,
          description: enhancedDescription,
          _serverKey: serverKey,
          _originalName: tool.name,
        };
      });

    serverSessions[serverKey].tools = tools;
    return tools;
  } catch (error) {
    console.error(`Error connecting to ${serverKey}: ${error.message}`);
    return [];
  }
}

async function callServerTool(serverKey, toolName, args) {
  const server = MCP_SERVERS[serverKey];
  if (!server) throw new Error(`Unknown server: ${serverKey}`);

  if (disabledServers.has(serverKey)) {
    throw new Error(`${serverKey} server is disabled (license not available or access denied)`);
  }

  await initializeServer(serverKey);

  const result = await makeAgent365Request(
    server.url,
    "tools/call",
    { name: toolName, arguments: args },
    Date.now()
  );

  if (result.error) {
    throw new Error(result.error.message);
  }

  // Process with tool name for better file naming
  return processResult(result.result, `${serverKey}_${toolName}`);
}

// ============================================================================
// MCP SERVER SETUP
// ============================================================================

const server = new Server(
  {
    name: "agent365-mcp-proxy",
    version: "1.1.0",
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

const toolServerMap = {};

// List tools handler
server.setRequestHandler(ListToolsRequestSchema, async () => {
  const allTools = [];
  const serverKeys = Object.keys(MCP_SERVERS).filter(k => !disabledServers.has(k));

  const results = await Promise.allSettled(
    serverKeys.map(key => getServerTools(key))
  );

  for (let i = 0; i < results.length; i++) {
    if (results[i].status === "fulfilled") {
      const tools = results[i].value;
      for (const tool of tools) {
        toolServerMap[tool.name] = {
          serverKey: serverKeys[i],
          originalName: tool._originalName,
        };
        const { _serverKey, _originalName, ...cleanTool } = tool;
        allTools.push(cleanTool);
      }
    }
  }

  // Always include auth tool so users can authenticate/re-authenticate
  allTools.unshift({
    name: "agent365_authenticate",
    description: "Start Microsoft 365 authentication. Call this if other agent365 tools fail with auth errors, or to check authentication status. Returns a device code and URL - tell the user to visit the URL and enter the code.",
    inputSchema: { type: "object", properties: {} },
  });

  if (allTools.length === 1) { // Only auth tool, no real tools loaded
    const disabledList = Array.from(disabledServers).join(", ");
    allTools.push({
      name: "agent365_status",
      description: `Authentication required. Call agent365_authenticate first to get login instructions.\n\nDisabled servers: ${disabledList || "none"}`,
      inputSchema: { type: "object", properties: {} },
    });
  }

  const activeServers = serverKeys.filter(k => !disabledServers.has(k)).length;
  const disabledCount = disabledServers.size;
  console.error(`Loaded ${allTools.length} tools from ${activeServers} servers (${disabledCount} disabled)`);

  return { tools: allTools };
});

// Call tool handler
server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  // Handle built-in auth tool
  if (name === "agent365_authenticate") {
    return await handleAuthenticate();
  }

  if (name === "agent365_status") {
    const token = await loadToken();
    if (token) {
      return {
        content: [{
          type: "text",
          text: "✅ Authenticated and ready. Other agent365_* tools should now be available.",
        }],
      };
    } else {
      return {
        content: [{
          type: "text",
          text: "❌ Not authenticated. Please call agent365_authenticate to get login instructions.",
        }],
      };
    }
  }

  // Check for dangerous operations in arguments (unless allowed)
  if (!ALLOW_DANGEROUS_TOOLS) {
    const argsStr = JSON.stringify(args || {}).toLowerCase();
    if (argsStr.includes("delete") || argsStr.includes("remove")) {
      return {
        content: [{
          type: "text",
          text: "This operation appears to involve deletion which is blocked for safety. Set AGENT365_ALLOW_DANGEROUS=true to enable, or perform deletions manually.",
        }],
        isError: true,
      };
    }
  }

  const mapping = toolServerMap[name];
  if (!mapping) {
    return {
      content: [{
        type: "text",
        text: `Unknown tool: ${name}. Available tools may need to be refreshed.`,
      }],
      isError: true,
    };
  }

  try {
    const result = await callServerTool(mapping.serverKey, mapping.originalName, args);
    return result;
  } catch (error) {
    return {
      content: [{
        type: "text",
        text: `Error calling ${name}: ${error.message}\n\nIf authentication expired, run: npx github:rapyuta-robotics/agent365-mcp auth`,
      }],
      isError: true,
    };
  }
});

// ============================================================================
// IN-CHAT AUTHENTICATION (returns device code in tool result)
// ============================================================================

let pendingAuthPromise = null;

async function handleAuthenticate() {
  const config = loadServerConfig();

  if (!config) {
    return {
      content: [{
        type: "text",
        text: `❌ **Configuration Missing**

No Tenant ID or Client ID found. These must be set in your MCP config:

\`\`\`json
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
\`\`\`

Get these values from your IT administrator, then restart your coding assistant.`,
      }],
      isError: true,
    };
  }

  // Check if already authenticated
  const existingToken = await loadToken();
  if (existingToken) {
    const tokenData = loadTokenData();
    return {
      content: [{
        type: "text",
        text: `✅ **Already Authenticated**

Account: ${tokenData?.account?.username || "Unknown"}
Status: Valid token present

You can use the other agent365_* tools now. If you're having issues, try calling this tool again after the token expires.`,
      }],
    };
  }

  // Start device code flow
  const pca = new PublicClientApplication({
    auth: {
      clientId: config.clientId,
      authority: `https://login.microsoftonline.com/${config.tenantId}`,
    },
    cache: {
      cachePlugin: createMsalCachePlugin(),
    },
  });

  const scopes = [`${AGENT365_API}/.default`];

  // Try silent refresh first
  const accounts = await pca.getTokenCache().getAllAccounts();
  if (accounts.length > 0) {
    try {
      const response = await pca.acquireTokenSilent({
        account: accounts[0],
        scopes: scopes,
      });
      saveToken(response, config);
      return {
        content: [{
          type: "text",
          text: `✅ **Token Refreshed**

Account: ${response.account?.username || "Unknown"}
Expires: ${response.expiresOn}

Authentication successful! You can now use the other agent365_* tools.`,
        }],
      };
    } catch (e) {
      // Fall through to device code
    }
  }

  // Device code flow - return the code to the user via tool result
  return new Promise((resolve) => {
    let deviceCode = null;
    let verificationUrl = null;

    const authPromise = pca.acquireTokenByDeviceCode({
      scopes: scopes,
      deviceCodeCallback: (response) => {
        deviceCode = response.userCode;
        verificationUrl = response.verificationUriComplete || response.verificationUri;

        // Save code to file as backup
        saveDeviceCodeToFile(deviceCode, response.verificationUri);

        // Copy to clipboard
        copyToClipboard(deviceCode);

        // Return immediately with the code - don't wait for auth to complete
        resolve({
          content: [{
            type: "text",
            text: `🔐 **Microsoft 365 Authentication Required**

**Your code: \`${deviceCode}\`** (copied to clipboard)

**Steps for user:**
1. Go to: ${verificationUrl}
2. Enter the code: **${deviceCode}**
3. Sign in with your Microsoft account

**For agent:** After the user confirms they've completed login, call \`agent365_authenticate\` again to verify authentication succeeded. The tool will return "Already Authenticated" or "Token Refreshed" when successful.

The code expires in 15 minutes.`,
          }],
        });
      },
    });

    // Handle auth completion in background
    authPromise.then((response) => {
      saveToken(response, config);
      console.error(`✅ Authentication completed for ${response.account?.username}`);
    }).catch((e) => {
      console.error(`❌ Authentication failed: ${e.message}`);
    });
  });
}

// ============================================================================
// AUTO-AUTHENTICATION ON FIRST USE
// ============================================================================

function getConfigPath() {
  return path.join(HOME, ".agent365-mcp", "config.json");
}

function loadServerConfig() {
  // Try environment variables first
  const tenantId = process.env.AGENT365_TENANT_ID || process.env.TENANT_ID;
  const clientId = process.env.AGENT365_CLIENT_ID || process.env.CLIENT_ID;

  if (tenantId && clientId) {
    return { tenantId, clientId };
  }

  // Try config file
  const configPath = getConfigPath();
  if (fs.existsSync(configPath)) {
    try {
      const config = JSON.parse(fs.readFileSync(configPath, "utf8"));
      if (config.tenantId && config.clientId) {
        return config;
      }
    } catch (e) {
      // Ignore
    }
  }

  // Try token file (might have config embedded)
  const tokenData = loadTokenData();
  if (tokenData?.tenantId && tokenData?.clientId) {
    return { tenantId: tokenData.tenantId, clientId: tokenData.clientId };
  }

  return null;
}

function openBrowser(url) {
  const { exec } = require("child_process");
  const platform = process.platform;

  let cmd;
  if (platform === "darwin") {
    cmd = `open "${url}"`;
  } else if (platform === "win32") {
    cmd = `start "" "${url}"`;
  } else {
    cmd = `xdg-open "${url}"`;
  }

  exec(cmd, (err) => {
    if (err) {
      console.error(`Could not open browser automatically. Please visit: ${url}`);
    }
  });
}

function copyToClipboard(text) {
  const { exec } = require("child_process");
  const platform = process.platform;

  let cmd;
  if (platform === "darwin") {
    cmd = `echo "${text}" | pbcopy`;
  } else if (platform === "win32") {
    cmd = `echo ${text} | clip`;
  } else {
    cmd = `echo "${text}" | xclip -selection clipboard 2>/dev/null || echo "${text}" | xsel --clipboard 2>/dev/null`;
  }

  exec(cmd, () => {}); // Ignore errors - clipboard is nice-to-have
}

function saveDeviceCodeToFile(code, url) {
  const codePath = path.join(HOME, ".agent365-mcp", "device-code.txt");
  const content = `
═══════════════════════════════════════════════════════════════
  MICROSOFT 365 AUTHENTICATION - DEVICE CODE
═══════════════════════════════════════════════════════════════

  Your code: ${code}

  1. Go to: ${url}
  2. Enter the code above
  3. Sign in with your Microsoft account

  This file: ${codePath}
═══════════════════════════════════════════════════════════════
`;
  try {
    fs.writeFileSync(codePath, content);
  } catch (e) {
    // Ignore
  }
  return codePath;
}

async function autoAuthenticate(config) {
  console.error("\n🔐 First-time authentication required...\n");

  const pca = new PublicClientApplication({
    auth: {
      clientId: config.clientId,
      authority: `https://login.microsoftonline.com/${config.tenantId}`,
    },
    cache: {
      cachePlugin: createMsalCachePlugin(),
    },
  });

  const scopes = [`${AGENT365_API}/.default`];

  // Try silent auth first (if we have cached refresh token)
  const accounts = await pca.getTokenCache().getAllAccounts();
  if (accounts.length > 0) {
    try {
      console.error(`Found existing session for ${accounts[0].username}, refreshing...`);
      const response = await pca.acquireTokenSilent({
        account: accounts[0],
        scopes: scopes,
      });

      // Save the new token
      saveToken(response, config);
      console.error("✅ Token refreshed successfully!\n");
      return true;
    } catch (e) {
      console.error("Silent refresh failed, starting interactive login...\n");
    }
  }

  // Device code flow with auto-browser open
  try {
    const response = await pca.acquireTokenByDeviceCode({
      scopes: scopes,
      deviceCodeCallback: (deviceCodeResponse) => {
        const code = deviceCodeResponse.userCode;
        const url = deviceCodeResponse.verificationUri;
        // Use complete URL if available (has code pre-filled)
        const completeUrl = deviceCodeResponse.verificationUriComplete || url;

        // Save code to file for easy access
        const codePath = saveDeviceCodeToFile(code, url);

        // Copy code to clipboard
        copyToClipboard(code);

        console.error("\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
        console.error("📱 AUTHENTICATION REQUIRED");
        console.error("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
        console.error(`\n  Your code: ${code}  (copied to clipboard)\n`);
        console.error(`  Code also saved to: ${codePath}\n`);
        console.error("  Opening browser automatically...");
        console.error("  If prompted, enter the code above.\n");
        console.error("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n");

        // Auto-open browser with complete URL (includes code if supported)
        openBrowser(completeUrl);
      },
    });

    // Save token with config embedded
    saveToken(response, config);

    console.error("\n✅ Authentication successful!");
    console.error(`📧 Logged in as: ${response.account?.username || "Unknown"}`);
    console.error(`🔄 Session valid for ~90 days\n`);
    return true;
  } catch (e) {
    console.error(`❌ Authentication failed: ${e.message}`);
    return false;
  }
}

function saveToken(response, config) {
  const tokenPath = TOKEN_CACHE_PATH;
  const dir = path.dirname(tokenPath);

  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }

  fs.writeFileSync(tokenPath, JSON.stringify({
    accessToken: response.accessToken,
    expiresOn: response.expiresOn,
    account: response.account,
    tenantId: config.tenantId,
    clientId: config.clientId,
  }));

  // Also save config separately
  const configPath = getConfigPath();
  fs.writeFileSync(configPath, JSON.stringify({
    tenantId: config.tenantId,
    clientId: config.clientId,
  }, null, 2));

  // Update in-memory cache
  cachedToken = response.accessToken;
  cachedTokenExpiry = response.expiresOn;
}

async function ensureAuthenticated() {
  // First check if we have valid token
  const existingToken = await loadToken();
  if (existingToken) {
    return true;
  }

  // Need to authenticate - do we have config?
  const config = loadServerConfig();
  if (!config) {
    console.error(`
╔══════════════════════════════════════════════════════════════════════╗
║                 CONFIGURATION REQUIRED                                ║
╚══════════════════════════════════════════════════════════════════════╝

No Tenant ID or Client ID found. Get these from your IT administrator.

Option 1: Set environment variables in your MCP config:
  "env": {
    "AGENT365_TENANT_ID": "your-tenant-id",
    "AGENT365_CLIENT_ID": "your-client-id"
  }

Option 2: Run setup first:
  npx github:rapyuta-robotics/agent365-mcp setup

`);
    return false;
  }

  // We have config - try to authenticate automatically
  return await autoAuthenticate(config);
}

// ============================================================================
// STARTUP
// ============================================================================

async function main() {
  // Log configuration on startup
  console.error("Agent 365 MCP Proxy v1.3.1");
  console.error(`Configuration:`);
  console.error(`  MAX_RESPONSE_SIZE: ${MAX_RESPONSE_SIZE}`);
  console.error(`  REQUEST_TIMEOUT: ${REQUEST_TIMEOUT}ms`);
  console.error(`  ALLOW_DANGEROUS_TOOLS: ${ALLOW_DANGEROUS_TOOLS}`);
  console.error(`  LARGE_FILE_DIR: ${LARGE_FILE_DIR || "(not set - will truncate)"}`);
  if (DISABLED_SERVERS.length > 0) {
    console.error(`  DISABLED_SERVERS: ${DISABLED_SERVERS.join(", ")}`);
  }

  // Check auth status (don't block - let user authenticate via tool)
  const token = await loadToken();
  if (token) {
    console.error("✅ Authenticated - ready to connect to Microsoft 365");
  } else {
    console.error("⚠️  Not authenticated - use agent365_authenticate tool to login");
  }

  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("MCP server ready.");
}

main().catch(console.error);
