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
      readSmallTextFile: "For structured documents, prefer word_GetDocumentContent (.docx) or excel_GetDocumentContent (.xlsx).",
      createSmallTextFile: "For local files from disk, use sharepoint_uploadLocalFile instead.",
      createSmallBinaryFile: "For local files from disk, use sharepoint_uploadLocalFile instead.",
    },
  },
  word: {
    url: "https://agent365.svc.cloud.microsoft/agents/servers/mcp_WordServer",
    prefix: "word",
    description: "Word documents - read content, comments, create docs",
    requiresCopilot: true,
    toolHints: {},
  },
  teams: {
    url: "https://agent365.svc.cloud.microsoft/agents/servers/mcp_TeamsServer",
    prefix: "teams",
    description: "Teams chats, channels, messages, meetings",
    requiresCopilot: true,
    toolHints: {},
  },
  copilot: {
    url: "https://agent365.svc.cloud.microsoft/agents/servers/mcp_M365Copilot",
    prefix: "copilot",
    description: "M365 Copilot - semantic search across all M365 content",
    requiresCopilot: true, // Strictly requires Copilot license
    toolHints: {},
  },
  mail: {
    url: "https://agent365.svc.cloud.microsoft/agents/servers/mcp_MailTools",
    prefix: "mail",
    description: "Outlook email - read, search, send, reply, attachments",
    requiresCopilot: true,
    toolHints: {},
  },
  calendar: {
    url: "https://agent365.svc.cloud.microsoft/agents/servers/mcp_CalendarTools",
    prefix: "calendar",
    description: "Outlook calendar - events, meetings, scheduling",
    requiresCopilot: true,
    toolHints: {
      ListEvents: "WARNING: Only returns master events — may miss recurring meeting instances. For recurring meetings use ListCalendarView instead.",
      ListCalendarView: "PREFERRED for finding meetings — expands recurring events into individual instances.",
    },
  },
  me: {
    url: "https://agent365.svc.cloud.microsoft/agents/servers/mcp_MeServer",
    prefix: "me",
    description: "User profile, org chart, directory lookup",
    requiresCopilot: true,
    toolHints: {},
  },
  excel: {
    url: "https://agent365.svc.cloud.microsoft/agents/servers/mcp_ExcelServer",
    prefix: "excel",
    description: "Excel spreadsheets - read data, comments, create workbooks",
    requiresCopilot: true,
    toolHints: {},
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
// GRAPH API TOKEN MANAGEMENT
// Separate token for direct Microsoft Graph API calls (large file upload,
// meeting transcripts). Uses same MSAL PCA pattern as Agent365 token.
// ============================================================================

const GRAPH_SCOPES = ["Files.ReadWrite.All", "Sites.ReadWrite.All", "OnlineMeetings.Read", "OnlineMeetingTranscript.Read.All"];
const GRAPH_TOKEN_PATH = path.join(HOME, ".agent365-mcp", "graph-tokens.json");
let cachedGraphToken = null;
let cachedGraphTokenExpiry = null;

function loadGraphTokenData() {
  try {
    if (fs.existsSync(GRAPH_TOKEN_PATH)) {
      return JSON.parse(fs.readFileSync(GRAPH_TOKEN_PATH, "utf8"));
    }
  } catch (e) {
    // Ignore
  }
  return null;
}

async function refreshGraphTokenSilently() {
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
      scopes: GRAPH_SCOPES,
    });

    const dir = path.dirname(GRAPH_TOKEN_PATH);
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
    }

    fs.writeFileSync(GRAPH_TOKEN_PATH, JSON.stringify({
      accessToken: response.accessToken,
      expiresOn: response.expiresOn,
      account: response.account,
      tenantId: tokenData.tenantId,
      clientId: tokenData.clientId,
    }));

    console.error(`Graph token refreshed silently, expires: ${response.expiresOn}`);
    return response.accessToken;
  } catch (e) {
    console.error(`Silent Graph token refresh failed: ${e.message}`);
    return null;
  }
}

async function loadGraphToken() {
  // Check in-memory cache first
  if (cachedGraphToken && cachedGraphTokenExpiry && new Date(cachedGraphTokenExpiry) > new Date()) {
    return cachedGraphToken;
  }

  // Check token file
  try {
    if (fs.existsSync(GRAPH_TOKEN_PATH)) {
      const data = JSON.parse(fs.readFileSync(GRAPH_TOKEN_PATH, "utf8"));
      if (data.expiresOn && new Date(data.expiresOn) > new Date()) {
        cachedGraphToken = data.accessToken;
        cachedGraphTokenExpiry = data.expiresOn;
        return data.accessToken;
      }
    }
  } catch (e) {
    // Ignore
  }

  // Token expired or missing - try silent refresh
  console.error("Graph token expired or missing, attempting silent refresh...");
  const refreshedToken = await refreshGraphTokenSilently();
  if (refreshedToken) {
    cachedGraphToken = refreshedToken;
    cachedGraphTokenExpiry = new Date(Date.now() + 3600 * 1000).toISOString();
    return refreshedToken;
  }

  return null;
}

async function acquireGraphToken() {
  const config = loadServerConfig();
  if (!config) {
    return null;
  }

  const pca = new PublicClientApplication({
    auth: {
      clientId: config.clientId,
      authority: `https://login.microsoftonline.com/${config.tenantId}`,
    },
    cache: {
      cachePlugin: createMsalCachePlugin(),
    },
  });

  // Try silent first
  const accounts = await pca.getTokenCache().getAllAccounts();
  if (accounts.length > 0) {
    try {
      const response = await pca.acquireTokenSilent({
        account: accounts[0],
        scopes: GRAPH_SCOPES,
      });
      saveGraphToken(response, config);
      return response.accessToken;
    } catch (e) {
      // Fall through to device code
    }
  }

  // Device code flow
  return new Promise((resolve) => {
    pca.acquireTokenByDeviceCode({
      scopes: GRAPH_SCOPES,
      deviceCodeCallback: (response) => {
        const deviceCode = response.userCode;
        const verificationUrl = response.verificationUriComplete || response.verificationUri;

        saveDeviceCodeToFile(deviceCode, response.verificationUri);
        copyToClipboard(deviceCode);

        // Store for the caller to use
        resolve({
          deviceCode,
          verificationUrl,
          authPromise: null, // Will be set below
        });
      },
    }).then((response) => {
      saveGraphToken(response, loadServerConfig());
      console.error(`Graph API authentication completed for ${response.account?.username}`);
    }).catch((e) => {
      console.error(`Graph API authentication failed: ${e.message}`);
    });
  });
}

function saveGraphToken(response, config) {
  const dir = path.dirname(GRAPH_TOKEN_PATH);
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }

  fs.writeFileSync(GRAPH_TOKEN_PATH, JSON.stringify({
    accessToken: response.accessToken,
    expiresOn: response.expiresOn,
    account: response.account,
    tenantId: config.tenantId,
    clientId: config.clientId,
  }));

  cachedGraphToken = response.accessToken;
  cachedGraphTokenExpiry = response.expiresOn;
}

/**
 * Make an HTTPS request to Microsoft Graph API.
 * Returns parsed JSON response, or an error object { error: string } on failure.
 */
async function makeGraphRequest(method, urlPath, body, extraHeaders) {
  const token = await loadGraphToken();
  if (!token) {
    return { error: "No Graph API token available. Call the agent365_graph_auth tool to authenticate with Microsoft Graph API." };
  }

  return new Promise((resolve, reject) => {
    // encodeURI preserves URL structure ($, &, =, ?, /, :, ') but encodes spaces and other unsafe chars
    const fullPath = encodeURI(`/v1.0${urlPath}`);
    const bodyStr = body ? JSON.stringify(body) : null;

    const options = {
      hostname: "graph.microsoft.com",
      path: fullPath,
      method: method,
      headers: {
        "Authorization": `Bearer ${token}`,
        "Content-Type": "application/json",
        ...(extraHeaders || {}),
      },
    };

    if (bodyStr) {
      options.headers["Content-Length"] = Buffer.byteLength(bodyStr);
    }

    const req = https.request(options, (res) => {
      let data = "";
      res.on("data", (chunk) => {
        data += chunk;
      });
      res.on("end", () => {
        try {
          if (res.statusCode >= 200 && res.statusCode < 300) {
            const contentType = res.headers["content-type"] || "";
            if (contentType.includes("application/json") || (!contentType && data.startsWith("{"))) {
              resolve(data ? JSON.parse(data) : {});
            } else {
              // Non-JSON response (VTT transcripts, plain text, etc.)
              resolve(data);
            }
          } else {
            resolve({ error: `Graph API error (${res.statusCode}): ${data.slice(0, 500)}` });
          }
        } catch (e) {
          // JSON parse failed — return raw text
          resolve(data);
        }
      });
      res.on("error", (err) => {
        resolve({ error: `Graph API response error: ${err.message}` });
      });
    });

    req.on("error", (err) => {
      resolve({ error: `Graph API request error: ${err.message}` });
    });

    req.setTimeout(REQUEST_TIMEOUT, () => {
      req.destroy();
      resolve({ error: `Graph API request timeout (${REQUEST_TIMEOUT / 1000}s)` });
    });

    if (bodyStr) {
      req.write(bodyStr);
    }
    req.end();
  });
}

/**
 * Make a raw HTTPS request (for chunked uploads where we need to send binary data).
 * Returns { statusCode, data } or { error }.
 */
async function makeGraphRawRequest(method, url, bodyBuffer, headers) {
  const token = await loadGraphToken();
  if (!token) {
    return { error: "No Graph API token available." };
  }

  return new Promise((resolve) => {
    const parsedUrl = new URL(url);
    const options = {
      hostname: parsedUrl.hostname,
      path: parsedUrl.pathname + parsedUrl.search,
      method: method,
      headers: {
        "Authorization": `Bearer ${token}`,
        ...headers,
      },
    };

    if (bodyBuffer) {
      options.headers["Content-Length"] = bodyBuffer.length;
    }

    const req = https.request(options, (res) => {
      let data = "";
      res.on("data", (chunk) => {
        data += chunk;
      });
      res.on("end", () => {
        try {
          resolve({
            statusCode: res.statusCode,
            data: data ? JSON.parse(data) : {},
          });
        } catch (e) {
          resolve({ statusCode: res.statusCode, data: data });
        }
      });
      res.on("error", (err) => {
        resolve({ error: `Upload response error: ${err.message}` });
      });
    });

    req.on("error", (err) => {
      resolve({ error: `Upload request error: ${err.message}` });
    });

    req.setTimeout(120000, () => { // 2 min timeout for chunk uploads
      req.destroy();
      resolve({ error: "Upload chunk request timeout" });
    });

    if (bodyBuffer) {
      req.write(bodyBuffer);
    }
    req.end();
  });
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
      clientInfo: { name: "agent365-proxy", version: "1.4.0" },
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
    version: "1.4.0",
  },
  {
    capabilities: {
      tools: { listChanged: true },
    },
  }
);

const toolServerMap = {};
let toolsLoadedPromise = null;
let toolsLoaded = false;

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
    description: "Start Microsoft 365 authentication. Call this if tools return auth errors. Returns a device code and URL for the user to authenticate in their browser. Call agent365_help for workflow guidance.",
    inputSchema: { type: "object", properties: {} },
  });

  // Built-in tool: upload local files to SharePoint/OneDrive
  allTools.push({
    name: "sharepoint_uploadLocalFile",
    description: "[SharePoint & OneDrive files] Upload a local file from your machine to SharePoint or OneDrive. Supports any file type. Files up to 4MB use direct upload; 4-250MB use Graph API chunked upload (requires agent365_graph_auth). Call agent365_help for workflow guidance.",
    inputSchema: {
      type: "object",
      properties: {
        localFilePath: {
          type: "string",
          description: "Absolute path to the local file to upload (e.g. /home/user/report.docx)",
        },
        documentLibraryId: {
          type: "string",
          description: "Document library (drive) ID. Use 'me' for user's OneDrive. Find via sharepoint_listDocumentLibrariesInSite.",
        },
        parentFolderId: {
          type: "string",
          description: "Target folder ID within the document library. Defaults to root. Find via sharepoint_getFolderChildren.",
          default: "root",
        },
        filename: {
          type: "string",
          description: "Override filename for the upload. If omitted, uses the original filename from localFilePath.",
        },
      },
      required: ["localFilePath", "documentLibraryId"],
    },
  });

  // Built-in tool: get meeting transcript via Graph API
  allTools.push({
    name: "teams_getMeetingTranscript",
    description: "[Teams Meetings] Get the VTT transcript of a Teams meeting. Requires meetingUrl (get from calendar_ListCalendarView). Requires Graph API auth (agent365_graph_auth). Call agent365_help for workflow guidance.",
    inputSchema: {
      type: "object",
      properties: {
        meetingUrl: {
          type: "string",
          description: "The Teams meeting join URL. Get this from calendar_ListCalendarView (look for onlineMeeting.joinUrl in the event response).",
        },
      },
      required: ["meetingUrl"],
    },
  });

  // Built-in tool: authenticate with Microsoft Graph API for advanced features
  allTools.push({
    name: "agent365_graph_auth",
    description: "Authenticate with Microsoft Graph API for advanced features (large file upload >4MB, meeting transcripts). Only needed if those tools report auth errors. Uses device code flow.",
    inputSchema: { type: "object", properties: {} },
  });

  // Built-in tool: workflow guide (returns full instructions on demand to save context)
  allTools.push({
    name: "agent365_help",
    description: "Get workflow guidance for Microsoft 365 tools. Call this FIRST when working with SharePoint, Teams, Mail, Calendar, Excel, or Word tools — returns step-by-step workflows, tool chaining patterns, and best practices. Low cost: only loads guidance when needed instead of bloating every tool description.",
    inputSchema: {
      type: "object",
      properties: {
        topic: {
          type: "string",
          description: "Optional topic filter: sharepoint, upload, mail, calendar, teams, transcript, excel, word, or leave empty for all workflows.",
        },
      },
    },
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

  toolsLoaded = true;
  return { tools: allTools };
});

// Call tool handler
server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  // Wait for tools/list to complete at least once (avoids race condition)
  if (!toolsLoaded) {
    const maxWait = 15000;
    const interval = 100;
    let waited = 0;
    while (!toolsLoaded && waited < maxWait) {
      await new Promise(r => setTimeout(r, interval));
      waited += interval;
    }
  }

  // Handle built-in auth tool
  if (name === "agent365_authenticate") {
    return await handleAuthenticate();
  }

  // Handle built-in upload tool
  if (name === "sharepoint_uploadLocalFile") {
    return await handleUploadLocalFile(args);
  }

  // Handle built-in meeting transcript tool
  if (name === "teams_getMeetingTranscript") {
    return await handleGetMeetingTranscript(args);
  }

  // Handle built-in Graph API auth tool
  if (name === "agent365_graph_auth") {
    return await handleGraphAuth();
  }

  if (name === "agent365_help") {
    return handleHelp(args);
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

Get these values from entra.microsoft.com (Tenant ID on Home, Client ID under Home > App Registrations > Agent 365 MCP). If you don't have Entra access, ask your IT administrator, then restart your coding assistant.`,
      }],
      isError: true,
    };
  }

  // Check if already authenticated
  const existingToken = await loadToken();
  if (existingToken) {
    const tokenData = loadTokenData();

    // Clear cached sessions and notify client to re-fetch tools
    // (handles the case where tools weren't loaded on initial startup)
    Object.keys(serverSessions).forEach(key => delete serverSessions[key]);
    Object.keys(toolServerMap).forEach(key => delete toolServerMap[key]);
    try {
      await server.sendToolListChanged();
      console.error("📢 Notified client to refresh tool list (already authenticated)");
    } catch (e) {
      // Ignore - client may not support notifications
    }

    return {
      content: [{
        type: "text",
        text: `✅ **Already Authenticated**

Account: ${tokenData?.account?.username || "Unknown"}
Status: Valid token present

Tool list has been refreshed. The agent365_* tools (SharePoint, Excel, Word, Teams, Mail, Calendar, Copilot) should now be available.`,
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

      // Clear cached sessions and notify client to re-fetch tools
      Object.keys(serverSessions).forEach(key => delete serverSessions[key]);
      Object.keys(toolServerMap).forEach(key => delete toolServerMap[key]);
      try {
        await server.sendToolListChanged();
        console.error("📢 Notified client to refresh tool list (token refreshed)");
      } catch (e) {
        // Ignore
      }

      return {
        content: [{
          type: "text",
          text: `✅ **Token Refreshed**

Account: ${response.account?.username || "Unknown"}
Expires: ${response.expiresOn}

Tool list has been refreshed. The agent365_* tools (SharePoint, Excel, Word, Teams, Mail, Calendar, Copilot) should now be available.`,
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
    authPromise.then(async (response) => {
      saveToken(response, config);
      console.error(`✅ Authentication completed for ${response.account?.username}`);

      // Clear cached sessions so tools are re-fetched with the new token
      Object.keys(serverSessions).forEach(key => delete serverSessions[key]);
      Object.keys(toolServerMap).forEach(key => delete toolServerMap[key]);

      // Notify client that tools have changed so it re-fetches the full tool list
      try {
        await server.sendToolListChanged();
        console.error("📢 Notified client to refresh tool list");
      } catch (e) {
        console.error(`Could not notify client of tool list change: ${e.message}`);
      }
    }).catch((e) => {
      console.error(`❌ Authentication failed: ${e.message}`);
    });
  });
}

// ============================================================================
// LOCAL FILE UPLOAD TO SHAREPOINT/ONEDRIVE
// Reads a local file and uploads via the upstream createSmallBinaryFile tool.
// ============================================================================

const UPLOAD_MAX_SIZE_SMALL = 4 * 1024 * 1024; // 4MB limit for Agent365 proxy upload
const UPLOAD_MAX_SIZE_LARGE = 250 * 1024 * 1024; // 250MB limit for Graph API chunked upload
const UPLOAD_CHUNK_SIZE = 5 * 320 * 1024; // 1.6MB chunks (must be multiple of 320KB)

async function handleUploadLocalFile(args) {
  const { localFilePath, documentLibraryId, parentFolderId, filename } = args || {};

  if (!localFilePath) {
    return {
      content: [{ type: "text", text: "Error: localFilePath is required." }],
      isError: true,
    };
  }

  if (!documentLibraryId) {
    return {
      content: [{ type: "text", text: "Error: documentLibraryId is required. Use 'me' for OneDrive, or find it via sharepoint_findSite + sharepoint_listDocumentLibrariesInSite." }],
      isError: true,
    };
  }

  // Resolve and validate file path
  const resolvedPath = path.resolve(localFilePath);
  if (!fs.existsSync(resolvedPath)) {
    return {
      content: [{ type: "text", text: `Error: File not found: ${resolvedPath}` }],
      isError: true,
    };
  }

  const stats = fs.statSync(resolvedPath);
  if (!stats.isFile()) {
    return {
      content: [{ type: "text", text: `Error: Path is not a file: ${resolvedPath}` }],
      isError: true,
    };
  }

  const uploadFilename = filename || path.basename(resolvedPath);

  if (stats.size > UPLOAD_MAX_SIZE_LARGE) {
    const sizeMB = (stats.size / (1024 * 1024)).toFixed(1);
    return {
      content: [{ type: "text", text: `Error: File too large (${sizeMB}MB). Maximum upload size is ${UPLOAD_MAX_SIZE_LARGE / (1024 * 1024)}MB.` }],
      isError: true,
    };
  }

  // Large files (>4MB) use Graph API chunked upload
  if (stats.size > UPLOAD_MAX_SIZE_SMALL) {
    return await uploadLargeFile(resolvedPath, uploadFilename, documentLibraryId, parentFolderId || "root");
  }

  if (stats.size === 0) {
    return {
      content: [{ type: "text", text: "Error: File is empty (0 bytes)." }],
      isError: true,
    };
  }

  // Determine if text or binary based on extension
  const textExtensions = new Set([
    ".txt", ".md", ".csv", ".json", ".xml", ".html", ".htm", ".css", ".js",
    ".ts", ".py", ".sh", ".bash", ".yaml", ".yml", ".toml", ".ini", ".cfg",
    ".log", ".sql", ".r", ".m", ".c", ".h", ".cpp", ".hpp", ".java",
    ".rb", ".go", ".rs", ".swift", ".kt", ".scala", ".pl", ".lua",
  ]);
  const ext = path.extname(resolvedPath).toLowerCase();
  const isText = textExtensions.has(ext);

  try {
    await initializeServer("sharepoint");

    if (isText) {
      const content = fs.readFileSync(resolvedPath, "utf8");
      const result = await makeAgent365Request(
        MCP_SERVERS.sharepoint.url,
        "tools/call",
        {
          name: "createSmallTextFile",
          arguments: {
            filename: uploadFilename,
            contentText: content,
            documentLibraryId: documentLibraryId,
            parentfolderId: parentFolderId || "root",
          },
        },
        Date.now()
      );

      if (result.error) {
        throw new Error(result.error.message);
      }

      return processResult(result.result, "sharepoint_uploadLocalFile");
    } else {
      // Binary file: base64 encode and use createSmallBinaryFile
      const content = fs.readFileSync(resolvedPath);
      const base64Content = content.toString("base64");

      const result = await makeAgent365Request(
        MCP_SERVERS.sharepoint.url,
        "tools/call",
        {
          name: "createSmallBinaryFile",
          arguments: {
            filename: uploadFilename,
            base64Content: base64Content,
            documentLibraryId: documentLibraryId,
            parentfolderId: parentFolderId || "root",
          },
        },
        Date.now()
      );

      if (result.error) {
        throw new Error(result.error.message);
      }

      return processResult(result.result, "sharepoint_uploadLocalFile");
    }
  } catch (error) {
    return {
      content: [{
        type: "text",
        text: `Error uploading ${uploadFilename}: ${error.message}\n\nIf authentication expired, run agent365_authenticate.`,
      }],
      isError: true,
    };
  }
}

// ============================================================================
// LARGE FILE UPLOAD VIA GRAPH API (>4MB, up to 250MB)
// Uses resumable upload sessions with chunked PUT requests.
// ============================================================================

async function uploadLargeFile(resolvedPath, uploadFilename, documentLibraryId, parentFolderId) {
  // Get Graph token
  const graphToken = await loadGraphToken();
  if (!graphToken) {
    return {
      content: [{
        type: "text",
        text: `Error: Large file upload (>4MB) requires Microsoft Graph API authentication. Call the agent365_graph_auth tool first to authenticate with Graph API permissions (Files.ReadWrite.All, Sites.ReadWrite.All).`,
      }],
      isError: true,
    };
  }

  try {
    // Step 1: Create upload session
    // Resolve folder path for correct placement (ID:path: syntax doesn't work reliably with createUploadSession)
    const driveId = documentLibraryId === "me" ? "me/drive" : `drives/${documentLibraryId}`;
    let sessionUrl;
    if (parentFolderId === "root") {
      sessionUrl = `/${driveId}/root:/${encodeURIComponent(uploadFilename)}:/createUploadSession`;
    } else {
      // Get the folder's path from Graph API, then use path-based upload
      const folderInfo = await makeGraphRequest("GET", `/${driveId}/items/${parentFolderId}?$select=name,parentReference`);
      if (folderInfo.error) {
        return {
          content: [{ type: "text", text: `Error resolving folder: ${folderInfo.error}` }],
          isError: true,
        };
      }
      const folderPath = folderInfo.parentReference?.path
        ? `${folderInfo.parentReference.path}/${folderInfo.name}`
        : `/${driveId}/root:/${folderInfo.name}`;
      sessionUrl = `${folderPath}/${encodeURIComponent(uploadFilename)}:/createUploadSession`;
    }

    const sessionResult = await makeGraphRequest("POST", sessionUrl, {
      item: {
        "@microsoft.graph.conflictBehavior": "rename",
      },
    });

    if (sessionResult.error) {
      return {
        content: [{
          type: "text",
          text: `Error creating upload session: ${sessionResult.error}`,
        }],
        isError: true,
      };
    }

    const uploadUrl = sessionResult.uploadUrl;
    if (!uploadUrl) {
      return {
        content: [{
          type: "text",
          text: "Error: No upload URL returned from Graph API createUploadSession.",
        }],
        isError: true,
      };
    }

    // Step 2: Read and upload file in chunks
    const fileSize = fs.statSync(resolvedPath).size;
    const fd = fs.openSync(resolvedPath, "r");
    let offset = 0;
    let lastResponse = null;

    try {
      while (offset < fileSize) {
        const chunkSize = Math.min(UPLOAD_CHUNK_SIZE, fileSize - offset);
        const chunk = Buffer.alloc(chunkSize);
        fs.readSync(fd, chunk, 0, chunkSize, offset);

        const rangeEnd = offset + chunkSize - 1;
        const contentRange = `bytes ${offset}-${rangeEnd}/${fileSize}`;

        console.error(`Uploading chunk: ${contentRange} (${(chunkSize / 1024).toFixed(0)}KB)`);

        const chunkResult = await makeGraphRawRequest("PUT", uploadUrl, chunk, {
          "Content-Range": contentRange,
          "Content-Type": "application/octet-stream",
        });

        if (chunkResult.error) {
          return {
            content: [{
              type: "text",
              text: `Error uploading chunk at offset ${offset}: ${chunkResult.error}`,
            }],
            isError: true,
          };
        }

        if (chunkResult.statusCode >= 400) {
          const errorData = typeof chunkResult.data === "string" ? chunkResult.data : JSON.stringify(chunkResult.data);
          return {
            content: [{
              type: "text",
              text: `Error uploading chunk at offset ${offset}: HTTP ${chunkResult.statusCode} - ${errorData.slice(0, 500)}`,
            }],
            isError: true,
          };
        }

        lastResponse = chunkResult.data;
        offset += chunkSize;
      }
    } finally {
      fs.closeSync(fd);
    }

    // Step 3: Return the created file metadata
    const sizeMB = (fileSize / (1024 * 1024)).toFixed(1);
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          message: `Successfully uploaded ${uploadFilename} (${sizeMB}MB) via Graph API chunked upload.`,
          file: lastResponse,
        }, null, 2),
      }],
    };
  } catch (error) {
    return {
      content: [{
        type: "text",
        text: `Error during large file upload of ${uploadFilename}: ${error.message}\n\nIf Graph API authentication expired, call agent365_graph_auth.`,
      }],
      isError: true,
    };
  }
}

// ============================================================================
// MEETING TRANSCRIPT RETRIEVAL VIA GRAPH API
// ============================================================================

function handleHelp(args) {
  const topic = (args?.topic || "").toLowerCase().trim();

  const sections = {
    auth: `## Authentication (required once per session)
- agent365_authenticate() → triggers device code login
- agent365_graph_auth() → separate auth for advanced features (large uploads >4MB, meeting transcripts)`,

    sharepoint: `## SharePoint / OneDrive — Finding Files
**Option A — Search directly:**
  sharepoint_findFileOrFolder(searchQuery) → returns items with id, webUrl, driveId

**Option B — Browse by site:**
  1. sharepoint_findSite(searchQuery) → get siteId
  2. sharepoint_listDocumentLibrariesInSite(siteId) → get documentLibraryId (driveId)
  3. sharepoint_getFolderChildren(documentLibraryId, parentFolderId) → browse files

**Reading files:**
  - .docx → word_GetDocumentContent(fileUrl)
  - .xlsx → excel_GetDocumentContent(fileUrl)
  - text files → sharepoint_readSmallTextFile(fileId, documentLibraryId)
  - binary files → sharepoint_readSmallBinaryFile(fileId, documentLibraryId)
  - From URL → sharepoint_getFileOrFolderMetadataByUrl(url) to resolve ids first`,

    upload: `## Uploading Files
**Local file → SharePoint/OneDrive:**
  1. Find destination: sharepoint_findSite → listDocumentLibrariesInSite → getFolderChildren
  2. sharepoint_uploadLocalFile(localFilePath, documentLibraryId, parentFolderId)
  - Files ≤4MB: direct upload via upstream API
  - Files 4-250MB: Graph API chunked upload (requires agent365_graph_auth)
  - For OneDrive: use documentLibraryId='me'

**Creating files to upload:**
  Use docx, xlsx, pptx, or pdf skills to create files locally, then upload with sharepoint_uploadLocalFile.

**Other upload methods:**
  - sharepoint_createSmallTextFile — upload text content directly (no local file needed)
  - sharepoint_createSmallBinaryFile — upload base64-encoded binary content
  - sharepoint_uploadFileFromUrl — copy from one SharePoint location to another`,

    mail: `## Email (Outlook)
**Reading:**
  mail_SearchMessages(query) → mail_GetMessage(id) → mail_GetAttachments(id) → mail_DownloadAttachment(id, attachmentId)

**Sending:**
  mail_SendEmailWithAttachments(to, subject, body, directAttachmentFilePaths=["/path/to/file"])
  - Supports local files via directAttachmentFilePaths
  - Supports SharePoint/OneDrive files via attachmentUris
  - Create attachments locally with docx, xlsx, pptx, pdf skills first

**Drafts:**
  mail_CreateDraftMessage → mail_UpdateDraft (add attachments) → mail_SendDraftMessage

**Replying:**
  mail_ReplyToMessage(id) — single recipient
  mail_ReplyAllToMessage(id) — all recipients
  mail_ForwardMessage(id) — forward to new recipients`,

    calendar: `## Calendar
**Finding meetings (IMPORTANT: use ListCalendarView for recurring meetings):**
  calendar_ListCalendarView(userIdentifier='me', startDateTime, endDateTime, subject)
  - Expands recurring events (standups, syncs, scrums, 1:1s) into individual instances
  - Use narrow date ranges (1-2 days) to avoid large responses

  calendar_ListEvents(startDateTime, endDateTime, meetingTitle)
  - Only returns master events — will MISS recurring meeting instances
  - Use only for one-off meetings

**Creating/modifying:**
  calendar_FindMeetingTimes → calendar_CreateEvent
  calendar_UpdateEvent, calendar_CancelEvent
  calendar_AcceptEvent, calendar_DeclineEvent, calendar_TentativelyAcceptEvent`,

    transcript: `## Meeting Transcripts
**Workflow:**
  1. calendar_ListCalendarView(userIdentifier='me', startDateTime, endDateTime, subject) → find the meeting
  2. Get onlineMeeting.joinUrl from the event response
  3. teams_getMeetingTranscript(meetingUrl=joinUrl) → returns VTT transcript

**Requirements:**
  - Transcription must have been enabled during the meeting
  - Requires Graph API auth: call agent365_graph_auth if you get auth errors
  - Use ListCalendarView (NOT ListEvents) — most meetings are recurring`,

    teams: `## Teams
**Browsing:**
  teams_ListTeams → teams_ListChannels(teamId) → teams_ListChannelMessages(teamId, channelId)
  teams_ListChats → teams_ListChatMessages(chatId)

**Searching:**
  teams_SearchTeamsMessages(searchQuery) — search across all Teams messages

**Posting (requires user confirmation):**
  teams_PostChannelMessage(teamId, channelId, message)
  teams_PostMessage(chatId, message)`,

    word: `## Word Documents
**Reading:** word_GetDocumentContent(fileUrl) — extracts text, comments, structure
  Find file first: sharepoint_findFileOrFolder(searchQuery) → use webUrl

**Creating:**
  word_CreateDocument(fileName, contentInHtml) — creates in OneDrive root, accepts HTML/plain text
  For richer formatting (TOC, headers/footers, page numbers): create locally with docx skill, then sharepoint_uploadLocalFile`,

    excel: `## Excel Spreadsheets
**Reading:** excel_GetDocumentContent(fileUrl) — returns cell data as structured text
  Find file first: sharepoint_findFileOrFolder(searchQuery='.xlsx') → use webUrl

**Creating:**
  excel_CreateWorkbook(fileName) — creates in OneDrive
  For complex spreadsheets (formulas, charts, formatting): create locally with xlsx skill, then sharepoint_uploadLocalFile`,

    me: `## User / Directory Lookup
  me_GetMyDetails — current user's profile
  me_GetUserDetails(identifier) — look up by name or email
  me_GetMultipleUsersDetails — batch lookup
  me_GetManagerDetails — current user's manager
  me_GetDirectReportsDetails — direct reports`,

    copilot: `## M365 Copilot (requires Copilot license)
  copilot_copilot_chat(query) — semantic search across ALL M365 content
  Use when you don't know exact file names or need to search across emails, chats, files, and meetings simultaneously.`,
  };

  let output;
  if (topic && sections[topic]) {
    output = sections[topic];
  } else {
    // Progressive disclosure: return compact overview, not all details
    output = `# Agent 365 MCP — Workflow Guide

## Available capabilities (call agent365_help with a topic for details):

- **auth** — Authentication setup (agent365_authenticate, agent365_graph_auth)
- **sharepoint** — Find files/folders, browse sites, read documents
- **upload** — Upload local files to SharePoint/OneDrive (supports docx, xlsx, pptx, pdf, etc up to 250MB)
- **mail** — Search, read, send, reply, forward emails with attachments
- **calendar** — Find meetings, create events, manage invitations
- **transcript** — Get Teams meeting transcripts (VTT)
- **teams** — Browse/search chats, channels, post messages
- **word** — Read/create Word documents
- **excel** — Read/create Excel spreadsheets
- **me** — User profiles, org chart, directory lookup
- **copilot** — Semantic search across all M365 content (requires Copilot license)

## Quick tips:
- Always authenticate first: agent365_authenticate()
- For recurring meetings (standups, syncs, scrums): use calendar_ListCalendarView, NOT calendar_ListEvents
- For reading .docx use word_GetDocumentContent, for .xlsx use excel_GetDocumentContent
- To upload local files: sharepoint_uploadLocalFile (use docx/xlsx/pptx/pdf skills to create files first)
- For meeting transcripts: calendar_ListCalendarView → get joinUrl → teams_getMeetingTranscript`;
  }

  return {
    content: [{ type: "text", text: output }],
  };
}

async function handleGetMeetingTranscript(args) {
  const { meetingUrl } = args || {};

  if (!meetingUrl) {
    return {
      content: [{
        type: "text",
        text: "Error: meetingUrl is required. Use calendar_ListEvents or calendar_ListCalendarView to find the meeting first, then pass its onlineMeeting.joinUrl here.",
      }],
      isError: true,
    };
  }

  // Get Graph token
  const graphToken = await loadGraphToken();
  if (!graphToken) {
    return {
      content: [{
        type: "text",
        text: `Error: Meeting transcript retrieval requires Microsoft Graph API authentication. Call the agent365_graph_auth tool first.`,
      }],
      isError: true,
    };
  }

  try {
    // Find meeting by join URL
    const result = await makeGraphRequest(
      "GET",
      `/me/onlineMeetings?$filter=joinWebUrl eq '${meetingUrl}'`,
    );

    if (result.error) {
      return {
        content: [{
          type: "text",
          text: `Error finding meeting by URL: ${result.error}`,
        }],
        isError: true,
      };
    }

    const meetings = result.value || [];
    if (meetings.length === 0) {
      return {
        content: [{
          type: "text",
          text: `No meeting found with the provided join URL. Make sure you are the organizer or a participant.`,
        }],
        isError: true,
      };
    }

    const meetingInfo = meetings[0];
    const meetingId = meetingInfo.id;

    // Get transcripts for the meeting
    const transcriptsResult = await makeGraphRequest(
      "GET",
      `/me/onlineMeetings/${meetingId}/transcripts`,
    );

    if (transcriptsResult.error) {
      return {
        content: [{
          type: "text",
          text: `Error getting transcripts for meeting "${meetingInfo.subject || meetingId}": ${transcriptsResult.error}`,
        }],
        isError: true,
      };
    }

    const transcripts = transcriptsResult.value || [];
    if (transcripts.length === 0) {
      return {
        content: [{
          type: "text",
          text: `No transcripts found for meeting "${meetingInfo.subject || "Unknown"}". Make sure transcription was enabled during the meeting.`,
        }],
        isError: true,
      };
    }

    // Get the content of the first (most recent) transcript
    const transcriptId = transcripts[0].id;
    const contentResult = await makeGraphRequest(
      "GET",
      `/me/onlineMeetings/${meetingId}/transcripts/${transcriptId}/content?$format=text/vtt`,
    );

    if (contentResult.error) {
      return {
        content: [{
          type: "text",
          text: `Error getting transcript content: ${contentResult.error}`,
        }],
        isError: true,
      };
    }

    // Format the response
    const transcriptText = typeof contentResult === "string"
      ? contentResult
      : JSON.stringify(contentResult, null, 2);

    return {
      content: [{
        type: "text",
        text: `Meeting: ${meetingInfo.subject || "Unknown"}\nDate: ${meetingInfo.startDateTime || "Unknown"}\nOrganizer: ${meetingInfo.participants?.organizer?.upn || "Unknown"}\n\n--- Transcript ---\n\n${transcriptText}`,
      }],
    };
  } catch (error) {
    return {
      content: [{
        type: "text",
        text: `Error retrieving meeting transcript: ${error.message}\n\nIf Graph API authentication expired, call agent365_graph_auth.`,
      }],
      isError: true,
    };
  }
}

// ============================================================================
// GRAPH API AUTHENTICATION HANDLER
// ============================================================================

async function handleGraphAuth() {
  const config = loadServerConfig();
  if (!config) {
    return {
      content: [{
        type: "text",
        text: `Error: No Tenant ID or Client ID found. Set AGENT365_TENANT_ID and AGENT365_CLIENT_ID in your MCP config.`,
      }],
      isError: true,
    };
  }

  // Check if already authenticated with Graph
  const existingToken = await loadGraphToken();
  if (existingToken) {
    const tokenData = loadGraphTokenData();
    return {
      content: [{
        type: "text",
        text: `Already authenticated with Microsoft Graph API.\n\nAccount: ${tokenData?.account?.username || "Unknown"}\nScopes: ${GRAPH_SCOPES.join(", ")}\n\nGraph API features (large file upload >4MB, meeting transcripts) are ready to use.`,
      }],
    };
  }

  // Start device code flow for Graph scopes
  const pca = new PublicClientApplication({
    auth: {
      clientId: config.clientId,
      authority: `https://login.microsoftonline.com/${config.tenantId}`,
    },
    cache: {
      cachePlugin: createMsalCachePlugin(),
    },
  });

  // Try silent auth first
  const accounts = await pca.getTokenCache().getAllAccounts();
  if (accounts.length > 0) {
    try {
      const response = await pca.acquireTokenSilent({
        account: accounts[0],
        scopes: GRAPH_SCOPES,
      });
      saveGraphToken(response, config);
      return {
        content: [{
          type: "text",
          text: `Graph API token refreshed successfully.\n\nAccount: ${response.account?.username || "Unknown"}\nExpires: ${response.expiresOn}\nScopes: ${GRAPH_SCOPES.join(", ")}\n\nGraph API features (large file upload >4MB, meeting transcripts) are ready to use.`,
        }],
      };
    } catch (e) {
      // Fall through to device code
    }
  }

  // Device code flow
  return new Promise((resolve) => {
    let resolved = false;

    const authPromise = pca.acquireTokenByDeviceCode({
      scopes: GRAPH_SCOPES,
      deviceCodeCallback: (response) => {
        const deviceCode = response.userCode;
        const verificationUrl = response.verificationUriComplete || response.verificationUri;

        saveDeviceCodeToFile(deviceCode, response.verificationUri);
        copyToClipboard(deviceCode);

        resolved = true;
        resolve({
          content: [{
            type: "text",
            text: `Microsoft Graph API Authentication Required\n\nYour code: ${deviceCode} (copied to clipboard)\n\nSteps for user:\n1. Go to: ${verificationUrl}\n2. Enter the code: ${deviceCode}\n3. Sign in with your Microsoft account\n4. Approve the additional permissions (Files.ReadWrite.All, Sites.ReadWrite.All, OnlineMeetings.Read, OnlineMeetingTranscript.Read.All)\n\nFor agent: After the user confirms login, call agent365_graph_auth again to verify. This enables large file upload (>4MB) and meeting transcript retrieval.\n\nThe code expires in 15 minutes.`,
          }],
        });
      },
    });

    // Handle completion in background
    if (authPromise && typeof authPromise.then === "function") {
      authPromise.then((response) => {
        saveGraphToken(response, loadServerConfig());
        console.error(`Graph API authentication completed for ${response.account?.username}`);
      }).catch((e) => {
        console.error(`Graph API authentication failed: ${e.message}`);
      });
    }

    // If deviceCodeCallback was not called (e.g., MSAL error), resolve with a fallback
    if (!resolved) {
      // Give a short delay for the callback to fire
      setTimeout(() => {
        if (!resolved) {
          resolve({
            content: [{
              type: "text",
              text: "Graph API authentication initiated. If no device code appeared, try running agent365_authenticate first, then retry agent365_graph_auth.",
            }],
          });
        }
      }, 1000);
    }
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

No Tenant ID or Client ID found. Get these from entra.microsoft.com (Tenant ID on Home, Client ID under Home > App Registrations > Agent 365 MCP), or ask your IT administrator if you don't have access.

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
  console.error("Agent 365 MCP Proxy v1.3.2");
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

// ============================================================================
// EXPORTS (for testing)
// ============================================================================

if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    // Graph token management
    GRAPH_SCOPES,
    GRAPH_TOKEN_PATH,
    loadGraphToken,
    makeGraphRequest,
    // Upload constants
    UPLOAD_MAX_SIZE_SMALL,
    UPLOAD_MAX_SIZE_LARGE,
    UPLOAD_CHUNK_SIZE,
    // Feature handlers
    handleUploadLocalFile,
    uploadLargeFile,
    handleGetMeetingTranscript,
    handleGraphAuth,
    // Test helpers
    _resetGraphTokenCache: () => {
      cachedGraphToken = null;
      cachedGraphTokenExpiry = null;
    },
    _setGraphTokenCache: (token, expiry) => {
      cachedGraphToken = token;
      cachedGraphTokenExpiry = expiry;
    },
  };
}
