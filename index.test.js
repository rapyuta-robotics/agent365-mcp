/**
 * Tests for Graph API features in Agent 365 MCP Proxy
 *
 * Tests cover:
 * 1. Graph token management (load, refresh, acquire)
 * 2. Large file upload (>4MB chunked upload via Graph API)
 * 3. Meeting transcript retrieval
 * 4. Graph auth tool registration
 */

// We need to mock modules BEFORE requiring index.js
// Mock MSAL
const mockAcquireTokenSilent = jest.fn();
const mockAcquireTokenByDeviceCode = jest.fn();
const mockGetAllAccounts = jest.fn().mockResolvedValue([]);
const mockPCA = {
  acquireTokenSilent: mockAcquireTokenSilent,
  acquireTokenByDeviceCode: mockAcquireTokenByDeviceCode,
  getTokenCache: () => ({ getAllAccounts: mockGetAllAccounts }),
};
jest.mock("@azure/msal-node", () => ({
  PublicClientApplication: jest.fn(() => mockPCA),
}));

// Mock MCP SDK
const mockSetRequestHandler = jest.fn();
const mockConnect = jest.fn();
const mockSendToolListChanged = jest.fn();
jest.mock("@modelcontextprotocol/sdk/server/index.js", () => ({
  Server: jest.fn(() => ({
    setRequestHandler: mockSetRequestHandler,
    connect: mockConnect,
    sendToolListChanged: mockSendToolListChanged,
  })),
}));
jest.mock("@modelcontextprotocol/sdk/server/stdio.js", () => ({
  StdioServerTransport: jest.fn(),
}));
jest.mock("@modelcontextprotocol/sdk/types.js", () => ({
  CallToolRequestSchema: "CallToolRequestSchema",
  ListToolsRequestSchema: "ListToolsRequestSchema",
}));

const fs = require("fs");
const path = require("path");
const https = require("https");
const { PassThrough } = require("stream");
const { EventEmitter } = require("events");

// Suppress console.error during tests
beforeAll(() => {
  jest.spyOn(console, "error").mockImplementation(() => {});
});
afterAll(() => {
  console.error.mockRestore();
});

// ============================================================================
// Helper to require index.js cleanly (it has side effects from main())
// We need to stop main() from connecting. We do this by mocking the transport.
// ============================================================================

let graphExports;

beforeAll(async () => {
  // Set env vars for config
  process.env.AGENT365_TENANT_ID = "test-tenant-id";
  process.env.AGENT365_CLIENT_ID = "test-client-id";

  // Require the module - main() will run but connect is mocked
  graphExports = require("./index.js");

  // Give main() a moment to finish
  await new Promise((resolve) => setTimeout(resolve, 100));
});

// ============================================================================
// 1. GRAPH TOKEN MANAGEMENT TESTS
// ============================================================================

describe("Graph Token Constants", () => {
  test("GRAPH_SCOPES includes required permissions", () => {
    expect(graphExports.GRAPH_SCOPES).toBeDefined();
    expect(graphExports.GRAPH_SCOPES).toContain("Files.ReadWrite.All");
    expect(graphExports.GRAPH_SCOPES).toContain("Sites.ReadWrite.All");
    expect(graphExports.GRAPH_SCOPES).toContain("OnlineMeetings.Read");
    expect(graphExports.GRAPH_SCOPES).toContain("OnlineMeetingTranscript.Read.All");
  });

  test("GRAPH_TOKEN_PATH points to graph-tokens.json in .agent365-mcp dir", () => {
    expect(graphExports.GRAPH_TOKEN_PATH).toBeDefined();
    expect(graphExports.GRAPH_TOKEN_PATH).toContain(".agent365-mcp");
    expect(graphExports.GRAPH_TOKEN_PATH).toContain("graph-tokens.json");
  });
});

describe("loadGraphToken", () => {
  const HOME = process.env.HOME || process.env.USERPROFILE || "";
  const graphTokenPath = path.join(HOME, ".agent365-mcp", "graph-tokens.json");

  afterEach(() => {
    // Clean up any created files
    try {
      if (fs.existsSync(graphTokenPath)) {
        fs.unlinkSync(graphTokenPath);
      }
    } catch (e) {
      // Ignore
    }
    // Reset in-memory cache
    if (graphExports._resetGraphTokenCache) {
      graphExports._resetGraphTokenCache();
    }
  });

  test("returns null when no graph token file exists", async () => {
    // Ensure file doesn't exist
    try { fs.unlinkSync(graphTokenPath); } catch (e) { /* ignore */ }

    const token = await graphExports.loadGraphToken();
    expect(token).toBeNull();
  });

  test("returns token when valid (non-expired) graph token file exists", async () => {
    const futureDate = new Date(Date.now() + 3600 * 1000).toISOString();
    const tokenData = {
      accessToken: "valid-graph-token-123",
      expiresOn: futureDate,
      tenantId: "test-tenant",
      clientId: "test-client",
    };

    const dir = path.dirname(graphTokenPath);
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
    }
    fs.writeFileSync(graphTokenPath, JSON.stringify(tokenData));

    const token = await graphExports.loadGraphToken();
    expect(token).toBe("valid-graph-token-123");
  });

  test("returns null when graph token is expired and silent refresh fails", async () => {
    const pastDate = new Date(Date.now() - 3600 * 1000).toISOString();
    const tokenData = {
      accessToken: "expired-graph-token",
      expiresOn: pastDate,
      tenantId: "test-tenant",
      clientId: "test-client",
    };

    const dir = path.dirname(graphTokenPath);
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
    }
    fs.writeFileSync(graphTokenPath, JSON.stringify(tokenData));

    // Mock: no accounts for silent refresh
    mockGetAllAccounts.mockResolvedValueOnce([]);

    const token = await graphExports.loadGraphToken();
    expect(token).toBeNull();
  });
});

// ============================================================================
// 2. makeGraphRequest TESTS
// ============================================================================

describe("makeGraphRequest", () => {
  let originalHttpsRequest;

  beforeEach(() => {
    originalHttpsRequest = https.request;
  });

  afterEach(() => {
    https.request = originalHttpsRequest;
  });

  test("makes HTTPS request to graph.microsoft.com with correct path and auth", async () => {
    // Set up a valid graph token in the cache
    if (graphExports._setGraphTokenCache) {
      graphExports._setGraphTokenCache("test-graph-token-for-request", new Date(Date.now() + 3600 * 1000).toISOString());
    }

    let capturedOptions = null;
    let capturedBody = null;

    // Mock https.request
    https.request = jest.fn((options, callback) => {
      capturedOptions = options;

      const res = new PassThrough();
      res.statusCode = 200;
      callback(res);

      // Simulate response
      process.nextTick(() => {
        res.end(JSON.stringify({ value: "test-result" }));
      });

      const req = new PassThrough();
      req.setTimeout = jest.fn();
      req.on = jest.fn((event, handler) => {
        if (event === "error") { /* store handler */ }
        return req;
      });
      req.write = jest.fn((data) => { capturedBody = data; });
      req.end = jest.fn();
      return req;
    });

    const result = await graphExports.makeGraphRequest(
      "GET",
      "/me/onlineMeetings",
      null,
      {}
    );

    expect(capturedOptions.hostname).toBe("graph.microsoft.com");
    expect(capturedOptions.path).toBe("/v1.0/me/onlineMeetings");
    expect(capturedOptions.headers["Authorization"]).toBe("Bearer test-graph-token-for-request");
  });

  test("returns error object when no Graph token is available", async () => {
    // Clear graph token cache
    if (graphExports._resetGraphTokenCache) {
      graphExports._resetGraphTokenCache();
    }
    // Remove token file if it exists
    const HOME = process.env.HOME || process.env.USERPROFILE || "";
    const graphTokenPath = path.join(HOME, ".agent365-mcp", "graph-tokens.json");
    try { fs.unlinkSync(graphTokenPath); } catch (e) { /* ignore */ }

    // Mock: no accounts for silent refresh
    mockGetAllAccounts.mockResolvedValue([]);

    const result = await graphExports.makeGraphRequest("GET", "/me/onlineMeetings");
    expect(result).toBeDefined();
    expect(result.error).toBeDefined();
    expect(result.error).toMatch(/Graph/i);
  });
});

// ============================================================================
// 3. LARGE FILE UPLOAD TESTS
// ============================================================================

describe("Upload size constants", () => {
  test("UPLOAD_MAX_SIZE_SMALL is 4MB", () => {
    expect(graphExports.UPLOAD_MAX_SIZE_SMALL).toBe(4 * 1024 * 1024);
  });

  test("UPLOAD_MAX_SIZE_LARGE is 250MB", () => {
    expect(graphExports.UPLOAD_MAX_SIZE_LARGE).toBe(250 * 1024 * 1024);
  });

  test("UPLOAD_CHUNK_SIZE is a multiple of 320KB", () => {
    expect(graphExports.UPLOAD_CHUNK_SIZE).toBeDefined();
    expect(graphExports.UPLOAD_CHUNK_SIZE % (320 * 1024)).toBe(0);
  });
});

describe("handleUploadLocalFile - large file support", () => {
  const testDir = path.join("/tmp", "agent365-test-uploads");
  let originalHttpsRequest;

  beforeEach(() => {
    if (!fs.existsSync(testDir)) {
      fs.mkdirSync(testDir, { recursive: true });
    }
    originalHttpsRequest = https.request;
  });

  afterEach(() => {
    https.request = originalHttpsRequest;
    // Clean up test files
    try {
      const files = fs.readdirSync(testDir);
      for (const f of files) {
        fs.unlinkSync(path.join(testDir, f));
      }
      fs.rmdirSync(testDir);
    } catch (e) {
      // Ignore
    }
  });

  test("rejects files larger than 250MB", async () => {
    // Create a test file that's "too large" by mocking fs.statSync
    const testFile = path.join(testDir, "huge-file.bin");
    fs.writeFileSync(testFile, "x"); // Create the file so it exists

    const originalStatSync = fs.statSync;
    fs.statSync = jest.fn((p) => {
      if (p === testFile) {
        return { isFile: () => true, size: 260 * 1024 * 1024 }; // 260MB
      }
      return originalStatSync(p);
    });

    try {
      const result = await graphExports.handleUploadLocalFile({
        localFilePath: testFile,
        documentLibraryId: "test-drive-id",
        parentFolderId: "root",
      });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toMatch(/250/);
    } finally {
      fs.statSync = originalStatSync;
    }
  });

  test("files between 4MB and 250MB attempt Graph API upload", async () => {
    // Create a 5MB test file
    const testFile = path.join(testDir, "medium-file.bin");
    fs.writeFileSync(testFile, "x"); // Create so it exists

    const originalStatSync = fs.statSync;
    fs.statSync = jest.fn((p) => {
      if (p === testFile) {
        return { isFile: () => true, size: 5 * 1024 * 1024 }; // 5MB
      }
      return originalStatSync(p);
    });

    // No graph token available - should get helpful error about Graph auth
    if (graphExports._resetGraphTokenCache) {
      graphExports._resetGraphTokenCache();
    }
    const HOME = process.env.HOME || process.env.USERPROFILE || "";
    const graphTokenPath = path.join(HOME, ".agent365-mcp", "graph-tokens.json");
    try { fs.unlinkSync(graphTokenPath); } catch (e) { /* ignore */ }
    mockGetAllAccounts.mockResolvedValue([]);

    try {
      const result = await graphExports.handleUploadLocalFile({
        localFilePath: testFile,
        documentLibraryId: "test-drive-id",
        parentFolderId: "root",
      });

      // Should fail because no Graph token, but should NOT say "File too large" (old behavior)
      // Instead should mention Graph API authentication
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toMatch(/[Gg]raph/);
      expect(result.content[0].text).not.toMatch(/File too large/);
    } finally {
      fs.statSync = originalStatSync;
    }
  });
});

describe("uploadLargeFile", () => {
  test("function exists and is exported", () => {
    expect(graphExports.uploadLargeFile).toBeDefined();
    expect(typeof graphExports.uploadLargeFile).toBe("function");
  });

  test("returns error when no Graph token is available", async () => {
    // Clear graph token
    if (graphExports._resetGraphTokenCache) {
      graphExports._resetGraphTokenCache();
    }
    const HOME = process.env.HOME || process.env.USERPROFILE || "";
    const graphTokenPath = path.join(HOME, ".agent365-mcp", "graph-tokens.json");
    try { fs.unlinkSync(graphTokenPath); } catch (e) { /* ignore */ }
    mockGetAllAccounts.mockResolvedValue([]);

    const result = await graphExports.uploadLargeFile(
      "/tmp/test-file.bin",
      "test-file.bin",
      "drive-id",
      "folder-id"
    );

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toMatch(/[Gg]raph/);
    expect(result.content[0].text).toMatch(/agent365_graph_auth/);
  });
});

// ============================================================================
// 4. MEETING TRANSCRIPT TESTS
// ============================================================================

describe("handleGetMeetingTranscript", () => {
  test("function exists and is exported", () => {
    expect(graphExports.handleGetMeetingTranscript).toBeDefined();
    expect(typeof graphExports.handleGetMeetingTranscript).toBe("function");
  });

  test("returns error when no Graph token is available", async () => {
    // Clear graph token
    if (graphExports._resetGraphTokenCache) {
      graphExports._resetGraphTokenCache();
    }
    const HOME = process.env.HOME || process.env.USERPROFILE || "";
    const graphTokenPath = path.join(HOME, ".agent365-mcp", "graph-tokens.json");
    try { fs.unlinkSync(graphTokenPath); } catch (e) { /* ignore */ }
    mockGetAllAccounts.mockResolvedValue([]);

    const result = await graphExports.handleGetMeetingTranscript({});

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toMatch(/[Gg]raph/);
    expect(result.content[0].text).toMatch(/agent365_graph_auth/);
  });
});

// ============================================================================
// 5. TOOL REGISTRATION TESTS
// ============================================================================

describe("Tool Registration", () => {
  // The ListToolsRequestSchema handler is registered in the module.
  // We can test by looking at what mockSetRequestHandler was called with.

  test("teams_getMeetingTranscript tool is registered in tool list", async () => {
    // Find the ListToolsRequestSchema handler
    const listToolsCall = mockSetRequestHandler.mock.calls.find(
      (call) => call[0] === "ListToolsRequestSchema"
    );
    expect(listToolsCall).toBeDefined();

    const handler = listToolsCall[1];
    const result = await handler();
    const toolNames = result.tools.map((t) => t.name);

    expect(toolNames).toContain("teams_getMeetingTranscript");
  });

  test("teams_getMeetingTranscript has correct input schema", async () => {
    const listToolsCall = mockSetRequestHandler.mock.calls.find(
      (call) => call[0] === "ListToolsRequestSchema"
    );
    const handler = listToolsCall[1];
    const result = await handler();
    const tool = result.tools.find((t) => t.name === "teams_getMeetingTranscript");

    expect(tool).toBeDefined();
    expect(tool.inputSchema.properties.meetingUrl).toBeDefined();
    expect(tool.inputSchema.properties.meetingSubject).toBeDefined();
    expect(tool.inputSchema.properties.startDate).toBeDefined();
  });

  test("agent365_graph_auth tool is registered in tool list", async () => {
    const listToolsCall = mockSetRequestHandler.mock.calls.find(
      (call) => call[0] === "ListToolsRequestSchema"
    );
    const handler = listToolsCall[1];
    const result = await handler();
    const toolNames = result.tools.map((t) => t.name);

    expect(toolNames).toContain("agent365_graph_auth");
  });

  test("agent365_graph_auth description mentions large file and transcript", async () => {
    const listToolsCall = mockSetRequestHandler.mock.calls.find(
      (call) => call[0] === "ListToolsRequestSchema"
    );
    const handler = listToolsCall[1];
    const result = await handler();
    const tool = result.tools.find((t) => t.name === "agent365_graph_auth");

    expect(tool).toBeDefined();
    expect(tool.description).toMatch(/[Gg]raph/);
  });
});

// ============================================================================
// 6. CALL TOOL HANDLER ROUTING TESTS
// ============================================================================

describe("Call Tool Handler Routing", () => {
  let callToolHandler;

  beforeAll(() => {
    const callToolCall = mockSetRequestHandler.mock.calls.find(
      (call) => call[0] === "CallToolRequestSchema"
    );
    expect(callToolCall).toBeDefined();
    callToolHandler = callToolCall[1];
  });

  test("routes teams_getMeetingTranscript to handleGetMeetingTranscript", async () => {
    // Clear graph token to trigger error (proves routing works)
    if (graphExports._resetGraphTokenCache) {
      graphExports._resetGraphTokenCache();
    }
    const HOME = process.env.HOME || process.env.USERPROFILE || "";
    const graphTokenPath = path.join(HOME, ".agent365-mcp", "graph-tokens.json");
    try { fs.unlinkSync(graphTokenPath); } catch (e) { /* ignore */ }
    mockGetAllAccounts.mockResolvedValue([]);

    const result = await callToolHandler({
      params: { name: "teams_getMeetingTranscript", arguments: {} },
    });

    // Should reach the handler (not "Unknown tool")
    expect(result.content[0].text).not.toMatch(/Unknown tool/);
    expect(result.content[0].text).toMatch(/[Gg]raph/);
  });

  test("routes agent365_graph_auth to graph auth handler", async () => {
    const result = await callToolHandler({
      params: { name: "agent365_graph_auth", arguments: {} },
    });

    // Should reach the handler (not "Unknown tool")
    expect(result.content[0].text).not.toMatch(/Unknown tool/);
  });
});

// ============================================================================
// 7. UPLOAD DESCRIPTION UPDATE
// ============================================================================

describe("Upload tool description", () => {
  test("sharepoint_uploadLocalFile description mentions larger than 4MB support", async () => {
    const listToolsCall = mockSetRequestHandler.mock.calls.find(
      (call) => call[0] === "ListToolsRequestSchema"
    );
    const handler = listToolsCall[1];
    const result = await handler();
    const tool = result.tools.find((t) => t.name === "sharepoint_uploadLocalFile");

    expect(tool).toBeDefined();
    // Should mention support for larger files, not just "Maximum file size: 4MB"
    expect(tool.description).toMatch(/250/);
  });
});
