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
const mockAcquireTokenInteractive = jest.fn();
const mockGetAllAccounts = jest.fn().mockResolvedValue([]);
const mockPCA = {
  acquireTokenSilent: mockAcquireTokenSilent,
  acquireTokenByDeviceCode: mockAcquireTokenByDeviceCode,
  acquireTokenInteractive: mockAcquireTokenInteractive,
  getTokenCache: () => ({ getAllAccounts: mockGetAllAccounts }),
};
jest.mock("@azure/msal-node", () => ({
  PublicClientApplication: jest.fn(() => mockPCA),
}));

// Mock child_process so tests never open a real browser or touch the clipboard
const mockSpawn = jest.fn(() => ({ on: jest.fn(), unref: jest.fn() }));
jest.mock("child_process", () => ({
  spawn: (...args) => mockSpawn(...args),
  exec: jest.fn(),
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
  const os = require("os");
  let testDir;
  let originalHttpsRequest;
  let originalUploadRoots;

  beforeEach(() => {
    testDir = fs.realpathSync(fs.mkdtempSync(path.join(os.tmpdir(), "agent365-test-uploads-")));
    originalUploadRoots = process.env.AGENT365_UPLOAD_ROOTS;
    process.env.AGENT365_UPLOAD_ROOTS = testDir;
    originalHttpsRequest = https.request;
  });

  afterEach(() => {
    https.request = originalHttpsRequest;
    if (originalUploadRoots === undefined) delete process.env.AGENT365_UPLOAD_ROOTS;
    else process.env.AGENT365_UPLOAD_ROOTS = originalUploadRoots;
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

  test("refuses a file outside the allowed upload directories", async () => {
    const outside = fs.realpathSync(fs.mkdtempSync(path.join(os.tmpdir(), "agent365-outside-")));
    const testFile = path.join(outside, "notes.txt");
    fs.writeFileSync(testFile, "x");
    try {
      const result = await graphExports.handleUploadLocalFile({
        localFilePath: testFile,
        documentLibraryId: "test-drive-id",
      });
      expect(result.isError).toBe(true);
      expect(result.content[0].text).toMatch(/outside the allowed upload directories/);
    } finally {
      fs.rmSync(outside, { recursive: true, force: true });
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

  describe("speaker attribution fallback", () => {
    let originalHttpsRequest;

    beforeEach(() => {
      originalHttpsRequest = https.request;
      if (graphExports._setGraphTokenCache) {
        graphExports._setGraphTokenCache(
          "test-graph-token-transcript",
          new Date(Date.now() + 3600 * 1000).toISOString(),
        );
      }
    });

    afterEach(() => {
      https.request = originalHttpsRequest;
      if (graphExports._resetGraphTokenCache) {
        graphExports._resetGraphTokenCache();
      }
    });

    /**
     * Drive handleGetMeetingTranscript through its three Graph calls.
     * `vttStatus` controls how the attributed (text/vtt) content call answers,
     * letting one helper cover both the happy path and the 403 fallback.
     */
    function mockGraphSequence(vttStatus) {
      const requests = [];

      https.request = jest.fn((options, callback) => {
        requests.push(options);
        const res = new PassThrough();
        res.headers = { "content-type": "application/json" };
        res.statusCode = 200;

        let payload;
        if (options.path.includes("$filter=joinWebUrl")) {
          payload = JSON.stringify({ value: [{ id: "MEETING-1", subject: "Sprint Review" }] });
        } else if (options.path.endsWith("/transcripts")) {
          payload = JSON.stringify({ value: [{ id: "TRANSCRIPT-1" }] });
        } else if (options.path.includes("$format=text/vtt")) {
          res.statusCode = vttStatus;
          res.headers = { "content-type": "text/plain" };
          payload = vttStatus === 200
            ? "WEBVTT\n\n00:00:01.000 --> 00:00:02.000\n<v Ada>Hello</v>"
            : JSON.stringify({
              error: {
                code: "Forbidden",
                message: "Speaker-attributed transcript content is disabled for this tenant.",
                innerError: { code: "SpeakerAttributionNotAllowed" },
              },
            });
        } else {
          res.headers = { "content-type": "text/plain" };
          payload = "00:00:01.000 --> 00:00:02.000\nHello";
        }

        callback(res);
        process.nextTick(() => res.end(payload));

        const req = new PassThrough();
        req.setTimeout = jest.fn();
        req.on = jest.fn(() => req);
        req.write = jest.fn();
        req.end = jest.fn();
        return req;
      });

      return requests;
    }

    test("falls back to un-attributed text when tenant blocks speaker attribution", async () => {
      const requests = mockGraphSequence(403);

      const result = await graphExports.handleGetMeetingTranscript({
        meetingUrl: "https://teams.microsoft.com/l/meetup-join/test",
      });

      expect(result.isError).toBeUndefined();

      // The retry must use the Accept header, not a $format query param.
      const retry = requests[requests.length - 1];
      expect(retry.path).not.toContain("$format");
      expect(retry.headers["Accept"]).toBe("application/vnd.microsoft.graph.transcript+text");

      // Content is returned, and the caller is warned that names are absent.
      expect(result.content[0].text).toContain("Hello");
      expect(result.content[0].text).toMatch(/does NOT identify who said each line/);
      expect(result.content[0].text).toMatch(/EnableAttributedTranscripts/);
    });

    test("keeps attributed VTT and adds no warning when attribution is allowed", async () => {
      const requests = mockGraphSequence(200);

      const result = await graphExports.handleGetMeetingTranscript({
        meetingUrl: "https://teams.microsoft.com/l/meetup-join/test",
      });

      expect(result.isError).toBeUndefined();
      // Exactly three calls: find meeting, list transcripts, fetch VTT — no retry.
      expect(requests).toHaveLength(3);
      expect(requests[2].path).toContain("$format=text/vtt");
      expect(result.content[0].text).toContain("WEBVTT");
      expect(result.content[0].text).not.toMatch(/does NOT identify who said each line/);
    });
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

// ============================================================================
// BROWSER SIGN-IN AND SIGN-IN-REQUIRED MARKER
// ============================================================================

describe("Browser sign-in", () => {
  const HOME = process.env.HOME || process.env.USERPROFILE || "";
  const agentDir = path.join(HOME, ".agent365-mcp");
  const tokenPath = path.join(agentDir, "tokens.json");
  const graphTokenPath = path.join(agentDir, "graph-tokens.json");

  const fakeResponse = {
    accessToken: "new-access-token",
    expiresOn: new Date(Date.now() + 3600 * 1000).toISOString(),
    account: { username: "user@example.com" },
  };

  // Simulate MSAL: open the browser, then finish sign-in when `finish()` is called
  function interactiveThatWaits() {
    let finish;
    let fail;
    mockAcquireTokenInteractive.mockImplementationOnce(async (request) => {
      await request.openBrowser("https://login.microsoftonline.com/test/oauth2/v2.0/authorize?client_id=x&state=y");
      return new Promise((resolve, reject) => {
        finish = () => resolve(fakeResponse);
        fail = (e) => reject(e);
      });
    });
    return { finish: () => finish(), fail: (e) => fail(e) };
  }

  function cleanup() {
    for (const p of [tokenPath, graphTokenPath, graphExports.SIGN_IN_REQUIRED_PATH]) {
      try { fs.unlinkSync(p); } catch (e) { /* ignore */ }
    }
  }

  beforeEach(() => {
    cleanup();
    mockAcquireTokenInteractive.mockReset();
    mockAcquireTokenByDeviceCode.mockReset();
    mockAcquireTokenSilent.mockReset();
    mockGetAllAccounts.mockReset();
    mockGetAllAccounts.mockResolvedValue([]);
    mockSpawn.mockClear();
    graphExports._resetPendingSignIns();
    graphExports._resetTokenCache();
    graphExports._resetGraphTokenCache();
  });

  afterAll(cleanup);

  test("browser is the default sign-in flow", () => {
    expect(graphExports.AUTH_FLOW).toBe("browser");
  });

  test("startInteractiveSignIn opens the browser without a shell and never uses device code", async () => {
    const sim = interactiveThatWaits();
    const session = await graphExports.startInteractiveSignIn(mockPCA, ["scope"]);

    expect(session.flow).toBe("browser");
    expect(session.url).toContain("login.microsoftonline.com");
    expect(mockAcquireTokenByDeviceCode).not.toHaveBeenCalled();
    expect(mockAcquireTokenInteractive).toHaveBeenCalledWith(
      expect.objectContaining({ scopes: ["scope"], openBrowser: expect.any(Function) })
    );
    expect(mockSpawn).toHaveBeenCalledTimes(1);
    const [cmd, args, opts] = mockSpawn.mock.calls[0];
    expect(["xdg-open", "open", "rundll32"]).toContain(cmd);
    expect(args[args.length - 1]).toBe(session.url);
    expect(opts.shell).toBeUndefined();

    sim.finish();
    await expect(session.done).resolves.toBe(fakeResponse);
  });

  test("startInteractiveSignIn uses device code only when asked for", async () => {
    mockAcquireTokenByDeviceCode.mockImplementationOnce(async (request) => {
      request.deviceCodeCallback({ userCode: "ABCD1234", verificationUri: "https://microsoft.com/devicelogin" });
      return fakeResponse;
    });
    const session = await graphExports.startInteractiveSignIn(mockPCA, ["scope"], "device_code");

    expect(session.flow).toBe("device_code");
    expect(session.userCode).toBe("ABCD1234");
    expect(mockAcquireTokenInteractive).not.toHaveBeenCalled();
  });

  test("startInteractiveSignIn rejects when the flow fails before it starts", async () => {
    mockAcquireTokenInteractive.mockRejectedValueOnce(new Error("loopback failed"));
    await expect(graphExports.startInteractiveSignIn(mockPCA, ["scope"])).rejects.toThrow("loopback failed");
  });

  test("beginSignIn reuses a pending sign-in instead of opening another window", async () => {
    const sim = interactiveThatWaits();
    const first = await graphExports.beginSignIn("agent365", mockPCA, ["scope"], jest.fn());
    const second = await graphExports.beginSignIn("agent365", mockPCA, ["scope"], jest.fn());

    expect(second).toBe(first);
    expect(mockAcquireTokenInteractive).toHaveBeenCalledTimes(1);
    sim.finish();
    await first.done;
  });

  test("handleAuthenticate returns browser instructions and saves the token when sign-in completes", async () => {
    const sim = interactiveThatWaits();
    const result = await graphExports.handleAuthenticate();
    const text = result.content[0].text;

    expect(result.isError).toBeUndefined();
    expect(text).toContain("browser window");
    expect(text).toContain("login.microsoftonline.com");
    expect(text).not.toContain("Your code");
    expect(mockAcquireTokenByDeviceCode).not.toHaveBeenCalled();

    sim.finish();
    await new Promise((resolve) => setTimeout(resolve, 10));
    const saved = JSON.parse(fs.readFileSync(tokenPath, "utf8"));
    expect(saved.accessToken).toBe("new-access-token");
  });

  test("handleGraphAuth returns browser instructions", async () => {
    const sim = interactiveThatWaits();
    const result = await graphExports.handleGraphAuth();
    const text = result.content[0].text;

    expect(text).toContain("browser window");
    expect(text).toContain("Files.ReadWrite.All");
    expect(mockAcquireTokenByDeviceCode).not.toHaveBeenCalled();
    sim.finish();
  });

  test("needsFreshSignIn matches rejected refresh tokens but not transient errors", () => {
    expect(graphExports.needsFreshSignIn({ errorCode: "invalid_grant" })).toBe(true);
    expect(graphExports.needsFreshSignIn({ name: "InteractionRequiredAuthError" })).toBe(true);
    expect(graphExports.needsFreshSignIn({ errorCode: "network_error" })).toBe(false);
    expect(graphExports.needsFreshSignIn(undefined)).toBe(false);
  });

  test("loadToken stops retrying silent refresh after Entra rejects the session", async () => {
    fs.mkdirSync(agentDir, { recursive: true });
    fs.writeFileSync(tokenPath, JSON.stringify({
      accessToken: "expired",
      expiresOn: new Date(Date.now() - 3600 * 1000).toISOString(),
      tenantId: "test-tenant",
      clientId: "test-client",
    }));
    mockGetAllAccounts.mockResolvedValue([{ username: "user@example.com" }]);
    const rejected = Object.assign(new Error("invalid_grant: AADSTS50173: The provided grant has expired"), { errorCode: "invalid_grant" });
    mockAcquireTokenSilent.mockRejectedValue(rejected);

    expect(await graphExports.loadToken()).toBeNull();
    expect(mockAcquireTokenSilent).toHaveBeenCalledTimes(1);
    expect(graphExports.isSignInRequired("agent365")).toBe(true);

    // A later call (or a new server process) must not hit Entra again
    expect(await graphExports.loadToken()).toBeNull();
    expect(mockAcquireTokenSilent).toHaveBeenCalledTimes(1);
  });

  test("a transient refresh error does not block later retries", async () => {
    fs.mkdirSync(agentDir, { recursive: true });
    fs.writeFileSync(tokenPath, JSON.stringify({
      accessToken: "expired",
      expiresOn: new Date(Date.now() - 3600 * 1000).toISOString(),
      tenantId: "test-tenant",
      clientId: "test-client",
    }));
    mockGetAllAccounts.mockResolvedValue([{ username: "user@example.com" }]);
    mockAcquireTokenSilent.mockRejectedValue(Object.assign(new Error("ENOTFOUND"), { errorCode: "network_error" }));

    await graphExports.loadToken();
    await graphExports.loadToken();
    expect(mockAcquireTokenSilent).toHaveBeenCalledTimes(2);
    expect(graphExports.isSignInRequired("agent365")).toBe(false);
  });

  test("a completed sign-in clears the sign-in-required marker", async () => {
    graphExports.markSignInRequired("agent365", { errorCode: "invalid_grant" });
    expect(graphExports.isSignInRequired("agent365")).toBe(true);

    const sim = interactiveThatWaits();
    await graphExports.handleAuthenticate();
    sim.finish();
    await new Promise((resolve) => setTimeout(resolve, 10));

    expect(graphExports.isSignInRequired("agent365")).toBe(false);
    expect(mockAcquireTokenSilent).not.toHaveBeenCalled();
  });
});

// ============================================================================
// UPLOAD SESSION REQUESTS
// ============================================================================

describe("makeUploadSessionRequest", () => {
  let originalHttpsRequest;

  beforeEach(() => {
    originalHttpsRequest = https.request;
    graphExports._setGraphTokenCache(
      "test-graph-token-upload",
      new Date(Date.now() + 3600 * 1000).toISOString(),
    );
  });

  afterEach(() => {
    https.request = originalHttpsRequest;
    graphExports._resetGraphTokenCache();
  });

  test("sends no Authorization header to the upload session host", async () => {
    let capturedOptions;
    https.request = jest.fn((options, callback) => {
      capturedOptions = options;
      const res = new PassThrough();
      res.statusCode = 202;
      callback(res);
      res.end("{}");
      const req = new EventEmitter();
      req.write = jest.fn();
      req.end = jest.fn();
      req.setTimeout = jest.fn();
      req.destroy = jest.fn();
      return req;
    });

    const result = await graphExports.makeUploadSessionRequest(
      "PUT",
      "https://contoso.sharepoint.com/_api/v2.0/drives/d/items/i/uploadSession?guid=1&tempauth=abc",
      Buffer.from("chunk"),
      { "Content-Range": "bytes 0-4/5", Authorization: "Bearer must-not-be-sent" },
    );

    expect(result.statusCode).toBe(202);
    expect(capturedOptions.hostname).toBe("contoso.sharepoint.com");
    const headerNames = Object.keys(capturedOptions.headers).map((h) => h.toLowerCase());
    expect(headerNames).not.toContain("authorization");
    expect(JSON.stringify(capturedOptions)).not.toContain("test-graph-token-upload");
  });

  test("makes no request to a host that is not a Microsoft upload host", async () => {
    https.request = jest.fn();
    const result = await graphExports.makeUploadSessionRequest(
      "PUT",
      "https://attacker.example/upload",
      Buffer.from("chunk"),
      {},
    );
    expect(result.error).toMatch(/not an expected Microsoft host/);
    expect(https.request).not.toHaveBeenCalled();
  });
});
