/**
 * Tests for the CLI's logout. HOME points at a temporary directory for the whole file,
 * so the real home directory is never read or written.
 */
const fs = require("fs");
const os = require("os");
const path = require("path");

jest.mock("@azure/msal-node", () => ({ PublicClientApplication: jest.fn() }));

describe("logout", () => {
  const saved = { HOME: process.env.HOME, USERPROFILE: process.env.USERPROFILE };
  let home;
  let logout;

  beforeAll(() => {
    home = fs.realpathSync(fs.mkdtempSync(path.join(os.tmpdir(), "agent365-cli-")));
    process.env.HOME = home;
    process.env.USERPROFILE = home;
    ({ logout } = require("./cli"));
    jest.spyOn(console, "log").mockImplementation(() => {});
  });

  afterAll(() => {
    for (const [key, value] of Object.entries(saved)) {
      if (value === undefined) delete process.env[key];
      else process.env[key] = value;
    }
    fs.rmSync(home, { recursive: true, force: true });
  });

  test("removes every saved token file, the Graph token included", () => {
    const dir = path.join(home, ".agent365-mcp");
    fs.mkdirSync(dir, { recursive: true });
    const files = ["tokens.json", "graph-tokens.json", "msal-cache.json", "sign-in-required.json"];
    for (const name of files) fs.writeFileSync(path.join(dir, name), "{}");

    logout();

    for (const name of files) {
      expect(fs.existsSync(path.join(dir, name))).toBe(false);
    }
  });

  test("does nothing when no token file exists", () => {
    expect(() => logout()).not.toThrow();
  });
});
