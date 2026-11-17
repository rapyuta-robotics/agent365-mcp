const fs = require("fs");
const os = require("os");
const path = require("path");
const { writePrivateFile, tightenStateDir } = require("./secure-fs");

const posix = process.platform !== "win32";
const mode = (p) => fs.statSync(p).mode & 0o777;

describe("secure-fs", () => {
  let root;

  beforeEach(() => {
    root = fs.mkdtempSync(path.join(os.tmpdir(), "agent365-secure-fs-"));
  });

  afterEach(() => {
    fs.rmSync(root, { recursive: true, force: true });
  });

  (posix ? test : test.skip)("creates the directory and file for the owner only", () => {
    const file = path.join(root, "state", "msal-cache.json");
    writePrivateFile(file, "{}");

    expect(fs.readFileSync(file, "utf8")).toBe("{}");
    expect(mode(path.dirname(file))).toBe(0o700);
    expect(mode(file)).toBe(0o600);
  });

  (posix ? test : test.skip)("replaces a world-readable file with an owner-only one", () => {
    const file = path.join(root, "tokens.json");
    fs.writeFileSync(file, "old", { mode: 0o644 });
    fs.chmodSync(file, 0o644);

    writePrivateFile(file, "new");

    expect(fs.readFileSync(file, "utf8")).toBe("new");
    expect(mode(file)).toBe(0o600);
    expect(fs.readdirSync(root)).toEqual(["tokens.json"]);
  });

  (posix ? test : test.skip)("tightens files written by earlier versions", () => {
    const dir = path.join(root, ".agent365-mcp");
    fs.mkdirSync(dir, { mode: 0o755 });
    fs.chmodSync(dir, 0o755);
    const file = path.join(dir, "msal-cache.json");
    fs.writeFileSync(file, "{}");
    fs.chmodSync(file, 0o644);

    tightenStateDir(dir);

    expect(mode(dir)).toBe(0o700);
    expect(mode(file)).toBe(0o600);
  });
});
