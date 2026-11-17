/**
 * Tests for the upload checks. Every path lives in a temporary directory; the real home
 * directory is never read or written.
 */
const fs = require("fs");
const os = require("os");
const path = require("path");
const { resolveUploadPath, checkUploadUrl, uploadRoots } = require("./upload-guard");

describe("resolveUploadPath", () => {
  let base;
  let home;
  let work;

  beforeEach(() => {
    base = fs.realpathSync(fs.mkdtempSync(path.join(os.tmpdir(), "agent365-guard-")));
    home = path.join(base, "home");
    work = path.join(home, "project");
    fs.mkdirSync(work, { recursive: true });
    fs.mkdirSync(path.join(home, ".ssh"));
    fs.mkdirSync(path.join(home, ".agent365-mcp"));
    fs.mkdirSync(path.join(home, ".config", "app"), { recursive: true });
    fs.mkdirSync(path.join(home, "Documents"));
    fs.writeFileSync(path.join(work, "report.docx"), "report");
    fs.writeFileSync(path.join(home, ".ssh", "id_ed25519"), "key");
    fs.writeFileSync(path.join(home, ".agent365-mcp", "tokens.json"), "{}");
    fs.writeFileSync(path.join(home, ".config", "app", "secret"), "s");
    fs.writeFileSync(path.join(home, "Documents", "plan.docx"), "plan");
  });

  afterEach(() => {
    fs.rmSync(base, { recursive: true, force: true });
  });

  const resolve = (p, env = {}) => resolveUploadPath(p, { home, cwd: work, env });

  test("allows a file in the working directory", () => {
    expect(resolve(path.join(work, "report.docx"))).toEqual({ path: path.join(work, "report.docx") });
  });

  test("resolves a relative path against the working directory", () => {
    expect(resolve("report.docx")).toEqual({ path: path.join(work, "report.docx") });
  });

  test("refuses a file outside the working directory", () => {
    expect(resolve(path.join(home, "Documents", "plan.docx")).error).toMatch(/outside the allowed/);
  });

  test("refuses a path that climbs out of the working directory", () => {
    expect(resolve("../Documents/plan.docx").error).toMatch(/outside the allowed/);
  });

  test("allows a directory listed in AGENT365_UPLOAD_ROOTS", () => {
    const env = { AGENT365_UPLOAD_ROOTS: [path.join(base, "elsewhere"), path.join(home, "Documents")].join(path.delimiter) };
    expect(resolve(path.join(home, "Documents", "plan.docx"), env)).toEqual({
      path: path.join(home, "Documents", "plan.docx"),
    });
    expect(resolve(path.join(work, "report.docx"), env).error).toMatch(/outside the allowed/);
  });

  test.each([
    [".ssh", "id_ed25519"],
    [".agent365-mcp", "tokens.json"],
    [".config", path.join("app", "secret")],
  ])("refuses ~/%s even when the home directory is an upload root", (dir, file) => {
    const result = resolve(path.join(home, dir, file), { AGENT365_UPLOAD_ROOTS: home });
    expect(result.error).toMatch(/not allowed/);
  });

  test("refuses a home dot-directory reached through a symlink", () => {
    fs.symlinkSync(path.join(home, ".ssh"), path.join(work, "keys"));
    expect(resolve(path.join(work, "keys", "id_ed25519"), { AGENT365_UPLOAD_ROOTS: home }).error).toMatch(/not allowed/);
    expect(resolve(path.join(work, "keys", "id_ed25519")).error).toBeDefined();
  });

  test("refuses a symlink in the working directory that points outside it", () => {
    fs.symlinkSync(path.join(home, "Documents", "plan.docx"), path.join(work, "link.docx"));
    expect(resolve(path.join(work, "link.docx")).error).toMatch(/outside the allowed/);
  });

  test("reports a missing file", () => {
    expect(resolve(path.join(work, "missing.docx")).error).toMatch(/File not found/);
  });

  test("the working directory is the default upload root", () => {
    expect(uploadRoots({}, work)).toEqual([work]);
  });
});

describe("checkUploadUrl", () => {
  test.each([
    "https://contoso.sharepoint.com/_api/v2.0/drives/d/items/i/uploadSession?guid=1",
    "https://contoso-my.sharepoint.com/personal/u/_api/v2.0/drive/items/i/uploadSession",
    "https://contoso.sharepoint-df.com/_api/v2.0/drives/d/uploadSession",
  ])("accepts %s", (url) => {
    expect(checkUploadUrl(url).hostname).toBe(new URL(url).hostname);
  });

  test.each([
    ["http://contoso.sharepoint.com/upload", /https/],
    ["https://attacker.example/upload", /not an expected Microsoft host/],
    ["https://sharepoint.com.attacker.example/upload", /not an expected Microsoft host/],
    ["https://evilsharepoint.com/upload", /not an expected Microsoft host/],
    ["https://user:pass@contoso.sharepoint.com/upload", /credentials/],
    ["not a url", /valid URL/],
  ])("refuses %s", (url, message) => {
    expect(() => checkUploadUrl(url)).toThrow(message);
  });
});
