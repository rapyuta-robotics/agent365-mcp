// Checks for local file uploads: which files may leave the machine, and where they may go.
const fs = require("fs");
const path = require("path");

// Hosts that Microsoft Graph upload sessions for SharePoint and OneDrive for work use.
const UPLOAD_HOST_SUFFIXES = [".sharepoint.com", ".sharepoint-df.com", ".sharepoint.us"];

function isInside(child, parent) {
  const relative = path.relative(parent, child);
  return relative === "" || (!relative.startsWith("..") && !path.isAbsolute(relative));
}

function realpathOrSelf(p) {
  try {
    return fs.realpathSync(p);
  } catch (e) {
    return path.resolve(p);
  }
}

// The directories uploads may come from: AGENT365_UPLOAD_ROOTS, or the working directory.
function uploadRoots(env = process.env, cwd = process.cwd()) {
  const configured = (env.AGENT365_UPLOAD_ROOTS || "")
    .split(path.delimiter)
    .map((root) => root.trim())
    .filter(Boolean);
  const roots = configured.length > 0 ? configured : [cwd];
  return roots.map((root) => realpathOrSelf(root));
}

// Returns { path } with the real path of an uploadable file, or { error }.
function resolveUploadPath(localFilePath, options = {}) {
  const env = options.env || process.env;
  const home = options.home !== undefined ? options.home : env.HOME || env.USERPROFILE || "";
  const cwd = options.cwd || process.cwd();

  let real;
  try {
    real = fs.realpathSync(path.resolve(cwd, localFilePath));
  } catch (e) {
    return { error: `File not found: ${path.resolve(cwd, localFilePath)}` };
  }

  if (home) {
    const realHome = realpathOrSelf(home);
    if (isInside(real, realHome)) {
      const first = path.relative(realHome, real).split(path.sep)[0];
      if (first.startsWith(".")) {
        return { error: `Uploads from ${path.join("~", first)} are not allowed.` };
      }
    }
  }

  const roots = uploadRoots(env, cwd);
  if (!roots.some((root) => isInside(real, root))) {
    return {
      error:
        `File is outside the allowed upload directories (${roots.join(", ")}). ` +
        "Move the file there, or set AGENT365_UPLOAD_ROOTS.",
    };
  }

  return { path: real };
}

// Returns the parsed upload session URL, or throws when it is not a Microsoft upload host.
function checkUploadUrl(uploadUrl) {
  let parsed;
  try {
    parsed = new URL(uploadUrl);
  } catch (e) {
    throw new Error("Upload URL is not a valid URL.");
  }
  if (parsed.protocol !== "https:") {
    throw new Error("Upload URL must use https.");
  }
  if (parsed.username || parsed.password) {
    throw new Error("Upload URL must not carry credentials.");
  }
  const host = parsed.hostname.toLowerCase();
  if (!UPLOAD_HOST_SUFFIXES.some((suffix) => host.endsWith(suffix))) {
    throw new Error(`Upload URL host is not an expected Microsoft host: ${host}`);
  }
  return parsed;
}

module.exports = {
  UPLOAD_HOST_SUFFIXES,
  uploadRoots,
  resolveUploadPath,
  checkUploadUrl,
};
