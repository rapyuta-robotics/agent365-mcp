/**
 * Owner-only storage for tokens, the MSAL cache and related state.
 *
 * The MSAL cache holds a refresh token that lasts about 90 days, so other
 * local users must not be able to read it.
 */

const fs = require("fs");
const path = require("path");

const HOME = process.env.HOME || process.env.USERPROFILE || "";
const STATE_DIR = path.join(HOME, ".agent365-mcp");

const DIR_MODE = 0o700;
const FILE_MODE = 0o600;

function ensurePrivateDir(dir) {
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true, mode: DIR_MODE });
  }
}

/**
 * Write a file readable only by its owner. The data goes to a temporary file
 * first and is renamed into place, so the mode applies even when the file
 * already existed with wider permissions, and readers never see a partial file.
 */
function writePrivateFile(file, data) {
  ensurePrivateDir(path.dirname(file));
  const tmp = `${file}.${process.pid}.${Date.now()}.tmp`;
  try {
    fs.writeFileSync(tmp, data, { mode: FILE_MODE });
    fs.chmodSync(tmp, FILE_MODE);
    fs.renameSync(tmp, file);
  } catch (e) {
    try { fs.unlinkSync(tmp); } catch (_) { /* already gone */ }
    throw e;
  }
}

/**
 * Restrict the state directory and the files in it to the owner. Fixes files
 * written by earlier versions with the default mode.
 */
function tightenStateDir(dir = STATE_DIR) {
  if (process.platform === "win32") return;
  try {
    if (!fs.existsSync(dir)) return;
    fs.chmodSync(dir, DIR_MODE);
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
      if (entry.isFile()) fs.chmodSync(path.join(dir, entry.name), FILE_MODE);
    }
  } catch (e) {
    console.error(`Could not restrict permissions on ${dir}: ${e.message}`);
  }
}

module.exports = { STATE_DIR, ensurePrivateDir, writePrivateFile, tightenStateDir };
