import * as fs from "fs";
import * as path from "path";
import { Config } from "./config.js";

/**
 * Local-filesystem and request-shaping guardrails.
 *
 * The model driving this server can be steered by the content it reads
 * (an email, a shared file, a chat message). These checks stop a prompt
 * injection from turning a "download" into "overwrite my startup folder", or an
 * "attach" into "mail out my SSH key".
 */

const CASE_INSENSITIVE_FS =
  process.platform === "win32" || process.platform === "darwin";

/** Path fragments that are never read or written, whatever the config says. */
const SENSITIVE_SEGMENTS = [
  ".ssh",
  ".gnupg",
  ".aws",
  ".azure",
  ".kube",
  ".docker",
  ".config",
  ".claude",
  ".claude-jcl",
  ".codex",
  ".gemini",
  ".cursor",
  ".copilot",
  ".vscode",
  ".git",
  ".ipython",
  ".jupyter",
  ".local",
  ".npm",
  ".nvm",
  ".cargo",
  ".m2",
  ".gradle",
  "appdata",
  "library/keychains",
  "start menu/programs/startup",
  "documents/powershell",
  "documents/windowspowershell",
];

/** Directory names that hold this server's own token caches. */
const TOKEN_CACHE_DIR = /^\.microsoft-graph-mcp(-[\w-]+)?$/i;

const SENSITIVE_FILE_PATTERNS = [
  /^\.env(\..*)?$/i,
  /^id_(rsa|dsa|ecdsa|ed25519)(\.pub)?$/i,
  /\.(pem|key|pfx|p12|kdbx|ovpn|ppk)$/i,
  /^(\.npmrc|\.netrc|\.git-credentials|\.pgpass|credentials(\.json)?)$/i,
  /^msal-cache\.json$/i,
  /^rt\.dat$/i,
  /^\.(bashrc|bash_profile|bash_logout|zshrc|zprofile|profile|inputrc|vimrc|rprofile)$/i,
  // Tool and agent configuration that is executed or obeyed on next start.
  /^(\.claude\.json|\.gitconfig|\.wslconfig|\.pypirc|_netrc|\.condarc|\.yarnrc(\.yml)?|desktop\.ini|autorun\.inf)$/i,
  /^(agents|claude|gemini|soul|user|handoff|memory|identity)\.md$/i,
  /^(\.cursorrules|\.windsurfrules|copilot-instructions\.md)$/i,
];

/** Names Windows treats as devices, with or without an extension. */
const RESERVED_DEVICE_NAME =
  /^(con|prn|aux|nul|com[0-9¹²³]|lpt[0-9¹²³])(\..*)?$/i;

/**
 * Rejects path forms whose meaning differs from how they look, before any
 * other check: NTFS alternate data streams (`evil.bat::$DATA` writes
 * `evil.bat`), device names, and components ending in a dot or space (which
 * Windows strips, so `evil.bat.` becomes `evil.bat`).
 */
function assertPlainPath(target: string, name: string): void {
  const rest = target.slice(path.parse(target).root.length);
  if (rest.includes(":")) {
    throw new PolicyError(
      `${name} must not contain ':' (alternate data streams are not allowed).`,
    );
  }
  for (const part of rest.split(/[\\/]+/).filter(Boolean)) {
    if (RESERVED_DEVICE_NAME.test(part)) {
      throw new PolicyError(`${name} uses a reserved device name ('${part}').`);
    }
    if (/[. ]$/.test(part) && part !== "." && part !== "..") {
      throw new PolicyError(
        `${name} has a component ending in a dot or space ('${part}').`,
      );
    }
  }
}

const EXECUTABLE_EXTENSIONS =
  /\.(exe|com|bat|cmd|ps1|psm1|psd1|vbs|vbe|js|jse|wsf|wsh|msi|msp|scr|lnk|dll|cpl|hta|reg|sh|command|app|jar)$/i;

function norm(p: string): string {
  const slashed = p.replace(/\\/g, "/");
  return CASE_INSENSITIVE_FS ? slashed.toLowerCase() : slashed;
}

function isWithin(child: string, root: string): boolean {
  const rel = path.relative(root, child);
  return rel === "" || (!rel.startsWith("..") && !path.isAbsolute(rel));
}

/** Resolves symlinks on the longest existing prefix of `p`. */
function realpathLoose(p: string): string {
  let current = p;
  const tail: string[] = [];
  while (!fs.existsSync(current)) {
    const parent = path.dirname(current);
    if (parent === current) return p;
    tail.unshift(path.basename(current));
    current = parent;
  }
  return path.join(fs.realpathSync.native(current), ...tail);
}

/**
 * `relToRoot` is the path below the allowed root it matched. Protected folder
 * names are checked there, not on the full path, so an explicitly allowed root
 * such as the Windows temp folder (which sits under AppData) still works.
 */
function assertNotSensitive(
  resolved: string,
  relToRoot: string,
  cfg: Config,
  verb: string,
): void {
  // Case-insensitive on every platform: stricter is safer, and the list is lowercase.
  const rel = `/${relToRoot}`.toLowerCase();

  for (const seg of SENSITIVE_SEGMENTS) {
    if (rel.includes(`/${seg}/`) || rel.endsWith(`/${seg}`)) {
      throw new PolicyError(
        `Refusing to ${verb} '${resolved}': it is inside a protected location (${seg}).`,
      );
    }
  }
  if (
    norm(resolved)
      .split("/")
      .some((s) => TOKEN_CACHE_DIR.test(s))
  ) {
    throw new PolicyError(
      `Refusing to ${verb} '${resolved}': it is inside this server's token cache.`,
    );
  }
  const base = path.basename(resolved);
  if (SENSITIVE_FILE_PATTERNS.some((re) => re.test(base))) {
    throw new PolicyError(
      `Refusing to ${verb} '${resolved}': the file name looks like a credential, shell profile or agent/tool config file.`,
    );
  }
  for (const fragment of cfg.denyPaths) {
    if (norm(resolved).includes(norm(fragment))) {
      throw new PolicyError(
        `Refusing to ${verb} '${resolved}': it matches a configured deny rule (GRAPH_MCP_DENY_PATHS).`,
      );
    }
  }
}

/** Returns the path relative to the most specific allowed root, or throws. */
function assertInsideRoots(
  resolved: string,
  roots: string[],
  envName: string,
  verb: string,
): string {
  const target = norm(resolved);
  let best: string | null = null;
  for (const r of roots) {
    const root = norm(realpathLoose(path.resolve(r)));
    if (
      isWithin(target, root) &&
      (best === null || root.length > best.length)
    ) {
      best = root;
    }
  }
  if (best === null) {
    throw new PolicyError(
      `Refusing to ${verb} '${resolved}': it is outside the allowed folders. Allowed: ${roots.join(", ")} (set ${envName} to change).`,
    );
  }
  return norm(path.relative(best, target));
}

export class PolicyError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "PolicyError";
  }
}

/**
 * Validates a local destination for a download. Returns the absolute path.
 * Refuses relative paths, protected locations, executables (unless allowed),
 * and overwriting an existing file unless `overwrite` is true.
 */
export function resolveWritePath(
  target: unknown,
  cfg: Config,
  overwrite = false,
): string {
  if (typeof target !== "string" || !target.trim()) {
    throw new PolicyError("destinationPath must be a non-empty string.");
  }
  if (!path.isAbsolute(target)) {
    throw new PolicyError(`destinationPath must be absolute, got '${target}'.`);
  }
  assertPlainPath(target, "destinationPath");
  const resolved = realpathLoose(path.resolve(target));
  const rel = assertInsideRoots(
    resolved,
    cfg.writeDirs,
    "GRAPH_MCP_WRITE_DIRS",
    "write",
  );
  assertNotSensitive(resolved, rel, cfg, "write");

  if (!cfg.allowExecutableWrites && EXECUTABLE_EXTENSIONS.test(resolved)) {
    throw new PolicyError(
      `Refusing to write '${resolved}': executable or script file types are blocked (set GRAPH_MCP_ALLOW_EXECUTABLE_WRITES=true to allow).`,
    );
  }
  if (fs.existsSync(resolved)) {
    if (fs.statSync(resolved).isDirectory()) {
      throw new PolicyError(
        `destinationPath '${resolved}' is a folder; include the file name.`,
      );
    }
    if (!overwrite) {
      throw new PolicyError(
        `'${resolved}' already exists. Pass overwrite: true to replace it.`,
      );
    }
  }
  return resolved;
}

/** Validates a local file that is about to be uploaded or attached. */
export function resolveReadPath(source: unknown, cfg: Config): string {
  if (typeof source !== "string" || !source.trim()) {
    throw new PolicyError("filePath must be a non-empty string.");
  }
  if (!path.isAbsolute(source)) {
    throw new PolicyError(`filePath must be absolute, got '${source}'.`);
  }
  assertPlainPath(source, "filePath");
  if (!fs.existsSync(source)) {
    throw new PolicyError(`File not found: ${source}`);
  }
  const resolved = fs.realpathSync.native(path.resolve(source));
  const rel = assertInsideRoots(
    resolved,
    cfg.readDirs,
    "GRAPH_MCP_READ_DIRS",
    "read",
  );
  assertNotSensitive(resolved, rel, cfg, "read");

  const stat = fs.statSync(resolved);
  if (!stat.isFile()) {
    throw new PolicyError(`'${resolved}' is not a regular file.`);
  }
  if (stat.size > cfg.maxTransferBytes) {
    throw new PolicyError(
      `'${resolved}' is ${stat.size} bytes, above the ${cfg.maxTransferBytes}-byte limit (GRAPH_MCP_MAX_TRANSFER_MB).`,
    );
  }
  return resolved;
}

/**
 * Encodes one URL path segment taken from model input (a message, event,
 * drive or site ID). Without this, an "ID" such as `../../users/x` could
 * redirect the request to a different Graph resource.
 */
export function seg(value: unknown, name: string): string {
  if (typeof value !== "string" || !value.trim()) {
    throw new PolicyError(`${name} must be a non-empty string.`);
  }
  if (value.length > 1024) {
    throw new PolicyError(`${name} is too long.`);
  }
  if (value === "." || value === "..") {
    throw new PolicyError(`${name} is not a valid ID.`);
  }
  return encodeURIComponent(value);
}

/**
 * SharePoint site IDs are commonly in `host,siteGuid,webGuid` form, or the
 * path form `host:/sites/name:`. Commas and colons are kept so both work;
 * everything else, including `/` and `?`, is encoded.
 */
export function siteSeg(value: unknown): string {
  return seg(value, "siteId").replace(/%2C/gi, ",").replace(/%3A/gi, ":");
}

/** Clamps a model-supplied page size into [1, max]. */
export function clampTop(
  value: unknown,
  fallback: number,
  max: number,
): number {
  const n = typeof value === "number" ? value : Number(value);
  if (!Number.isFinite(n)) return fallback;
  return Math.min(Math.max(Math.trunc(n), 1), max);
}

const GRAPH_ORIGIN = "https://graph.microsoft.com";

/** Query options a genuine next-page link may carry. */
const PAGING_PARAMS = new Set([
  "$skiptoken",
  "$skip",
  "$top",
  "$select",
  "$filter",
  "$orderby",
  "$count",
]);

/**
 * Accepts an `@odata.nextLink` only if it is a next page of the same kind of
 * resource the tool lists. The Graph client attaches the bearer token to any
 * absolute URL, so the origin must be Graph; and without the path and query
 * checks, a "next page" could be any Graph GET (for example a mail search
 * with `$expand=attachments` passed to a OneDrive tool).
 *
 * @param resource final path segment the tool lists, e.g. "messages" or "children"
 */
export function assertGraphUrl(value: unknown, resource: string): string {
  const u = parseGraphUrl(value);
  const pathname = u.pathname.toLowerCase();
  if (
    !pathname.startsWith("/v1.0/") ||
    !pathname.endsWith(`/${resource.toLowerCase()}`)
  ) {
    throw new PolicyError(
      `nextLink must be a next page of '${resource}' from the previous result.`,
    );
  }
  for (const key of u.searchParams.keys()) {
    if (!PAGING_PARAMS.has(key.toLowerCase())) {
      throw new PolicyError(`nextLink carries an unexpected option '${key}'.`);
    }
  }
  return u.toString();
}

function parseGraphUrl(value: unknown): URL {
  if (typeof value !== "string") {
    throw new PolicyError("nextLink must be a string.");
  }
  let u: URL;
  try {
    u = new URL(value);
  } catch {
    throw new PolicyError("nextLink is not a valid URL.");
  }
  if (u.origin !== GRAPH_ORIGIN || u.username || u.password) {
    throw new PolicyError(
      `nextLink must start with ${GRAPH_ORIGIN}/ (got ${u.origin}).`,
    );
  }
  return u;
}

/** Escapes a value for use inside a single-quoted OData string literal. */
export function odataString(value: unknown): string {
  return String(value ?? "").replace(/'/g, "''");
}
