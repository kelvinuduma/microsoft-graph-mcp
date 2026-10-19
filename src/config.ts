import * as dotenv from "dotenv";
import * as os from "os";
import * as path from "path";

/**
 * Loads configuration once, in this order (earlier wins, real environment
 * variables always win):
 *   1. the process environment (what the MCP client passes in its `env` block)
 *   2. `.env.<profile>` next to the package (profile-specific, git-ignored)
 *   3. `.env` next to the package (shared, git-ignored)
 *   4. `.env` in the current working directory
 *
 * `quiet: true` matters: dotenv >= 17 logs to stdout by default, which would
 * corrupt the MCP stdio JSON-RPC channel.
 */
const PACKAGE_ROOT = path.resolve(__dirname, "..");

function loadEnvFiles(): void {
  if (process.env.MICROSOFT_ENV_FILE) {
    dotenv.config({
      path: path.resolve(process.env.MICROSOFT_ENV_FILE),
      quiet: true,
    });
  }
  const profile = normaliseProfile(process.env.MICROSOFT_PROFILE);
  dotenv.config({
    path: path.join(PACKAGE_ROOT, `.env.${profile}`),
    quiet: true,
  });
  dotenv.config({ path: path.join(PACKAGE_ROOT, ".env"), quiet: true });
  dotenv.config({ quiet: true });
}

function normaliseProfile(raw: string | undefined): string {
  const p = (raw || "").trim().toLowerCase();
  return !p || p === "default" ? "corporate" : p;
}

function flag(name: string, fallback = false): boolean {
  const v = process.env[name];
  if (v === undefined || v === "") return fallback;
  return ["1", "true", "yes", "on"].includes(v.trim().toLowerCase());
}

function list(name: string): string[] {
  return (process.env[name] || "")
    .split(path.delimiter)
    .map((s) => s.trim())
    .filter(Boolean);
}

function positiveInt(name: string, fallback: number): number {
  const n = Number.parseInt(process.env[name] || "", 10);
  return Number.isFinite(n) && n > 0 ? n : fallback;
}

export interface Config {
  clientId: string;
  tenantId: string;
  /** "corporate" (default) or any other name, e.g. "personal". */
  profile: string;
  isCorporateProfile: boolean;
  isConsumer: boolean;
  toolPrefix: string;
  /** Tools that notify other people: send mail, Teams messages, meeting changes. */
  enableOutbound: boolean;
  /** Local roots the server may write downloads into. */
  writeDirs: string[];
  /** Local roots the server may read attachments from. */
  readDirs: string[];
  /** Extra path fragments to refuse for both reads and writes. */
  denyPaths: string[];
  allowExecutableWrites: boolean;
  maxTransferBytes: number;
  /** IANA or Windows time-zone name used for calendar input and output. */
  timeZone: string;
}

let cached: Config | null = null;

export function loadConfig(): Config {
  if (cached) return cached;
  loadEnvFiles();

  const tenantId = process.env.MICROSOFT_TENANT_ID || "common";
  const profile = normaliseProfile(process.env.MICROSOFT_PROFILE);
  const home = os.homedir();

  cached = {
    clientId: process.env.MICROSOFT_CLIENT_ID || "",
    tenantId,
    profile,
    isCorporateProfile: profile === "corporate",
    isConsumer: tenantId.toLowerCase() === "consumers",
    toolPrefix: process.env.TOOL_PREFIX || "",
    enableOutbound: flag("GRAPH_MCP_ENABLE_OUTBOUND"),
    // Narrow defaults: a home-wide write root cannot be made safe with a
    // denylist (e.g. overwriting ~/.claude.json or ~/.gitconfig runs code).
    writeDirs: list("GRAPH_MCP_WRITE_DIRS").length
      ? list("GRAPH_MCP_WRITE_DIRS")
      : [path.join(home, "Downloads", "graph-mcp")],
    readDirs: list("GRAPH_MCP_READ_DIRS").length
      ? list("GRAPH_MCP_READ_DIRS")
      : ["Documents", "Downloads", "Desktop"].map((d) => path.join(home, d)),
    denyPaths: list("GRAPH_MCP_DENY_PATHS"),
    allowExecutableWrites: flag("GRAPH_MCP_ALLOW_EXECUTABLE_WRITES"),
    maxTransferBytes:
      positiveInt("GRAPH_MCP_MAX_TRANSFER_MB", 150) * 1024 * 1024,
    timeZone: process.env.GRAPH_MCP_TIMEZONE || "UTC",
  };
  return cached;
}

/** Test hook: forget the cached config so a test can change the environment. */
export function resetConfigForTests(): void {
  cached = null;
}
