import * as msal from "@azure/msal-node";
import {
  PersistenceCreator,
  PersistenceCachePlugin,
  DataProtectionScope,
} from "@azure/msal-node-extensions";
import * as path from "path";
import * as os from "os";
import * as fs from "fs";
import { execFile } from "child_process";
import { Config } from "./config.js";
import { scopesFor } from "./scopes.js";

function profileSuffix(cfg: Config): string {
  return cfg.isCorporateProfile ? "" : `-${cfg.profile}`;
}

export function getMsalCachePath(cfg: Config): string {
  if (process.env.MICROSOFT_CACHE_PATH) {
    return path.resolve(process.env.MICROSOFT_CACHE_PATH);
  }
  return path.join(
    os.homedir(),
    `.microsoft-graph-mcp${profileSuffix(cfg)}`,
    "msal-cache.json",
  );
}

/** Raised when the user has to sign in; the message tells them how. */
export class AuthRequiredError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "AuthRequiredError";
  }
}

type DeviceCodeInfo = {
  message: string;
  userCode?: string;
  verificationUri?: string;
};

/**
 * Best-effort convenience: copy the device code and open the sign-in page.
 * Uses execFile with an argument array (no shell string interpolation).
 */
function assistDeviceLogin(info: DeviceCodeInfo): void {
  if (/^(1|true|yes)$/i.test(process.env.GRAPH_MCP_NO_BROWSER || "")) return;
  const url = info.verificationUri || "https://microsoft.com/devicelogin";
  const quiet = () => undefined;
  try {
    if (process.platform === "win32") {
      if (info.userCode) {
        // stdin, not a command argument: nothing in the code is ever parsed.
        const p = execFile("clip.exe", [], quiet);
        p.stdin?.end(info.userCode);
      }
      execFile("rundll32.exe", ["url.dll,FileProtocolHandler", url], quiet);
    } else if (process.platform === "darwin") {
      if (info.userCode) {
        const p = execFile("pbcopy", [], quiet);
        p.stdin?.end(info.userCode);
      }
      execFile("open", [url], quiet);
    } else {
      execFile("xdg-open", [url], quiet);
    }
  } catch {
    // The printed message is enough on its own.
  }
}

export class MSGraphAuth {
  private clientApp: msal.PublicClientApplication | null = null;
  private account: msal.AccountInfo | null = null;
  private scopes: string[];
  private pendingDeviceFlow: Promise<void> | null = null;
  private pendingMessage = "";

  constructor(private readonly cfg: Config) {
    this.scopes = scopesFor(cfg);
  }

  public get grantedScopes(): string[] {
    return [...this.scopes];
  }

  public get username(): string | null {
    return this.account?.username ?? null;
  }

  /**
   * Forgets the cached sign-in. On macOS and Linux the tokens live in the
   * Keychain / libsecret, not the file, so the persistence entry itself must
   * be deleted; unlinking the file alone would leave the refresh token behind.
   */
  public async clearCache(): Promise<void> {
    const cachePath = getMsalCachePath(this.cfg);
    try {
      const persistence = await this.createPersistence();
      await persistence.delete();
    } catch (e: any) {
      console.error(
        `[Auth] Could not clear the secure store entry: ${e.message}`,
      );
    }
    try {
      if (fs.existsSync(cachePath)) fs.unlinkSync(cachePath);
    } catch (e: any) {
      console.error(`[Auth] Could not delete token cache file: ${e.message}`);
    }
    this.account = null;
    console.error(
      `[Auth] Cleared token cache for profile '${this.cfg.profile}'.`,
    );
  }

  private async createPersistence() {
    const suffix = profileSuffix(this.cfg);
    return PersistenceCreator.createPersistence({
      cachePath: getMsalCachePath(this.cfg),
      dataProtectionScope: DataProtectionScope.CurrentUser,
      serviceName: `microsoft-graph-mcp${suffix}`,
      accountName: `msal-token-cache${suffix}`,
      usePlaintextFileOnLinux: process.env.MSAL_USE_PLAINTEXT_CACHE === "true",
    });
  }

  /** Builds the MSAL public client with an encrypted persistent token cache. */
  public async initialize(): Promise<void> {
    if (!this.cfg.clientId) {
      throw new Error("MICROSOFT_CLIENT_ID is not set.");
    }
    let persistence;
    try {
      persistence = await this.createPersistence();
    } catch (err: any) {
      console.error(
        `[Auth] Token cache unavailable (${err.message}); resetting it.`,
      );
      await this.clearCache();
      persistence = await this.createPersistence();
    }
    this.clientApp = new msal.PublicClientApplication({
      auth: {
        clientId: this.cfg.clientId,
        authority: `https://login.microsoftonline.com/${encodeURIComponent(this.cfg.tenantId)}`,
      },
      cache: { cachePlugin: new PersistenceCachePlugin(persistence) },
    });
    console.error(
      `[Auth] Ready (profile: ${this.cfg.profile}, tenant: ${this.cfg.isConsumer ? "consumers" : "organisational"}, outbound: ${this.cfg.enableOutbound ? "on" : "off"})`,
    );
  }

  private async loadCachedAccount(): Promise<void> {
    if (!this.clientApp) return;
    try {
      const accounts = await this.clientApp.getTokenCache().getAllAccounts();
      this.account = accounts[0] ?? null;
    } catch (error: any) {
      console.error(
        `[Auth] Token cache unreadable (${error.message}); clearing it.`,
      );
      await this.clearCache();
    }
  }

  private async trySilent(): Promise<string | null> {
    if (!this.clientApp) throw new Error("MSGraphAuth not initialised.");
    if (!this.account) await this.loadCachedAccount();
    if (!this.account) return null;
    try {
      const result = await this.clientApp.acquireTokenSilent({
        account: this.account,
        scopes: this.scopes,
      });
      return result?.accessToken || null;
    } catch (error: any) {
      const msg = String(error?.message || error);
      if (
        msg.includes("Encryption/Decryption failed") ||
        msg.includes("2148073483")
      ) {
        await this.clearCache();
      } else if (!(error instanceof msal.InteractionRequiredAuthError)) {
        console.error(`[Auth] Silent token refresh failed: ${msg}`);
      }
      return null;
    }
  }

  private startDeviceFlow(
    onCode: (info: DeviceCodeInfo) => void,
  ): Promise<msal.AuthenticationResult | null> {
    if (!this.clientApp) throw new Error("MSGraphAuth not initialised.");
    return this.clientApp.acquireTokenByDeviceCode({
      scopes: this.scopes,
      deviceCodeCallback: (info) => {
        this.pendingMessage = info.message;
        console.error(`\n[Auth] ${info.message}\n`);
        assistDeviceLogin(info);
        onCode(info);
      },
    });
  }

  /**
   * Interactive sign-in for the CLI: blocks until the user completes the
   * device-code flow. Retries transient network failures.
   */
  public async signInInteractive(): Promise<string> {
    const silent = await this.trySilent();
    if (silent) return silent;
    for (let attempt = 1; ; attempt++) {
      try {
        const result = await this.startDeviceFlow(() => undefined);
        if (!result?.accessToken) throw new Error("No access token returned.");
        this.account = result.account;
        return result.accessToken;
      } catch (err: any) {
        const msg = String(err?.message || err);
        const transient =
          msg.includes("network_error") ||
          msg.includes("Network request failed");
        if (!transient || attempt >= 3) throw err;
        console.error(`[Auth] Network error (${msg}); retrying in 3 s...`);
        await new Promise((r) => setTimeout(r, 3000));
      }
    }
  }

  /**
   * Token for tool calls. Never blocks the MCP channel on a sign-in: if no
   * cached token works, it starts a device-code flow in the background and
   * throws AuthRequiredError carrying the sign-in instructions. Once the user
   * finishes signing in, the next call succeeds.
   */
  public async getAccessToken(): Promise<string> {
    const silent = await this.trySilent();
    if (silent) return silent;

    if (!this.pendingDeviceFlow) {
      let announce: () => void = () => undefined;
      const announced = new Promise<void>((resolve) => (announce = resolve));
      this.pendingDeviceFlow = this.startDeviceFlow(() => announce())
        .then((result) => {
          if (result?.account) this.account = result.account;
          console.error("[Auth] Sign-in completed.");
        })
        .catch((err) => {
          console.error(
            `[Auth] Sign-in did not complete: ${err?.message || err}`,
          );
        })
        .finally(() => {
          this.pendingDeviceFlow = null;
          this.pendingMessage = "";
        });
      await Promise.race([
        announced,
        new Promise((r) => setTimeout(r, 10_000)),
      ]);
    }

    throw new AuthRequiredError(
      this.pendingMessage
        ? `Microsoft sign-in required. ${this.pendingMessage} Then retry this request.`
        : "Microsoft sign-in required, but no device code was issued. Run `npm run auth` in the server folder.",
    );
  }
}
