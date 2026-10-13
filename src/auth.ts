import * as msal from "@azure/msal-node";
import {
  PersistenceCreator,
  PersistenceCachePlugin,
  DataProtectionScope,
} from "@azure/msal-node-extensions";
import * as path from "path";
import * as os from "os";
import "isomorphic-fetch";

/**
 * Delegated Microsoft Graph scopes — aligned to the implemented tools only
 * (least privilege). `Group.ReadWrite.All` was removed (no tool uses it);
 * `Sites.Read.All` and `Chat.ReadWrite` were added for the SharePoint and
 * Teams tools, which previously 403'd for lack of consent.
 */
const GRAPH_SCOPES = [
  "User.Read",
  "Mail.ReadWrite",
  "Mail.Send",
  "Calendars.ReadWrite",
  "Files.ReadWrite",
  "Sites.Read.All",
  "Chat.ReadWrite",
];

/**
 * Persistent token-cache location, OUTSIDE the repository (in the user profile,
 * so it is never committed). On Windows the file is encrypted at rest with
 * DPAPI, bound to the current Windows user.
 */
const CACHE_PATH = path.join(
  os.homedir(),
  ".microsoft-graph-mcp",
  "msal-cache.json",
);

export class MSGraphAuth {
  private clientApp: msal.PublicClientApplication | null = null;
  private account: msal.AccountInfo | null = null;
  private scopes: string[] = GRAPH_SCOPES;

  /**
   * Initializes the MSAL public client with a DPAPI-encrypted persistent cache.
   *
   * Device Code / delegated flow only. The original confidential
   * (client-credentials) path was removed: every tool targets `/me`, which is
   * invalid in app-only mode, so that path could never have worked.
   */
  public async initialize(): Promise<void> {
    const tenantId = process.env.MICROSOFT_TENANT_ID || "common";
    const clientId = process.env.MICROSOFT_CLIENT_ID;

    if (!clientId) {
      throw new Error("MICROSOFT_CLIENT_ID environment variable is missing.");
    }

    const authority = `https://login.microsoftonline.com/${tenantId}`;

    // Secure token persistence: DPAPI (Windows), Keychain (macOS), LibSecret (Linux).
    // Falls back gracefully on Linux systems without a desktop keyring (e.g. WSL/containers).
    let persistence;
    try {
      persistence = await PersistenceCreator.createPersistence({
        cachePath: CACHE_PATH,
        dataProtectionScope: DataProtectionScope.CurrentUser,
        serviceName: "microsoft-graph-mcp",
        accountName: "msal-token-cache",
        usePlaintextFileOnLinux: process.env.MSAL_USE_PLAINTEXT_CACHE === "true",
      });
    } catch (err) {
      if (process.platform === "linux") {
        console.warn(
          "[Auth Warning] System keyring not accessible. Falling back to plaintext cache on Linux. (Set MSAL_USE_PLAINTEXT_CACHE=true to suppress this check)."
        );
        persistence = await PersistenceCreator.createPersistence({
          cachePath: CACHE_PATH,
          dataProtectionScope: DataProtectionScope.CurrentUser,
          serviceName: "microsoft-graph-mcp",
          accountName: "msal-token-cache",
          usePlaintextFileOnLinux: true,
        });
      } else {
        throw err;
      }
    }

    const config: msal.Configuration = {
      auth: { clientId, authority },
      cache: {
        cachePlugin: new PersistenceCachePlugin(persistence),
      },
    };

    this.clientApp = new msal.PublicClientApplication(config);
    console.error(
      "[Auth] Public client ready (device-code flow, DPAPI-encrypted persistent cache)",
    );
  }

  /**
   * Hydrates `this.account` from the persistent cache — needed after a process
   * restart, when the in-memory field is null but the cache file still holds a
   * valid account. Provided for you; call it before attempting a silent token.
   */
  private async loadCachedAccount(): Promise<void> {
    if (!this.clientApp) return;
    const accounts = await this.clientApp.getTokenCache().getAllAccounts();
    if (accounts.length > 0) {
      this.account = accounts[0];
    }
  }

  /**
   * Device-code callback: prints the sign-in instructions to STDERR so they do
   * not corrupt the MCP stdio (stdout) JSON-RPC channel.
   */
  private deviceCodeCallback(response: { message: string }): void {
    console.error("\n=========================================");
    console.error("MICROSOFT GRAPH AUTHENTICATION REQUIRED");
    console.error(response.message);
    console.error("=========================================\n");
  }

  public async getAccessToken(): Promise<string> {
    if (!this.clientApp) {
      throw new Error("MSGraphAuth not initialized. Call initialize() first.");
    }

    // 1. If account is null, hydrate it from the persistent cache
    if (!this.account) {
      await this.loadCachedAccount();
    }

    // 2. If we now have an account, try silent acquisition first
    if (this.account) {
      try {
        const silentRequest = {
          account: this.account,
          scopes: this.scopes,
        };
        const result = await this.clientApp.acquireTokenSilent(silentRequest);
        if (result && result.accessToken) {
          return result.accessToken;
        }
      } catch (error: any) {
        console.error(
          `[Auth] Silent token acquisition failed: ${error.message}. Falling back to Device Code Flow.`,
        );
      }
    }

    // 3. Fall back to Device Code Flow
    const deviceCodeRequest = {
      deviceCodeCallback: (response: any) => this.deviceCodeCallback(response),
      scopes: this.scopes,
    };

    const result = await this.clientApp.acquireTokenByDeviceCode(deviceCodeRequest);
    if (!result || !result.accessToken) {
      throw new Error("Failed to acquire access token via device code.");
    }

    this.account = result.account;
    return result.accessToken;
  }
}
