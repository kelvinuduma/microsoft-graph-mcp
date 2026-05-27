import * as msal from "@azure/msal-node";
import "isomorphic-fetch";

export class MSGraphAuth {
  private clientAppConfidential: msal.ConfidentialClientApplication | null = null;
  private clientAppPublic: msal.PublicClientApplication | null = null;
  private scopes: string[];

  constructor() {
    this.scopes = [
      "https://graph.microsoft.com/.default"
    ];
  }

  /**
   * Initializes the MSAL application configuration based on environment variables.
   */
  public initialize(): void {
    const tenantId = process.env.MICROSOFT_TENANT_ID || "common";
    const clientId = process.env.MICROSOFT_CLIENT_ID;
    const clientSecret = process.env.MICROSOFT_CLIENT_SECRET;

    if (!clientId) {
      throw new Error("MICROSOFT_CLIENT_ID environment variable is missing.");
    }

    const authority = `https://login.microsoftonline.com/${tenantId}`;

    if (clientSecret) {
      // Initialize Confidential Client Application (App-only permissions)
      const config: msal.Configuration = {
        auth: {
          clientId,
          authority,
          clientSecret,
        }
      };
      this.clientAppConfidential = new msal.ConfidentialClientApplication(config);
      console.error("[Auth] Initialized Confidential Client (Client Credentials Flow)");
    } else {
      // Initialize Public Client Application (Interactive / Device Code Flow for user delegated permissions)
      const config: msal.Configuration = {
        auth: {
          clientId,
          authority,
        }
      };
      this.clientAppPublic = new msal.PublicClientApplication(config);
      console.error("[Auth] Initialized Public Client (Device Code Flow)");
    }
  }

  /**
   * Acquires a fresh access token for Microsoft Graph.
   */
  public async getAccessToken(): Promise<string> {
    if (this.clientAppConfidential) {
      // Client Credentials flow
      const tokenRequest: msal.ClientCredentialRequest = {
        scopes: this.scopes,
      };
      const response = await this.clientAppConfidential.acquireTokenByClientCredential(tokenRequest);
      if (!response || !response.accessToken) {
        throw new Error("Failed to acquire access token via client credentials.");
      }
      return response.accessToken;
    } else if (this.clientAppPublic) {
      // Device Code Flow
      const deviceCodeRequest: msal.DeviceCodeRequest = {
        deviceCodeCallback: (response) => {
          // Print instructions to stderr so they don't corrupt the MCP stdio communication protocol channel!
          console.error("\n=========================================");
          console.error("MICROSOFT GRAPH AUTHENTICATION REQUIRED");
          console.error(response.message);
          console.error("=========================================\n");
        },
        scopes: [
          "User.Read",
          "Mail.ReadWrite",
          "Mail.Send",
          "Calendars.ReadWrite",
          "Files.ReadWrite",
          "Group.ReadWrite.All"
        ],
      };

      const response = await this.clientAppPublic.acquireTokenByDeviceCode(deviceCodeRequest);
      if (!response || !response.accessToken) {
        throw new Error("Failed to acquire access token via device code.");
      }
      return response.accessToken;
    } else {
      throw new Error("MSGraphAuth has not been initialized. Call initialize() first.");
    }
  }
}
