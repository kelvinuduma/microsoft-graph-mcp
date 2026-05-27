# Microsoft Graph MCP Server

This is a Model Context Protocol (MCP) server that exposes **21 tools** allowing Claude Desktop (or any other MCP-compatible host) to interact securely with your Microsoft Graph API (Outlook Mail, Calendar, OneDrive, Teams, SharePoint, and Power BI).

## Prerequisites

1. **Azure/Entra ID App Registration**:
   - Go to [Azure Portal](https://portal.azure.com/) -> **Microsoft Entra ID** -> **App registrations** -> **New registration**.
   - Set Name to `Microsoft Graph MCP Server`.
   - Select **Supported account types** (e.g., *Accounts in any organizational directory and personal Microsoft accounts* for multi-tenant access).
   - Under **Authentication** -> **Advanced settings**, set **Allow public client flows** to **Yes** (required if using Device Code Flow).
   - In **API permissions**, add delegated permissions for the services you wish to use:
     - `User.Read`
     - `Mail.ReadWrite`
     - `Mail.Send`
     - `Calendars.ReadWrite`
     - `Files.ReadWrite`
     - `Group.ReadWrite.All`
   - Click **Grant admin consent** (if you have tenant administrator access).

2. **Retrieve App Identifiers**:
   - Copy the **Application (client) ID**.
   - Copy the **Directory (tenant) ID** (or use `common`).

---

## Configuration

1. Create a `.env` file in this directory based on the `.env.example`:
   ```bash
   MICROSOFT_CLIENT_ID=your-client-id-here
   MICROSOFT_TENANT_ID=common
   ```

2. Install dependencies:
   ```bash
   npm install
   ```

3. Build the server:
   ```bash
   npm run build
   ```

---

## Local Testing & Verification

1. Start the compiled server:
   ```bash
   npm run start
   ```

2. If using **Device Code Flow** (no client secret provided), the console (stderr) will output an authentication link:
   ```text
   =========================================
   MICROSOFT GRAPH AUTHENTICATION REQUIRED
   To sign in, use a web browser to open the page https://microsoft.com/devicelogin and enter the code XXXXXXXXX to authenticate.
   =========================================
   ```
3. Open the link, enter the code, sign in, and accept the app permissions. Once completed, the server will connect.

---

## Claude Desktop Configuration

To load this server in Claude Desktop, add the server to your `claude_desktop_config.json` (typically located at `%APPDATA%\Claude\claude_desktop_config.json` on Windows):

```json
{
  "mcpServers": {
    "microsoft-graph": {
      "command": "node",
      "args": ["C:/Users/kelvi/.gemini/antigravity/scratch/microsoft-graph-mcp/build/index.js"],
      "env": {
        "MICROSOFT_CLIENT_ID": "your-azure-client-id",
        "MICROSOFT_TENANT_ID": "common"
      }
    }
  }
}
```
Replace the Client ID and absolute paths with your own. Restart Claude Desktop after updating the configuration.
