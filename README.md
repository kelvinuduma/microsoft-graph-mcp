# Microsoft Graph MCP Server

A [Model Context Protocol](https://modelcontextprotocol.io) (MCP) server that exposes **20 tools** letting Claude Desktop (or any MCP-compatible host) work with your Microsoft 365 data through the Microsoft Graph API — **Outlook Mail, Calendar, OneDrive, Teams, and SharePoint**.

Authentication uses the MSAL **device-code (delegated) flow** with a **persistent, DPAPI-encrypted token cache**, so you sign in once and the server reuses the cached session across restarts.

## Tools (20)

| Area             | Tools                                                                           |
| ---------------- | ------------------------------------------------------------------------------- |
| **Outlook Mail** | list messages, get message, send mail, create draft, search messages            |
| **Calendar**     | list events, get event, create event, update event, delete event                |
| **OneDrive**     | list root, list folder, search files, download file, upload file, create folder |
| **SharePoint**   | list sites, get site details                                                    |
| **Teams**        | list chats, send chat message                                                   |

## Prerequisites

### Entra ID (Azure AD) app registration

1. [Azure Portal](https://portal.azure.com/) → **Microsoft Entra ID** → **App registrations** → **New registration**.
2. **Name:** `Microsoft Graph MCP Server`.
3. **Supported account types:** _Accounts in this organizational directory only_ (single-tenant) is recommended for the tightest sign-in surface. Note your **Directory (tenant) ID**.
4. **Authentication** → **Advanced settings** → set **Allow public client flows** to **Yes** (required for the device-code flow).
5. **API permissions** → add these **delegated** Microsoft Graph permissions (the least-privilege set matching the 20 tools):
   - `User.Read`
   - `Mail.ReadWrite`
   - `Mail.Send`
   - `Calendars.ReadWrite`
   - `Files.ReadWrite`
   - `Sites.Read.All`
   - `Chat.ReadWrite`
6. Click **Grant admin consent** — `Sites.Read.All` and `Chat.ReadWrite` require it.
7. Copy the **Application (client) ID** and **Directory (tenant) ID**.

## Configuration

Create a `.env` from `.env.example`:

```bash
MICROSOFT_CLIENT_ID=your-application-client-id
MICROSOFT_TENANT_ID=your-directory-tenant-id   # your tenant ID (single-tenant) or "common" (multi-tenant/personal)
```

Install and build:

```bash
npm install
npm run build
```

## First sign-in (run once in a terminal)

Start the server directly **once** in a terminal so you can see and complete the device-code prompt:

```bash
npm run start
```

The sign-in instructions print to **stderr**:

```text
=========================================
MICROSOFT GRAPH AUTHENTICATION REQUIRED
To sign in, use a web browser to open https://microsoft.com/devicelogin and enter the code XXXXXXXXX.
=========================================
```

Open the link, enter the code, sign in, and consent. The token is then cached (DPAPI-encrypted, under `%USERPROFILE%\.microsoft-graph-mcp\`), so later runs — including when Claude Desktop launches the server — authenticate **silently**.

> **Why a terminal first?** Claude Desktop captures the server's stderr into log files, so the device-code prompt is not visible in the chat UI. Completing the first sign-in in a terminal seeds the persistent cache; Claude Desktop then reuses it silently.

## Claude Desktop configuration

Add the server to `claude_desktop_config.json` (on Windows: `%APPDATA%\Claude\claude_desktop_config.json`). Use the **absolute path to your built `build/index.js`**:

```json
{
  "mcpServers": {
    "microsoft-graph": {
      "command": "node",
      "args": ["C:/path/to/microsoft-graph-mcp/build/index.js"],
      "env": {
        "MICROSOFT_CLIENT_ID": "your-application-client-id",
        "MICROSOFT_TENANT_ID": "your-directory-tenant-id"
      }
    }
  }
}
```

Restart Claude Desktop after editing the config.

## Security notes

- **Delegated permissions only.** The server acts as the signed-in user and requests the least-privilege scope set above. No client secret is used or stored.
- **Token cache** is DPAPI-encrypted at rest (Windows) and lives in your user profile, never in the repository.
- Your `.env` is gitignored — never commit credentials.
