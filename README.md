# Microsoft Graph MCP Server

A hardened [Model Context Protocol](https://modelcontextprotocol.io) server that gives AI agents (Claude Desktop, Claude Code, Codex, Cursor and others) delegated access to Microsoft 365: Outlook mail, Calendar, OneDrive, SharePoint and Teams.

It runs locally over stdio, signs in with the device-code flow, keeps tokens encrypted at rest, and assumes the model can be manipulated by the content it reads. The defaults are **drafts, not sends**, **safe folders, not the whole disk**, and **Microsoft endpoints only**.

## Highlights

- **30 tools** across mail, calendar, OneDrive, SharePoint, Teams and diagnostics (28 shown by default, 21 for personal accounts).
- **Outbound actions off by default.** Sending mail, posting to Teams, and creating, changing or cancelling meetings with attendees need `GRAPH_MCP_ENABLE_OUTBOUND=true`. Without it the model creates drafts you review and send, and the `Mail.Send` scope is not requested.
- **Local file guardrails.** Downloads go to `~/Downloads/graph-mcp` by default, and attachments come only from allowed folders. Alternate data streams, device names and trailing-dot tricks are rejected. Credential files, key stores, shell profiles, startup folders, agent configs and this server's own token cache are always refused. No silent overwrites. Executables are blocked unless allowed.
- **Request hardening.** Every ID from the model is URL-encoded. Pagination links must be a genuine next page of the same resource on `graph.microsoft.com`, so the token can't be sent elsewhere and a "next page" can't become a different query. Transfers have size limits and timeouts.
- **Prompt-injection notice** on every tool that returns third-party text (mail, chats, invites).
- **Non-blocking sign-in.** The MCP handshake never waits for a login. If a token is needed, the tool returns the device code to the user and the next call succeeds.
- **Actionable errors.** For example, a removed Exchange licence is reported as such, not as a bare `404`.
- **Multiple profiles.** Run a work and a personal account side by side, each with its own encrypted token cache and tool prefix.

## Tools

| Area        | Tools                                                                                                                                                                                                                                                                             |
| ----------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Diagnostics | `graph_health_check`                                                                                                                                                                                                                                                              |
| Mail        | `outlook_list_messages`, `outlook_get_message`, `outlook_search_messages`, `outlook_create_draft`, `outlook_create_reply_draft`, `outlook_list_attachments`, `outlook_download_attachment`, `outlook_add_attachment`, `outlook_file_attachment_to_onedrive`, `outlook_send_mail`¹ |
| Calendar    | `calendar_list_events`, `calendar_get_event`, `calendar_create_event`², `calendar_update_event`², `calendar_delete_event`²                                                                                                                                                        |
| OneDrive    | `onedrive_list_root`, `onedrive_list_folder`, `onedrive_search_files`, `onedrive_download_file`, `onedrive_upload_file`, `onedrive_create_folder`                                                                                                                                 |
| SharePoint³ | `sharepoint_list_sites`, `sharepoint_get_site_details`, `sharepoint_list_drive_items`, `sharepoint_download_file`, `sharepoint_search_files`                                                                                                                                      |
| Teams³      | `teams_list_chats`, `teams_list_chat_messages`, `teams_send_chat_message`¹                                                                                                                                                                                                        |

¹ Only listed when outbound tools are enabled.
² Refused for meetings with attendees unless outbound tools are enabled, because Graph notifies the attendees.
³ Work or school accounts only.

## Requirements

- Node.js 20 or later
- An Entra ID app registration (public client) with **Allow public client flows** turned on, and the delegated permissions below

| Permission                     | Used by                        | Requested when                             |
| ------------------------------ | ------------------------------ | ------------------------------------------ |
| `User.Read`                    | health check, profile          | always                                     |
| `Mail.ReadWrite`               | mail read, drafts, attachments | always                                     |
| `Calendars.ReadWrite`          | calendar                       | always                                     |
| `Files.ReadWrite`              | OneDrive                       | always                                     |
| `Sites.Read.All`               | SharePoint                     | work/school                                |
| `Chat.Read` / `Chat.ReadWrite` | Teams read / send              | work/school (ReadWrite only with outbound) |
| `Mail.Send`                    | `outlook_send_mail`            | outbound enabled                           |

## Install

```bash
git clone https://github.com/kelvinuduma/microsoft-graph-mcp.git
cd microsoft-graph-mcp
npm ci
npm run build
npm test
```

## Sign in

Sign in once per profile from a terminal. The token is cached encrypted (DPAPI on Windows, Keychain on macOS, libsecret on Linux):

```bash
# macOS / Linux
MICROSOFT_CLIENT_ID=<app-id> MICROSOFT_TENANT_ID=<tenant-id> npm run auth
```

```powershell
# Windows: reads .env / .env.<profile>, copies the code, opens the browser
.\scripts\login.ps1                    # profile "corporate"
.\scripts\login.ps1 -Profile personal
.\scripts\login.ps1 -SignOut           # forget a profile's sign-in
```

The command ends with a health check (account, mailbox status, licences). You can also sign in later from inside the agent: the first tool call returns the device code.

## Configure your MCP client

Claude Desktop (`claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "microsoft-graph": {
      "command": "node",
      "args": ["/absolute/path/to/microsoft-graph-mcp/build/index.js"],
      "env": {
        "MICROSOFT_CLIENT_ID": "<app-id>",
        "MICROSOFT_TENANT_ID": "<tenant-id>",
        "GRAPH_MCP_TIMEZONE": "W. Central Africa Standard Time"
      }
    },
    "microsoft-graph-personal": {
      "command": "node",
      "args": ["/absolute/path/to/microsoft-graph-mcp/build/index.js"],
      "env": {
        "MICROSOFT_CLIENT_ID": "<app-id>",
        "MICROSOFT_TENANT_ID": "consumers",
        "MICROSOFT_PROFILE": "personal",
        "TOOL_PREFIX": "personal_"
      }
    }
  }
}
```

Codex (`~/.codex/config.toml`):

```toml
[mcp_servers.microsoft-graph]
command = "node"
args = ["/absolute/path/to/microsoft-graph-mcp/build/index.js"]

[mcp_servers.microsoft-graph.env]
MICROSOFT_CLIENT_ID = "<app-id>"
MICROSOFT_TENANT_ID = "<tenant-id>"
```

## Settings

Settings come from the client's `env` block first, then `.env.<profile>`, then `.env` (both next to `package.json`, both git-ignored).

| Variable                            | Default     | Purpose                                                          |
| ----------------------------------- | ----------- | ---------------------------------------------------------------- |
| `MICROSOFT_CLIENT_ID`               | (required)  | App registration (client) ID                                     |
| `MICROSOFT_TENANT_ID`               | `common`    | Tenant ID, `organizations`, or `consumers` for personal accounts |
| `MICROSOFT_PROFILE`                 | `corporate` | Profile name; selects the token cache and `.env.<profile>`       |
| `TOOL_PREFIX`                       | (none)      | Prefix for tool names when running several profiles              |
| `GRAPH_MCP_ENABLE_OUTBOUND`         | `false`     | Enable tools that notify other people                            |
| `GRAPH_MCP_WRITE_DIRS`              | home + temp | Allowed download folders (`;` on Windows, `:` elsewhere)         |
| `GRAPH_MCP_READ_DIRS`               | home        | Allowed attachment source folders                                |
| `GRAPH_MCP_DENY_PATHS`              | (none)      | Extra path fragments to refuse for reads and writes              |
| `GRAPH_MCP_ALLOW_EXECUTABLE_WRITES` | `false`     | Allow saving `.exe`, `.ps1`, `.js` and similar files             |
| `GRAPH_MCP_MAX_TRANSFER_MB`         | `150`       | Largest download, attachment or filing                           |
| `GRAPH_MCP_TIMEZONE`                | `UTC`       | Calendar time zone (Windows or IANA name)                        |
| `GRAPH_MCP_NO_BROWSER`              | `false`     | Don't open a browser or copy the code during sign-in             |
| `MICROSOFT_CACHE_PATH`              | per profile | Override the token cache file                                    |
| `MSAL_USE_PLAINTEXT_CACHE`          | `false`     | Linux without a keyring only                                     |

## Troubleshooting

| Symptom                                                                               | Cause and fix                                                                                                                             |
| ------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------- |
| `MailboxNotEnabledForRESTAPI` / health check says "no active Exchange Online licence" | The account's licence was removed or expired. An administrator must reassign it; a soft-deleted mailbox is recoverable for about 30 days. |
| Tool returns "Microsoft sign-in required ... enter the code"                          | Open the link, enter the code, retry. Or run `npm run auth` / `scripts\login.ps1`.                                                        |
| "is outside the allowed folders" / "protected location"                               | A guardrail fired. Pick another path, or widen `GRAPH_MCP_WRITE_DIRS` / `GRAPH_MCP_READ_DIRS` deliberately.                               |
| "outbound tools are disabled"                                                         | Use a draft tool, or set `GRAPH_MCP_ENABLE_OUTBOUND=true` for that server.                                                                |
| Token cache errors after a Windows reinstall or profile move                          | The DPAPI key changed. The server resets the cache; sign in again.                                                                        |

## Development

```bash
npm run dev     # tsc --watch
npm test        # build + node:test unit tests (guardrails, tool filtering, scopes)
```

The code is split into `config.ts` (settings), `auth.ts` (MSAL and token cache), `scopes.ts`, `policy.ts` (guardrails), `tools.ts` (catalogue), `handlers.ts` (Graph calls), `errors.ts` and `index.ts` (MCP wiring).

## Security

See [SECURITY.md](SECURITY.md) for the threat model and how to report a vulnerability.

## License

MIT
