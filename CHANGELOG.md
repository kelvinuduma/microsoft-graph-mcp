# Changelog

## 2.0.0 (2026-10-19)

Security hardening and refactor. **Breaking:** outbound tools are now opt-in.

### Security

- Outbound tools (`outlook_send_mail`, `teams_send_chat_message`, and meeting changes that notify attendees) are disabled unless `GRAPH_MCP_ENABLE_OUTBOUND=true`. `Mail.Send` / `Chat.ReadWrite` are requested only then.
- Local file guardrails for downloads and attachments: narrow default roots (writes: `~/Downloads/graph-mcp`; reads: Documents/Downloads/Desktop), rejection of NTFS alternate data streams, device names and trailing dots/spaces, deny list for agent and tool config files (`.claude.json`, `.gitconfig`, `AGENTS.md`...), built-in deny list (keys, credentials, agent configs, startup folders, token caches), symlink resolution, no silent overwrite, executable block, size limit.
- Model-supplied IDs are URL-encoded as path segments. Pagination links are restricted to next pages of the same resource on `graph.microsoft.com` (no `$expand`/`$search`), preventing token leakage and query smuggling.
- Pre-authenticated download URLs are removed from tool output.
- Third-party content (mail, chats, invites) is returned with a prompt-injection notice.
- Sign-in helpers use `execFile` with argument arrays instead of shell strings. The device code reaches the clipboard via `clip.exe` stdin.
- Sign-out and cache resets also delete the Keychain / libsecret entry on macOS and Linux.
- OneDrive writes into a folder shared with other people count as outbound (permissions checked before upload, filing or folder creation).
- Removed hard-coded deployment paths and organisation-specific values from the code; deny rules are now configuration (`GRAPH_MCP_DENY_PATHS`).

### Reliability

- The MCP handshake no longer blocks on sign-in. The first tool call returns the device code and the next call succeeds.
- dotenv runs with `quiet: true` (dotenv 17+ logs to stdout by default, which corrupts stdio JSON-RPC).
- Actionable error messages with Graph error codes and request IDs, including licence/mailbox diagnosis (`MailboxNotEnabledForRESTAPI`).
- Drive uploads above 4 MB use upload sessions. Transfers have timeouts.
- `$filter` and `$orderby` are no longer combined on message lists (avoids `InefficientFilter`).

### Added

- `graph_health_check`: account, scopes, mailbox reachability, licence status.
- `outlook_create_reply_draft`: reply / reply-all draft in the original thread.
- `teams_list_chat_messages`.
- cc/bcc on drafts and mail. Plain-text body option on `outlook_get_message`. Folder choice and paging on lists.
- Calendar time zone support (`timeZone` argument, `GRAPH_MCP_TIMEZONE`). Teams meeting links on new events.
- `onedrive_upload_file` `conflictBehavior` (default `rename`, previously silent replace).
- MCP tool annotations (`readOnlyHint`, `destructiveHint`, `openWorldHint`).
- macOS support for the sign-in convenience (pbcopy/open).
- `scripts/login.ps1 -Profile <name> [-SignOut]` replaces the per-account login scripts.
- Unit tests (`npm test`) for guardrails, tool filtering and scopes.

### Changed

- Dependencies: `@modelcontextprotocol/sdk` 1.32, `@azure/msal-node` 7.1, `@azure/msal-node-extensions` 5.5, `dotenv` 18, TypeScript 7. Removed `isomorphic-fetch` (Node 20+ has `fetch`). Node 20 or later is required.
- CI: Node 20/22/24 matrix, tests, `npm audit`, identifier leak check, SHA-pinned actions.
