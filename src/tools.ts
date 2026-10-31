import { Tool } from "@modelcontextprotocol/sdk/types.js";
import { Config } from "./config.js";

/**
 * Tool catalogue. `meta` drives filtering and response handling:
 * - outbound:      notifies other people; hidden unless GRAPH_MCP_ENABLE_OUTBOUND=true
 * - orgOnly:       needs a work/school tenant (SharePoint, Teams)
 * - untrusted:     returns third-party content (mail, chats, invites); the
 *                  response is prefixed with a prompt-injection notice
 */
export interface ToolMeta {
  outbound?: boolean;
  orgOnly?: boolean;
  untrusted?: boolean;
}

export type ToolDef = Tool & { meta: ToolMeta };

const str = (description: string) => ({ type: "string", description });
const num = (description: string) => ({ type: "number", description });
const bool = (description: string) => ({ type: "boolean", description });
const emails = (description: string) => ({
  type: "array",
  items: { type: "string" },
  description,
});

const READ = { readOnlyHint: true, openWorldHint: false };
const WRITE = {
  readOnlyHint: false,
  destructiveHint: false,
  openWorldHint: false,
};
const LOCAL_WRITE = {
  readOnlyHint: false,
  destructiveHint: false,
  openWorldHint: false,
};

const nextLink = str(
  "Optional @odata.nextLink from a previous page of results. Must be a https://graph.microsoft.com/ URL.",
);
const destinationPath = str(
  "Absolute local path for the saved file, including the file name. Must be inside an allowed folder (GRAPH_MCP_WRITE_DIRS).",
);
const overwrite = bool(
  "Replace the local file if it already exists (default: false).",
);
const timeZone = str(
  "Optional time-zone name for the given times, e.g. 'Africa/Lagos' or 'W. Central Africa Standard Time' (default: GRAPH_MCP_TIMEZONE or UTC).",
);

export const MS_GRAPH_TOOLS: ToolDef[] = [
  // --- DIAGNOSTICS ---
  {
    name: "graph_health_check",
    description:
      "Check the connection: signed-in account, profile, requested scopes, whether outbound tools are enabled, and whether the mailbox is reachable (detects removed or expired licences).",
    inputSchema: { type: "object", properties: {} },
    annotations: { title: "Health check", ...READ },
    meta: {},
  },

  // --- OUTLOOK (MAIL) ---
  {
    name: "outlook_list_messages",
    description:
      "List recent messages in the Inbox (or another well-known folder), newest first.",
    inputSchema: {
      type: "object",
      properties: {
        top: num("Number of messages (default 10, max 50)"),
        filter: str('Optional OData filter, e.g. "isRead eq false"'),
        folder: str(
          "Optional well-known folder: inbox (default), sentitems, drafts, archive, deleteditems, junkemail",
        ),
        nextLink,
      },
    },
    annotations: { title: "List messages", ...READ },
    meta: { untrusted: true },
  },
  {
    name: "outlook_get_message",
    description: "Get one message by ID, including its body.",
    inputSchema: {
      type: "object",
      properties: {
        messageId: str("Message ID from a list or search result"),
        bodyType: {
          type: "string",
          enum: ["text", "html"],
          description: "Body format (default text, which is smaller)",
        },
      },
      required: ["messageId"],
    },
    annotations: { title: "Get message", ...READ },
    meta: { untrusted: true },
  },
  {
    name: "outlook_search_messages",
    description:
      "Search all mail folders using Microsoft Search (KQL), e.g. 'from:alice subject:invoice'.",
    inputSchema: {
      type: "object",
      properties: {
        query: str("KQL search text"),
        top: num("Number of results (default 25, max 50)"),
      },
      required: ["query"],
    },
    annotations: { title: "Search mail", ...READ },
    meta: { untrusted: true },
  },
  {
    name: "outlook_create_draft",
    description: "Create a new draft in the Drafts folder. Nothing is sent.",
    inputSchema: {
      type: "object",
      properties: {
        to: emails("Recipient addresses"),
        cc: emails("Optional Cc addresses"),
        bcc: emails("Optional Bcc addresses"),
        subject: str("Subject line"),
        body: str("Body; HTML is detected automatically, otherwise plain text"),
        importance: {
          type: "string",
          enum: ["low", "normal", "high"],
          description: "Importance",
        },
      },
      required: ["to", "subject", "body"],
    },
    annotations: { title: "Create draft", ...WRITE },
    meta: {},
  },
  {
    name: "outlook_create_reply_draft",
    description:
      "Create a reply (or reply-all) draft in the original thread, with optional text above the quoted message. Nothing is sent.",
    inputSchema: {
      type: "object",
      properties: {
        messageId: str("ID of the message being replied to"),
        comment: str("Optional reply text (HTML or plain text)"),
        replyAll: bool("Reply to all recipients (default false)"),
      },
      required: ["messageId"],
    },
    annotations: { title: "Create reply draft", ...WRITE },
    meta: {},
  },
  {
    name: "outlook_send_mail",
    description:
      "Send an email immediately. Only available when outbound tools are enabled. Confirm recipients and content with the user first.",
    inputSchema: {
      type: "object",
      properties: {
        to: emails("Recipient addresses"),
        cc: emails("Optional Cc addresses"),
        bcc: emails("Optional Bcc addresses"),
        subject: str("Subject line"),
        body: str("Body; HTML is detected automatically, otherwise plain text"),
        importance: {
          type: "string",
          enum: ["low", "normal", "high"],
          description: "Importance",
        },
      },
      required: ["to", "subject", "body"],
    },
    annotations: {
      title: "Send mail",
      readOnlyHint: false,
      destructiveHint: true,
      openWorldHint: true,
    },
    meta: { outbound: true },
  },
  {
    name: "outlook_list_attachments",
    description:
      "List a message's attachments (ID, name, type, size, inline flag).",
    inputSchema: {
      type: "object",
      properties: { messageId: str("Message ID") },
      required: ["messageId"],
    },
    annotations: { title: "List attachments", ...READ },
    meta: {},
  },
  {
    name: "outlook_download_attachment",
    description: "Save a file attachment to local disk.",
    inputSchema: {
      type: "object",
      properties: {
        messageId: str("Message ID"),
        attachmentId: str("Attachment ID"),
        destinationPath,
        overwrite,
      },
      required: ["messageId", "attachmentId", "destinationPath"],
    },
    annotations: { title: "Download attachment", ...LOCAL_WRITE },
    meta: {},
  },
  {
    name: "outlook_add_attachment",
    description:
      "Attach a local file to an existing draft. Files over 3 MB use a resumable upload session.",
    inputSchema: {
      type: "object",
      properties: {
        messageId: str("Draft message ID"),
        filePath: str(
          "Absolute path of the local file (must be inside GRAPH_MCP_READ_DIRS)",
        ),
        filename: str("Optional display name"),
        contentType: str("Optional MIME type, e.g. application/pdf"),
      },
      required: ["messageId", "filePath"],
    },
    annotations: { title: "Add attachment", ...WRITE },
    meta: {},
  },
  {
    name: "outlook_file_attachment_to_onedrive",
    description: "Copy an email attachment straight into OneDrive.",
    inputSchema: {
      type: "object",
      properties: {
        messageId: str("Message ID"),
        attachmentId: str("Attachment ID"),
        parentFolderId: str("Optional OneDrive folder ID (default: root)"),
        filename: str(
          "Optional file name in OneDrive (default: attachment name)",
        ),
      },
      required: ["messageId", "attachmentId"],
    },
    annotations: { title: "File attachment to OneDrive", ...WRITE },
    meta: {},
  },

  // --- CALENDAR ---
  {
    name: "calendar_list_events",
    description:
      "List events (including recurring instances) between two times.",
    inputSchema: {
      type: "object",
      properties: {
        startDateTime: str("ISO 8601 start, e.g. 2026-10-20T00:00:00"),
        endDateTime: str("ISO 8601 end, e.g. 2026-10-27T00:00:00"),
        top: num("Max events (default 50, max 100)"),
        timeZone,
      },
      required: ["startDateTime", "endDateTime"],
    },
    annotations: { title: "List events", ...READ },
    meta: { untrusted: true },
  },
  {
    name: "calendar_get_event",
    description: "Get one event by ID.",
    inputSchema: {
      type: "object",
      properties: { eventId: str("Event ID"), timeZone },
      required: ["eventId"],
    },
    annotations: { title: "Get event", ...READ },
    meta: { untrusted: true },
  },
  {
    name: "calendar_create_event",
    description:
      "Create an event. Adding attendees sends them invitations, so that needs outbound tools enabled.",
    inputSchema: {
      type: "object",
      properties: {
        subject: str("Subject"),
        startDateTime: str(
          "ISO 8601 start, without offset when timeZone is given",
        ),
        endDateTime: str("ISO 8601 end"),
        timeZone,
        body: str("Optional description"),
        attendees: emails("Optional attendee addresses (sends invitations)"),
        location: str("Optional location"),
        isOnlineMeeting: bool("Create a Teams meeting link (default false)"),
      },
      required: ["subject", "startDateTime", "endDateTime"],
    },
    annotations: { title: "Create event", ...WRITE },
    meta: {},
  },
  {
    name: "calendar_update_event",
    description:
      "Update an event. Changing a meeting that has attendees notifies them, so that needs outbound tools enabled.",
    inputSchema: {
      type: "object",
      properties: {
        eventId: str("Event ID"),
        subject: str("New subject"),
        startDateTime: str("New ISO 8601 start"),
        endDateTime: str("New ISO 8601 end"),
        timeZone,
        body: str("New description"),
        location: str("New location"),
      },
      required: ["eventId"],
    },
    annotations: { title: "Update event", ...WRITE },
    meta: {},
  },
  {
    name: "calendar_delete_event",
    description:
      "Delete an event. Deleting a meeting with attendees sends cancellations, so that needs outbound tools enabled.",
    inputSchema: {
      type: "object",
      properties: { eventId: str("Event ID") },
      required: ["eventId"],
    },
    annotations: {
      title: "Delete event",
      readOnlyHint: false,
      destructiveHint: true,
      openWorldHint: false,
    },
    meta: {},
  },

  // --- ONEDRIVE ---
  {
    name: "onedrive_list_root",
    description: "List items in the OneDrive root folder.",
    inputSchema: {
      type: "object",
      properties: { top: num("Max items (default 50, max 200)"), nextLink },
    },
    annotations: { title: "List OneDrive root", ...READ },
    meta: {},
  },
  {
    name: "onedrive_list_folder",
    description: "List items in a OneDrive folder.",
    inputSchema: {
      type: "object",
      properties: {
        folderId: str("Folder ID"),
        top: num("Max items (default 50, max 200)"),
        nextLink,
      },
      required: ["folderId"],
    },
    annotations: { title: "List OneDrive folder", ...READ },
    meta: {},
  },
  {
    name: "onedrive_search_files",
    description: "Search OneDrive for files matching a query.",
    inputSchema: {
      type: "object",
      properties: {
        query: str("Search text"),
        top: num("Max results (default 25, max 200)"),
      },
      required: ["query"],
    },
    annotations: { title: "Search OneDrive", ...READ },
    meta: {},
  },
  {
    name: "onedrive_download_file",
    description:
      "Get a file's metadata, or save it locally when destinationPath is given.",
    inputSchema: {
      type: "object",
      properties: { fileId: str("File ID"), destinationPath, overwrite },
      required: ["fileId"],
    },
    annotations: { title: "Download OneDrive file", ...LOCAL_WRITE },
    meta: {},
  },
  {
    name: "onedrive_upload_file",
    description: "Create a small text file in OneDrive (up to 4 MB).",
    inputSchema: {
      type: "object",
      properties: {
        filename: str("File name, e.g. notes.txt"),
        content: str("Text content"),
        parentFolderId: str("Optional folder ID (default: root)"),
        conflictBehavior: {
          type: "string",
          enum: ["rename", "fail", "replace"],
          description: "If the name exists: rename (default), fail, or replace",
        },
      },
      required: ["filename", "content"],
    },
    annotations: { title: "Upload to OneDrive", ...WRITE },
    meta: {},
  },
  {
    name: "onedrive_create_folder",
    description:
      "Create a folder in OneDrive (renamed automatically if the name exists).",
    inputSchema: {
      type: "object",
      properties: {
        folderName: str("Folder name"),
        parentFolderId: str("Optional parent folder ID"),
      },
      required: ["folderName"],
    },
    annotations: { title: "Create OneDrive folder", ...WRITE },
    meta: {},
  },

  // --- SHAREPOINT ---
  {
    name: "sharepoint_list_sites",
    description: "Search SharePoint sites the user can access.",
    inputSchema: {
      type: "object",
      properties: { searchQuery: str("Optional search text (default: all)") },
    },
    annotations: { title: "List SharePoint sites", ...READ },
    meta: { orgOnly: true },
  },
  {
    name: "sharepoint_get_site_details",
    description: "List the document libraries (drives) of a SharePoint site.",
    inputSchema: {
      type: "object",
      properties: { siteId: str("Site ID") },
      required: ["siteId"],
    },
    annotations: { title: "Get SharePoint site libraries", ...READ },
    meta: { orgOnly: true },
  },
  {
    name: "sharepoint_list_drive_items",
    description: "List items in a SharePoint library root or subfolder.",
    inputSchema: {
      type: "object",
      properties: {
        siteId: str("Site ID"),
        driveId: str("Library (drive) ID"),
        folderId: str("Optional folder ID (default: library root)"),
        top: num("Max items (default 50, max 200)"),
        nextLink,
      },
      required: ["siteId", "driveId"],
    },
    annotations: { title: "List SharePoint items", ...READ },
    meta: { orgOnly: true },
  },
  {
    name: "sharepoint_download_file",
    description:
      "Get a SharePoint file's metadata, or save it locally when destinationPath is given.",
    inputSchema: {
      type: "object",
      properties: {
        siteId: str("Site ID"),
        driveId: str("Library (drive) ID"),
        fileId: str("File ID"),
        destinationPath,
        overwrite,
      },
      required: ["siteId", "driveId", "fileId"],
    },
    annotations: { title: "Download SharePoint file", ...LOCAL_WRITE },
    meta: { orgOnly: true },
  },
  {
    name: "sharepoint_search_files",
    description:
      "Search a SharePoint library (or the site's default library) by keyword.",
    inputSchema: {
      type: "object",
      properties: {
        siteId: str("Site ID"),
        query: str("Search text"),
        driveId: str(
          "Optional library ID (default: the site's default library)",
        ),
        top: num("Max results (default 25, max 200)"),
      },
      required: ["siteId", "query"],
    },
    annotations: { title: "Search SharePoint", ...READ },
    meta: { orgOnly: true },
  },

  // --- TEAMS ---
  {
    name: "teams_list_chats",
    description:
      "List recent Teams chats with their topic and last message preview.",
    inputSchema: {
      type: "object",
      properties: { top: num("Max chats (default 10, max 50)") },
    },
    annotations: { title: "List Teams chats", ...READ },
    meta: { orgOnly: true, untrusted: true },
  },
  {
    name: "teams_list_chat_messages",
    description: "Read recent messages in a Teams chat, newest first.",
    inputSchema: {
      type: "object",
      properties: {
        chatId: str("Chat ID"),
        top: num("Max messages (default 20, max 50)"),
      },
      required: ["chatId"],
    },
    annotations: { title: "Read Teams chat", ...READ },
    meta: { orgOnly: true, untrusted: true },
  },
  {
    name: "teams_send_chat_message",
    description:
      "Send a message to a Teams chat. Only available when outbound tools are enabled. Confirm with the user first.",
    inputSchema: {
      type: "object",
      properties: {
        chatId: str("Chat ID"),
        content: str("Message text (HTML or plain text)"),
      },
      required: ["chatId", "content"],
    },
    annotations: {
      title: "Send Teams message",
      readOnlyHint: false,
      destructiveHint: true,
      openWorldHint: true,
    },
    meta: { orgOnly: true, outbound: true },
  },
];

/** Tools this configuration exposes, with the optional name prefix applied. */
export function getTools(cfg: Config): Tool[] {
  return MS_GRAPH_TOOLS.filter(
    (t) =>
      !(t.meta.orgOnly && cfg.isConsumer) &&
      !(t.meta.outbound && !cfg.enableOutbound),
  ).map(({ meta: _meta, ...tool }) => ({
    ...tool,
    name: `${cfg.toolPrefix}${tool.name}`,
  }));
}

export function findTool(baseName: string): ToolDef | undefined {
  return MS_GRAPH_TOOLS.find((t) => t.name === baseName);
}
