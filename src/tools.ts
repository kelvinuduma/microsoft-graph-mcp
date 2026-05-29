import { Tool } from "@modelcontextprotocol/sdk/types.js";

export const MS_GRAPH_TOOLS: Tool[] = [
  // --- OUTLOOK (MAIL) ---
  {
    name: "outlook_list_messages",
    description:
      "Retrieve a list of recent email messages from the user's Inbox.",
    inputSchema: {
      type: "object",
      properties: {
        top: {
          type: "number",
          description: "Number of emails to return (default: 10, max: 50)",
        },
        filter: {
          type: "string",
          description: 'OData filter query (e.g. "isRead eq false")',
        },
      },
    },
  },
  {
    name: "outlook_get_message",
    description: "Get detailed contents of a specific email message by its ID.",
    inputSchema: {
      type: "object",
      properties: {
        messageId: {
          type: "string",
          description: "The unique ID of the email message",
        },
      },
      required: ["messageId"],
    },
  },
  {
    name: "outlook_send_mail",
    description: "Send a new email message to one or more recipients.",
    inputSchema: {
      type: "object",
      properties: {
        to: {
          type: "array",
          items: { type: "string" },
          description: "Recipient email addresses",
        },
        subject: { type: "string", description: "Email subject line" },
        body: {
          type: "string",
          description: "HTML or plain text email body content",
        },
        importance: {
          type: "string",
          enum: ["low", "normal", "high"],
          description: "Email importance level",
        },
      },
      required: ["to", "subject", "body"],
    },
  },
  {
    name: "outlook_create_draft",
    description: "Create a new draft email in the user's Drafts folder.",
    inputSchema: {
      type: "object",
      properties: {
        to: {
          type: "array",
          items: { type: "string" },
          description: "Recipient email addresses",
        },
        subject: { type: "string", description: "Email subject line" },
        body: {
          type: "string",
          description: "HTML or plain text body content",
        },
      },
      required: ["to", "subject", "body"],
    },
  },
  {
    name: "outlook_search_messages",
    description:
      "Search the user's email messages using keyword query language (KQL).",
    inputSchema: {
      type: "object",
      properties: {
        query: {
          type: "string",
          description: "The search query (e.g., 'from:boss', 'project status')",
        },
      },
      required: ["query"],
    },
  },

  // --- CALENDAR ---
  {
    name: "calendar_list_events",
    description: "List upcoming calendar events for a specific time range.",
    inputSchema: {
      type: "object",
      properties: {
        startDateTime: {
          type: "string",
          description: "ISO 8601 start time (e.g., '2026-05-27T00:00:00Z')",
        },
        endDateTime: {
          type: "string",
          description: "ISO 8601 end time (e.g., '2026-05-30T23:59:59Z')",
        },
      },
      required: ["startDateTime", "endDateTime"],
    },
  },
  {
    name: "calendar_get_event",
    description:
      "Get detailed information about a specific calendar event by its ID.",
    inputSchema: {
      type: "object",
      properties: {
        eventId: {
          type: "string",
          description: "The unique ID of the calendar event",
        },
      },
      required: ["eventId"],
    },
  },
  {
    name: "calendar_create_event",
    description: "Create a new event or meeting in the user's calendar.",
    inputSchema: {
      type: "object",
      properties: {
        subject: { type: "string", description: "Subject of the meeting" },
        startDateTime: {
          type: "string",
          description:
            "ISO 8601 start date-time (e.g., '2026-05-28T14:00:00Z')",
        },
        endDateTime: {
          type: "string",
          description: "ISO 8601 end date-time (e.g., '2026-05-28T15:00:00Z')",
        },
        body: {
          type: "string",
          description: "Description or body text for the calendar event",
        },
        attendees: {
          type: "array",
          items: { type: "string" },
          description: "Attendee email addresses",
        },
        location: { type: "string", description: "Location of the meeting" },
      },
      required: ["subject", "startDateTime", "endDateTime"],
    },
  },
  {
    name: "calendar_update_event",
    description: "Update details of an existing calendar event.",
    inputSchema: {
      type: "object",
      properties: {
        eventId: {
          type: "string",
          description: "The unique ID of the event to update",
        },
        subject: { type: "string", description: "New subject of the meeting" },
        startDateTime: {
          type: "string",
          description: "New ISO 8601 start date-time",
        },
        endDateTime: {
          type: "string",
          description: "New ISO 8601 end date-time",
        },
        body: { type: "string", description: "New description body text" },
        location: { type: "string", description: "New location" },
      },
      required: ["eventId"],
    },
  },
  {
    name: "calendar_delete_event",
    description: "Delete or cancel a calendar event by its ID.",
    inputSchema: {
      type: "object",
      properties: {
        eventId: {
          type: "string",
          description: "The unique ID of the event to delete",
        },
      },
      required: ["eventId"],
    },
  },

  // --- ONEDRIVE & SHAREPOINT ---
  {
    name: "onedrive_list_root",
    description: "List all items in the root folder of the user's OneDrive.",
    inputSchema: {
      type: "object",
      properties: {},
    },
  },
  {
    name: "onedrive_list_folder",
    description:
      "List all items inside a specific OneDrive folder by its folder ID.",
    inputSchema: {
      type: "object",
      properties: {
        folderId: {
          type: "string",
          description: "The unique ID of the OneDrive folder",
        },
      },
      required: ["folderId"],
    },
  },
  {
    name: "onedrive_search_files",
    description: "Search for files in the user's OneDrive matching a query.",
    inputSchema: {
      type: "object",
      properties: {
        query: {
          type: "string",
          description: "Search query text (e.g. 'report')",
        },
      },
      required: ["query"],
    },
  },
  {
    name: "onedrive_download_file",
    description:
      "Retrieve download content link or file details for a file ID.",
    inputSchema: {
      type: "object",
      properties: {
        fileId: { type: "string", description: "The unique ID of the file" },
      },
      required: ["fileId"],
    },
  },
  {
    name: "onedrive_upload_file",
    description: "Upload content to create or replace a file in OneDrive.",
    inputSchema: {
      type: "object",
      properties: {
        filename: {
          type: "string",
          description: "Name of the file (e.g. 'document.txt')",
        },
        content: { type: "string", description: "Text content of the file" },
        parentFolderId: {
          type: "string",
          description: "Optional folder ID to upload into (defaults to root)",
        },
      },
      required: ["filename", "content"],
    },
  },
  {
    name: "onedrive_create_folder",
    description: "Create a new folder in OneDrive.",
    inputSchema: {
      type: "object",
      properties: {
        folderName: {
          type: "string",
          description: "Name of the folder to create",
        },
        parentFolderId: {
          type: "string",
          description: "Optional parent folder ID",
        },
      },
      required: ["folderName"],
    },
  },
  {
    name: "sharepoint_list_sites",
    description: "Search and list SharePoint sites the user has access to.",
    inputSchema: {
      type: "object",
      properties: {
        searchQuery: {
          type: "string",
          description: "Search query to filter sites",
        },
      },
    },
  },
  {
    name: "sharepoint_get_site_details",
    description:
      "Retrieve document libraries and lists for a specific SharePoint Site ID.",
    inputSchema: {
      type: "object",
      properties: {
        siteId: {
          type: "string",
          description: "The unique ID of the SharePoint site",
        },
      },
      required: ["siteId"],
    },
  },

  // --- TEAMS ---
  {
    name: "teams_list_chats",
    description: "List recent active chat threads from Microsoft Teams.",
    inputSchema: {
      type: "object",
      properties: {
        top: {
          type: "number",
          description: "Max threads to return (default: 10)",
        },
      },
    },
  },
  {
    name: "teams_send_chat_message",
    description: "Send a message to a specific Teams chat thread.",
    inputSchema: {
      type: "object",
      properties: {
        chatId: {
          type: "string",
          description: "The unique ID of the Teams chat thread",
        },
        content: {
          type: "string",
          description: "The message body content (plain text or HTML)",
        },
      },
      required: ["chatId", "content"],
    },
  },
];
