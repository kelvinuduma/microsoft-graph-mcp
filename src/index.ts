import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import { Client } from "@microsoft/microsoft-graph-client";
import { MSGraphAuth } from "./auth.js";
import { MS_GRAPH_TOOLS } from "./tools.js";
import * as dotenv from "dotenv";
import * as path from "path";
import "isomorphic-fetch";

// Load environment variables from repo .env and current working directory
dotenv.config({ path: path.resolve(__dirname, "../.env") });
dotenv.config();

class MSGraphMCPServer {
  private server: Server;
  private auth: MSGraphAuth;
  private graphClient!: Client;

  constructor() {
    this.server = new Server(
      {
        name: "microsoft-graph-mcp",
        version: "1.0.0",
      },
      {
        capabilities: {
          tools: {},
        },
      },
    );

    this.auth = new MSGraphAuth();
    this.setupHandlers();
  }

  /**
   * Initializes MSAL Authentication and Microsoft Graph Client.
   */
  private async initGraphClient(): Promise<void> {
    try {
      await this.auth.initialize();

      // Warm up authentication (forces device-code flow in terminal if not already cached)
      console.error("[Auth] Warming up authentication...");
      await this.auth.getAccessToken();
      console.error("[Auth] Authentication token successfully verified/cached");

      // Initialize Graph Client using custom Auth Provider wrapping MSAL auth
      this.graphClient = Client.initWithMiddleware({
        authProvider: {
          getAccessToken: async () => {
            return await this.auth.getAccessToken();
          },
        },
      });

      console.error(
        "[Graph Client] Successfully initialized Microsoft Graph Client",
      );
    } catch (error: any) {
      console.error(
        `[Init Error] Failed to initialize Graph Client: ${error.message}`,
      );
      process.exit(1);
    }
  }

  /**
   * Configures MCP server request handlers.
   */
  private setupHandlers(): void {
    // 1. List available tools
    this.server.setRequestHandler(ListToolsRequestSchema, async () => {
      return {
        tools: MS_GRAPH_TOOLS,
      };
    });

    // 2. Dispatch tool calls
    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      const { name, arguments: args } = request.params;
      console.error(`[Request] Executing tool: ${name}`);

      try {
        const result = await this.executeTool(name, args || {});
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(result, null, 2),
            },
          ],
        };
      } catch (error: any) {
        console.error(`[Tool Execution Error] ${name} failed:`, error);
        return {
          isError: true,
          content: [
            {
              type: "text",
              text: `Error executing ${name}: ${error.message || error}`,
            },
          ],
        };
      }
    });
  }

  /**
   * Routes tool calls to Microsoft Graph SDK queries.
   */
  private async executeTool(name: string, args: any): Promise<any> {
    switch (name) {
      // --- OUTLOOK (MAIL) ---
      case "outlook_list_messages": {
        const top = args.top || 10;
        let query = this.graphClient
          .api("/me/mailFolders/inbox/messages")
          .top(top);
        if (args.filter) {
          query = query.filter(args.filter);
        }
        return await query
          .select("id,subject,from,receivedDateTime,isRead,bodyPreview")
          .get();
      }

      case "outlook_get_message": {
        return await this.graphClient
          .api(`/me/messages/${args.messageId}`)
          .get();
      }

      case "outlook_send_mail": {
        const recipients = args.to.map((email: string) => ({
          emailAddress: { address: email },
        }));
        const message = {
          subject: args.subject,
          body: {
            contentType: args.body.includes("<") ? "html" : "text",
            content: args.body,
          },
          toRecipients: recipients,
          importance: args.importance || "normal",
        };
        await this.graphClient.api("/me/sendMail").post({ message });
        return { status: "success", message: "Email sent successfully" };
      }

      case "outlook_create_draft": {
        const recipients = args.to.map((email: string) => ({
          emailAddress: { address: email },
        }));
        const message = {
          subject: args.subject,
          body: {
            contentType: args.body.includes("<") ? "html" : "text",
            content: args.body,
          },
          toRecipients: recipients,
        };
        return await this.graphClient.api("/me/messages").post(message);
      }

      case "outlook_search_messages": {
        return await this.graphClient
          .api("/me/messages")
          .search(args.query)
          .select("id,subject,from,receivedDateTime,bodyPreview")
          .get();
      }

      // --- CALENDAR ---
      case "calendar_list_events": {
        return await this.graphClient
          .api("/me/calendarView")
          .query({
            startDateTime: args.startDateTime,
            endDateTime: args.endDateTime,
          })
          .select("id,subject,start,end,location,attendees")
          .get();
      }

      case "calendar_get_event": {
        return await this.graphClient.api(`/me/events/${args.eventId}`).get();
      }

      case "calendar_create_event": {
        const event = {
          subject: args.subject,
          body: args.body
            ? {
                contentType: "text",
                content: args.body,
              }
            : undefined,
          start: {
            dateTime: args.startDateTime,
            timeZone: "UTC",
          },
          end: {
            dateTime: args.endDateTime,
            timeZone: "UTC",
          },
          location: args.location ? { displayName: args.location } : undefined,
          attendees: args.attendees
            ? args.attendees.map((email: string) => ({
                emailAddress: { address: email },
                type: "required",
              }))
            : [],
        };
        return await this.graphClient.api("/me/events").post(event);
      }

      case "calendar_update_event": {
        const updates: any = {};
        if (args.subject) updates.subject = args.subject;
        if (args.body) {
          updates.body = { contentType: "text", content: args.body };
        }
        if (args.startDateTime) {
          updates.start = { dateTime: args.startDateTime, timeZone: "UTC" };
        }
        if (args.endDateTime) {
          updates.end = { dateTime: args.endDateTime, timeZone: "UTC" };
        }
        if (args.location) {
          updates.location = { displayName: args.location };
        }
        return await this.graphClient
          .api(`/me/events/${args.eventId}`)
          .update(updates);
      }

      case "calendar_delete_event": {
        await this.graphClient.api(`/me/events/${args.eventId}`).delete();
        return { status: "success", message: `Event ${args.eventId} deleted` };
      }

      // --- ONEDRIVE & SHAREPOINT ---
      case "onedrive_list_root": {
        return await this.graphClient.api("/me/drive/root/children").get();
      }

      case "onedrive_list_folder": {
        return await this.graphClient
          .api(`/me/drive/items/${args.folderId}/children`)
          .get();
      }

      case "onedrive_search_files": {
        const q = String(args.query).replace(/'/g, "''");
        return await this.graphClient
          .api(`/me/drive/root/search(q='${q}')`)
          .get();
      }

      case "onedrive_download_file": {
        return await this.graphClient
          .api(`/me/drive/items/${args.fileId}`)
          .get();
      }

      case "onedrive_upload_file": {
        const folderPath = args.parentFolderId
          ? `/me/drive/items/${args.parentFolderId}:/${args.filename}:/content`
          : `/me/drive/root:/${args.filename}:/content`;
        return await this.graphClient.api(folderPath).put(args.content);
      }

      case "onedrive_create_folder": {
        const parentPath = args.parentFolderId
          ? `/me/drive/items/${args.parentFolderId}/children`
          : "/me/drive/root/children";
        const folder = {
          name: args.folderName,
          folder: {},
          "@microsoft.graph.conflictBehavior": "rename",
        };
        return await this.graphClient.api(parentPath).post(folder);
      }

      case "sharepoint_list_sites": {
        const apiPath = args.searchQuery
          ? `/sites?search=${encodeURIComponent(args.searchQuery)}`
          : "/sites?search=*";
        return await this.graphClient.api(apiPath).get();
      }

      case "sharepoint_get_site_details": {
        return await this.graphClient.api(`/sites/${args.siteId}/drives`).get();
      }

      // --- TEAMS ---
      case "teams_list_chats": {
        const top = args.top || 10;
        return await this.graphClient.api("/me/chats").top(top).get();
      }

      case "teams_send_chat_message": {
        const message = {
          body: {
            content: args.content,
          },
        };
        return await this.graphClient
          .api(`/chats/${args.chatId}/messages`)
          .post(message);
      }

      default:
        throw new Error(`Tool handler for ${name} is not implemented.`);
    }
  }

  /**
   * Starts the MCP server on Stdio transport.
   */
  public async start(): Promise<void> {
    await this.initGraphClient();
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
    console.error("[MCP Server] Microsoft Graph MCP Server listening on Stdio");
  }
}

// Start Server
const serverInstance = new MSGraphMCPServer();
serverInstance.start().catch((err) => {
  console.error("[Fatal Error] Server failed to start:", err);
  process.exit(1);
});
