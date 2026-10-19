import * as fs from "fs";
import * as path from "path";
import { Client, ResponseType } from "@microsoft/microsoft-graph-client";
import { Config } from "./config.js";
import { MSGraphAuth } from "./auth.js";
import { describeError, errorCode } from "./errors.js";
import {
  PolicyError,
  assertGraphUrl,
  clampTop,
  isSharedWithOthers,
  odataString,
  resolveReadPath,
  resolveWritePath,
  seg,
  siteSeg,
} from "./policy.js";

const SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024; // Graph simple upload / PUT content
const ATTACHMENT_INLINE_LIMIT = 3 * 1024 * 1024; // above this, use an upload session
const CHUNK = 320 * 1024 * 10; // 3.2 MB, a multiple of 320 KiB as Graph requires
const HTTP_TIMEOUT_MS = 120_000;

const MAIL_FOLDERS = new Set([
  "inbox",
  "sentitems",
  "drafts",
  "archive",
  "deleteditems",
  "junkemail",
]);
const EMAIL_RE = /^[^\s@<>,;]+@[^\s@<>,;]+\.[^\s@<>,;]+$/;

type Args = Record<string, any>;

function looksLikeHtml(text: string): boolean {
  return /<\/?[a-z][\s\S]*>/i.test(text);
}

function body(content: unknown) {
  const text = String(content ?? "");
  return { contentType: looksLikeHtml(text) ? "html" : "text", content: text };
}

function recipients(value: unknown, field: string, required = false) {
  if (value === undefined || value === null) {
    if (required) throw new PolicyError(`${field} is required.`);
    return undefined;
  }
  const list = Array.isArray(value) ? value : [value];
  if (required && list.length === 0)
    throw new PolicyError(`${field} must list at least one address.`);
  if (list.length > 100)
    throw new PolicyError(`${field} has more than 100 addresses.`);
  return list.map((v) => {
    const address = String(v).trim();
    if (!EMAIL_RE.test(address))
      throw new PolicyError(
        `'${address}' in ${field} is not a valid email address.`,
      );
    return { emailAddress: { address } };
  });
}

/** Drops OData bookkeeping fields that cost tokens and carry no meaning. */
function tidy(value: any, depth = 0): any {
  if (depth > 6 || value === null || typeof value !== "object") return value;
  if (Array.isArray(value)) return value.map((v) => tidy(v, depth + 1));
  if (Buffer.isBuffer(value)) return `<${value.length} bytes>`;
  const out: Record<string, any> = {};
  for (const [k, v] of Object.entries(value)) {
    if (
      k === "@odata.context" ||
      k === "@odata.etag" ||
      k === "@microsoft.graph.downloadUrl"
    )
      continue;
    out[k] = tidy(v, depth + 1);
  }
  return out;
}

function writeLocalFile(dest: string, data: Buffer, overwrite: boolean): void {
  fs.mkdirSync(path.dirname(dest), { recursive: true });
  // "wx" fails if the file appeared after the policy check (no silent clobber).
  fs.writeFileSync(dest, data, { flag: overwrite ? "w" : "wx" });
}

export class GraphHandlers {
  constructor(
    private readonly client: Client,
    private readonly auth: MSGraphAuth,
    private readonly cfg: Config,
  ) {}

  private requireOutbound(action: string): void {
    if (!this.cfg.enableOutbound) {
      throw new PolicyError(
        `${action} would notify other people, and outbound tools are disabled. Create a draft instead, or set GRAPH_MCP_ENABLE_OUTBOUND=true for this server.`,
      );
    }
  }

  private assertSize(bytes: number | undefined, what: string): void {
    if (typeof bytes === "number" && bytes > this.cfg.maxTransferBytes) {
      throw new PolicyError(
        `${what} is ${bytes} bytes, above the ${this.cfg.maxTransferBytes}-byte limit (GRAPH_MCP_MAX_TRANSFER_MB).`,
      );
    }
  }

  private async page(
    apiPath: string,
    args: Args,
    top: number,
    select?: string,
  ) {
    if (args.nextLink)
      return this.client.api(assertGraphUrl(args.nextLink, "children")).get();
    const req = this.client.api(apiPath).top(top);
    return select ? req.select(select).get() : req.get();
  }

  /** Downloads a drive item's bytes, preferring the pre-authenticated URL. */
  private async fetchDriveItem(
    item: any,
    contentPath: string,
  ): Promise<Buffer> {
    if (item.folder)
      throw new PolicyError(`'${item.name}' is a folder, not a file.`);
    this.assertSize(item.size, `'${item.name}'`);
    const url = item["@microsoft.graph.downloadUrl"];
    if (url) {
      const resp = await fetch(url, {
        signal: AbortSignal.timeout(HTTP_TIMEOUT_MS),
      });
      if (!resp.ok) throw new Error(`Download failed: HTTP ${resp.status}`);
      return Buffer.from(await resp.arrayBuffer());
    }
    const data = await this.client
      .api(contentPath)
      .responseType(ResponseType.ARRAYBUFFER)
      .get();
    return Buffer.from(data);
  }

  /** Uploads bytes to a drive folder; uses an upload session above 4 MB. */
  private async uploadToDrive(
    folderPrefix: string,
    filename: string,
    data: Buffer,
    conflictBehavior: string,
  ): Promise<any> {
    const name = seg(filename, "filename");
    if (data.length <= SIMPLE_UPLOAD_LIMIT) {
      return this.client
        .api(
          `${folderPrefix}:/${name}:/content?@microsoft.graph.conflictBehavior=${conflictBehavior}`,
        )
        .put(data);
    }
    const session = await this.client
      .api(`${folderPrefix}:/${name}:/createUploadSession`)
      .post({
        item: { "@microsoft.graph.conflictBehavior": conflictBehavior },
      });
    return this.putChunks(session.uploadUrl, data);
  }

  private async putChunks(uploadUrl: string, data: Buffer): Promise<any> {
    let last: any = null;
    for (let offset = 0; offset < data.length; offset += CHUNK) {
      const end = Math.min(offset + CHUNK, data.length);
      const resp = await fetch(uploadUrl, {
        method: "PUT",
        headers: {
          "Content-Length": String(end - offset),
          "Content-Range": `bytes ${offset}-${end - 1}/${data.length}`,
        },
        body: data.subarray(offset, end),
        signal: AbortSignal.timeout(HTTP_TIMEOUT_MS),
      });
      if (![200, 201, 202].includes(resp.status)) {
        throw new Error(
          `Upload failed at byte ${offset}: HTTP ${resp.status} ${await resp.text()}`,
        );
      }
      const text = await resp.text();
      last = text ? JSON.parse(text) : null;
    }
    return last;
  }

  private async eventHasAttendees(eventId: string): Promise<boolean> {
    const ev = await this.client
      .api(`/me/events/${seg(eventId, "eventId")}`)
      .select("attendees")
      .get();
    return Array.isArray(ev.attendees) && ev.attendees.length > 0;
  }

  /**
   * Resolves the OneDrive folder a write tool targets. Writing into a folder
   * that is shared with other people (directly, by inheritance, or through a
   * sharing link) publishes the content to them, so it counts as outbound.
   * The drive root is never shared and needs no lookup.
   */
  private async writableDriveFolder(parentFolderId: unknown): Promise<string> {
    if (
      parentFolderId === undefined ||
      parentFolderId === null ||
      parentFolderId === ""
    ) {
      return "/me/drive/root";
    }
    const folder = `/me/drive/items/${seg(parentFolderId, "parentFolderId")}`;
    if (!this.cfg.enableOutbound) {
      const perms = await this.client.api(`${folder}/permissions`).get();
      if (isSharedWithOthers(perms.value)) {
        this.requireOutbound(
          "Writing into a OneDrive folder that is shared with other people",
        );
      }
    }
    return folder;
  }

  private tz(args: Args): string {
    return typeof args.timeZone === "string" && args.timeZone.trim()
      ? args.timeZone.trim()
      : this.cfg.timeZone;
  }

  public async execute(name: string, args: Args): Promise<any> {
    return tidy(await this.dispatch(name, args));
  }

  private async dispatch(name: string, args: Args): Promise<any> {
    switch (name) {
      // --- DIAGNOSTICS ---
      case "graph_health_check":
        return this.healthCheck();

      // --- OUTLOOK (MAIL) ---
      case "outlook_list_messages": {
        const folder = String(args.folder || "inbox").toLowerCase();
        if (!MAIL_FOLDERS.has(folder))
          throw new PolicyError(`Unknown folder '${folder}'.`);
        if (args.nextLink)
          return this.client
            .api(assertGraphUrl(args.nextLink, "messages"))
            .get();
        let req = this.client
          .api(`/me/mailFolders/${folder}/messages`)
          .top(clampTop(args.top, 10, 50))
          .select(
            "id,subject,from,receivedDateTime,isRead,hasAttachments,importance,bodyPreview,conversationId",
          );
        // Graph rejects many $filter + $orderby combinations ("InefficientFilter"),
        // so only sort explicitly when there is no filter (the default order is newest first).
        req = args.filter
          ? req.filter(String(args.filter))
          : req.orderby("receivedDateTime desc");
        return req.get();
      }

      case "outlook_get_message": {
        const type = args.bodyType === "html" ? "html" : "text";
        return this.client
          .api(`/me/messages/${seg(args.messageId, "messageId")}`)
          .header("Prefer", `outlook.body-content-type="${type}"`)
          .select(
            "id,subject,from,sender,toRecipients,ccRecipients,replyTo,receivedDateTime,sentDateTime,importance,isRead,hasAttachments,conversationId,internetMessageId,webLink,body",
          )
          .get();
      }

      case "outlook_search_messages": {
        const q = String(args.query ?? "")
          .replace(/\\/g, "\\\\")
          .replace(/"/g, '\\"');
        if (!q.trim()) throw new PolicyError("query must not be empty.");
        return this.client
          .api("/me/messages")
          .search(`"${q}"`)
          .top(clampTop(args.top, 25, 50))
          .select(
            "id,subject,from,receivedDateTime,hasAttachments,bodyPreview,conversationId",
          )
          .get();
      }

      case "outlook_create_draft": {
        const draft = await this.client.api("/me/messages").post({
          subject: String(args.subject ?? ""),
          body: body(args.body),
          toRecipients: recipients(args.to, "to", true),
          ccRecipients: recipients(args.cc, "cc"),
          bccRecipients: recipients(args.bcc, "bcc"),
          importance: args.importance || "normal",
        });
        return {
          status: "draft_created",
          id: draft.id,
          subject: draft.subject,
          webLink: draft.webLink,
        };
      }

      case "outlook_create_reply_draft": {
        const action = args.replyAll ? "createReplyAll" : "createReply";
        const draft = await this.client
          .api(`/me/messages/${seg(args.messageId, "messageId")}/${action}`)
          .post({ comment: String(args.comment ?? "") });
        return {
          status: "reply_draft_created",
          id: draft.id,
          subject: draft.subject,
          to: draft.toRecipients,
          cc: draft.ccRecipients,
          webLink: draft.webLink,
        };
      }

      case "outlook_send_mail": {
        this.requireOutbound("Sending mail");
        await this.client.api("/me/sendMail").post({
          message: {
            subject: String(args.subject ?? ""),
            body: body(args.body),
            toRecipients: recipients(args.to, "to", true),
            ccRecipients: recipients(args.cc, "cc"),
            bccRecipients: recipients(args.bcc, "bcc"),
            importance: args.importance || "normal",
          },
          saveToSentItems: true,
        });
        return { status: "sent" };
      }

      case "outlook_list_attachments":
        return this.client
          .api(`/me/messages/${seg(args.messageId, "messageId")}/attachments`)
          .select("id,name,contentType,size,isInline,lastModifiedDateTime")
          .get();

      case "outlook_download_attachment": {
        const overwrite = args.overwrite === true;
        const dest = resolveWritePath(
          args.destinationPath,
          this.cfg,
          overwrite,
        );
        const att = await this.client
          .api(
            `/me/messages/${seg(args.messageId, "messageId")}/attachments/${seg(args.attachmentId, "attachmentId")}`,
          )
          .get();
        if (
          att["@odata.type"] !== "#microsoft.graph.fileAttachment" ||
          !att.contentBytes
        ) {
          throw new PolicyError(
            "Only file attachments can be saved (this is an item or reference attachment).",
          );
        }
        this.assertSize(att.size, `Attachment '${att.name}'`);
        const data = Buffer.from(att.contentBytes, "base64");
        writeLocalFile(dest, data, overwrite);
        return {
          status: "saved",
          name: att.name,
          savedPath: dest,
          size: data.length,
        };
      }

      case "outlook_add_attachment": {
        const src = resolveReadPath(args.filePath, this.cfg);
        const data = fs.readFileSync(src);
        const filename = String(args.filename || path.basename(src));
        const contentType = String(
          args.contentType || "application/octet-stream",
        );
        const base = `/me/messages/${seg(args.messageId, "messageId")}/attachments`;
        if (data.length <= ATTACHMENT_INLINE_LIMIT) {
          const res = await this.client.api(base).post({
            "@odata.type": "#microsoft.graph.fileAttachment",
            name: filename,
            contentType,
            contentBytes: data.toString("base64"),
          });
          return {
            status: "attached",
            attachmentId: res.id,
            name: filename,
            size: data.length,
          };
        }
        const session = await this.client
          .api(`${base}/createUploadSession`)
          .post({
            AttachmentItem: {
              attachmentType: "file",
              name: filename,
              size: data.length,
              contentType,
            },
          });
        await this.putChunks(session.uploadUrl, data);
        return {
          status: "attached",
          name: filename,
          size: data.length,
          method: "upload_session",
        };
      }

      case "outlook_file_attachment_to_onedrive": {
        const att = await this.client
          .api(
            `/me/messages/${seg(args.messageId, "messageId")}/attachments/${seg(args.attachmentId, "attachmentId")}`,
          )
          .get();
        if (!att.contentBytes)
          throw new PolicyError("Only file attachments can be filed.");
        this.assertSize(att.size, `Attachment '${att.name}'`);
        const folder = await this.writableDriveFolder(args.parentFolderId);
        const item = await this.uploadToDrive(
          folder,
          String(args.filename || att.name),
          Buffer.from(att.contentBytes, "base64"),
          "rename",
        );
        return {
          status: "filed",
          name: item?.name,
          id: item?.id,
          webUrl: item?.webUrl,
        };
      }

      // --- CALENDAR ---
      case "calendar_list_events":
        return this.client
          .api("/me/calendarView")
          .header("Prefer", `outlook.timezone="${this.tz(args)}"`)
          .query({
            startDateTime: String(args.startDateTime),
            endDateTime: String(args.endDateTime),
          })
          .top(clampTop(args.top, 50, 100))
          .orderby("start/dateTime")
          .select(
            "id,subject,start,end,location,organizer,attendees,isOnlineMeeting,onlineMeeting,isCancelled,showAs",
          )
          .get();

      case "calendar_get_event":
        return this.client
          .api(`/me/events/${seg(args.eventId, "eventId")}`)
          .header("Prefer", `outlook.timezone="${this.tz(args)}"`)
          .get();

      case "calendar_create_event": {
        const attendees = recipients(args.attendees, "attendees") ?? [];
        if (attendees.length)
          this.requireOutbound("Creating a meeting with attendees");
        const tz = this.tz(args);
        const online = args.isOnlineMeeting === true && !this.cfg.isConsumer;
        const ev = await this.client.api("/me/events").post({
          subject: String(args.subject ?? ""),
          body: args.body ? body(args.body) : undefined,
          start: { dateTime: String(args.startDateTime), timeZone: tz },
          end: { dateTime: String(args.endDateTime), timeZone: tz },
          location: args.location
            ? { displayName: String(args.location) }
            : undefined,
          attendees: attendees.map((a) => ({ ...a, type: "required" })),
          ...(online
            ? {
                isOnlineMeeting: true,
                onlineMeetingProvider: "teamsForBusiness",
              }
            : {}),
        });
        return {
          status: "created",
          id: ev.id,
          subject: ev.subject,
          start: ev.start,
          end: ev.end,
          webLink: ev.webLink,
        };
      }

      case "calendar_update_event": {
        if (
          !this.cfg.enableOutbound &&
          (await this.eventHasAttendees(args.eventId))
        ) {
          this.requireOutbound("Changing a meeting that has attendees");
        }
        const tz = this.tz(args);
        const updates: Args = {};
        if (args.subject) updates.subject = String(args.subject);
        if (args.body) updates.body = body(args.body);
        if (args.startDateTime)
          updates.start = {
            dateTime: String(args.startDateTime),
            timeZone: tz,
          };
        if (args.endDateTime)
          updates.end = { dateTime: String(args.endDateTime), timeZone: tz };
        if (args.location)
          updates.location = { displayName: String(args.location) };
        if (!Object.keys(updates).length)
          throw new PolicyError("Nothing to update.");
        const ev = await this.client
          .api(`/me/events/${seg(args.eventId, "eventId")}`)
          .update(updates);
        return {
          status: "updated",
          id: ev?.id ?? args.eventId,
          subject: ev?.subject,
          start: ev?.start,
          end: ev?.end,
        };
      }

      case "calendar_delete_event": {
        if (
          !this.cfg.enableOutbound &&
          (await this.eventHasAttendees(args.eventId))
        ) {
          this.requireOutbound("Deleting a meeting that has attendees");
        }
        await this.client
          .api(`/me/events/${seg(args.eventId, "eventId")}`)
          .delete();
        return { status: "deleted", id: args.eventId };
      }

      // --- ONEDRIVE ---
      case "onedrive_list_root":
        return this.page(
          "/me/drive/root/children",
          args,
          clampTop(args.top, 50, 200),
          "id,name,size,folder,file,webUrl,lastModifiedDateTime",
        );

      case "onedrive_list_folder":
        return this.page(
          `/me/drive/items/${seg(args.folderId, "folderId")}/children`,
          args,
          clampTop(args.top, 50, 200),
          "id,name,size,folder,file,webUrl,lastModifiedDateTime",
        );

      case "onedrive_search_files":
        return this.client
          .api(
            `/me/drive/root/search(q='${encodeURIComponent(odataString(args.query))}')`,
          )
          .top(clampTop(args.top, 25, 200))
          .select(
            "id,name,size,folder,file,webUrl,lastModifiedDateTime,parentReference",
          )
          .get();

      case "onedrive_download_file": {
        const id = seg(args.fileId, "fileId");
        const item = await this.client.api(`/me/drive/items/${id}`).get();
        if (!args.destinationPath) return item;
        const overwrite = args.overwrite === true;
        const dest = resolveWritePath(
          args.destinationPath,
          this.cfg,
          overwrite,
        );
        const data = await this.fetchDriveItem(
          item,
          `/me/drive/items/${id}/content`,
        );
        writeLocalFile(dest, data, overwrite);
        return {
          status: "saved",
          name: item.name,
          savedPath: dest,
          size: data.length,
        };
      }

      case "onedrive_upload_file": {
        const behavior = ["rename", "fail", "replace"].includes(
          args.conflictBehavior,
        )
          ? args.conflictBehavior
          : "rename";
        const data = Buffer.from(String(args.content ?? ""), "utf8");
        if (data.length > SIMPLE_UPLOAD_LIMIT)
          throw new PolicyError("Text uploads are limited to 4 MB.");
        const folder = await this.writableDriveFolder(args.parentFolderId);
        const item = await this.uploadToDrive(
          folder,
          String(args.filename),
          data,
          behavior,
        );
        return {
          status: "uploaded",
          id: item?.id,
          name: item?.name,
          webUrl: item?.webUrl,
        };
      }

      case "onedrive_create_folder": {
        const parent = `${await this.writableDriveFolder(args.parentFolderId)}/children`;
        return this.client.api(parent).post({
          name: String(args.folderName),
          folder: {},
          "@microsoft.graph.conflictBehavior": "rename",
        });
      }

      // --- SHAREPOINT ---
      case "sharepoint_list_sites":
        return this.client
          .api(
            `/sites?search=${encodeURIComponent(String(args.searchQuery || "*"))}`,
          )
          .get();

      case "sharepoint_get_site_details":
        return this.client
          .api(`/sites/${siteSeg(args.siteId)}/drives`)
          .select("id,name,description,webUrl,driveType")
          .get();

      case "sharepoint_list_drive_items": {
        const drive = `/sites/${siteSeg(args.siteId)}/drives/${seg(args.driveId, "driveId")}`;
        const apiPath = args.folderId
          ? `${drive}/items/${seg(args.folderId, "folderId")}/children`
          : `${drive}/root/children`;
        return this.page(
          apiPath,
          args,
          clampTop(args.top, 50, 200),
          "id,name,size,folder,file,webUrl,lastModifiedDateTime",
        );
      }

      case "sharepoint_download_file": {
        const itemPath = `/sites/${siteSeg(args.siteId)}/drives/${seg(args.driveId, "driveId")}/items/${seg(args.fileId, "fileId")}`;
        const item = await this.client.api(itemPath).get();
        if (!args.destinationPath) return item;
        const overwrite = args.overwrite === true;
        const dest = resolveWritePath(
          args.destinationPath,
          this.cfg,
          overwrite,
        );
        const data = await this.fetchDriveItem(item, `${itemPath}/content`);
        writeLocalFile(dest, data, overwrite);
        return {
          status: "saved",
          name: item.name,
          savedPath: dest,
          size: data.length,
          webUrl: item.webUrl,
        };
      }

      case "sharepoint_search_files": {
        const q = encodeURIComponent(odataString(args.query));
        const site = `/sites/${siteSeg(args.siteId)}`;
        const apiPath = args.driveId
          ? `${site}/drives/${seg(args.driveId, "driveId")}/root/search(q='${q}')`
          : `${site}/drive/root/search(q='${q}')`;
        return this.client
          .api(apiPath)
          .top(clampTop(args.top, 25, 200))
          .select(
            "id,name,size,folder,file,webUrl,lastModifiedDateTime,parentReference",
          )
          .get();
      }

      // --- TEAMS ---
      case "teams_list_chats":
        return this.client
          .api("/me/chats")
          .expand("lastMessagePreview")
          .top(clampTop(args.top, 10, 50))
          .select("id,topic,chatType,lastUpdatedDateTime,webUrl")
          .get();

      case "teams_list_chat_messages":
        return this.client
          .api(`/chats/${seg(args.chatId, "chatId")}/messages`)
          .top(clampTop(args.top, 20, 50))
          .orderby("createdDateTime desc")
          .get();

      case "teams_send_chat_message": {
        this.requireOutbound("Sending a Teams message");
        const msg = await this.client
          .api(`/chats/${seg(args.chatId, "chatId")}/messages`)
          .post({ body: body(args.content) });
        return {
          status: "sent",
          id: msg.id,
          createdDateTime: msg.createdDateTime,
        };
      }

      default:
        throw new PolicyError(`Unknown tool '${name}'.`);
    }
  }

  private async healthCheck() {
    const report: Args = {
      profile: this.cfg.profile,
      accountType: this.cfg.isConsumer
        ? "personal (consumers)"
        : "work or school",
      outboundToolsEnabled: this.cfg.enableOutbound,
      requestedScopes: this.auth.grantedScopes,
      writeDirs: this.cfg.writeDirs,
      readDirs: this.cfg.readDirs,
      denyRules: this.cfg.denyPaths.length,
    };
    const me = await this.client
      .api("/me")
      .select("displayName,mail,userPrincipalName,accountEnabled")
      .get();
    report.account = {
      displayName: me.displayName,
      mail: me.mail || me.userPrincipalName,
    };

    try {
      const inbox = await this.client
        .api("/me/mailFolders/inbox")
        .select("totalItemCount,unreadItemCount")
        .get();
      report.mailbox = {
        status: "ok",
        totalItems: inbox.totalItemCount,
        unread: inbox.unreadItemCount,
      };
    } catch (e: any) {
      report.mailbox = {
        status: "unavailable",
        code: errorCode(e),
        detail: describeError(e),
      };
    }

    if (!this.cfg.isConsumer) {
      try {
        const lic = await this.client.api("/me/licenseDetails").get();
        const skus = (lic.value || []).map((l: any) => l.skuPartNumber);
        const exchangeActive = (lic.value || []).some((l: any) =>
          (l.servicePlans || []).some(
            (p: any) =>
              /^EXCHANGE_S_(STANDARD|ENTERPRISE|DESKLESS)/.test(
                p.servicePlanName,
              ) && p.provisioningStatus === "Success",
          ),
        );
        report.licences = { skus, exchangeOnlinePlanActive: exchangeActive };
      } catch (e: any) {
        report.licences = { status: "unavailable", detail: describeError(e) };
      }
    }
    report.verdict =
      report.mailbox.status === "ok"
        ? "Healthy"
        : report.licences?.exchangeOnlinePlanActive === false
          ? "Mailbox unavailable: no active Exchange Online licence on this account. Ask an administrator to reassign it."
          : "Mailbox unavailable: see mailbox.detail.";
    return report;
  }
}
