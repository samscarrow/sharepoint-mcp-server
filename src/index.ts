#!/usr/bin/env node

/**
 * SharePoint MCP Server
 * 
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 */

import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  ListResourcesRequestSchema,
  ReadResourceRequestSchema,
  McpError,
  ErrorCode,
} from "@modelcontextprotocol/sdk/types.js";
import fs from "node:fs";
import path from "node:path";
import JSZip from "jszip";
import { PDFParse } from "pdf-parse";

/**
 * Environment variables required for SharePoint authentication
 */
const { SHAREPOINT_URL, TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;

if (!SHAREPOINT_URL || !TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
  throw new Error(
    "Required environment variables: SHAREPOINT_URL, TENANT_ID, CLIENT_ID, CLIENT_SECRET"
  );
}

const TOKEN_FILE =
  process.env.TOKEN_FILE ||
  path.join(process.env.HOME || "/tmp", ".config", "bay-view-graph", "tokens.json");

const DELEGATED_SCOPES = "Mail.ReadWrite Mail.Send Files.Read Files.Read.All Files.ReadWrite Files.ReadWrite.All Sites.ReadWrite.All Calendars.Read Calendars.ReadBasic Calendars.ReadWrite Calendars.Read.Shared ChannelMessage.Send Tasks.ReadWrite offline_access";

/**
 * Interface for Microsoft Graph API responses
 */
interface GraphResponse {
  value?: any[];
  [key: string]: any;
}

/**
 * SharePoint MCP Server implementation
 * Provides tools and resources for interacting with Microsoft SharePoint via Microsoft Graph API
 */
class SharePointServer {
  private server: Server;
  private accessToken: string | null = null;
  private tokenExpiry: number = 0;
  private userAccessToken: string | null = null;
  private userTokenExpiry: number = 0;

  constructor() {
    this.server = new Server(
      {
        name: "sharepoint-mcp-server",
        version: "0.1.0",
      },
      {
        capabilities: {
          tools: {},
          resources: {},
        },
      }
    );

    this.setupHandlers();
    this.setupErrorHandling();
  }

  /**
   * Get access token for Microsoft Graph API
   */
  private async getAccessToken(): Promise<string> {
    if (this.accessToken && Date.now() < this.tokenExpiry) {
      return this.accessToken;
    }

    const tenantId = TENANT_ID!;
    const clientId = CLIENT_ID!;
    const clientSecret = CLIENT_SECRET!;

    const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
    const params = new URLSearchParams({
      client_id: clientId,
      client_secret: clientSecret,
      scope: "https://graph.microsoft.com/.default",
      grant_type: "client_credentials",
    });

    try {
      const response = await fetch(tokenUrl, {
        method: "POST",
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
        },
        body: params,
      });

      if (!response.ok) {
        throw new Error(`Token request failed: ${response.status} ${response.statusText}`);
      }

      const data = await response.json();
      this.accessToken = data.access_token;
      this.tokenExpiry = Date.now() + (data.expires_in * 1000) - 60000; // Refresh 1 minute early

      return this.accessToken!;
    } catch (error) {
      throw new Error(`Failed to get access token: ${error}`);
    }
  }

  /**
   * Make authenticated request to Microsoft Graph API
   */
  private async graphRequest(endpoint: string, method: string = "GET", body?: any, timeoutMs = 20000): Promise<any> {
    const token = await this.getAccessToken();
    const url = endpoint.startsWith("https://") ? endpoint : `https://graph.microsoft.com/v1.0${endpoint}`;

    const headers: Record<string, string> = {
      "Authorization": `Bearer ${token}`,
      "Content-Type": "application/json",
    };

    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);

    const options: RequestInit = {
      method,
      headers,
      signal: controller.signal,
    };

    if (body && method !== "GET") {
      options.body = JSON.stringify(body);
    }

    try {
      const response = await fetch(url, options);
      clearTimeout(timer);

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Graph API request failed: ${response.status} ${response.statusText} - ${errorText}`);
      }

      return await response.json();
    } catch (error: any) {
      clearTimeout(timer);
      if (error?.name === "AbortError") {
        throw new Error(`Graph API request timed out after ${timeoutMs}ms: ${endpoint}`);
      }
      throw new Error(`Graph API request error: ${error}`);
    }
  }

  /**
   * Get delegated user access token (for email operations via /me/ endpoints).
   * Reads refresh token from disk and refreshes as needed.
   */
  private async getUserAccessToken(): Promise<string> {
    if (this.userAccessToken && Date.now() < this.userTokenExpiry) {
      return this.userAccessToken;
    }

    if (!fs.existsSync(TOKEN_FILE)) {
      throw new Error(
        `No user token found at ${TOKEN_FILE}. Run "npm run oauth-login" to authenticate.`
      );
    }

    const tokens = JSON.parse(fs.readFileSync(TOKEN_FILE, "utf-8"));

    // Use cached access token if still valid
    if (tokens.access_token && Date.now() < tokens.expires_at) {
      this.userAccessToken = tokens.access_token;
      this.userTokenExpiry = tokens.expires_at;
      return this.userAccessToken!;
    }

    // Refresh
    if (!tokens.refresh_token) {
      throw new Error('No refresh token found. Run "npm run oauth-login" to re-authenticate.');
    }

    const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
    const params = new URLSearchParams({
      client_id: CLIENT_ID!,
      client_secret: CLIENT_SECRET!,
      refresh_token: tokens.refresh_token,
      grant_type: "refresh_token",
      scope: DELEGATED_SCOPES,
    });

    const response = await fetch(tokenUrl, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: params,
    });

    if (!response.ok) {
      const errText = await response.text();
      throw new Error(
        `Token refresh failed (${response.status}): ${errText}. Run "npm run oauth-login" to re-authenticate.`
      );
    }

    const data: any = await response.json();

    // Persist updated tokens
    const updated = {
      access_token: data.access_token,
      refresh_token: data.refresh_token || tokens.refresh_token,
      expires_at: Date.now() + data.expires_in * 1000 - 60000,
    };
    fs.writeFileSync(TOKEN_FILE, JSON.stringify(updated, null, 2), { mode: 0o600 });

    this.userAccessToken = data.access_token;
    this.userTokenExpiry = updated.expires_at;
    return this.userAccessToken!;
  }

  /**
   * Make authenticated request to Microsoft Graph API using the delegated user token.
   */
  private async graphRequestAsUser(
    endpoint: string,
    method: string = "GET",
    body?: any,
    timeoutMs = 20000
  ): Promise<any> {
    const token = await this.getUserAccessToken();
    const url = endpoint.startsWith("https://") ? endpoint : `https://graph.microsoft.com/v1.0${endpoint}`;

    const headers: Record<string, string> = {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    };

    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);

    const options: RequestInit = { method, headers, signal: controller.signal };
    if (body && method !== "GET") {
      options.body = JSON.stringify(body);
    }

    try {
      const response = await fetch(url, options);
      clearTimeout(timer);

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(
          `Graph API request failed: ${response.status} ${response.statusText} - ${errorText}`
        );
      }

      // sendMail returns 202 with no body
      const ct = response.headers.get("content-type") || "";
      if (response.status === 202 || !ct.includes("application/json")) {
        return { success: true };
      }

      return response.json();
    } catch (error: any) {
      clearTimeout(timer);
      if (error?.name === "AbortError") {
        throw new Error(`Graph API request timed out after ${timeoutMs}ms: ${endpoint}`);
      }
      throw new Error(`Graph API request error: ${error}`);
    }
  }

  /**
   * Setup error handling for the server
   */
  private setupErrorHandling(): void {
    this.server.onerror = (error) => console.error("[MCP Error]", error);
    
    process.on("SIGINT", async () => {
      await this.server.close();
      process.exit(0);
    });
  }

  /**
   * Setup all request handlers for tools and resources
   */
  private setupHandlers(): void {
    this.setupToolHandlers();
    this.setupResourceHandlers();
  }

  /**
   * Setup tool handlers for SharePoint operations
   */
  private setupToolHandlers(): void {
    this.server.setRequestHandler(ListToolsRequestSchema, async () => ({
      tools: [
        {
          name: "search_files",
          description:
            "Search for files and documents. scope=tenant searches all SharePoint sites (app auth); scope=me searches your OneDrive (delegated auth). Pass a returned webUrl as get_file_content.filePath to read the file.",
          inputSchema: {
            type: "object",
            properties: {
              query: {
                type: "string",
                description: "The search query string",
              },
              limit: {
                type: "number",
                description: "Maximum number of results to return (default: 10)",
                default: 10,
              },
              scope: {
                type: "string",
                enum: ["tenant", "me"],
                description: "Search scope: tenant (all SharePoint, default) or me (your OneDrive)",
                default: "tenant",
              },
            },
            required: ["query"],
          },
        },
        {
          name: "list_sites",
          description: "List SharePoint sites accessible to the application",
          inputSchema: {
            type: "object",
            properties: {
              search: {
                type: "string",
                description: "Optional search term to filter sites",
              },
            },
          },
        },
        {
          name: "get_site_info",
          description: "Get detailed information about a specific SharePoint site",
          inputSchema: {
            type: "object",
            properties: {
              siteUrl: {
                type: "string",
                description: "The SharePoint site URL (e.g., https://tenant.sharepoint.com/sites/sitename)",
              },
            },
            required: ["siteUrl"],
          },
        },
        {
          name: "list_site_drives",
          description:
            "List document libraries (drives). scope=tenant lists drives in a SharePoint site; scope=me lists your OneDrive drives.",
          inputSchema: {
            type: "object",
            properties: {
              siteUrl: {
                type: "string",
                description: "The SharePoint site URL (required when scope=tenant)",
              },
              scope: {
                type: "string",
                enum: ["tenant", "me"],
                description: "Scope: tenant (SharePoint site drives, default) or me (your OneDrive drives)",
                default: "tenant",
              },
            },
          },
        },
        {
          name: "list_drive_items",
          description:
            "List files and folders. scope=tenant lists from a SharePoint site drive; scope=me lists from your OneDrive.",
          inputSchema: {
            type: "object",
            properties: {
              siteUrl: {
                type: "string",
                description: "The SharePoint site URL (required when scope=tenant)",
              },
              driveId: {
                type: "string",
                description: "The drive ID (optional, uses default drive if not specified)",
              },
              folderPath: {
                type: "string",
                description: "Optional folder path to list items from (default: root)",
              },
              scope: {
                type: "string",
                enum: ["tenant", "me"],
                description: "Scope: tenant (SharePoint, default) or me (your OneDrive)",
                default: "tenant",
              },
              nextLink: {
                type: "string",
                description: "Pagination cursor — pass the nextLink value from a previous response to fetch the next page",
              },
            },
          },
        },
        {
          name: "get_file_content",
          description:
            "Get the content of a file. .pdf, .docx, .xlsx, and .pptx are automatically converted to text (xlsx as TSV per sheet). Other binary formats are returned raw and may be unreadable. scope=tenant reads from SharePoint; scope=me reads from your OneDrive.",
          inputSchema: {
            type: "object",
            properties: {
              siteUrl: {
                type: "string",
                description: "The SharePoint site URL (required when scope=tenant)",
              },
              filePath: {
                type: "string",
                description: "The drive-relative path to the file, or a full SharePoint/OneDrive webUrl returned by search_files",
              },
              driveId: {
                type: "string",
                description: "The drive ID (optional, uses default drive if not specified)",
              },
              scope: {
                type: "string",
                enum: ["tenant", "me"],
                description: "Scope: tenant (SharePoint, default) or me (your OneDrive)",
                default: "tenant",
              },
            },
            required: ["filePath"],
          },
        },
        {
          name: "list_emails",
          description:
            "List recent emails from your mailbox (delegated auth). Returns subject, from, date, and preview.",
          inputSchema: {
            type: "object",
            properties: {
              folder: {
                type: "string",
                description:
                  "Mail folder to read from (default: inbox). Options: inbox, sentitems, drafts, deleteditems, archive",
                default: "inbox",
              },
              top: {
                type: "number",
                description: "Number of emails to return (default: 10, max: 50)",
                default: 10,
              },
              filter: {
                type: "string",
                description:
                  'OData filter expression (e.g. "from/emailAddress/address eq \'someone@example.com\'" or "isRead eq false")',
              },
              nextLink: {
                type: "string",
                description: "Pagination cursor — pass the nextLink value from a previous response to fetch the next page",
              },
            },
          },
        },
        {
          name: "get_email",
          description: "Get the full content of a specific email by message ID",
          inputSchema: {
            type: "object",
            properties: {
              messageId: {
                type: "string",
                description: "The email message ID",
              },
            },
            required: ["messageId"],
          },
        },
        {
          name: "search_emails",
          description: "Search emails by keyword in your mailbox",
          inputSchema: {
            type: "object",
            properties: {
              query: {
                type: "string",
                description: "Search query (searches subject, body, sender, recipients)",
              },
              top: {
                type: "number",
                description: "Number of results to return (default: 10, max: 50)",
                default: 10,
              },
              nextLink: {
                type: "string",
                description: "Pagination cursor — pass the nextLink value from a previous response to fetch the next page",
              },
            },
            required: ["query"],
          },
        },
        {
          name: "send_email",
          description: "Send an email from your mailbox",
          inputSchema: {
            type: "object",
            properties: {
              to: {
                type: "array",
                items: { type: "string" },
                description: "Recipient email addresses",
              },
              subject: {
                type: "string",
                description: "Email subject line",
              },
              body: {
                type: "string",
                description: "Email body (HTML supported)",
              },
              cc: {
                type: "array",
                items: { type: "string" },
                description: "CC recipient email addresses",
              },
              bcc: {
                type: "array",
                items: { type: "string" },
                description: "BCC recipient email addresses",
              },
              bodyType: {
                type: "string",
                enum: ["HTML", "Text"],
                description: "Body content type (default: HTML)",
                default: "HTML",
              },
            },
            required: ["to", "subject", "body"],
          },
        },
        {
          name: "reply_email",
          description: "Reply to an email by message ID",
          inputSchema: {
            type: "object",
            properties: {
              messageId: {
                type: "string",
                description: "The message ID to reply to",
              },
              comment: {
                type: "string",
                description: "Reply body (HTML supported)",
              },
              replyAll: {
                type: "boolean",
                description: "Reply to all recipients (default: false)",
                default: false,
              },
            },
            required: ["messageId", "comment"],
          },
        },
        {
          name: "list_mail_folders",
          description: "List mail folders in your mailbox (inbox, custom folders, etc.)",
          inputSchema: {
            type: "object",
            properties: {},
          },
        },
        {
          name: "move_email",
          description: "Move an email to a different folder",
          inputSchema: {
            type: "object",
            properties: {
              messageId: {
                type: "string",
                description: "The message ID to move",
              },
              destinationFolder: {
                type: "string",
                description:
                  "Destination folder name (e.g. inbox, archive, deleteditems) or folder ID",
              },
            },
            required: ["messageId", "destinationFolder"],
          },
        },
        {
          name: "mark_email_read",
          description: "Mark an email as read or unread",
          inputSchema: {
            type: "object",
            properties: {
              messageId: {
                type: "string",
                description: "The message ID",
              },
              isRead: {
                type: "boolean",
                description: "true to mark as read, false to mark as unread",
              },
            },
            required: ["messageId", "isRead"],
          },
        },
        {
          name: "delete_email",
          description: "Delete an email (moves to Deleted Items)",
          inputSchema: {
            type: "object",
            properties: {
              messageId: {
                type: "string",
                description: "The message ID to delete",
              },
            },
            required: ["messageId"],
          },
        },
        {
          name: "get_email_attachments",
          description:
            "List attachments on an email, including inline images. Returns name, contentType, isInline, contentId, and base64 contentBytes.",
          inputSchema: {
            type: "object",
            properties: {
              messageId: {
                type: "string",
                description: "The message ID",
              },
              saveSignature: {
                type: "boolean",
                description:
                  "If true, saves the first inline image as the default signature to ~/.config/bay-view-graph/signature.png",
                default: false,
              },
            },
            required: ["messageId"],
          },
        },
        {
          name: "list_calendar_events",
          description: "List calendar events from the user's Outlook calendar within a date range.",
          inputSchema: {
            type: "object",
            properties: {
              start: {
                type: "string",
                description: "Start datetime in ISO 8601 format (e.g. 2026-06-01T00:00:00). Defaults to today.",
              },
              end: {
                type: "string",
                description: "End datetime in ISO 8601 format (e.g. 2026-09-01T00:00:00). Defaults to 90 days from start.",
              },
              top: {
                type: "number",
                description: "Maximum number of events to return (default: 50, max: 100)",
                default: 50,
              },
              calendarId: {
                type: "string",
                description: "Optional calendar ID. Omit to use the default calendar.",
              },
              nextLink: {
                type: "string",
                description: "Pagination cursor — pass the nextLink value from a previous response to fetch the next page",
              },
            },
          },
        },
        {
          name: "get_calendar_event",
          description: "Get a specific calendar event by ID.",
          inputSchema: {
            type: "object",
            properties: {
              eventId: {
                type: "string",
                description: "The event ID",
              },
            },
            required: ["eventId"],
          },
        },
        {
          name: "create_calendar_event",
          description: "Create a new calendar event.",
          inputSchema: {
            type: "object",
            properties: {
              subject: {
                type: "string",
                description: "Event title",
              },
              start: {
                type: "string",
                description: "Start datetime in ISO 8601 format (e.g. 2026-07-29T14:00:00)",
              },
              end: {
                type: "string",
                description: "End datetime in ISO 8601 format",
              },
              timeZone: {
                type: "string",
                description: "IANA time zone (e.g. America/Detroit). Defaults to America/Detroit.",
                default: "America/Detroit",
              },
              location: {
                type: "string",
                description: "Location string",
              },
              body: {
                type: "string",
                description: "Event body/description (plain text)",
              },
              attendees: {
                type: "array",
                description: "List of attendee email addresses",
                items: { type: "string" },
              },
              calendarId: {
                type: "string",
                description: "Optional calendar ID. Omit to use the default calendar.",
              },
            },
            required: ["subject", "start", "end"],
          },
        },
        {
          name: "update_calendar_event",
          description: "Update an existing calendar event.",
          inputSchema: {
            type: "object",
            properties: {
              eventId: {
                type: "string",
                description: "The event ID to update",
              },
              subject: {
                type: "string",
                description: "Event title",
              },
              start: {
                type: "string",
                description: "Start datetime in ISO 8601 format",
              },
              end: {
                type: "string",
                description: "End datetime in ISO 8601 format",
              },
              timeZone: {
                type: "string",
                description: "IANA time zone (default: America/Detroit)",
                default: "America/Detroit",
              },
              location: {
                type: "string",
                description: "Location string",
              },
              body: {
                type: "string",
                description: "Event body/description",
              },
              attendees: {
                type: "array",
                items: { type: "string" },
                description: "List of attendee email addresses",
              },
              calendarId: {
                type: "string",
                description: "Optional calendar ID",
              },
            },
            required: ["eventId"],
          },
        },
        {
          name: "delete_calendar_event",
          description: "Delete a calendar event.",
          inputSchema: {
            type: "object",
            properties: {
              eventId: {
                type: "string",
                description: "The event ID to delete",
              },
              calendarId: {
                type: "string",
                description: "Optional calendar ID",
              },
            },
            required: ["eventId"],
          },
        },
        {
          name: "respond_to_event",
          description:
            "Accept, decline, or tentatively accept a calendar event invitation.",
          inputSchema: {
            type: "object",
            properties: {
              eventId: {
                type: "string",
                description: "The event ID to respond to",
              },
              response: {
                type: "string",
                enum: ["accept", "decline", "tentative"],
                description: "Response type",
              },
              comment: {
                type: "string",
                description: "Optional comment to include with the response",
              },
              sendResponse: {
                type: "boolean",
                description: "Whether to send response to organizer (default: true)",
                default: true,
              },
            },
            required: ["eventId", "response"],
          },
        },
        {
          name: "list_lists",
          description: "List all SharePoint lists in a site.",
          inputSchema: {
            type: "object",
            properties: {
              siteUrl: {
                type: "string",
                description: "SharePoint site URL",
              },
              filter: {
                type: "string",
                description: "Optional OData $filter expression",
              },
            },
            required: ["siteUrl"],
          },
        },
        {
          name: "get_list_items",
          description: "Get items from a SharePoint list with optional filtering and pagination.",
          inputSchema: {
            type: "object",
            properties: {
              siteUrl: {
                type: "string",
                description: "SharePoint site URL",
              },
              listId: {
                type: "string",
                description: "List ID or name",
              },
              filter: {
                type: "string",
                description: "Optional OData $filter expression on field values",
              },
              select: {
                type: "string",
                description: "Comma-separated field names to retrieve (default: all)",
              },
              top: {
                type: "number",
                description: "Number of items to return (default: 50, max: 200)",
                default: 50,
              },
              nextLink: {
                type: "string",
                description: "Pagination cursor — pass the nextLink value from a previous response to fetch the next page",
              },
            },
            required: ["siteUrl", "listId"],
          },
        },
        {
          name: "get_list_item",
          description: "Get a specific item from a SharePoint list.",
          inputSchema: {
            type: "object",
            properties: {
              siteUrl: {
                type: "string",
                description: "SharePoint site URL",
              },
              listId: {
                type: "string",
                description: "List ID or name",
              },
              itemId: {
                type: "string",
                description: "Item ID (numeric)",
              },
            },
            required: ["siteUrl", "listId", "itemId"],
          },
        },
        {
          name: "create_list_item",
          description: "Create a new item in a SharePoint list.",
          inputSchema: {
            type: "object",
            properties: {
              siteUrl: {
                type: "string",
                description: "SharePoint site URL",
              },
              listId: {
                type: "string",
                description: "List ID or name",
              },
              fields: {
                type: "object",
                description: "Object mapping field names to values (e.g. {Title: 'My Item', Status: 'Active'})",
              },
            },
            required: ["siteUrl", "listId", "fields"],
          },
        },
        {
          name: "update_list_item",
          description: "Update fields in a SharePoint list item.",
          inputSchema: {
            type: "object",
            properties: {
              siteUrl: {
                type: "string",
                description: "SharePoint site URL",
              },
              listId: {
                type: "string",
                description: "List ID or name",
              },
              itemId: {
                type: "string",
                description: "Item ID (numeric)",
              },
              fields: {
                type: "object",
                description: "Object mapping field names to new values",
              },
            },
            required: ["siteUrl", "listId", "itemId", "fields"],
          },
        },
        {
          name: "delete_list_item",
          description: "Delete an item from a SharePoint list.",
          inputSchema: {
            type: "object",
            properties: {
              siteUrl: {
                type: "string",
                description: "SharePoint site URL",
              },
              listId: {
                type: "string",
                description: "List ID or name",
              },
              itemId: {
                type: "string",
                description: "Item ID (numeric)",
              },
            },
            required: ["siteUrl", "listId", "itemId"],
          },
        },
        {
          name: "upload_file",
          description:
            "Upload or overwrite a file in SharePoint or OneDrive. Provide content as plain text (contentEncoding=text) or base64 (contentEncoding=base64).",
          inputSchema: {
            type: "object",
            properties: {
              filePath: {
                type: "string",
                description: "Destination path including filename (e.g. Documents/report.txt)",
              },
              content: {
                type: "string",
                description: "File content as UTF-8 text or base64 string",
              },
              contentEncoding: {
                type: "string",
                enum: ["text", "base64"],
                description: "Encoding of the content field (default: text)",
                default: "text",
              },
              contentType: {
                type: "string",
                description: "MIME type (default: text/plain)",
                default: "text/plain",
              },
              siteUrl: {
                type: "string",
                description: "SharePoint site URL (required when scope=tenant)",
              },
              driveId: {
                type: "string",
                description: "Drive ID (optional, uses default drive if omitted)",
              },
              scope: {
                type: "string",
                enum: ["tenant", "me"],
                description: "Scope: tenant (SharePoint, default) or me (OneDrive)",
                default: "tenant",
              },
              conflictBehavior: {
                type: "string",
                enum: ["replace", "rename", "fail"],
                description: "What to do if the file already exists (default: replace)",
                default: "replace",
              },
            },
            required: ["filePath", "content"],
          },
        },
        {
          name: "create_folder",
          description: "Create a folder in SharePoint or OneDrive.",
          inputSchema: {
            type: "object",
            properties: {
              folderName: {
                type: "string",
                description: "Name of the new folder",
              },
              parentPath: {
                type: "string",
                description: "Parent folder path (omit or use empty string for root)",
                default: "",
              },
              siteUrl: {
                type: "string",
                description: "SharePoint site URL (required when scope=tenant)",
              },
              driveId: {
                type: "string",
                description: "Drive ID (optional, uses default drive if omitted)",
              },
              scope: {
                type: "string",
                enum: ["tenant", "me"],
                description: "Scope: tenant (SharePoint, default) or me (OneDrive)",
                default: "tenant",
              },
              conflictBehavior: {
                type: "string",
                enum: ["rename", "replace", "fail"],
                description: "Conflict behavior if folder exists (default: rename)",
                default: "rename",
              },
            },
            required: ["folderName"],
          },
        },
        {
          name: "delete_drive_item",
          description: "Delete a file or folder from SharePoint or OneDrive.",
          inputSchema: {
            type: "object",
            properties: {
              itemPath: {
                type: "string",
                description: "Path to the file or folder to delete",
              },
              siteUrl: {
                type: "string",
                description: "SharePoint site URL (required when scope=tenant)",
              },
              driveId: {
                type: "string",
                description: "Drive ID (optional, uses default drive if omitted)",
              },
              scope: {
                type: "string",
                enum: ["tenant", "me"],
                description: "Scope: tenant (SharePoint, default) or me (OneDrive)",
                default: "tenant",
              },
            },
            required: ["itemPath"],
          },
        },
        {
          name: "move_rename_drive_item",
          description:
            "Move and/or rename a file or folder in SharePoint or OneDrive. Provide newName to rename, newParentPath to move, or both.",
          inputSchema: {
            type: "object",
            properties: {
              itemPath: {
                type: "string",
                description: "Current path of the item",
              },
              newName: {
                type: "string",
                description: "New filename (optional)",
              },
              newParentPath: {
                type: "string",
                description: "Destination folder path to move the item into (optional)",
              },
              siteUrl: {
                type: "string",
                description: "SharePoint site URL (required when scope=tenant)",
              },
              driveId: {
                type: "string",
                description: "Drive ID (optional, uses default drive if omitted)",
              },
              scope: {
                type: "string",
                enum: ["tenant", "me"],
                description: "Scope: tenant (SharePoint, default) or me (OneDrive)",
                default: "tenant",
              },
            },
            required: ["itemPath"],
          },
        },
      ],
    }));

    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      try {
        switch (request.params.name) {
          case "search_files":
            return await this.handleSearchFiles(request.params.arguments);
          case "list_sites":
            return await this.handleListSites(request.params.arguments);
          case "get_site_info":
            return await this.handleGetSiteInfo(request.params.arguments);
          case "list_site_drives":
            return await this.handleListSiteDrives(request.params.arguments);
          case "list_drive_items":
            return await this.handleListDriveItems(request.params.arguments);
          case "get_file_content":
            return await this.handleGetFileContent(request.params.arguments);
          case "list_emails":
            return await this.handleListEmails(request.params.arguments);
          case "get_email":
            return await this.handleGetEmail(request.params.arguments);
          case "search_emails":
            return await this.handleSearchEmails(request.params.arguments);
          case "send_email":
            return await this.handleSendEmail(request.params.arguments);
          case "reply_email":
            return await this.handleReplyEmail(request.params.arguments);
          case "list_mail_folders":
            return await this.handleListMailFolders(request.params.arguments);
          case "move_email":
            return await this.handleMoveEmail(request.params.arguments);
          case "mark_email_read":
            return await this.handleMarkEmailRead(request.params.arguments);
          case "delete_email":
            return await this.handleDeleteEmail(request.params.arguments);
          case "get_email_attachments":
            return await this.handleGetEmailAttachments(request.params.arguments);
          case "list_calendar_events":
            return await this.handleListCalendarEvents(request.params.arguments);
          case "get_calendar_event":
            return await this.handleGetCalendarEvent(request.params.arguments);
          case "create_calendar_event":
            return await this.handleCreateCalendarEvent(request.params.arguments);
          case "update_calendar_event":
            return await this.handleUpdateCalendarEvent(request.params.arguments);
          case "delete_calendar_event":
            return await this.handleDeleteCalendarEvent(request.params.arguments);
          case "respond_to_event":
            return await this.handleRespondToEvent(request.params.arguments);
          case "list_lists":
            return await this.handleListLists(request.params.arguments);
          case "get_list_items":
            return await this.handleGetListItems(request.params.arguments);
          case "get_list_item":
            return await this.handleGetListItem(request.params.arguments);
          case "create_list_item":
            return await this.handleCreateListItem(request.params.arguments);
          case "update_list_item":
            return await this.handleUpdateListItem(request.params.arguments);
          case "delete_list_item":
            return await this.handleDeleteListItem(request.params.arguments);
          case "upload_file":
            return await this.handleUploadFile(request.params.arguments);
          case "create_folder":
            return await this.handleCreateFolder(request.params.arguments);
          case "delete_drive_item":
            return await this.handleDeleteDriveItem(request.params.arguments);
          case "move_rename_drive_item":
            return await this.handleMoveRenameDriveItem(request.params.arguments);
          default:
            throw new McpError(
              ErrorCode.MethodNotFound,
              `Unknown tool: ${request.params.name}`
            );
        }
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : String(error);
        throw new McpError(ErrorCode.InternalError, `SharePoint operation failed: ${errorMessage}`);
      }
    });
  }

  /**
   * Setup resource handlers for SharePoint resources
   */
  private setupResourceHandlers(): void {
    this.server.setRequestHandler(ListResourcesRequestSchema, async () => {
      try {
        const response = await this.graphRequest("/sites?$select=id,displayName,name,webUrl");
        const sites = response.value || [];
        
        return {
          resources: sites.map((site: any) => ({
            uri: `sharepoint://site/${site.id}`,
            mimeType: "application/json",
            name: site.displayName || site.name,
            description: `SharePoint site: ${site.displayName || site.name} (${site.webUrl})`,
          })),
        };
      } catch (error) {
        console.error("Error listing resources:", error);
        return { resources: [] };
      }
    });

    this.server.setRequestHandler(ReadResourceRequestSchema, async (request) => {
      const url = new URL(request.params.uri);
      
      if (url.protocol === "sharepoint:" && url.pathname.startsWith("/site/")) {
        const siteId = url.pathname.replace("/site/", "");
        try {
          const site = await this.graphRequest(`/sites/${siteId}`);
          return {
            contents: [{
              uri: request.params.uri,
              mimeType: "application/json",
              text: JSON.stringify(site, null, 2),
            }],
          };
        } catch (error) {
          throw new McpError(ErrorCode.InternalError, `Failed to read site resource: ${error}`);
        }
      }
      
      throw new McpError(ErrorCode.InvalidParams, `Unsupported resource URI: ${request.params.uri}`);
    });
  }

  /**
   * Extract site ID from SharePoint URL
   */
  private async getSiteIdFromUrl(siteUrl: string): Promise<string> {
    try {
      const url = new URL(siteUrl);
      const hostname = url.hostname;
      const pathname = url.pathname;
      
      const response = await this.graphRequest(`/sites/${hostname}:${pathname}`);
      return response.id;
    } catch (error) {
      throw new Error(`Failed to get site ID from URL ${siteUrl}: ${error}`);
    }
  }

  private isHttpUrl(value: string): boolean {
    try {
      const parsed = new URL(value);
      return parsed.protocol === "https:" || parsed.protocol === "http:";
    } catch {
      return false;
    }
  }

  private encodeSharingUrl(webUrl: string): string {
    return `u!${Buffer.from(webUrl)
      .toString("base64")
      .replace(/=+$/g, "")
      .replace(/\+/g, "-")
      .replace(/\//g, "_")}`;
  }

  private pathForExtension(filePath: string): string {
    if (!this.isHttpUrl(filePath)) {
      return filePath;
    }

    return new URL(filePath).pathname;
  }

  private decodeXmlEntities(s: string): string {
    return s
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(Number(n)))
      .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCharCode(parseInt(h, 16)))
      .replace(/&amp;/g, "&");
  }

  private xmlInnerText(xml: string): string {
    return this.decodeXmlEntities(xml.replace(/<[^>]+>/g, ""));
  }

  private colLetterToIndex(col: string): number {
    let n = 0;
    for (const ch of col) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  }

  private async extractXlsxText(buf: ArrayBuffer): Promise<string> {
    const zip = await JSZip.loadAsync(buf);

    const sharedStrings: string[] = [];
    const ssFile = zip.file("xl/sharedStrings.xml");
    if (ssFile) {
      const xml = await ssFile.async("string");
      const siMatches = xml.match(/<si\b[^>]*>[\s\S]*?<\/si>/g) || [];
      for (const si of siMatches) {
        const parts = [...si.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((m) =>
          this.decodeXmlEntities(m[1])
        );
        sharedStrings.push(parts.join(""));
      }
    }

    const wbFile = zip.file("xl/workbook.xml");
    const wbRelsFile = zip.file("xl/_rels/workbook.xml.rels");
    type SheetEntry = { name: string; target: string };
    const sheets: SheetEntry[] = [];
    if (wbFile && wbRelsFile) {
      const wbXml = await wbFile.async("string");
      const relsXml = await wbRelsFile.async("string");
      const relMap = new Map<string, string>();
      for (const m of relsXml.matchAll(/<Relationship\b[^>]*\/>/g)) {
        const tag = m[0];
        const id = /Id="([^"]+)"/.exec(tag)?.[1];
        const target = /Target="([^"]+)"/.exec(tag)?.[1];
        if (id && target) relMap.set(id, target);
      }
      for (const m of wbXml.matchAll(/<sheet\b[^>]*\/>/g)) {
        const tag = m[0];
        const name = /name="([^"]+)"/.exec(tag)?.[1] ?? "Sheet";
        const rid = /r:id="([^"]+)"/.exec(tag)?.[1];
        const rel = rid ? relMap.get(rid) : undefined;
        if (rel) {
          const target = rel.startsWith("/") ? rel.slice(1) : `xl/${rel.replace(/^\.?\//, "")}`;
          sheets.push({ name: this.decodeXmlEntities(name), target });
        }
      }
    }

    if (sheets.length === 0) {
      const fallback = zip.file(/^xl\/worksheets\/sheet\d+\.xml$/);
      for (const f of fallback) sheets.push({ name: f.name.split("/").pop() ?? f.name, target: f.name });
    }

    const out: string[] = [];
    for (const sh of sheets) {
      const f = zip.file(sh.target);
      if (!f) continue;
      const xml = await f.async("string");
      out.push(`=== Sheet: ${sh.name} ===`);
      const rows = xml.match(/<row\b[^>]*>[\s\S]*?<\/row>/g) || [];
      for (const row of rows) {
        const cells = row.match(/<c\b[^>]*(?:\/>|>[\s\S]*?<\/c>)/g) || [];
        const rowOut: string[] = [];
        let lastIdx = -1;
        for (const cell of cells) {
          const ref = /r="([A-Z]+)\d+"/.exec(cell)?.[1];
          const idx = ref ? this.colLetterToIndex(ref) : lastIdx + 1;
          while (rowOut.length < idx) rowOut.push("");
          const t = /t="([^"]+)"/.exec(cell)?.[1] ?? "n";
          let value = "";
          if (t === "inlineStr") {
            const parts = [...cell.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((m) =>
              this.decodeXmlEntities(m[1])
            );
            value = parts.join("");
          } else {
            const v = /<v[^>]*>([\s\S]*?)<\/v>/.exec(cell)?.[1];
            if (v != null) {
              if (t === "s") {
                const i = Number(v);
                value = sharedStrings[i] ?? "";
              } else if (t === "b") {
                value = v === "1" ? "TRUE" : "FALSE";
              } else {
                value = this.decodeXmlEntities(v);
              }
            }
          }
          rowOut.push(value);
          lastIdx = idx;
        }
        out.push(rowOut.join("\t").replace(/\t+$/, ""));
      }
      out.push("");
    }
    return out.join("\n").trim();
  }

  private async extractPptxText(buf: ArrayBuffer): Promise<string> {
    const zip = await JSZip.loadAsync(buf);
    const slideFiles = zip.file(/^ppt\/slides\/slide\d+\.xml$/);
    slideFiles.sort((a, b) => {
      const na = Number(/slide(\d+)\.xml$/.exec(a.name)?.[1] ?? 0);
      const nb = Number(/slide(\d+)\.xml$/.exec(b.name)?.[1] ?? 0);
      return na - nb;
    });
    const out: string[] = [];
    for (const f of slideFiles) {
      const xml = await f.async("string");
      const n = /slide(\d+)\.xml$/.exec(f.name)?.[1] ?? "?";
      const parts = [...xml.matchAll(/<a:t[^>]*>([\s\S]*?)<\/a:t>/g)].map((m) =>
        this.decodeXmlEntities(m[1])
      );
      out.push(`=== Slide ${n} ===`);
      out.push(parts.join("\n"));
      out.push("");
    }
    return out.join("\n").trim();
  }

  /**
   * Handle search files tool request
   */
  private async handleSearchFiles(args: any) {
    const query = args?.query;
    const limit = args?.limit || 10;
    const scope = args?.scope || "tenant";

    if (typeof query !== "string") {
      throw new McpError(ErrorCode.InvalidParams, "Query parameter must be a string");
    }

    try {
      if (scope === "me") {
        const endpoint = `/me/drive/root/search(q='${encodeURIComponent(query)}')?$top=${limit}`;
        const response = await this.graphRequestAsUser(endpoint);
        const items = (response.value || []).map((item: any) => ({
          id: item.id,
          name: item.name,
          webUrl: item.webUrl,
          size: item.size,
          lastModified: item.lastModifiedDateTime,
          path: item.parentReference?.path,
          mimeType: item.file?.mimeType,
        }));
        return {
          content: [{ type: "text", text: JSON.stringify(items, null, 2) }],
        };
      }

      // tenant scope — search API returns nested hitsContainers; flatten to essentials
      const searchRequest = {
        requests: [{
          entityTypes: ["driveItem"],
          query: { queryString: query },
          region: "US",
          size: limit,
        }],
      };

      const searchResults = await this.graphRequest("/search/query", "POST", searchRequest);

      // Flatten Graph Search response to minimal items
      const hits: any[] = [];
      for (const container of searchResults.value || []) {
        for (const hit of container.hitsContainers || []) {
          for (const h of hit.hits || []) {
            const r = h.resource || {};
            hits.push({
              name: r.name,
              webUrl: r.webUrl,
              size: r.size,
              lastModified: r.lastModifiedDateTime,
              path: r.parentReference?.path || r.parentReference?.name,
              siteTitle: r.parentReference?.siteId ? undefined : undefined,
              mimeType: r.file?.mimeType,
              hitHighlights: h.hitHighlights?.map((hl: any) => hl.value).filter(Boolean),
            });
          }
        }
      }

      return {
        content: [{ type: "text", text: JSON.stringify(hits, null, 2) }],
      };
    } catch (error) {
      throw new Error(`Search failed: ${error}`);
    }
  }

  /**
   * Handle list sites tool request
   */
  private async handleListSites(args: any) {
    const searchTerm = args?.search;
    
    try {
      let endpoint: string;
      if (searchTerm) {
        endpoint = `/sites?search=${encodeURIComponent(searchTerm)}&$select=id,displayName,name,webUrl,description`;
      } else {
        endpoint = "/sites?$select=id,displayName,name,webUrl,description";
      }

      const response = await this.graphRequest(endpoint);
      const sites = response.value || [];

      return {
        content: [{
          type: "text",
          text: JSON.stringify(sites, null, 2),
        }],
      };
    } catch (error) {
      throw new Error(`Failed to list sites: ${error}`);
    }
  }

  /**
   * Handle get site info tool request
   */
  private async handleGetSiteInfo(args: any) {
    const siteUrl = args?.siteUrl;

    if (typeof siteUrl !== "string") {
      throw new McpError(ErrorCode.InvalidParams, "siteUrl parameter must be a string");
    }

    try {
      const siteId = await this.getSiteIdFromUrl(siteUrl);
      const site = await this.graphRequest(
        `/sites/${siteId}?$select=id,displayName,description,webUrl,createdDateTime,lastModifiedDateTime&$expand=drive($select=id,name,driveType,webUrl,quota)`
      );

      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            id: site.id,
            name: site.displayName,
            description: site.description,
            webUrl: site.webUrl,
            created: site.createdDateTime,
            lastModified: site.lastModifiedDateTime,
            defaultDrive: site.drive ? {
              id: site.drive.id,
              name: site.drive.name,
              driveType: site.drive.driveType,
              webUrl: site.drive.webUrl,
              totalSize: site.drive.quota?.total,
              usedSize: site.drive.quota?.used,
            } : null,
          }, null, 2),
        }],
      };
    } catch (error) {
      throw new Error(`Failed to get site info: ${error}`);
    }
  }

  /**
   * Handle list site drives tool request
   */
  private async handleListSiteDrives(args: any) {
    const siteUrl = args?.siteUrl;
    const scope = args?.scope || "tenant";

    const projectDrive = (d: any) => ({
      id: d.id,
      name: d.name,
      driveType: d.driveType,
      webUrl: d.webUrl,
      totalSize: d.quota?.total,
      usedSize: d.quota?.used,
    });

    try {
      if (scope === "me") {
        const response = await this.graphRequestAsUser(
          "/me/drives?$select=id,name,driveType,webUrl,quota"
        );
        const drives = (response.value || []).map(projectDrive);
        return {
          content: [{ type: "text", text: JSON.stringify(drives, null, 2) }],
        };
      }

      // tenant scope
      if (typeof siteUrl !== "string") {
        throw new McpError(ErrorCode.InvalidParams, "siteUrl is required when scope=tenant");
      }

      const siteId = await this.getSiteIdFromUrl(siteUrl);
      const response = await this.graphRequest(
        `/sites/${siteId}/drives?$select=id,name,driveType,webUrl,quota`
      );
      const drives = (response.value || []).map(projectDrive);

      return {
        content: [{ type: "text", text: JSON.stringify(drives, null, 2) }],
      };
    } catch (error) {
      throw new Error(`Failed to list drives: ${error}`);
    }
  }

  /**
   * Handle list drive items tool request
   */
  private async handleListDriveItems(args: any) {
    const siteUrl = args?.siteUrl;
    const driveId = args?.driveId;
    const folderPath = args?.folderPath;
    const scope = args?.scope || "tenant";

    const SELECT = "$select=id,name,webUrl,size,lastModifiedDateTime,folder,file";
    const TOP = "$top=50";
    const nextLinkArg = args?.nextLink;

    const projectItem = (item: any) => ({
      name: item.name,
      type: item.folder ? "folder" : "file",
      size: item.size,
      lastModified: item.lastModifiedDateTime,
      webUrl: item.webUrl,
      ...(item.folder ? { childCount: item.folder.childCount } : {}),
      ...(item.file?.mimeType ? { mimeType: item.file.mimeType } : {}),
    });

    try {
      let response: any;

      if (nextLinkArg) {
        response = scope === "me"
          ? await this.graphRequestAsUser(nextLinkArg)
          : await this.graphRequest(nextLinkArg);
      } else if (scope === "me") {
        const base = driveId ? `/me/drives/${driveId}` : "/me/drive";
        const endpoint = folderPath
          ? `${base}/root:/${folderPath}:/children?${SELECT}&${TOP}`
          : `${base}/root/children?${SELECT}&${TOP}`;
        response = await this.graphRequestAsUser(endpoint);
      } else {
        if (typeof siteUrl !== "string") {
          throw new McpError(ErrorCode.InvalidParams, "siteUrl is required when scope=tenant");
        }
        const siteId = await this.getSiteIdFromUrl(siteUrl);
        const base = driveId
          ? `/sites/${siteId}/drives/${driveId}`
          : `/sites/${siteId}/drive`;
        const endpoint = folderPath
          ? `${base}/root:/${folderPath}:/children?${SELECT}&${TOP}`
          : `${base}/root/children?${SELECT}&${TOP}`;
        response = await this.graphRequest(endpoint);
      }

      const items = (response.value || []).map(projectItem);
      const result: any = { items };
      if (response["@odata.nextLink"]) {
        result.nextLink = response["@odata.nextLink"];
      }

      return {
        content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
      };
    } catch (error) {
      throw new Error(`Failed to list drive items: ${error}`);
    }
  }

  /**
   * Handle get file content tool request
   */
  private async handleGetFileContent(args: any) {
    const siteUrl = args?.siteUrl;
    const filePath = args?.filePath;
    const driveId = args?.driveId;
    const scope = args?.scope || "tenant";

    if (typeof filePath !== "string") {
      throw new McpError(ErrorCode.InvalidParams, "filePath parameter must be a string");
    }

    const pathForExtension = this.pathForExtension(filePath);
    const isDocx = /\.docx$/i.test(pathForExtension);
    const isPdf = /\.pdf$/i.test(pathForExtension);
    const isXlsx = /\.xlsx$/i.test(pathForExtension);
    const isPptx = /\.pptx$/i.test(pathForExtension);

    try {
      let endpoint: string;
      let token: string;

      if (this.isHttpUrl(filePath)) {
        endpoint = `/shares/${this.encodeSharingUrl(filePath)}/driveItem/content`;
        token = scope === "me" ? await this.getUserAccessToken() : await this.getAccessToken();
      } else if (scope === "me") {
        endpoint = driveId
          ? `/me/drives/${driveId}/root:/${filePath}:/content`
          : `/me/drive/root:/${filePath}:/content`;
        token = await this.getUserAccessToken();
      } else {
        if (typeof siteUrl !== "string") {
          throw new McpError(ErrorCode.InvalidParams, "siteUrl is required when scope=tenant");
        }
        const siteId = await this.getSiteIdFromUrl(siteUrl);
        endpoint = driveId
          ? `/sites/${siteId}/drives/${driveId}/root:/${filePath}:/content`
          : `/sites/${siteId}/drive/root:/${filePath}:/content`;
        token = await this.getAccessToken();
      }

      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), 20000);

      let response: Response;
      try {
        response = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, {
          headers: { Authorization: `Bearer ${token}` },
          signal: controller.signal,
        });
      } catch (err: any) {
        clearTimeout(timer);
        if (err?.name === "AbortError") {
          throw new Error(`File download timed out after 20s: ${filePath}`);
        }
        throw err;
      }
      clearTimeout(timer);

      if (!response.ok) {
        throw new Error(`Failed to get file content: ${response.status} ${response.statusText}`);
      }

      const MAX_BYTES = 5_000_000; // 5 MB
      const contentLength = Number(response.headers.get("content-length") ?? 0);
      if (contentLength > MAX_BYTES) {
        throw new Error(`File too large (${(contentLength / 1024).toFixed(0)} KB). Download it directly instead.`);
      }

      if (isDocx) {
        const buf = await response.arrayBuffer();
        const zip = await JSZip.loadAsync(buf);
        const xmlFile = zip.file("word/document.xml");
        if (!xmlFile) {
          throw new Error("Not a valid .docx file (missing word/document.xml)");
        }
        const xml = await xmlFile.async("string");
        // Extract paragraph text: preserve paragraph breaks, strip all XML tags
        const text = xml
          .replace(/<w:p[ >]/g, "\n<w:p>")
          .replace(/<[^>]+>/g, "")
          .replace(/\n{3,}/g, "\n\n")
          .trim();
        return {
          content: [{ type: "text", text }],
        };
      }

      if (isXlsx) {
        const buf = await response.arrayBuffer();
        const text = await this.extractXlsxText(buf);
        if (text.length > MAX_BYTES) {
          throw new Error(`File content too large to return (${(text.length / 1024).toFixed(0)} KB).`);
        }
        return { content: [{ type: "text", text }] };
      }

      if (isPptx) {
        const buf = await response.arrayBuffer();
        const text = await this.extractPptxText(buf);
        if (text.length > MAX_BYTES) {
          throw new Error(`File content too large to return (${(text.length / 1024).toFixed(0)} KB).`);
        }
        return { content: [{ type: "text", text }] };
      }

      const contentType = response.headers.get("content-type") ?? "";
      if (isPdf || contentType.toLowerCase().includes("application/pdf")) {
        const pdfBytes = new Uint8Array(await response.arrayBuffer());
        let parser: PDFParse | undefined;
        try {
          parser = new PDFParse({ data: pdfBytes });
          const result = await parser.getText();
          const text = result.text.trim();
          if (!text) {
            throw new Error(
              `PDF text extraction returned no text from ${result.total} page(s); the PDF may be scanned/image-only and need OCR.`
            );
          }
          if (text.length > MAX_BYTES) {
            throw new Error(`File content too large to return (${(text.length / 1024).toFixed(0)} KB).`);
          }
          return {
            content: [{ type: "text", text }],
          };
        } finally {
          await parser?.destroy();
        }
      }

      const content = await response.text();
      if (content.length > MAX_BYTES) {
        throw new Error(`File content too large to return (${(content.length / 1024).toFixed(0)} KB).`);
      }

      return {
        content: [{ type: "text", text: content }],
      };
    } catch (error) {
      throw new Error(`Failed to get file content: ${error}`);
    }
  }

  private driveBase(scope: string, siteId: string, driveId?: string): string {
    if (scope === "me") {
      return driveId ? `/me/drives/${driveId}` : "/me/drive";
    }
    return driveId ? `/sites/${siteId}/drives/${driveId}` : `/sites/${siteId}/drive`;
  }

  private async handleUploadFile(args: any) {
    const {
      filePath, content, contentEncoding = "text", contentType = "text/plain",
      siteUrl, driveId, scope = "tenant", conflictBehavior = "replace",
    } = args || {};

    if (!filePath || content === undefined) {
      throw new McpError(ErrorCode.InvalidParams, "filePath and content are required");
    }

    const bodyBuffer = contentEncoding === "base64"
      ? Buffer.from(content, "base64")
      : Buffer.from(content, "utf-8");

    let siteId = "";
    if (scope === "tenant") {
      if (!siteUrl) throw new McpError(ErrorCode.InvalidParams, "siteUrl is required when scope=tenant");
      siteId = await this.getSiteIdFromUrl(siteUrl);
    }

    const base = this.driveBase(scope, siteId, driveId);
    const endpoint = `${base}/root:/${filePath}:/content?@microsoft.graph.conflictBehavior=${conflictBehavior}`;
    const token = scope === "me" ? await this.getUserAccessToken() : await this.getUserAccessToken();

    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 30000);

    let response: Response;
    try {
      response = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, {
        method: "PUT",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": contentType },
        body: bodyBuffer,
        signal: controller.signal,
      });
    } catch (err: any) {
      clearTimeout(timer);
      if (err?.name === "AbortError") throw new Error(`Upload timed out after 30s: ${filePath}`);
      throw err;
    }
    clearTimeout(timer);

    if (!response.ok) {
      const errText = await response.text();
      throw new Error(`Upload failed: ${response.status} ${response.statusText} - ${errText}`);
    }

    const item = await response.json();
    return {
      content: [{
        type: "text",
        text: JSON.stringify({ id: item.id, name: item.name, webUrl: item.webUrl, size: item.size }, null, 2),
      }],
    };
  }

  private async handleCreateFolder(args: any) {
    const {
      folderName, parentPath = "", siteUrl, driveId, scope = "tenant", conflictBehavior = "rename",
    } = args || {};

    if (!folderName) throw new McpError(ErrorCode.InvalidParams, "folderName is required");

    let siteId = "";
    if (scope === "tenant") {
      if (!siteUrl) throw new McpError(ErrorCode.InvalidParams, "siteUrl is required when scope=tenant");
      siteId = await this.getSiteIdFromUrl(siteUrl);
    }

    const base = this.driveBase(scope, siteId, driveId);
    const endpoint = parentPath
      ? `${base}/root:/${parentPath}:/children`
      : `${base}/root/children`;

    const token = await this.getUserAccessToken();
    const url = `https://graph.microsoft.com/v1.0${endpoint}`;

    const response = await fetch(url, {
      method: "POST",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        name: folderName,
        folder: {},
        "@microsoft.graph.conflictBehavior": conflictBehavior,
      }),
    });

    if (!response.ok) {
      const errText = await response.text();
      throw new Error(`Create folder failed: ${response.status} ${response.statusText} - ${errText}`);
    }

    const item = await response.json();
    return {
      content: [{
        type: "text",
        text: JSON.stringify({ id: item.id, name: item.name, webUrl: item.webUrl }, null, 2),
      }],
    };
  }

  private async handleDeleteDriveItem(args: any) {
    const { itemPath, siteUrl, driveId, scope = "tenant" } = args || {};

    if (!itemPath) throw new McpError(ErrorCode.InvalidParams, "itemPath is required");

    let siteId = "";
    if (scope === "tenant") {
      if (!siteUrl) throw new McpError(ErrorCode.InvalidParams, "siteUrl is required when scope=tenant");
      siteId = await this.getSiteIdFromUrl(siteUrl);
    }

    const base = this.driveBase(scope, siteId, driveId);
    const endpoint = `${base}/root:/${itemPath}`;
    const token = await this.getUserAccessToken();

    const response = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, {
      method: "DELETE",
      headers: { Authorization: `Bearer ${token}` },
    });

    if (!response.ok) {
      const errText = await response.text();
      throw new Error(`Delete failed: ${response.status} ${response.statusText} - ${errText}`);
    }

    return {
      content: [{ type: "text", text: JSON.stringify({ success: true, deleted: itemPath }) }],
    };
  }

  private async handleMoveRenameDriveItem(args: any) {
    const { itemPath, newName, newParentPath, siteUrl, driveId, scope = "tenant" } = args || {};

    if (!itemPath) throw new McpError(ErrorCode.InvalidParams, "itemPath is required");
    if (!newName && !newParentPath) throw new McpError(ErrorCode.InvalidParams, "Provide newName, newParentPath, or both");

    let siteId = "";
    if (scope === "tenant") {
      if (!siteUrl) throw new McpError(ErrorCode.InvalidParams, "siteUrl is required when scope=tenant");
      siteId = await this.getSiteIdFromUrl(siteUrl);
    }

    const base = this.driveBase(scope, siteId, driveId);
    const token = await this.getUserAccessToken();

    // Resolve item ID
    const itemMeta = await fetch(`https://graph.microsoft.com/v1.0${base}/root:/${itemPath}?$select=id,parentReference`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!itemMeta.ok) throw new Error(`Could not resolve item at path: ${itemPath}`);
    const item = await itemMeta.json();

    const patch: any = {};
    if (newName) patch.name = newName;

    if (newParentPath) {
      // Resolve destination folder ID
      const destMeta = await fetch(`https://graph.microsoft.com/v1.0${base}/root:/${newParentPath}?$select=id`, {
        headers: { Authorization: `Bearer ${token}` },
      });
      if (!destMeta.ok) throw new Error(`Could not resolve destination folder: ${newParentPath}`);
      const dest = await destMeta.json();
      patch.parentReference = { id: dest.id };
    }

    const patchEndpoint = scope === "me"
      ? `/me/drive/items/${item.id}`
      : `/sites/${siteId}/drive/items/${item.id}`;

    const response = await fetch(`https://graph.microsoft.com/v1.0${patchEndpoint}`, {
      method: "PATCH",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify(patch),
    });

    if (!response.ok) {
      const errText = await response.text();
      throw new Error(`Move/rename failed: ${response.status} ${response.statusText} - ${errText}`);
    }

    const updated = await response.json();
    return {
      content: [{
        type: "text",
        text: JSON.stringify({ id: updated.id, name: updated.name, webUrl: updated.webUrl }, null, 2),
      }],
    };
  }

  /**
   * Handle list emails tool request (delegated /me/ endpoint)
   */
  private async handleListEmails(args: any) {
    const folder = args?.folder || "inbox";
    const top = Math.min(args?.top || 10, 50);
    const filter = args?.filter;
    const nextLinkArg = args?.nextLink;

    try {
      let endpoint: string;
      if (nextLinkArg) {
        endpoint = nextLinkArg;
      } else {
        endpoint = `/me/mailFolders/${folder}/messages?$top=${top}&$select=id,subject,from,toRecipients,receivedDateTime,isRead,bodyPreview&$orderby=receivedDateTime desc`;
        if (filter) {
          endpoint += `&$filter=${encodeURIComponent(filter)}`;
        }
      }

      const response = await this.graphRequestAsUser(endpoint);
      const messages = (response.value || []).map((m: any) => ({
        id: m.id,
        subject: m.subject,
        from: m.from?.emailAddress?.address,
        fromName: m.from?.emailAddress?.name,
        to: m.toRecipients?.map((r: any) => r.emailAddress?.address),
        date: m.receivedDateTime,
        isRead: m.isRead,
        preview: m.bodyPreview,
      }));

      const result: any = { messages };
      if (response["@odata.nextLink"]) {
        result.nextLink = response["@odata.nextLink"];
      }

      return {
        content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
      };
    } catch (error) {
      throw new Error(`Failed to list emails: ${error}`);
    }
  }

  /**
   * Handle get email tool request (delegated /me/ endpoint)
   */
  private async handleGetEmail(args: any) {
    const messageId = args?.messageId;

    if (typeof messageId !== "string") {
      throw new McpError(ErrorCode.InvalidParams, "messageId parameter must be a string");
    }

    try {
      const message = await this.graphRequestAsUser(
        `/me/messages/${messageId}?$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,hasAttachments,importance`
      );

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(
              {
                id: message.id,
                subject: message.subject,
                from: message.from?.emailAddress,
                to: message.toRecipients?.map((r: any) => r.emailAddress),
                cc: message.ccRecipients?.map((r: any) => r.emailAddress),
                date: message.receivedDateTime,
                hasAttachments: message.hasAttachments,
                importance: message.importance,
                body: message.body?.content,
                bodyType: message.body?.contentType,
              },
              null,
              2
            ),
          },
        ],
      };
    } catch (error) {
      throw new Error(`Failed to get email: ${error}`);
    }
  }

  /**
   * Handle search emails tool request (delegated /me/ endpoint)
   */
  private async handleSearchEmails(args: any) {
    const query = args?.query;
    const top = Math.min(args?.top || 10, 50);
    const nextLinkArg = args?.nextLink;

    if (!nextLinkArg && typeof query !== "string") {
      throw new McpError(ErrorCode.InvalidParams, "query parameter must be a string");
    }

    try {
      const endpoint = nextLinkArg
        || `/me/messages?$search="${encodeURIComponent(query)}"&$top=${top}&$select=id,subject,from,receivedDateTime,bodyPreview`;

      const response = await this.graphRequestAsUser(endpoint);
      const messages = (response.value || []).map((m: any) => ({
        id: m.id,
        subject: m.subject,
        from: m.from?.emailAddress?.address,
        fromName: m.from?.emailAddress?.name,
        date: m.receivedDateTime,
        preview: m.bodyPreview,
      }));

      const result: any = { messages };
      if (response["@odata.nextLink"]) {
        result.nextLink = response["@odata.nextLink"];
      }

      return {
        content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
      };
    } catch (error) {
      throw new Error(`Failed to search emails: ${error}`);
    }
  }

  /**
   * Handle send email tool request (delegated /me/ endpoint)
   */
  private async handleSendEmail(args: any) {
    const toRaw = args?.to;
    const to: string[] = typeof toRaw === "string" ? [toRaw] : toRaw;
    const subject: string = args?.subject;
    const body: string = args?.body;
    const ccRaw = args?.cc;
    const cc: string[] = typeof ccRaw === "string" ? [ccRaw] : (ccRaw || []);
    const bccRaw = args?.bcc;
    const bcc: string[] = typeof bccRaw === "string" ? [bccRaw] : (bccRaw || []);
    const bodyType: string = args?.bodyType || "HTML";

    if (!Array.isArray(to) || !to.length || !subject || !body) {
      throw new McpError(ErrorCode.InvalidParams, "to, subject, and body are required");
    }

    const toRecipients = to.map((email) => ({ emailAddress: { address: email } }));
    const ccRecipients = cc.map((email) => ({ emailAddress: { address: email } }));
    const bccRecipients = bcc.map((email) => ({ emailAddress: { address: email } }));

    // Auto-append signature if image file is saved
    const sigFile = path.join(
      path.dirname(TOKEN_FILE),
      "signature.png"
    );
    const attachments: any[] = [];
    let finalBody = body;

    if (bodyType === "HTML" && fs.existsSync(sigFile)) {
      const sigBytes = fs.readFileSync(sigFile).toString("base64");
      const sigCid = "bay-view-signature";
      attachments.push({
        "@odata.type": "#microsoft.graph.fileAttachment",
        name: "signature.png",
        contentType: "image/png",
        contentBytes: sigBytes,
        isInline: true,
        contentId: sigCid,
      });
      finalBody =
        body +
        `<br><br><a href="https://bayviewassociation.org" style="text-decoration:none">` +
        `<img src="cid:${sigCid}" width="484" height="215" style="max-width:780px; display:block">` +
        `</a>`;
    }

    const message: any = {
      subject,
      body: { contentType: bodyType, content: finalBody },
      toRecipients,
      ccRecipients,
      bccRecipients,
    };
    if (attachments.length > 0) {
      message.attachments = attachments;
    }

    await this.graphRequestAsUser("/me/sendMail", "POST", { message });

    return {
      content: [
        {
          type: "text",
          text: JSON.stringify({
            success: true,
            message: `Email sent to ${to.join(", ")}`,
            signatureIncluded: attachments.length > 0,
          }),
        },
      ],
    };
  }

  /**
   * Handle reply email tool request (delegated /me/ endpoint)
   */
  private async handleReplyEmail(args: any) {
    const messageId: string = args?.messageId;
    const comment: string = args?.comment;
    const replyAll: boolean = args?.replyAll || false;

    if (!messageId || !comment) {
      throw new McpError(ErrorCode.InvalidParams, "messageId and comment are required");
    }

    // Check for signature file to decide whether to use createReply (supports attachments)
    const sigFile = path.join(path.dirname(TOKEN_FILE), "signature.png");

    if (fs.existsSync(sigFile)) {
      // Use createReply so we can attach the inline signature image
      const createEndpoint = replyAll
        ? `/me/messages/${messageId}/createReplyAll`
        : `/me/messages/${messageId}/createReply`;

      const draft: any = await this.graphRequestAsUser(createEndpoint, "POST", {});
      const draftId: string = draft.id;

      const sigBytes = fs.readFileSync(sigFile).toString("base64");
      const sigCid = "bay-view-signature";
      const htmlBody =
        comment +
        `<br><br><a href="https://bayviewassociation.org" style="text-decoration:none">` +
        `<img src="cid:${sigCid}" width="484" height="215" style="max-width:780px; display:block">` +
        `</a>`;

      // Update the draft body and add the inline attachment
      await this.graphRequestAsUser(`/me/messages/${draftId}`, "PATCH", {
        body: { contentType: "HTML", content: htmlBody },
      });
      await this.graphRequestAsUser(`/me/messages/${draftId}/attachments`, "POST", {
        "@odata.type": "#microsoft.graph.fileAttachment",
        name: "signature.png",
        contentType: "image/png",
        contentBytes: sigBytes,
        isInline: true,
        contentId: sigCid,
      });

      await this.graphRequestAsUser(`/me/messages/${draftId}/send`, "POST", {});
    } else {
      const endpoint = replyAll
        ? `/me/messages/${messageId}/replyAll`
        : `/me/messages/${messageId}/reply`;

      await this.graphRequestAsUser(endpoint, "POST", { comment });
    }

    return {
      content: [
        {
          type: "text",
          text: JSON.stringify({
            success: true,
            message: `Reply sent${replyAll ? " to all" : ""}`,
          }),
        },
      ],
    };
  }

  /**
   * Handle list mail folders tool request
   */
  private async handleListMailFolders(_args: any) {
    try {
      const response = await this.graphRequestAsUser(
        "/me/mailFolders?$select=id,displayName,totalItemCount,unreadItemCount&$top=50"
      );
      const folders = (response.value || []).map((f: any) => ({
        id: f.id,
        name: f.displayName,
        totalItems: f.totalItemCount,
        unreadItems: f.unreadItemCount,
      }));

      return {
        content: [{ type: "text", text: JSON.stringify(folders, null, 2) }],
      };
    } catch (error) {
      throw new Error(`Failed to list mail folders: ${error}`);
    }
  }

  /**
   * Handle move email tool request
   */
  private async handleMoveEmail(args: any) {
    const messageId: string = args?.messageId;
    const destinationFolder: string = args?.destinationFolder;

    if (!messageId || !destinationFolder) {
      throw new McpError(ErrorCode.InvalidParams, "messageId and destinationFolder are required");
    }

    try {
      // Well-known folder names can be used directly; otherwise treat as folder ID
      const result = await this.graphRequestAsUser(
        `/me/messages/${messageId}/move`,
        "POST",
        { destinationId: destinationFolder }
      );

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify({
              success: true,
              message: `Email moved to ${destinationFolder}`,
            }),
          },
        ],
      };
    } catch (error) {
      throw new Error(`Failed to move email: ${error}`);
    }
  }

  /**
   * Handle mark email read/unread tool request
   */
  private async handleMarkEmailRead(args: any) {
    const messageId: string = args?.messageId;
    const isRead: boolean = args?.isRead;

    if (!messageId || typeof isRead !== "boolean") {
      throw new McpError(ErrorCode.InvalidParams, "messageId and isRead (boolean) are required");
    }

    try {
      await this.graphRequestAsUser(
        `/me/messages/${messageId}`,
        "PATCH",
        { isRead }
      );

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify({
              success: true,
              message: `Email marked as ${isRead ? "read" : "unread"}`,
            }),
          },
        ],
      };
    } catch (error) {
      throw new Error(`Failed to update email: ${error}`);
    }
  }

  /**
   * Handle delete email tool request (soft delete — moves to Deleted Items)
   */
  private async handleDeleteEmail(args: any) {
    const messageId: string = args?.messageId;

    if (!messageId) {
      throw new McpError(ErrorCode.InvalidParams, "messageId is required");
    }

    try {
      const token = await this.getUserAccessToken();
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/messages/${messageId}`,
        {
          method: "DELETE",
          headers: { Authorization: `Bearer ${token}` },
        }
      );

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`${response.status} ${response.statusText} - ${errorText}`);
      }

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify({ success: true, message: "Email deleted (moved to Deleted Items)" }),
          },
        ],
      };
    } catch (error) {
      throw new Error(`Failed to delete email: ${error}`);
    }
  }

  /**
   * Handle get email attachments tool request
   * Optionally saves the first inline image as the default signature file.
   */
  private async handleGetEmailAttachments(args: any) {
    const messageId: string = args?.messageId;
    const saveSignature: boolean = args?.saveSignature || false;

    if (!messageId) {
      throw new McpError(ErrorCode.InvalidParams, "messageId is required");
    }

    try {
      const response = await this.graphRequestAsUser(
        `/me/messages/${messageId}/attachments`
      );
      const attachments = (response.value || []).map((a: any) => ({
        id: a.id,
        name: a.name,
        contentType: a.contentType,
        size: a.size,
        isInline: a.isInline,
        contentId: a.contentId,
        contentBytes: a.contentBytes,
      }));

      if (saveSignature) {
        const inlineImg = attachments.find(
          (a: any) => a.isInline && a.contentType?.startsWith("image/")
        );
        if (inlineImg?.contentBytes) {
          const sigFile = path.join(path.dirname(TOKEN_FILE), "signature.png");
          fs.mkdirSync(path.dirname(sigFile), { recursive: true });
          fs.writeFileSync(sigFile, Buffer.from(inlineImg.contentBytes, "base64"));
          return {
            content: [
              {
                type: "text",
                text: JSON.stringify({
                  saved: true,
                  file: sigFile,
                  name: inlineImg.name,
                  contentType: inlineImg.contentType,
                  size: inlineImg.size,
                  attachmentCount: attachments.length,
                }),
              },
            ],
          };
        }
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({
                saved: false,
                reason: "No inline image attachment found",
                attachmentCount: attachments.length,
              }),
            },
          ],
        };
      }

      // Return metadata only (no contentBytes in summary to keep response small)
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(
              attachments.map(({ contentBytes: _cb, ...rest }: any) => rest),
              null,
              2
            ),
          },
        ],
      };
    } catch (error) {
      throw new Error(`Failed to get email attachments: ${error}`);
    }
  }

  /**
   * Handle list calendar events tool request (delegated /me/ endpoint)
   */
  private async handleListCalendarEvents(args: any) {
    const now = new Date();
    const defaultEnd = new Date(now);
    defaultEnd.setDate(defaultEnd.getDate() + 90);

    const start = args?.start || now.toISOString();
    const end = args?.end || defaultEnd.toISOString();
    const top = Math.min(args?.top || 50, 100);
    const calendarId = args?.calendarId;
    const nextLinkArg = args?.nextLink;

    try {
      let endpoint: string;
      if (nextLinkArg) {
        endpoint = nextLinkArg;
      } else {
        const base = calendarId ? `/me/calendars/${calendarId}/events` : `/me/calendarView`;
        endpoint =
          `${base}?startDateTime=${encodeURIComponent(start)}&endDateTime=${encodeURIComponent(end)}` +
          `&$top=${top}&$select=id,subject,start,end,location,organizer,attendees,bodyPreview,isAllDay,isCancelled` +
          `&$orderby=start/dateTime`;
      }

      const response = await this.graphRequestAsUser(endpoint);
      const events = (response.value || []).map((e: any) => ({
        id: e.id,
        subject: e.subject,
        start: e.start,
        end: e.end,
        location: e.location?.displayName,
        organizer: e.organizer?.emailAddress,
        attendees: e.attendees?.map((a: any) => ({
          email: a.emailAddress?.address,
          name: a.emailAddress?.name,
          status: a.status?.response,
        })),
        isAllDay: e.isAllDay,
        isCancelled: e.isCancelled,
        preview: e.bodyPreview,
      }));

      const result: any = { events };
      if (response["@odata.nextLink"]) {
        result.nextLink = response["@odata.nextLink"];
      }

      return {
        content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
      };
    } catch (error) {
      throw new Error(`Failed to list calendar events: ${error}`);
    }
  }

  /**
   * Handle get calendar event tool request (delegated /me/ endpoint)
   */
  private async handleGetCalendarEvent(args: any) {
    const eventId = args?.eventId;

    if (typeof eventId !== "string") {
      throw new McpError(ErrorCode.InvalidParams, "eventId parameter must be a string");
    }

    try {
      const event = await this.graphRequestAsUser(
        `/me/events/${eventId}?$select=id,subject,start,end,location,organizer,attendees,body,isAllDay,isCancelled,recurrence`
      );

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(
              {
                id: event.id,
                subject: event.subject,
                start: event.start,
                end: event.end,
                location: event.location?.displayName,
                organizer: event.organizer?.emailAddress,
                attendees: event.attendees?.map((a: any) => ({
                  email: a.emailAddress?.address,
                  name: a.emailAddress?.name,
                  status: a.status?.response,
                })),
                isAllDay: event.isAllDay,
                isCancelled: event.isCancelled,
                recurrence: event.recurrence,
                body: event.body?.content,
                bodyType: event.body?.contentType,
              },
              null,
              2
            ),
          },
        ],
      };
    } catch (error) {
      throw new Error(`Failed to get calendar event: ${error}`);
    }
  }

  /**
   * Handle create calendar event tool request (delegated /me/ endpoint)
   */
  private async handleCreateCalendarEvent(args: any) {
    const { subject, start, end, timeZone = "America/Detroit", location, body, attendees, calendarId } = args || {};

    if (!subject || !start || !end) {
      throw new McpError(ErrorCode.InvalidParams, "subject, start, and end are required");
    }

    const eventBody: any = {
      subject,
      start: { dateTime: start, timeZone },
      end: { dateTime: end, timeZone },
    };

    if (location) {
      eventBody.location = { displayName: location };
    }

    if (body) {
      eventBody.body = { contentType: "text", content: body };
    }

    if (attendees && Array.isArray(attendees) && attendees.length > 0) {
      eventBody.attendees = attendees.map((email: string) => ({
        emailAddress: { address: email },
        type: "required",
      }));
    }

    try {
      const endpoint = calendarId ? `/me/calendars/${calendarId}/events` : `/me/events`;
      const event = await this.graphRequestAsUser(endpoint, "POST", eventBody);

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(
              {
                id: event.id,
                subject: event.subject,
                start: event.start,
                end: event.end,
                location: event.location?.displayName,
                webLink: event.webLink,
              },
              null,
              2
            ),
          },
        ],
      };
    } catch (error) {
      throw new Error(`Failed to create calendar event: ${error}`);
    }
  }

  private async handleUpdateCalendarEvent(args: any) {
    const { eventId, subject, start, end, timeZone = "America/Detroit", location, body, attendees, calendarId } = args || {};
    if (!eventId) throw new McpError(ErrorCode.InvalidParams, "eventId is required");

    const patch: any = {};
    if (subject) patch.subject = subject;
    if (start) patch.start = { dateTime: start, timeZone };
    if (end) patch.end = { dateTime: end, timeZone };
    if (location) patch.location = { displayName: location };
    if (body) patch.body = { contentType: "text", content: body };
    if (attendees && Array.isArray(attendees)) {
      patch.attendees = attendees.map((email: string) => ({
        emailAddress: { address: email },
        type: "required",
      }));
    }

    try {
      const endpoint = calendarId ? `/me/calendars/${calendarId}/events/${eventId}` : `/me/events/${eventId}`;
      const event = await this.graphRequestAsUser(endpoint, "PATCH", patch);
      return {
        content: [{
          type: "text",
          text: JSON.stringify({ id: event.id, subject: event.subject, start: event.start, end: event.end, webLink: event.webLink }, null, 2),
        }],
      };
    } catch (error) {
      throw new Error(`Failed to update calendar event: ${error}`);
    }
  }

  private async handleDeleteCalendarEvent(args: any) {
    const { eventId, calendarId } = args || {};
    if (!eventId) throw new McpError(ErrorCode.InvalidParams, "eventId is required");

    try {
      const endpoint = calendarId ? `/me/calendars/${calendarId}/events/${eventId}` : `/me/events/${eventId}`;
      const token = await this.getUserAccessToken();
      const response = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, {
        method: "DELETE",
        headers: { Authorization: `Bearer ${token}` },
      });
      if (!response.ok) throw new Error(`${response.status} ${response.statusText}`);
      return { content: [{ type: "text", text: JSON.stringify({ success: true, deleted: eventId }) }] };
    } catch (error) {
      throw new Error(`Failed to delete calendar event: ${error}`);
    }
  }

  private async handleRespondToEvent(args: any) {
    const { eventId, response: responseType, comment, sendResponse = true } = args || {};
    if (!eventId || !responseType) throw new McpError(ErrorCode.InvalidParams, "eventId and response are required");
    if (!["accept", "decline", "tentative"].includes(responseType)) {
      throw new McpError(ErrorCode.InvalidParams, "response must be accept, decline, or tentative");
    }

    try {
      const actionMap = { accept: "accept", decline: "decline", tentative: "tentativelyAccept" };
      const action = actionMap[responseType as keyof typeof actionMap];
      const endpoint = `/me/events/${eventId}/${action}`;
      const token = await this.getUserAccessToken();
      const body = comment ? { comment, sendResponse } : { sendResponse };

      const response = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, {
        method: "POST",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify(body),
      });
      if (!response.ok) throw new Error(`${response.status} ${response.statusText}`);
      return { content: [{ type: "text", text: JSON.stringify({ success: true, response: responseType }) }] };
    } catch (error) {
      throw new Error(`Failed to respond to event: ${error}`);
    }
  }

  private async handleListLists(args: any) {
    const { siteUrl, filter } = args || {};
    if (!siteUrl) throw new McpError(ErrorCode.InvalidParams, "siteUrl is required");

    try {
      const siteId = await this.getSiteIdFromUrl(siteUrl);
      let endpoint = `/sites/${siteId}/lists?$select=id,name,displayName,description,webUrl,list`;
      if (filter) endpoint += `&$filter=${encodeURIComponent(filter)}`;

      const response = await this.graphRequestAsUser(endpoint);
      const lists = (response.value || []).map((l: any) => ({
        id: l.id,
        name: l.name,
        displayName: l.displayName,
        description: l.description,
        webUrl: l.webUrl,
        template: l.list?.template,
      }));
      return { content: [{ type: "text", text: JSON.stringify(lists, null, 2) }] };
    } catch (error) {
      throw new Error(`Failed to list lists: ${error}`);
    }
  }

  private async handleGetListItems(args: any) {
    const { siteUrl, listId, filter, select, top = 50, nextLink } = args || {};
    if (!siteUrl || !listId) throw new McpError(ErrorCode.InvalidParams, "siteUrl and listId are required");

    try {
      const siteId = await this.getSiteIdFromUrl(siteUrl);
      let endpoint = `/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=${Math.min(top, 200)}`;
      if (select) endpoint += `&$select=${encodeURIComponent(select)}`;
      if (filter) endpoint += `&$filter=${encodeURIComponent(filter)}`;
      if (nextLink) {
        // Extract $skiptoken from a previous nextLink and append
        const skipToken = new URL(nextLink).searchParams.get("$skiptoken");
        if (skipToken) endpoint += `&$skiptoken=${encodeURIComponent(skipToken)}`;
      }

      const response = await this.graphRequestAsUser(endpoint);
      const items = (response.value || []).map((i: any) => ({
        id: i.id,
        fields: i.fields,
        webUrl: i.webUrl,
        createdDateTime: i.createdDateTime,
        lastModifiedDateTime: i.lastModifiedDateTime,
      }));
      const result: any = { items };
      if (response["@odata.nextLink"]) result.nextLink = response["@odata.nextLink"];
      return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
    } catch (error) {
      throw new Error(`Failed to get list items: ${error}`);
    }
  }

  private async handleGetListItem(args: any) {
    const { siteUrl, listId, itemId } = args || {};
    if (!siteUrl || !listId || !itemId) throw new McpError(ErrorCode.InvalidParams, "siteUrl, listId, and itemId are required");

    try {
      const siteId = await this.getSiteIdFromUrl(siteUrl);
      const endpoint = `/sites/${siteId}/lists/${listId}/items/${itemId}?$expand=fields`;
      const item = await this.graphRequestAsUser(endpoint);
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            id: item.id,
            fields: item.fields,
            webUrl: item.webUrl,
            createdDateTime: item.createdDateTime,
            lastModifiedDateTime: item.lastModifiedDateTime,
          }, null, 2),
        }],
      };
    } catch (error) {
      throw new Error(`Failed to get list item: ${error}`);
    }
  }

  private async handleCreateListItem(args: any) {
    const { siteUrl, listId, fields } = args || {};
    if (!siteUrl || !listId || !fields) throw new McpError(ErrorCode.InvalidParams, "siteUrl, listId, and fields are required");

    try {
      const siteId = await this.getSiteIdFromUrl(siteUrl);
      const endpoint = `/sites/${siteId}/lists/${listId}/items`;
      const item = await this.graphRequestAsUser(endpoint, "POST", { fields });
      return {
        content: [{
          type: "text",
          text: JSON.stringify({ id: item.id, fields: item.fields, webUrl: item.webUrl }, null, 2),
        }],
      };
    } catch (error) {
      throw new Error(`Failed to create list item: ${error}`);
    }
  }

  private async handleUpdateListItem(args: any) {
    const { siteUrl, listId, itemId, fields } = args || {};
    if (!siteUrl || !listId || !itemId || !fields) {
      throw new McpError(ErrorCode.InvalidParams, "siteUrl, listId, itemId, and fields are required");
    }

    try {
      const siteId = await this.getSiteIdFromUrl(siteUrl);
      const endpoint = `/sites/${siteId}/lists/${listId}/items/${itemId}/fields`;
      const updated = await this.graphRequestAsUser(endpoint, "PATCH", fields);
      return { content: [{ type: "text", text: JSON.stringify(updated, null, 2) }] };
    } catch (error) {
      throw new Error(`Failed to update list item: ${error}`);
    }
  }

  private async handleDeleteListItem(args: any) {
    const { siteUrl, listId, itemId } = args || {};
    if (!siteUrl || !listId || !itemId) throw new McpError(ErrorCode.InvalidParams, "siteUrl, listId, and itemId are required");

    try {
      const siteId = await this.getSiteIdFromUrl(siteUrl);
      const endpoint = `/sites/${siteId}/lists/${listId}/items/${itemId}`;
      const token = await this.getUserAccessToken();
      const response = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, {
        method: "DELETE",
        headers: { Authorization: `Bearer ${token}` },
      });
      if (!response.ok) throw new Error(`${response.status} ${response.statusText}`);
      return { content: [{ type: "text", text: JSON.stringify({ success: true, deleted: itemId }) }] };
    } catch (error) {
      throw new Error(`Failed to delete list item: ${error}`);
    }
  }

  /**
   * Start the MCP server
   */
  async run(): Promise<void> {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
    console.error("SharePoint MCP server running on stdio");
  }
}

/**
 * Main entry point
 */
const server = new SharePointServer();
server.run().catch((error) => {
  console.error("Failed to start SharePoint MCP server:", error);
  process.exit(1);
});
