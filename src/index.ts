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

const DELEGATED_SCOPES = "Mail.ReadWrite Mail.Send Files.Read.All offline_access";

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
  private async graphRequest(endpoint: string, method: string = "GET", body?: any): Promise<any> {
    const token = await this.getAccessToken();
    const url = `https://graph.microsoft.com/v1.0${endpoint}`;

    const headers: Record<string, string> = {
      "Authorization": `Bearer ${token}`,
      "Content-Type": "application/json",
    };

    const options: RequestInit = {
      method,
      headers,
    };

    if (body && method !== "GET") {
      options.body = JSON.stringify(body);
    }

    try {
      const response = await fetch(url, options);

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Graph API request failed: ${response.status} ${response.statusText} - ${errorText}`);
      }

      return await response.json();
    } catch (error) {
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
    body?: any
  ): Promise<any> {
    const token = await this.getUserAccessToken();
    const url = `https://graph.microsoft.com/v1.0${endpoint}`;

    const headers: Record<string, string> = {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    };

    const options: RequestInit = { method, headers };
    if (body && method !== "GET") {
      options.body = JSON.stringify(body);
    }

    const response = await fetch(url, options);

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
            "Search for files and documents. scope=tenant searches all SharePoint sites (app auth); scope=me searches your OneDrive (delegated auth).",
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
            },
          },
        },
        {
          name: "get_file_content",
          description:
            "Get the content of a file (text files only). scope=tenant reads from SharePoint; scope=me reads from your OneDrive.",
          inputSchema: {
            type: "object",
            properties: {
              siteUrl: {
                type: "string",
                description: "The SharePoint site URL (required when scope=tenant)",
              },
              filePath: {
                type: "string",
                description: "The path to the file",
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
      let endpoint: string;
      let response: any;

      if (scope === "me") {
        const base = driveId ? `/me/drives/${driveId}` : "/me/drive";
        endpoint = folderPath
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
        endpoint = folderPath
          ? `${base}/root:/${folderPath}:/children?${SELECT}&${TOP}`
          : `${base}/root/children?${SELECT}&${TOP}`;
        response = await this.graphRequest(endpoint);
      }

      const items = (response.value || []).map(projectItem);
      const hasMore = !!response["@odata.nextLink"];

      const result: any = { items };
      if (hasMore) {
        result.has_more = true;
        result.note = "More items exist. Use a more specific folderPath to drill down.";
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

    try {
      let endpoint: string;
      let token: string;

      if (scope === "me") {
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

      const response = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, {
        headers: { Authorization: `Bearer ${token}` },
      });

      if (!response.ok) {
        throw new Error(`Failed to get file content: ${response.status} ${response.statusText}`);
      }

      const content = await response.text();

      return {
        content: [{ type: "text", text: content }],
      };
    } catch (error) {
      throw new Error(`Failed to get file content: ${error}`);
    }
  }

  /**
   * Handle list emails tool request (delegated /me/ endpoint)
   */
  private async handleListEmails(args: any) {
    const folder = args?.folder || "inbox";
    const top = Math.min(args?.top || 10, 50);
    const filter = args?.filter;

    try {
      let endpoint = `/me/mailFolders/${folder}/messages?$top=${top}&$select=id,subject,from,toRecipients,receivedDateTime,isRead,bodyPreview&$orderby=receivedDateTime desc`;
      if (filter) {
        endpoint += `&$filter=${encodeURIComponent(filter)}`;
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

      return {
        content: [{ type: "text", text: JSON.stringify(messages, null, 2) }],
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

    if (typeof query !== "string") {
      throw new McpError(ErrorCode.InvalidParams, "query parameter must be a string");
    }

    try {
      const endpoint = `/me/messages?$search="${encodeURIComponent(query)}"&$top=${top}&$select=id,subject,from,receivedDateTime,bodyPreview`;

      const response = await this.graphRequestAsUser(endpoint);
      const messages = (response.value || []).map((m: any) => ({
        id: m.id,
        subject: m.subject,
        from: m.from?.emailAddress?.address,
        fromName: m.from?.emailAddress?.name,
        date: m.receivedDateTime,
        preview: m.bodyPreview,
      }));

      return {
        content: [{ type: "text", text: JSON.stringify(messages, null, 2) }],
      };
    } catch (error) {
      throw new Error(`Failed to search emails: ${error}`);
    }
  }

  /**
   * Handle send email tool request (delegated /me/ endpoint)
   */
  private async handleSendEmail(args: any) {
    const to: string[] = args?.to;
    const subject: string = args?.subject;
    const body: string = args?.body;
    const cc: string[] = args?.cc || [];
    const bcc: string[] = args?.bcc || [];
    const bodyType: string = args?.bodyType || "HTML";

    if (!Array.isArray(to) || !to.length || !subject || !body) {
      throw new McpError(ErrorCode.InvalidParams, "to, subject, and body are required");
    }

    const toRecipients = to.map((email) => ({ emailAddress: { address: email } }));
    const ccRecipients = cc.map((email) => ({ emailAddress: { address: email } }));
    const bccRecipients = bcc.map((email) => ({ emailAddress: { address: email } }));

    const payload = {
      message: {
        subject,
        body: { contentType: bodyType, content: body },
        toRecipients,
        ccRecipients,
        bccRecipients,
      },
    };

    await this.graphRequestAsUser("/me/sendMail", "POST", payload);

    return {
      content: [
        {
          type: "text",
          text: JSON.stringify({
            success: true,
            message: `Email sent to ${to.join(", ")}`,
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

    const endpoint = replyAll
      ? `/me/messages/${messageId}/replyAll`
      : `/me/messages/${messageId}/reply`;

    await this.graphRequestAsUser(endpoint, "POST", { comment });

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
