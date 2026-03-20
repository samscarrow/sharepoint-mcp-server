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

/**
 * Environment variables required for SharePoint authentication
 */
const { SHAREPOINT_URL, TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;

if (!SHAREPOINT_URL || !TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
  throw new Error(
    "Required environment variables: SHAREPOINT_URL, TENANT_ID, CLIENT_ID, CLIENT_SECRET"
  );
}

const MAIL_USER_HARDCODED = "worshipservices@bayviewassociation.org";

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
          description: "Search for files and documents in SharePoint using Microsoft Graph Search API",
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
          description: "List document libraries (drives) in a SharePoint site",
          inputSchema: {
            type: "object",
            properties: {
              siteUrl: {
                type: "string",
                description: "The SharePoint site URL",
              },
            },
            required: ["siteUrl"],
          },
        },
        {
          name: "list_drive_items",
          description: "List files and folders in a SharePoint document library",
          inputSchema: {
            type: "object",
            properties: {
              siteUrl: {
                type: "string",
                description: "The SharePoint site URL",
              },
              driveId: {
                type: "string",
                description: "The drive ID (optional, uses default drive if not specified)",
              },
              folderPath: {
                type: "string",
                description: "Optional folder path to list items from (default: root)",
              },
            },
            required: ["siteUrl"],
          },
        },
        {
          name: "get_file_content",
          description: "Get the content of a specific file from SharePoint (text files only)",
          inputSchema: {
            type: "object",
            properties: {
              siteUrl: {
                type: "string",
                description: "The SharePoint site URL",
              },
              filePath: {
                type: "string",
                description: "The path to the file",
              },
              driveId: {
                type: "string",
                description: "The drive ID (optional, uses default drive if not specified)",
              },
            },
            required: ["siteUrl", "filePath"],
          },
        },
        {
          name: "list_emails",
          description: "List recent emails from the worshipservices@ mailbox. Returns subject, from, date, and preview.",
          inputSchema: {
            type: "object",
            properties: {
              folder: {
                type: "string",
                description: "Mail folder to read from (default: inbox). Options: inbox, sentitems, drafts, deleteditems, archive",
                default: "inbox",
              },
              top: {
                type: "number",
                description: "Number of emails to return (default: 10, max: 50)",
                default: 10,
              },
              filter: {
                type: "string",
                description: "OData filter expression (e.g. \"from/emailAddress/address eq 'someone@example.com'\" or \"isRead eq false\")",
              },
            },
          },
        },
        {
          name: "get_email",
          description: "Get the full content of a specific email by ID from the worshipservices@ mailbox",
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
          description: "Search emails by keyword in the worshipservices@ mailbox",
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

    if (typeof query !== "string") {
      throw new McpError(ErrorCode.InvalidParams, "Query parameter must be a string");
    }

    try {
      const searchRequest = {
        requests: [{
          entityTypes: ["driveItem"],
          query: {
            queryString: query,
          },
          region: "US",
          size: limit,
        }],
      };

      const searchResults = await this.graphRequest("/search/query", "POST", searchRequest);

      return {
        content: [{
          type: "text",
          text: JSON.stringify(searchResults, null, 2),
        }],
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
      const site = await this.graphRequest(`/sites/${siteId}?$expand=drive`);

      return {
        content: [{
          type: "text",
          text: JSON.stringify(site, null, 2),
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

    if (typeof siteUrl !== "string") {
      throw new McpError(ErrorCode.InvalidParams, "siteUrl parameter must be a string");
    }

    try {
      const siteId = await this.getSiteIdFromUrl(siteUrl);
      const response = await this.graphRequest(`/sites/${siteId}/drives`);
      const drives = response.value || [];

      return {
        content: [{
          type: "text",
          text: JSON.stringify(drives, null, 2),
        }],
      };
    } catch (error) {
      throw new Error(`Failed to list site drives: ${error}`);
    }
  }

  /**
   * Handle list drive items tool request
   */
  private async handleListDriveItems(args: any) {
    const siteUrl = args?.siteUrl;
    const driveId = args?.driveId;
    const folderPath = args?.folderPath;

    if (typeof siteUrl !== "string") {
      throw new McpError(ErrorCode.InvalidParams, "siteUrl parameter must be a string");
    }

    try {
      const siteId = await this.getSiteIdFromUrl(siteUrl);
      
      let endpoint: string;
      if (driveId) {
        if (folderPath) {
          endpoint = `/sites/${siteId}/drives/${driveId}/root:/${folderPath}:/children`;
        } else {
          endpoint = `/sites/${siteId}/drives/${driveId}/root/children`;
        }
      } else {
        if (folderPath) {
          endpoint = `/sites/${siteId}/drive/root:/${folderPath}:/children`;
        } else {
          endpoint = `/sites/${siteId}/drive/root/children`;
        }
      }

      const response = await this.graphRequest(endpoint);
      const items = response.value || [];

      return {
        content: [{
          type: "text",
          text: JSON.stringify(items, null, 2),
        }],
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

    if (typeof siteUrl !== "string" || typeof filePath !== "string") {
      throw new McpError(ErrorCode.InvalidParams, "siteUrl and filePath parameters must be strings");
    }

    try {
      const siteId = await this.getSiteIdFromUrl(siteUrl);
      
      let endpoint: string;
      if (driveId) {
        endpoint = `/sites/${siteId}/drives/${driveId}/root:/${filePath}:/content`;
      } else {
        endpoint = `/sites/${siteId}/drive/root:/${filePath}:/content`;
      }

      const token = await this.getAccessToken();
      const response = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, {
        headers: {
          "Authorization": `Bearer ${token}`,
        },
      });

      if (!response.ok) {
        throw new Error(`Failed to get file content: ${response.status} ${response.statusText}`);
      }

      const content = await response.text();

      return {
        content: [{
          type: "text",
          text: content,
        }],
      };
    } catch (error) {
      throw new Error(`Failed to get file content: ${error}`);
    }
  }

  /**
   * Handle list emails tool request
   */
  private async handleListEmails(args: any) {
    const user = MAIL_USER_HARDCODED;
    const folder = args?.folder || "inbox";
    const top = Math.min(args?.top || 10, 50);
    const filter = args?.filter;

    try {
      let endpoint = `/users/${user}/mailFolders/${folder}/messages?$top=${top}&$select=id,subject,from,toRecipients,receivedDateTime,isRead,bodyPreview&$orderby=receivedDateTime desc`;
      if (filter) {
        endpoint += `&$filter=${encodeURIComponent(filter)}`;
      }

      const response = await this.graphRequest(endpoint);
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
        content: [{
          type: "text",
          text: JSON.stringify(messages, null, 2),
        }],
      };
    } catch (error) {
      throw new Error(`Failed to list emails: ${error}`);
    }
  }

  /**
   * Handle get email tool request
   */
  private async handleGetEmail(args: any) {
    const user = MAIL_USER_HARDCODED;
    const messageId = args?.messageId;

    if (typeof messageId !== "string") {
      throw new McpError(ErrorCode.InvalidParams, "messageId parameter must be a string");
    }

    try {
      const message = await this.graphRequest(
        `/users/${user}/messages/${messageId}?$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,hasAttachments,importance`
      );

      return {
        content: [{
          type: "text",
          text: JSON.stringify({
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
          }, null, 2),
        }],
      };
    } catch (error) {
      throw new Error(`Failed to get email: ${error}`);
    }
  }

  /**
   * Handle search emails tool request
   */
  private async handleSearchEmails(args: any) {
    const user = MAIL_USER_HARDCODED;
    const query = args?.query;
    const top = Math.min(args?.top || 10, 50);

    if (typeof query !== "string") {
      throw new McpError(ErrorCode.InvalidParams, "query parameter must be a string");
    }

    try {
      const endpoint = `/users/${user}/messages?$search="${encodeURIComponent(query)}"&$top=${top}&$select=id,subject,from,receivedDateTime,bodyPreview`;

      const response = await this.graphRequest(endpoint);
      const messages = (response.value || []).map((m: any) => ({
        id: m.id,
        subject: m.subject,
        from: m.from?.emailAddress?.address,
        fromName: m.from?.emailAddress?.name,
        date: m.receivedDateTime,
        preview: m.bodyPreview,
      }));

      return {
        content: [{
          type: "text",
          text: JSON.stringify(messages, null, 2),
        }],
      };
    } catch (error) {
      throw new Error(`Failed to search emails: ${error}`);
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
