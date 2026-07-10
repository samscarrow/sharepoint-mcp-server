/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 */

/**
 * Streamable HTTP entry point.
 *
 * Serves the same MCP server as `index.ts` (stdio) over the transport defined in
 * MCP 2025-06-18, so no stdio->HTTP bridge process is needed.
 *
 * Stateless mode (`sessionIdGenerator: undefined`): the server never issues an
 * `Mcp-Session-Id`, so clients are never required to send one. This matches how the
 * claude.ai connector actually behaves — it POSTs requests without a session id and
 * without a preceding `initialize`. A *stateful* transport rejects those with
 * `400 Bad Request: No valid session ID provided`.
 *
 * Per the SDK's stateless guidance, each request gets its own Server + Transport so
 * concurrent requests cannot collide on JSON-RPC ids. These are plain in-process
 * objects: no subprocess is forked, and `res.on("close")` disposes both. Graph token
 * caches are module-scoped in index.ts, so a fresh Server does not re-authenticate.
 */

import express from "express";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { SharePointServer } from "./index.js";

const PORT = Number(process.env.PORT ?? 3000);
const MCP_PATH = process.env.MCP_PATH ?? "/mcp";

const app = express();
app.use(express.json({ limit: "8mb" }));

app.get("/healthz", (_req, res) => {
  res.status(200).send("ok");
});

app.post(MCP_PATH, async (req, res) => {
  const server = new SharePointServer();
  const transport = new StreamableHTTPServerTransport({
    sessionIdGenerator: undefined,
  });

  // Dispose both when the response ends. Without this the transport is never closed
  // and the per-request Server is retained.
  res.on("close", () => {
    void transport.close();
    void server.close();
  });

  try {
    await server.connect(transport);
    await transport.handleRequest(req, res, req.body);
  } catch (error) {
    console.error("Error handling MCP request:", error);
    if (!res.headersSent) {
      res.status(500).json({
        jsonrpc: "2.0",
        error: { code: -32603, message: "Internal server error" },
        id: null,
      });
    }
  }
});

// GET/DELETE on the MCP endpoint are only meaningful for stateful sessions.
// Per spec a server that does not offer them replies 405.
app.get(MCP_PATH, (_req, res) => res.status(405).send("Method Not Allowed"));
app.delete(MCP_PATH, (_req, res) => res.status(405).send("Method Not Allowed"));

const httpServer = app.listen(PORT, () => {
  console.error(`SharePoint MCP server running on http://0.0.0.0:${PORT}${MCP_PATH}`);
});

for (const signal of ["SIGTERM", "SIGINT"] as const) {
  process.on(signal, () => {
    console.error(`Received ${signal}, shutting down`);
    httpServer.close(() => process.exit(0));
    setTimeout(() => process.exit(1), 5000).unref();
  });
}
