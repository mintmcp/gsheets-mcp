import express, { type Request, type Response } from "express";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { createServer } from "./server.js";
import { requestContext } from "./auth.js";

export const MCP_PATH = "/mcp";

export function createApp() {
  const app = express();
  app.use(express.json({ limit: "10mb" }));

  app.get("/healthz", (_req, res) => {
    res.json({ status: "ok" });
  });
  app.get("/health", (_req, res) => {
    res.json({ status: "ok" });
  });

  app.post(MCP_PATH, async (req: Request, res: Response) => {
    const authHeader = req.header("authorization") ?? req.header("Authorization") ?? "";
    const accessToken = authHeader.startsWith("Bearer ")
      ? authHeader.slice("Bearer ".length).trim()
      : "";

    // A fresh server AND transport per request: the MCP Protocol holds one
    // transport at a time, so a shared server drops overlapping calls with
    // "Already connected to a transport".
    const server = createServer();
    const transport = new StreamableHTTPServerTransport({
      sessionIdGenerator: undefined,
    });

    requestContext.run({ accessToken }, async () => {
      try {
        res.on("close", () => {
          transport.close().catch(() => {});
          server.close().catch(() => {});
        });
        await server.connect(transport);
        await transport.handleRequest(req, res, req.body);
      } catch (err) {
        console.error("[gsheets-hosted] MCP request error:", err);
        if (!res.headersSent) {
          res.status(500).json({
            jsonrpc: "2.0",
            error: {
              code: -32603,
              message: err instanceof Error ? err.message : "Internal error",
            },
            id: null,
          });
        }
      }
    });
  });

  return app;
}
