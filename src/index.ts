import express, { type Request, type Response } from "express";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { createServer } from "./server.js";
import { requireAccessToken } from "./auth.js";
import { jsonRpcError, messagesOf, responseIdFor } from "./jsonrpc.js";

const PORT = Number(process.env.PORT) || 8000;
const MCP_PATH = "/mcp";

const app = express();
app.use(express.json({ limit: "10mb" }));

app.get("/healthz", (_req, res) => {
  res.json({ status: "ok" });
});
app.get("/health", (_req, res) => {
  res.json({ status: "ok" });
});

app.post(MCP_PATH, requireAccessToken, async (req: Request, res: Response) => {
  // A fresh McpServer AND transport per request. The MCP Protocol allows only
  // one transport at a time, so a shared module-level server throws
  // "Already connected to a transport" as soon as two tool calls overlap.
  const server = createServer();
  const transport = new StreamableHTTPServerTransport({
    sessionIdGenerator: undefined,
  });

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
      res
        .status(500)
        .json(
          jsonRpcError(
            -32603,
            err instanceof Error ? err.message : "Internal error",
            responseIdFor(messagesOf(req.body)),
          ),
        );
    }
  }
});

app.listen(PORT, "0.0.0.0", () => {
  console.log(`[gsheets-hosted] listening on 0.0.0.0:${PORT}${MCP_PATH}`);
});
