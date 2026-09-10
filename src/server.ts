import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { tools } from "./tools/index.js";
import { grantedScopes, isToolGranted } from "./scopes.js";

const SERVER_NAME = "Google Sheets";
const SERVER_VERSION = "0.1.0";

export function toolSurface(granted: Set<string> | null) {
  const registered: string[] = [];
  const skipped: string[] = [];
  for (const [toolName, toolConfig] of Object.entries(tools)) {
    (isToolGranted((toolConfig as any).handler?.scope, granted) ? registered : skipped).push(toolName);
  }
  return { registered, skipped };
}

export function createServer(granted = grantedScopes()): McpServer {
  const server = new McpServer({ name: SERVER_NAME, version: SERVER_VERSION });
  const { registered } = toolSurface(granted);

  for (const toolName of registered) {
    const t = (tools as any)[toolName];
    server.registerTool(
      toolName,
      {
        description: t.description,
        inputSchema: t.schema,
        outputSchema: t.outputSchema,
        annotations: {
          readOnlyHint: t.readOnlyHint ?? false,
          destructiveHint: t.destructiveHint ?? false,
        },
      },
      async (args: Record<string, unknown>) => t.handler(args),
    );
  }

  return server;
}
