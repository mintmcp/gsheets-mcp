import { readTools } from './read.js';
import { writeTools } from './write.js';
import { formatTools } from './format.js';

export type ToolMap = typeof readTools & typeof writeTools & typeof formatTools;

/**
 * Merging three modules into one namespace means a duplicated tool name would
 * silently overwrite rather than fail, so uniqueness is asserted once at
 * module load. A server is built per request, so doing it inside getTools()
 * re-ran this scan on every call.
 */
function mergeTools(): ToolMap {
  const merged: Record<string, unknown> = {};
  for (const group of [readTools, writeTools, formatTools]) {
    for (const [name, tool] of Object.entries(group)) {
      if (name in merged) {
        throw new Error(`Duplicate tool name "${name}" across tool modules`);
      }
      merged[name] = tool;
    }
  }
  return merged as ToolMap;
}

export const tools = mergeTools();
