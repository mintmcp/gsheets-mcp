import { readTools } from './read.js';
import { writeTools } from './write.js';
import { formatTools } from './format.js';

export const tools = { ...readTools, ...writeTools, ...formatTools };
export type ToolMap = typeof tools;
