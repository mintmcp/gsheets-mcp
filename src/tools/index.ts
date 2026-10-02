import { readTools } from './read.js';
import { writeTools } from './write.js';
import { formatTools } from './format.js';
import { chartTools } from './charts.js';

export const tools = { ...readTools, ...writeTools, ...formatTools, ...chartTools };
export type ToolMap = typeof tools;
