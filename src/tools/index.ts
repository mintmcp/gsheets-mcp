import { readTools } from './read.js';
import { writeTools } from './write.js';
import { formatTools } from './format.js';
import { chartTools } from './charts.js';
import { structureTools } from './structure.js';
import { rangeTools } from './ranges.js';

export const tools = { ...readTools, ...writeTools, ...formatTools, ...chartTools, ...structureTools, ...rangeTools };
export type ToolMap = typeof tools;
