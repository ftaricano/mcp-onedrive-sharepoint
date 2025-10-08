/**
 * Advanced tools index
 * Exports all advanced tool categories
 */

// Import all advanced tool modules
import { collaborationTools, collaborationHandlers } from './collaboration.js';
import { syncTools, syncHandlers } from './sync.js';
import { analyticsTools, analyticsHandlers } from './analytics.js';
import { excelTools, excelHandlers } from './excel.js';

// Export all advanced tools
export const advancedTools = [
  ...collaborationTools,
  ...syncTools,
  ...analyticsTools,
  ...excelTools
];

// Export all advanced handlers
export const advancedHandlers = {
  ...collaborationHandlers,
  ...syncHandlers,
  ...analyticsHandlers,
  ...excelHandlers
};

// Export individual categories for selective imports
export {
  collaborationTools,
  collaborationHandlers,
  syncTools,
  syncHandlers,
  analyticsTools,
  analyticsHandlers,
  excelTools,
  excelHandlers
};