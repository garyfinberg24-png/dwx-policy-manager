/**
 * EventViewer shared style constants — Forest Teal design system
 */

export const Colors = {
  // Primary palette
  tealPrimary: '#0d9488',
  tealDark: '#0f766e',
  tealLight: '#ccfbf1',
  tealPale: '#f0fdfa',

  // Text
  textPrimary: '#0f172a',
  textSecondary: '#334155',
  textMuted: '#64748b',
  textSlate: '#94a3b8',

  // Borders
  border: '#e2e8f0',
  borderLight: '#f1f5f9',

  // Backgrounds
  bgPage: '#f1f5f9',
  bgCard: '#ffffff',
  bgHeader: '#f8fafc',
  bgSidebar: '#f1f5f9',

  // Status
  success: '#059669',
  successLight: '#d1fae5',
  warning: '#d97706',
  warningLight: '#fef3c7',
  error: '#dc2626',
  errorLight: '#fee2e2',
  critical: '#7f1d1d',

  // AI / Purple
  aiPrimary: '#7c3aed',
  aiDark: '#6d28d9',
  aiDeep: '#4c1d95',
  aiLight: '#f5f3ff',
  aiBorder: '#ddd6fe',
  aiPale: '#ede9fe',

  // Additional
  blue: '#2563eb',
  blueLight: '#dbeafe',
  purple: '#7c3aed',
  slate: '#475569',
};

/** Severity → colour mapping */
export const SeverityColors: Record<string, { bg: string; text: string; border: string }> = {
  Verbose: { bg: '#f1f5f9', text: '#94a3b8', border: '#e2e8f0' },
  Information: { bg: '#f0f9ff', text: '#2563eb', border: '#bfdbfe' },
  Warning: { bg: '#fffbeb', text: '#d97706', border: '#fde68a' },
  Error: { bg: '#fef2f2', text: '#dc2626', border: '#fecaca' },
  Critical: { bg: '#7f1d1d', text: '#ffffff', border: '#991b1b' },
};

/** Channel → colour mapping */
export const ChannelColors: Record<string, { bg: string; text: string }> = {
  Application: { bg: '#ede9fe', text: '#7c3aed' },
  Console: { bg: '#e0f2fe', text: '#0284c7' },
  Network: { bg: '#fef3c7', text: '#b45309' },
  Audit: { bg: '#d1fae5', text: '#047857' },
  DLQ: { bg: '#fee2e2', text: '#b91c1c' },
  System: { bg: '#f1f5f9', text: '#475569' },
};

/** HTTP method → colour mapping */
export const MethodColors: Record<string, { bg: string; text: string }> = {
  GET: { bg: '#dbeafe', text: '#1d4ed8' },
  POST: { bg: '#d1fae5', text: '#047857' },
  PATCH: { bg: '#fef3c7', text: '#b45309' },
  PUT: { bg: '#fef3c7', text: '#b45309' },
  DELETE: { bg: '#fee2e2', text: '#b91c1c' },
};

/** Health status → indicator colour */
export const HealthColors: Record<string, string> = {
  Healthy: '#22c55e',
  Degraded: '#f59e0b',
  Unhealthy: '#ef4444',
};

/** Tab definitions for the Event Viewer */
export const EVENT_VIEWER_TABS = [
  { key: 'stream', label: 'Event Stream' },
  { key: 'network', label: 'Network Monitor' },
  { key: 'investigate', label: 'Investigation Board' },
  { key: 'health', label: 'System Health' },
  { key: 'ai', label: 'AI Triage' },
  { key: 'performance', label: 'Performance' },
  { key: 'troubleshooter', label: 'Troubleshooter' },
] as const;

export type EventViewerTabKey = typeof EVENT_VIEWER_TABS[number]['key'];
