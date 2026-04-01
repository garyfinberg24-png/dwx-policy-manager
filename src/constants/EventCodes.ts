/**
 * Event Code Definitions for DWx Event Viewer
 * Each code has RegExp patterns for auto-classification, severity, and category.
 */

import {
  EventSeverity,
  EventClassification,
  IEventCodeDefinition,
} from '../models/IEventViewer';

// ============================================================================
// APPLICATION EVENT CODES (APP-001 to APP-020)
// ============================================================================

const APP_CODES: IEventCodeDefinition[] = [
  {
    code: 'APP-001',
    patterns: [/ErrorBoundary/i, /Unhandled render error/i, /componentDidCatch/i],
    severity: EventSeverity.Critical,
    category: EventClassification.Bug,
    description: 'Component render error caught by ErrorBoundary',
    suggestedAction: 'Check stack trace for null/undefined references in render methods',
  },
  {
    code: 'APP-002',
    patterns: [/Failed to initialize/i, /Service init.*fail/i],
    severity: EventSeverity.Error,
    category: EventClassification.Bug,
    description: 'Service initialisation failure',
    suggestedAction: 'Verify SharePoint list exists and user has access',
  },
  {
    code: 'APP-003',
    patterns: [/Failed to create policy/i, /Failed to update policy/i, /Failed to delete policy/i],
    severity: EventSeverity.Error,
    category: EventClassification.Bug,
    description: 'Policy CRUD operation failure',
    suggestedAction: 'Check SP list schema matches service code',
  },
  {
    code: 'APP-004',
    patterns: [/Failed to load/i, /Failed to fetch.*data/i, /getData.*fail/i],
    severity: EventSeverity.Error,
    category: EventClassification.Bug,
    description: 'Data loading failure',
    suggestedAction: 'Check network tab for failed requests, verify list permissions',
  },
  {
    code: 'APP-005',
    patterns: [/notification.*fail/i, /Failed to send notification/i],
    severity: EventSeverity.Warning,
    category: EventClassification.External,
    description: 'Notification delivery failure (non-critical)',
    suggestedAction: 'Check Logic App status and email queue',
  },
  {
    code: 'APP-006',
    patterns: [/audit.*fail/i, /Failed to log audit/i],
    severity: EventSeverity.Warning,
    category: EventClassification.Bug,
    description: 'Audit log write failure',
    suggestedAction: 'Verify PM_PolicyAuditLog list exists and column names match',
  },
  {
    code: 'APP-007',
    patterns: [/JSON\.parse/i, /Unexpected token/i, /JSON.*invalid/i],
    severity: EventSeverity.Error,
    category: EventClassification.Bug,
    description: 'JSON parse error on external data',
    suggestedAction: 'Check data source for malformed JSON, ensure try/catch guards',
  },
  {
    code: 'APP-008',
    patterns: [/validation.*fail/i, /invalid.*data/i, /required field/i],
    severity: EventSeverity.Warning,
    category: EventClassification.Configuration,
    description: 'Data validation failure',
    suggestedAction: 'Review input data against schema requirements',
  },
  {
    code: 'APP-009',
    patterns: [/DWx.*fail/i, /Hub.*unavailable/i, /cross-app.*fail/i],
    severity: EventSeverity.Warning,
    category: EventClassification.External,
    description: 'DWx Hub integration failure',
    suggestedAction: 'DWx Hub may be unavailable — app continues standalone',
  },
  {
    code: 'APP-010',
    patterns: [/permission.*denied/i, /access.*denied/i, /unauthorized/i],
    severity: EventSeverity.Warning,
    category: EventClassification.Security,
    description: 'Permission or access denied',
    suggestedAction: 'Verify user role and SP group membership',
  },
  {
    code: 'APP-020',
    patterns: [/setState.*unmount/i, /_isMounted.*guard/i, /after unmount/i],
    severity: EventSeverity.Warning,
    category: EventClassification.Bug,
    description: 'setState called after component unmount',
    suggestedAction: 'Ensure _isMounted guard covers all async callbacks',
  },
];

// ============================================================================
// NETWORK EVENT CODES (NET-001 to NET-020)
// ============================================================================

const NET_CODES: IEventCodeDefinition[] = [
  {
    code: 'NET-000',
    patterns: [/^(GET|POST|PATCH|DELETE|PUT)\s.*\s(200|201|204)$/i],
    severity: EventSeverity.Verbose,
    category: EventClassification.Unknown,
    description: 'Successful HTTP request',
  },
  {
    code: 'NET-001',
    patterns: [/slow.*request/i, /threshold.*exceeded/i],
    severity: EventSeverity.Warning,
    category: EventClassification.Performance,
    description: 'Slow API request (above threshold)',
    suggestedAction: 'Check SP list indexing, reduce $top/payload size',
  },
  {
    code: 'NET-002',
    patterns: [/\b400\b/, /Bad Request/i],
    severity: EventSeverity.Error,
    category: EventClassification.Bug,
    description: 'HTTP 400 Bad Request',
    suggestedAction: 'Check request payload and SP column names',
  },
  {
    code: 'NET-003',
    patterns: [/\b401\b/, /Authentication Required/i],
    severity: EventSeverity.Error,
    category: EventClassification.Security,
    description: 'HTTP 401 Authentication Required',
    suggestedAction: 'User session may have expired — page refresh required',
  },
  {
    code: 'NET-004',
    patterns: [/\b403\b/, /Forbidden/i, /Access Denied/i],
    severity: EventSeverity.Error,
    category: EventClassification.Security,
    description: 'HTTP 403 Forbidden',
    suggestedAction: 'User lacks permissions on the target resource',
  },
  {
    code: 'NET-005',
    patterns: [/\b404\b/, /Not Found/i],
    severity: EventSeverity.Warning,
    category: EventClassification.Configuration,
    description: 'HTTP 404 Not Found',
    suggestedAction: 'Verify list/endpoint exists, check URL construction',
  },
  {
    code: 'NET-006',
    patterns: [/\b409\b/, /Conflict/i],
    severity: EventSeverity.Warning,
    category: EventClassification.Bug,
    description: 'HTTP 409 Conflict',
    suggestedAction: 'Item may have been modified by another user',
  },
  {
    code: 'NET-010',
    patterns: [/\b429\b/, /Too Many Requests/i, /throttl/i],
    severity: EventSeverity.Error,
    category: EventClassification.Performance,
    description: 'HTTP 429 Throttled — SharePoint rate limit exceeded',
    suggestedAction: 'Reduce request frequency, implement request deduplication',
  },
  {
    code: 'NET-011',
    patterns: [/\b500\b/, /Internal Server Error/i],
    severity: EventSeverity.Error,
    category: EventClassification.External,
    description: 'HTTP 500 Internal Server Error',
    suggestedAction: 'Server-side issue — retry or check SharePoint health',
  },
  {
    code: 'NET-012',
    patterns: [/\b502\b|\b503\b/, /Service Unavailable/i, /Bad Gateway/i],
    severity: EventSeverity.Error,
    category: EventClassification.External,
    description: 'HTTP 502/503 Service Unavailable',
    suggestedAction: 'SharePoint or Azure Function may be temporarily unavailable',
  },
  {
    code: 'NET-020',
    patterns: [/timeout/i, /abort/i, /AbortError/i],
    severity: EventSeverity.Error,
    category: EventClassification.Performance,
    description: 'Request timeout or aborted',
    suggestedAction: 'Increase timeout or reduce payload size',
  },
  {
    code: 'NET-021',
    patterns: [/Failed to fetch/i, /NetworkError/i, /network.*fail/i, /ERR_NETWORK/i],
    severity: EventSeverity.Error,
    category: EventClassification.External,
    description: 'Network connectivity failure',
    suggestedAction: 'Check network connection, VPN, or CORS configuration',
  },
  {
    code: 'NET-022',
    patterns: [/CORS/i, /cross-origin/i, /blocked by.*policy/i],
    severity: EventSeverity.Error,
    category: EventClassification.Configuration,
    description: 'CORS policy blocked request',
    suggestedAction: 'Check Azure Function CORS settings or API permissions',
  },
];

// ============================================================================
// CONSOLE EVENT CODES (CON-001 to CON-005)
// ============================================================================

const CON_CODES: IEventCodeDefinition[] = [
  {
    code: 'CON-001',
    patterns: [/.*/],
    severity: EventSeverity.Error,
    category: EventClassification.Bug,
    description: 'Console error output',
    suggestedAction: 'Review error message and stack trace',
  },
  {
    code: 'CON-002',
    patterns: [/.*/],
    severity: EventSeverity.Warning,
    category: EventClassification.Bug,
    description: 'Console warning output',
  },
  {
    code: 'CON-003',
    patterns: [/.*/],
    severity: EventSeverity.Information,
    category: EventClassification.Unknown,
    description: 'Console info output',
  },
  {
    code: 'CON-004',
    patterns: [/.*/],
    severity: EventSeverity.Verbose,
    category: EventClassification.Unknown,
    description: 'Console debug output',
  },
  {
    code: 'CON-005',
    patterns: [/unhandled.*rejection/i, /uncaught.*exception/i],
    severity: EventSeverity.Critical,
    category: EventClassification.Bug,
    description: 'Unhandled promise rejection or uncaught exception',
    suggestedAction: 'Add error handling to the async operation',
  },
];

// ============================================================================
// SECURITY EVENT CODES (SEC-001 to SEC-005)
// ============================================================================

const SEC_CODES: IEventCodeDefinition[] = [
  {
    code: 'SEC-001',
    patterns: [/UnauthorizedAccess/i, /unauthorized.*attempt/i],
    severity: EventSeverity.Warning,
    category: EventClassification.Security,
    description: 'Unauthorized access attempt logged by audit service',
    suggestedAction: 'Review user permissions and role assignments',
  },
  {
    code: 'SEC-002',
    patterns: [/ValidationFailed/i, /OData.*injection/i, /sanitize/i],
    severity: EventSeverity.Warning,
    category: EventClassification.Security,
    description: 'Input validation failure',
    suggestedAction: 'Review input sanitisation and OData filters',
  },
  {
    code: 'SEC-003',
    patterns: [/SuspiciousActivity/i],
    severity: EventSeverity.Error,
    category: EventClassification.Security,
    description: 'Suspicious activity detected',
    suggestedAction: 'Review audit log for details',
  },
  {
    code: 'SEC-004',
    patterns: [/PII.*exposure/i, /REDACTED/i],
    severity: EventSeverity.Warning,
    category: EventClassification.Security,
    description: 'PII exposure or redaction event',
    suggestedAction: 'Verify PII redaction is working correctly',
  },
];

// ============================================================================
// SYSTEM EVENT CODES (SYS-001 to SYS-005)
// ============================================================================

const SYS_CODES: IEventCodeDefinition[] = [
  {
    code: 'SYS-001',
    patterns: [/list.*not.*found/i, /not.*provisioned/i, /list.*missing/i],
    severity: EventSeverity.Error,
    category: EventClassification.Configuration,
    description: 'SharePoint list not provisioned',
    suggestedAction: 'Run provisioning script for the missing list',
  },
  {
    code: 'SYS-002',
    patterns: [/column.*not.*found/i, /field.*missing/i, /schema.*mismatch/i],
    severity: EventSeverity.Error,
    category: EventClassification.Configuration,
    description: 'SharePoint list column mismatch',
    suggestedAction: 'Check provisioning script for missing columns',
  },
  {
    code: 'SYS-003',
    patterns: [/configuration.*missing/i, /config.*not.*found/i],
    severity: EventSeverity.Warning,
    category: EventClassification.Configuration,
    description: 'Missing application configuration',
    suggestedAction: 'Set the configuration value in Admin Centre',
  },
  {
    code: 'SYS-004',
    patterns: [/version.*mismatch/i, /compatibility/i],
    severity: EventSeverity.Warning,
    category: EventClassification.Configuration,
    description: 'Version or compatibility issue',
    suggestedAction: 'Verify component versions are compatible',
  },
  {
    code: 'SYS-005',
    patterns: [/signalAppReady/i, /app.*ready/i],
    severity: EventSeverity.Information,
    category: EventClassification.Unknown,
    description: 'Application ready signal',
  },
];

// ============================================================================
// DLQ EVENT CODES (DLQ-001 to DLQ-003)
// ============================================================================

const DLQ_CODES: IEventCodeDefinition[] = [
  {
    code: 'DLQ-001',
    patterns: [/notification.*dead.*letter/i, /notification.*dlq/i, /notification.*failed.*retries/i],
    severity: EventSeverity.Error,
    category: EventClassification.External,
    description: 'Notification delivery permanently failed — moved to DLQ',
    suggestedAction: 'Check Logic App status and recipient validity',
  },
  {
    code: 'DLQ-002',
    patterns: [/sync.*dead.*letter/i, /sync.*dlq/i, /workflow.*sync.*fail/i],
    severity: EventSeverity.Error,
    category: EventClassification.Bug,
    description: 'Workflow sync permanently failed — moved to DLQ',
    suggestedAction: 'Check SP list schema and data integrity',
  },
  {
    code: 'DLQ-003',
    patterns: [/approval.*dead.*letter/i, /approval.*dlq/i],
    severity: EventSeverity.Error,
    category: EventClassification.Bug,
    description: 'Approval operation permanently failed — moved to DLQ',
    suggestedAction: 'Check approval chain configuration',
  },
];

// ============================================================================
// COMBINED CODE REGISTRY
// ============================================================================

/**
 * All event code definitions in priority order.
 * More specific codes should appear before generic ones.
 * CON-001..004 are catch-all codes for console levels (wildcard pattern)
 * and should be matched by console level, not pattern.
 */
export const EVENT_CODES: IEventCodeDefinition[] = [
  // Specific codes first (order matters — first match wins)
  ...SEC_CODES,
  ...DLQ_CODES,
  ...SYS_CODES,
  ...APP_CODES,
  ...NET_CODES,
  // Console codes last (catch-all patterns)
  ...CON_CODES,
];

/**
 * Lookup a single event code definition by code string
 */
export function getEventCodeDefinition(code: string): IEventCodeDefinition | undefined {
  return EVENT_CODES.find(ec => ec.code === code);
}

/**
 * Get all codes for a given category
 */
export function getCodesByCategory(category: EventClassification): IEventCodeDefinition[] {
  return EVENT_CODES.filter(ec => ec.category === category);
}

/**
 * Slow request threshold in milliseconds for NET-001 classification
 */
export const SLOW_REQUEST_THRESHOLD_MS = 2000;

/**
 * Telemetry endpoint patterns to exclude from network interception
 */
export const TELEMETRY_EXCLUSION_PATTERNS: RegExp[] = [
  /dc\.services\.visualstudio\.com/i,
  /browser\.events\.data\.microsoft\.com/i,
  /browser\.pipe\.aria\.microsoft\.com/i,
  /\.clarity\.ms/i,
];

/**
 * Asset file extensions to tag as CDN/asset requests
 */
export const ASSET_EXTENSIONS: RegExp = /\.(js|css|png|jpg|jpeg|gif|svg|woff|woff2|ttf|eot|ico|map)(\?|$)/i;
