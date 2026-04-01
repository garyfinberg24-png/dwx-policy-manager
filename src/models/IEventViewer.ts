/**
 * Event Viewer Models
 * Interfaces, enums, and configuration for the DWx Event Viewer diagnostic tool.
 */

// ============================================================================
// ENUMS
// ============================================================================

/**
 * Event severity levels — aligned with LoggingService.SeverityLevel
 */
export enum EventSeverity {
  Verbose = 0,
  Information = 1,
  Warning = 2,
  Error = 3,
  Critical = 4,
}

/**
 * Event data channels — where the event originated
 */
export enum EventChannel {
  Application = 'Application',
  Console = 'Console',
  Network = 'Network',
  Audit = 'Audit',
  DLQ = 'DLQ',
  System = 'System',
}

/**
 * Console event sub-origins for granular filtering
 */
export enum ConsoleOrigin {
  App = 'app',
  Framework = 'framework',
  Library = 'library',
  React = 'react',
  Browser = 'browser',
}

/**
 * Investigation classification categories
 */
export enum EventClassification {
  Bug = 'Bug',
  Performance = 'Performance',
  Security = 'Security',
  Configuration = 'Configuration',
  External = 'External',
  Unknown = 'Unknown',
}

// ============================================================================
// CORE EVENT INTERFACES
// ============================================================================

/**
 * Core event entry — the fundamental unit stored in the ring buffer
 */
export interface IEventEntry {
  /** Unique event ID (generated: `evt_${timestamp}_${random}`) */
  id: string;
  /** ISO timestamp of when the event occurred */
  timestamp: string;
  /** Severity level */
  severity: EventSeverity;
  /** Data channel (Application, Console, Network, etc.) */
  channel: EventChannel;
  /** Source component or service name (e.g. 'PolicyService', 'ErrorBoundary') */
  source: string;
  /** Human-readable event message */
  message: string;
  /** Classified event code (e.g. 'APP-001', 'NET-010') — set by EventClassifier */
  eventCode?: string;
  /** Error stack trace (if applicable) */
  stackTrace?: string;
  /** Additional metadata as key-value pairs */
  metadata?: Record<string, unknown>;
  /** Session ID for cross-event correlation */
  sessionId?: string;
  /** Page URL where event occurred */
  url?: string;
  /** Whether this event has been persisted to PM_EventLog */
  persisted?: boolean;
  /** SP list item ID if persisted */
  persistedItemId?: number;
  /** Whether auto-persist is flagged (Error/Critical severity) */
  autoPersist?: boolean;
  /** Console sub-origin for Console channel events */
  consoleOrigin?: ConsoleOrigin;
}

/**
 * Network event — extends IEventEntry with HTTP-specific fields
 */
export interface INetworkEvent extends IEventEntry {
  /** Request URL */
  requestUrl: string;
  /** HTTP method (GET, POST, PATCH, DELETE) */
  httpMethod: string;
  /** HTTP status code */
  httpStatus?: number;
  /** Request duration in milliseconds */
  duration?: number;
  /** Response size in bytes (from Content-Length header) */
  responseSize?: number;
  /** SharePoint list name extracted from URL (e.g. 'PM_Policies') */
  spListName?: string;
  /** Whether this is a CDN/static asset request */
  isAssetRequest?: boolean;
}

// ============================================================================
// FILTER & QUERY INTERFACES
// ============================================================================

/**
 * Filter criteria for querying events
 */
export interface IEventFilter {
  /** Filter by channel(s) */
  channels?: EventChannel[];
  /** Filter by severity level(s) */
  severities?: EventSeverity[];
  /** Filter by minimum severity (inclusive) */
  minSeverity?: EventSeverity;
  /** Filter by source component/service */
  source?: string;
  /** Full-text search across message, source, eventCode, URL */
  searchText?: string;
  /** Filter by event code(s) */
  eventCodes?: string[];
  /** Filter by time range — start (ISO string) */
  startTime?: string;
  /** Filter by time range — end (ISO string) */
  endTime?: string;
  /** Filter by SP list name (network events only) */
  spListName?: string;
  /** Filter by HTTP status range (network events only) */
  httpStatusMin?: number;
  httpStatusMax?: number;
  /** Filter by minimum duration in ms (network events only) */
  minDuration?: number;
  /** Include/exclude asset requests (network events) */
  includeAssets?: boolean;
  /** Filter by classification */
  classification?: EventClassification;
  /** Show only investigated/uninvestigated */
  isInvestigated?: boolean;
  /** Show only persisted events */
  persistedOnly?: boolean;
}

// ============================================================================
// PERSISTENCE INTERFACES (PM_EventLog SharePoint list)
// ============================================================================

/**
 * Shape of an event record persisted to PM_EventLog SharePoint list
 */
export interface IPersistedEvent {
  Id?: number;
  Title: string;
  EventCode?: string;
  Severity: string;
  Channel: string;
  Source: string;
  Message: string;
  StackTrace?: string;
  CorrelationId?: string;
  SessionId?: string;
  UserLogin?: string;
  EventTimestamp: string;
  Duration?: number;
  Url?: string;
  HttpMethod?: string;
  HttpStatus?: number;
  InvestigationNotes?: string;
  Classification?: string;
  IsInvestigated?: boolean;
  AutoPersisted?: boolean;
  Metadata?: string;
}

// ============================================================================
// EVENT CODE / CLASSIFICATION
// ============================================================================

/**
 * Definition for a classified event code
 */
export interface IEventCodeDefinition {
  /** Event code (e.g. 'APP-001') */
  code: string;
  /** RegExp patterns that match this code */
  patterns: RegExp[];
  /** Default severity for this code */
  severity: EventSeverity;
  /** Classification category */
  category: EventClassification;
  /** Human-readable description */
  description: string;
  /** Suggested action text */
  suggestedAction?: string;
}

/**
 * Result from the EventClassifier
 */
export interface IEventClassificationResult {
  eventCode: string;
  category: EventClassification;
  description: string;
  suggestedAction?: string;
}

// ============================================================================
// INVESTIGATION BOARD
// ============================================================================

/**
 * Grouped event summary for the Investigation Board
 */
export interface IEventGroup {
  /** Event code */
  eventCode: string;
  /** Description from EventCodeDefinition */
  description: string;
  /** Severity of the code */
  severity: EventSeverity;
  /** Total occurrences */
  count: number;
  /** First occurrence timestamp */
  firstSeen: string;
  /** Last occurrence timestamp */
  lastSeen: string;
  /** Number of unique sessions affected */
  sessionsAffected?: number;
  /** Classification assigned by admin */
  classification?: EventClassification;
  /** Whether marked as investigated */
  isInvestigated?: boolean;
  /** Investigation notes */
  notes?: string;
  /** Sparkline data — occurrence counts per time bucket */
  sparklineData?: number[];
  /** Individual events in this group */
  events: IEventEntry[];
}

// ============================================================================
// SYSTEM HEALTH
// ============================================================================

/**
 * Health status for a service
 */
export enum HealthStatus {
  Healthy = 'Healthy',
  Degraded = 'Degraded',
  Unhealthy = 'Unhealthy',
}

/**
 * Service health card data
 */
export interface IServiceHealth {
  /** Service name */
  name: string;
  /** Overall health status */
  status: HealthStatus;
  /** Total requests in the monitoring window */
  requestCount: number;
  /** Error count in the monitoring window */
  errorCount: number;
  /** Average latency in ms (network services) */
  avgLatency?: number;
  /** Success rate as percentage */
  successRate: number;
  /** Last error message */
  lastError?: string;
  /** Timestamp of last error */
  lastErrorTime?: string;
}

/**
 * Session information for the System Health tab
 */
export interface ISessionInfo {
  sessionId: string;
  userId: string;
  userRole: string;
  browser: string;
  startTime: string;
  currentPage: string;
  appVersion: string;
  appInsightsConnected: boolean;
  spSiteUrl: string;
}

// ============================================================================
// AI TRIAGE
// ============================================================================

/**
 * Request payload for AI event triage
 */
export interface IEventTriageRequest {
  /** Conversation mode — always 'event-triage' for Event Viewer */
  mode: 'event-triage';
  /** The question or instruction for AI */
  message: string;
  /** Event context to analyse */
  eventContext: {
    events: IEventEntry[];
    sessionInfo?: ISessionInfo;
    dlqStats?: {
      pending: number;
      processing: number;
      resolved: number;
      abandoned: number;
    };
    networkStats?: {
      totalRequests: number;
      avgLatency: number;
      errorRate: number;
      throttledCount: number;
    };
  };
  /** Previous conversation messages (for Ask AI) */
  conversationHistory?: Array<{ role: string; content: string }>;
  /** User's role */
  userRole: string;
}

/**
 * Response from AI event triage
 */
export interface IEventTriageResponse {
  /** AI-generated analysis text */
  analysis: string;
  /** Identified root causes */
  rootCauses?: IRootCause[];
  /** Overall session health assessment */
  healthAssessment?: string;
  /** Suggested follow-up questions */
  suggestedActions?: string[];
  /** Model confidence (0-100) */
  confidence?: number;
}

/**
 * Individual root cause identified by AI
 */
export interface IRootCause {
  /** Title of the root cause */
  title: string;
  /** Severity assessment */
  severity: 'Critical' | 'High' | 'Medium' | 'Low';
  /** Detailed explanation */
  explanation: string;
  /** Why it happened */
  whyItHappened?: string;
  /** Related event codes */
  relatedEventCodes: string[];
  /** Number of events attributed to this cause */
  affectedEventCount: number;
  /** Suggested fix steps */
  suggestedFix?: IFixStep[];
  /** Confidence score (0-100) */
  confidence: number;
  /** Classification */
  classification: EventClassification;
  /** Whether AI believes this is auto-fixable */
  autoFixable?: boolean;
}

/**
 * A suggested fix step from AI triage
 */
export interface IFixStep {
  /** Step number */
  step: number;
  /** Step title */
  title: string;
  /** Step description */
  description: string;
  /** Code snippet (if applicable) */
  codeSnippet?: string;
  /** File path where fix should be applied */
  filePath?: string;
}

/**
 * Root Cause Analysis report
 */
export interface IRCAReport {
  /** Report ID */
  reportId: string;
  /** Generation timestamp */
  generatedAt: string;
  /** Session ID analysed */
  sessionId: string;
  /** Session duration */
  sessionDuration: string;
  /** Total events analysed */
  totalEvents: number;
  /** Error count */
  errorCount: number;
  /** Executive summary paragraph */
  executiveSummary: string;
  /** Root causes identified */
  rootCauses: IRootCause[];
  /** Number of auto-fixable issues */
  autoFixableCount: number;
  /** Overall confidence score */
  overallConfidence: number;
  /** AI model used */
  modelUsed: string;
  /** Analyst notes (user-provided) */
  analystNotes?: string;
}

// ============================================================================
// CONFIGURATION
// ============================================================================

/**
 * Event Viewer admin configuration
 */
export interface IEventViewerConfig {
  /** Whether Event Viewer is enabled */
  enabled: boolean;
  /** Ring buffer size for application events */
  appBufferSize: number;
  /** Ring buffer size for console events */
  consoleBufferSize: number;
  /** Ring buffer size for network events */
  networkBufferSize: number;
  /** Minimum severity for auto-persistence ('Error' or 'Critical') */
  autoPersistThreshold: EventSeverity;
  /** Whether AI triage is enabled */
  aiTriageEnabled: boolean;
  /** Azure Function URL for AI triage */
  aiFunctionUrl: string;
  /** Event retention period in days */
  retentionDays: number;
  /** Whether to hide CDN/asset requests by default */
  hideCdnByDefault: boolean;
}

/**
 * Default configuration values
 */
export const DEFAULT_EVENT_VIEWER_CONFIG: IEventViewerConfig = {
  enabled: true,
  appBufferSize: 1000,
  consoleBufferSize: 500,
  networkBufferSize: 500,
  autoPersistThreshold: EventSeverity.Error,
  aiTriageEnabled: false,
  aiFunctionUrl: '',
  retentionDays: 90,
  hideCdnByDefault: true,
};

// ============================================================================
// EVENT BUFFER STATS
// ============================================================================

/**
 * Statistics from the EventBuffer
 */
export interface IEventBufferStats {
  /** Count of application channel events in buffer */
  appCount: number;
  /** Count of console channel events in buffer */
  consoleCount: number;
  /** Count of network channel events in buffer */
  networkCount: number;
  /** Total events across all buffers */
  totalCount: number;
  /** Count of error-severity events */
  errorCount: number;
  /** Count of warning-severity events */
  warningCount: number;
  /** Count of critical-severity events */
  criticalCount: number;
  /** Buffer capacity (max sizes) */
  capacity: {
    app: number;
    console: number;
    network: number;
  };
}

// ============================================================================
// PERFORMANCE OPTIMIZER
// ============================================================================

/**
 * Performance sub-score categories
 */
export interface IPerformanceSubScore {
  label: string;
  key: string;
  score: number;
  detail: string;
}

/**
 * Overall performance score
 */
export interface IPerformanceScore {
  overall: number;
  subScores: IPerformanceSubScore[];
  issueCount: number;
}

/**
 * A detected performance issue with tuneable controls
 */
export interface IPerformanceIssue {
  id: string;
  title: string;
  description: string;
  severity: 'high' | 'medium' | 'low';
  impactPercent: number;
  controls: IOptimizationControl[];
  prediction: string;
  applied: boolean;
  configKeys: Record<string, string>;
}

/**
 * A tuneable optimization control (slider or toggle)
 */
export interface IOptimizationControl {
  type: 'slider' | 'toggle';
  label: string;
  configKey: string;
  /** For sliders */
  min?: number;
  max?: number;
  step?: number;
  value: number | boolean;
  unit?: string;
  /** For toggles */
  onLabel?: string;
  offLabel?: string;
}

/**
 * Before/after comparison metrics
 */
export interface IPerformanceComparison {
  metric: string;
  current: string;
  projected: string;
  improved: boolean;
}

/**
 * AI performance recommendation
 */
export interface IAIPerformanceRecommendation {
  id: string;
  title: string;
  impact: 'high' | 'medium' | 'low';
  analysis: string;
  codeSnippet?: string;
  prediction: string;
  actionType: 'config' | 'script' | 'code';
  actionLabel: string;
  configKeys?: Record<string, string>;
  dismissed: boolean;
}
