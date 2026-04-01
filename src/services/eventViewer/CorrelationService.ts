/**
 * CorrelationService — Groups related events into correlation chains.
 * Detects patterns like: policy publish → API calls → notifications → completion
 * by analysing event timing, source, and content proximity.
 */

import { IEventEntry, INetworkEvent, EventSeverity } from '../../models/IEventViewer';

// ============================================================================
// TYPES
// ============================================================================

export interface ICorrelationChain {
  /** Unique chain ID */
  id: string;
  /** Human-readable label for the chain */
  label: string;
  /** Chain type for categorisation */
  type: 'policy-save' | 'approval-flow' | 'notification-send' | 'quiz-generate' | 'data-load' | 'error-cascade' | 'unknown';
  /** Events in this chain (ordered by timestamp) */
  events: IEventEntry[];
  /** Start time of first event */
  startTime: string;
  /** End time of last event */
  endTime: string;
  /** Total duration in ms */
  durationMs: number;
  /** Whether any event in the chain has Error or Critical severity */
  hasErrors: boolean;
  /** SP list name involved (if applicable) */
  primaryTarget?: string;
}

// ============================================================================
// CHAIN DETECTION PATTERNS
// ============================================================================

interface IChainPattern {
  type: ICorrelationChain['type'];
  label: string;
  /** Patterns to match in event source or message */
  sourcePatterns: RegExp[];
  /** URL patterns for network events */
  urlPatterns?: RegExp[];
  /** Maximum time window between first and last event (ms) */
  maxWindowMs: number;
}

const CHAIN_PATTERNS: IChainPattern[] = [
  {
    type: 'policy-save',
    label: 'Policy Save',
    sourcePatterns: [/PolicyService/i, /PolicyAuditService/i],
    urlPatterns: [/PM_Policies/i, /PM_PolicyVersions/i, /PM_PolicyAuditLog/i],
    maxWindowMs: 10000,
  },
  {
    type: 'approval-flow',
    label: 'Approval Flow',
    sourcePatterns: [/ApprovalService/i, /ApprovalNotification/i],
    urlPatterns: [/PM_Approvals/i, /PM_ApprovalHistory/i, /PM_NotificationQueue/i],
    maxWindowMs: 15000,
  },
  {
    type: 'notification-send',
    label: 'Notification',
    sourcePatterns: [/NotificationRouter/i, /PolicyNotification/i, /EmailQueue/i],
    urlPatterns: [/PM_Notifications/i, /PM_NotificationQueue/i],
    maxWindowMs: 10000,
  },
  {
    type: 'quiz-generate',
    label: 'Quiz Generation',
    sourcePatterns: [/QuizBuilder/i, /generate-quiz/i],
    urlPatterns: [/generate-quiz-questions/i, /PM_PolicyQuiz/i],
    maxWindowMs: 30000,
  },
  {
    type: 'data-load',
    label: 'Data Load',
    sourcePatterns: [/PolicyHubService/i, /AdminConfigService/i],
    urlPatterns: [/\/_api\/web\/lists/i],
    maxWindowMs: 5000,
  },
];

// ============================================================================
// SERVICE
// ============================================================================

export class CorrelationService {

  /**
   * Analyse events and extract correlation chains.
   */
  public static buildChains(events: IEventEntry[]): ICorrelationChain[] {
    if (!events || events.length === 0) return [];

    const chains: ICorrelationChain[] = [];
    const used = new Set<string>();

    // Sort events by timestamp
    const sorted = events.slice().sort((a, b) => a.timestamp.localeCompare(b.timestamp));

    // Phase 1: Pattern-based chain detection
    for (const pattern of CHAIN_PATTERNS) {
      const matchingEvents = sorted.filter(e => {
        if (used.has(e.id)) return false;
        return CorrelationService._matchesPattern(e, pattern);
      });

      // Group into time windows
      const groups = CorrelationService._groupByTimeWindow(matchingEvents, pattern.maxWindowMs);

      for (const group of groups) {
        if (group.length < 2) continue; // Chains need at least 2 events

        const chainId = 'chain_' + Date.now() + '_' + Math.random().toString(36).substring(2, 6);
        const startTime = group[0].timestamp;
        const endTime = group[group.length - 1].timestamp;
        const durationMs = new Date(endTime).getTime() - new Date(startTime).getTime();

        chains.push({
          id: chainId,
          label: pattern.label,
          type: pattern.type,
          events: group,
          startTime,
          endTime,
          durationMs,
          hasErrors: group.some(e => e.severity >= EventSeverity.Error),
          primaryTarget: CorrelationService._extractTarget(group),
        });

        // Mark events as used
        for (const e of group) {
          used.add(e.id);
        }
      }
    }

    // Phase 2: Error cascade detection — group errors within 2s of each other
    const unusedErrors = sorted.filter(e =>
      !used.has(e.id) && e.severity >= EventSeverity.Error
    );

    const errorGroups = CorrelationService._groupByTimeWindow(unusedErrors, 2000);
    for (const group of errorGroups) {
      if (group.length < 2) continue;

      const chainId = 'chain_err_' + Date.now() + '_' + Math.random().toString(36).substring(2, 6);
      const startTime = group[0].timestamp;
      const endTime = group[group.length - 1].timestamp;

      chains.push({
        id: chainId,
        label: 'Error Cascade',
        type: 'error-cascade',
        events: group,
        startTime,
        endTime,
        durationMs: new Date(endTime).getTime() - new Date(startTime).getTime(),
        hasErrors: true,
      });

      for (const e of group) {
        used.add(e.id);
      }
    }

    // Sort chains by start time descending (newest first)
    return chains.sort((a, b) => b.startTime.localeCompare(a.startTime));
  }

  /** Check if an event matches a chain pattern */
  private static _matchesPattern(event: IEventEntry, pattern: IChainPattern): boolean {
    // Check source patterns
    for (const re of pattern.sourcePatterns) {
      if (re.test(event.source) || re.test(event.message)) return true;
    }

    // Check URL patterns for network events
    if (pattern.urlPatterns && (event as INetworkEvent).requestUrl) {
      const url = (event as INetworkEvent).requestUrl;
      for (const re of pattern.urlPatterns) {
        if (re.test(url)) return true;
      }
    }

    return false;
  }

  /** Group events into clusters within a time window */
  private static _groupByTimeWindow(events: IEventEntry[], windowMs: number): IEventEntry[][] {
    if (events.length === 0) return [];

    const groups: IEventEntry[][] = [];
    let currentGroup: IEventEntry[] = [events[0]];

    for (let i = 1; i < events.length; i++) {
      const prevTime = new Date(events[i - 1].timestamp).getTime();
      const currTime = new Date(events[i].timestamp).getTime();

      if (currTime - prevTime <= windowMs) {
        currentGroup.push(events[i]);
      } else {
        groups.push(currentGroup);
        currentGroup = [events[i]];
      }
    }

    groups.push(currentGroup);
    return groups;
  }

  /** Extract the primary SP list target from a chain */
  private static _extractTarget(events: IEventEntry[]): string | undefined {
    for (const e of events) {
      const net = e as INetworkEvent;
      if (net.spListName) return net.spListName;
    }
    return undefined;
  }
}
