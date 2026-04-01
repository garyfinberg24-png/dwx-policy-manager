/**
 * WatchRuleService — Custom alert rules that evaluate against the EventBuffer.
 * Admins can define conditions (e.g. >5 errors in 1 minute) and get notified
 * when rules trigger.
 */

import { EventBuffer } from './EventBuffer';
import { IEventEntry, EventSeverity } from '../../models/IEventViewer';

// ============================================================================
// TYPES
// ============================================================================

export interface IWatchRule {
  id: string;
  name: string;
  description: string;
  enabled: boolean;
  /** Condition type */
  condition: 'error-count' | 'severity-threshold' | 'slow-requests' | 'throttle-detect' | 'source-errors' | 'custom-pattern';
  /** Condition parameters */
  params: {
    /** Threshold count */
    threshold?: number;
    /** Time window in seconds */
    windowSeconds?: number;
    /** Minimum severity to count */
    minSeverity?: EventSeverity;
    /** Duration threshold in ms (for slow requests) */
    durationMs?: number;
    /** Source name filter */
    source?: string;
    /** Message pattern (regex string) */
    pattern?: string;
  };
}

export interface IWatchRuleAlert {
  ruleId: string;
  ruleName: string;
  message: string;
  triggeredAt: string;
  matchCount: number;
  severity: 'critical' | 'warning' | 'info';
}

// ============================================================================
// BUILT-IN RULES
// ============================================================================

export const DEFAULT_WATCH_RULES: IWatchRule[] = [
  {
    id: 'wr-error-burst',
    name: 'Error Burst',
    description: 'More than 5 errors within 60 seconds',
    enabled: true,
    condition: 'error-count',
    params: { threshold: 5, windowSeconds: 60, minSeverity: EventSeverity.Error },
  },
  {
    id: 'wr-critical-any',
    name: 'Critical Event',
    description: 'Any critical severity event detected',
    enabled: true,
    condition: 'severity-threshold',
    params: { threshold: 1, windowSeconds: 300, minSeverity: EventSeverity.Critical },
  },
  {
    id: 'wr-throttle',
    name: 'Throttle Detection',
    description: 'SP returns HTTP 429 (Too Many Requests)',
    enabled: true,
    condition: 'throttle-detect',
    params: { threshold: 1, windowSeconds: 120 },
  },
  {
    id: 'wr-slow-requests',
    name: 'Slow Requests',
    description: 'More than 3 requests over 3 seconds',
    enabled: true,
    condition: 'slow-requests',
    params: { threshold: 3, windowSeconds: 120, durationMs: 3000 },
  },
  {
    id: 'wr-policy-service-errors',
    name: 'PolicyService Errors',
    description: 'PolicyService producing errors',
    enabled: true,
    condition: 'source-errors',
    params: { threshold: 2, windowSeconds: 120, source: 'PolicyService' },
  },
  {
    id: 'wr-auth-failures',
    name: 'Auth Failures',
    description: '401/403 responses detected',
    enabled: true,
    condition: 'custom-pattern',
    params: { threshold: 2, windowSeconds: 60, pattern: '(401|403|Unauthorized|Forbidden)' },
  },
];

// ============================================================================
// SERVICE
// ============================================================================

export class WatchRuleService {
  private _rules: IWatchRule[];
  private _firedAlerts: Map<string, number> = new Map(); // ruleId → last fired timestamp

  constructor(rules?: IWatchRule[]) {
    this._rules = rules || DEFAULT_WATCH_RULES.slice();
  }

  public getRules(): IWatchRule[] {
    return this._rules.slice();
  }

  public setRules(rules: IWatchRule[]): void {
    this._rules = rules;
  }

  public toggleRule(ruleId: string, enabled: boolean): void {
    const rule = this._rules.find(r => r.id === ruleId);
    if (rule) rule.enabled = enabled;
  }

  /**
   * Evaluate all enabled rules against the current buffer.
   * Returns alerts for rules that fired.
   * Uses cooldown (30s per rule) to avoid alert spam.
   */
  public evaluate(buffer: EventBuffer): IWatchRuleAlert[] {
    const now = Date.now();
    const alerts: IWatchRuleAlert[] = [];
    const allEvents = buffer.getAll();
    const networkEvents = buffer.getNetworkEvents();

    for (const rule of this._rules) {
      if (!rule.enabled) continue;

      // Cooldown check — don't re-fire within 30 seconds
      const lastFired = this._firedAlerts.get(rule.id) || 0;
      if (now - lastFired < 30000) continue;

      const alert = this._evaluateRule(rule, allEvents, networkEvents, now);
      if (alert) {
        this._firedAlerts.set(rule.id, now);
        alerts.push(alert);
      }
    }

    return alerts;
  }

  /** Get all currently firing alerts (no cooldown reset) */
  public getActiveAlerts(buffer: EventBuffer): IWatchRuleAlert[] {
    const now = Date.now();
    const allEvents = buffer.getAll();
    const networkEvents = buffer.getNetworkEvents();
    const alerts: IWatchRuleAlert[] = [];

    for (const rule of this._rules) {
      if (!rule.enabled) continue;
      const alert = this._evaluateRule(rule, allEvents, networkEvents, now);
      if (alert) alerts.push(alert);
    }

    return alerts;
  }

  private _evaluateRule(
    rule: IWatchRule,
    allEvents: IEventEntry[],
    networkEvents: any[],
    now: number
  ): IWatchRuleAlert | null {
    const windowMs = (rule.params.windowSeconds || 60) * 1000;
    const cutoff = now - windowMs;
    const threshold = rule.params.threshold || 1;

    switch (rule.condition) {
      case 'error-count': {
        const minSev = rule.params.minSeverity ?? EventSeverity.Error;
        const count = allEvents.filter(e =>
          e.severity >= minSev && new Date(e.timestamp).getTime() > cutoff
        ).length;
        if (count >= threshold) {
          return {
            ruleId: rule.id, ruleName: rule.name,
            message: `${count} error${count !== 1 ? 's' : ''} in last ${rule.params.windowSeconds}s (threshold: ${threshold})`,
            triggeredAt: new Date().toISOString(), matchCount: count,
            severity: count >= threshold * 2 ? 'critical' : 'warning',
          };
        }
        break;
      }
      case 'severity-threshold': {
        const minSev = rule.params.minSeverity ?? EventSeverity.Critical;
        const count = allEvents.filter(e =>
          e.severity >= minSev && new Date(e.timestamp).getTime() > cutoff
        ).length;
        if (count >= threshold) {
          return {
            ruleId: rule.id, ruleName: rule.name,
            message: `${count} critical event${count !== 1 ? 's' : ''} detected`,
            triggeredAt: new Date().toISOString(), matchCount: count,
            severity: 'critical',
          };
        }
        break;
      }
      case 'throttle-detect': {
        const count = networkEvents.filter((e: any) =>
          e.httpStatus === 429 && new Date(e.timestamp).getTime() > cutoff
        ).length;
        if (count >= threshold) {
          return {
            ruleId: rule.id, ruleName: rule.name,
            message: `${count} throttled request${count !== 1 ? 's' : ''} (HTTP 429) in last ${rule.params.windowSeconds}s`,
            triggeredAt: new Date().toISOString(), matchCount: count,
            severity: 'critical',
          };
        }
        break;
      }
      case 'slow-requests': {
        const durMs = rule.params.durationMs || 3000;
        const count = networkEvents.filter((e: any) =>
          (e.duration || 0) > durMs && new Date(e.timestamp).getTime() > cutoff
        ).length;
        if (count >= threshold) {
          return {
            ruleId: rule.id, ruleName: rule.name,
            message: `${count} slow request${count !== 1 ? 's' : ''} (>${durMs}ms)`,
            triggeredAt: new Date().toISOString(), matchCount: count,
            severity: 'warning',
          };
        }
        break;
      }
      case 'source-errors': {
        const src = rule.params.source || '';
        const count = allEvents.filter(e =>
          e.source === src && e.severity >= EventSeverity.Error && new Date(e.timestamp).getTime() > cutoff
        ).length;
        if (count >= threshold) {
          return {
            ruleId: rule.id, ruleName: rule.name,
            message: `${src}: ${count} error${count !== 1 ? 's' : ''} in last ${rule.params.windowSeconds}s`,
            triggeredAt: new Date().toISOString(), matchCount: count,
            severity: 'warning',
          };
        }
        break;
      }
      case 'custom-pattern': {
        const pat = rule.params.pattern ? new RegExp(rule.params.pattern, 'i') : null;
        if (!pat) break;
        const count = allEvents.filter(e =>
          new Date(e.timestamp).getTime() > cutoff && (pat.test(e.message) || (e.eventCode && pat.test(e.eventCode)))
        ).length;
        if (count >= threshold) {
          return {
            ruleId: rule.id, ruleName: rule.name,
            message: `Pattern "${rule.params.pattern}" matched ${count} time${count !== 1 ? 's' : ''}`,
            triggeredAt: new Date().toISOString(), matchCount: count,
            severity: 'warning',
          };
        }
        break;
      }
    }

    return null;
  }
}
