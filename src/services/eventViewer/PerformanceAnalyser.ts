/**
 * PerformanceAnalyser — Calculates performance scores and detects optimisation
 * issues from captured network events.
 */

import { EventBuffer } from './EventBuffer';
import {
  INetworkEvent,
  IPerformanceScore,
  IPerformanceSubScore,
  IPerformanceIssue,
  IPerformanceComparison,
} from '../../models/IEventViewer';
import { SLOW_REQUEST_THRESHOLD_MS } from '../../constants/EventCodes';

// ============================================================================
// THRESHOLDS
// ============================================================================

const LATENCY_EXCELLENT = 200;
const LATENCY_GOOD = 500;
const LATENCY_FAIR = 1000;

const MAX_REQUESTS_PER_MINUTE = 15;
const THROTTLE_PENALTY = 20; // per 429 event
const ERROR_PENALTY = 5;     // per 4xx/5xx event

// ============================================================================
// ANALYSER
// ============================================================================

export class PerformanceAnalyser {

  /**
   * Calculate overall performance score from network events.
   */
  public static calculateScore(buffer: EventBuffer): IPerformanceScore {
    const events = buffer.getNetworkEvents().filter(e => !e.isAssetRequest);
    if (events.length === 0) {
      return {
        overall: 100,
        subScores: PerformanceAnalyser._emptySubScores(),
        issueCount: 0,
      };
    }

    const latencyScore = PerformanceAnalyser._calcLatencyScore(events);
    const throttlingScore = PerformanceAnalyser._calcThrottlingScore(events);
    const errorScore = PerformanceAnalyser._calcErrorScore(events);
    const payloadScore = PerformanceAnalyser._calcPayloadScore(events);
    const volumeScore = PerformanceAnalyser._calcVolumeScore(events);

    const subScores: IPerformanceSubScore[] = [latencyScore, throttlingScore, errorScore, payloadScore, volumeScore];
    const overall = Math.round(subScores.reduce((sum, s) => sum + s.score, 0) / subScores.length);

    const issues = PerformanceAnalyser.detectIssues(buffer);

    return { overall, subScores, issueCount: issues.length };
  }

  /**
   * Detect performance issues and generate optimisation cards with controls.
   */
  public static detectIssues(buffer: EventBuffer): IPerformanceIssue[] {
    const events = buffer.getNetworkEvents().filter(e => !e.isAssetRequest);
    const issues: IPerformanceIssue[] = [];

    // Issue: Excessive queries to same list
    const listCounts = PerformanceAnalyser._countByList(events);
    const topLists = Object.entries(listCounts).filter(([_, count]) => count > 10);
    if (topLists.length > 0) {
      const [listName, count] = topLists.sort((a, b) => b[1] - a[1])[0];
      const has429 = events.some(e => e.httpStatus === 429 && e.spListName === listName);
      issues.push({
        id: 'excessive-queries',
        title: `Excessive API Calls to ${listName}`,
        description: `${count} requests to ${listName} this session.${has429 ? ' Causing 429 throttling.' : ' Risk of throttling.'}`,
        severity: has429 ? 'high' : 'medium',
        impactPercent: has429 ? 85 : 60,
        controls: [
          { type: 'slider', label: 'Cache TTL (seconds)', configKey: 'Perf.CacheTTL', min: 0, max: 120, step: 5, value: 30, unit: 's' },
          { type: 'toggle', label: 'Request deduplication', configKey: 'Perf.RequestDedup', value: true, onLabel: 'Enabled', offLabel: 'Disabled' },
        ],
        prediction: `Est. -${has429 ? 60 : 40}% requests${has429 ? ', eliminates 429s' : ''}`,
        applied: false,
        configKeys: { 'Perf.CacheTTL': '30', 'Perf.RequestDedup': 'true' },
      });
    }

    // Issue: Large payloads (select * heuristic)
    const largePayloads = events.filter(e => {
      // Heuristic: if URL doesn't contain $select, it's probably select(*)
      const url = e.requestUrl || '';
      return url.indexOf('_api/') !== -1 && url.indexOf('$select') === -1 && e.httpMethod === 'GET';
    });
    if (largePayloads.length >= 2) {
      issues.push({
        id: 'large-payloads',
        title: 'Oversized Payloads (select * detected)',
        description: `${largePayloads.length} queries without column selection — returning all fields.`,
        severity: 'medium',
        impactPercent: 55,
        controls: [
          { type: 'toggle', label: 'Lean query mode (specific columns)', configKey: 'Perf.LeanQueries', value: true, onLabel: 'Enabled', offLabel: 'Disabled' },
          { type: 'slider', label: 'Default $top limit', configKey: 'Perf.DefaultTopLimit', min: 25, max: 500, step: 25, value: 100 },
        ],
        prediction: 'Est. -40% payload size, -120ms latency',
        applied: false,
        configKeys: { 'Perf.LeanQueries': 'true', 'Perf.DefaultTopLimit': '100' },
      });
    }

    // Issue: Too many concurrent requests
    const concurrentPeaks = PerformanceAnalyser._detectConcurrency(events);
    if (concurrentPeaks > 6) {
      issues.push({
        id: 'concurrent-requests',
        title: 'No Concurrent Request Limit',
        description: `Up to ${concurrentPeaks} simultaneous SP API calls detected. SharePoint throttles at 6+ concurrent requests.`,
        severity: 'medium',
        impactPercent: 50,
        controls: [
          { type: 'slider', label: 'Max concurrent requests', configKey: 'Perf.MaxConcurrent', min: 1, max: 10, step: 1, value: 4 },
        ],
        prediction: 'Est. eliminates concurrent throttling',
        applied: false,
        configKeys: { 'Perf.MaxConcurrent': '4' },
      });
    }

    // Issue: Slow queries on specific lists
    const slowLists = PerformanceAnalyser._findSlowLists(events);
    for (const sl of slowLists) {
      issues.push({
        id: `slow-query-${sl.listName}`,
        title: `Slow Query on ${sl.listName}`,
        description: `Average ${sl.avgDuration}ms latency. ${sl.hasFilter ? '' : 'Query lacks $filter — loading all items.'}`,
        severity: 'low',
        impactPercent: 30,
        controls: [
          { type: 'toggle', label: 'Server-side filtering', configKey: `Perf.ServerFilter.${sl.listName}`, value: true, onLabel: 'Enabled', offLabel: 'Disabled' },
        ],
        prediction: `Est. -${Math.round((1 - 200 / sl.avgDuration) * 100)}% latency on ${sl.listName}`,
        applied: false,
        configKeys: { [`Perf.ServerFilter.${sl.listName}`]: 'true' },
      });
    }

    // Sort by impact descending
    return issues.sort((a, b) => b.impactPercent - a.impactPercent);
  }

  /**
   * Generate before/after comparison metrics.
   */
  public static generateComparison(buffer: EventBuffer, issues: IPerformanceIssue[]): IPerformanceComparison[] {
    const events = buffer.getNetworkEvents().filter(e => !e.isAssetRequest);
    if (events.length === 0) return [];

    const totalDuration = events.reduce((s, e) => s + (e.duration || 0), 0);
    const avgLatency = Math.round(totalDuration / events.length);
    const durations = events.map(e => e.duration || 0).sort((a, b) => a - b);
    const p95 = durations[Math.floor(durations.length * 0.95)] || 0;
    const throttled = events.filter(e => e.httpStatus === 429).length;
    const reqPerMin = PerformanceAnalyser._requestsPerMinute(events);

    // Project improvements based on detected issues
    const hasDedup = issues.some(i => i.id === 'excessive-queries' && !i.applied);
    const hasLean = issues.some(i => i.id === 'large-payloads' && !i.applied);

    const projLatency = Math.round(avgLatency * (hasLean ? 0.6 : 0.85));
    const projP95 = Math.round(p95 * (hasDedup ? 0.5 : 0.7));
    const projReqMin = Math.round(reqPerMin * (hasDedup ? 0.4 : 0.7));
    const projThrottled = hasDedup ? 0 : throttled;

    const currentScore = PerformanceAnalyser.calculateScore(buffer).overall;
    const projectedScore = Math.min(100, currentScore + issues.filter(i => !i.applied).length * 5);

    return [
      { metric: 'Avg Latency', current: `${avgLatency}ms`, projected: `~${projLatency}ms`, improved: projLatency < avgLatency },
      { metric: 'P95 Latency', current: `${p95}ms`, projected: `~${projP95}ms`, improved: projP95 < p95 },
      { metric: 'Requests / min', current: `${reqPerMin}`, projected: `~${projReqMin}`, improved: projReqMin < reqPerMin },
      { metric: '429 Throttled', current: `${throttled}`, projected: `${projThrottled}`, improved: projThrottled < throttled },
      { metric: 'Score', current: `${currentScore}`, projected: `~${projectedScore}`, improved: projectedScore > currentScore },
    ];
  }

  // ==========================================================================
  // PRIVATE — Sub-score calculators
  // ==========================================================================

  private static _calcLatencyScore(events: INetworkEvent[]): IPerformanceSubScore {
    const durations = events.filter(e => e.duration !== undefined).map(e => e.duration!);
    if (durations.length === 0) return { label: 'API Latency', key: 'latency', score: 100, detail: 'No data' };

    const avg = Math.round(durations.reduce((s, d) => s + d, 0) / durations.length);
    let score = 100;
    if (avg > LATENCY_FAIR) score = 40;
    else if (avg > LATENCY_GOOD) score = 65;
    else if (avg > LATENCY_EXCELLENT) score = 85;

    return { label: 'API Latency', key: 'latency', score, detail: `Avg ${avg}ms` };
  }

  private static _calcThrottlingScore(events: INetworkEvent[]): IPerformanceSubScore {
    const throttled = events.filter(e => e.httpStatus === 429).length;
    const score = Math.max(0, 100 - throttled * THROTTLE_PENALTY);
    return { label: 'Throttling', key: 'throttling', score, detail: `${throttled}x 429 errors` };
  }

  private static _calcErrorScore(events: INetworkEvent[]): IPerformanceSubScore {
    const errors = events.filter(e => e.httpStatus && e.httpStatus >= 400 && e.httpStatus !== 429).length;
    const score = Math.max(0, 100 - errors * ERROR_PENALTY);
    const rate = events.length > 0 ? ((errors / events.length) * 100).toFixed(1) : '0';
    return { label: 'Error Rate', key: 'errorRate', score, detail: `${rate}% failure` };
  }

  private static _calcPayloadScore(events: INetworkEvent[]): IPerformanceSubScore {
    const noSelect = events.filter(e => {
      const url = e.requestUrl || '';
      return url.indexOf('_api/') !== -1 && url.indexOf('$select') === -1 && e.httpMethod === 'GET';
    }).length;
    const score = noSelect === 0 ? 100 : Math.max(30, 100 - noSelect * 15);
    return { label: 'Payload', key: 'payload', score, detail: `${noSelect}x select(*)` };
  }

  private static _calcVolumeScore(events: INetworkEvent[]): IPerformanceSubScore {
    const rpm = PerformanceAnalyser._requestsPerMinute(events);
    let score = 100;
    if (rpm > MAX_REQUESTS_PER_MINUTE * 2) score = 40;
    else if (rpm > MAX_REQUESTS_PER_MINUTE) score = 65;
    else if (rpm > MAX_REQUESTS_PER_MINUTE * 0.7) score = 85;

    return { label: 'Request Vol.', key: 'volume', score, detail: `${rpm} req/min` };
  }

  private static _emptySubScores(): IPerformanceSubScore[] {
    return [
      { label: 'API Latency', key: 'latency', score: 100, detail: 'No data' },
      { label: 'Throttling', key: 'throttling', score: 100, detail: '0x 429' },
      { label: 'Error Rate', key: 'errorRate', score: 100, detail: '0%' },
      { label: 'Payload', key: 'payload', score: 100, detail: 'OK' },
      { label: 'Request Vol.', key: 'volume', score: 100, detail: '0 req/min' },
    ];
  }

  // ==========================================================================
  // PRIVATE — Helpers
  // ==========================================================================

  private static _countByList(events: INetworkEvent[]): Record<string, number> {
    const map: Record<string, number> = {};
    for (let i = 0; i < events.length; i++) {
      const name = events[i].spListName;
      if (name) map[name] = (map[name] || 0) + 1;
    }
    return map;
  }

  private static _detectConcurrency(events: INetworkEvent[]): number {
    // Simple heuristic: count overlapping requests by timestamp windows
    let maxConcurrent = 0;
    for (let i = 0; i < events.length; i++) {
      const start = new Date(events[i].timestamp).getTime();
      const end = start + (events[i].duration || 0);
      let concurrent = 1;
      for (let j = 0; j < events.length; j++) {
        if (i === j) continue;
        const otherStart = new Date(events[j].timestamp).getTime();
        if (otherStart >= start && otherStart <= end) concurrent++;
      }
      if (concurrent > maxConcurrent) maxConcurrent = concurrent;
    }
    return maxConcurrent;
  }

  private static _findSlowLists(events: INetworkEvent[]): Array<{ listName: string; avgDuration: number; hasFilter: boolean }> {
    const listDurations: Record<string, number[]> = {};
    const listFilters: Record<string, boolean> = {};

    for (let i = 0; i < events.length; i++) {
      const name = events[i].spListName;
      if (!name || !events[i].duration) continue;
      if (!listDurations[name]) { listDurations[name] = []; listFilters[name] = false; }
      listDurations[name].push(events[i].duration!);
      if ((events[i].requestUrl || '').indexOf('$filter') !== -1) listFilters[name] = true;
    }

    const results: Array<{ listName: string; avgDuration: number; hasFilter: boolean }> = [];
    for (const [name, durations] of Object.entries(listDurations)) {
      const avg = Math.round(durations.reduce((s, d) => s + d, 0) / durations.length);
      if (avg > SLOW_REQUEST_THRESHOLD_MS) {
        results.push({ listName: name, avgDuration: avg, hasFilter: listFilters[name] });
      }
    }
    return results;
  }

  private static _requestsPerMinute(events: INetworkEvent[]): number {
    if (events.length < 2) return events.length;
    const timestamps = events.map(e => new Date(e.timestamp).getTime());
    const span = (Math.max(...timestamps) - Math.min(...timestamps)) / 60000;
    return span > 0 ? Math.round(events.length / span) : events.length;
  }
}
