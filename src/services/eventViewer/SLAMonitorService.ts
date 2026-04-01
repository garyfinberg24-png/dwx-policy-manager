/**
 * SLAMonitorService — Monitors response time SLAs per SP list.
 * Computes P50/P95/P99 latencies from NetworkInterceptor data.
 */

import { EventBuffer } from './EventBuffer';
import { INetworkEvent } from '../../models/IEventViewer';

// ============================================================================
// TYPES
// ============================================================================

export interface IListSLA {
  listName: string;
  requestCount: number;
  p50: number;
  p95: number;
  p99: number;
  maxLatency: number;
  errorRate: number;
  /** Breach if p95 exceeds target */
  targetMs: number;
  breached: boolean;
}

export interface ISLAMonitorSummary {
  lists: IListSLA[];
  overallP50: number;
  overallP95: number;
  breachCount: number;
}

// Default SLA targets per list category
const DEFAULT_SLA_MS = 2000; // 2 seconds P95 target

// ============================================================================
// SERVICE
// ============================================================================

export class SLAMonitorService {

  /**
   * Compute SLA metrics from current network event buffer.
   */
  public static compute(buffer: EventBuffer, targetMs?: number): ISLAMonitorSummary {
    const target = targetMs || DEFAULT_SLA_MS;
    const networkEvents = buffer.getNetworkEvents();

    // Group by SP list name
    const listMap: Record<string, INetworkEvent[]> = {};
    for (const e of networkEvents) {
      const name = (e as INetworkEvent).spListName;
      if (!name) continue;
      if (!listMap[name]) listMap[name] = [];
      listMap[name].push(e as INetworkEvent);
    }

    const lists: IListSLA[] = [];
    const allDurations: number[] = [];

    for (const [listName, events] of Object.entries(listMap)) {
      const durations = events
        .map(e => e.duration || 0)
        .filter(d => d > 0)
        .sort((a, b) => a - b);

      if (durations.length === 0) continue;
      allDurations.push(...durations);

      const errorCount = events.filter(e => (e.httpStatus || 0) >= 400).length;

      lists.push({
        listName,
        requestCount: events.length,
        p50: SLAMonitorService._percentile(durations, 50),
        p95: SLAMonitorService._percentile(durations, 95),
        p99: SLAMonitorService._percentile(durations, 99),
        maxLatency: durations[durations.length - 1],
        errorRate: events.length > 0 ? Math.round((errorCount / events.length) * 1000) / 10 : 0,
        targetMs: target,
        breached: SLAMonitorService._percentile(durations, 95) > target,
      });
    }

    // Sort by P95 descending (worst first)
    lists.sort((a, b) => b.p95 - a.p95);

    allDurations.sort((a, b) => a - b);

    return {
      lists,
      overallP50: allDurations.length > 0 ? SLAMonitorService._percentile(allDurations, 50) : 0,
      overallP95: allDurations.length > 0 ? SLAMonitorService._percentile(allDurations, 95) : 0,
      breachCount: lists.filter(l => l.breached).length,
    };
  }

  private static _percentile(sorted: number[], p: number): number {
    if (sorted.length === 0) return 0;
    const idx = Math.ceil((p / 100) * sorted.length) - 1;
    return sorted[Math.max(0, Math.min(idx, sorted.length - 1))];
  }
}
