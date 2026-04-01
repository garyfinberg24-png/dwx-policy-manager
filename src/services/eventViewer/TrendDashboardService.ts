/**
 * TrendDashboardService — Loads historical event data from PM_EventLog
 * and computes error trends over time for the Trend Dashboard view.
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

// ============================================================================
// TYPES
// ============================================================================

export interface ITrendDataPoint {
  /** Date label (e.g. "Mar 28") */
  label: string;
  /** ISO date */
  date: string;
  /** Error count for this period */
  errors: number;
  /** Warning count */
  warnings: number;
  /** Total events */
  total: number;
}

export interface ITrendSummary {
  dataPoints: ITrendDataPoint[];
  totalErrors: number;
  totalWarnings: number;
  totalEvents: number;
  trendDirection: 'improving' | 'stable' | 'worsening';
  topSources: Array<{ source: string; count: number }>;
  durationMs: number;
}

// ============================================================================
// SERVICE
// ============================================================================

export class TrendDashboardService {
  private _sp: SPFI;

  constructor(sp: SPFI) {
    this._sp = sp;
  }

  /**
   * Load trend data for the last N days from PM_EventLog.
   */
  public async loadTrends(days: number = 7): Promise<ITrendSummary> {
    const startTime = Date.now();
    const cutoffDate = new Date(Date.now() - days * 24 * 60 * 60 * 1000).toISOString();

    try {
      const items = await this._sp.web.lists.getByTitle('PM_EventLog')
        .items
        .filter(`Created ge '${cutoffDate}'`)
        .select('Id', 'EventSeverity', 'EventSource', 'Created')
        .orderBy('Created', false)
        .top(2000)();

      // Group by day
      const dayMap: Record<string, { errors: number; warnings: number; total: number }> = {};
      const sourceMap: Record<string, number> = {};

      for (const item of items) {
        const date = new Date(item.Created).toISOString().slice(0, 10);
        if (!dayMap[date]) dayMap[date] = { errors: 0, warnings: 0, total: 0 };
        dayMap[date].total++;

        const sev = item.EventSeverity;
        if (sev === 'Error' || sev === 'Critical') dayMap[date].errors++;
        if (sev === 'Warning') dayMap[date].warnings++;

        const src = item.EventSource || 'Unknown';
        sourceMap[src] = (sourceMap[src] || 0) + 1;
      }

      // Build data points for each day in range
      const dataPoints: ITrendDataPoint[] = [];
      for (let i = days - 1; i >= 0; i--) {
        const d = new Date(Date.now() - i * 24 * 60 * 60 * 1000);
        const dateStr = d.toISOString().slice(0, 10);
        const monthDay = d.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
        const dayData = dayMap[dateStr] || { errors: 0, warnings: 0, total: 0 };
        dataPoints.push({
          label: monthDay,
          date: dateStr,
          errors: dayData.errors,
          warnings: dayData.warnings,
          total: dayData.total,
        });
      }

      // Trend direction: compare last 2 days vs previous 2 days
      const recentErrors = dataPoints.slice(-2).reduce((s, p) => s + p.errors, 0);
      const previousErrors = dataPoints.slice(-4, -2).reduce((s, p) => s + p.errors, 0);
      const trendDirection = recentErrors < previousErrors ? 'improving'
        : recentErrors > previousErrors ? 'worsening'
        : 'stable';

      // Top sources
      const topSources = Object.entries(sourceMap)
        .map(([source, count]) => ({ source, count }))
        .sort((a, b) => b.count - a.count)
        .slice(0, 5);

      return {
        dataPoints,
        totalErrors: items.filter(i => i.EventSeverity === 'Error' || i.EventSeverity === 'Critical').length,
        totalWarnings: items.filter(i => i.EventSeverity === 'Warning').length,
        totalEvents: items.length,
        trendDirection,
        topSources,
        durationMs: Date.now() - startTime,
      };
    } catch (err) {
      return {
        dataPoints: [],
        totalErrors: 0,
        totalWarnings: 0,
        totalEvents: 0,
        trendDirection: 'stable',
        topSources: [],
        durationMs: Date.now() - startTime,
      };
    }
  }
}
