/**
 * SessionComparisonService — Compares two sessions from PM_EventLog
 * to show what changed (more errors, new sources, timing differences).
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

// ============================================================================
// TYPES
// ============================================================================

export interface ISessionSnapshot {
  sessionId: string;
  eventCount: number;
  errorCount: number;
  warningCount: number;
  sources: Record<string, number>;
  startTime: string;
  endTime: string;
}

export interface ISessionDiff {
  sessionA: ISessionSnapshot;
  sessionB: ISessionSnapshot;
  eventCountDelta: number;
  errorCountDelta: number;
  newSources: string[];
  removedSources: string[];
  durationMs: number;
}

// ============================================================================
// SERVICE
// ============================================================================

export class SessionComparisonService {
  private _sp: SPFI;

  constructor(sp: SPFI) {
    this._sp = sp;
  }

  /**
   * Load the two most recent sessions from PM_EventLog and compare.
   */
  public async compareRecentSessions(): Promise<ISessionDiff | null> {
    const startTime = Date.now();

    try {
      // Get distinct session IDs (most recent first)
      const items = await this._sp.web.lists.getByTitle('PM_EventLog')
        .items
        .select('SessionId', 'EventSeverity', 'EventSource', 'Created')
        .orderBy('Created', false)
        .top(2000)();

      // Group by session
      const sessionMap: Record<string, typeof items> = {};
      for (const item of items) {
        const sid = item.SessionId || 'unknown';
        if (!sessionMap[sid]) sessionMap[sid] = [];
        sessionMap[sid].push(item);
      }

      const sessionIds = Object.keys(sessionMap);
      if (sessionIds.length < 2) return null;

      const snapA = this._buildSnapshot(sessionIds[0], sessionMap[sessionIds[0]]);
      const snapB = this._buildSnapshot(sessionIds[1], sessionMap[sessionIds[1]]);

      const sourcesA = Object.keys(snapA.sources);
      const sourcesB = Object.keys(snapB.sources);
      const newSources = sourcesA.filter(s => sourcesB.indexOf(s) === -1);
      const removedSources = sourcesB.filter(s => sourcesA.indexOf(s) === -1);

      return {
        sessionA: snapA,
        sessionB: snapB,
        eventCountDelta: snapA.eventCount - snapB.eventCount,
        errorCountDelta: snapA.errorCount - snapB.errorCount,
        newSources,
        removedSources,
        durationMs: Date.now() - startTime,
      };
    } catch (_) {
      return null;
    }
  }

  private _buildSnapshot(sessionId: string, items: any[]): ISessionSnapshot {
    const sources: Record<string, number> = {};
    let errorCount = 0;
    let warningCount = 0;

    for (const item of items) {
      const src = item.EventSource || 'Unknown';
      sources[src] = (sources[src] || 0) + 1;
      const sev = item.EventSeverity;
      if (sev === 'Error' || sev === 'Critical') errorCount++;
      if (sev === 'Warning') warningCount++;
    }

    return {
      sessionId,
      eventCount: items.length,
      errorCount,
      warningCount,
      sources,
      startTime: items.length > 0 ? items[items.length - 1].Created : '',
      endTime: items.length > 0 ? items[0].Created : '',
    };
  }
}
